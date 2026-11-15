#nullable enable

using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

/// <summary>
/// Normalizes an ISO/IEC 29500 STRICT-conformance WordprocessingML package (namespace family
/// <c>http://purl.oclc.org/ooxml/*</c>, Word's "Strict Open XML Document" save format) into its
/// transitional equivalent. Word performs the same translation transparently on open; Docxodus'
/// XDocument-based pipelines (IR reader, comparers, revision processing) only understand the
/// transitional namespaces, so strict inputs are converted up front, and so is a strict part inside a
/// transitional package. Transitional inputs with no strict part are returned unchanged (same instance,
/// no copy).
/// </summary>
internal static class StrictOoxmlNormalizer
{
    private const string StrictMarker = "http://purl.oclc.org/ooxml/";

    // Strict → transitional URI map (ISO/IEC 29500-4 Annex A, the families a WML package carries).
    // ORDER MATTERS: the extendedProperties entries must precede the generic
    // "officeDocument/relationships" prefix entry — transitional renames that family with a hyphen
    // ("extended-properties"), so the generic prefix rewrite would otherwise produce a URI that
    // exists in neither conformance class.
    private static readonly (string Strict, string Transitional)[] UriMap =
    {
        ("http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties",
         "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"),
        ("http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
         "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"),
        // One prefix rule covers xmlns:r AND every relationship Type in the .rels streams.
        ("http://purl.oclc.org/ooxml/officeDocument/relationships",
         "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
        ("http://purl.oclc.org/ooxml/wordprocessingml/main",
         "http://schemas.openxmlformats.org/wordprocessingml/2006/main"),
        ("http://purl.oclc.org/ooxml/drawingml/main",
         "http://schemas.openxmlformats.org/drawingml/2006/main"),
        ("http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing",
         "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"),
        ("http://purl.oclc.org/ooxml/drawingml/picture",
         "http://schemas.openxmlformats.org/drawingml/2006/picture"),
        ("http://purl.oclc.org/ooxml/drawingml/chart",
         "http://schemas.openxmlformats.org/drawingml/2006/chart"),
        ("http://purl.oclc.org/ooxml/drawingml/chartDrawing",
         "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing"),
        ("http://purl.oclc.org/ooxml/drawingml/diagram",
         "http://schemas.openxmlformats.org/drawingml/2006/diagram"),
        ("http://purl.oclc.org/ooxml/officeDocument/math",
         "http://schemas.openxmlformats.org/officeDocument/2006/math"),
        ("http://purl.oclc.org/ooxml/officeDocument/sharedTypes",
         "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes"),
        ("http://purl.oclc.org/ooxml/officeDocument/customXml",
         "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
        ("http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes",
         "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"),
        ("http://purl.oclc.org/ooxml/officeDocument/bibliography",
         "http://schemas.openxmlformats.org/officeDocument/2006/bibliography"),
        ("http://purl.oclc.org/ooxml/schemaLibrary/main",
         "http://schemas.openxmlformats.org/schemaLibrary/2006/main"),
    };

    /// <summary>
    /// Returns the transitional-conformance equivalent of <paramref name="doc"/>, or
    /// <paramref name="doc"/> itself (no copy) when no part of it is strict.
    /// </summary>
    internal static WmlDocument NormalizeToTransitional(WmlDocument doc)
    {
        var bytes = NormalizeToTransitional(doc.DocumentByteArray);
        return ReferenceEquals(bytes, doc.DocumentByteArray) ? doc : new WmlDocument(doc.FileName, bytes);
    }

    /// <summary>
    /// Returns the transitional-conformance equivalent of raw DOCX bytes. This avoids constructing
    /// a <see cref="WmlDocument"/> for the overwhelmingly common transitional case, which matters
    /// for byte-based rendering entry points.
    /// </summary>
    internal static byte[] NormalizeToTransitional(byte[] docxBytes) => NormalizeToTransitional(docxBytes, trusted: false);

    /// <param name="trusted">The bytes are the session's own checkpoint, already read through the
    /// depth door once, so they are not scanned again (issue #1389).</param>
    private static byte[] NormalizeToTransitional(byte[] docxBytes, bool trusted)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        // The XML depth door (issue #1389), before any tree is built. A strict package is measured as
        // the unfold leaves it: its text rewrites can make an entry well-formed, and so deep, that was
        // not XML before (the unfold guards its own parse).
        var result = StrictEntries(docxBytes) is { } strictEntries ? NormalizeStrict(docxBytes, strictEntries) : docxBytes;
        if (!trusted) Internal.PartDepthGuard.EnsurePackage(result);
        return result;
    }

    /// <summary>
    /// <see cref="NormalizeToTransitional(byte[])"/> plus <see cref="PackageOpenRepair.Repair"/>, for
    /// code about to open the package to read or change its content. A package too damaged to convert
    /// is returned unchanged, so the open that follows reports the damage exactly as it did before
    /// conversion existed.
    /// </summary>
    internal static byte[] NormalizeForReading(byte[] docxBytes, bool trusted = false)
    {
        try
        {
            var transitional = NormalizeToTransitional(docxBytes, trusted);
            var repaired = PackageOpenRepair.Repair(transitional);
            // Repair can give an untyped part the type its relationship implies, exposing a part the
            // packaging layer did not see before: measure what it wrote.
            if (!trusted && !ReferenceEquals(repaired, transitional)) Internal.PartDepthGuard.EnsurePackage(repaired);
            return repaired;
        }
        catch (Exception e) when (e is InvalidDataException or IOException)
        {
            return docxBytes;
        }
    }

    private static byte[] NormalizeStrict(byte[] documentBytes, Predicate<string> converts)
    {
        using var ms = new MemoryStream();
        ms.Write(documentBytes, 0, documentBytes.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if ((!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                     !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) ||
                    !converts(entry.FullName))
                    continue;

                string text;
                try
                {
                    using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                    text = reader.ReadToEnd();
                }
                catch (InvalidDataException)
                {
                    // A damaged entry is copied through unchanged: a reader that never opens it
                    // still gets the rest of the package converted, and one that does fails on it
                    // as it would have without the conversion.
                    continue;
                }
                if (!text.Contains(StrictMarker, StringComparison.Ordinal))
                    continue;

                var rewritten = text;
                foreach (var (strict, transitional) in UriMap)
                    rewritten = rewritten.Replace(strict, transitional, StringComparison.Ordinal);
                // Word stamps w:conformance="strict" on w:document; a transitional package
                // never carries the attribute, so drop it rather than leave a stale marker.
                rewritten = rewritten
                    .Replace(" w:conformance=\"strict\"", string.Empty, StringComparison.Ordinal)
                    .Replace(" conformance=\"strict\"", string.Empty, StringComparison.Ordinal);
                rewritten = RewriteTree(rewritten);

                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return ms.ToArray();
    }

    private static readonly XNamespace TransitionalWpDrawing =
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace TransitionalA =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName TransitionalWDrawing =
        XName.Get("drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    private const string Ms2010WpPrefix = "http://schemas.microsoft.com/office/word/2010/wordprocessing";

    /// <summary>
    /// The rewrites the flat URI substitution cannot make, because they need the part's tree: the
    /// shape-namespace fold and the values Strict writes in a form transitional does not have. A part
    /// that is not well-formed, or nests too deep, is left to the package's own depth measurement.
    /// </summary>
    private static string RewriteTree(string xml)
    {
        XDocument doc;
        try
        {
            // Depth-limited: the recursive re-homing must not meet a deeper tree, and this entry may be
            // one the read step never measured; the package is measured again afterwards.
            doc = Internal.DepthLimitedXmlReader.Parse(xml, LoadOptions.PreserveWhitespace);
        }
        catch (System.Xml.XmlException)
        {
            return xml;
        }
        if (!(UnfoldGraphicDataNamespaces(doc) | ConvertStrictValues(doc)))
            return xml;
        // StringWriter reports UTF-16, which XDocument.Save stamps into the XML declaration — but
        // the entry is written as UTF-8, making the SDK's reader refuse the part ("There is no
        // Unicode byte order mark"). Declare UTF-8 explicitly.
        using var sw = new Utf8StringWriter();
        doc.SaveLf(sw, SaveOptions.DisableFormatting);
        return sw.ToString();
    }

    // Nested so this Strict-only table stays out of the class's AOT-compiled type initializer and off the Transitional path.
    private static class StrictValueTables
    {
        private static readonly XNamespace ChartStyle = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

        /// <summary>
        /// The attributes of DrawingML's percentage types (ST_Percentage, ST_PositivePercentage,
        /// ST_FixedPercentage, ST_PositiveFixedPercentage and the text spacing, font scale and bullet size
        /// types built on them), by element: the DrawingML main schema, the Office 2010 drawing extensions
        /// (a14) and the chart style part. Keyed by element, so free text that reads like a percentage
        /// (a description, a tooltip, a number format of "0%") is never touched.
        /// </summary>
        internal static readonly ILookup<XName, XName> PercentageAttributes = new (XNamespace Ns, string Attributes, string Elements)[]
        {
            (A.a, "val", "alpha alphaMod alphaOff blue blueMod blueOff buSzPct green greenMod greenOff hueMod lum lumMod lumOff " +
                         "red redMod redOff sat satMod satOff shade spcPct tint"),
            (A.a, "l t r b", "fillRect fillToRect srcRect tileRect"),
            (A.a, "r g b", "scrgbClr"),
            (A.a, "sat lum", "hsl hslClr"),
            (A.a, "sx sy", "outerShdw reflection tile xfrm"),
            (A.a, "stA stPos endA endPos", "reflection"),
            (A.a, "baseline", "defRPr endParaRPr rPr"),
            (A.a, "bright contrast", "lum"),
            (A.a, "thresh", "alphaBiLevel biLevel"),
            (A.a, "a", "alphaRepl"),
            (A.a, "amt", "alphaModFix tint"),
            (A.a, "d sp", "ds"),
            (A.a, "lim", "miter"),
            (A.a, "pos", "gs"),
            (A.a, "tx ty", "relOff"),
            (A.a, "zoom", "camera"),
            (A.a, "fontScale lnSpcReduction", "normAutofit"),
            (ChartStyle, "baseline", "defRPr"),
            (A14.a14, "bright contrast", "brightnessContrast"),
            (A14.a14, "sat", "saturation"),
            (A14.a14, "amount", "sharpenSoften"),
            (A14.a14, "t b l r", "backgroundRemoval"),
            (A14.a14, "x1 y1 x2 y2", "foregroundMark backgroundMark"),
            (A14.a14, "trans", "artisticBlur artisticCement artisticChalkSketch artisticCrisscrossEtching artisticCutout " +
                               "artisticFilmGrain artisticGlass artisticGlowDiffused artisticGlowEdges artisticLightScreen " +
                               "artisticLineDrawing artisticMarker artisticMosiaicBubbles artisticPaintBrush artisticPaintStrokes " +
                               "artisticPastelsSmooth artisticPencilGrayscale artisticPencilSketch artisticPhotocopy " +
                               "artisticPlasticWrap artisticTexturizer artisticWatercolorSponge"),
        }.SelectMany(row => row.Elements.Split(' ').SelectMany(element =>
            row.Attributes.Split(' ').Select(attribute => (Element: row.Ns + element, Attribute: (XName)attribute))))
        .ToLookup(p => p.Element, p => p.Attribute);

        /// <summary>The relative size and position elements whose text is a DrawingML percentage.</summary>
        internal static readonly XName[] PercentageText =
            { WP14.pctWidth, WP14.pctHeight, WP14.wp14 + "pctPosHOffset", WP14.pctPosVOffset };
    }

    /// <summary>
    /// Rewrites the values Strict writes in a form transitional does not have: a DrawingML percentage
    /// ("50%") becomes thousandths of a percent ("50000"), and a <c>w:cnfStyle</c>, which in Strict
    /// carries only its flags, gains the <c>w:val</c> transitional requires, spelled from those flags.
    /// A value already in its transitional form is left as it is.
    /// </summary>
    private static bool ConvertStrictValues(XDocument doc)
    {
        var changed = false;
        foreach (var element in doc.Descendants().ToList())
        {
            foreach (var attribute in StrictValueTables.PercentageAttributes[element.Name].Select(element.Attribute).OfType<XAttribute>())
            {
                if (Thousandths(attribute.Value) is not { } value) continue;
                attribute.Value = value;
                changed = true;
            }
            if (StrictValueTables.PercentageText.Contains(element.Name) && !element.HasElements && Thousandths(element.Value) is { } text)
            {
                element.Value = text;
                changed = true;
            }
            if (element.Name == W.cnfStyle && element.Attribute(W.val) is null)
            {
                // Strict's ST_OnOff is xsd:boolean, so a flag is on when it is "1" or "true".
                var val = string.Concat(W.CnfStyleValFlags.Select(flag => (string?)element.Attribute(flag) is "1" or "true" ? '1' : '0'));
                element.ReplaceAttributes(new XAttribute(W.val, val), element.Attributes().ToList());
                changed = true;
            }
        }
        return changed;
    }

    /// <summary>"50%" as thousandths of a percent ("50000", "-12.5%" as "-12500"); null for a value that
    /// is not a percentage or does not fit the transitional types' integer range.</summary>
    private static string? Thousandths(string value) =>
        value.EndsWith('%') && decimal.TryParse(value.AsSpan(0, value.Length - 1),
            NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out var percent)
        && Math.Abs(percent) <= int.MaxValue / 1000m
            ? ((int)Math.Round(percent * 1000m, MidpointRounding.AwayFromZero)).ToString(CultureInfo.InvariantCulture)
            : null;

    /// <summary>
    /// Un-fold Word's strict-save namespace fold: strict packages write MS-2010
    /// wordprocessingShape/Group/Canvas payload elements (<c>wsp/spPr/bodyPr/…</c>) IN the strict
    /// wordprocessingDrawing namespace, with only <c>a:graphicData/@uri</c> naming the real payload
    /// namespace. Word un-folds on open; after the flat URI substitution those elements sit in the
    /// TRANSITIONAL wpDrawing namespace — names that do not exist there — and LibreOffice silently
    /// drops the whole shape. Re-home each such descendant to the <c>@uri</c> namespace, stopping at
    /// nested <c>w:drawing</c> boundaries (a drawing inside a <c>wps:txbx</c> keeps its own genuine
    /// wpDrawing container elements).
    /// </summary>
    private static bool UnfoldGraphicDataNamespaces(XDocument doc)
    {
        bool changed = false;
        foreach (var graphicData in doc.Descendants(TransitionalA + "graphicData"))
        {
            var uri = (string?)graphicData.Attribute("uri");
            if (uri is null || !uri.StartsWith(Ms2010WpPrefix, StringComparison.Ordinal))
                continue;
            XNamespace target = uri;
            void Rehome(XElement el)
            {
                foreach (var child in el.Elements())
                {
                    if (child.Name == TransitionalWDrawing)
                        continue;
                    if (child.Name.Namespace == TransitionalWpDrawing)
                    {
                        child.Name = target + child.Name.LocalName;
                        changed = true;
                    }
                    Rehome(child);
                }
            }
            Rehome(graphicData);
        }
        return changed;
    }

    private sealed class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }

    /// <summary>
    /// The entries to convert, or null when there are none. A package is strict when its main document
    /// part's root element lives in the strict WordprocessingML namespace, and then every entry is. In a
    /// transitional package, a part whose own root element is in a strict namespace (a chart a strict
    /// producer wrote, issue #1764) is converted alone. The main part is resolved through _rels/.rels
    /// (either conformance class's officeDocument relationship type), falling back to the conventional
    /// word/document.xml path.
    /// </summary>
    private static Predicate<string>? StrictEntries(byte[] documentBytes)
    {
        try
        {
            using var ms = new MemoryStream(documentBytes, writable: false);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var main = FindMainDocumentEntry(zip);
            if (main is not null && HasStrictRoot(main))
                return _ => true;
            var parts = zip.Entries.Where(e => e != main && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && HasStrictRoot(e)).Select(e => e.FullName).ToHashSet(StringComparer.Ordinal);
            return parts.Count == 0 ? null : parts.Contains;
        }
        catch (InvalidDataException)
        {
            return null; // not a zip — let the downstream open path produce its own error
        }
    }

    // Only the root element's namespace decides, so read just its start tag: this check runs on every
    // part of every package the pipelines open, transitional or not, before the depth door measures
    // anything, so it reads no further than a start tag can reasonably run, however far the part
    // inflates. A part it cannot read is left to the downstream open path to report.
    private static bool HasStrictRoot(ZipArchiveEntry entry)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = System.Xml.XmlReader.Create(stream, RootTagOnly);
            return reader.MoveToContent() == System.Xml.XmlNodeType.Element &&
                   reader.NamespaceURI.StartsWith(StrictMarker, StringComparison.Ordinal);
        }
        catch (Exception e) when (e is System.Xml.XmlException or InvalidDataException or IOException)
        {
            return false;
        }
    }

    private static readonly System.Xml.XmlReaderSettings RootTagOnly = new() { MaxCharactersInDocument = 1 << 20 };

    private static ZipArchiveEntry? FindMainDocumentEntry(ZipArchive zip)
    {
        var rels = zip.GetEntry("_rels/.rels");
        if (rels is not null)
        {
            try
            {
                using var stream = rels.Open();
                var relsDoc = XDocument.Load(stream);
                XNamespace pr = "http://schemas.openxmlformats.org/package/2006/relationships";
                var target = relsDoc.Root?
                    .Elements(pr + "Relationship")
                    .FirstOrDefault(r =>
                    {
                        var type = (string?)r.Attribute("Type");
                        return type is
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" or
                            "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";
                    })
                    ?.Attribute("Target")?.Value;
                if (target is not null)
                    return zip.GetEntry(target.TrimStart('/'));
            }
            catch (System.Xml.XmlException)
            {
                // fall through to the conventional path
            }
        }
        return zip.GetEntry("word/document.xml");
    }
}

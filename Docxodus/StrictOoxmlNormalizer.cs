#nullable enable

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Normalizes an ISO/IEC 29500 STRICT-conformance WordprocessingML package (namespace family
/// <c>http://purl.oclc.org/ooxml/*</c>, Word's "Strict Open XML Document" save format) into its
/// transitional equivalent. Word performs the same translation transparently on open; Docxodus'
/// XDocument-based pipelines (IR reader, comparers, revision processing) only understand the
/// transitional namespaces, so strict inputs are converted up front. Transitional inputs are
/// returned unchanged (same instance, no copy).
/// </summary>
internal static class StrictOoxmlNormalizer
{
    private const string StrictMarker = "http://purl.oclc.org/ooxml/";
    private static readonly XNamespace StrictW = "http://purl.oclc.org/ooxml/wordprocessingml/main";

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
    /// <paramref name="doc"/> itself (no copy) when it is not a strict package.
    /// </summary>
    internal static WmlDocument NormalizeToTransitional(WmlDocument doc)
    {
        if (!IsStrict(doc))
            return doc;

        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                    !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    continue;

                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();
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
                rewritten = UnfoldGraphicDataNamespaces(rewritten);

                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return new WmlDocument(doc.FileName, ms.ToArray());
    }

    private static readonly XNamespace TransitionalWpDrawing =
        "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace TransitionalA =
        "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XName TransitionalWDrawing =
        XName.Get("drawing", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    private const string Ms2010WpPrefix = "http://schemas.microsoft.com/office/word/2010/wordprocessing";

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
    private static string UnfoldGraphicDataNamespaces(string xml)
    {
        if (!xml.Contains("graphicData", StringComparison.Ordinal))
            return xml;
        XDocument doc;
        try
        {
            doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        }
        catch (System.Xml.XmlException)
        {
            return xml;
        }
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
        if (!changed)
            return xml;
        // StringWriter reports UTF-16, which XDocument.Save stamps into the XML declaration — but
        // the entry is written as UTF-8, making the SDK's reader refuse the part ("There is no
        // Unicode byte order mark"). Declare UTF-8 explicitly.
        using var sw = new Utf8StringWriter();
        doc.Save(sw, SaveOptions.DisableFormatting);
        return sw.ToString();
    }

    private sealed class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }

    /// <summary>
    /// A package is strict when its main document part's root element lives in the strict
    /// WordprocessingML namespace. The main part is resolved through _rels/.rels (either
    /// conformance class's officeDocument relationship type), falling back to the conventional
    /// word/document.xml path.
    /// </summary>
    internal static bool IsStrict(WmlDocument doc)
    {
        try
        {
            using var ms = new MemoryStream(doc.DocumentByteArray, writable: false);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var main = FindMainDocumentEntry(zip);
            if (main is null)
                return false;
            using var stream = main.Open();
            var root = XDocument.Load(stream).Root;
            return root is not null && root.Name.Namespace == StrictW;
        }
        catch (InvalidDataException)
        {
            return false; // not a zip — let the downstream open path produce its own error
        }
        catch (System.Xml.XmlException)
        {
            return false; // malformed main part — same reasoning
        }
    }

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

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>
/// Resolves the small set of malformed or compatibility-markup shapes which Word repairs on open,
/// before DocxDiff reads the package. This keeps the IR reader and the renderer on the same valid
/// OOXML view. The compatibility rules are deliberately conservative:
/// <list type="bullet">
/// <item>A <c>mc:Choice</c> requiring only VML namespaces (Word's strict-save watermark shape,
/// <c>Requires="v"</c>) is unwrapped to its bare <c>w:pict</c> payload — LibreOffice does not
/// render the wrapped form, Word Compare emits it bare.</item>
/// <item>When NO choice is understood (e.g. the obsolete Office 2008/6/28 draft
/// wordprocessingShape namespace), the <c>mc:Fallback</c> content is inlined — Word renders the
/// fallback VML; LibreOffice renders nothing for the original.</item>
/// </list>
/// Modern DrawingML choices (canonical 2010 wps/wpg/wpc) keep their wrapper — every reader
/// understands them and Word Compare preserves them. It also coalesces direct, disjoint duplicate
/// <c>w:pPr</c> elements: Word repairs those into one paragraph-properties element before the
/// paragraph content, whereas leaving the second one after revision runs produces invalid OOXML
/// and layout drift in LibreOffice. Ambiguous or revision-bearing duplicates are left untouched.
/// Untouched documents are returned as the same instance (no copy).
/// </summary>
internal static class MarkupCompatibilityNormalizer
{
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    private static readonly HashSet<string> VmlNamespaces = new(StringComparer.Ordinal)
    {
        "urn:schemas-microsoft-com:vml",
        "urn:schemas-microsoft-com:office:office",
        "urn:schemas-microsoft-com:office:word",
    };

    /// <summary>Namespaces a modern Word build understands in a <c>Requires</c> list. Anything
    /// outside this set (notably pre-release draft namespaces) makes the choice unreadable.</summary>
    private static readonly HashSet<string> UnderstoodNamespaces = new(StringComparer.Ordinal)
    {
        "urn:schemas-microsoft-com:vml",
        "urn:schemas-microsoft-com:office:office",
        "urn:schemas-microsoft-com:office:word",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
        "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
        "http://schemas.microsoft.com/office/word/2010/wordml",
        "http://schemas.microsoft.com/office/word/2012/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml",
        "http://schemas.microsoft.com/office/word/2018/wordml/cex",
        "http://schemas.microsoft.com/office/drawing/2010/main",
        "http://schemas.microsoft.com/office/drawing/2014/main",
    };

    internal static WmlDocument Normalize(WmlDocument doc)
    {
        // Two passes, because almost every document needs no repair at all and the expensive work
        // is proving that. The first pass streams each part looking for the two shapes the repairs
        // react to; it builds no DOM and reads the archive read-only, so it never pays for
        // ZipArchiveMode.Update's entry buffering either. Only a document that has a candidate part
        // reaches the second pass, and only its candidate parts are parsed and rewritten.
        var candidates = FindCandidateParts(doc.DocumentByteArray);
        if (candidates is null)
            return doc;

        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        var anyChanged = false;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!candidates.Contains(entry.FullName))
                    continue;

                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();

                var rewritten = NormalizePart(text);
                if (rewritten is null)
                    continue;

                anyChanged = true;
                using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
                writer.BaseStream.SetLength(0);
                writer.Write(rewritten);
            }
        }
        return anyChanged ? new WmlDocument(doc.FileName, ms.ToArray()) : doc;
    }

    /// <summary>The <c>.xml</c> entries that carry a shape one of the repairs reacts to, or
    /// <c>null</c> when the package carries none.</summary>
    private static HashSet<string>? FindCandidateParts(byte[] package)
    {
        HashSet<string>? candidates = null;
        using var probe = new MemoryStream(package, writable: false);
        using var zip = new ZipArchive(probe, ZipArchiveMode.Read);
        foreach (var entry in zip.Entries)
        {
            if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                continue;
            using var stream = entry.Open();
            if (!CarriesRepairableShape(stream))
                continue;
            (candidates ??= new HashSet<string>(StringComparer.Ordinal)).Add(entry.FullName);
        }

        return candidates;
    }

    /// <summary>
    /// Stream one part and answer the only two questions the repairs ask: is there an
    /// <c>mc:AlternateContent</c> anywhere, and is there a paragraph carrying two or more DIRECT
    /// <c>pPr</c> children. Matching is by local name, which keeps the gate a superset of what the
    /// repairs actually act on (they are namespace-exact) — a false positive costs one parse of one
    /// part and still returns it unchanged, whereas a false negative would be a correctness bug.
    /// Malformed XML answers "no", which is what <see cref="NormalizePart"/> concludes anyway.
    /// </summary>
    private static bool CarriesRepairableShape(Stream part)
    {
        // Per-depth view of the open element stack: what opened at each depth, and — for a depth
        // holding a paragraph — how many direct pPr children it has seen so far. Opening a new
        // element at a depth resets that depth's count, so sibling paragraphs never pool.
        var nameAtDepth = new List<string>();
        var pPrAtDepth = new List<int>();

        try
        {
            using var reader = XmlReader.Create(part, StreamingProbeSettings);
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element)
                    continue;

                var name = reader.LocalName;
                if (name == "AlternateContent")
                    return true;

                var depth = reader.Depth;
                while (nameAtDepth.Count <= depth)
                {
                    nameAtDepth.Add(string.Empty);
                    pPrAtDepth.Add(0);
                }

                nameAtDepth[depth] = name;
                pPrAtDepth[depth] = 0;

                if (name == "pPr" && depth > 0 && nameAtDepth[depth - 1] == "p"
                    && ++pPrAtDepth[depth - 1] > 1)
                    return true;
            }
        }
        catch (System.Xml.XmlException)
        {
            return false;
        }

        return false;
    }

    private static readonly XmlReaderSettings StreamingProbeSettings = new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        IgnoreComments = true,
        IgnoreProcessingInstructions = true,
        IgnoreWhitespace = true,
        CloseInput = false,
    };

    /// <summary>Returns rewritten part XML, or null when no conservative repair was applicable.</summary>
    private static string? NormalizePart(string xml)
    {
        XDocument doc;
        try
        {
            doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        }
        catch (System.Xml.XmlException)
        {
            return null;
        }

        var changed = ResolveAlternateContent(doc);
        changed |= CoalesceDisjointDuplicateParagraphProperties(doc);
        if (!changed)
            return null;

        using var sw = new Utf8StringWriter();
        doc.Save(sw, SaveOptions.DisableFormatting);
        return sw.ToString();
    }

    /// <summary>Resolve supported <c>mc:AlternateContent</c> wrappers in an already parsed part.</summary>
    private static bool ResolveAlternateContent(XDocument doc)
    {
        var changed = false;
        foreach (var ac in doc.Descendants(Mc + "AlternateContent").ToList())
        {
            var selected = ac.Elements(Mc + "Choice")
                .FirstOrDefault(c => RequiredNamespaces(c).All(UnderstoodNamespaces.Contains));
            if (selected is not null)
            {
                // Only VML-only choices are unwrapped; modern DrawingML wrappers stay.
                var required = RequiredNamespaces(selected).ToList();
                if (required.Count == 0 || !required.All(VmlNamespaces.Contains))
                    continue;
                ac.ReplaceWith(selected.Nodes());
                changed = true;
            }
            else
            {
                var fallback = ac.Element(Mc + "Fallback");
                if (fallback is null)
                    continue;
                ac.ReplaceWith(fallback.Nodes());
                changed = true;
            }
        }
        return changed;
    }

    /// <summary>
    /// Repair only the unambiguous duplicate-<c>w:pPr</c> shape Word coalesces. A group is safe
    /// when every direct property child has a distinct QName and its attributes do not conflict.
    /// Property-change/revision markup is intentionally excluded: its before/after semantics make
    /// a mechanical merge lossy. The merged properties are put back in schema order before every
    /// paragraph content child, which is important when a comparer has already inserted revision
    /// runs between the original malformed property elements.
    /// </summary>
    private static bool CoalesceDisjointDuplicateParagraphProperties(XDocument doc)
    {
        var changed = false;
        foreach (var paragraph in doc.Descendants(W.p).ToList())
        {
            var properties = paragraph.Elements(W.pPr).ToList();
            if (properties.Count < 2 || !CanCoalesce(properties))
                continue;

            var attributes = new Dictionary<XName, string>();
            var children = new List<XElement>();
            foreach (var propertiesElement in properties)
            {
                foreach (var attribute in propertiesElement.Attributes())
                    if (!attribute.IsNamespaceDeclaration)
                        attributes[attribute.Name] = attribute.Value;
                children.AddRange(propertiesElement.Elements().Select(e => new XElement(e)));
            }

            var merged = new XElement(
                W.pPr,
                attributes.Select(a => new XAttribute(a.Key, a.Value)),
                children);
            merged = (XElement)WordprocessingMLUtil.WmlOrderElementsPerStandard(merged);

            // The source's first pPr can itself be misplaced. AddFirst deliberately repairs that
            // too, then remove every original pPr after the clone is safely attached.
            paragraph.AddFirst(merged);
            foreach (var propertiesElement in properties)
                propertiesElement.Remove();
            changed = true;
        }
        return changed;
    }

    private static bool CanCoalesce(IReadOnlyCollection<XElement> properties)
    {
        var seenChildren = new HashSet<XName>();
        var attributes = new Dictionary<XName, string>();
        foreach (var propertiesElement in properties)
        {
            // pPrChange contains a previous pPr snapshot; revision-bearing rPr has similarly
            // nontrivial history semantics. Do not guess which state should win.
            if (propertiesElement.Descendants().Any(IsRevisionMarkup))
                return false;

            foreach (var child in propertiesElement.Elements())
                if (!seenChildren.Add(child.Name))
                    return false;

            foreach (var attribute in propertiesElement.Attributes())
            {
                if (attribute.IsNamespaceDeclaration)
                    continue;
                if (attributes.TryGetValue(attribute.Name, out var prior) && prior != attribute.Value)
                    return false;
                attributes[attribute.Name] = attribute.Value;
            }
        }
        return true;
    }

    private static bool IsRevisionMarkup(XElement element) =>
        element.Name == W.pPrChange ||
        element.Name == W.rPrChange ||
        element.Name == W.ins ||
        element.Name == W.del ||
        element.Name == W.moveFrom ||
        element.Name == W.moveTo;

    /// <summary>The namespaces a choice's <c>Requires</c> prefix list resolves to in scope.
    /// An unresolvable prefix yields an empty marker that never matches the understood set.</summary>
    private static IEnumerable<string> RequiredNamespaces(XElement choice)
    {
        var requires = (string?)choice.Attribute("Requires");
        if (string.IsNullOrWhiteSpace(requires))
            yield break;
        foreach (var prefix in requires.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            yield return choice.GetNamespaceOfPrefix(prefix)?.NamespaceName ?? string.Empty;
    }

    private sealed class Utf8StringWriter : StringWriter
    {
        public override Encoding Encoding => Encoding.UTF8;
    }
}

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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
        using var ms = new MemoryStream();
        ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
        var anyChanged = false;
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries.ToList())
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    continue;

                string text;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    text = reader.ReadToEnd();
                // Most parts need no XML parse. A literal pPr check is deliberately broad enough
                // to cover nonstandard Word prefixes too; a harmless false positive only parses
                // the part and still returns it unchanged.
                if (!text.Contains("AlternateContent", StringComparison.Ordinal) &&
                    !text.Contains("pPr", StringComparison.Ordinal))
                    continue;

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

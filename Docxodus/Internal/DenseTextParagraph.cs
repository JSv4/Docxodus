#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>
/// A formatting template for a dense, explicitly formatted ASCII paragraph.
/// The ordinary converter resolves the paragraph and each distinct run format;
/// expansion then restores the original text and breaks. The live OOXML is
/// never abbreviated. This avoids resolving the same handful of formats many
/// thousands of times when incrementally rendering a text picture.
/// </summary>
internal sealed class DenseTextParagraph
{
    private readonly List<(int Format, XElement[] Content)> _runs = new();
    private readonly List<XElement> _formats = new();
    private readonly string _prefix = "DXTEXT";

    internal XElement Template { get; private set; } = null!;

    internal static DenseTextParagraph? TryCompact(XElement paragraph)
    {
        if (paragraph.Name != W.p || paragraph.Elements(W.r).Take(128).Count() < 128)
            return null;
        var pPr = paragraph.Element(W.pPr);
        if ((string?)pPr?.Element(W.spacing)?.Attribute(W.lineRule) != "exact"
            || pPr.Elements().Any(e => e.Name != W.spacing && e.Name != W.shd)
            || paragraph.Elements().Any(e => e.Name != W.pPr && e.Name != W.r)) return null;

        var result = new DenseTextParagraph();
        var keys = new Dictionary<string, int>(StringComparer.Ordinal);
        XElement? geometry = null;
        foreach (var run in paragraph.Elements(W.r))
        {
            var properties = run.Element(W.rPr);
            if (properties is null) return null;
            int seen = 0;
            string? key = null;
            foreach (var p in properties.Elements())
            {
                int flag = p.Name == W.rFonts ? 1 : p.Name == W.sz ? 2 : p.Name == W.color ? 4
                    : p.Name == W.szCs ? 8 : p.Name == W.b ? 16 : p.Name == W.bCs ? 32
                    : p.Name == W.spacing ? 64 : 0;
                if (flag == 0 || (seen & flag) != 0 || p.HasElements || p.Value.Length != 0) return null;
                seen |= flag;
                if (p.Name == W.color)
                {
                    if (p.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != PtOpenXml.Unid && a.Name != W.val))
                        return null;
                    key = (string?)p.Attribute(W.val);
                }
            }
            if ((seen & 7) != 7 || key is null) return null;
            geometry ??= properties;
            if (!SameGeometry(geometry, properties)) return null;
            var content = run.Elements().Where(e => e.Name != W.rPr).ToArray();
            if (content.Length == 0 || content.Any(e => e.Name != W.t && e.Name != W.br)) return null;
            if (content.Any(e => e.Name == W.t && !IsAscii(e.Value)
                || e.Name == W.br && e.Attributes().Any(a => a.Name != PtOpenXml.Unid))) return null;
            if (!keys.TryGetValue(key, out int index))
            {
                var format = new XElement(properties);
                foreach (var element in format.DescendantsAndSelf()) element.Attribute(PtOpenXml.Unid)?.Remove();
                index = keys.Count;
                if (index >= 256) return null;
                keys.Add(key, index);
                result._formats.Add(format);
            }
            result._runs.Add((index, content));
        }

        if (result._runs.Count < result._formats.Count * 4) return null;
        // The picture changes which ink occurs first. Keep the formatting
        // template stable so a later frame can reuse its resolved styles.
        var order = keys.OrderBy(k => k.Key, StringComparer.Ordinal).Select(k => k.Value).ToArray();
        var remap = new int[order.Length];
        var sorted = order.Select(i => result._formats[i]).ToArray();
        for (int i = 0; i < order.Length; i++) remap[order[i]] = i;
        result._formats.Clear();
        result._formats.AddRange(sorted);
        for (int i = 0; i < result._runs.Count; i++)
            result._runs[i] = (remap[result._runs[i].Format], result._runs[i].Content);
        result.Template = new XElement(paragraph.Name, paragraph.Attributes(), pPr);
        foreach (var element in result.Template.Descendants()) element.Attribute(PtOpenXml.Unid)?.Remove();
        for (int i = 0; i < result._formats.Count; i++)
            result.Template.Add(new XElement(W.r, result._formats[i],
                new XElement(W.t, result.Marker(i)), new XElement(W.br)));
        return result;
    }

    private static bool IsAscii(string text)
    {
        foreach (char c in text) if (c < ' ' || c > '~') return false;
        return true;
    }

    // All accepted runs have the same geometry. Compare the small property
    // trees directly, ignoring bookkeeping IDs, instead of constructing and
    // hashing a long formatting string for every run in every frame.
    private static bool SameGeometry(XElement a, XElement b)
    {
        using var left = a.Elements().Where(e => e.Name != W.color).GetEnumerator();
        using var right = b.Elements().Where(e => e.Name != W.color).GetEnumerator();
        while (left.MoveNext())
        {
            if (!right.MoveNext() || left.Current.Name != right.Current.Name) return false;
            var x = left.Current.FirstAttribute;
            var y = right.Current.FirstAttribute;
            while (true)
            {
                while (x is not null && (x.IsNamespaceDeclaration || x.Name == PtOpenXml.Unid)) x = x.NextAttribute;
                while (y is not null && (y.IsNamespaceDeclaration || y.Name == PtOpenXml.Unid)) y = y.NextAttribute;
                if (x is null || y is null) { if (x != y) return false; break; }
                if (x.Name != y.Name || x.Value != y.Value) return false;
                x = x.NextAttribute; y = y.NextAttribute;
            }
        }
        return !right.MoveNext();
    }

    private string Marker(int index) => _prefix + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "END";

    internal bool TryExpand(XElement html)
    {
        var templates = new XElement[_formats.Count];
        for (int i = 0; i < templates.Length; i++)
        {
            var marker = Marker(i);
            var span = html.Descendants().FirstOrDefault(e => e.Name.LocalName == "span"
                && e.Nodes().OfType<XText>().Any(t => t.Value == marker));
            if (span is null || span.Parent != html) return false;
            templates[i] = span;
        }
        var expanded = new List<XElement>(_runs.Count);
        foreach (var (format, content) in _runs)
        {
            var span = new XElement(templates[format].Name, templates[format].Attributes());
            foreach (var node in content)
            {
                if (node.Name == W.br) span.Add(new XElement(Xhtml.br));
                else span.Add(new XText(node.Value.Replace(' ', '\u00a0')));
            }
            expanded.Add(span);
        }
        html.ReplaceNodes(expanded);
        return true;
    }
}

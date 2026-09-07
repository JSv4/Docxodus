#nullable enable
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        string? geometry = null;
        var keyBuilder = new StringBuilder(256);
        foreach (var run in paragraph.Elements(W.r))
        {
            var properties = run.Element(W.rPr);
            if (properties is null || properties.Element(W.rFonts) is null
                || properties.Element(W.sz) is null || properties.Element(W.color) is null)
                return null;
            if (properties.Elements().Any(e => e.Name != W.rFonts && e.Name != W.sz
                && e.Name != W.szCs && e.Name != W.color && e.Name != W.b && e.Name != W.bCs
                && e.Name != W.spacing)) return null;
            var content = run.Elements().Where(e => e.Name != W.rPr).ToArray();
            if (content.Length == 0 || content.Any(e => e.Name != W.t && e.Name != W.br)) return null;
            if (content.Any(e => e.Name == W.t && e.Value.Any(c => c < ' ' || c > '~')
                || e.Name == W.br && e.Attributes().Any(a => a.Name != PtOpenXml.Unid))) return null;

            // A formatting key ignores bookkeeping IDs without cloning and
            // serializing an entire property subtree for every repeated run.
            keyBuilder.Clear();
            foreach (var property in properties.Elements())
            {
                keyBuilder.Append(property.Name).Append('[');
                foreach (var attr in property.Attributes())
                    if (!attr.IsNamespaceDeclaration && attr.Name != PtOpenXml.Unid)
                        keyBuilder.Append(attr.Name).Append('=').Append(attr.Value.Length)
                            .Append(':').Append(attr.Value);
                keyBuilder.Append(']');
            }
            string key = keyBuilder.ToString();
            if (!keys.TryGetValue(key, out int index))
            {
                var format = new XElement(properties);
                foreach (var element in format.DescendantsAndSelf()) element.Attribute(PtOpenXml.Unid)?.Remove();
                var shape = new XElement(format);
                shape.Element(W.color)?.Remove();
                string shapeKey = shape.ToString(SaveOptions.DisableFormatting);
                geometry ??= shapeKey;
                if (shapeKey != geometry) return null;
                index = keys.Count;
                if (index >= 256) return null;
                keys.Add(key, index);
                result._formats.Add(format);
            }
            result._runs.Add((index, content));
        }

        if (result._runs.Count < result._formats.Count * 4) return null;
        result.Template = new XElement(paragraph.Name, paragraph.Attributes(), pPr);
        for (int i = 0; i < result._formats.Count; i++)
            result.Template.Add(new XElement(W.r, result._formats[i],
                new XElement(W.t, result.Marker(i)), new XElement(W.br)));
        return result;
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

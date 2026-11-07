using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>Original body boundaries for the Word-confirmed automatic paragraph gaps.</summary>
internal sealed class AutomaticParagraphSpacing
{
    private static readonly XNamespace Rendering = "urn:docxodus:rendering";
    private static readonly XName Before = Rendering + "autoBefore";
    private static readonly XName After = Rendering + "autoAfter";
    private readonly WordprocessingDocument _document;
    private readonly XElement? _body;
    private readonly bool _enabled;
    private readonly Dictionary<XElement, (bool Eligible, int? List)> _facts = new();

    internal AutomaticParagraphSpacing(WordprocessingDocument document)
    {
        _document = document;
        var main = document.MainDocumentPart!;
        _body = main.GetXDocument().Root?.Element(W.body);
        var compatibility = main.DocumentSettingsPart?.GetXDocument().Root?
            .Element(W.compat)?.Element(W.doNotUseHTMLParagraphAutoSpacing);
        // First/last boundaries, other stories, revisions and compatibility-mode spacing
        // still need separate Word measurements. Do not infer those from a shell's body.
        _enabled = !(compatibility != null && compatibility.Attribute(W.val).ToBoolean() != false)
            && !RevisionAccepter.HasTrackedRevisions(document)
            && new[] { main.GetXDocument(), main.StyleDefinitionsPart?.GetXDocument(),
                main.NumberingDefinitionsPart?.GetXDocument() }.Any(x => x?.Descendants(W.spacing)
                    .Any(s => s.Attribute(W.beforeAutospacing).ToBoolean() == true
                        || s.Attribute(W.afterAutospacing).ToBoolean() == true) == true);
    }

    internal void Stamp(XElement source, XElement clone)
    {
        if (!_enabled && ReferenceEquals(source, clone)) return;
        using var originals = source.DescendantsAndSelf(W.p).GetEnumerator();
        using var copies = clone.DescendantsAndSelf(W.p).GetEnumerator();
        while (originals.MoveNext() && copies.MoveNext())
        {
            // Block clones already carry their source facts, including excluded edges.
            if (copies.Current.Attribute(Before) != null) continue;
            copies.Current.SetAttributeValue(Before, Gap(originals.Current, before: true));
            copies.Current.SetAttributeValue(After, Gap(originals.Current, before: false));
        }
    }

    private string Gap(XElement paragraph, bool before)
    {
        var own = Facts(paragraph);
        if (!own.Eligible) return "-";
        XNode? node = before ? paragraph.PreviousNode : paragraph.NextNode;
        while (node != null && node is not XElement) node = before ? node.PreviousNode : node.NextNode;
        if (node is not XElement neighbor || !Facts(neighbor).Eligible) return "-";
        return own.List is { } id && id == Facts(neighbor).List ? "0" : "280";
    }

    private (bool Eligible, int? List) Facts(XElement paragraph)
    {
        if (_facts.TryGetValue(paragraph, out var facts)) return facts;
        if (!_enabled || paragraph.Name != W.p || paragraph.Parent != _body) return (false, null);
        var properties = FormattingAssembler.ResolveEffectiveParagraphProperties(_document, paragraph);
        var listParagraph = paragraph;
        if (_document.MainDocumentPart?.NumberingDefinitionsPart is { } numbering)
        {
            if (_document.MainDocumentPart.StyleDefinitionsPart != null)
                ListItemRetriever.RetrieveListItem(_document, paragraph);
            else
            {
                // styles.xml is optional. Resolve its empty cascade without adding a part
                // or leaving partially initialized numbering annotations on the live source.
                listParagraph = new XElement(W.p, paragraph.Element(W.pPr));
                ListItemRetriever.InitListItemInfo(numbering.GetXDocument(), new XDocument(new XElement(W.styles)), listParagraph);
            }
        }
        var item = listParagraph.Annotation<ListItemRetriever.ListItemInfo>();
        var list = item is { IsListItem: true, IsZeroNumId: false } && item.NumId > 0 ? item : null;
        var levelProperties = list?.Lvl(ListItemRetriever.GetEffectiveLevel(listParagraph))?.Element(W.pPr);
        static bool Boundary(XElement? pPr) => pPr?.Elements().Any(e =>
            e.Name == W.framePr || e.Name == W.pBdr || e.Name == W.sectPr) == true;
        facts = (!Boundary(properties) && !Boundary(levelProperties), list?.NumId);
        _facts.Add(paragraph, facts);
        return facts;
    }

    internal static void Apply(XElement root)
    {
        foreach (var paragraph in root.DescendantsAndSelf(W.p))
        {
            var spacing = paragraph.Element(W.pPr)?.Element(W.spacing);
            if (spacing == null) continue;
            ApplySide(Before, W.beforeAutospacing, W.before);
            ApplySide(After, W.afterAutospacing, W.after);
            void ApplySide(XName context, XName automatic, XName value)
            {
                var gap = (string?)paragraph.Attribute(context);
                if (gap is "0" or "280" && spacing.Attribute(automatic).ToBoolean() == true)
                    spacing.SetAttributeValue(value, gap);
            }
        }
    }
}

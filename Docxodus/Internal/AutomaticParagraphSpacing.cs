using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>Word's spacing between adjacent paragraphs, resolved from each paragraph's original neighbors.</summary>
internal sealed class AutomaticParagraphSpacing
{
    private static readonly XNamespace Rendering = "urn:docxodus:rendering";
    private static readonly XName Before = Rendering + "autoBefore";
    private static readonly XName After = Rendering + "autoAfter";
    private static readonly XName Added = Rendering + "addedBefore";
    private readonly WordprocessingDocument _document;
    private readonly XElement? _body;
    // Word gives the automatic flags their 14 pt, or, under doNotUseHTMLParagraphAutoSpacing,
    // ignores the flags and adds space after to the next paragraph's space before.
    private readonly bool _automatic;
    private readonly bool _additive;
    private readonly Dictionary<XElement, (bool Eligible, int? List, decimal After)> _facts = new();
    private readonly Dictionary<XElement, (XElement? Previous, XElement? Next)> _neighbors = new();
    private readonly HashSet<XElement> _stories = new();
    private readonly ParagraphLineGrid? _lineGrid;
    private List<XElement>? _sections;

    internal AutomaticParagraphSpacing(WordprocessingDocument document)
    {
        _document = document;
        var main = document.MainDocumentPart!;
        _body = main.GetXDocument().Root?.Element(W.body);
        var compatibility = main.DocumentSettingsPart?.GetXDocument().Root?
            .Element(W.compat)?.Element(W.doNotUseHTMLParagraphAutoSpacing);
        _additive = compatibility != null && compatibility.Attribute(W.val).ToBoolean() != false;
        _lineGrid = _additive ? new ParagraphLineGrid(document) : null;
        _automatic = !_additive
            && new[] { main.GetXDocument(), main.StyleDefinitionsPart?.GetXDocument(),
                main.NumberingDefinitionsPart?.GetXDocument() }.Any(x => x?.Descendants(W.spacing)
                    .Any(s => s.Attribute(W.beforeAutospacing).ToBoolean() == true
                        || s.Attribute(W.afterAutospacing).ToBoolean() == true) == true);
    }

    internal void Stamp(XElement source, XElement clone)
    {
        if (!_automatic && !_additive && ReferenceEquals(source, clone)) return;
        using var originals = source.DescendantsAndSelf(W.p).GetEnumerator();
        using var copies = clone.DescendantsAndSelf(W.p).GetEnumerator();
        while (originals.MoveNext() && copies.MoveNext())
        {
            // Block clones already carry their source facts, including excluded edges.
            if (copies.Current.Attribute(Before) != null) continue;
            copies.Current.SetAttributeValue(Before, Gap(originals.Current, before: true));
            copies.Current.SetAttributeValue(After, Gap(originals.Current, before: false));
            copies.Current.SetAttributeValue(Added, AddedBefore(originals.Current));
        }
    }

    private string Gap(XElement paragraph, bool before)
    {
        if (!_automatic || Neighbors(paragraph) is not (var previous, var next)) return "-";
        // Word prints no automatic space at the start or end of the body or a table cell.
        if ((before ? previous : next) is not { } neighbor) return "0";
        if (neighbor.Name == W.tbl)
            return neighbor.Element(W.tblPr)?.Element(W.tblpPr) == null
                && !neighbor.Descendants(W.trPr).Elements().Any(e => e.Name == W.ins || e.Name == W.del) ? "280" : "-";
        if (!Facts(neighbor).Eligible) return "-";
        return Facts(paragraph).List is { } id && id == Facts(neighbor).List ? "0" : "280";
    }

    // The previous paragraph's space after, which Word adds to this paragraph's space before. A
    // section break onto a new page leaves only this paragraph's space before at the top of it.
    private string? AddedBefore(XElement paragraph) =>
        _additive && Neighbors(paragraph)?.Previous is { } previous && Facts(previous).Eligible
            && !StartsPage(previous) ? Facts(previous).After.ToString(CultureInfo.InvariantCulture) : null;

    private bool StartsPage(XElement paragraph)
    {
        if (paragraph.Element(W.pPr)?.Element(W.sectPr) is not { } end) return false;
        _sections ??= _body?.Descendants(W.sectPr).Where(s => s.Parent?.Name != W.sectPrChange).ToList() ?? new();
        var following = _sections.ElementAtOrDefault(_sections.IndexOf(end) + 1);
        return (string?)following?.Element(W.type)?.Attribute(W.val) != "continuous";
    }

    private (XElement? Previous, XElement? Next)? Neighbors(XElement paragraph)
    {
        if (Story(paragraph) is not { } story || !Facts(paragraph).Eligible) return null;
        if (_stories.Add(story))
        {
            var blocks = Blocks(story).ToList();
            for (int i = 0; i < blocks.Count; i++)
                if (blocks[i].Name == W.p)
                    _neighbors[blocks[i]] = (i > 0 ? blocks[i - 1] : null, i + 1 < blocks.Count ? blocks[i + 1] : null);
        }
        return _neighbors.TryGetValue(paragraph, out var neighbors) ? neighbors : null;
    }

    // The body or table cell whose block sequence holds the paragraph. Content controls and
    // custom XML are transparent; text boxes, notes and other stories are not measured.
    private XElement? Story(XElement paragraph)
    {
        var parent = paragraph.Parent;
        while (parent != null && (parent.Name == W.sdtContent || parent.Name == W.sdt || parent.Name == W.customXml))
            parent = parent.Parent;
        return parent != null && (parent == _body || parent.Name == W.tc && parent.Document == _body?.Document
            && !parent.Ancestors(W.txbxContent).Any()) ? parent : null;
    }

    private static IEnumerable<XElement> Blocks(XElement container) => container.Elements().SelectMany(e =>
        e.Name == W.p || e.Name == W.tbl ? new[] { e }
        : e.Name == W.sdt ? e.Elements(W.sdtContent).SelectMany(Blocks)
        : e.Name == W.customXml ? Blocks(e) : Enumerable.Empty<XElement>());

    private (bool Eligible, int? List, decimal After) Facts(XElement paragraph)
    {
        if (_facts.TryGetValue(paragraph, out var facts)) return facts;
        if (!_automatic && !_additive || paragraph.Name != W.p) return default;
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
        // Borders and section breaks keep the automatic gap. Frames sit outside the text flow,
        // and a tracked paragraph mark joins or splits paragraphs, so neither has one neighbor.
        var mark = paragraph.Element(W.pPr)?.Element(W.rPr);
        var revised = mark?.Elements().Any(e => e.Name == W.ins || e.Name == W.del
            || e.Name == W.moveFrom || e.Name == W.moveTo) == true;
        var spacing = properties?.Element(W.spacing);
        facts = (!revised && properties?.Element(W.framePr) == null && levelProperties?.Element(W.framePr) == null,
            list?.NumId, ParagraphLineGrid.Spacing(spacing, W.afterLines,
                WordprocessingMLUtil.AttributeToDecimalTwips(spacing?.Attribute(W.after)), _lineGrid?.UnitOf(paragraph)) ?? 0);
        _facts.Add(paragraph, facts);
        return facts;
    }

    internal static void Apply(XElement root)
    {
        foreach (var paragraph in root.DescendantsAndSelf(W.p))
        {
            var spacing = paragraph.Element(W.pPr)?.Element(W.spacing);
            if ((decimal?)paragraph.Attribute(Added) is { } added)
            {
                // In a cell a table style can set the space after, which the cascade read before
                // assembly lacks, so a cell rendered with the paragraph above reads that instead.
                if (paragraph.Parent?.Name == W.tc
                    && paragraph.ElementsBeforeSelf().LastOrDefault(WmlToHtmlConverter.HoldsBlock) is { } rendered)
                {
                    var above = rendered.Element(W.pPr)?.Element(W.spacing);
                    added = ParagraphLineGrid.Spacing(above, W.afterLines,
                        WordprocessingMLUtil.AttributeToDecimalTwips(above?.Attribute(W.after))) ?? 0;
                }
                if (added > 0 && paragraph.Element(W.pPr) is { } properties)
                {
                    if (spacing == null) properties.Add(spacing = new XElement(W.spacing));
                    // Every renderer collapses adjacent margins to the larger one, so a space
                    // before that already holds the space after above it prints Word's sum.
                    spacing.SetAttributeValue(W.before, added + (ParagraphLineGrid.Spacing(spacing, W.beforeLines,
                        WordprocessingMLUtil.AttributeToDecimalTwips(spacing.Attribute(W.before))) ?? 0));
                    spacing.SetAttributeValue(W.beforeLines, null);
                }
            }
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

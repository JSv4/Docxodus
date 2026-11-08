using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>The original main-body section grid, before block shells or conversion rewrites.</summary>
internal sealed class ParagraphLineGrid
{
    private static readonly XName Pitch = XNamespace.Get("urn:docxodus:rendering") + "lineGrid";
    private readonly Dictionary<XElement, string> _pitches = new();

    internal ParagraphLineGrid(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (body == null || !body.Descendants(W.docGrid).Any(g => (string?)g.Attribute(W.type) == "lines")
            || RevisionAccepter.HasTrackedRevisions(document)) return;
        XElement? section = null;
        // A break belongs to the section it ends. One reverse walk avoids a forward
        // governing-section search for every paragraph in a long document.
        foreach (var block in body.Elements().Reverse())
        {
            section = BlockMetadataOps.SectionBreakIn(block) ?? section;
            var grid = section?.Element(W.docGrid);
            var pitch = WordprocessingMLUtil.AttributeToTwips(grid?.Attribute(W.linePitch));
            if (block.Name == W.p && (string?)grid?.Attribute(W.type) == "lines" && pitch > 0)
                _pitches[block] = (pitch.Value / 20m).ToString(CultureInfo.InvariantCulture) + "pt";
        }
    }

    internal void Stamp(XElement source, XElement clone)
    {
        if (_pitches.Count == 0 && ReferenceEquals(source, clone)) return;
        using var originals = source.DescendantsAndSelf(W.p).GetEnumerator();
        using var copies = clone.DescendantsAndSelf(W.p).GetEnumerator();
        while (originals.MoveNext() && copies.MoveNext())
            if (copies.Current.Attribute(Pitch) == null)
                copies.Current.SetAttributeValue(Pitch, _pitches.GetValueOrDefault(originals.Current, "0"));
    }

    internal static void AddStyle(XElement paragraph, Dictionary<string, string> style)
    {
        var pitch = (string?)paragraph.Attribute(Pitch);
        var properties = paragraph.Element(W.pPr);
        var spacing = properties?.Element(W.spacing);
        if (pitch == null || pitch == "0" || properties?.Element(W.snapToGrid)?.Attribute(W.val).ToBoolean() == false
            || ((string?)spacing?.Attribute(W.lineRule) ?? "auto") != "auto"
            || (WordprocessingMLUtil.AttributeToTwips(spacing?.Attribute(W.line)) ?? 240m) != 240m
            || properties?.Elements().Any(e => e.Name == W.framePr || e.Name == W.pBdr) == true
            || paragraph.Descendants().Attributes(PtOpenXml.ListItemRun).Any()) return;
        style["--docx-line-grid"] = pitch;
    }
}

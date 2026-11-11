using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Docxodus.Internal;

/// <summary>
/// The original main-body section grid, before block shells or conversion rewrites. Only body
/// paragraphs carry it: Word 2024 snaps no header or footer line and counts their line-count
/// spacing in 12pt lines, whatever the section's grid (#1551).
/// </summary>
internal sealed class ParagraphLineGrid
{
    private static readonly XName Pitch = XNamespace.Get("urn:docxodus:rendering") + "lineGrid";
    // The pitch that measures line-count spacing, which also holds in cells and unsnapped paragraphs.
    private static readonly XName Unit = XNamespace.Get("urn:docxodus:rendering") + "lineUnit";
    private readonly Dictionary<XElement, string> _pitches = new();

    internal ParagraphLineGrid(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (body == null || !body.Descendants(W.docGrid).Any(SnapsLines)) return;
        XElement? section = null;
        // A break belongs to the section it ends. One reverse walk avoids a forward
        // governing-section search for every paragraph in a long document.
        foreach (var block in body.Elements().Reverse())
        {
            section = BlockMetadataOps.SectionBreakIn(block) ?? section;
            var grid = section?.Element(W.docGrid);
            var pitch = WordprocessingMLUtil.AttributeToTwips(grid?.Attribute(W.linePitch));
            if (SnapsLines(grid) && pitch > 0)
                _pitches[block] = (pitch.Value / 20m).ToString(CultureInfo.InvariantCulture) + "pt";
        }
    }

    /// <summary>Word puts lines on the pitch for each grid type but <c>default</c>, the type an absent attribute means.</summary>
    private static bool SnapsLines(XElement? grid) =>
        (string?)grid?.Attribute(W.type) is "lines" or "linesAndChars" or "snapToChars";

    internal void Stamp(XElement source, XElement clone)
    {
        if (_pitches.Count == 0 && ReferenceEquals(source, clone)) return;
        using var originals = source.DescendantsAndSelf(W.p).GetEnumerator();
        using var copies = clone.DescendantsAndSelf(W.p).GetEnumerator();
        while (originals.MoveNext() && copies.MoveNext())
        {
            if (copies.Current.Attribute(Pitch) != null) continue;
            var block = originals.Current.AncestorsAndSelf().FirstOrDefault(e => e.Parent?.Name == W.body);
            var pitch = block == null ? null : _pitches.GetValueOrDefault(block);
            // Only a body paragraph snaps: Word never snaps cell text (compatibility mode 15).
            copies.Current.SetAttributeValue(Pitch, block == originals.Current ? pitch ?? "0" : "0");
            copies.Current.SetAttributeValue(Unit, pitch);
        }
    }

    /// <summary>
    /// Spacing in hundredths of a line (<c>w:beforeLines</c>, <c>w:afterLines</c>) replaces its twip
    /// value. Word 2024 counts that line as the section's grid pitch, or 12pt without a line grid,
    /// whatever the font, also in cells and unsnapped paragraphs; a zero count keeps the twips (#1551).
    /// </summary>
    internal static decimal? Spacing(XElement? spacing, XName lines, decimal? twips) =>
        Spacing(spacing, lines, twips, (string?)spacing?.Parent?.Parent?.Attribute(Unit));

    /// <summary>The same, for a source paragraph's resolved spacing, with <see cref="UnitOf"/>'s pitch.</summary>
    internal static decimal? Spacing(XElement? spacing, XName lines, decimal? twips, string? pitch)
    {
        if (!int.TryParse((string?)spacing?.Attribute(lines), NumberStyles.Integer, CultureInfo.InvariantCulture,
            out var hundredths) || hundredths <= 0) return twips;
        return hundredths * (pitch == null ? 12m : Points(pitch)) / 5m;
    }

    private static decimal Points(string pitch) => decimal.Parse(pitch[..^2], CultureInfo.InvariantCulture);

    /// <summary>The pitch a source paragraph's line-count spacing counts, or null for 12pt.</summary>
    internal string? UnitOf(XElement paragraph) =>
        paragraph.AncestorsAndSelf().FirstOrDefault(e => e.Parent?.Name == W.body) is { } block
            ? _pitches.GetValueOrDefault(block) : null;

    /// <summary>
    /// Word 2024 gives an <c>atLeast</c> line <c>max(whole grid lines, minimum)</c> and ignores the grid
    /// for an <c>exact</c> one (#1551). An at-least line keeps the converter's line height for its
    /// minimum: that is the line wherever the grid pass doesn't measure the text's grid lines.
    /// </summary>
    internal static void AddStyle(XElement paragraph, Dictionary<string, string> style)
    {
        var pitch = (string?)paragraph.Attribute(Pitch);
        var properties = paragraph.Element(W.pPr);
        if (pitch == null || pitch == "0" || properties?.Element(W.snapToGrid)?.Attribute(W.val).ToBoolean() == false
            || (string?)properties?.Element(W.spacing)?.Attribute(W.lineRule) is not (null or "auto" or "atLeast")
            || properties?.Elements().Any(e => e.Name == W.framePr || e.Name == W.pBdr) == true
            || paragraph.Descendants().Attributes(PtOpenXml.ListItemRun).Any()) return;
        style["--docx-line-grid"] = pitch;
    }
}

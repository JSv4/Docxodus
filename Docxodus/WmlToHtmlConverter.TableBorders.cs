using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Docxodus.Internal;

namespace Docxodus;

public static partial class WmlToHtmlConverter
{
    private sealed record TableBorderEdge(string Side, int Start, int End, XElement Border)
    {
        public decimal Width => Math.Clamp((int?)Border.Attribute(W.sz) ?? 0, 2, 96) / 8m;
    }

    private sealed class ExactTableCellBorders
    {
        public required XElement Cell { get; init; }
        public required int FirstRow { get; init; }
        public required int LastRow { get; init; }
        public required int Start { get; init; }
        public required int End { get; init; }
        public decimal Left { get; set; }
        public decimal Right { get; set; }
        public decimal Top { get; set; }
        public decimal Bottom { get; set; }
    }

    private sealed record TableBorderSlot(XElement Source, ExactTableCellBorders Owner, int Start, int End);

    private sealed class ExactTableBorders
    {
        public Dictionary<XElement, ExactTableCellBorders> Cells { get; } = new();
        public Dictionary<XElement, List<TableBorderEdge>> Rows { get; } = new();
        public Dictionary<XElement, List<TableBorderEdge>> HeaderJoins { get; } = new();
        public decimal Top { get; set; }
        public decimal Bottom { get; set; }
    }

    // Resolve segments before the legacy whole-side pass can spread a winner across a merged edge.
    // The grid contains physical intervals, never an allocation proportional to an authored span.
    private static bool PrepareExactTableBorders(XElement table)
    {
        table.RemoveAnnotations<ExactTableBorders>();
        var rows = WordprocessingMLUtil.TableRows(table).ToList();
        var widths = table.Element(W.tblGrid)?.Elements(W.gridCol).Select(c => (int?)c.Attribute(W._w) ?? 0).ToList();
        if (!IsFixedLayoutTable(table) || widths == null || widths.Count == 0 || widths.Any(w => w <= 0)
            || rows.Count == 0) return false;
        var grids = rows.Select(TableGridModel.RowGrid).ToList();
        foreach (var cell in grids.SelectMany(r => r)) cell.Cell.RemoveAnnotations<ExactTableCellBorders>();
        if (grids.Any(r => r.Any(c => c.End > widths.Count))) return false;
        var borders = grids.SelectMany(r => r).SelectMany(c => c.Cell.Element(W.tcPr)?.Element(W.tcBorders)?.Elements()
            ?? Enumerable.Empty<XElement>()).Where(b => (string?)b.Attribute(W.val) is not (null or "nil" or "none")).ToList();
        if (borders.Count == 0 || borders.Any(b => (string?)b.Attribute(W.val) != "single" || (int?)b.Attribute(W.sz) is not > 0))
            return false;

        var model = new ExactTableBorders();
        var spans = TableGridModel.VerticalMergeSpans(table);
        var slots = new List<List<TableBorderSlot>>();
        for (var row = 0; row < rows.Count; row++)
        {
            var current = new List<TableBorderSlot>();
            foreach (var cell in grids[row])
            {
                ExactTableCellBorders? owner = null;
                if (row > 0 && spans.TryGetValue(cell.Cell, out var span) && span == 0)
                    owner = slots[row - 1].FirstOrDefault(c => c.Start == cell.Start && c.End == cell.End)?.Owner;
                if (owner == null)
                {
                    owner = new ExactTableCellBorders { Cell = cell.Cell, FirstRow = row,
                        LastRow = row + (spans.TryGetValue(cell.Cell, out var count) ? Math.Max(1, count) : 1) - 1,
                        Start = cell.Start, End = cell.End };
                    model.Cells.Add(cell.Cell, owner);
                }
                current.Add(new TableBorderSlot(cell.Cell, owner, cell.Start, cell.End));
            }
            slots.Add(current);
            model.Rows.Add(rows[row], new List<TableBorderEdge>());
        }

        var rtl = IsOnOffPropertyEnabled(table.Element(W.tblPr)?.Element(W.bidiVisual));
        for (var row = 0; row <= rows.Count; row++)
        {
            var above = row > 0 ? slots[row - 1] : new List<TableBorderSlot>();
            var below = row < rows.Count ? slots[row] : new List<TableBorderSlot>();
            foreach (var (a, b, edge) in ResolveTableRowBoundary(above, below))
            {
                if (a != null) model.Rows[rows[row - 1]].Add(edge with { Side = "bottom" });
                if (b != null) model.Rows[rows[row]].Add(edge);
            }
        }
        // A repeated header meets a new body row on each continuation. Resolve these possible
        // joins here, using the same source-border conflict rule as the original row neighbors.
        var headerCount = rows.TakeWhile(r => IsOnOffPropertyEnabled(r.Element(W.trPr)?.Element(W.tblHeader))).Count();
        if (headerCount > 0 && headerCount < rows.Count)
            for (var row = headerCount; row < rows.Count; row++)
                model.HeaderJoins[rows[row]] = ResolveTableRowBoundary(slots[headerCount - 1], slots[row])
                    .Select(join => join.Edge).ToList();

        for (var row = 0; row < rows.Count; row++)
        {
            var cells = slots[row];
            var starts = cells.ToDictionary(c => c.Start);
            var ends = cells.ToDictionary(c => c.End);
            var boundaries = cells.SelectMany(c => new[] { c.Start, c.End }).Distinct().OrderBy(c => c);
            foreach (var column in boundaries)
            {
                starts.TryGetValue(column, out var b);
                ends.TryGetValue(column, out var a);
                if (a?.Owner == b?.Owner) continue;
                var winner = WinningTableBorder(a?.Source, rtl ? W.left : W.right, b?.Source, rtl ? W.right : W.left);
                if (winner == null) continue;
                var edge = new TableBorderEdge("vertical", column, column, winner);
                model.Rows[rows[row]].Add(edge);
                if (a != null)
                {
                    if (rtl) a.Owner.Left = Math.Max(a.Owner.Left, edge.Width);
                    else a.Owner.Right = Math.Max(a.Owner.Right, edge.Width);
                }
                if (b != null)
                {
                    if (rtl) b.Owner.Right = Math.Max(b.Owner.Right, edge.Width);
                    else b.Owner.Left = Math.Max(b.Owner.Left, edge.Width);
                }
            }
        }
        decimal RowWidth(int row, string side) => model.Rows[rows[row]].Where(e => e.Side == side).Select(e => e.Width).DefaultIfEmpty().Max();
        foreach (var cell in model.Cells.Values)
        {
            cell.Top = RowWidth(cell.FirstRow, "top");
            cell.Bottom = RowWidth(Math.Min(rows.Count - 1, cell.LastRow), "bottom");
            cell.Cell.AddAnnotation(cell);
        }
        model.Top = RowWidth(0, "top");
        model.Bottom = RowWidth(rows.Count - 1, "bottom");
        table.AddAnnotation(model);
        return true;
    }

    private static IEnumerable<(TableBorderSlot? Above, TableBorderSlot? Below, TableBorderEdge Edge)> ResolveTableRowBoundary(
        List<TableBorderSlot> above, List<TableBorderSlot> below)
    {
        var boundaries = above.Concat(below).SelectMany(c => new[] { c.Start, c.End }).Distinct().OrderBy(c => c).ToList();
        var ai = 0;
        var bi = 0;
        for (var i = 1; i < boundaries.Count; i++)
        {
            var start = boundaries[i - 1];
            while (ai < above.Count && above[ai].End <= start) ai++;
            while (bi < below.Count && below[bi].End <= start) bi++;
            var a = ai < above.Count && above[ai].Start <= start ? above[ai] : null;
            var b = bi < below.Count && below[bi].Start <= start ? below[bi] : null;
            if (a?.Owner == b?.Owner) continue;
            var winner = WinningTableBorder(a?.Source, W.bottom, b?.Source, W.top);
            if (winner != null) yield return (a, b, new TableBorderEdge("top", start, boundaries[i], winner));
        }
    }

    private static XElement? WinningTableBorder(XElement? firstCell, XName firstSide, XElement? secondCell, XName secondSide)
    {
        static XElement? Visible(XElement? cell, XName side)
        {
            var border = cell?.Element(W.tcPr)?.Element(W.tcBorders)?.Element(side);
            return (string?)border?.Attribute(W.val) is null or "nil" or "none" ? null : border;
        }
        var first = Visible(firstCell, firstSide);
        var second = Visible(secondCell, secondSide);
        if (first == null) return second == null ? null : new XElement(second);
        if (second == null) return new XElement(first);
        // Reuse the conflict rule, then snapshot the complete winner before legacy CSS adjustment.
        var copy = new XElement(first);
        ResolveCellBorder(copy, new XElement(second));
        return new XElement(XNode.DeepEquals(new XElement("border", copy.Attributes()), new XElement("border", first.Attributes())) ? first : second);
    }

    private static void RecordExactTableBorderMetadata(XElement source, XElement html,
        Dictionary<string, string> style, decimal? indent, string alignment, bool rtl, bool floating, decimal? oldLeading)
    {
        var model = source.Annotation<ExactTableBorders>();
        if (model == null) return;
        html.SetAttributeValue("data-docx-exact-table-borders", source.Annotation<TableRenderingGrid>()!.LeadingColumns);
        style["--docx-table-border-top"] = PointsCss(model.Top);
        style["--docx-table-border-bottom"] = PointsCss(model.Bottom);
        var first = model.Cells.Values.FirstOrDefault();
        if (first == null) return;
        var half = (rtl ? first.Right : first.Left) / 2m;
        var margin = CellMarginTwips(first.Cell, rtl ? W.right : W.left) / 20m;
        var legacy = (string?)source.Attribute(LegacyTableIndentAttribute);
        if (!floating && alignment is "left" or "start")
        {
            var center = legacy == "text" ? (indent ?? 0m) - Math.Max(margin, half)
                : legacy != null ? indent ?? 0m : (indent ?? 0m) + half;
            html.SetAttributeValue("data-docx-grid-margin-side", rtl ? "right" : "left");
            style["--docx-table-grid-margin"] = PointsCss(center);
        }
        else if (floating)
        {
            html.SetAttributeValue("data-docx-grid-margin-side", rtl ? "right" : "left");
            style["--docx-table-grid-margin"] = PointsCss((oldLeading ?? 0m) + half);
        }
        else if (alignment is "right" or "end")
        {
            var end = model.Cells.Values.Max(c => c.End);
            var trailing = model.Cells.Values.Where(c => c.End == end).Max(c => rtl ? c.Left : c.Right);
            html.SetAttributeValue("data-docx-grid-margin-side", rtl ? "left" : "right");
            html.SetAttributeValue("data-docx-grid-align-end", "true");
            style["--docx-table-grid-margin"] = PointsCss(trailing / 2m);
        }
    }

    private static void RecordExactCellBorderMetadata(XElement source, XElement html, Dictionary<string, string> style)
    {
        var borders = source.Annotation<ExactTableCellBorders>();
        if (borders == null) return;
        html.SetAttributeValue("data-docx-exact-cell-borders", "true");
        foreach (var (side, width) in new[] { ("top", borders.Top), ("right", borders.Right), ("bottom", borders.Bottom), ("left", borders.Left) })
            style["--docx-cell-border-" + side] = PointsCss(width);
    }

    private static void RecordExactRowBorderMetadata(XElement source, XElement html, ThemeColorScheme? theme)
    {
        var table = source.Ancestors(W.tbl).FirstOrDefault();
        var model = table?.Annotation<ExactTableBorders>();
        if (model == null || !model.Rows.TryGetValue(source, out var edges)) return;
        html.SetAttributeValue("data-docx-table-border-edges", TableBorderEdgesJson(edges, theme));
        if (model.HeaderJoins.TryGetValue(source, out var join))
            html.SetAttributeValue("data-docx-table-header-edges", TableBorderEdgesJson(join, theme));
        var height = TableRowHeightInTwips(source.Element(W.trPr)?.Element(W.trHeight)?.Attribute(W.val));
        if (height != null) html.SetAttributeValue("data-docx-source-row-height", (height.Value / 20m).ToString(CultureInfo.InvariantCulture));
    }

    private static string TableBorderEdgesJson(List<TableBorderEdge> edges, ThemeColorScheme? theme)
    {
        using var buffer = new MemoryStream();
        using (var json = new Utf8JsonWriter(buffer))
        {
            json.WriteStartArray();
            foreach (var edge in edges)
            {
                var color = ResolveThemeColor(edge.Border, W.color, W.themeColor, W.themeTint, W.themeShade, theme);
                color = color == null || color == "auto" ? "currentColor" : ConvertColor(color);
                json.WriteStartArray();
                json.WriteStringValue(edge.Side);
                json.WriteNumberValue(edge.Start);
                json.WriteNumberValue(edge.End);
                json.WriteNumberValue(edge.Width);
                json.WriteStringValue(color);
                json.WriteEndArray();
            }
            json.WriteEndArray();
        }
        return Encoding.UTF8.GetString(buffer.ToArray());
    }
}

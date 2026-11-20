// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Docxodus;

/// <summary>A table's preferred width (<c>w:tblW</c>): automatic, a fixed width in twips, or a
/// percentage of the text width.</summary>
public enum TableWidthType { Auto, Twips, Percent }

/// <summary>What <see cref="DocxSession.SetCellProperties"/> applies to: the anchored cell, every
/// cell of its row, every cell covering its grid column, or every cell of the table.</summary>
public enum TableCellScope { Cell, Row, Column, Table }

/// <summary>A table's default cell margins (<c>w:tblCellMar</c>), in twips. A null edge is left
/// as it is.</summary>
public sealed record TableCellMargins
{
    public int? TopTwips { get; init; }
    public int? LeftTwips { get; init; }
    public int? BottomTwips { get; init; }
    public int? RightTwips { get; init; }
}

/// <summary>
/// Word's Table Properties, Table and Alt Text tabs, for <see cref="DocxSession.SetTableProperties"/>.
/// Every member is optional: null leaves that property as it is.
/// </summary>
public sealed record TablePropertiesOp
{
    /// <summary>The preferred width's kind; with <see cref="Width"/> (twips for
    /// <see cref="TableWidthType.Twips"/>, percent of the text width for
    /// <see cref="TableWidthType.Percent"/>; ignored for <see cref="TableWidthType.Auto"/>).</summary>
    public TableWidthType? WidthType { get; init; }
    public double? Width { get; init; }

    /// <summary>The table's alignment between the margins (<c>w:jc</c>): "left", "center" or
    /// "right".</summary>
    public string? Alignment { get; init; }

    /// <summary>Indent from the left margin (<c>w:tblInd</c>), twips, may be negative. 0 removes it.</summary>
    public int? IndentTwips { get; init; }

    /// <summary>Default cell margins (<c>w:tblCellMar</c>).</summary>
    public TableCellMargins? CellMargins { get; init; }

    /// <summary>Alt text (<c>w:tblCaption</c>, <c>w:tblDescription</c>); "" removes it.</summary>
    public string? AltTitle { get; init; }
    public string? AltDescription { get; init; }
}

/// <summary>Cell properties for <see cref="DocxSession.SetCellProperties"/>. Null leaves a
/// property as it is.</summary>
public sealed record CellPropertiesOp
{
    /// <summary>Vertical alignment of the cell's text (<c>w:vAlign</c>): "top", "center" or
    /// "bottom". "top", Word's default, is written as no <c>w:vAlign</c>.</summary>
    public string? VerticalAlignment { get; init; }
}

/// <summary>What <see cref="DocxSession.GetTableProperties"/> reads: the table's, the anchored
/// cell's row's and the cell's own direct properties, as Word's Table Properties dialog shows them.
/// A null value is not set on the element itself.</summary>
public sealed record TablePropertiesInfo
{
    public required string TableAnchorId { get; init; }
    public TableWidthType? WidthType { get; init; }
    /// <summary>Twips, or percent for <see cref="TableWidthType.Percent"/>.</summary>
    public double? Width { get; init; }
    public string? Alignment { get; init; }
    public int? IndentTwips { get; init; }
    public TableCellMargins? CellMargins { get; init; }
    public string? AltTitle { get; init; }
    public string? AltDescription { get; init; }
    /// <summary>"fixed" or "autofit" (<c>w:tblLayout</c>), or null.</summary>
    public string? Layout { get; init; }
    /// <summary>The table grid's column widths (<c>w:tblGrid</c>), twips.</summary>
    public required IReadOnlyList<int> ColumnWidthsTwips { get; init; }
    /// <summary>The anchored cell's first grid column and how many it spans.</summary>
    public int ColumnIndex { get; init; }
    public int ColumnSpan { get; init; } = 1;
    public int RowIndex { get; init; }
    public int RowCount { get; init; }
    public int? RowHeightTwips { get; init; }
    public TableRowHeightRule? RowHeightRule { get; init; }
    public bool AllowBreakAcrossPages { get; init; } = true;
    public bool RepeatHeader { get; init; }
    public string? VerticalAlignment { get; init; }
    public int? CellWidthTwips { get; init; }
}

public sealed partial class DocxSession
{
    private static readonly string[] TblCellMarEdgeOrder = { "top", "start", "left", "bottom", "end", "right" };

    /// <summary>
    /// Set the table properties of the table containing <paramref name="cellAnchorId"/> (a
    /// <c>tc</c> anchor, a paragraph in a cell, or the table's <c>tbl</c> anchor): its preferred
    /// width, alignment, indent, default cell margins and alt text, as Word's Table Properties
    /// dialog does. One undo step; with changes tracked, one <c>w:tblPrChange</c>.
    /// </summary>
    public EditResult SetTableProperties(string cellAnchorId, TablePropertiesOp op)
    {
        ArgumentNullException.ThrowIfNull(op);
        if (ResolveTableFor(cellAnchorId, out var tbl, out var target) is { } err) return err;

        if (op.Alignment is { } a && a is not ("left" or "center" or "right"))
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                $"alignment must be left, center or right; got '{a}'", cellAnchorId);
        if (op.WidthType is TableWidthType.Twips or TableWidthType.Percent && op.Width is not > 0)
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "a twips or percent width needs a positive width", cellAnchorId);
        if (op.WidthType is TableWidthType.Percent && op.Width > 100)
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "a percent width is at most 100", cellAnchorId);
        if (op.CellMargins is { } m && new[] { m.TopTwips, m.LeftTwips, m.BottomTwips, m.RightTwips }.Any(v => v is < 0))
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                "cell margins in twips must be >= 0", cellAnchorId);

        var existing = tbl!.Element(W.tblPr);
        if (RefuseNestedTrackedPropertyChange(new[] { (existing, W.tblPrChange) }, cellAnchorId) is { } pending)
            return pending;
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var before = new XElement(existing ?? new XElement(W.tblPr));

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            var tblPr = GetOrCreateTblPr(tbl);
            if (op.WidthType is { } type)
            {
                var (w, kind) = type switch
                {
                    TableWidthType.Twips => ((int)Math.Round(op.Width!.Value), "dxa"),
                    TableWidthType.Percent => ((int)Math.Round(op.Width!.Value * 50), "pct"),
                    _ => (0, "auto"),
                };
                SetChildInOrder(tblPr, new XElement(W.tblW, new XAttribute(W._w, w), new XAttribute(W.type, kind)),
                    TblPrChildOrder);
            }
            if (op.Alignment is { } alignment)
            {
                if (alignment == "left") tblPr.Elements(W.jc).Remove();
                else SetChildInOrder(tblPr, new XElement(W.jc, new XAttribute(W.val, alignment)), TblPrChildOrder);
            }
            if (op.IndentTwips is { } indent)
            {
                if (indent == 0) tblPr.Elements(W.tblInd).Remove();
                else SetChildInOrder(tblPr,
                    new XElement(W.tblInd, new XAttribute(W._w, indent), new XAttribute(W.type, "dxa")), TblPrChildOrder);
            }
            if (op.CellMargins is { } margins)
            {
                var mar = tblPr.Element(W.tblCellMar);
                if (mar is null) { mar = new XElement(W.tblCellMar); SetChildInOrder(tblPr, mar, TblPrChildOrder); }
                void Edge(XName name, XName? alias, int? twips)
                {
                    if (twips is not { } v) return;
                    if (alias is not null) mar.Elements(alias).Remove();
                    SetChildInOrder(mar, new XElement(name, new XAttribute(W._w, v), new XAttribute(W.type, "dxa")),
                        TblCellMarEdgeOrder);
                }
                Edge(W.top, null, margins.TopTwips);
                Edge(W.left, W.start, margins.LeftTwips);
                Edge(W.bottom, null, margins.BottomTwips);
                Edge(W.right, W.end, margins.RightTwips);
                if (!mar.HasElements) mar.Remove();
            }
            void Text(XName name, string? value)
            {
                if (value is null) return;
                if (value.Length == 0) tblPr.Elements(name).Remove();
                else SetChildInOrder(tblPr, new XElement(name, new XAttribute(W.val, value)), TblPrChildOrder);
            }
            Text(W.tblCaption, op.AltTitle);
            Text(W.tblDescription, op.AltDescription);

            if (tracked && !PropertySnapshotEquals(before, tblPr, W.tblPrChange))
                TrackPropertyMutation(tblPr, before, W.tblPrChange,
                    _revisionAuthor ?? "docxodus", NextTrackedFormatRevisionDate());
            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>
    /// Set cell properties (the vertical alignment of their text) on the cell
    /// <paramref name="cellAnchorId"/>, its row, its grid column or the whole table. One undo
    /// step; with changes tracked, a <c>w:tcPrChange</c> on each changed cell.
    /// </summary>
    public EditResult SetCellProperties(string cellAnchorId, CellPropertiesOp op, TableCellScope scope = TableCellScope.Cell)
    {
        ArgumentNullException.ThrowIfNull(op);
        if (ResolveCell(cellAnchorId, out _, out var tc, out var tr, out var tbl, out var target) is { } err)
            return err;
        if (op.VerticalAlignment is { } v && v is not ("top" or "center" or "bottom"))
            return EditResult.Fail(EditErrorCode.InvalidTableStyling,
                $"vertical alignment must be top, center or bottom; got '{v}'", cellAnchorId);

        List<XElement> cells;
        switch (scope)
        {
            case TableCellScope.Row:
                cells = tr!.Elements(W.tc).ToList();
                break;
            case TableCellScope.Table:
                cells = tbl!.Elements(W.tr).SelectMany(row => row.Elements(W.tc)).ToList();
                break;
            case TableCellScope.Column:
                var mine = RowGrid(tr!).First(c => c.Tc == tc);
                cells = tbl!.Elements(W.tr).SelectMany(RowGrid)
                    .Where(c => c.Start < mine.End && mine.Start < c.End).Select(c => c.Tc).ToList();
                break;
            default:
                cells = new List<XElement> { tc! };
                break;
        }

        if (RefuseNestedTrackedPropertyChange(cells.Select(cell => (cell.Element(W.tcPr), W.tcPrChange)),
                cellAnchorId) is { } pending)
            return pending;
        var tracked = _trackedChanges == TrackedChangeMode.RenderInline;
        var old = tracked ? cells.ToDictionary(cell => cell,
            cell => new XElement(cell.Element(W.tcPr) ?? new XElement(W.tcPr))) : null;

        _history.RecordPreOp(TakeSnapshot());
        try
        {
            foreach (var cell in cells)
            {
                if (op.VerticalAlignment is { } valign)
                {
                    if (valign == "top") cell.Element(W.tcPr)?.Elements(W.vAlign).Remove();
                    else SetChildInOrder(GetOrCreateTcPr(cell), new XElement(W.vAlign, new XAttribute(W.val, valign)),
                        TcPrChildOrder);
                }
            }
            if (tracked)
            {
                var author = _revisionAuthor ?? "docxodus";
                var date = NextTrackedFormatRevisionDate();
                foreach (var pair in old!)
                    if (!PropertySnapshotEquals(pair.Value, pair.Key.Element(W.tcPr),
                            W.tcPrChange, W.cellIns, W.cellDel, W.cellMerge))
                        TrackPropertyMutation(GetOrCreateTcPr(pair.Key), pair.Value,
                            W.tcPrChange, author, date, W.cellIns, W.cellDel, W.cellMerge);
            }
            return TableStyleResult(target!);
        }
        catch (Exception ex)
        {
            return FailInternal(ex, cellAnchorId);
        }
    }

    /// <summary>Read the properties Word's Table Properties dialog shows for the table, row and
    /// cell of <paramref name="cellAnchorId"/> (a <c>tc</c> anchor or a paragraph in a cell), or
    /// null when the anchor is not in a table.</summary>
    public TablePropertiesInfo? GetTableProperties(string cellAnchorId)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(cellAnchorId);
        var target = FindAnchor(cellAnchorId);
        var element = target?.Resolve(_doc!);
        if (element is null) return null;
        var tc = element.Name == W.tc ? element : element.Ancestors(W.tc).FirstOrDefault();
        var tr = tc?.Parent;
        var tbl = tr?.Parent;
        if (tc is null || tr?.Name != W.tr || tbl?.Name != W.tbl) return null;

        var tblPr = tbl.Element(W.tblPr);
        var tblW = tblPr?.Element(W.tblW);
        TableWidthType? widthType = (string?)tblW?.Attribute(W.type) switch
        {
            "dxa" => TableWidthType.Twips,
            "pct" => TableWidthType.Percent,
            "auto" or "nil" => TableWidthType.Auto,
            null when tblW is not null => TableWidthType.Twips,
            _ => null,
        };
        double? width = null;
        if (widthType is TableWidthType.Twips or TableWidthType.Percent)
        {
            var raw = ((string?)tblW!.Attribute(W._w) ?? "").Trim();
            var percentText = raw.EndsWith('%');
            if (double.TryParse(percentText ? raw[..^1] : raw, NumberStyles.Float, CultureInfo.InvariantCulture, out var n))
                width = widthType == TableWidthType.Percent && !percentText ? n / 50 : n;
        }
        int? Twips(XElement? e) => int.TryParse((string?)e?.Attribute(W._w), NumberStyles.Integer,
            CultureInfo.InvariantCulture, out var n) && (string?)e!.Attribute(W.type) is null or "dxa" ? n : null;
        var mar = tblPr?.Element(W.tblCellMar);
        var margins = mar is null ? null : new TableCellMargins
        {
            TopTwips = Twips(mar.Element(W.top)),
            LeftTwips = Twips(mar.Element(W.left) ?? mar.Element(W.start)),
            BottomTwips = Twips(mar.Element(W.bottom)),
            RightTwips = Twips(mar.Element(W.right) ?? mar.Element(W.end)),
        };
        var columns = tbl.Element(W.tblGrid)?.Elements(W.gridCol)
            .Select(col => int.TryParse((string?)col.Attribute(W._w), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0)
            .ToList() ?? new List<int>();
        var gridCells = RowGrid(tr);
        var at = gridCells.FindIndex(c => c.Tc == tc);
        var trPr = tr.Element(W.trPr);
        var height = trPr?.Element(W.trHeight);
        int? heightTwips = int.TryParse((string?)height?.Attribute(W.val), NumberStyles.Integer, CultureInfo.InvariantCulture, out var h) ? h : null;
        var rows = tbl.Elements(W.tr).ToList();
        var tcPr = tc.Element(W.tcPr);
        var tableAnchor = AnchorForElement(tbl);
        return new TablePropertiesInfo
        {
            TableAnchorId = tableAnchor?.Id ?? "",
            WidthType = widthType,
            Width = width,
            Alignment = (string?)tblPr?.Element(W.jc)?.Attribute(W.val) switch
            {
                "start" => "left",
                "end" => "right",
                var j => j,
            },
            IndentTwips = Twips(tblPr?.Element(W.tblInd)),
            CellMargins = margins,
            AltTitle = (string?)tblPr?.Element(W.tblCaption)?.Attribute(W.val),
            AltDescription = (string?)tblPr?.Element(W.tblDescription)?.Attribute(W.val),
            Layout = (string?)tblPr?.Element(W.tblLayout)?.Attribute(W.type),
            ColumnWidthsTwips = columns,
            ColumnIndex = at < 0 ? 0 : gridCells[at].Start,
            ColumnSpan = at < 0 ? 1 : Math.Max(1, gridCells[at].End - gridCells[at].Start),
            RowIndex = rows.IndexOf(tr),
            RowCount = rows.Count,
            RowHeightTwips = heightTwips,
            RowHeightRule = heightTwips is null ? null : (string?)height!.Attribute(W.hRule) switch
            {
                "exact" => TableRowHeightRule.Exact,
                "auto" => TableRowHeightRule.Auto,
                _ => TableRowHeightRule.AtLeast,
            },
            AllowBreakAcrossPages = !ReadOnOff(trPr?.Element(W.cantSplit)),
            RepeatHeader = ReadOnOff(trPr?.Element(W.tblHeader)),
            VerticalAlignment = (string?)tcPr?.Element(W.vAlign)?.Attribute(W.val),
            CellWidthTwips = Twips(tcPr?.Element(W.tcW)),
        };
    }

    private static bool ReadOnOff(XElement? e) =>
        e is not null && (string?)e.Attribute(W.val) is null or "1" or "true" or "on";

    /// <summary>The table of a <c>tbl</c> anchor, or of the cell a <c>tc</c> or in-cell paragraph
    /// anchor is in.</summary>
    private EditResult? ResolveTableFor(string anchorId, out XElement? tbl, out AnchorTarget? target)
    {
        target = FindAnchor(anchorId);
        tbl = target?.Anchor.Kind == "tbl" ? target.Resolve(_doc!) : null;
        if (tbl?.Name == W.tbl)
            return MutationRefusal() is { } refusal ? EditResult.Fail(refusal) : null;
        return ResolveCell(anchorId, out _, out _, out _, out tbl, out target);
    }
}

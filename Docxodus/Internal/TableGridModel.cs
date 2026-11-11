// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Xml.Linq;

namespace Docxodus.Internal;

/// <summary>A physical cell's geometry in the Word table grid. <see cref="End"/> is exclusive.</summary>
internal readonly record struct TableGridCell(XElement Cell, int Start, int Span)
{
    internal XElement Tc => Cell;
    internal int End => Start + Span;
}

/// <summary>
/// Single owner of Word table-grid geometry and canonical structural metadata. Mutation code,
/// live-session resolution, and byte-oriented structure analysis all use this model so omitted
/// columns, spans, and vertical merges cannot acquire competing interpretations.
/// </summary>
internal static class TableGridModel
{
    private static int ValOf(XElement? e) => Math.Max(0, (int?)e?.Attribute(W.val) ?? 0);

    internal static int GridBefore(XElement row) => ValOf(row.Element(W.trPr)?.Element(W.gridBefore));

    internal static int GridAfter(XElement row) => ValOf(row.Element(W.trPr)?.Element(W.gridAfter));

    // Compare can already supply revised spans beside historical deleted cells. Such a row
    // must not absorb the deleted spans again when projecting its current geometry.
    internal static bool RetainedCellsFillGrid(XElement row)
    {
        int columns = row.Ancestors(W.tbl).FirstOrDefault()?.Element(W.tblGrid)?.Elements(W.gridCol).Count() ?? 0;
        int retained = WordprocessingMLUtil.RowCells(row).Where(tc => !tc.Elements(W.tcPr).Elements(W.cellDel).Any())
            .Sum(tc => (int?)tc.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);
        retained += (int?)row.Element(W.trPr)?.Element(W.gridBefore)?.Attribute(W.val) ?? 0;
        retained += (int?)row.Element(W.trPr)?.Element(W.gridAfter)?.Attribute(W.val) ?? 0;
        return columns > 0 && retained == columns;
    }

    internal static List<TableGridCell> RowGrid(XElement row)
    {
        int column = GridBefore(row);
        var cells = new List<TableGridCell>();
        foreach (var cell in WordprocessingMLUtil.RowCells(row))
        {
            int span = Math.Max(1, (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);
            cells.Add(new TableGridCell(cell, column, span));
            column += span;
        }
        return cells;
    }

    /// <summary>
    /// Takes cells out of their row: the one resolution of a cell that leaves, whether by an
    /// accepted <c>w:cellDel</c> or a rejected <c>w:cellIns</c>. A row left narrower than its
    /// table is made whole again, as Word 2024 does it: each leaving cell's columns go to the
    /// nearest staying cell before it, those that start the row to the first staying cell after
    /// them, and the widened cell is as wide as the grid columns it then spans. How wide the
    /// table is depends on what the stated grid describes. It describes the table without its
    /// deleted cells, so a row that loses a deleted cell must still fill it. It describes the
    /// table with its inserted cells, so it says nothing about a row without them; there the
    /// rows that keep every cell are the witness, and with none, nothing is widened. A row that
    /// is already whole is left alone, whoever wrote the mark: the property changes resolved
    /// with it (a restored grid, a restored span) already describe the row. A content-control or
    /// custom-XML wrapper left without a cell goes with its cell. <paramref name="leaving"/>
    /// names every cell leaving the table, in any row, and <paramref name="inserted"/> says they
    /// are rejected insertions. Returns <c>false</c>, changing nothing,
    /// when every cell of the row leaves: the row itself goes, which is the caller's to do.
    /// </summary>
    internal static bool RemoveCells(XElement row, IReadOnlySet<XElement> leaving, bool inserted)
    {
        var cells = WordprocessingMLUtil.RowCells(row).ToList();
        var staying = cells.Where(cell => !leaving.Contains(cell)).ToList();
        if (staying.Count == 0) return false;

        int Span(XElement cell) => Math.Max(1, (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);
        int Columns(XElement r) => GridBefore(r) + GridAfter(r)
            + WordprocessingMLUtil.RowCells(r).Where(cell => !leaving.Contains(cell)).Sum(Span);
        var table = row.Ancestors(W.tbl).FirstOrDefault();
        int stated = table?.Element(W.tblGrid)?.Elements(W.gridCol).Count() ?? 0;
        int tableColumns = inserted
            ? (table is null ? Enumerable.Empty<XElement>() : WordprocessingMLUtil.TableRows(table))
                .Where(r => !WordprocessingMLUtil.RowCells(r).Any(leaving.Contains)).Select(Columns).DefaultIfEmpty(0).Max()
            : stated > 0 ? stated : GridBefore(row) + GridAfter(row) + cells.Sum(Span);
        int missing = tableColumns - Columns(row);

        var widened = new HashSet<XElement>();
        foreach (var cell in cells.Where(leaving.Contains))
        {
            int taken = Math.Min(Span(cell), missing);
            if (taken > 0)
            {
                int position = cells.IndexOf(cell);
                var taker = staying.LastOrDefault(other => cells.IndexOf(other) < position) ?? staying[0];
                var properties = taker.Element(W.tcPr);
                if (properties is null) taker.AddFirst(properties = new XElement(W.tcPr));
                var gridSpan = properties.Element(W.gridSpan);
                if (gridSpan is null)
                {
                    // CT_TcPr order: cnfStyle, tcW, gridSpan, then everything else.
                    gridSpan = new XElement(W.gridSpan);
                    var after = properties.Elements().LastOrDefault(e => e.Name == W.cnfStyle || e.Name == W.tcW);
                    if (after is null) properties.AddFirst(gridSpan);
                    else after.AddAfterSelf(gridSpan);
                }
                gridSpan.SetAttributeValue(W.val, Span(taker) + taken);
                widened.Add(taker);
                missing -= taken;
            }

            var emptied = cell;
            while (emptied.Parent is { } parent && parent != row
                && WordprocessingMLUtil.RowCells(parent).All(other => other == cell))
                emptied = parent;
            emptied.Remove();
        }

        foreach (var shape in RowGrid(row).Where(shape => widened.Contains(shape.Cell)))
        {
            var width = shape.Cell.Element(W.tcPr)!.Element(W.tcW);
            var spanned = table!.Elements(W.tblGrid).Elements(W.gridCol).Skip(shape.Start).Take(shape.Span)
                .Select(column => long.TryParse((string?)column.Attribute(W._w), out long twips) ? twips : (long?)null).ToList();
            if ((string?)width?.Attribute(W.type) == "dxa" && spanned.Count == shape.Span && !spanned.Contains(null))
                width.SetAttributeValue(W._w, spanned.Sum());
        }
        return true;
    }

    internal static TableGridCell? CellCovering(IEnumerable<TableGridCell> grid, int column)
    {
        foreach (var cell in grid)
            if (column >= cell.Start && column < cell.End)
                return cell;
        return null;
    }

    internal static XElement? AlignedCell(XElement row, TableGridCell shape)
    {
        foreach (var cell in RowGrid(row))
            if (cell.Start == shape.Start && cell.End == shape.End)
                return cell.Cell;
        return null;
    }

    internal static TableVerticalMergeRole VerticalMergeRole(XElement cell)
    {
        var merge = cell.Element(W.tcPr)?.Element(W.vMerge);
        if (merge is null) return TableVerticalMergeRole.None;
        return string.Equals((string?)merge.Attribute(W.val), "restart", StringComparison.OrdinalIgnoreCase)
            ? TableVerticalMergeRole.Restart
            : TableVerticalMergeRole.Continue;
    }

    internal static int GridColumnCount(XElement table)
    {
        int explicitCount = table.Element(W.tblGrid)?.Elements(W.gridCol).Count() ?? 0;
        int rowCount = WordprocessingMLUtil.TableRows(table)
            .Select(row =>
            {
                var grid = RowGrid(row);
                int end = grid.Count == 0 ? GridBefore(row) : grid[^1].End;
                return end + GridAfter(row);
            })
            .DefaultIfEmpty(0)
            .Max();
        return Math.Max(explicitCount, rowCount);
    }

    internal static TableMetadata BuildMetadata(XElement table, Func<XElement, Anchor?> anchorForElement)
    {
        var tableAnchor = anchorForElement(table)
            ?? throw new InvalidOperationException("table has no canonical anchor");
        var rowElements = WordprocessingMLUtil.TableRows(table).ToList();
        var verticalSpans = VerticalMergeSpans(table);
        var rows = new List<TableRowMetadata>(rowElements.Count);

        for (int rowIndex = 0; rowIndex < rowElements.Count; rowIndex++)
        {
            var row = rowElements[rowIndex];
            var rowAnchor = anchorForElement(row)
                ?? throw new InvalidOperationException("table row has no canonical anchor");
            var cells = new List<TableCellMetadata>();
            foreach (var geometry in RowGrid(row))
            {
                var cellAnchor = anchorForElement(geometry.Cell)
                    ?? throw new InvalidOperationException("table cell has no canonical anchor");
                var role = VerticalMergeRole(geometry.Cell);
                int rowSpan = role == TableVerticalMergeRole.Continue
                    ? 0
                    : role == TableVerticalMergeRole.Restart
                        ? verticalSpans[geometry.Cell]
                        : 1;
                cells.Add(new TableCellMetadata
                {
                    Anchor = cellAnchor,
                    TableAnchorId = tableAnchor.Id,
                    RowAnchorId = rowAnchor.Id,
                    RowIndex = rowIndex,
                    ColumnIndex = geometry.Start,
                    RowSpan = rowSpan,
                    ColumnSpan = geometry.Span,
                    VerticalMerge = role,
                    ParagraphAnchors = geometry.Cell.Elements(W.p)
                        .Select(anchorForElement)
                        .Where(anchor => anchor is not null)
                        .Select(anchor => anchor!.Value)
                        .ToList(),
                });
            }
            rows.Add(new TableRowMetadata
            {
                Anchor = rowAnchor,
                TableAnchorId = tableAnchor.Id,
                RowIndex = rowIndex,
                GridBefore = GridBefore(row),
                GridAfter = GridAfter(row),
                Cells = cells,
            });
        }

        int columnCount = GridColumnCount(table);
        var gridColumns = table.Element(W.tblGrid)?.Elements(W.gridCol).ToList() ?? new List<XElement>();
        var columns = new List<TableColumnMetadata>(columnCount);
        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
        {
            Anchor columnAnchor;
            int width = 0;
            if (columnIndex < gridColumns.Count)
            {
                columnAnchor = anchorForElement(gridColumns[columnIndex])
                    ?? throw new InvalidOperationException("table grid column has no canonical anchor");
                width = Math.Max(0, (int?)gridColumns[columnIndex].Attribute(W._w) ?? 0);
            }
            else
            {
                // OOXML normally has w:tblGrid. Keep malformed/legacy tables inspectable without
                // mutating a read call; the next structural mutation materializes real gridCols.
                string unid = UnidHelper.ShortHash(
                    $"{tableAnchor.Unid}:virtual-col:{columnIndex}", hexChars: 32);
                columnAnchor = new Anchor($"col:{tableAnchor.Scope}:{unid}", "col", tableAnchor.Scope, unid);
            }
            columns.Add(new TableColumnMetadata
            {
                Anchor = columnAnchor,
                TableAnchorId = tableAnchor.Id,
                ColumnIndex = columnIndex,
                WidthTwips = width,
                IsVirtual = columnIndex >= gridColumns.Count,
                CellAnchorIds = rows
                    .SelectMany(row => row.Cells)
                    .Where(cell => columnIndex >= cell.ColumnIndex
                        && columnIndex < cell.ColumnIndex + cell.ColumnSpan)
                    .Select(cell => cell.Anchor.Id)
                    .ToList(),
            });
        }

        return new TableMetadata { Anchor = tableAnchor, Columns = columns, Rows = rows };
    }

    internal static TableCellMetadata? CellAt(TableMetadata table, int rowIndex, int columnIndex)
    {
        if (rowIndex < 0 || rowIndex >= table.Rows.Count || columnIndex < 0) return null;
        return table.Rows[rowIndex].Cells.FirstOrDefault(cell =>
            columnIndex >= cell.ColumnIndex && columnIndex < cell.ColumnIndex + cell.ColumnSpan);
    }

    internal static IReadOnlyList<TableAnchorLocation> Locations(TableMetadata metadata)
    {
        var locations = new List<TableAnchorLocation>
        {
            new() { Anchor = metadata.Anchor, EntityKind = TableAnchorEntityKind.Table },
        };
        locations.AddRange(metadata.Columns.Select(column => new TableAnchorLocation
        {
            Anchor = column.Anchor,
            EntityKind = TableAnchorEntityKind.Column,
            ColumnIndex = column.ColumnIndex,
            IsVirtual = column.IsVirtual,
        }));
        foreach (var row in metadata.Rows)
        {
            locations.Add(new TableAnchorLocation
            {
                Anchor = row.Anchor,
                EntityKind = TableAnchorEntityKind.Row,
                RowIndex = row.RowIndex,
            });
            locations.AddRange(row.Cells.Select(cell => new TableAnchorLocation
            {
                Anchor = cell.Anchor,
                EntityKind = TableAnchorEntityKind.Cell,
                RowIndex = cell.RowIndex,
                ColumnIndex = cell.ColumnIndex,
                RowSpan = cell.RowSpan,
                ColumnSpan = cell.ColumnSpan,
            }));
        }
        return locations;
    }

    internal static TableAnchorMapping Map(TableMetadata? before, TableMetadata? after)
    {
        var oldLocations = before is null ? Array.Empty<TableAnchorLocation>() : Locations(before);
        var newLocations = after is null ? Array.Empty<TableAnchorLocation>() : Locations(after);
        var oldById = oldLocations.ToDictionary(location => location.Anchor.Id, StringComparer.Ordinal);
        var newById = newLocations.ToDictionary(location => location.Anchor.Id, StringComparer.Ordinal);
        return new TableAnchorMapping
        {
            Retained = oldLocations
                .Where(location => newById.ContainsKey(location.Anchor.Id))
                .Select(location => new RetainedTableAnchor(location, newById[location.Anchor.Id]))
                .ToList(),
            Added = newLocations.Where(location => !oldById.ContainsKey(location.Anchor.Id)).ToList(),
            Invalidated = oldLocations.Where(location => !newById.ContainsKey(location.Anchor.Id)).ToList(),
        };
    }

    internal static Dictionary<XElement, int> VerticalMergeSpans(XElement table)
    {
        var spans = new Dictionary<XElement, int>();
        var previous = new Dictionary<(int Start, int Span), XElement>();
        foreach (var row in WordprocessingMLUtil.TableRows(table))
        {
            var current = new Dictionary<(int Start, int Span), XElement>();
            foreach (var cell in RowGrid(row))
            {
                var interval = (cell.Start, cell.Span);
                if (VerticalMergeRole(cell.Cell) == TableVerticalMergeRole.Restart)
                {
                    current[interval] = cell.Cell;
                    spans[cell.Cell] = 1;
                }
                else if (VerticalMergeRole(cell.Cell) == TableVerticalMergeRole.Continue
                    && previous.TryGetValue(interval, out var restart))
                {
                    current[interval] = restart;
                    spans[restart]++;
                    spans[cell.Cell] = 0;
                }
            }
            previous = current;
        }
        return spans;
    }
}

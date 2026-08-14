// Copyright (c) Microsoft. All rights reserved.
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

    internal static List<TableGridCell> RowGrid(XElement row)
    {
        int column = GridBefore(row);
        var cells = new List<TableGridCell>();
        foreach (var cell in row.Elements(W.tc))
        {
            int span = Math.Max(1, (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1);
            cells.Add(new TableGridCell(cell, column, span));
            column += span;
        }
        return cells;
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
        int rowCount = table.Elements(W.tr)
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
        var rowElements = table.Elements(W.tr).ToList();
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
                        ? VerticalSpan(rowElements, rowIndex, geometry)
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

    private static int VerticalSpan(IReadOnlyList<XElement> rows, int rowIndex, TableGridCell restart)
    {
        int span = 1;
        for (int index = rowIndex + 1; index < rows.Count; index++)
        {
            var aligned = AlignedCell(rows[index], restart);
            if (aligned is null || VerticalMergeRole(aligned) != TableVerticalMergeRole.Continue) break;
            span++;
        }
        return span;
    }
}

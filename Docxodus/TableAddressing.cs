// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus;

/// <summary>The role of a physical <c>w:tc</c> in a vertical-merge run.</summary>
public enum TableVerticalMergeRole
{
    None,
    Restart,
    Continue,
}

/// <summary>The structural table identity represented by a table-anchor mapping entry.</summary>
public enum TableAnchorEntityKind
{
    Table,
    Row,
    Column,
    Cell,
}

/// <summary>Metadata for one physical <c>w:tc</c>, addressed by its canonical <c>tc</c> anchor.</summary>
public sealed record TableCellMetadata
{
    required public Anchor Anchor { get; init; }
    required public string TableAnchorId { get; init; }
    required public string RowAnchorId { get; init; }
    public int RowIndex { get; init; }
    public int ColumnIndex { get; init; }
    public int RowSpan { get; init; } = 1;
    public int ColumnSpan { get; init; } = 1;
    public TableVerticalMergeRole VerticalMerge { get; init; }

    /// <summary>Direct cell paragraphs only. Paragraphs in nested tables belong to their own cells.</summary>
    public IReadOnlyList<Anchor> ParagraphAnchors { get; init; } = Array.Empty<Anchor>();
}

/// <summary>Metadata for one physical <c>w:tr</c>.</summary>
public sealed record TableRowMetadata
{
    required public Anchor Anchor { get; init; }
    required public string TableAnchorId { get; init; }
    public int RowIndex { get; init; }
    public int GridBefore { get; init; }
    public int GridAfter { get; init; }
    public IReadOnlyList<TableCellMetadata> Cells { get; init; } = Array.Empty<TableCellMetadata>();
}

/// <summary>Metadata for one table grid column, identified by its <c>w:gridCol</c> anchor.</summary>
public sealed record TableColumnMetadata
{
    required public Anchor Anchor { get; init; }
    required public string TableAnchorId { get; init; }
    public int ColumnIndex { get; init; }
    public int WidthTwips { get; init; }

    /// <summary>True only when an absent/underspecified <c>w:tblGrid</c> required a read-only
    /// coordinate identity. A shape/width transaction materializes a real gridCol anchor and
    /// reports this virtual identity invalidated.</summary>
    public bool IsVirtual { get; init; }

    /// <summary>Physical cells covering this grid column, top-to-bottom.</summary>
    public IReadOnlyList<string> CellAnchorIds { get; init; } = Array.Empty<string>();
}

/// <summary>
/// The canonical table-addressing view of one <c>w:tbl</c>. Table, row, column, and cell
/// identities are explicit; cells use zero-based Word table-grid coordinates.
/// </summary>
public sealed record TableMetadata
{
    required public Anchor Anchor { get; init; }
    public IReadOnlyList<TableColumnMetadata> Columns { get; init; } = Array.Empty<TableColumnMetadata>();
    public IReadOnlyList<TableRowMetadata> Rows { get; init; } = Array.Empty<TableRowMetadata>();
}

/// <summary>Result of resolving a table anchor to its metadata.</summary>
public sealed record TableMetadataResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public TableMetadata? Metadata { get; init; }
}

/// <summary>Result of either direction of canonical cell-anchor/coordinate resolution.</summary>
public sealed record TableCellResolutionResult
{
    public bool Success { get; init; }
    public EditError? Error { get; init; }
    public TableCellMetadata? Cell { get; init; }
}

/// <summary>A structural table anchor plus its location at one point in a mutation.</summary>
public sealed record TableAnchorLocation
{
    required public Anchor Anchor { get; init; }
    public TableAnchorEntityKind EntityKind { get; init; }
    public int? RowIndex { get; init; }
    public int? ColumnIndex { get; init; }
    public int? RowSpan { get; init; }
    public int? ColumnSpan { get; init; }
    public bool IsVirtual { get; init; }
}

/// <summary>A stable structural identity retained across a table mutation.</summary>
public sealed record RetainedTableAnchor(TableAnchorLocation Before, TableAnchorLocation After);

/// <summary>
/// Deterministic structural identity map for a table mutation. Retained entries are ordered by
/// their old location; added entries by their new location; invalidated entries by their old location.
/// </summary>
public sealed record TableAnchorMapping
{
    public IReadOnlyList<RetainedTableAnchor> Retained { get; init; } = Array.Empty<RetainedTableAnchor>();
    public IReadOnlyList<TableAnchorLocation> Added { get; init; } = Array.Empty<TableAnchorLocation>();
    public IReadOnlyList<TableAnchorLocation> Invalidated { get; init; } = Array.Empty<TableAnchorLocation>();
}

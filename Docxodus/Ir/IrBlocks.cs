#nullable enable

using System.Xml.Linq;

namespace Docxodus.Ir;

/// <summary>
/// Base type for block-level IR content. Every block carries a stable <see cref="Anchor"/>, a
/// <see cref="ContentHash"/> (text/structure digest) and a <see cref="FormatFingerprint"/>
/// (formatting digest), both computed by the reader (Task 4).
/// </summary>
/// <remarks>
/// <see cref="Source"/> is an <see cref="IrProvenance"/> whose equality is neutral (it equals any
/// other provenance), so it is excluded from a block's value equality even though it is a record
/// property. Block child collections use <see cref="IrNodeList{T}"/> so that node-for-node value
/// equality composes correctly down the tree (§8 determinism guarantee).
/// </remarks>
internal abstract record IrBlock
{
    public required IrAnchor Anchor { get; init; }
    public required IrHash ContentHash { get; init; }
    public required IrHash FormatFingerprint { get; init; }

    /// <summary>Back-reference to source OOXML; equality-neutral (does not affect record equality).</summary>
    public IrProvenance Source { get; init; } = new();
}

/// <summary>
/// A paragraph: its direct formatting (<see cref="Format"/>), optional list membership
/// (<see cref="List"/>), and inline children. No effective/cascaded format member — that is a
/// computed view added in M1.3.
/// </summary>
internal sealed record IrParagraph : IrBlock
{
    /// <summary>Direct paragraph formatting (`w:pPr`); cascade resolution is an M1.3 view, not stored here.</summary>
    public required IrParaFormat Format { get; init; }
    public IrListInfo? List { get; init; }
    public required IrNodeList<IrInline> Inlines { get; init; }
}

/// <summary>A table: its rows plus a digest of the unmodeled `w:tblPr`/`w:tblGrid` properties.</summary>
internal sealed record IrTable : IrBlock
{
    public required IrNodeList<IrRow> Rows { get; init; }

    /// <summary>Canonical hash of unmodeled table-level props (`w:tblPr`/`w:tblGrid`).</summary>
    public required IrHash UnmodeledTablePropsDigest { get; init; }
}

/// <summary>A table row. <paramref name="Source"/> is equality-neutral provenance.</summary>
internal sealed record IrRow(IrAnchor Anchor, IrNodeList<IrCell> Cells, IrHash ContentHash)
{
    public IrProvenance Source { get; init; } = new();
}

/// <summary>
/// A table cell: its block children plus grid span and vertical-merge state.
/// <paramref name="Source"/> is equality-neutral provenance.
/// </summary>
internal sealed record IrCell(IrAnchor Anchor, IrNodeList<IrBlock> Blocks,
                              int GridSpan, IrVMerge VMerge, IrHash ContentHash)
{
    public IrProvenance Source { get; init; } = new();
}

/// <summary>A section break carrying its direct section formatting (`w:sectPr`).</summary>
internal sealed record IrSectionBreak : IrBlock
{
    public required IrSectionFormat Format { get; init; }
}

/// <summary>
/// An unmodeled block-level element preserved opaquely. Its <see cref="IrBlock.ContentHash"/> is
/// the canonical hash of the source XML and its <see cref="IrBlock.FormatFingerprint"/> is the
/// hash of empty bytes (it has no modeled formatting).
/// </summary>
internal sealed record IrOpaqueBlock : IrBlock
{
    public required XName ElementName { get; init; }
}

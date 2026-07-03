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

    /// <summary>
    /// Back-reference to source OOXML; equality-neutral (does not affect record equality). Also carries
    /// the <see cref="IrProvenance.FromBlockSdt"/> flag (see there) which the markdown emitter uses to
    /// mirror the oracle's block-level-SDT skip without perturbing block value equality.
    /// </summary>
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

    /// <summary>
    /// The auto-number marker Word would render for this paragraph, resolved by the reader via
    /// <c>ListItemRetriever.RetrieveListItem</c> against the live package — the exact string the
    /// markdown projection's <c>ResolveListMarker</c>/<c>ResolveHeadingNumberPrefix</c> consume (e.g.
    /// <c>"1."</c>, <c>"a."</c>, <c>"1.1"</c>, a bullet glyph, or "First Article"). Null when the
    /// paragraph is not a list item / carries no resolvable numbering.
    /// <para>
    /// <b>Why this is an IR fact, not an emitter computation.</b> Word's numbering is a stateful
    /// document-order counter walk (per numId/ilvl, with continuation/restart/style-inherited cases)
    /// that needs the live <c>NumberingDefinitionsPart</c>/<c>StyleDefinitionsPart</c> and cross-
    /// paragraph state — facts the per-paragraph <see cref="List"/> info does NOT capture. The reader
    /// has the live package, so it resolves the marker once at read time (additive, equality-
    /// participating: the same document yields the same marker deterministically). The emitter then
    /// applies the projection's display rules (bullet→<c>-</c> for list items, single-glyph→null for
    /// heading prefixes) to this raw marker, guaranteeing byte-parity with the oracle without
    /// re-implementing the counter walk or peeking at <see cref="IrBlock.Source"/>.
    /// </para>
    /// <para>
    /// <b>Equality note.</b> Because this is an init-only record member it participates in
    /// <see cref="IrParagraph"/>'s record value-equality, which makes that equality <em>stricter</em>
    /// than <see cref="IrBlock.ContentHash"/> equality (the resolved marker is rendering-derived state,
    /// not content). Phase-2 diff code must key alignment/equality on the hashes
    /// (<c>ContentHash</c>/<c>FormatFingerprint</c>), never on record <c>==</c>.
    /// </para>
    /// </summary>
    public string? ResolvedListMarker { get; init; }

    /// <summary>
    /// When this paragraph's `w:pPr` carries a `w:sectPr` (an in-document section transition), the
    /// anchor of that section break (its own `pt:Unid`, kind `sec`). Null for the common case of a
    /// paragraph with no section transition. Captured by the reader so the markdown projection can
    /// emit the `{#sec:…}` + thematic-break that Word renders at the section boundary, and so the
    /// anchor index carries the `sec` entry — both of which the oracle derives from the same in-pPr
    /// sectPr. The trailing top-level body `w:sectPr` (last-section metadata, not a transition) is a
    /// standalone <see cref="IrSectionBreak"/> block instead, never this field. The paragraph's
    /// content/format hashes are unaffected (the pPr walk already excludes the sectPr); two reads of
    /// the same document yield the same deterministic sectPr Unid here, so determinism is preserved.
    /// </summary>
    public IrAnchor? InlineSectionBreakAnchor { get; init; }

    /// <summary>
    /// The oracle's <c>WmlToMarkdownConverter.IsListItem</c> verdict for this paragraph: a purely
    /// <em>structural</em> predicate — true iff a <c>w:numPr</c> is present inline (<c>w:pPr/w:numPr</c>)
    /// OR anywhere up the <c>pStyle → basedOn</c> chain, <b>regardless of whether that numPr carries a
    /// resolvable <c>numId</c></b>. This is deliberately distinct from <see cref="List"/>
    /// (<see cref="IrListInfo"/>), which is null when the numPr has no numId / unresolvable numId
    /// (genuine "no list membership" semantics). The markdown emitter's trailing-blank-line rule keys
    /// on this structural verdict to reproduce the oracle byte-for-byte for the legal/heading edge
    /// case where a <c>Heading{N}</c>/<c>Subtitle</c> style chain contributes a bare
    /// <c>w:numPr</c>/<c>w:ilvl</c> (no numId): the oracle treats it as a list item for spacing while
    /// the anchor kind is still <c>h</c> and <see cref="List"/> is null. Captured by the reader (which
    /// has the live package) so the emitter never re-walks the style chain. Equality-participating but
    /// fully determined by the paragraph's pPr + the document's styles, so deterministic.
    /// </summary>
    public bool IsListItemForLayout { get; init; }
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

    /// <summary>
    /// True when this row was delivered by a table-level <c>w:sdt</c> wrapping a <c>w:tr</c> (e.g. a
    /// repeating-section content control), rather than being a direct <c>w:tr</c> child of the
    /// <c>w:tbl</c>. Equality-participating (the same table read twice yields the same flag; a row
    /// moving in/out of an SDT wrapper is a structural change the diff engine must see).
    /// <para>
    /// The markdown emitter's table walk excludes SDT-delivered rows so it mirrors the ORACLE
    /// (<c>WmlToMarkdownConverter</c> walks <c>tbl.Elements(w:tr)</c> — direct rows only — so it never
    /// renders an SDT-delivered row). The IR keeps the row (no content loss) and indexes it (the
    /// oracle's anchor index DOES include it, since that walk uses <c>Descendants</c>).
    /// </para>
    /// </summary>
    public bool FromTableSdt { get; init; }
}

/// <summary>
/// A table cell: its block children plus grid span and vertical-merge state.
/// <paramref name="Source"/> is equality-neutral provenance.
/// </summary>
internal sealed record IrCell(IrAnchor Anchor, IrNodeList<IrBlock> Blocks,
                              int GridSpan, IrVMerge VMerge, IrHash ContentHash)
{
    public IrProvenance Source { get; init; } = new();

    /// <summary>
    /// Canonical hash of the cell's SHELL — the whole <c>w:tcPr</c> (width, gridSpan, vMerge, borders,
    /// shading, …); <c>default(IrHash)</c> when the cell has no <c>w:tcPr</c>. Folded into
    /// <see cref="ContentHash"/> by the reader so a shell-only edit is visible to the diff engine, and
    /// exposed here so the N-way composite merger can distinguish WHICH reviewer changed a cell's shell
    /// (attribute a single shell edit / conflict competing ones) without source-element access.
    /// </summary>
    public IrHash ShellDigest { get; init; }

    /// <summary>
    /// True when this cell was delivered by a row-level <c>w:sdt</c> wrapping a <c>w:tc</c>
    /// (the SDT-unwrap discipline in <c>IrReader.BuildRow</c>), rather than being a direct
    /// <c>w:tc</c> child of the <c>w:tr</c>. It is EQUALITY-PARTICIPATING (a positional structural
    /// fact: the same row read twice yields the same flag, and a cell moving in/out of an SDT
    /// wrapper is a genuine structural change the Phase 2 diff engine must see — so the cell is
    /// present in the IR and its ContentHash).
    /// <para>
    /// The markdown emitter's GFM/opaque table walk excludes SDT-delivered cells so it mirrors the
    /// ORACLE exactly: <c>WmlToMarkdownConverter</c>'s table path walks
    /// <c>Elements(w:tr).Elements(w:tc)</c> — direct <c>w:tc</c> children only — so it never sees a
    /// cell an SDT delivers. The IR's richer view keeps the cell (no content loss); the emitter
    /// narrows to the oracle's view for byte parity.
    /// </para>
    /// </summary>
    public bool FromRowSdt { get; init; }
}

/// <summary>A section break carrying its direct section formatting (`w:sectPr`).</summary>
internal sealed record IrSectionBreak : IrBlock
{
    public required IrSectionFormat Format { get; init; }
}

/// <summary>
/// An unmodeled block-level element preserved opaquely. Its <see cref="IrBlock.ContentHash"/> is
/// the canonical hash of the source XML and its <see cref="IrBlock.FormatFingerprint"/> is the
/// cached empty-unmodeled-container digest (it has no modeled formatting).
/// </summary>
internal sealed record IrOpaqueBlock : IrBlock
{
    public required XName ElementName { get; init; }
}

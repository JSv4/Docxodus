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
    /// Back-reference to source OOXML; equality-neutral (does not affect record equality).
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

    /// <summary>Display inlines when custom XML remains opaque in <see cref="Inlines"/> for comparison.
    /// Read as one paragraph so fields can cross wrapper boundaries; excluded from content/format hashes.</summary>
    public IrNodeList<IrInline>? ProjectedInlines { get; init; }

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
    /// The modeled section-format of this paragraph's inline `w:pPr/w:sectPr` (an in-document section
    /// transition), or null when absent. Distinct from <see cref="InlineSectionBreakAnchor"/> (the anchor):
    /// this carries the page-setup PROPERTIES so a mid-document sectPr-only change is diffable. Folded into
    /// <see cref="IrBlock.FormatFingerprint"/> (not <see cref="IrBlock.ContentHash"/> — page metadata, not
    /// content), so such a change classifies FormatOnly and routes to a paired-paragraph emit path.
    /// </summary>
    public IrSectionFormat? InlineSectionFormat { get; init; }

    /// <summary>
    /// Canonical hash of this paragraph's `w:pPr` PROPERTIES — its children except the paragraph-mark
    /// `w:rPr`, the inline `w:sectPr`, and the `w:pPrChange` marker (the pPrChange-comparable projection,
    /// flattened so empty ≡ absent). NOT in <see cref="IrBlock.ContentHash"/> or the fingerprint — a
    /// standalone projection so the N-way composite merger can attribute a paragraph-property change to a
    /// reviewer (and consensus/conflict competing ones) without source-element access, exactly as
    /// <see cref="IrCell.ShellDigest"/> does for a cell shell. Consolidate sub-project B.
    /// </summary>
    public IrHash PPrDigest { get; init; }

    /// <summary>
    /// Resolver-aware canonical digest of this paragraph's outer inline <c>w:sdt</c>/<c>w:smartTag</c>
    /// carriers, including their metadata, full contained OOXML, and position in the inline tree. The reader
    /// deliberately splices those wrappers from <see cref="Inlines"/> to preserve the markdown projection,
    /// so this separate structural fact lets the diff/render paths recognize a carrier-only change and avoid
    /// slicing an atomic wrapper into incompatible revisions. It is intentionally separate from
    /// <see cref="IrBlock.ContentHash"/>: ordinary text alignment remains stable, while a paired edit whose
    /// envelope digest differs is explicitly lowered to a whole-paragraph replacement.
    /// <para><c>default(IrHash)</c> means the paragraph contains no inline carrier; non-empty carrier streams
    /// always have a canonical SHA-256 digest.</para>
    /// </summary>
    public IrHash InlineEnvelopeDigest { get; init; }

    /// <summary>
    /// Canonical digest of the non-hyperlink fields carried by this paragraph: field-code instruction, simple
    /// versus complex representation, inline position, and source scaffolding that normal content/token hashes
    /// intentionally omit. It is deliberately separate from <see cref="IrBlock.ContentHash"/> so a run-based
    /// field result still aligns like its visible text, while a code/state-only change (or any direct simple-field
    /// mutation) is lowered to a reversible whole paragraph replacement instead of leaking the right field into
    /// Reject.
    /// <para>It also records, in order, the field characters of a complex field that spans into or out of the
    /// paragraph (an <c>end</c> or <c>separate</c> with no <c>begin</c> in the paragraph, a <c>begin</c> left open
    /// at its end — a Word table of contents), so moving one of them between paragraphs is a change (issue
    /// #1156).</para>
    /// <para><c>default(IrHash)</c> means no non-hyperlink field occurs in the modeled inline tree and no field spans
    /// into or out of the paragraph.</para>
    /// </summary>
    public IrHash FieldEnvelopeDigest { get; init; }

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

/// <summary>A table: its rows plus per-element digests of its shell (`w:tblPr` and `w:tblGrid`).</summary>
internal sealed record IrTable : IrBlock
{
    public required IrNodeList<IrRow> Rows { get; init; }

    /// <summary>
    /// Canonical hash of the table's `w:tblPr` shell — precisely, all non-`w:tr`/non-`w:sdt`/non-`w:tblGrid`
    /// table children (the `w:tblPr` element plus any stray markup). Folded into
    /// <see cref="IrBlock.FormatFingerprint"/> (not <see cref="IrBlock.ContentHash"/>) so a table-property-only
    /// change classifies FormatOnly; split out from the pre-2026-07-03 single lump so the markup renderer can
    /// attribute a `w:tblPrChange` to exactly the tblPr edit.
    /// </summary>
    public required IrHash TblPrDigest { get; init; }

    /// <summary>Canonical hash of the table's `w:tblGrid` (empty-container hash when absent). Folded into
    /// <see cref="IrBlock.FormatFingerprint"/>; drives `w:tblGridChange` attribution.</summary>
    public required IrHash TblGridDigest { get; init; }
}

/// <summary>
/// A block-level wrapper envelope: a <c>w:sdt</c> content control or a <c>w:customXml</c> element. The
/// wrapper remains one structural block so its metadata and child ordering are never lost;
/// <see cref="Blocks"/> is retained for recursive anchor/flat-text traversal rather than flattening the
/// wrapper into its parent scope.
/// </summary>
/// <remarks>
/// <see cref="EnvelopeDigest"/> is the resolver-aware canonical hash of the OUTER wrapper, not merely
/// <c>w:sdtPr</c>/<c>w:customXmlPr</c>. Thus the modeled identity includes its properties, end properties,
/// nested wrappers, and exact content-tree shape. The parsed child blocks remain equality-participating
/// too, while source provenance remains equality-neutral.
/// </remarks>
internal sealed record IrSdtBlock : IrBlock
{
    /// <summary>Direct block children of <c>w:sdtContent</c> (or of <c>w:customXml</c>, after its
    /// <c>w:customXmlPr</c>), in document order.</summary>
    public required IrNodeList<IrBlock> Blocks { get; init; }

    /// <summary>Resolver-aware canonical digest of the complete outer wrapper.</summary>
    public required IrHash EnvelopeDigest { get; init; }

    /// <summary>True for a <c>w:customXml</c> wrapper, false for a <c>w:sdt</c> content control. The
    /// markdown projection renders a custom-XML wrapper's content but not a content control's.</summary>
    public bool IsCustomXml { get; init; }

    /// <summary>The canonical digest of the wrapper alone, without the blocks it carries: a
    /// <c>w:customXml</c> element's attributes and <c>w:customXmlPr</c>, or a content control's
    /// <c>w:sdtPr</c> and <c>w:sdtEndPr</c>. It tells a change to the wrapper from a change to its content:
    /// the diff aligns the children of two wrappers with equal shells inside one copy of the wrapper.</summary>
    public IrHash ShellDigest { get; init; }

    /// <summary>Whether a bookmark or comment-range marker sits between blocks in the wrapper's content, outside
    /// every paragraph: directly in the wrapper or in a table row or cell inside it. The reader drops such
    /// markers (rule N3), so the wrapper's blocks alone do not describe its content. A nested wrapper's own
    /// markers are its own: the enclosing wrapper keeps or replaces it as one child, markers and all.</summary>
    public bool HasBlockLevelMarkers { get; init; }

    /// <summary>Whether a complex field opens in one of the wrapper's paragraphs and closes in another, as the
    /// TOC field of a table of contents does (it begins in the first entry paragraph and ends paragraphs later).
    /// Such a wrapper is still replaced whole. The flag was added when the paragraph diff moved a begin or end
    /// character between paragraphs without a mark; the paragraph diff now records those characters itself
    /// (<see cref="IrParagraph.FieldEnvelopeDigest"/>, issue #1156), and relaxing this fallback is a separate
    /// change. Like <see cref="HasBlockLevelMarkers"/>, it stops at a nested wrapper.</summary>
    public bool HasParagraphSpanningField { get; init; }

    /// <summary>Whether two wrappers are the same kind with the same own markup (<see cref="ShellDigest"/>), so
    /// only the blocks they carry can differ. This is the wrapper's identity: the aligner pairs two such
    /// wrappers as one edited wrapper, and the semantic diff compares their children, even when
    /// <see cref="CanDiffInPlace"/> says the redline must still draw the pair as a whole replacement.</summary>
    public static bool SameShell(IrSdtBlock left, IrSdtBlock right) =>
        left.IsCustomXml == right.IsCustomXml && left.ShellDigest == right.ShellDigest;

    /// <summary>Whether the diff may align two wrappers' children inside one copy of the wrapper: the same
    /// shell, and on neither side content the children alone would not carry: a bookmark or comment range
    /// between blocks (<see cref="HasBlockLevelMarkers"/>) or a field spanning paragraphs
    /// (<see cref="HasParagraphSpanningField"/>). Otherwise the wrapper is replaced whole, which copies each
    /// side's markers and field characters with it.</summary>
    public static bool CanDiffInPlace(IrSdtBlock left, IrSdtBlock right) =>
        SameShell(left, right) && DiffsInPlace(left) && DiffsInPlace(right);

    private static bool DiffsInPlace(IrSdtBlock wrapper) =>
        !wrapper.HasBlockLevelMarkers && !wrapper.HasParagraphSpanningField;
}

/// <summary>A table row. <paramref name="Source"/> is equality-neutral provenance.</summary>
internal sealed record IrRow(IrAnchor Anchor, IrNodeList<IrCell> Cells, IrHash ContentHash)
{
    public IrProvenance Source { get; init; } = new();

    /// <summary>
    /// The number of leading table-grid columns omitted by this row's cells
    /// (<c>w:trPr/w:gridBefore</c>).  Zero means the first cell starts at grid column zero.
    /// This is equality-participating table geometry, not merely source provenance.
    /// </summary>
    public int GridBefore { get; init; }

    /// <summary>
    /// The number of trailing table-grid columns omitted by this row's cells
    /// (<c>w:trPr/w:gridAfter</c>).  Zero means the last cell reaches the row's declared grid extent.
    /// This is equality-participating table geometry, not merely source provenance.
    /// </summary>
    public int GridAfter { get; init; }

    /// <summary>
    /// Canonical hash of the row's SHELL — all non-`w:tc`/non-`w:sdt` row children (the `w:trPr` element
    /// plus any `w:tblPrEx`/stray markup). Folded into the table's <see cref="IrBlock.FormatFingerprint"/>
    /// (not <see cref="IrBlock.ContentHash"/>) so a row-property-only change classifies FormatOnly. The markup
    /// renderer compares the `w:trPr` element specifically for `w:trPrChange` attribution — a `w:tblPrEx`-only
    /// change flips this digest (so it is detected) but emits no `w:trPrChange` (documented v1 untracked).
    /// </summary>
    public IrHash TrPrDigest { get; init; }

    /// <summary>
    /// Canonical hash of the row's `w:trPr` FLATTENED children ONLY (excludes `w:tblPrEx`; empty ≡ absent).
    /// This is the exactly-trackable subset the markup's `w:trPrChange` attribution and the revision surface
    /// compare — so `GetRevisions`, `Compare`'s markup, and the round-trip contract agree (a `w:tblPrEx`-only
    /// change is untracked in ALL three, not reported by one and ignored by another). NOT in the fingerprint.
    /// </summary>
    public IrHash TrPrShellDigest { get; init; }

    /// <summary>
    /// Canonical hash of the row's `w:tblPrEx` FLATTENED children (empty ≡ absent) — the row-level table
    /// property exceptions, tracked independently of `w:trPr` (`w:tblPrExChange` markup). NOT in the
    /// fingerprint (<see cref="TrPrDigest"/> already carries tblPrEx there); a parallel projection so the
    /// markup + revision surfaces attribute a tblPrEx change without disturbing the trPr digests.
    /// </summary>
    public IrHash TrPrExDigest { get; init; }

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
    /// Canonical hash of the cell's `w:tcPr` FLATTENED children (empty ≡ absent) — the revision-surface
    /// analogue of <see cref="ShellDigest"/>. <see cref="ShellDigest"/> (the whole `w:tcPr` element, folded
    /// into <see cref="IrBlock.ContentHash"/>) distinguishes an empty `&lt;w:tcPr/&gt;` from an absent one;
    /// the markup's `CleanShell` and this digest do NOT, so the `w:tcPrChange` markup and the `TableCell`
    /// revision agree (no spurious revision from an empty-vs-absent shell a reject-cycle can leave). NOT in
    /// <see cref="IrBlock.ContentHash"/> — a separate projection, so the alignment substrate is unchanged.
    /// </summary>
    public IrHash TcPrShellDigest { get; init; }

    /// <summary>
    /// Ordered fold of the cell's blocks' <see cref="IrBlock.FormatFingerprint"/>s — the format-side
    /// counterpart of <see cref="ContentHash"/>. A nested table's <c>w:tblPr</c>/<c>w:tblGrid</c> (or a
    /// paragraph's <c>w:pPr</c>) edit moves no content hash, so <see cref="ContentHash"/> alone cannot
    /// tell the semantic projector that this cell's blocks need descending; gating recursion on this
    /// fingerprint too is what gives a format-only nested-table change its typed records (issue #511).
    /// The cell's own <see cref="ShellDigest"/> is deliberately NOT folded in — it is already compared
    /// (and reported) at the cell level.
    /// </summary>
    public IrHash FormatFingerprint { get; init; }

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

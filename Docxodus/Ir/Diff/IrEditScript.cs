#nullable enable

using System;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// The kind of a block-level edit operation in an <see cref="IrEditScript"/> (M2.2 Task 2). Each kind
/// is the edit-script projection of one <see cref="IrAlignmentKind"/> entry — with the single
/// exception that a <see cref="IrAlignmentKind.Moved"/>/<see cref="IrAlignmentKind.MovedModified"/>
/// alignment entry projects to TWO ops (a source op and a destination op), see <see cref="MoveBlock"/>.
/// </summary>
internal enum IrEditOpKind
{
    /// <summary>
    /// A block unchanged in content AND format (projects an <see cref="IrAlignmentKind.Unchanged"/>
    /// entry). Both <see cref="IrEditOp.LeftAnchor"/> and <see cref="IrEditOp.RightAnchor"/> are set;
    /// <see cref="IrEditOp.TokenDiff"/> is null.
    /// </summary>
    EqualBlock,

    /// <summary>
    /// A block whose text is unchanged but whose block-level formatting differs (projects an
    /// <see cref="IrAlignmentKind.FormatOnly"/> entry). Both anchors set; <see cref="IrEditOp.TokenDiff"/>
    /// is null (the format delta is at block fingerprint granularity — intra-block format-change tokens
    /// only arise inside a <see cref="ModifyBlock"/>).
    /// </summary>
    FormatOnlyBlock,

    /// <summary>
    /// A block present on both sides but neither content- nor format-equal (projects an
    /// <see cref="IrAlignmentKind.Modified"/> entry). Both anchors set. For a PARAGRAPH pair
    /// <see cref="IrEditOp.TokenDiff"/> carries the intra-block token diff; for a TABLE pair
    /// <see cref="IrEditOp.TableDiff"/> carries the nested row/cell diff (M2.2 Task 4); for any other
    /// non-paragraph pair (opaque / section break) both are null (no sub-block model).
    /// </summary>
    ModifyBlock,

    /// <summary>
    /// A right-only block (projects an <see cref="IrAlignmentKind.Inserted"/> entry). Only
    /// <see cref="IrEditOp.RightAnchor"/> is set; <see cref="IrEditOp.LeftAnchor"/> is null.
    /// </summary>
    InsertBlock,

    /// <summary>
    /// A left-only block (projects an <see cref="IrAlignmentKind.Deleted"/> entry). Only
    /// <see cref="IrEditOp.LeftAnchor"/> is set; <see cref="IrEditOp.RightAnchor"/> is null.
    /// </summary>
    DeleteBlock,

    /// <summary>
    /// One side of an exact-content move (projects HALF of an <see cref="IrAlignmentKind.Moved"/>
    /// entry). A move produces TWO <see cref="MoveBlock"/> ops sharing one <see cref="IrEditOp.MoveGroupId"/>:
    /// the SOURCE op (<see cref="IrEditOp.IsMoveSource"/> = true, <see cref="IrEditOp.LeftAnchor"/> set,
    /// emitted at the position the left block would have been deleted from) and the DESTINATION op
    /// (<see cref="IrEditOp.IsMoveSource"/> = false, <see cref="IrEditOp.RightAnchor"/> set, emitted at
    /// the right block's position). <see cref="IrEditOp.TokenDiff"/> is null (a plain move is
    /// exact-content; the destination reproduces the source text verbatim).
    /// </summary>
    MoveBlock,

    /// <summary>
    /// One side of a fuzzy move-and-edit (projects HALF of an <see cref="IrAlignmentKind.MovedModified"/>
    /// entry). Structurally identical to <see cref="MoveBlock"/> (source + destination op sharing a
    /// <see cref="IrEditOp.MoveGroupId"/>) but the DESTINATION op carries a non-null
    /// <see cref="IrEditOp.TokenDiff"/> describing the in-move edit.
    /// <para><b>Reachability.</b> Emitted: the similarity-based cross-gap fuzzy-move detection
    /// (<see cref="IrBlockAligner"/>'s <c>DetectCrossGapMoves</c>) produces
    /// <see cref="IrAlignmentKind.MovedModified"/> for a relocated-and-edited block, which the builder
    /// projects to a <see cref="MoveModifyBlock"/> source + destination pair (the destination carrying the
    /// in-move <see cref="IrEditOp.TokenDiff"/>).</para>
    /// </summary>
    MoveModifyBlock,

    /// <summary>
    /// One LEFT paragraph whose content migrated, in order, across N≥2 RIGHT paragraphs (a paragraph
    /// SPLIT — M2.6). The singular side rides <see cref="IrEditOp.LeftAnchor"/>; the ordered right
    /// anchors ride <see cref="IrEditOp.SplitMergeAnchors"/> with one complete per-segment token diff
    /// each in <see cref="IrEditOp.SegmentDiffs"/> (slice-local left spans; the slices tile the left
    /// token stream exactly — the partition invariant the apply-verifier enforces).
    /// <para><b>N:M is rejected by <c>AssertSplitMergePairing</c> + never emitted by the builder; the
    /// field set physically permits it (nullable fields), so the pairing assert is load-bearing —
    /// a SplitBlock must carry a null <see cref="IrEditOp.RightAnchor"/>.</b></para>
    /// </summary>
    SplitBlock,

    /// <summary>
    /// N≥2 adjacent LEFT paragraphs fused into one RIGHT paragraph (a paragraph MERGE — the byte-mirror
    /// of <see cref="SplitBlock"/>; M2.6). Singular side rides <see cref="IrEditOp.RightAnchor"/>;
    /// the ordered left anchors ride <see cref="IrEditOp.SplitMergeAnchors"/>; <see cref="IrEditOp.SegmentDiffs"/>
    /// holds one diff per left block against the corresponding slice of the right token stream.
    /// <see cref="IrEditOp.LeftAnchor"/> must be null (pairing-assert-enforced; see SplitBlock note).
    /// Shipped alongside split as apply-path CONFIDENCE for the N↔1 reconstruction machinery + fuzzer
    /// coverage — no corpus deviation demands it (the two retained deviations are both splits).
    /// </summary>
    MergeBlock,
}

/// <summary>
/// One block-level edit operation: an anchor-addressed edit referring to a left block, a right block,
/// or both. Anchor strings are the blocks' <see cref="IrAnchor.ToString"/> form (<c>kind:scope:unid</c>),
/// resolvable in the originating document's <see cref="IrDocument.AnchorIndex"/>.
/// </summary>
/// <remarks>
/// Field presence by <see cref="Kind"/>:
/// <list type="bullet">
/// <item><see cref="IrEditOpKind.EqualBlock"/> / <see cref="IrEditOpKind.FormatOnlyBlock"/>: both
/// anchors set; <see cref="TokenDiff"/>, <see cref="MoveGroupId"/>, <see cref="IsMoveSource"/> null.</item>
/// <item><see cref="IrEditOpKind.ModifyBlock"/>: both anchors set; for a paragraph pair
/// <see cref="TokenDiff"/> is non-null and <see cref="TableDiff"/> null; for a TABLE pair
/// <see cref="TableDiff"/> is non-null and <see cref="TokenDiff"/> null (M2.2 Task 4); for any other
/// non-paragraph pair both are null.</item>
/// <item><see cref="IrEditOpKind.InsertBlock"/>: <see cref="RightAnchor"/> set, <see cref="LeftAnchor"/> null.</item>
/// <item><see cref="IrEditOpKind.DeleteBlock"/>: <see cref="LeftAnchor"/> set, <see cref="RightAnchor"/> null.</item>
/// <item><see cref="IrEditOpKind.MoveBlock"/> / <see cref="IrEditOpKind.MoveModifyBlock"/>:
/// <see cref="MoveGroupId"/> and <see cref="IsMoveSource"/> set. The SOURCE op (<see cref="IsMoveSource"/>
/// = true) sets <see cref="LeftAnchor"/>; the DESTINATION op (<see cref="IsMoveSource"/> = false) sets
/// <see cref="RightAnchor"/>. A MoveModify DESTINATION additionally carries <see cref="TokenDiff"/>.</item>
/// <item><see cref="IrEditOpKind.SplitBlock"/>: <see cref="LeftAnchor"/> set, <see cref="RightAnchor"/> null;
/// <see cref="SplitMergeAnchors"/> carries the N≥2 right-doc anchors in document order; <see cref="SegmentDiffs"/>
/// carries one <see cref="IrTokenDiff"/> per right segment (same count), with slice-local left spans that tile
/// the left token stream (partition invariant); all move fields null.</item>
/// <item><see cref="IrEditOpKind.MergeBlock"/>: <see cref="RightAnchor"/> set, <see cref="LeftAnchor"/> null;
/// <see cref="SplitMergeAnchors"/> carries the N≥2 left anchors; <see cref="SegmentDiffs"/> holds one diff
/// per left block against the corresponding slice of the right token stream; mirror of SplitBlock otherwise.</item>
/// <item><see cref="BodyFullRewriteGroupId"/>: set only on the matching InsertBlock/DeleteBlock pair
/// emitted from one BODY-level 1×1 full-lexical-rewrite gap. It is renderer-only provenance; nested
/// scopes and every other operation leave it null.</item>
/// </list>
/// </remarks>
internal sealed record IrEditOp(
    IrEditOpKind Kind,
    string? LeftAnchor,
    string? RightAnchor,
    IrTokenDiff? TokenDiff,
    int? MoveGroupId,
    bool? IsMoveSource,
    IrTableDiff? TableDiff = null,
    IrNodeList<IrTextboxDiff>? TextboxDiffs = null,
    IrNodeList<string>? SplitMergeAnchors = null,
    IrNodeList<IrTokenDiff>? SegmentDiffs = null,
    int? BodyFullRewriteGroupId = null);

/// <summary>
/// The nested inner-block diff of ONE textbox pair inside a Modified paragraph (M2.4 Task 1). A paragraph
/// can contain several textboxes (one <see cref="IrTextbox"/> placeholder token each); when a Modified
/// paragraph pair's textbox placeholders differ, the textboxes are paired POSITIONALLY within the
/// paragraph (the i-th left textbox with the i-th right textbox), their inner block lists aligned with the
/// SAME <see cref="IrBlockAligner.AlignBlocks"/> machinery the body/cell paths use, and the resulting
/// block ops attached here — mirroring the <see cref="IrTableDiff"/> nesting on a Modified table pair.
/// </summary>
/// <remarks>
/// <para><b>Positional pairing.</b> Textboxes pair by their document order within the paragraph. A textbox
/// surplus on one side (the paragraph gained/lost a textbox) yields a diff whose <see cref="Ops"/> are all
/// inserts (right-only) or all deletes (left-only) over that lone textbox's blocks.</para>
/// <para><b>No double-reporting.</b> When a paragraph carries textbox diffs, the paragraph's own
/// <see cref="IrEditOp.TokenDiff"/> is rebuilt to treat the differing placeholder tokens as Equal, so the
/// textbox change is reported ONCE — through these nested ops, not also as a token insert/delete of the
/// opaque placeholder (which has no surface text anyway). Documented on
/// <see cref="IrEditScriptBuilder"/>.</para>
/// </remarks>
internal sealed record IrTextboxDiff(IrNodeList<IrEditOp> Ops);

/// <summary>
/// The block-level diff of ONE note scope (a single footnote or endnote, M2.4 Task 1). Carries the note's
/// kind, its RIGHT-side id (<see cref="NoteId"/> — used to resolve the note in the produced/right document and
/// to stamp the scope context) and its LEFT-side id (<see cref="LeftNoteId"/>), plus the ordered block edit
/// ops produced by aligning the matched note's left/right block lists.
/// </summary>
/// <remarks>
/// <para><b>Whole-note insert/delete.</b> A note present on only one side has no counterpart to align: a
/// right-only note's <see cref="Ops"/> are all <see cref="IrEditOpKind.InsertBlock"/> over its blocks; a
/// left-only note's are all <see cref="IrEditOpKind.DeleteBlock"/>. A matched note runs the full block
/// aligner over its two block lists, exactly like a cell.</para>
/// <para><b>Distinct left/right ids (M2.5 Task 3).</b> Under the oracle's note correspondence a note pairs by
/// body-reference order, NOT by raw <c>w:id</c>, so a matched pair can carry DIFFERENT left and right ids
/// (e.g. an inserted reference shifts the numbering). <see cref="NoteId"/> is the SCOPE id used to resolve the
/// note in the produced/right document and to stamp revisions — the right id for a matched or inserted note,
/// the left id for a deleted-only note. <see cref="LeftNoteId"/> is the left-store id (set for matched and
/// deleted notes, null for a right-only / inserted note). A consumer resolving the LEFT store uses
/// <see cref="LeftNoteId"/>; one resolving the RIGHT/output store uses <see cref="NoteId"/>.</para>
/// </remarks>
internal sealed record IrNoteDiff(IrNoteKind Kind, string NoteId, IrNodeList<IrEditOp> Ops, string? LeftNoteId = null);

/// <summary>
/// The block-level diff of ONE header/footer story (the header/footer scope-diff campaign, 2026-07-03).
/// A story is the effective header (or footer) of one occurrence kind for one body section; stories pair
/// across documents by (section ordinal, kind) with Word's previous-section inheritance rule, then
/// de-duplicate to distinct (left part, right part) pairs — an inherited story referenced by several
/// sections is diffed once, at its first cell.
/// </summary>
/// <remarks>
/// <para><b>Naming convention (the <see cref="IrNoteDiff"/> convention).</b> <see cref="ScopeName"/> is
/// the RIGHT document's scope name (<c>hdr1</c>/<c>ftr2</c>…) for a matched or inserted story — the
/// produced content is right-shaped — and the LEFT scope name for a deleted-only story.
/// <see cref="LeftScopeName"/> is the left scope name when a left story exists, null for a right-only
/// (inserted) story. Scope names are per-document positional labels; the STABLE pairing axis is the
/// (section, kind) cell, which is why both part URIs are carried for the markup renderer.</para>
/// <para><b>Whole-story insert/delete.</b> A story present on only one side has no counterpart to align:
/// a right-only story's <see cref="Ops"/> are all <see cref="IrEditOpKind.InsertBlock"/>; a left-only
/// story's are all <see cref="IrEditOpKind.DeleteBlock"/>. A matched story runs the full block aligner
/// over its two block lists, exactly like a note or a cell.</para>
/// <para><b>Part URIs.</b> <see cref="LeftPartUri"/>/<see cref="RightPartUri"/> are the source part URIs
/// (<see cref="Docxodus.Ir.IrScope.PartUri"/>) — the markup renderer locates the output part (a clone of
/// the left package) by <see cref="LeftPartUri"/> and clones an inserted story's shell/relationships from
/// <see cref="RightPartUri"/>.</para>
/// </remarks>
internal sealed record IrHeaderFooterDiff(
    bool IsHeader,
    IrHeaderFooterKind Kind,
    int SectionIndex,
    string ScopeName,
    string? LeftScopeName,
    Uri? LeftPartUri,
    Uri? RightPartUri,
    IrNodeList<IrEditOp> Ops);

/// <summary>
/// The kind of a row-level operation in an <see cref="IrTableDiff"/> (M2.2 Task 4). Rows carry a
/// <c>ContentHash</c> but no <c>FormatFingerprint</c>, so there is no row-level FormatOnly — the kinds
/// reduce to content classifications.
/// </summary>
internal enum IrRowOpKind
{
    /// <summary>Rows whose <c>ContentHash</c> is equal, on the in-order spine. Both anchors set; no cell ops.</summary>
    EqualRow,

    /// <summary>Rows paired in a gap but content-differing. Both anchors set; <see cref="IrRowOp.CellOps"/> carries the per-cell diff.</summary>
    ModifyRow,

    /// <summary>A right-only row. <see cref="IrRowOp.RightRowAnchor"/> set, left null, no cell ops.</summary>
    InsertRow,

    /// <summary>A left-only row. <see cref="IrRowOp.LeftRowAnchor"/> set, right null, no cell ops.</summary>
    DeleteRow,

    /// <summary>An exact-content row relocated off the spine (only emitted when free; see <see cref="IrTableDiffer"/>).</summary>
    MovedRow,
}

/// <summary>
/// One row-level operation in a table diff. For <see cref="IrRowOpKind.ModifyRow"/>,
/// <see cref="CellOps"/> carries the per-cell diff (ordinary unit-span grids may use a monotone body-key
/// alignment; complex topology retains positional pairing); all other kinds carry a null
/// <see cref="CellOps"/>. <see cref="MoveGroupId"/> links a <see cref="IrRowOpKind.MovedRow"/> source
/// and destination (one row op per side, like the block-level move convention).
/// </summary>
internal sealed record IrRowOp(
    IrRowOpKind Kind,
    string? LeftRowAnchor,
    string? RightRowAnchor,
    IrNodeList<IrCellOp>? CellOps,
    int? MoveGroupId = null,
    bool? IsMoveSource = null);

/// <summary>
/// One cell-level operation inside a <see cref="IrRowOpKind.ModifyRow"/>. Ordinary, direct unit-span rows
/// can pair stable cells by a monotone body-hash spine across a right-only insertion; all other shapes pair
/// positionally (gridSpan/vMerge topology remains a later capability). A paired cell's
/// paragraph blocks are aligned with the same block machinery and each Modified paragraph pair carries a
/// token diff in <see cref="BlockOps"/>, so a cell-text edit surfaces as a token diff IN THE CELL rather
/// than a whole-table blob. An unpaired cell (column count differs) carries one anchor and a null
/// <see cref="BlockOps"/>.
/// </summary>
internal sealed record IrCellOp(
    string? LeftCellAnchor,
    string? RightCellAnchor,
    IrNodeList<IrEditOp>? BlockOps);

/// <summary>
/// The nested diff of one Modified table pair (M2.2 Task 4): an ordered list of <see cref="IrRowOp"/>
/// in right-table row order (with deleted rows interleaved at their left positions, mirroring the
/// block-level convention).
/// </summary>
internal sealed record IrTableDiff(IrNodeList<IrRowOp> RowOps);

/// <summary>
/// The diff-as-data product: an ordered, anchor-addressed, JSON-round-trippable, apply-verifiable
/// sequence of block-level <see cref="IrEditOp"/>s describing how to transform a left
/// <see cref="IrDocument"/>'s body into a right document's body.
/// </summary>
/// <remarks>
/// <para><b>Ordering.</b> Operations mirror the <see cref="IrBlockAligner"/>'s right-document entry
/// order, with one refinement: a <see cref="IrAlignmentKind.Moved"/> alignment entry expands to two ops
/// (source + destination). The DESTINATION op is emitted at the moved entry's position (right order);
/// the SOURCE op is interleaved at the position the moved left block WOULD have occupied under the
/// aligner's left-anchored deletion convention (see <see cref="IrEditScriptBuilder"/>). This makes the
/// script read as a unified diff that both deletes the block at its old position and inserts it at its
/// new one, while the shared <see cref="IrEditOp.MoveGroupId"/> records that the two are one move.</para>
/// <para><b>Apply invariant.</b> Applying the script to the left IR reconstructs the right body at the
/// text level (per-block token text for paragraphs, ContentHash for non-paragraph blocks). This is
/// proven by the test-side <c>IrEditScriptVerifier</c> over every synthetic case and the full WC corpus.</para>
/// <para><b>Note scopes (M2.4 Task 1).</b> <see cref="NoteOps"/> carries the per-note block diffs for the
/// footnote and endnote scopes, in a DETERMINISTIC document order appended AFTER the body
/// <see cref="Operations"/>: footnotes first (by note id, numeric ascending), then endnotes (by note id,
/// numeric ascending). This mirrors <see cref="WmlComparer.GetRevisions"/>'s coverage — body, then
/// footnotes, then endnotes. Each <see cref="IrNoteDiff"/>'s anchors live in the note's own
/// <c>fn</c>/<c>en</c> scope, so they never collide with body anchors.</para>
/// <para><b>Header/footer scopes (2026-07-03 campaign).</b> <see cref="HeaderFooterOps"/> carries the
/// per-story block diffs for the header and footer scopes — a capability the WmlComparer oracle LACKS
/// (it ignores header/footer differences entirely); DocxDiff is strictly ahead here, mirroring Word
/// Compare's own default-on "Headers and footers" granularity. Deterministic order: headers first
/// (section ordinal ascending, then kind Default&lt;First&lt;Even), then footers likewise. Null when no
/// story changed or <see cref="IrDiffSettings.CompareHeadersFooters"/> is false, so pre-campaign scripts
/// serialize byte-identically. Anchors live in the per-part <c>hdrN</c>/<c>ftrN</c> scopes.</para>
/// </remarks>
internal sealed record IrEditScript(
    IrNodeList<IrEditOp> Operations,
    IrNodeList<IrNoteDiff>? NoteOps = null,
    IrNodeList<IrHeaderFooterDiff>? HeaderFooterOps = null);

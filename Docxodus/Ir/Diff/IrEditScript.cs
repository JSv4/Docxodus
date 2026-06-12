#nullable enable

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
    /// <para><b>Reachability.</b> The M2.1/M2.2-Task-2 aligner never emits
    /// <see cref="IrAlignmentKind.MovedModified"/> (similarity-based fuzzy moves are M2.2 Task 3), so
    /// this op kind is UNTESTED-UNTIL-TASK-3. The builder branch that produces it is written now so the
    /// surface is stable; it activates automatically when the aligner starts producing the kind.</para>
    /// </summary>
    MoveModifyBlock,
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
    IrNodeList<IrTextboxDiff>? TextboxDiffs = null);

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
/// kind + id (so a consumer can resolve it back to the store and so the renderer can stamp the scope
/// context) and the ordered block edit ops produced by aligning the matched note's left/right block lists.
/// </summary>
/// <remarks>
/// <para><b>Whole-note insert/delete.</b> A note present on only one side has no counterpart to align: a
/// right-only note's <see cref="Ops"/> are all <see cref="IrEditOpKind.InsertBlock"/> over its blocks; a
/// left-only note's are all <see cref="IrEditOpKind.DeleteBlock"/>. A matched note runs the full block
/// aligner over its two block lists, exactly like a cell.</para>
/// </remarks>
internal sealed record IrNoteDiff(IrNoteKind Kind, string NoteId, IrNodeList<IrEditOp> Ops);

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
/// <see cref="CellOps"/> carries the positional per-cell diff; all other kinds carry a null
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
/// One cell-level operation inside a <see cref="IrRowOpKind.ModifyRow"/>. Cells pair POSITIONALLY within
/// the row (grid-aware pairing — gridSpan/vMerge-aware — is M2.3+; documented). A paired cell's
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
/// numeric ascending). This mirrors <see cref="WmlComparer.GetRevisions"/>'s coverage exactly — body, then
/// footnotes, then endnotes (it does NOT diff header/footer scopes, so neither do we). Each
/// <see cref="IrNoteDiff"/>'s anchors live in the note's own <c>fn</c>/<c>en</c> scope, so they never
/// collide with body anchors.</para>
/// </remarks>
internal sealed record IrEditScript(
    IrNodeList<IrEditOp> Operations,
    IrNodeList<IrNoteDiff>? NoteOps = null);

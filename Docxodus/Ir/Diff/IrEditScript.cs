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
    /// <see cref="IrEditOp.TokenDiff"/> carries the intra-block token diff; for a non-paragraph pair
    /// (table / opaque / section break) it is null in M2.2 Task 2 — table row/cell granularity is Task 4.
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
/// <item><see cref="IrEditOpKind.ModifyBlock"/>: both anchors set; <see cref="TokenDiff"/> non-null for
/// paragraph pairs, null for non-paragraph pairs (Task 4).</item>
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
    bool? IsMoveSource);

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
/// </remarks>
internal sealed record IrEditScript(IrNodeList<IrEditOp> Operations);

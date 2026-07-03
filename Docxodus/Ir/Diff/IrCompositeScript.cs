#nullable enable

using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// One token op tagged with the reviewer who authored it, used when composing multi-reviewer
/// modified paragraphs. <see cref="SourceReviewer"/> indexes the caller's reviewers list;
/// -1 means base-sourced (Equal/Delete spans).
/// </summary>
internal sealed record IrAuthoredTokenOp(IrTokenOp Op, string Author, int SourceReviewer);

/// <summary>
/// Maps one contributing reviewer (<see cref="Reviewer"/>, indexing the caller's reviewers list) to the
/// <c>kind:scope:unid</c> RIGHT anchor of that reviewer's edited paragraph for a composed
/// <see cref="IrCompositeOp"/>. The composite renderer needs each contributing reviewer's OWN right
/// paragraph to source inserted runs: an <see cref="IrAuthoredTokenOp"/>'s Insert RightStart/RightEnd
/// index THAT reviewer's right-token list. <see cref="IrCompositeOp.SourceRightAnchors"/> carries one
/// entry per contributing reviewer so the renderer can resolve each right paragraph independently.
/// </summary>
internal sealed record IrSourceRightAnchor(int Reviewer, string Anchor);

/// <summary>The kind of one composed table cell (see <see cref="IrAuthoredCellOp"/>).</summary>
internal enum IrAuthoredCellKind
{
    /// <summary>A base-paired cell: base passthrough (<see cref="IrAuthoredCellOp.ComposedBlockOps"/> null)
    /// or per-cell composed content. The only kind that existed before column composition.</summary>
    Content,

    /// <summary>A reviewer-INSERTED cell (column add): no base counterpart; the whole cell is cloned from
    /// the reviewer (<see cref="IrAuthoredCellOp.ShellSourceReviewer"/>/<see cref="IrAuthoredCellOp.ShellRightCellAnchor"/>)
    /// and rendered with <c>w:tcPr/w:cellIns</c> + ins-marked content, so accept keeps the new cell and
    /// reject removes it (restoring the base column count).</summary>
    InsertCell,

    /// <summary>A reviewer-DELETED base cell (column remove): rendered as the base cell with
    /// <c>w:tcPr/w:cellDel</c> + del-marked content, so accept removes the cell and reject restores it.</summary>
    DeleteCell,
}

/// <summary>
/// One CELL in a composed multi-reviewer table row (FOLLOW-ON B). <see cref="BaseCellAnchor"/> is the
/// <c>tc:</c> anchor of the cell in the BASE table (null for an <see cref="IrAuthoredCellKind.InsertCell"/>,
/// which has no base counterpart). <see cref="ComposedBlockOps"/> is the cell-local composite of the
/// reviewers' edits to that cell's paragraph blocks — the SAME <see cref="IrCompositeOp"/> shape the body
/// produces, recursively composed at cell-paragraph granularity (so a cell is a mini-body). It is null when
/// no reviewer changed the cell (base passthrough): the renderer then emits the base cell verbatim. The
/// merged apply/JSON truth still lives in the parent op's <see cref="IrEditOp.TableDiff"/>; this is the
/// renderer/revision ATTRIBUTION view.
/// <para><b>Cell shell sourcing.</b> <see cref="ShellSourceReviewer"/>/<see cref="ShellRightCellAnchor"/>
/// select which document supplies the cell's SHELL (<c>w:tcPr</c> — width/gridSpan/vMerge/borders/shading)
/// in the composed output: -1/null (the default) = the BASE cell's shell; a reviewer index + that
/// reviewer's right-cell anchor = the reviewer's shell (set by the merger when exactly one reviewer — or
/// several agreeing reviewers, or a policy-resolved winner — changed the shell, so a width/merge-only edit
/// composes instead of silently reverting to the base shell). For an <see cref="IrAuthoredCellKind.InsertCell"/>
/// they name the inserting reviewer's whole cell; for a <see cref="IrAuthoredCellKind.DeleteCell"/>
/// <see cref="ShellSourceReviewer"/> is the deleting reviewer (attribution only; the shell stays base).</para>
/// <para><see cref="Author"/> is the acting reviewer's name for an InsertCell/DeleteCell (revision
/// attribution on the <c>w:cellIns</c>/<c>w:cellDel</c> marks); "" for Content.</para>
/// </summary>
internal sealed record IrAuthoredCellOp(string? BaseCellAnchor, IrNodeList<IrCompositeOp>? ComposedBlockOps,
    int ShellSourceReviewer = -1, string? ShellRightCellAnchor = null,
    IrAuthoredCellKind Kind = IrAuthoredCellKind.Content, string Author = "");

/// <summary>
/// One ROW in a composed multi-reviewer table (FOLLOW-ON B). <see cref="Kind"/> mirrors
/// <see cref="IrRowOpKind"/>: EqualRow → base row verbatim; InsertRow/DeleteRow → a whole-row insert/delete
/// authored to <see cref="SourceReviewer"/>/<see cref="Author"/>; ModifyRow → per-cell composed via
/// <see cref="ComposedCells"/>. <see cref="BaseRowAnchor"/> is the <c>tr:</c> anchor of the base row this op
/// derives from (null for an InsertRow, which has no base counterpart). <see cref="SourceReviewer"/> -1 =
/// base-sourced (an EqualRow, or a ModifyRow whose cells carry their own per-cell reviewers); a whole
/// Insert/Delete row carries the relocating reviewer's index. <see cref="ComposedCells"/> is non-null only
/// for ModifyRow.
/// </summary>
internal sealed record IrAuthoredRowOp(
    IrRowOpKind Kind,
    string? BaseRowAnchor,
    int SourceReviewer,
    string Author,
    IrNodeList<IrAuthoredCellOp>? ComposedCells,
    string? RightRowAnchor = null);

/// <summary>
/// An edit op tagged with its contributing reviewer. For a composed multi-reviewer Modify,
/// <see cref="Op"/>'s <c>TokenDiff</c> is the MERGED diff (apply/json truth) and
/// <see cref="AuthoredTokens"/> carries per-span authorship for the renderer; for all
/// single-source ops <see cref="AuthoredTokens"/> is null and <see cref="Author"/>/<see cref="SourceReviewer"/>
/// apply. <see cref="SourceReviewer"/> -1 = base-sourced.
/// <para><see cref="ConflictId"/> is non-null when the op is the winner-representative of a conflict
/// (the losing competitors are recorded in <see cref="IrCompositeScript.Conflicts"/>).</para>
/// <para><see cref="SourceRightAnchors"/> is non-null ONLY on a composed multi-reviewer Modify (alongside
/// <see cref="AuthoredTokens"/>): it carries, per contributing reviewer, the right paragraph anchor the
/// renderer must resolve to source that reviewer's inserted runs (their <see cref="IrAuthoredTokenOp"/>
/// Insert spans index that reviewer's right-token list). For all single-source ops it is null and
/// <see cref="Op"/>'s own <c>RightAnchor</c> suffices. Additive/optional — absent from older scripts and
/// from the JSON wire shape, so existing tests/serialization are unaffected.</para>
/// <para><see cref="AuthoredRows"/> is non-null ONLY on a COMPOSED multi-reviewer table (FOLLOW-ON B): the
/// op is a ModifyBlock whose <see cref="Op"/>'s <see cref="IrEditOp.TableDiff"/> remains the MERGED
/// apply/JSON truth (built from the composed row/cell ops' <c>.Op</c> projections) and <see cref="AuthoredRows"/>
/// is the renderer/revision ATTRIBUTION view — per row: EqualRow → base verbatim, Insert/Delete row → authored
/// whole row, ModifyRow → per-cell composed (each cell a recursive <see cref="IrCompositeOp"/> mini-body).
/// Additive/optional — default null preserves all existing behavior and the JSON wire shape.</para>
/// </summary>
internal sealed record IrCompositeOp(
    IrEditOp Op,
    string Author,
    int SourceReviewer,
    IrNodeList<IrAuthoredTokenOp>? AuthoredTokens = null,
    int? ConflictId = null,
    IrNodeList<IrSourceRightAnchor>? SourceRightAnchors = null,
    IrNodeList<IrAuthoredRowOp>? AuthoredRows = null);

/// <summary>
/// One reviewer's competing result for a conflicted span. <see cref="Author"/> is the reviewer
/// name; <see cref="ResultText"/> is the flat text the reviewer's edit would have produced.
/// </summary>
internal sealed record IrConflictCompetitor(string Author, string ResultText);

/// <summary>
/// A base span (<see cref="TokenStart"/>..<see cref="TokenEnd"/>, [x,x) for a block-level conflict)
/// edited DIFFERENTLY by 2+ reviewers. <see cref="Id"/> links back to the <see cref="IrCompositeOp.ConflictId"/>
/// of the winning op in <see cref="IrCompositeScript.Operations"/>. <see cref="AppliedPolicy"/> records
/// which <see cref="ConflictResolution"/> was used; <see cref="Competitors"/> lists every competing edit
/// (including the one that won, for audit purposes).
/// </summary>
internal sealed record IrConflict(
    int Id,
    string BaseAnchor,
    int TokenStart,
    int TokenEnd,
    Docxodus.ConflictResolution AppliedPolicy,
    IrNodeList<IrConflictCompetitor> Competitors);

/// <summary>
/// The N-way composed diff of ONE note scope (a single footnote or endnote) — the consolidate-side
/// counterpart to <see cref="IrNoteDiff"/>. For a BASE-matched note, <see cref="BaseNoteId"/> is the note's
/// id in the SHARED BASE document (the composite output is base-shaped, so this is also the output-part id
/// the ops apply to) and <see cref="Ops"/> is the reviewers' per-block composition over the note's base
/// blocks (the SAME <see cref="IrCompositeOp"/> stream shape the body produces). For a reviewer-INSERTED
/// note, <see cref="BaseNoteId"/> is null, <see cref="SourceReviewer"/>/<see cref="ReviewerNoteId"/> name
/// the inserting reviewer and the note's id in THAT reviewer's document (the renderer clones the definition
/// from there under a fresh output id), and <see cref="Ops"/> are the reviewer's InsertBlock ops authored to
/// them.
/// </summary>
internal sealed record IrCompositeNoteDiff(
    IrNoteKind Kind,
    string? BaseNoteId,
    IrNodeList<IrCompositeOp> Ops,
    int SourceReviewer = -1,
    string? ReviewerNoteId = null);

/// <summary>
/// One reviewer's note-id correspondence entry: the reviewer's note id (<see cref="ReviewerNoteId"/>, the
/// id space that reviewer-SOURCED cloned body content carries on its note references) mapped to the BASE
/// note id it corresponds to (<see cref="BaseNoteId"/>; null for a reviewer-inserted note, which gets a
/// fresh output id at render time). The composite renderer uses these to rewrite reviewer-sourced body
/// references into the base-anchored output id space before the body-order renumber pass — without this, a
/// reviewer whose note numbering shifted (they inserted a note) would leave stale/colliding ids on the refs
/// their cloned content carries. An id absent from a reviewer's map is untouched by that reviewer's diff
/// and therefore already equals the base id (identity).
/// </summary>
internal sealed record IrReviewerNoteIdMap(int Reviewer, IrNoteKind Kind, string ReviewerNoteId, string? BaseNoteId);

/// <summary>
/// The composite diff-as-data product: base-anchor-ordered authored ops + the conflict list.
/// <see cref="NoteOps"/> carries the N-way composed per-note diffs (footnotes/endnotes) when any reviewer
/// edited a note scope, in deterministic order (footnotes then endnotes; base-matched notes by numeric id
/// ascending, then reviewer-inserted notes by reviewer/id) — mirroring the <see cref="IrEditScript.NoteOps"/>
/// convention. <see cref="NoteIdMaps"/> carries every reviewer's note-id correspondence for the renderer's
/// reference rewrite (see <see cref="IrReviewerNoteIdMap"/>); null when no reviewer has any note diff.
/// </summary>
internal sealed record IrCompositeScript(
    IrNodeList<IrCompositeOp> Operations,
    IrNodeList<IrConflict> Conflicts,
    IrNodeList<IrCompositeNoteDiff>? NoteOps = null,
    IrNodeList<IrReviewerNoteIdMap>? NoteIdMaps = null);

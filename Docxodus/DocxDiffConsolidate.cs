#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir.Diff;

namespace Docxodus;

/// <summary>How a consolidate resolves a span edited with DIFFERING edits by two+ reviewers.</summary>
public enum ConflictResolution
{
    /// <summary>Leave the base text at the conflicted span; record every competitor. The default.</summary>
    BaseWins,
    /// <summary>Apply the first reviewer (list order) inline; record the others.</summary>
    FirstReviewerWins,
    /// <summary>Emit each competing edit in sequence at the site; record all.</summary>
    StackAll,
}

/// <summary>
/// One reviewer's contribution to a <see cref="DocxDiff"/> consolidate: their revised copy of the
/// shared base document and the author name stamped on their edits in the consolidated output.
/// </summary>
public sealed class DocxDiffReviewer
{
    /// <summary>
    /// The reviewer's revised copy of the shared base document. Must not be null when passed to a
    /// consolidate call.
    /// </summary>
    public WmlDocument Document { get; set; } = null!;

    /// <summary>
    /// The author name stamped on this reviewer's revisions in the consolidated output.
    /// </summary>
    public string Author { get; set; } = string.Empty;
}

/// <summary>
/// Settings for a <see cref="DocxDiff"/> consolidate: the base diff settings plus the conflict
/// resolution policy. Composes <see cref="DocxDiffSettings"/> rather than inheriting it.
/// </summary>
/// <remarks>
/// <para><b>Composition, not inheritance.</b> <see cref="DocxDiffSettings"/> is <c>sealed</c>, so
/// <see cref="DocxDiffConsolidateSettings"/> wraps it via the <see cref="Diff"/> property. To set
/// per-reviewer author names, use <see cref="DocxDiffReviewer.Author"/>; the
/// <see cref="DocxDiffSettings.AuthorForRevisions"/> on <see cref="Diff"/> is used only for
/// base-vs-reviewer pairwise diffs and may be overridden per reviewer internally.</para>
/// </remarks>
public sealed class DocxDiffConsolidateSettings
{
    /// <summary>
    /// The base diff settings applied to each pairwise base↔reviewer comparison. Defaults to a
    /// fresh <see cref="DocxDiffSettings"/> (all defaults, including deterministic dates).
    /// </summary>
    public DocxDiffSettings Diff { get; set; } = new();

    /// <summary>
    /// How the consolidate resolves a base span edited DIFFERENTLY by two or more reviewers.
    /// Default <see cref="ConflictResolution.BaseWins"/>: the base text is kept and all competing
    /// edits are recorded in the conflict list.
    /// </summary>
    public ConflictResolution ConflictResolution { get; set; } = ConflictResolution.BaseWins;
}

/// <summary>
/// A consumer revision from a <c>DocxDiff.GetConsolidatedRevisions</c> call — the
/// <see cref="DocxDiffRevision"/> shape augmented with the contributing reviewer's identity and,
/// when the revision participates in a conflict, the conflict link.
/// </summary>
/// <remarks>
/// <para><b>Author vs. reviewer.</b> <see cref="Author"/> is always the contributing reviewer's
/// name (from <see cref="DocxDiffReviewer.Author"/>). For equal/base-wins spans it reflects the
/// base.</para>
/// <para><b>Conflict link.</b> <see cref="ConflictId"/> is non-null only when this revision was
/// chosen as the winner of a conflict; the corresponding <see cref="DocxDiffConflict"/> in the
/// conflict list carries all competitors. (Producer wired in Task 3.3.)</para>
/// <para><b>Anchors.</b> <see cref="LeftAnchor"/> addresses the base document; <see cref="RightAnchor"/>
/// addresses the contributing reviewer's document. Presence rules mirror
/// <see cref="DocxDiffRevision"/>.</para>
/// </remarks>
public sealed class DocxDiffConsolidatedRevision
{
    /// <summary>The kind of change.</summary>
    public DocxDiffRevisionType Type { get; }

    /// <summary>The affected text.</summary>
    public string Text { get; }

    /// <summary>The contributing reviewer's author name, or the base author for equal/unchanged spans.</summary>
    public string Author { get; }

    /// <summary>The ISO-8601 revision date (deterministic epoch by default).</summary>
    public string Date { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.Moved"/> revisions, the id linking source and destination;
    /// null otherwise.
    /// </summary>
    public int? MoveGroupId { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.Moved"/> revisions: true at the source (moved FROM here),
    /// false at the destination (moved TO here); null for non-move revisions.
    /// </summary>
    public bool? IsMoveSource { get; }

    /// <summary>
    /// For <see cref="DocxDiffRevisionType.FormatChanged"/> revisions, the old/new modeled properties
    /// and the changed names; null otherwise.
    /// </summary>
    public DocxDiffFormatChange? FormatChange { get; }

    /// <summary>
    /// The base-document block anchor (<c>kind:scope:unid</c>), or null when the revision has no left
    /// side (a pure insertion from a reviewer). Mirrors <see cref="DocxDiffRevision.LeftAnchor"/>.
    /// </summary>
    public string? LeftAnchor { get; }

    /// <summary>
    /// The contributing reviewer's document block anchor (<c>kind:scope:unid</c>), or null when the
    /// revision has no right side (a pure deletion). Mirrors <see cref="DocxDiffRevision.RightAnchor"/>.
    /// </summary>
    public string? RightAnchor { get; }

    /// <summary>
    /// When non-null, this revision is the winning representative of a conflict with this id; look up
    /// the <see cref="DocxDiffConflict"/> with the matching <see cref="DocxDiffConflict.Id"/> in the
    /// conflict list to find all competing edits.
    /// </summary>
    public int? ConflictId { get; }

    internal DocxDiffConsolidatedRevision(
        DocxDiffRevisionType type,
        string text,
        string author,
        string date,
        int? moveGroupId,
        bool? isMoveSource,
        DocxDiffFormatChange? formatChange,
        string? leftAnchor,
        string? rightAnchor,
        int? conflictId)
    {
        Type = type;
        Text = text;
        Author = author;
        Date = date;
        MoveGroupId = moveGroupId;
        IsMoveSource = isMoveSource;
        FormatChange = formatChange;
        LeftAnchor = leftAnchor;
        RightAnchor = rightAnchor;
        ConflictId = conflictId;
    }
}

/// <summary>
/// One conflict from a <c>DocxDiff.Consolidate</c> call: a base span that was edited DIFFERENTLY by
/// two or more reviewers. The <see cref="AppliedPolicy"/> records which
/// <see cref="ConflictResolution"/> was used; <see cref="Competitors"/> lists every competing edit,
/// including the one whose text was placed in the output document (for audit purposes).
/// </summary>
/// <remarks>
/// <para><b>Span encoding.</b> <see cref="TokenStart"/> and <see cref="TokenEnd"/> are token indices
/// into the base paragraph identified by <see cref="BaseAnchor"/>, forming the half-open interval
/// [<see cref="TokenStart"/>, <see cref="TokenEnd"/>). An empty interval (both equal) represents a
/// block-level conflict (the whole block differs, e.g. an inserted vs. deleted paragraph at the same
/// base position). Token indices correspond to the IR tokenizer's output; they are suitable for
/// machine consumers that inspect the edit script — do not rely on them being character offsets.</para>
/// <para><b>Linking to the consolidated revision list.</b> <see cref="Id"/> matches the
/// <see cref="DocxDiffConsolidatedRevision.ConflictId"/> on the winning revision in the consolidated
/// revision list, so conflicts can be correlated to the revision that was actually placed in the
/// output document.</para>
/// </remarks>
public sealed class DocxDiffConflict
{
    /// <summary>
    /// The conflict id; matches <see cref="DocxDiffConsolidatedRevision.ConflictId"/> on the winning
    /// revision in the consolidated revision list.
    /// </summary>
    public int Id { get; }

    /// <summary>
    /// The <c>kind:scope:unid</c> anchor of the base paragraph where the conflict occurs, in the base
    /// document passed to the consolidate call.
    /// For a BLOCK-LEVEL conflict (empty token interval, <see cref="TokenStart"/> == <see cref="TokenEnd"/>),
    /// BaseAnchor is the contested block's anchor — which may be a table (<c>tbl:</c>) or row (<c>tr:</c>)
    /// anchor not present in a paragraph anchor index; token-level conflicts carry a <c>p:</c> anchor that
    /// resolves in the projection/AnchorIndex.
    /// </summary>
    public string BaseAnchor { get; }

    /// <summary>
    /// The (inclusive) start token index of the conflicted span in the base paragraph. Together with
    /// <see cref="TokenEnd"/> this forms the half-open interval [<see cref="TokenStart"/>,
    /// <see cref="TokenEnd"/>). Both equal means a block-level conflict (no specific token span).
    /// </summary>
    public int TokenStart { get; }

    /// <summary>
    /// The (exclusive) end token index of the conflicted span in the base paragraph. See
    /// <see cref="TokenStart"/>.
    /// </summary>
    public int TokenEnd { get; }

    /// <summary>
    /// The conflict resolution policy that was applied to choose the winning text for this conflict.
    /// </summary>
    public ConflictResolution AppliedPolicy { get; }

    /// <summary>
    /// Every competing edit for this conflict (including the winning one, for audit purposes), in
    /// reviewer-list order.
    /// </summary>
    public IReadOnlyList<DocxDiffConflictCompetitor> Competitors { get; }

    internal DocxDiffConflict(
        int id,
        string baseAnchor,
        int tokenStart,
        int tokenEnd,
        ConflictResolution appliedPolicy,
        IReadOnlyList<DocxDiffConflictCompetitor> competitors)
    {
        Id = id;
        BaseAnchor = baseAnchor;
        TokenStart = tokenStart;
        TokenEnd = tokenEnd;
        AppliedPolicy = appliedPolicy;
        Competitors = competitors;
    }

    internal static DocxDiffConflict FromIr(IrConflict c) => new(
        id: c.Id,
        baseAnchor: c.BaseAnchor,
        tokenStart: c.TokenStart,
        tokenEnd: c.TokenEnd,
        appliedPolicy: c.AppliedPolicy,
        competitors: c.Competitors.Select(comp => new DocxDiffConflictCompetitor(comp.Author, comp.ResultText)).ToList());
}

/// <summary>
/// One competitor in a <see cref="DocxDiffConflict"/>: the reviewer who authored the competing edit
/// and the flat text their edit would have produced at the conflicted span.
/// </summary>
public sealed class DocxDiffConflictCompetitor
{
    /// <summary>The reviewer's author name (from <see cref="DocxDiffReviewer.Author"/>).</summary>
    public string Author { get; }

    /// <summary>
    /// The flat text this reviewer's edit would have produced at the conflicted span. Empty string
    /// means a deletion.
    /// </summary>
    public string ResultText { get; }

    internal DocxDiffConflictCompetitor(string author, string resultText)
    {
        Author = author;
        ResultText = resultText;
    }
}

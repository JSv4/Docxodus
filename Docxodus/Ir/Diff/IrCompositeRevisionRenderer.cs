#nullable enable

using System.Collections.Generic;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Renders an <see cref="IrCompositeScript"/> (the N-way merge of several reviewers' pairwise edit scripts
/// against one shared base) into a flat, ordered list of attributed consumer revisions — the consolidate-side
/// counterpart to <see cref="IrRevisionRenderer"/>. Each emitted tuple carries the produced
/// <see cref="IrRevision"/>, the CONTRIBUTING reviewer's author name (which overrides the diff settings'
/// single author), and the conflict id when the revision participates in a conflicted span.
/// </summary>
/// <remarks>
/// <para><b>Per-op attribution.</b> A single-source composite op (its <see cref="IrCompositeOp.AuthoredTokens"/>
/// is null — every InsertBlock/DeleteBlock/ModifyBlock/Move/Split/Merge in v1, plus base-sourced EqualBlock)
/// is projected through the two-way <see cref="IrRevisionRenderer"/> with the RIGHT document pointed at the
/// contributing reviewer's IR (or the base, for a base-sourced op) and the settings' author overridden to the
/// reviewer's name. Every revision the two-way renderer produces for that op inherits the reviewer's author
/// and the op's block-level <see cref="IrCompositeOp.ConflictId"/> (non-null only on a block-level conflict
/// winner under <c>FirstReviewerWins</c>/<c>StackAll</c>).</para>
///
/// <para><b>Composed multi-reviewer paragraph.</b> A composed ModifyBlock (<see cref="IrCompositeOp.AuthoredTokens"/>
/// non-null) projects one revision per authored token op: an Insert span → an Inserted revision authored to the
/// span's reviewer (text from THAT reviewer's right tokens, resolved via
/// <see cref="IrCompositeOp.SourceRightAnchors"/>); a Delete span → a Deleted revision authored to the deleting
/// reviewer (text from the base block); an Equal span → no revision; a FormatChanged span → a FormatChanged
/// revision authored to the span's reviewer. Each revision is stamped with the conflict id of the base-token
/// span it falls within when that span is recorded in <see cref="IrCompositeScript.Conflicts"/> for this base
/// block (so a token-level conflict — which the merger records separately from the op — links to its revisions).</para>
///
/// <para><b>Granularity.</b> Single-source ops honour <see cref="IrDiffSettings.RevisionGranularity"/> via the
/// reused two-way renderer; the composed path is always per-authored-token (Fine), the engine's truth for a
/// multi-author paragraph.</para>
///
/// <para><b>Determinism.</b> Output is a pure function of the (deterministic) composite script, the documents,
/// and the settings — composite ops are emitted in base-anchor order, authored token ops in merge order.</para>
/// </remarks>
internal static class IrCompositeRevisionRenderer
{
    /// <summary>
    /// Render <paramref name="script"/> into attributed revisions. <paramref name="reviewers"/> supplies, in
    /// <see cref="IrCompositeOp.SourceReviewer"/> index order, each reviewer's author + IR (the source the
    /// reviewer's inserted/modified content resolves against). <paramref name="settings"/> supplies
    /// date/granularity; the per-op author comes from the reviewer, not <c>settings</c>.
    /// </summary>
    public static IReadOnlyList<(IrRevision Rev, string Author, int? ConflictId)> Render(
        IrCompositeScript script,
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings)
    {
        // Move-source pre-pass over the WHOLE composite script. Single-source ops are each rendered in
        // their own one-op mini-script (so IrRevisionRenderer honours per-op granularity/author), but a
        // MoveModifyBlock destination's token diff Delete spans index its SOURCE block's tokens — and the
        // two-way renderer recovers that source anchor from the source op (IsMoveSource=true) being present
        // in the same script. In a one-op mini-script the source op is absent, so the destination's
        // in-transit Delete text would render empty. Map each MoveGroupId to its source op here so a
        // destination op can carry its paired source op into the mini-script and stay paired.
        var moveSourceOp = new Dictionary<int, IrEditOp>();
        foreach (var op in script.Operations)
            if (op.AuthoredTokens == null && op.Op.IsMoveSource == true && op.Op.MoveGroupId is { } gid)
                moveSourceOp[gid] = op.Op;

        var result = new List<(IrRevision, string, int?)>();
        foreach (var op in script.Operations)
        {
            if (op.AuthoredRows != null)
                RenderComposedTable(op, baseIr, reviewers, settings, script.Conflicts, moveSourceOp, result);
            else if (op.AuthoredTokens != null)
                RenderComposed(op, baseIr, reviewers, settings, script.Conflicts, result);
            else
                RenderSingleSource(op, baseIr, reviewers, settings, moveSourceOp, result);
        }

        // NOTE SCOPES: project each composed note diff's ops through the same per-op paths, appended after
        // the body ops (footnotes then endnotes — the merger's NoteOps ordering), mirroring the two-way
        // renderer's body-then-notes revision order.
        if (script.NoteOps is { } noteOps)
        {
            foreach (var noteDiff in noteOps)
            {
                foreach (var op in noteDiff.Ops)
                {
                    if (op.AuthoredTokens != null)
                        RenderComposed(op, baseIr, reviewers, settings, script.Conflicts, result);
                    else
                        RenderSingleSource(op, baseIr, reviewers, settings, moveSourceOp, result);
                }
            }
        }
        return result;
    }

    // ------------------------------------------------------------------ composed multi-reviewer table

    /// <summary>
    /// Project a COMPOSED multi-reviewer table (FOLLOW-ON B): walk its <see cref="IrCompositeOp.AuthoredRows"/>,
    /// projecting each cell-block's composite op through the SAME per-block projection the body uses
    /// (<see cref="RenderComposed"/> for a composed paragraph, <see cref="RenderSingleSource"/> otherwise), so
    /// each revision is attributed to its cell-block reviewer and stamped with any cell-level conflict id. An
    /// EqualRow and a base-passthrough cell emit nothing; an InsertRow/DeleteRow projects its whole row through
    /// the single-source path.
    /// </summary>
    private static void RenderComposedTable(
        IrCompositeOp op,
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings,
        IrNodeList<IrConflict> conflicts,
        IReadOnlyDictionary<int, IrEditOp> moveSourceOp,
        List<(IrRevision, string, int?)> sink)
    {
        foreach (var rowOp in op.AuthoredRows!)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                    break; // unchanged base row → no revision

                case IrRowOpKind.InsertRow:
                {
                    // A reviewer-inserted whole row: one Inserted revision per cell paragraph, authored to the
                    // inserting reviewer. The merged TableDiff carries the right row anchor; we surface a
                    // coarse row-level Inserted revision (text from the reviewer's right row) so the row is not
                    // invisible to consumers. Row text is resolved from the reviewer's IR by the right anchor.
                    if (rowOp.SourceReviewer >= 0 && rowOp.SourceReviewer < reviewers.Count
                        && rowOp.RightRowAnchor is { } rra)
                    {
                        string text = RowText(reviewers[rowOp.SourceReviewer].Ir, rra, settings);
                        if (text.Length > 0)
                            sink.Add((
                                new IrRevision(IrRevisionType.Inserted, text, rowOp.Author,
                                    settings.DateTimeForRevisions, RightAnchor: rra),
                                rowOp.Author, op.ConflictId));
                    }
                    break;
                }

                case IrRowOpKind.DeleteRow:
                {
                    // A reviewer-deleted whole base row: one Deleted revision (text from the base row).
                    if (rowOp.BaseRowAnchor is { } bra)
                    {
                        string text = RowText(baseIr, bra, settings);
                        if (text.Length > 0)
                            sink.Add((
                                new IrRevision(IrRevisionType.Deleted, text, rowOp.Author,
                                    settings.DateTimeForRevisions, LeftAnchor: bra),
                                rowOp.Author, op.ConflictId));
                    }
                    break;
                }

                case IrRowOpKind.ModifyRow:
                {
                    if (rowOp.ComposedCells is not { } cells)
                        break;
                    foreach (var cellOp in cells)
                    {
                        // A reviewer-added cell (column add) → one Inserted revision with the reviewer
                        // cell's text; a reviewer-removed base cell (column remove) → one Deleted revision
                        // with the base cell's text — so column edits are visible to consumers.
                        if (cellOp.Kind == IrAuthoredCellKind.InsertCell)
                        {
                            if (cellOp.ShellSourceReviewer >= 0 && cellOp.ShellSourceReviewer < reviewers.Count
                                && cellOp.ShellRightCellAnchor is { } ica)
                            {
                                string text = CellTextByAnchor(reviewers[cellOp.ShellSourceReviewer].Ir, ica, settings);
                                if (text.Length > 0)
                                    sink.Add((
                                        new IrRevision(IrRevisionType.Inserted, text, cellOp.Author,
                                            settings.DateTimeForRevisions, RightAnchor: ica),
                                        cellOp.Author, op.ConflictId));
                            }
                            continue;
                        }
                        if (cellOp.Kind == IrAuthoredCellKind.DeleteCell)
                        {
                            if (cellOp.BaseCellAnchor is { } dca)
                            {
                                string text = CellTextByAnchor(baseIr, dca, settings);
                                if (text.Length > 0)
                                    sink.Add((
                                        new IrRevision(IrRevisionType.Deleted, text, cellOp.Author,
                                            settings.DateTimeForRevisions, LeftAnchor: dca),
                                        cellOp.Author, op.ConflictId));
                            }
                            continue;
                        }
                        if (cellOp.ComposedBlockOps is not { } blockOps)
                            continue; // base passthrough → no revision
                        foreach (var cellBlock in blockOps)
                        {
                            if (cellBlock.AuthoredTokens != null)
                                RenderComposed(cellBlock, baseIr, reviewers, settings, conflicts, sink);
                            else
                                RenderSingleSource(cellBlock, baseIr, reviewers, settings, moveSourceOp, sink);
                        }
                    }
                    break;
                }
            }
        }
    }

    /// <summary>The concatenated paragraph text of ONE cell resolved by anchor (cells are not in
    /// <see cref="IrDocument.AnchorIndex"/>; nested tables recursed), for cell insert/delete revision text.</summary>
    private static string CellTextByAnchor(IrDocument ir, string cellAnchor, IrDiffSettings settings)
    {
        foreach (var block in ir.AnchorIndex.Values)
            if (block is IrTable tbl && FindCellText(tbl, cellAnchor, settings) is { } text)
                return text;
        return string.Empty;
    }

    private static string? FindCellText(IrTable tbl, string cellAnchor, IrDiffSettings settings)
    {
        foreach (var row in tbl.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.Anchor.ToString() == cellAnchor)
                {
                    var sb = new StringBuilder();
                    foreach (var b in cell.Blocks)
                        if (b is IrParagraph p)
                            foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                                sb.Append(t.Text);
                    return sb.ToString();
                }
                foreach (var b in cell.Blocks)
                    if (b is IrTable nested && FindCellText(nested, cellAnchor, settings) is { } text)
                        return text;
            }
        return null;
    }

    /// <summary>The concatenated cell text of a row resolved by anchor, for whole-row insert/delete revision
    /// text. A row is not in <see cref="IrDocument.AnchorIndex"/>, so walk the document's tables.</summary>
    private static string RowText(IrDocument ir, string rowAnchor, IrDiffSettings settings)
    {
        foreach (var block in ir.AnchorIndex.Values)
            if (block is IrTable tbl)
                foreach (var row in tbl.Rows)
                    if (row.Anchor.ToString() == rowAnchor)
                    {
                        var sb = new StringBuilder();
                        foreach (var cell in row.Cells)
                            foreach (var b in cell.Blocks)
                                if (b is IrParagraph p)
                                    foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                                        sb.Append(t.Text);
                        return sb.ToString();
                    }
        return string.Empty;
    }

    // ------------------------------------------------------------------ single-source ops

    /// <summary>
    /// Project a single-source composite op through the two-way <see cref="IrRevisionRenderer"/>: point the
    /// right document at the contributing reviewer (or the base, for a base-sourced op), override the author,
    /// and stamp the op's (block-level) conflict id on every produced revision.
    /// </summary>
    private static void RenderSingleSource(
        IrCompositeOp op,
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings,
        IReadOnlyDictionary<int, IrEditOp> moveSourceOp,
        List<(IrRevision, string, int?)> sink)
    {
        bool baseSourced = op.SourceReviewer < 0;
        var rightIr = baseSourced ? baseIr : reviewers[op.SourceReviewer].Ir;
        string author = baseSourced ? settings.AuthorForRevisions : op.Author;

        // Mini-script projected against (base, contributing reviewer). The author override flows through
        // settings so every revision the two-way renderer stamps carries the reviewer's name. A move
        // DESTINATION op (IsMoveSource==false) needs its paired SOURCE op in the same script so the
        // two-way renderer's move-source pre-pass can resolve the source anchor (otherwise a
        // MoveModifyBlock destination's in-transit Delete text renders empty). Prepend the paired source
        // op when this op is a move destination; the source op is itself emitted separately as its own
        // composite op, so it is rendered exactly once there — here it only seeds the pre-pass map. We
        // therefore keep only the destination op's revisions and drop the (duplicate) source revisions.
        IrEditOp[] miniOps = new[] { op.Op };
        bool isMoveDestination = false;
        if (op.Op.IsMoveSource == false && op.Op.MoveGroupId is { } gid &&
            moveSourceOp.TryGetValue(gid, out var srcOp))
        {
            isMoveDestination = true;
            miniOps = new[] { srcOp, op.Op };
        }

        var oneOpScript = new IrEditScript(IrNodeList.From(miniOps));
        var perOpSettings = settings with { AuthorForRevisions = author };
        var rendered = IrRevisionRenderer.Render(oneOpScript, baseIr, rightIr, perOpSettings);

        foreach (var rev in rendered)
        {
            // When we seeded the pre-pass with the paired source op, suppress the source-side revisions it
            // produced (IsMoveSource==true): the source op renders them itself as its own composite op.
            if (isMoveDestination && rev.IsMoveSource == true)
                continue;
            sink.Add((rev, rev.Author, op.ConflictId));
        }
    }

    // ------------------------------------------------------------------ composed multi-reviewer paragraph

    /// <summary>
    /// Project a composed multi-reviewer ModifyBlock: one revision per authored token op, attributed to the
    /// span's reviewer, with the per-token-span conflict id resolved from the base block's recorded conflicts.
    /// </summary>
    private static void RenderComposed(
        IrCompositeOp op,
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings,
        IrNodeList<IrConflict> conflicts,
        List<(IrRevision, string, int?)> sink)
    {
        var authored = op.AuthoredTokens!;
        string? baseAnchor = op.Op.LeftAnchor;
        var leftTokens = ParagraphTokens(baseAnchor, baseIr, settings);

        // Per-reviewer right paragraph tokens (an authored Insert's RightStart/RightEnd index THAT reviewer's
        // right-token list), resolved from SourceRightAnchors and cached.
        var rightAnchorByReviewer = new Dictionary<int, string>();
        if (op.SourceRightAnchors != null)
            foreach (var sra in op.SourceRightAnchors)
                rightAnchorByReviewer[sra.Reviewer] = sra.Anchor;
        var rightTokensCache = new Dictionary<int, IReadOnlyList<IrDiffToken>>();

        // Conflict spans recorded for this base block, for per-token-span linking.
        var blockConflicts = new List<IrConflict>();
        if (baseAnchor != null)
            foreach (var c in conflicts)
                if (c.BaseAnchor == baseAnchor)
                    blockConflicts.Add(c);

        foreach (var authoredOp in authored)
        {
            var tokenOp = authoredOp.Op;
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
                    break;

                case IrTokenOpKind.Insert:
                {
                    var rightTokens = RightTokensFor(
                        authoredOp.SourceReviewer, rightAnchorByReviewer, reviewers, settings, rightTokensCache);
                    string text = RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    int? cid = ConflictIdAt(blockConflicts, tokenOp.LeftStart);
                    sink.Add((
                        new IrRevision(IrRevisionType.Inserted, text, authoredOp.Author, settings.DateTimeForRevisions,
                            LeftAnchor: baseAnchor,
                            RightAnchor: rightAnchorByReviewer.TryGetValue(authoredOp.SourceReviewer, out var ra) ? ra : null),
                        authoredOp.Author, cid));
                    break;
                }

                case IrTokenOpKind.Delete:
                {
                    string text = RawText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    int? cid = ConflictIdAt(blockConflicts, tokenOp.LeftStart);
                    sink.Add((
                        new IrRevision(IrRevisionType.Deleted, text, authoredOp.Author, settings.DateTimeForRevisions,
                            LeftAnchor: baseAnchor,
                            RightAnchor: rightAnchorByReviewer.TryGetValue(authoredOp.SourceReviewer, out var ra2) ? ra2 : null),
                        authoredOp.Author, cid));
                    break;
                }

                case IrTokenOpKind.FormatChanged:
                {
                    var rightTokens = RightTokensFor(
                        authoredOp.SourceReviewer, rightAnchorByReviewer, reviewers, settings, rightTokensCache);
                    IrRunFormat? oldFmt = tokenOp.LeftStart < leftTokens.Count ? leftTokens[tokenOp.LeftStart].Format : null;
                    IrRunFormat? newFmt = tokenOp.RightStart < rightTokens.Count ? rightTokens[tokenOp.RightStart].Format : null;
                    var details = IrModeledFormat.FormatChangeDetails(oldFmt, newFmt);
                    string text = RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    int? cid = ConflictIdAt(blockConflicts, tokenOp.LeftStart);
                    sink.Add((
                        new IrRevision(IrRevisionType.FormatChanged, text, authoredOp.Author, settings.DateTimeForRevisions,
                            FormatChange: details, LeftAnchor: baseAnchor,
                            RightAnchor: rightAnchorByReviewer.TryGetValue(authoredOp.SourceReviewer, out var ra3) ? ra3 : null),
                        authoredOp.Author, cid));
                    break;
                }
            }
        }
    }

    /// <summary>Resolve (and cache) the right-token list of reviewer <paramref name="reviewer"/>'s paragraph
    /// for the composed base block; an empty list for a reviewer with no recorded right anchor.</summary>
    private static IReadOnlyList<IrDiffToken> RightTokensFor(
        int reviewer,
        IReadOnlyDictionary<int, string> rightAnchorByReviewer,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings,
        Dictionary<int, IReadOnlyList<IrDiffToken>> cache)
    {
        if (cache.TryGetValue(reviewer, out var cached))
            return cached;
        IReadOnlyList<IrDiffToken> tokens = System.Array.Empty<IrDiffToken>();
        if (reviewer >= 0 && reviewer < reviewers.Count && rightAnchorByReviewer.TryGetValue(reviewer, out var anchor))
            tokens = ParagraphTokens(anchor, reviewers[reviewer].Ir, settings);
        cache[reviewer] = tokens;
        return tokens;
    }

    /// <summary>The id of the recorded conflict whose base-token span links <paramref name="basePos"/>, or
    /// null when no conflict span links it.
    /// <para>Two conflict shapes need different matching against an authored token op's
    /// <c>LeftStart</c> anchor:</para>
    /// <list type="bullet">
    /// <item><b>Zero-width insert conflict</b> (<c>[pos, pos)</c>, recorded by
    /// <c>ComposeInsertsAt</c> for insert-vs-insert): the competing inserts are all anchored exactly at
    /// <c>pos</c>, so match the anchor position exactly.</item>
    /// <item><b>Delete/replace conflict</b> (<c>[pos, pos+1)</c>, recorded by <c>ComposeBaseTokenAt</c>):
    /// the deleted base token's revision anchors at <c>pos</c>, but a replacement Insert is anchored at
    /// the differ's <c>Delete.LeftEnd</c> = the right boundary <c>TokenEnd</c> (not <c>TokenStart</c>).
    /// Matching the half-open interval INCLUSIVE of the right boundary links both the delete and its
    /// replacement insert to the conflict.</item>
    /// </list>
    /// <para><b>Over-linking.</b> Including the right boundary means an authored op anchored exactly at a
    /// delete/replace conflict's <c>TokenEnd</c> matches. Equal spans emit no revision (the Equal case
    /// returns early), so they never over-link. In the compose path the op anchored at a delete's
    /// <c>LeftEnd</c> is the replacement Insert keyed to that deleted span (the merger excludes
    /// replacement-Inserts from the standalone insert path and emits them only via the delete path), so
    /// the right-boundary match links the intended replacement rather than an unrelated insert. The only
    /// residual risk — a standalone, non-replacement Insert that coincidentally anchors at a delete
    /// conflict's <c>TokenEnd</c> — does not arise here: such an insert would be anchored at the next base
    /// position and recorded (if conflicting) as its own zero-width conflict at that position.</para></summary>
    private static int? ConflictIdAt(List<IrConflict> blockConflicts, int basePos)
    {
        foreach (var c in blockConflicts)
        {
            bool match = c.TokenStart == c.TokenEnd
                ? basePos == c.TokenStart                          // zero-width insert conflict: match anchor
                : c.TokenStart <= basePos && basePos <= c.TokenEnd; // delete/replace: include right boundary
            if (match)
                return c.Id;
        }
        return null;
    }

    // ------------------------------------------------------------------ token helpers

    /// <summary>Tokens of a paragraph resolved by anchor; empty list for a missing/non-paragraph anchor.</summary>
    private static IReadOnlyList<IrDiffToken> ParagraphTokens(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p)
            return IrDiffTokenizer.Tokenize(p, settings);
        return System.Array.Empty<IrDiffToken>();
    }

    /// <summary>Concatenate the raw <see cref="IrDiffToken.Text"/> over a half-open token span.</summary>
    private static string RawText(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return string.Empty;
        var sb = new StringBuilder();
        for (int i = start; i < end && i < tokens.Count; i++)
            sb.Append(tokens[i].Text);
        return sb.ToString();
    }
}

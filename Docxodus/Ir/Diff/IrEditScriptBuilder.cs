#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Builds an <see cref="IrEditScript"/> from two documents (M2.2 Task 2): runs the
/// <see cref="IrBlockAligner"/>, then projects each alignment entry to one or two block-level edit ops,
/// token-diffing Modified paragraph pairs along the way.
/// </summary>
/// <remarks>
/// <para><b>Move source-interleave rule (deterministic, documented, apply-verifier-proven).</b> The
/// aligner emits ONE entry per <see cref="IrAlignmentKind.Moved"/> pair, at the moved block's RIGHT
/// position. The edit script needs TWO ops — a source (delete-from-old-position) and a destination
/// (insert-at-new-position) — so the script reads as a unified diff. We place them thus:</para>
/// <list type="number">
/// <item>The DESTINATION op (<c>IsMoveSource=false</c>, <c>RightAnchor</c> set) is emitted IN PLACE,
/// at the moved entry's position in the aligner's right-ordered entry list — exactly where the aligner
/// put the entry.</item>
/// <item>The SOURCE op (<c>IsMoveSource=true</c>, <c>LeftAnchor</c> set) is interleaved using the SAME
/// left-anchored unified-diff convention the aligner uses for <see cref="IrAlignmentKind.Deleted"/>
/// entries: the source op trails the op of the nearest PAIRED-IN-PLACE left block preceding the moved
/// left block on the LEFT side; sources preceding every such left block go at the very front, in left
/// order. We reconstruct that adjacency from the alignment entries (which carry the left block of every
/// paired entry) plus the left document's block order, so the rule reuses the aligner's published
/// convention rather than duplicating its private interleave helper.</item>
/// </list>
/// <para><b>MoveGroupId allocation.</b> Ascending starting at 1, assigned in DESTINATION order — i.e.
/// the order moved entries appear in the aligner's right-ordered entry list. Deterministic because the
/// entry order is.</para>
/// <para><b>Determinism.</b> Every step is a pure function of the (deterministic) alignment entries and
/// the left block order; no dictionary iteration order is observed for output.</para>
/// </remarks>
internal static class IrEditScriptBuilder
{
    /// <summary>The left side of a move (source), keyed by the moved left block's body index.</summary>
    private readonly record struct MoveInfo(int GroupId, IrBlock LeftBlock, IrEditOpKind OpKind);

    public static IrEditScript Build(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        ArgumentNullException.ThrowIfNull(settings);

        var alignment = IrBlockAligner.Align(left, right, settings);
        var bodyOps = ProjectAlignment(left.Body.Blocks, alignment, settings);
        var noteOps = BuildNoteOps(left, right, settings);
        return new IrEditScript(IrNodeList.From(bodyOps),
            noteOps.Count == 0 ? null : IrNodeList.From(noteOps));
    }

    // ------------------------------------------------------------------ note scopes (M2.4 Task 1)

    /// <summary>
    /// Diff the footnote and endnote stores of <paramref name="left"/> vs <paramref name="right"/>, in the
    /// DETERMINISTIC document order <see cref="IrEditScript.NoteOps"/> documents: footnotes (by note id,
    /// numeric ascending) then endnotes (by note id, numeric ascending). For each note id present in either
    /// store: a matched note aligns its left/right block lists with the body block aligner and projects the
    /// alignment to block ops (so a footnote-text edit surfaces as a ModifyBlock token diff inside the note,
    /// exactly like a body paragraph); an only-left note becomes all-Deleted blocks; an only-right note
    /// all-Inserted blocks. Mirrors <see cref="WmlComparer.GetRevisions"/>'s footnote+endnote coverage —
    /// header/footer scopes are deliberately NOT diffed (the oracle does not diff them either).
    /// </summary>
    private static List<IrNoteDiff> BuildNoteOps(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        var result = new List<IrNoteDiff>();
        result.AddRange(BuildOneStore(left.Footnotes, right.Footnotes, IrNoteKind.Footnote, settings));
        result.AddRange(BuildOneStore(left.Endnotes, right.Endnotes, IrNoteKind.Endnote, settings));
        return result;
    }

    private static List<IrNoteDiff> BuildOneStore(
        IrNoteStore left, IrNoteStore right, IrNoteKind kind, IrDiffSettings settings)
    {
        // The union of note ids on both sides, ordered numeric-ascending (a non-numeric id sorts last by
        // its ordinal string) so the per-scope op stream is deterministic and matches the oracle's note
        // traversal order (notes are authored/numbered ascending).
        var ids = new SortedSet<string>(left.Notes.Keys.Concat(right.Notes.Keys), NoteIdComparer.Instance);

        var diffs = new List<IrNoteDiff>();
        foreach (var id in ids)
        {
            bool hasLeft = left.Notes.TryGetValue(id, out var leftScope);
            bool hasRight = right.Notes.TryGetValue(id, out var rightScope);

            List<IrEditOp> ops;
            if (hasLeft && hasRight)
            {
                var alignment = IrBlockAligner.AlignBlocks(leftScope!.Blocks, rightScope!.Blocks, settings);
                ops = ProjectAlignment(leftScope.Blocks, alignment, settings);
            }
            else if (hasRight)
            {
                ops = rightScope!.Blocks
                    .Select(b => new IrEditOp(IrEditOpKind.InsertBlock, null, b.Anchor.ToString(), null, null, null))
                    .ToList();
            }
            else
            {
                ops = leftScope!.Blocks
                    .Select(b => new IrEditOp(IrEditOpKind.DeleteBlock, b.Anchor.ToString(), null, null, null, null))
                    .ToList();
            }

            // A matched note whose alignment is entirely EqualBlock/FormatOnly carries no real change; only
            // emit a note diff when something actually changed, so an unedited note produces zero revisions.
            if (ops.Any(o => o.Kind is not IrEditOpKind.EqualBlock))
                diffs.Add(new IrNoteDiff(kind, id, IrNodeList.From(ops)));
        }
        return diffs;
    }

    /// <summary>Numeric-ascending note-id order (id is a <c>w:id</c> integer string); non-numeric ids sort
    /// after all numeric ids by ordinal string, so the order is total and deterministic.</summary>
    private sealed class NoteIdComparer : IComparer<string>
    {
        public static readonly NoteIdComparer Instance = new();

        public int Compare(string? x, string? y)
        {
            bool xn = int.TryParse(x, NumberStyles.Integer, CultureInfo.InvariantCulture, out int xi);
            bool yn = int.TryParse(y, NumberStyles.Integer, CultureInfo.InvariantCulture, out int yi);
            if (xn && yn) return xi.CompareTo(yi);
            if (xn) return -1;
            if (yn) return 1;
            return string.CompareOrdinal(x, y);
        }
    }

    /// <summary>
    /// Project an alignment over <paramref name="leftBlocks"/> into the ordered block edit-op list
    /// (right order, with move/delete sources interleaved). Shared by <see cref="Build"/> (body) and
    /// <see cref="IrTableDiffer"/> (cell block lists) so both produce identical op shapes. Move group ids
    /// are LOCAL to this projection (1..N in destination order) — for cell projections that means ids are
    /// scoped to the cell, which is exactly right since a row/cell move never crosses cells in M2.2.
    /// </summary>
    public static List<IrEditOp> ProjectAlignment(
        IrNodeList<IrBlock> leftBlocks, IrBlockAlignment alignment, IrDiffSettings settings)
    {
        // Left block index by reference identity → used to order move-source interleaving by left position.
        var leftIndex = BuildLeftIndexMap(leftBlocks);

        // Pass 1: assign MoveGroupIds in destination (right-entry) order, ascending from 1, capturing
        // each move's source block + the op kind (MoveBlock vs MoveModifyBlock), keyed by left index.
        var moves = new Dictionary<int, MoveInfo>(); // left-block index → move info
        int nextGroup = 1;
        foreach (var entry in alignment.Entries)
        {
            if (entry.Kind is IrAlignmentKind.Moved or IrAlignmentKind.MovedModified)
            {
                int li = leftIndex[entry.Left!];
                var opKind = entry.Kind == IrAlignmentKind.MovedModified
                    ? IrEditOpKind.MoveModifyBlock
                    : IrEditOpKind.MoveBlock;
                moves[li] = new MoveInfo(nextGroup++, entry.Left!, opKind);
            }
        }

        // Bucket move-source ops by the left index of the nearest preceding paired-in-place left block
        // (left-anchored convention; -1 = front), walking the LEFT document order.
        var sourcesAfterLeft = BuildSourceInterleave(leftBlocks, alignment, leftIndex, moves);

        var ops = new List<IrEditOp>();

        // Front move-sources (those preceding every paired-in-place left block).
        EmitSources(sourcesAfterLeft, -1, moves, ops);

        foreach (var entry in alignment.Entries)
        {
            switch (entry.Kind)
            {
                case IrAlignmentKind.Unchanged:
                    ops.Add(new IrEditOp(IrEditOpKind.EqualBlock,
                        entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                        null, null, null));
                    break;

                case IrAlignmentKind.FormatOnly:
                    ops.Add(new IrEditOp(IrEditOpKind.FormatOnlyBlock,
                        entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                        null, null, null));
                    break;

                case IrAlignmentKind.Modified:
                    ops.Add(MakeModifyOp(entry.Left!, entry.Right!, settings));
                    break;

                case IrAlignmentKind.Inserted:
                    ops.Add(new IrEditOp(IrEditOpKind.InsertBlock,
                        null, entry.Right!.Anchor.ToString(), null, null, null));
                    break;

                case IrAlignmentKind.Deleted:
                    ops.Add(new IrEditOp(IrEditOpKind.DeleteBlock,
                        entry.Left!.Anchor.ToString(), null, null, null, null));
                    break;

                case IrAlignmentKind.Moved:
                case IrAlignmentKind.MovedModified:
                {
                    // Emit the DESTINATION op in place; the SOURCE op was interleaved separately.
                    var move = moves[leftIndex[entry.Left!]];
                    // MoveModifyBlock (from a MovedModified alignment, M2.2 Task 3) carries the in-move
                    // token diff on its destination — tokenize source (left) vs destination (right) so the
                    // op describes "relocated AND edited"; a plain Moved destination carries none.
                    var tokenDiff = move.OpKind == IrEditOpKind.MoveModifyBlock
                        ? TokenDiffFor(entry.Left!, entry.Right!, settings)
                        : null;
                    ops.Add(new IrEditOp(
                        move.OpKind, null, entry.Right!.Anchor.ToString(),
                        tokenDiff, move.GroupId, IsMoveSource: false));
                    break;
                }
            }

            // After a paired-in-place left block's entry, flush move-sources anchored to it.
            if (entry.Left is not null && IsPairedInPlace(entry.Kind))
                EmitSources(sourcesAfterLeft, leftIndex[entry.Left], moves, ops);
        }

        return ops;
    }

    // ------------------------------------------------------------------ modify op (token / table diff)

    /// <summary>
    /// Build a <see cref="IrEditOpKind.ModifyBlock"/> op for a Modified pair. A paragraph pair carries a
    /// token diff; a TABLE pair carries a nested <see cref="IrTableDiff"/> (M2.2 Task 4) — so a cell-text
    /// edit surfaces as a token diff inside the cell, not a whole-table blob; any other non-paragraph
    /// pair (opaque / section break) carries neither.
    /// </summary>
    private static IrEditOp MakeModifyOp(IrBlock leftBlock, IrBlock rightBlock, IrDiffSettings settings)
    {
        if (leftBlock is IrTable lt && rightBlock is IrTable rt)
            return new IrEditOp(IrEditOpKind.ModifyBlock,
                leftBlock.Anchor.ToString(), rightBlock.Anchor.ToString(),
                null, null, null, IrTableDiffer.Diff(lt, rt, settings));

        if (leftBlock is IrParagraph lp && rightBlock is IrParagraph rp)
            return MakeParagraphModifyOp(lp, rp, settings);

        return new IrEditOp(IrEditOpKind.ModifyBlock,
            leftBlock.Anchor.ToString(), rightBlock.Anchor.ToString(),
            TokenDiffFor(leftBlock, rightBlock, settings), null, null);
    }

    /// <summary>
    /// Build the ModifyBlock for a Modified PARAGRAPH pair (M2.4 Task 1: textbox interiors). When both
    /// paragraphs carry textboxes whose placeholder tokens differ, recurse: pair the textboxes positionally
    /// within the paragraph, align each pair's inner blocks, and attach the nested ops as
    /// <see cref="IrEditOp.TextboxDiffs"/> (mirroring the table-diff nesting). The paragraph's OWN token
    /// diff then EXCLUDES the placeholder-token change (the differ keys on a MASKED token list whose textbox
    /// placeholders share one constant key, so they pair as Equal) — the textbox change is reported once,
    /// through the nested ops, never also as an opaque-placeholder token op.
    /// </summary>
    private static IrEditOp MakeParagraphModifyOp(IrParagraph lp, IrParagraph rp, IrDiffSettings settings)
    {
        var leftBoxes = CollectTextboxes(lp.Inlines);
        var rightBoxes = CollectTextboxes(rp.Inlines);

        // Build the textbox diffs (positional pairing + surplus insert/delete). Only keep them when at least
        // one carries a real change; if every box is unchanged we leave the paragraph as a plain token diff.
        var textboxDiffs = BuildTextboxDiffs(leftBoxes, rightBoxes, settings);
        bool nest = textboxDiffs is not null;

        // When nesting, mask the placeholder tokens so the paragraph token diff does not also report the
        // textbox change; otherwise tokenize normally.
        var leftTokens = IrDiffTokenizer.Tokenize(lp, settings);
        var rightTokens = IrDiffTokenizer.Tokenize(rp, settings);
        var diffLeft = nest ? MaskTextboxKeys(leftTokens) : leftTokens;
        var diffRight = nest ? MaskTextboxKeys(rightTokens) : rightTokens;
        var tokenDiff = IrTokenDiffer.Diff(diffLeft, diffRight, settings);

        return new IrEditOp(IrEditOpKind.ModifyBlock,
            lp.Anchor.ToString(), rp.Anchor.ToString(),
            tokenDiff, null, null, null,
            nest ? IrNodeList.From(textboxDiffs!) : null);
    }

    // ------------------------------------------------------------------ textbox interiors (M2.4 Task 1)

    /// <summary>
    /// Collect the <see cref="IrTextbox"/> inlines of a paragraph in DOCUMENT ORDER, recursing transparently
    /// through fields' cached results and hyperlinks exactly as <see cref="IrDiffTokenizer"/> does — so the
    /// i-th collected textbox corresponds to the i-th Textbox placeholder TOKEN, which is what positional
    /// pairing relies on.
    /// </summary>
    private static List<IrTextbox> CollectTextboxes(IReadOnlyList<IrInline> inlines)
    {
        var boxes = new List<IrTextbox>();
        WalkForTextboxes(inlines, boxes);
        return boxes;
    }

    private static void WalkForTextboxes(IReadOnlyList<IrInline> inlines, List<IrTextbox> sink)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextbox tbx:
                    sink.Add(tbx);
                    break;
                case IrFieldRun field:
                    WalkForTextboxes(field.CachedResult, sink);
                    break;
                case IrHyperlink link:
                    WalkForTextboxes(link.Inlines, sink);
                    break;
            }
        }
    }

    /// <summary>
    /// Pair the paragraph's textboxes positionally and diff each pair's inner blocks. Returns null when there
    /// is no real textbox change to report — either side has no textboxes, OR every positionally-paired
    /// textbox is content-equal AND there is no surplus. A surplus textbox (a paragraph gained/lost a box)
    /// yields an all-insert / all-delete inner diff.
    /// </summary>
    private static List<IrTextboxDiff>? BuildTextboxDiffs(
        List<IrTextbox> leftBoxes, List<IrTextbox> rightBoxes, IrDiffSettings settings)
    {
        if (leftBoxes.Count == 0 && rightBoxes.Count == 0)
            return null;

        int paired = Math.Min(leftBoxes.Count, rightBoxes.Count);
        var diffs = new List<IrTextboxDiff>();
        bool anyChange = false;

        for (int i = 0; i < paired; i++)
        {
            var alignment = IrBlockAligner.AlignBlocks(leftBoxes[i].Blocks, rightBoxes[i].Blocks, settings);
            var ops = ProjectAlignment(leftBoxes[i].Blocks, alignment, settings);
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            if (ops.Any(o => o.Kind is not IrEditOpKind.EqualBlock))
                anyChange = true;
        }
        for (int i = paired; i < leftBoxes.Count; i++)
        {
            var ops = leftBoxes[i].Blocks
                .Select(b => new IrEditOp(IrEditOpKind.DeleteBlock, b.Anchor.ToString(), null, null, null, null))
                .ToList();
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            anyChange = anyChange || ops.Count > 0;
        }
        for (int i = paired; i < rightBoxes.Count; i++)
        {
            var ops = rightBoxes[i].Blocks
                .Select(b => new IrEditOp(IrEditOpKind.InsertBlock, null, b.Anchor.ToString(), null, null, null))
                .ToList();
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            anyChange = anyChange || ops.Count > 0;
        }

        return anyChange ? diffs : null;
    }

    /// <summary>The constant match key textbox placeholders collapse to when masked (so a textbox change does
    /// not surface in the paragraph's own token diff — it is reported through the nested ops instead).</summary>
    private const string MaskedTextboxKey = "tbx";

    /// <summary>
    /// Return a token list identical to <paramref name="tokens"/> except every <see cref="IrDiffTokenKind.Textbox"/>
    /// token's <see cref="IrDiffToken.MatchKey"/> is replaced by <see cref="MaskedTextboxKey"/>. Index
    /// positions are preserved, so token-op spans still index the REAL tokens; only equality is neutralized.
    /// </summary>
    private static IReadOnlyList<IrDiffToken> MaskTextboxKeys(IReadOnlyList<IrDiffToken> tokens)
    {
        var masked = new List<IrDiffToken>(tokens.Count);
        foreach (var t in tokens)
            masked.Add(t.Kind == IrDiffTokenKind.Textbox ? t with { MatchKey = MaskedTextboxKey } : t);
        return masked;
    }

    /// <summary>
    /// Token-diff a Modified (or MovedModified) pair. Paragraph pairs are tokenized + Myers-diffed;
    /// non-paragraph pairs other than tables (opaque blocks, section breaks) get a null TokenDiff — they
    /// have no sub-block token model. Tables are handled by <see cref="MakeModifyOp"/> via the table diff.
    /// </summary>
    private static IrTokenDiff? TokenDiffFor(IrBlock leftBlock, IrBlock rightBlock, IrDiffSettings settings)
    {
        if (leftBlock is IrParagraph lp && rightBlock is IrParagraph rp)
        {
            var leftTokens = IrDiffTokenizer.Tokenize(lp, settings);
            var rightTokens = IrDiffTokenizer.Tokenize(rp, settings);
            return IrTokenDiffer.Diff(leftTokens, rightTokens, settings);
        }

        return null;
    }

    // ------------------------------------------------------------------ move interleave

    /// <summary>Map each left block to its index by reference identity (for deterministic ordering).</summary>
    private static Dictionary<IrBlock, int> BuildLeftIndexMap(IrNodeList<IrBlock> blocks)
    {
        var map = new Dictionary<IrBlock, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < blocks.Count; i++)
            map[blocks[i]] = i;
        return map;
    }

    /// <summary>
    /// Bucket each move-source left block under the left index of the nearest preceding PAIRED-IN-PLACE
    /// left block (left-anchored convention; -1 = front). "Paired-in-place" = the left block participated
    /// as the left partner of an Unchanged/FormatOnly/Modified op (a move destination never carries a
    /// left block; a Deleted left block is itself removed and does not anchor). We walk the LEFT document
    /// order so the adjacency exactly mirrors the aligner's deletion interleave.
    /// </summary>
    private static Dictionary<int, List<int>> BuildSourceInterleave(
        IrNodeList<IrBlock> blocks, IrBlockAlignment alignment,
        Dictionary<IrBlock, int> leftIndex, Dictionary<int, MoveInfo> moves)
    {
        var pairedInPlace = new HashSet<int>();
        foreach (var entry in alignment.Entries)
        {
            if (entry.Left is not null && IsPairedInPlace(entry.Kind))
                pairedInPlace.Add(leftIndex[entry.Left]);
        }

        var sourcesAfterLeft = new Dictionary<int, List<int>>();
        int lastPairedLeft = -1;
        for (int i = 0; i < blocks.Count; i++)
        {
            if (moves.ContainsKey(i)) // this left block is a move source
            {
                if (!sourcesAfterLeft.TryGetValue(lastPairedLeft, out var list))
                    sourcesAfterLeft[lastPairedLeft] = list = new List<int>();
                list.Add(i);
            }
            else if (pairedInPlace.Contains(i))
            {
                lastPairedLeft = i;
            }
        }

        return sourcesAfterLeft;
    }

    private static bool IsPairedInPlace(IrAlignmentKind kind) =>
        kind is IrAlignmentKind.Unchanged or IrAlignmentKind.FormatOnly or IrAlignmentKind.Modified;

    /// <summary>Emit the move-SOURCE ops bucketed under <paramref name="anchorLeftIndex"/>, in left order.</summary>
    private static void EmitSources(
        Dictionary<int, List<int>> sourcesAfterLeft, int anchorLeftIndex,
        Dictionary<int, MoveInfo> moves, List<IrEditOp> ops)
    {
        if (!sourcesAfterLeft.TryGetValue(anchorLeftIndex, out var list))
            return;
        foreach (int li in list) // ascending left order
        {
            var move = moves[li];
            // The source op mirrors the destination's kind; the token diff lives only on the destination.
            ops.Add(new IrEditOp(
                move.OpKind, move.LeftBlock.Anchor.ToString(), null, null, move.GroupId, IsMoveSource: true));
        }
    }
}

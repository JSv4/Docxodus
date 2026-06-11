#nullable enable

using System;
using System.Collections.Generic;
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

        // Left block index by reference identity → used to order move-source interleaving by left position.
        var leftIndex = BuildLeftIndexMap(left);

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
        var sourcesAfterLeft = BuildSourceInterleave(left, alignment, leftIndex, moves);

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
                    ops.Add(new IrEditOp(IrEditOpKind.ModifyBlock,
                        entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                        TokenDiffFor(entry.Left!, entry.Right!, settings), null, null));
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

        return new IrEditScript(IrNodeList.From(ops));
    }

    // ------------------------------------------------------------------ token diff

    /// <summary>
    /// Token-diff a Modified (or MovedModified) pair. Paragraph pairs are tokenized + Myers-diffed;
    /// non-paragraph pairs (tables, opaque blocks, section breaks) get a null TokenDiff in M2.2 Task 2 —
    /// table row/cell granularity is Task 4, and the other non-paragraph kinds have no sub-block token model.
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

    /// <summary>Map each left body block to its index by reference identity (for deterministic ordering).</summary>
    private static Dictionary<IrBlock, int> BuildLeftIndexMap(IrDocument left)
    {
        var map = new Dictionary<IrBlock, int>(ReferenceEqualityComparer.Instance);
        var blocks = left.Body.Blocks;
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
        IrDocument left, IrBlockAlignment alignment,
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
        var blocks = left.Body.Blocks;
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

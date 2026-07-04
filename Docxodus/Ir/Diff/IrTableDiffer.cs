#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Structural row/cell diff of a Modified table pair (M2.2 Task 4). Produces an <see cref="IrTableDiff"/>:
/// rows aligned by <c>ContentHash</c>, cells paired positionally within paired rows, and each paired
/// cell's paragraph blocks recursed through the SAME block alignment + token diff machinery — so a
/// cell-text edit surfaces as a token diff inside that cell rather than a whole-table blob.
/// </summary>
/// <remarks>
/// <para><b>Row alignment — self-contained unique-hash + LIS + positional gap fill.</b> Rows carry a
/// <c>ContentHash</c> but no <c>FormatFingerprint</c>, and an <see cref="IrRow"/> is not an
/// <see cref="IrBlock"/>, so the body block aligner's <see cref="IrBlock"/>/fingerprint-keyed machinery
/// does not apply directly. Rather than refactor that aligner around a hash-provider interface (large
/// churn for little reuse), this is a focused row aligner that mirrors the SAME design at row grain:
/// (1) anchor rows whose <c>ContentHash</c> is unique on each side; (2) take the LIS over the anchored
/// pairs by (leftIndex, rightIndex) as the in-order spine = EqualRow; anchored pairs off the spine =
/// MovedRow; (3) gap-fill the remainder positionally — paired rows are ModifyRow, surplus left rows are
/// DeleteRow, surplus right rows InsertRow. Deterministic throughout (integer-indexed, no dictionary
/// enumeration for output).</para>
/// <para><b>Moved rows are "free only".</b> A row is MovedRow exactly when it is an off-spine exact-hash
/// anchor — the same by-construction move the block aligner gets from anchoring. We do NOT run fuzzy
/// cross-gap row moves (that is block-level Task 3 territory; rows rarely relocate-and-edit, and the
/// added cost/false-positive surface is not worth it for M2.2). Documented limitation.</para>
/// <para><b>Cells pair positionally.</b> Within a ModifyRow, cell i on the left pairs with cell i on the
/// right. Grid-aware pairing (gridSpan / vMerge-aware column matching) is M2.3+; a column
/// insert/delete therefore shows as a tail of unpaired cells plus shifted ModifyCell pairs, which is
/// acceptable for M2.2's "cell edits surface as token diffs" goal.</para>
/// </remarks>
internal static class IrTableDiffer
{
    public static IrTableDiff Diff(IrTable left, IrTable right, IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);

        var leftRows = left.Rows;
        var rightRows = right.Rows;
        int nLeft = leftRows.Count;
        int nRight = rightRows.Count;

        var leftKind = new IrRowOpKind?[nLeft];
        var rightKind = new IrRowOpKind?[nRight];
        var leftMatch = new int[nLeft];
        var rightMatch = new int[nRight];
        Array.Fill(leftMatch, -1);
        Array.Fill(rightMatch, -1);

        // --- Anchor: rows whose ContentHash is unique on each side pair up.
        var candidates = CollectRowAnchors(leftRows, rightRows, leftMatch, rightMatch);

        // --- Spine: LIS over anchored pairs (sorted by left index) by right index. On-spine = EqualRow,
        // off-spine = MovedRow.
        candidates.Sort((a, b) => a.Left.CompareTo(b.Left));
        var onSpine = Lis(candidates);
        for (int c = 0; c < candidates.Count; c++)
        {
            var (li, rj) = (candidates[c].Left, candidates[c].Right);
            var kind = onSpine.Contains(c) ? IrRowOpKind.EqualRow : IrRowOpKind.MovedRow;
            leftKind[li] = kind;
            rightKind[rj] = kind;
        }

        var spinePairs = onSpine
            .Select(c => (candidates[c].Left, candidates[c].Right))
            .OrderBy(p => p.Left)
            .ToList();

        // --- Gap fill: positional pairing of the remaining free rows between spine pairs.
        FillRowGaps(leftRows, rightRows, spinePairs, leftKind, rightKind, leftMatch, rightMatch);

        // --- Emit row ops in right order with left-anchored deletion interleave (+ a move group id pass).
        return new IrTableDiff(IrNodeList.From(
            EmitRowOps(leftRows, rightRows, leftKind, rightKind, leftMatch, rightMatch, settings)));
    }

    // ------------------------------------------------------------------ row anchoring / spine

    private readonly record struct RowCand(int Left, int Right);

    private static List<RowCand> CollectRowAnchors(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows, int[] leftMatch, int[] rightMatch)
    {
        var leftByHash = UniqueByHash(leftRows);
        var rightByHash = UniqueByHash(rightRows);
        var candidates = new List<RowCand>();

        for (int i = 0; i < leftRows.Count; i++)
        {
            var h = leftRows[i].ContentHash;
            if (!leftByHash.TryGetValue(h, out int li) || li != i)
                continue;
            if (!rightByHash.TryGetValue(h, out int rj))
                continue;
            leftMatch[i] = rj;
            rightMatch[rj] = i;
            candidates.Add(new RowCand(i, rj));
        }
        return candidates;
    }

    /// <summary>Hash → index for ContentHashes occurring exactly once in the list.</summary>
    private static Dictionary<IrHash, int> UniqueByHash(IrNodeList<IrRow> rows)
    {
        var counts = new Dictionary<IrHash, int>();
        var first = new Dictionary<IrHash, int>();
        for (int i = 0; i < rows.Count; i++)
        {
            var h = rows[i].ContentHash;
            counts[h] = counts.TryGetValue(h, out int c) ? c + 1 : 1;
            if (!first.ContainsKey(h))
                first[h] = i;
        }
        var unique = new Dictionary<IrHash, int>();
        foreach (var kv in first)
            if (counts[kv.Key] == 1)
                unique[kv.Key] = kv.Value;
        return unique;
    }

    /// <summary>LIS by right index over candidates already sorted by left index (patience sort).</summary>
    private static HashSet<int> Lis(List<RowCand> candidates)
    {
        int n = candidates.Count;
        var result = new HashSet<int>();
        if (n == 0)
            return result;

        var tails = new List<int>();
        var prev = new int[n];
        for (int i = 0; i < n; i++)
        {
            prev[i] = -1;
            int right = candidates[i].Right;
            int lo = 0, hi = tails.Count;
            while (lo < hi)
            {
                int mid = (lo + hi) >> 1;
                if (candidates[tails[mid]].Right < right)
                    lo = mid + 1;
                else
                    hi = mid;
            }
            if (lo > 0)
                prev[i] = tails[lo - 1];
            if (lo == tails.Count)
                tails.Add(i);
            else
                tails[lo] = i;
        }
        for (int i = tails[^1]; i != -1; i = prev[i])
            result.Add(i);
        return result;
    }

    // ------------------------------------------------------------------ gap fill

    private static void FillRowGaps(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows,
        List<(int Left, int Right)> spinePairs,
        IrRowOpKind?[] leftKind, IrRowOpKind?[] rightKind, int[] leftMatch, int[] rightMatch)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneRowGap(prevLeft + 1, sl, prevRight + 1, sr, leftKind, rightKind, leftMatch, rightMatch);
            prevLeft = sl;
            prevRight = sr;
        }
        FillOneRowGap(prevLeft + 1, leftRows.Count, prevRight + 1, rightRows.Count,
            leftKind, rightKind, leftMatch, rightMatch);
    }

    /// <summary>
    /// Positional gap fill: free left rows in [leftFrom,leftTo) pair in order with free right rows in
    /// [rightFrom,rightTo) → ModifyRow; the surplus left → DeleteRow, surplus right → InsertRow.
    /// </summary>
    private static void FillOneRowGap(
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        IrRowOpKind?[] leftKind, IrRowOpKind?[] rightKind, int[] leftMatch, int[] rightMatch)
    {
        var freeLeft = new List<int>();
        for (int i = leftFrom; i < leftTo; i++)
            if (leftMatch[i] == -1)
                freeLeft.Add(i);
        var freeRight = new List<int>();
        for (int j = rightFrom; j < rightTo; j++)
            if (rightMatch[j] == -1)
                freeRight.Add(j);

        int paired = Math.Min(freeLeft.Count, freeRight.Count);
        for (int k = 0; k < paired; k++)
        {
            int li = freeLeft[k], rj = freeRight[k];
            leftKind[li] = IrRowOpKind.ModifyRow;
            rightKind[rj] = IrRowOpKind.ModifyRow;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
        }
        for (int k = paired; k < freeLeft.Count; k++)
            leftKind[freeLeft[k]] = IrRowOpKind.DeleteRow;
        for (int k = paired; k < freeRight.Count; k++)
            rightKind[freeRight[k]] = IrRowOpKind.InsertRow;
    }

    // ------------------------------------------------------------------ emit

    private static List<IrRowOp> EmitRowOps(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows,
        IrRowOpKind?[] leftKind, IrRowOpKind?[] rightKind, int[] leftMatch, int[] rightMatch,
        IrDiffSettings settings)
    {
        // Move group ids in destination (right) order, keyed by left row index.
        var moveGroup = new Dictionary<int, int>();
        int nextGroup = 1;
        for (int j = 0; j < rightRows.Count; j++)
            if (rightKind[j] == IrRowOpKind.MovedRow)
                moveGroup[rightMatch[j]] = nextGroup++;

        // Deleted + moved-source rows interleave by the nearest preceding paired-in-place left row.
        var sourcesAfterLeft = new Dictionary<int, List<int>>();
        int lastPaired = -1;
        for (int i = 0; i < leftRows.Count; i++)
        {
            if (leftKind[i] is IrRowOpKind.DeleteRow or IrRowOpKind.MovedRow)
            {
                if (!sourcesAfterLeft.TryGetValue(lastPaired, out var list))
                    sourcesAfterLeft[lastPaired] = list = new List<int>();
                list.Add(i);
            }
            else if (leftMatch[i] != -1) // EqualRow / ModifyRow paired in place
            {
                lastPaired = i;
            }
        }

        var ops = new List<IrRowOp>();
        EmitRowSources(sourcesAfterLeft, -1, leftRows, leftKind, moveGroup, ops);

        for (int j = 0; j < rightRows.Count; j++)
        {
            var kind = rightKind[j] ?? IrRowOpKind.InsertRow;
            int li = rightMatch[j];
            switch (kind)
            {
                case IrRowOpKind.EqualRow:
                    ops.Add(new IrRowOp(IrRowOpKind.EqualRow,
                        leftRows[li].Anchor.ToString(), rightRows[j].Anchor.ToString(), null));
                    break;
                case IrRowOpKind.ModifyRow:
                    ops.Add(new IrRowOp(IrRowOpKind.ModifyRow,
                        leftRows[li].Anchor.ToString(), rightRows[j].Anchor.ToString(),
                        IrNodeList.From(DiffCells(leftRows[li], rightRows[j], settings))));
                    break;
                case IrRowOpKind.MovedRow:
                    ops.Add(new IrRowOp(IrRowOpKind.MovedRow,
                        null, rightRows[j].Anchor.ToString(), null,
                        moveGroup[li], IsMoveSource: false));
                    break;
                case IrRowOpKind.InsertRow:
                    ops.Add(new IrRowOp(IrRowOpKind.InsertRow,
                        null, rightRows[j].Anchor.ToString(), null));
                    break;
            }

            if (li != -1 && (kind == IrRowOpKind.EqualRow || kind == IrRowOpKind.ModifyRow))
                EmitRowSources(sourcesAfterLeft, li, leftRows, leftKind, moveGroup, ops);
        }

        return ops;
    }

    private static void EmitRowSources(
        Dictionary<int, List<int>> sourcesAfterLeft, int anchorLeft,
        IrNodeList<IrRow> leftRows, IrRowOpKind?[] leftKind, Dictionary<int, int> moveGroup, List<IrRowOp> ops)
    {
        if (!sourcesAfterLeft.TryGetValue(anchorLeft, out var list))
            return;
        foreach (int li in list)
        {
            if (leftKind[li] == IrRowOpKind.MovedRow)
                ops.Add(new IrRowOp(IrRowOpKind.MovedRow,
                    leftRows[li].Anchor.ToString(), null, null, moveGroup[li], IsMoveSource: true));
            else
                ops.Add(new IrRowOp(IrRowOpKind.DeleteRow, leftRows[li].Anchor.ToString(), null, null));
        }
    }

    // ------------------------------------------------------------------ cells

    /// <summary>
    /// Positionally pair the cells of two ModifyRow rows and diff each pair's block list. Surplus cells
    /// on either side become single-anchor IrCellOps with null BlockOps (column add/remove).
    /// </summary>
    private static List<IrCellOp> DiffCells(IrRow left, IrRow right, IrDiffSettings settings)
    {
        var leftCells = left.Cells;
        var rightCells = right.Cells;
        int paired = Math.Min(leftCells.Count, rightCells.Count);
        var cellOps = new List<IrCellOp>(Math.Max(leftCells.Count, rightCells.Count));

        for (int k = 0; k < paired; k++)
        {
            var lc = leftCells[k];
            var rc = rightCells[k];
            // Equal-content cells contribute a cell op with no block ops (nothing to recurse).
            var blockOps = lc.ContentHash.Equals(rc.ContentHash)
                ? null
                : IrNodeList.From(DiffCellBlocks(lc, rc, settings));
            cellOps.Add(new IrCellOp(lc.Anchor.ToString(), rc.Anchor.ToString(), blockOps));
        }
        for (int k = paired; k < leftCells.Count; k++)
            cellOps.Add(new IrCellOp(leftCells[k].Anchor.ToString(), null, null));
        for (int k = paired; k < rightCells.Count; k++)
            cellOps.Add(new IrCellOp(null, rightCells[k].Anchor.ToString(), null));

        return cellOps;
    }

    /// <summary>
    /// Align a cell's block lists with the shared block aligner and project to block edit ops (the same
    /// projection the body builder uses), so a cell-text edit lands as a ModifyBlock carrying a token
    /// diff. Cells nest one level: a table-in-a-cell Modified pair recurses again through
    /// <see cref="IrEditScriptBuilder.ProjectBlockOp"/>.
    /// </summary>
    private static List<IrEditOp> DiffCellBlocks(IrCell left, IrCell right, IrDiffSettings settings)
    {
        var alignment = IrBlockAligner.AlignBlocks(left.Blocks, right.Blocks, settings);
        return IrEditScriptBuilder.ProjectAlignment(left.Blocks, alignment, settings);
    }
}

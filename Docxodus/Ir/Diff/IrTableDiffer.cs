#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Structural row/cell diff of a Modified table pair (M2.2 Task 4). Produces an <see cref="IrTableDiff"/>:
/// rows aligned by <c>ContentHash</c>, cells aligned through a conservative ordinary-grid path when
/// possible (otherwise positionally), and each paired cell's paragraph blocks recursed through the SAME
/// block alignment + token diff machinery — so a cell-text edit surfaces as a token diff inside that cell
/// rather than a whole-table blob.
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
/// <para><b>Ordinary-grid cell alignment.</b> For direct, unit-span, non-vertically-merged rows with no
/// <c>gridBefore</c>/<c>gridAfter</c> offset, a unique body-hash + LIS spine preserves stable cells across a
/// right-only cell insertion. The body hash intentionally omits <c>w:tcPr</c>, so a table-grid or width
/// change does not destroy an otherwise stable cell anchor. Every left cell must be on that spine: no
/// positional gap pairing is mixed into this phase, avoiding a false pairing between an edited cell and an
/// adjacent inserted cell. Any unspined left cell, horizontal span, vertical merge, row offset, or
/// SDT-delivered cell retains the conservative positional path. Full gridSpan/vMerge topology is deliberately
/// a later capability.</para>
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

    private readonly record struct IndexCand(int Left, int Right);

    private static List<IndexCand> CollectRowAnchors(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows, int[] leftMatch, int[] rightMatch)
    {
        var leftByHash = UniqueByHash(leftRows);
        var rightByHash = UniqueByHash(rightRows);
        var candidates = new List<IndexCand>();

        for (int i = 0; i < leftRows.Count; i++)
        {
            var h = leftRows[i].ContentHash;
            if (!leftByHash.TryGetValue(h, out int li) || li != i)
                continue;
            if (!rightByHash.TryGetValue(h, out int rj))
                continue;
            leftMatch[i] = rj;
            rightMatch[rj] = i;
            candidates.Add(new IndexCand(i, rj));
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
    private static HashSet<int> Lis(List<IndexCand> candidates)
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
    /// Diff the cells of two ModifyRow rows.  Ordinary rectangular rows first receive a body-key/LIS
    /// alignment that can preserve cells after a right-only insertion; every other shape keeps the
    /// established positional pairing.
    /// </summary>
    private static List<IrCellOp> DiffCells(IrRow left, IrRow right, IrDiffSettings settings)
    {
        if (CanUseOrdinaryGridAlignment(left, right) &&
            TryAlignOrdinaryCellInsertions(left, right, settings, out var aligned))
            return aligned;

        return DiffCellsPositionally(left, right, settings);
    }

    /// <summary>
    /// This is intentionally narrower than general OOXML table-grid alignment.  A unit-span, direct-cell,
    /// no-offset row has a one-to-one physical-cell/grid-column relationship, so a monotone cell spine is
    /// safe.  Spans, vertical merges, grid offsets and row/cell SDT wrappers need topology-aware rendering;
    /// retain the old positional behavior until that layer exists.
    /// </summary>
    private static bool CanUseOrdinaryGridAlignment(IrRow left, IrRow right)
    {
        if (left.FromTableSdt || right.FromTableSdt ||
            left.GridBefore != 0 || left.GridAfter != 0 ||
            right.GridBefore != 0 || right.GridAfter != 0)
            return false;

        foreach (var cell in left.Cells)
            if (cell.GridSpan != 1 || cell.VMerge != IrVMerge.None || cell.FromRowSdt)
                return false;
        foreach (var cell in right.Cells)
            if (cell.GridSpan != 1 || cell.VMerge != IrVMerge.None || cell.FromRowSdt)
                return false;
        return true;
    }

    /// <summary>
    /// Align only safe right-only insertion shapes.  The anchors are unique CELL-BODY hashes (not full
    /// <see cref="IrCell.ContentHash"/>): full hashes deliberately include tcPr, but a column insertion often
    /// changes tcW/table-grid geometry for every otherwise unchanged cell. LIS keeps the anchor spine
    /// monotone, and every left cell must be represented by that spine before any unmatched right cells are
    /// admitted as insertions. This intentionally declines mixed insertion-plus-edit gaps rather than
    /// positionally guessing which right cell is new. A left-only cell remains unsupported by the two-way
    /// renderer, so this method declines the alignment and lets the old conservative path handle it.
    /// </summary>
    private static bool TryAlignOrdinaryCellInsertions(
        IrRow left, IrRow right, IrDiffSettings settings, out List<IrCellOp> cellOps)
    {
        var leftCells = left.Cells;
        var rightCells = right.Cells;
        var leftMatch = new int[leftCells.Count];
        var rightMatch = new int[rightCells.Count];
        Array.Fill(leftMatch, -1);
        Array.Fill(rightMatch, -1);

        var candidates = CollectCellBodyAnchors(leftCells, rightCells);
        candidates.Sort((a, b) => a.Left.CompareTo(b.Left));
        var onSpine = Lis(candidates);
        foreach (int c in onSpine)
        {
            var (li, rj) = (candidates[c].Left, candidates[c].Right);
            leftMatch[li] = rj;
            rightMatch[rj] = li;
        }

        // Pure insertion-only admission: each source cell must be a unique, in-order body anchor. If one
        // source cell was edited or ambiguous, positional gap filling could pair it with a newly inserted
        // right cell (for example A/B/C → A/X/B2/C), producing misleading cellIns topology. Preserve the
        // established conservative path until a topology-aware matcher can prove that case.
        if (leftMatch.Any(match => match < 0))
        {
            cellOps = new List<IrCellOp>();
            return false;
        }

        cellOps = EmitAlignedCellOps(leftCells, rightCells, leftMatch, rightMatch, settings);

        // Do not change the established remove/merge path in this phase.  A right-only cell is the only
        // new renderable shape; without one the aligned output is no more capable than positional pairing.
        return cellOps.Any(c => c.LeftCellAnchor == null) &&
               !cellOps.Any(c => c.RightCellAnchor == null);
    }

    /// <summary>Unique body-hash anchors for cells; tcPr differences are intentionally ignored.</summary>
    private static List<IndexCand> CollectCellBodyAnchors(
        IrNodeList<IrCell> leftCells, IrNodeList<IrCell> rightCells)
    {
        var leftByHash = UniqueCellBodyHashes(leftCells);
        var rightByHash = UniqueCellBodyHashes(rightCells);
        var candidates = new List<IndexCand>();
        for (int i = 0; i < leftCells.Count; i++)
        {
            var h = CellBodyHash(leftCells[i]);
            if (!leftByHash.TryGetValue(h, out int li) || li != i ||
                !rightByHash.TryGetValue(h, out int rj))
                continue;
            candidates.Add(new IndexCand(i, rj));
        }
        return candidates;
    }

    private static Dictionary<IrHash, int> UniqueCellBodyHashes(IrNodeList<IrCell> cells)
    {
        var counts = new Dictionary<IrHash, int>();
        var first = new Dictionary<IrHash, int>();
        for (int i = 0; i < cells.Count; i++)
        {
            var h = CellBodyHash(cells[i]);
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

    /// <summary>
    /// Canonical identity of a cell's block body, deliberately omitting its tcPr shell.  This mirrors the
    /// reader's ContentHash framing so a nested table/image/opaque child remains distinguishable, while a
    /// pure width/gridSpan/shading change cannot destroy an otherwise stable cell anchor.
    /// </summary>
    private static IrHash CellBodyHash(IrCell cell)
    {
        var builder = new IrContentHashBuilder();
        builder.AppendStructure(IrContentHashBuilder.StructureCell);
        foreach (var block in cell.Blocks)
            builder.AppendHash(block.ContentHash);
        return builder.Build();
    }

    /// <summary>
    /// Emit in right-cell order, interleaving any unpaired left source before the next paired right cell.
    /// The ordinary insertion path admits no left-only result, but retaining complete monotone emission here
    /// makes the shape explicit and lets the caller decline unsupported results before rendering.
    /// </summary>
    private static List<IrCellOp> EmitAlignedCellOps(
        IrNodeList<IrCell> leftCells, IrNodeList<IrCell> rightCells,
        int[] leftMatch, int[] rightMatch, IrDiffSettings settings)
    {
        var ops = new List<IrCellOp>(Math.Max(leftCells.Count, rightCells.Count));
        int nextLeft = 0;
        for (int rj = 0; rj < rightCells.Count; rj++)
        {
            int li = rightMatch[rj];
            if (li == -1)
            {
                ops.Add(new IrCellOp(null, rightCells[rj].Anchor.ToString(), null));
                continue;
            }

            while (nextLeft < li)
            {
                ops.Add(new IrCellOp(leftCells[nextLeft].Anchor.ToString(), null, null));
                nextLeft++;
            }
            ops.Add(PairedCellOp(leftCells[li], rightCells[rj], settings));
            nextLeft = li + 1;
        }
        while (nextLeft < leftCells.Count)
        {
            ops.Add(new IrCellOp(leftCells[nextLeft].Anchor.ToString(), null, null));
            nextLeft++;
        }
        return ops;
    }

    /// <summary>
    /// The established positional fallback for spans/merges/offsets and ambiguous ordinary rows.
    /// Surplus cells remain single-anchor operations with null block ops.
    /// </summary>
    private static List<IrCellOp> DiffCellsPositionally(IrRow left, IrRow right, IrDiffSettings settings)
    {
        var leftCells = left.Cells;
        var rightCells = right.Cells;
        int paired = Math.Min(leftCells.Count, rightCells.Count);
        var cellOps = new List<IrCellOp>(Math.Max(leftCells.Count, rightCells.Count));

        for (int k = 0; k < paired; k++)
            cellOps.Add(PairedCellOp(leftCells[k], rightCells[k], settings));
        for (int k = paired; k < leftCells.Count; k++)
            cellOps.Add(new IrCellOp(leftCells[k].Anchor.ToString(), null, null));
        for (int k = paired; k < rightCells.Count; k++)
            cellOps.Add(new IrCellOp(null, rightCells[k].Anchor.ToString(), null));

        return cellOps;
    }

    /// <summary>Build one paired cell op, recursing only when its full content+shell hash differs.</summary>
    private static IrCellOp PairedCellOp(IrCell left, IrCell right, IrDiffSettings settings)
    {
        var blockOps = left.ContentHash.Equals(right.ContentHash)
            ? null
            : IrNodeList.From(DiffCellBlocks(left, right, settings));
        return new IrCellOp(left.Anchor.ToString(), right.Anchor.ToString(), blockOps);
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

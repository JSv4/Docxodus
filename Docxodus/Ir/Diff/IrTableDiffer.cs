#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Structural row/cell diff of a Modified table pair (M2.2 Task 4). Produces an <see cref="IrTableDiff"/>:
/// rows anchored by <c>ContentHash</c> and cost-aligned in ordinary gaps, cells aligned through a conservative
/// ordinary-grid path when possible (otherwise positionally), and each paired cell's paragraph blocks recursed
/// through the SAME block alignment + token diff machinery — so a cell-text edit surfaces as a token diff
/// inside that cell rather than a whole-table blob.
/// </summary>
/// <remarks>
/// <para><b>Row alignment — self-contained unique-hash + LIS + costed gap fill.</b> Rows carry a
/// <c>ContentHash</c> but no <c>FormatFingerprint</c>, and an <see cref="IrRow"/> is not an
/// <see cref="IrBlock"/>, so the body block aligner's <see cref="IrBlock"/>/fingerprint-keyed machinery
/// does not apply directly. Rather than refactor that aligner around a hash-provider interface (large
/// churn for little reuse), this is a focused row aligner that mirrors the SAME design at row grain:
/// (1) anchor rows whose <c>ContentHash</c> is unique on each side; (2) take the LIS over the anchored
/// pairs by (leftIndex, rightIndex) as the in-order spine = EqualRow; anchored pairs off the spine =
/// MovedRow; (3) cost-fill ordinary-grid gaps with a bounded monotone dynamic program. The program first
/// minimizes unpaired rows, then maximizes a cached content affinity, so an inserted row plus an edited
/// retained row does not shift every following pairing. Complex topology and oversized gaps retain the
/// established positional fallback. Deterministic throughout (integer-indexed, no dictionary enumeration
/// for output).</para>
/// <para><b>Moved rows are "free only".</b> A row is MovedRow exactly when it is an off-spine exact-hash
/// anchor — the same by-construction move the block aligner gets from anchoring. We do NOT run fuzzy
/// cross-gap row moves (that is block-level Task 3 territory; rows rarely relocate-and-edit, and the
/// added cost/false-positive surface is not worth it for M2.2). Documented limitation.</para>
/// <para><b>Ordinary-grid cell alignment.</b> For direct, unit-span, non-vertically-merged rows with no
/// <c>gridBefore</c>/<c>gridAfter</c> offset, a unique body-hash + LIS spine plus the same bounded gap DP
/// preserves stable cells across a right-only insertion even when a retained cell was edited. The body hash
/// intentionally omits <c>w:tcPr</c>, so a table-grid or width change does not destroy an otherwise stable
/// cell anchor. The path is admitted only when every left cell is represented and at least one right-only
/// cell remains: native <c>w:cellIns</c> is reversible, while generic left-only <c>w:cellDel</c> topology is
/// not. Any horizontal span, vertical merge, row offset, SDT-delivered cell, left-only result, or oversized
/// gap retains the conservative positional path. Full gridSpan/vMerge topology is deliberately a later
/// capability.</para>
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

        // --- Gap fill: costed monotone pairing for ordinary rows; positional fallback otherwise.
        FillRowGaps(leftRows, rightRows, spinePairs, leftKind, rightKind, leftMatch, rightMatch, settings);

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
        IrRowOpKind?[] leftKind, IrRowOpKind?[] rightKind, int[] leftMatch, int[] rightMatch,
        IrDiffSettings settings)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneRowGap(leftRows, rightRows, prevLeft + 1, sl, prevRight + 1, sr,
                leftKind, rightKind, leftMatch, rightMatch, settings);
            prevLeft = sl;
            prevRight = sr;
        }
        FillOneRowGap(leftRows, rightRows, prevLeft + 1, leftRows.Count, prevRight + 1, rightRows.Count,
            leftKind, rightKind, leftMatch, rightMatch, settings);
    }

    /// <summary>
    /// Cost-fill free rows in one LIS-bounded gap. Direct, unit-span rows use the bounded monotone alignment;
    /// all other shapes preserve positional pairing. The latter is intentionally retained for table topology
    /// the renderer cannot reason about from physical-cell order alone.
    /// </summary>
    private static void FillOneRowGap(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows,
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        IrRowOpKind?[] leftKind, IrRowOpKind?[] rightKind, int[] leftMatch, int[] rightMatch,
        IrDiffSettings settings)
    {
        var freeLeft = new List<int>();
        for (int i = leftFrom; i < leftTo; i++)
            if (leftMatch[i] == -1)
                freeLeft.Add(i);
        var freeRight = new List<int>();
        for (int j = rightFrom; j < rightTo; j++)
            if (rightMatch[j] == -1)
                freeRight.Add(j);

        if (TryAlignOrdinaryRowGap(leftRows, rightRows, freeLeft, freeRight, settings, out var alignment))
        {
            foreach (var step in alignment)
            {
                switch (step.Kind)
                {
                    case MonotoneStepKind.Pair:
                    {
                        int li = freeLeft[step.LeftIndex];
                        int rj = freeRight[step.RightIndex];
                        var kind = leftRows[li].ContentHash.Equals(rightRows[rj].ContentHash)
                            ? IrRowOpKind.EqualRow
                            : IrRowOpKind.ModifyRow;
                        leftKind[li] = kind;
                        rightKind[rj] = kind;
                        leftMatch[li] = rj;
                        rightMatch[rj] = li;
                        break;
                    }
                    case MonotoneStepKind.Delete:
                        leftKind[freeLeft[step.LeftIndex]] = IrRowOpKind.DeleteRow;
                        break;
                    case MonotoneStepKind.Insert:
                        rightKind[freeRight[step.RightIndex]] = IrRowOpKind.InsertRow;
                        break;
                }
            }
            return;
        }

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

    /// <summary>
    /// Direct, unit-span rows have a physical-cell order that is safe to align independently of the table
    /// shell.  The renderer can represent every row-level insert/delete, so unlike cells this alignment may
    /// legitimately produce unmatched units on either side.
    /// </summary>
    private static bool TryAlignOrdinaryRowGap(
        IrNodeList<IrRow> leftRows, IrNodeList<IrRow> rightRows,
        List<int> freeLeft, List<int> freeRight, IrDiffSettings settings,
        out List<MonotoneStep> alignment)
    {
        alignment = new List<MonotoneStep>();
        if (freeLeft.Count == 0 || freeRight.Count == 0 ||
            !FitsDpBudget(freeLeft.Count, freeRight.Count))
            return false;

        foreach (int index in freeLeft)
            if (!CanUseOrdinaryRowAlignment(leftRows[index]))
                return false;
        foreach (int index in freeRight)
            if (!CanUseOrdinaryRowAlignment(rightRows[index]))
                return false;

        var leftBodies = new IrHash[freeLeft.Count];
        var rightBodies = new IrHash[freeRight.Count];
        var leftSignatures = new AlignmentSignature[freeLeft.Count];
        var rightSignatures = new AlignmentSignature[freeRight.Count];
        for (int i = 0; i < freeLeft.Count; i++)
        {
            var row = leftRows[freeLeft[i]];
            leftBodies[i] = RowBodyHash(row);
            leftSignatures[i] = RowSignature(row, settings);
        }
        for (int j = 0; j < freeRight.Count; j++)
        {
            var row = rightRows[freeRight[j]];
            rightBodies[j] = RowBodyHash(row);
            rightSignatures[j] = RowSignature(row, settings);
        }

        alignment = BuildMonotoneAlignment(freeLeft.Count, freeRight.Count,
            (i, j) => BodyAffinity(leftBodies[i], rightBodies[j], leftSignatures[i], rightSignatures[j]));
        return true;
    }

    private static bool CanUseOrdinaryRowAlignment(IrRow row)
    {
        if (row.FromTableSdt || row.GridBefore != 0 || row.GridAfter != 0)
            return false;
        foreach (var cell in row.Cells)
            if (cell.GridSpan != 1 || cell.VMerge != IrVMerge.None || cell.FromRowSdt)
                return false;
        return true;
    }

    // ------------------------------------------------------------------ bounded monotone gap alignment

    // The cap keeps both individual dimensions and the O(gap²) affinity work bounded. Larger gaps retain the
    // old linear positional path instead of turning a malformed or extremely wide table into a throughput cliff.
    private const int MaxDpMatrixCells = 16_384;
    private const int MaxDpGapUnits = 127;
    private const int MaxAffinity = 1_000;
    private const int SignatureEdgeChars = 128;

    private enum MonotoneStepKind : byte
    {
        Pair,
        Delete,
        Insert,
    }

    private readonly record struct MonotoneStep(MonotoneStepKind Kind, int LeftIndex, int RightIndex);

    /// <summary>
    /// Lexicographic score for a gap path.  Unpaired units dominate: this preserves the established
    /// in-place replacement behavior for a genuine rewrite.  Among equally paired paths, the affinity
    /// penalty selects the correspondence that best preserves nearby edited content around an insertion.
    /// </summary>
    private readonly record struct AlignmentScore(int Unpaired, int AffinityPenalty)
    {
        public bool IsStrictlyBetterThan(AlignmentScore other) =>
            Unpaired < other.Unpaired || (Unpaired == other.Unpaired && AffinityPenalty < other.AffinityPenalty);
    }

    private static bool FitsDpBudget(int leftCount, int rightCount) =>
        leftCount <= MaxDpGapUnits && rightCount <= MaxDpGapUnits &&
        (long)(leftCount + 1) * (rightCount + 1) <= MaxDpMatrixCells;

    /// <summary>
    /// Build a deterministic monotone edit alignment.  Pair takes precedence over delete, then insert when
    /// full scores tie; this keeps ambiguous equal-affinity cases stable and closest to historical positional
    /// behavior while still allowing a stronger affinity to move an insertion to its actual position.
    /// </summary>
    private static List<MonotoneStep> BuildMonotoneAlignment(
        int leftCount, int rightCount, Func<int, int, int> affinity)
    {
        int width = rightCount + 1;
        var scores = new AlignmentScore[(leftCount + 1) * width];
        var steps = new MonotoneStepKind[scores.Length];

        for (int i = 1; i <= leftCount; i++)
        {
            scores[i * width] = new AlignmentScore(i, 0);
            steps[i * width] = MonotoneStepKind.Delete;
        }
        for (int j = 1; j <= rightCount; j++)
        {
            scores[j] = new AlignmentScore(j, 0);
            steps[j] = MonotoneStepKind.Insert;
        }

        for (int i = 1; i <= leftCount; i++)
        {
            for (int j = 1; j <= rightCount; j++)
            {
                int affinityScore = affinity(i - 1, j - 1);
                var best = scores[(i - 1) * width + j - 1] with
                {
                    AffinityPenalty = scores[(i - 1) * width + j - 1].AffinityPenalty + (MaxAffinity - affinityScore)
                };
                var bestKind = MonotoneStepKind.Pair;

                var delete = scores[(i - 1) * width + j] with
                {
                    Unpaired = scores[(i - 1) * width + j].Unpaired + 1
                };
                if (delete.IsStrictlyBetterThan(best))
                {
                    best = delete;
                    bestKind = MonotoneStepKind.Delete;
                }

                var insert = scores[i * width + j - 1] with
                {
                    Unpaired = scores[i * width + j - 1].Unpaired + 1
                };
                if (insert.IsStrictlyBetterThan(best))
                {
                    best = insert;
                    bestKind = MonotoneStepKind.Insert;
                }

                scores[i * width + j] = best;
                steps[i * width + j] = bestKind;
            }
        }

        var result = new List<MonotoneStep>(Math.Max(leftCount, rightCount));
        for (int i = leftCount, j = rightCount; i != 0 || j != 0;)
        {
            switch (steps[i * width + j])
            {
                case MonotoneStepKind.Pair:
                    result.Add(new MonotoneStep(MonotoneStepKind.Pair, i - 1, j - 1));
                    i--;
                    j--;
                    break;
                case MonotoneStepKind.Delete:
                    result.Add(new MonotoneStep(MonotoneStepKind.Delete, i - 1, -1));
                    i--;
                    break;
                case MonotoneStepKind.Insert:
                    result.Add(new MonotoneStep(MonotoneStepKind.Insert, -1, j - 1));
                    j--;
                    break;
                default:
                    throw new DocxodusException("Invalid table-gap alignment backpointer.");
            }
        }
        result.Reverse();
        return result;
    }

    /// <summary>
    /// Compact token signature retaining only its leading and trailing material.  This makes affinity
    /// inexpensive for the bounded DP while keeping an insertion/deletion near either edge distinguishable.
    /// Match keys are used rather than display text so the heuristic follows the diff engine's own identity
    /// rules for case folding, links and inline atoms.
    /// </summary>
    private sealed class AlignmentSignature
    {
        private readonly StringBuilder _prefix = new(SignatureEdgeChars);
        private readonly StringBuilder _suffix = new(SignatureEdgeChars);
        private string? _prefixText;
        private string? _suffixText;

        public int Length { get; private set; }

        public void Append(string value)
        {
            if (value.Length == 0)
                return;

            _prefixText = null;
            _suffixText = null;
            Length = value.Length > int.MaxValue - Length ? int.MaxValue : Length + value.Length;
            if (_prefix.Length < SignatureEdgeChars)
            {
                int count = Math.Min(SignatureEdgeChars - _prefix.Length, value.Length);
                _prefix.Append(value, 0, count);
            }

            if (value.Length >= SignatureEdgeChars)
            {
                _suffix.Clear();
                _suffix.Append(value, value.Length - SignatureEdgeChars, SignatureEdgeChars);
            }
            else
            {
                _suffix.Append(value);
                if (_suffix.Length > SignatureEdgeChars)
                    _suffix.Remove(0, _suffix.Length - SignatureEdgeChars);
            }
        }

        public void AppendMarker(char marker)
        {
            _prefixText = null;
            _suffixText = null;
            if (Length != int.MaxValue)
                Length++;
            if (_prefix.Length < SignatureEdgeChars)
                _prefix.Append(marker);
            _suffix.Append(marker);
            if (_suffix.Length > SignatureEdgeChars)
                _suffix.Remove(0, _suffix.Length - SignatureEdgeChars);
        }

        public int AffinityTo(AlignmentSignature other)
        {
            if (Length == 0 || other.Length == 0)
                return 0;

            string leftPrefix = _prefixText ??= _prefix.ToString();
            string rightPrefix = other._prefixText ??= other._prefix.ToString();
            string leftSuffix = _suffixText ??= _suffix.ToString();
            string rightSuffix = other._suffixText ??= other._suffix.ToString();
            int shared = CommonPrefixLength(leftPrefix, rightPrefix) + CommonSuffixLength(leftSuffix, rightSuffix);
            shared = Math.Min(shared, Math.Min(Length, other.Length));
            return (int)Math.Min(MaxAffinity,
                (long)MaxAffinity * 2 * shared / ((long)Length + other.Length));
        }
    }

    private static int CommonPrefixLength(string left, string right)
    {
        int length = Math.Min(left.Length, right.Length);
        int index = 0;
        while (index < length && left[index] == right[index])
            index++;
        return index;
    }

    private static int CommonSuffixLength(string left, string right)
    {
        int leftIndex = left.Length - 1;
        int rightIndex = right.Length - 1;
        int count = 0;
        while (leftIndex >= 0 && rightIndex >= 0 && left[leftIndex] == right[rightIndex])
        {
            leftIndex--;
            rightIndex--;
            count++;
        }
        return count;
    }

    private static int BodyAffinity(
        IrHash leftBody, IrHash rightBody, AlignmentSignature left, AlignmentSignature right) =>
        leftBody.Equals(rightBody) ? MaxAffinity : left.AffinityTo(right);

    private static AlignmentSignature RowSignature(IrRow row, IrDiffSettings settings)
    {
        var signature = new AlignmentSignature();
        foreach (var cell in row.Cells)
            AppendCellSignature(cell, signature, settings);
        return signature;
    }

    private static AlignmentSignature CellSignature(IrCell cell, IrDiffSettings settings)
    {
        var signature = new AlignmentSignature();
        AppendCellSignature(cell, signature, settings);
        return signature;
    }

    private static void AppendCellSignature(IrCell cell, AlignmentSignature signature, IrDiffSettings settings)
    {
        signature.AppendMarker('\u0002'); // cell boundary (not legal XML text, so it cannot alias a MatchKey)
        foreach (var block in cell.Blocks)
        {
            signature.AppendMarker('\u0003'); // block boundary
            if (block is IrParagraph paragraph)
            {
                foreach (var token in IrDiffTokenizer.Tokenize(paragraph, settings))
                {
                    signature.Append(token.MatchKey);
                    signature.AppendMarker('\u0004'); // token boundary
                }
            }
            else
            {
                // A nested table/opaque/SDT has no cheap flat token stream at this grain. Its content hash
                // keeps it distinguishable without recursively launching another table alignment.
                signature.Append(block.ContentHash.ToHex());
            }
        }
    }

    private static IrHash RowBodyHash(IrRow row)
    {
        var builder = new IrContentHashBuilder();
        builder.AppendStructure(IrContentHashBuilder.StructureRow);
        foreach (var cell in row.Cells)
            builder.AppendHash(CellBodyHash(cell));
        return builder.Build();
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
    /// alignment plus bounded monotone gap fill that can preserve cells after a right-only insertion;
    /// every other shape keeps the established positional pairing.
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
        return CanUseOrdinaryRowAlignment(left) && CanUseOrdinaryRowAlignment(right);
    }

    /// <summary>
    /// Align safe right-only insertion shapes.  The anchors are unique CELL-BODY hashes (not full
    /// <see cref="IrCell.ContentHash"/>): full hashes deliberately include tcPr, but a column insertion often
    /// changes tcW/table-grid geometry for every otherwise unchanged cell. LIS keeps the anchor spine
    /// monotone, and a bounded monotone DP cost-fills each free gap by content affinity. This admits a mixed
    /// insertion-plus-edit (for example <c>A/B/C → A/X/B2/C</c>) without shifting B onto X. A left-only cell
    /// remains unsupported by the two-way renderer, so any such result declines this path and lets the old
    /// conservative positional path handle it.
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

        var spinePairs = onSpine
            .Select(c => (candidates[c].Left, candidates[c].Right))
            .OrderBy(pair => pair.Left)
            .ToList();
        if (!TryFillOrdinaryCellGaps(leftCells, rightCells, spinePairs, leftMatch, rightMatch, settings))
        {
            cellOps = new List<IrCellOp>();
            return false;
        }

        // Native cell insertion is reversible with tblGridChange; generic cell deletion is not. Require every
        // left cell to have a source pair, then admit only a real right-growth result. Count-stable ordinary
        // rows retain their historical positional path until move/topology semantics are designed explicitly.
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

    private static bool TryFillOrdinaryCellGaps(
        IrNodeList<IrCell> leftCells, IrNodeList<IrCell> rightCells,
        List<(int Left, int Right)> spinePairs, int[] leftMatch, int[] rightMatch,
        IrDiffSettings settings)
    {
        int prevLeft = -1;
        int prevRight = -1;
        foreach (var (left, right) in spinePairs)
        {
            if (!TryFillOneOrdinaryCellGap(leftCells, rightCells, prevLeft + 1, left,
                    prevRight + 1, right, leftMatch, rightMatch, settings))
                return false;
            prevLeft = left;
            prevRight = right;
        }
        return TryFillOneOrdinaryCellGap(leftCells, rightCells, prevLeft + 1, leftCells.Count,
            prevRight + 1, rightCells.Count, leftMatch, rightMatch, settings);
    }

    /// <summary>
    /// Fill one free cell gap.  Empty-sided gaps are already unambiguous; nonempty two-sided gaps use the
    /// shared bounded DP.  A matrix beyond the budget deliberately returns false so callers retain the
    /// established positional behavior rather than paying quadratic cost on an adversarially wide row.
    /// </summary>
    private static bool TryFillOneOrdinaryCellGap(
        IrNodeList<IrCell> leftCells, IrNodeList<IrCell> rightCells,
        int leftFrom, int leftTo, int rightFrom, int rightTo, int[] leftMatch, int[] rightMatch,
        IrDiffSettings settings)
    {
        var freeLeft = new List<int>();
        for (int i = leftFrom; i < leftTo; i++)
            if (leftMatch[i] == -1)
                freeLeft.Add(i);
        var freeRight = new List<int>();
        for (int j = rightFrom; j < rightTo; j++)
            if (rightMatch[j] == -1)
                freeRight.Add(j);

        if (freeLeft.Count == 0 || freeRight.Count == 0)
            return true;
        if (!FitsDpBudget(freeLeft.Count, freeRight.Count))
            return false;

        var leftBodies = new IrHash[freeLeft.Count];
        var rightBodies = new IrHash[freeRight.Count];
        var leftSignatures = new AlignmentSignature[freeLeft.Count];
        var rightSignatures = new AlignmentSignature[freeRight.Count];
        for (int i = 0; i < freeLeft.Count; i++)
        {
            var cell = leftCells[freeLeft[i]];
            leftBodies[i] = CellBodyHash(cell);
            leftSignatures[i] = CellSignature(cell, settings);
        }
        for (int j = 0; j < freeRight.Count; j++)
        {
            var cell = rightCells[freeRight[j]];
            rightBodies[j] = CellBodyHash(cell);
            rightSignatures[j] = CellSignature(cell, settings);
        }

        foreach (var step in BuildMonotoneAlignment(freeLeft.Count, freeRight.Count,
                     (i, j) => BodyAffinity(leftBodies[i], rightBodies[j], leftSignatures[i], rightSignatures[j])))
        {
            if (step.Kind != MonotoneStepKind.Pair)
                continue;
            int li = freeLeft[step.LeftIndex];
            int rj = freeRight[step.RightIndex];
            leftMatch[li] = rj;
            rightMatch[rj] = li;
        }
        return true;
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
    /// Canonical identity of a cell's block body, deliberately omitting its tcPr shell. This mirrors the
    /// reader's ContentHash framing so a nested table/image/opaque child remains distinguishable, while a pure
    /// width/gridSpan/shading change cannot destroy an otherwise stable cell anchor. Paragraph-local inline/field
    /// carriers are rolled in too: they are transparent in the paragraph's visible-content hash but must prevent
    /// an edited cell from being anchored as structurally equal.
    /// </summary>
    private static IrHash CellBodyHash(IrCell cell)
    {
        var builder = new IrContentHashBuilder();
        builder.AppendStructure(IrContentHashBuilder.StructureCell);
        foreach (var block in cell.Blocks)
        {
            builder.AppendHash(block.ContentHash);
            IrReader.AppendNestedStructuralCarrierHash(block, builder);
        }
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

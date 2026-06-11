#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// M2.1 block-level alignment engine. Aligns two documents' BODY block lists into a typed
/// <see cref="IrBlockAlignment"/> using unique-hash anchoring (histogram-diff style) plus a
/// longest-increasing-subsequence spine, with moves falling out of the anchoring by construction.
/// </summary>
/// <remarks>
/// <para><b>Granularity.</b> Tables, section breaks and opaque blocks align as WHOLE units — keyed on
/// their <c>ContentHash</c>/<c>FormatFingerprint</c> like any other block. Row/cell-level table
/// alignment is M2.2+.</para>
/// <para><b>Settings.</b> <see cref="IrDiffSettings"/> is accepted for surface stability /
/// future-proofing; M2.1 alignment keys purely on the reader-computed hashes and does not consult
/// the settings. (The token diff that M2.2 runs inside <see cref="IrAlignmentKind.Modified"/> gaps
/// is where the settings start to matter.)</para>
/// <para><b>Determinism.</b> All sorts are stable / total-ordered by integer index; no dictionary
/// iteration order is observed (dictionaries are used only for O(1) lookup, never enumerated for
/// output). Two <see cref="Align"/> calls on the same inputs produce sequence-equal entries.</para>
/// </remarks>
internal static class IrBlockAligner
{
    /// <summary>
    /// Align the body block lists of <paramref name="left"/> and <paramref name="right"/>.
    /// </summary>
    public static IrBlockAlignment Align(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        _ = settings; // accepted for future-proofing; M2.1 alignment is hash-only (see remarks).

        var leftBlocks = left.Body.Blocks;
        var rightBlocks = right.Body.Blocks;
        int nLeft = leftBlocks.Count;
        int nRight = rightBlocks.Count;

        // pairedLeft[i] / pairedRight[j] hold the kind once a block is consumed (anchor or gap fill);
        // null means "still free". leftMatch[i] = the right index it paired with (or -1).
        var leftKind = new IrAlignmentKind?[nLeft];
        var rightKind = new IrAlignmentKind?[nRight];
        var leftMatch = new int[nLeft];
        var rightMatch = new int[nRight];
        Array.Fill(leftMatch, -1);
        Array.Fill(rightMatch, -1);

        // --- Anchor pass A: key (ContentHash, FormatFingerprint), unique-each-side → candidate Unchanged.
        // --- Anchor pass B: key ContentHash alone (over A-unpaired), unique-each-side → candidate FormatOnly.
        var candidates = new List<Candidate>();
        CollectAnchors(leftBlocks, rightBlocks, KeyAB, IrAlignmentKind.Unchanged,
            leftMatch, rightMatch, candidates);
        CollectAnchors(leftBlocks, rightBlocks, KeyContentOnly, IrAlignmentKind.FormatOnly,
            leftMatch, rightMatch, candidates);

        // --- Spine: longest increasing subsequence over candidates (sorted by left index) by right
        // index. On-spine candidates keep their anchor kind (Unchanged/FormatOnly); off-spine become Moved.
        candidates.Sort((a, b) => a.LeftIndex.CompareTo(b.LeftIndex));
        var onSpine = LongestIncreasingSubsequence(candidates);

        for (int c = 0; c < candidates.Count; c++)
        {
            var cand = candidates[c];
            if (onSpine.Contains(c))
            {
                leftKind[cand.LeftIndex] = cand.AnchorKind;
                rightKind[cand.RightIndex] = cand.AnchorKind;
            }
            else
            {
                // Off-spine exact/content anchor = relocated. Format equality does not refine the
                // kind in M2.1: a moved+reformatted exact-content block is still plain Moved.
                leftKind[cand.LeftIndex] = IrAlignmentKind.Moved;
                rightKind[cand.RightIndex] = IrAlignmentKind.Moved;
            }
        }

        // --- Gap fill: between consecutive spine pairs (and the head/tail gaps), pair the remaining
        // (non-Moved, non-anchored) left and right blocks. Blocks already consumed as Moved or anchored
        // do NOT participate — they are skipped when walking the gaps.
        //
        // Build the ordered list of spine pairs (left index, right index), both ascending in lockstep.
        var spinePairs = onSpine
            .Select(c => (Left: candidates[c].LeftIndex, Right: candidates[c].RightIndex))
            .OrderBy(p => p.Left)
            .ToList();

        FillGaps(leftBlocks, rightBlocks, spinePairs, leftKind, rightKind, leftMatch, rightMatch);

        // --- Emit in right order with left-anchored deletion interleave.
        var entries = EmitEntries(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch);
        return new IrBlockAlignment(IrNodeList.From(entries));
    }

    // ------------------------------------------------------------------ anchoring

    private readonly record struct Candidate(int LeftIndex, int RightIndex, IrAlignmentKind AnchorKind);

    private static (IrHash, IrHash) KeyAB(IrBlock b) => (b.ContentHash, b.FormatFingerprint);

    // ContentHash-only key, widened to the same tuple shape with a zero second component so the two
    // passes can share CollectAnchors' generic dictionary type without boxing.
    private static (IrHash, IrHash) KeyContentOnly(IrBlock b) => (b.ContentHash, default);

    /// <summary>
    /// Find blocks whose key occurs exactly once on each side (among blocks not already paired),
    /// pairing them up as candidates of <paramref name="anchorKind"/>. For the FormatOnly pass, a
    /// pair whose fingerprints actually match is skipped (it was already an A candidate or would be a
    /// redundant Unchanged) — only genuine same-content/different-format pairs become FormatOnly.
    /// </summary>
    private static void CollectAnchors(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        Func<IrBlock, (IrHash, IrHash)> key, IrAlignmentKind anchorKind,
        int[] leftMatch, int[] rightMatch, List<Candidate> candidates)
    {
        var leftByKey = BuildUniqueIndex(leftBlocks, leftMatch, key);
        var rightByKey = BuildUniqueIndex(rightBlocks, rightMatch, key);

        // Iterate the LEFT blocks in index order (not the dictionary) so output is order-deterministic.
        for (int i = 0; i < leftBlocks.Count; i++)
        {
            if (leftMatch[i] != -1)
                continue;
            var k = key(leftBlocks[i]);
            if (!leftByKey.TryGetValue(k, out int li) || li != i)
                continue; // not the unique left occurrence of this key
            if (!rightByKey.TryGetValue(k, out int rj))
                continue; // no unique right counterpart
            if (rightMatch[rj] != -1)
                continue;

            if (anchorKind == IrAlignmentKind.FormatOnly &&
                leftBlocks[i].FormatFingerprint.Equals(rightBlocks[rj].FormatFingerprint))
                continue; // identical content+format would be Unchanged, not FormatOnly

            leftMatch[i] = rj;
            rightMatch[rj] = i;
            candidates.Add(new Candidate(i, rj, anchorKind));
        }
    }

    /// <summary>
    /// Build key → index for keys occurring exactly ONCE among the still-unpaired blocks; keys with
    /// 0 or ≥2 unpaired occurrences are absent (so non-unique boilerplate never anchors globally).
    /// </summary>
    private static Dictionary<(IrHash, IrHash), int> BuildUniqueIndex(
        IrNodeList<IrBlock> blocks, int[] matched, Func<IrBlock, (IrHash, IrHash)> key)
    {
        var counts = new Dictionary<(IrHash, IrHash), int>();
        var firstIndex = new Dictionary<(IrHash, IrHash), int>();
        for (int i = 0; i < blocks.Count; i++)
        {
            if (matched[i] != -1)
                continue;
            var k = key(blocks[i]);
            counts[k] = counts.TryGetValue(k, out int c) ? c + 1 : 1;
            if (!firstIndex.ContainsKey(k))
                firstIndex[k] = i;
        }

        var unique = new Dictionary<(IrHash, IrHash), int>();
        foreach (var kv in firstIndex)
            if (counts[kv.Key] == 1)
                unique[kv.Key] = kv.Value;
        return unique;
    }

    // ------------------------------------------------------------------ LIS spine

    /// <summary>
    /// Standard O(k log k) longest increasing subsequence by <see cref="Candidate.RightIndex"/> over
    /// <paramref name="candidates"/> (already sorted ascending by left index). Returns the set of
    /// candidate-list indices that lie on one chosen longest increasing subsequence. Ties are broken
    /// deterministically by the patience-sort tail discipline (strictly increasing right index).
    /// </summary>
    private static HashSet<int> LongestIncreasingSubsequence(List<Candidate> candidates)
    {
        int n = candidates.Count;
        var result = new HashSet<int>();
        if (n == 0)
            return result;

        // tails[len-1] = candidate-index whose right value ends an increasing subsequence of length len.
        var tails = new List<int>();
        var prev = new int[n];
        for (int i = 0; i < n; i++)
        {
            prev[i] = -1;
            int right = candidates[i].RightIndex;

            // Binary search for the first tail whose right value is >= this right (strictly increasing).
            int lo = 0, hi = tails.Count;
            while (lo < hi)
            {
                int mid = (lo + hi) >> 1;
                if (candidates[tails[mid]].RightIndex < right)
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

        // Reconstruct from the last tail back through prev.
        for (int i = tails[tails.Count - 1]; i != -1; i = prev[i])
            result.Add(i);
        return result;
    }

    // ------------------------------------------------------------------ gap fill

    /// <summary>
    /// Walk each gap delimited by consecutive spine pairs (plus the head gap before the first and the
    /// tail gap after the last). A gap is the contiguous spans of left indices and right indices that
    /// lie strictly between the two delimiting spine pairs. Within a gap, only blocks STILL FREE
    /// (not anchored, not Moved) participate.
    /// </summary>
    private static void FillGaps(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        List<(int Left, int Right)> spinePairs,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, sl, prevRight + 1, sr,
                leftKind, rightKind, leftMatch, rightMatch);
            prevLeft = sl;
            prevRight = sr;
        }
        // Tail gap (after the last spine pair, or the whole document if there were no spine pairs).
        FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, leftBlocks.Count, prevRight + 1, rightBlocks.Count,
            leftKind, rightKind, leftMatch, rightMatch);
    }

    /// <summary>
    /// Fill one gap: free left indices in [leftFrom, leftTo) and free right indices in
    /// [rightFrom, rightTo). Refinement first (cheap, deterministic, still linear): in-order pair
    /// equal (ContentHash,FormatFingerprint) keys as Unchanged then equal ContentHash as FormatOnly —
    /// this resolves "N identical boilerplate paragraphs, one deleted" to N-1 Unchanged + 1 Deleted
    /// with zero Moved/Modified. Then positional-pair the remaining free blocks as Modified; surplus
    /// left → Deleted, surplus right → Inserted.
    /// </summary>
    private static void FillOneGap(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch)
    {
        var freeLeft = new List<int>();
        for (int i = leftFrom; i < leftTo; i++)
            if (leftMatch[i] == -1)
                freeLeft.Add(i);
        var freeRight = new List<int>();
        for (int j = rightFrom; j < rightTo; j++)
            if (rightMatch[j] == -1)
                freeRight.Add(j);

        // Refinement pass 1: in-order exact (content+format) pairing → Unchanged.
        InOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch,
            requireFormatEqual: true, kind: IrAlignmentKind.Unchanged);
        // Refinement pass 2: in-order content-only pairing → FormatOnly.
        InOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch,
            requireFormatEqual: false, kind: IrAlignmentKind.FormatOnly);

        // Drop the now-consumed entries, preserving order.
        freeLeft.RemoveAll(i => leftMatch[i] != -1);
        freeRight.RemoveAll(j => rightMatch[j] != -1);

        // Positional pairing of the remainder → Modified; surplus → Deleted / Inserted.
        int pairCount = Math.Min(freeLeft.Count, freeRight.Count);
        for (int p = 0; p < pairCount; p++)
        {
            int li = freeLeft[p];
            int rj = freeRight[p];
            leftKind[li] = IrAlignmentKind.Modified;
            rightKind[rj] = IrAlignmentKind.Modified;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
        }
        for (int p = pairCount; p < freeLeft.Count; p++)
            leftKind[freeLeft[p]] = IrAlignmentKind.Deleted;
        for (int p = pairCount; p < freeRight.Count; p++)
            rightKind[freeRight[p]] = IrAlignmentKind.Inserted;
    }

    /// <summary>
    /// In-order first-to-first matching within a gap. For each free right block (in order), pair it
    /// with the FIRST still-free left block (in order) whose key matches under this pass's gate
    /// (content-equal, plus format-equal for Unchanged / format-differ for FormatOnly). This is the
    /// greedy first-to-first matching the plan specifies — it resolves repeated-boilerplate gaps
    /// (identical content+format) into one-to-one Unchanged pairs with the surplus falling out as
    /// Deleted/Inserted, with zero Moved/Modified. It is O(gap²) in the worst case but the dominant
    /// boilerplate case (a single shared key) is effectively linear; gaps are bounded by the spacing
    /// between unique anchors, so this never reintroduces a global O(n²).
    /// </summary>
    private static void InOrderRefine(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        List<int> freeLeft, List<int> freeRight,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        bool requireFormatEqual, IrAlignmentKind kind)
    {
        foreach (int rj in freeRight)
        {
            if (rightMatch[rj] != -1)
                continue;
            foreach (int candLeft in freeLeft)
            {
                if (leftMatch[candLeft] != -1)
                    continue;
                if (!leftBlocks[candLeft].ContentHash.Equals(rightBlocks[rj].ContentHash))
                    continue;
                bool formatEqual = leftBlocks[candLeft].FormatFingerprint.Equals(rightBlocks[rj].FormatFingerprint);
                if (requireFormatEqual != formatEqual)
                    continue; // Unchanged needs format-equal; FormatOnly needs format-differ

                leftKind[candLeft] = kind;
                rightKind[rj] = kind;
                leftMatch[candLeft] = rj;
                rightMatch[rj] = candLeft;
                break;
            }
        }
    }

    // ------------------------------------------------------------------ emit

    /// <summary>
    /// Emit entries in RIGHT-document order, interleaving Deleted (left-only) entries using the
    /// left-anchored unified-diff convention: each deleted left block is emitted right after the entry
    /// of the nearest PAIRED left block preceding it; deletions before any paired left block go first.
    /// </summary>
    private static List<IrAlignedBlock> EmitEntries(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch)
    {
        // Group deleted left indices by the left index of the nearest preceding PAIRED left block.
        // anchorLeftIndex = the left index whose right-side entry a deletion trails; -1 = emit at front.
        var deletionsAfterLeft = new Dictionary<int, List<int>>();
        int lastPairedLeft = -1;
        for (int i = 0; i < leftBlocks.Count; i++)
        {
            if (leftKind[i] == IrAlignmentKind.Deleted)
            {
                if (!deletionsAfterLeft.TryGetValue(lastPairedLeft, out var list))
                    deletionsAfterLeft[lastPairedLeft] = list = new List<int>();
                list.Add(i);
            }
            else if (leftMatch[i] != -1)
            {
                lastPairedLeft = i;
            }
        }

        var entries = new List<IrAlignedBlock>();

        // Front deletions (those preceding every paired left block).
        EmitDeletions(deletionsAfterLeft, -1, leftBlocks, entries);

        for (int j = 0; j < rightBlocks.Count; j++)
        {
            var kind = rightKind[j] ?? IrAlignmentKind.Inserted;
            int li = rightMatch[j];
            IrBlock? leftBlock = li != -1 ? leftBlocks[li] : null;
            entries.Add(new IrAlignedBlock(kind, leftBlock, rightBlocks[j]));

            // After emitting a paired right block, flush deletions anchored to its left partner.
            if (li != -1)
                EmitDeletions(deletionsAfterLeft, li, leftBlocks, entries);
        }

        return entries;
    }

    private static void EmitDeletions(
        Dictionary<int, List<int>> deletionsAfterLeft, int anchorLeftIndex,
        IrNodeList<IrBlock> leftBlocks, List<IrAlignedBlock> entries)
    {
        if (!deletionsAfterLeft.TryGetValue(anchorLeftIndex, out var list))
            return;
        foreach (int li in list) // already in ascending left order
            entries.Add(new IrAlignedBlock(IrAlignmentKind.Deleted, leftBlocks[li], null));
    }
}

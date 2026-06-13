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
        => AlignBlocks(left.Body.Blocks, right.Body.Blocks, settings);

    /// <summary>
    /// Align two raw block lists (M2.2 Task 4 generalization). The public <see cref="Align"/> calls this
    /// with the bodies; <see cref="IrTableDiffer"/> calls it on a table CELL's block list to recurse the
    /// same machinery into cell contents. Identical semantics — anchoring, LIS spine, gap fill, fuzzy
    /// moves — just over an arbitrary block list rather than a document body.
    /// </summary>
    public static IrBlockAlignment AlignBlocks(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks, IrDiffSettings settings)
    {
        // M2.2 Task 3: settings now drive similarity-based in-gap pairing + cross-gap fuzzy moves.
        // One per-call similarity scorer carries the tokenization cache (each block tokenized at most
        // once across all the candidate-pair scorings below).
        var similarity = new IrBlockSimilarity(settings);

        int nLeft = leftBlocks.Count;
        int nRight = rightBlocks.Count;

        // leftKind[i] / rightKind[j] hold the kind once a block is consumed (anchor or gap fill);
        // null means "still free". leftMatch[i] = the right index it paired with (or -1).
        var leftKind = new IrAlignmentKind?[nLeft];
        var rightKind = new IrAlignmentKind?[nRight];
        var leftMatch = new int[nLeft];
        var rightMatch = new int[nRight];
        Array.Fill(leftMatch, -1);
        Array.Fill(rightMatch, -1);

        // --- Anchor pass A: key (ContentHash, FormatFingerprint), unique-each-side → candidate Unchanged.
        // --- Anchor pass B: key ContentHash alone (over A-unpaired), unique-each-side → Unchanged or
        //     FormatOnly decided by FormatEqual (boundary-normalized modeled-only under ModeledOnly; the
        //     stored fingerprint under Full). M2.2 Task 4: this is where unmodeled-rPr noise that flipped
        //     the stored fingerprint (lang/bCs/iCs/…) is reclassified Unchanged instead of FormatOnly.
        var candidates = new List<Candidate>();
        CollectAnchors(leftBlocks, rightBlocks, KeyAB, IrAlignmentKind.Unchanged,
            leftMatch, rightMatch, candidates, settings);
        CollectAnchors(leftBlocks, rightBlocks, KeyContentOnly, IrAlignmentKind.FormatOnly,
            leftMatch, rightMatch, candidates, settings);

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

        FillGaps(leftBlocks, rightBlocks, spinePairs, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);

        // --- Cross-gap fuzzy moves: over the GLOBAL leftover Deleted × Inserted sets (after all gap
        // fill), re-pair similar blocks as Moved / MovedModified. Runs AFTER gap fill so it sees the
        // final Deleted/Inserted leftovers, never blocks already consumed in-place.
        DetectCrossGapMoves(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);

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
    /// pairing them up. Pass A (<paramref name="anchorKind"/> = Unchanged, key includes the fingerprint)
    /// only ever pairs exact content+format matches. Pass B (<paramref name="anchorKind"/> = FormatOnly,
    /// ContentHash-only key) pairs content-equal blocks and then DECIDES the kind via
    /// <see cref="FormatEqual"/>: format-equal (boundary-normalized modeled-only under ModeledOnly, the
    /// stored fingerprint under Full) → Unchanged, else FormatOnly. This is what makes unmodeled-rPr
    /// noise that flips the stored fingerprint reclassify as Unchanged under the default policy.
    /// </summary>
    private static void CollectAnchors(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        Func<IrBlock, (IrHash, IrHash)> key, IrAlignmentKind anchorKind,
        int[] leftMatch, int[] rightMatch, List<Candidate> candidates, IrDiffSettings settings)
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

            // Pass B refines its content-equal pair into Unchanged (format-equal) or FormatOnly.
            var resolvedKind = anchorKind == IrAlignmentKind.FormatOnly
                ? (FormatEqual(leftBlocks[i], rightBlocks[rj], settings)
                    ? IrAlignmentKind.Unchanged : IrAlignmentKind.FormatOnly)
                : anchorKind;

            leftMatch[i] = rj;
            rightMatch[rj] = i;
            candidates.Add(new Candidate(i, rj, resolvedKind));
        }
    }

    /// <summary>
    /// Diff-time format equality of two content-equal blocks under the settings' format-comparison
    /// policy. Under <see cref="IrFormatComparison.Full"/> (and for any non-paragraph pair) it is the
    /// stored block <c>FormatFingerprint</c>. Under <see cref="IrFormatComparison.ModeledOnly"/> for a
    /// paragraph pair it is the BOUNDARY-NORMALIZED modeled-only block signature — the per-token
    /// (MatchKey, modeled-format) sequence — which is invariant to the run-resegmentation churn that
    /// flips the stored fingerprint (the M2.1 finding), so unmodeled rPr noise no longer reads as a
    /// format change.
    /// </summary>
    private static bool FormatEqual(IrBlock left, IrBlock right, IrDiffSettings settings)
    {
        if (settings.FormatComparison == IrFormatComparison.ModeledOnly
            && left is IrParagraph lp && right is IrParagraph rp)
            return IrModeledFormat.BlockSignature(lp, settings) == IrModeledFormat.BlockSignature(rp, settings);

        return left.FormatFingerprint.Equals(right.FormatFingerprint);
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
        int[] leftMatch, int[] rightMatch,
        IrBlockSimilarity similarity, IrDiffSettings settings)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, sl, prevRight + 1, sr,
                leftKind, rightKind, leftMatch, rightMatch, similarity, settings);
            prevLeft = sl;
            prevRight = sr;
        }
        // Tail gap (after the last spine pair, or the whole document if there were no spine pairs).
        FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, leftBlocks.Count, prevRight + 1, rightBlocks.Count,
            leftKind, rightKind, leftMatch, rightMatch, similarity, settings);
    }

    /// <summary>
    /// Fill one gap: free left indices in [leftFrom, leftTo) and free right indices in
    /// [rightFrom, rightTo). Refinement first (cheap, deterministic, still linear): in-order pair
    /// equal (ContentHash,FormatFingerprint) keys as Unchanged then equal ContentHash as FormatOnly —
    /// this resolves "N identical boilerplate paragraphs, one deleted" to N-1 Unchanged + 1 Deleted
    /// with zero Moved/Modified. Then SIMILARITY-pair the remaining free blocks as Modified (M2.2
    /// Task 3, replacing the M2.1 blind positional pairing); surplus left → Deleted, surplus right →
    /// Inserted. The similarity pairing is what lets a cross-positioned in-gap edit land as Modified
    /// instead of falling out as Delete+Insert when the gap's free blocks are not aligned positionally.
    /// </summary>
    private static void FillOneGap(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        IrBlockSimilarity similarity, IrDiffSettings settings)
    {
        var freeLeft = new List<int>();
        for (int i = leftFrom; i < leftTo; i++)
            if (leftMatch[i] == -1)
                freeLeft.Add(i);
        var freeRight = new List<int>();
        for (int j = rightFrom; j < rightTo; j++)
            if (rightMatch[j] == -1)
                freeRight.Add(j);

        // Refinement pass 1: in-order content-equal + format-equal pairing → Unchanged.
        InOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch,
            requireFormatEqual: true, kind: IrAlignmentKind.Unchanged, settings: settings);
        // Refinement pass 2: in-order content-equal + format-DIFFERING pairing → FormatOnly.
        InOrderRefine(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind, leftMatch, rightMatch,
            requireFormatEqual: false, kind: IrAlignmentKind.FormatOnly, settings: settings);

        // Drop the now-consumed entries, preserving order.
        freeLeft.RemoveAll(i => leftMatch[i] != -1);
        freeRight.RemoveAll(j => rightMatch[j] != -1);

        // Similarity pairing of the remainder → Modified; leftovers → Deleted / Inserted.
        //
        // Greedy best-score: repeatedly take the highest-scoring free left×right pair whose score is
        // ≥ BlockSimilarityThreshold; consume both; repeat. Ties break by smallest left index, then
        // smallest right index (so the choice is a deterministic function of the gap's block order).
        // Cost: each round scans the (≤ |freeLeft|·|freeRight|) candidate grid once; with at most
        // min(|freeLeft|,|freeRight|) rounds, that is gap-bounded G²·(tokenization) — the same G²/2-class
        // bound the in-order refinement documents — and the per-call tokenization cache means every block
        // in the gap is tokenized at most once regardless of how many candidate pairs reference it.
        SimilarityPair(leftBlocks, rightBlocks, freeLeft, freeRight, leftKind, rightKind,
            leftMatch, rightMatch, similarity, settings.BlockSimilarityThreshold);

        // Collect what the similarity pass left unpaired (still in ascending index order).
        var leftoverLeft = new List<int>();
        foreach (int li in freeLeft)
            if (leftMatch[li] == -1)
                leftoverLeft.Add(li);
        var leftoverRight = new List<int>();
        foreach (int rj in freeRight)
            if (rightMatch[rj] == -1)
                leftoverRight.Add(rj);

        // Unambiguous table residue → Modified regardless of score (M2.4b Workstream C). A table can only
        // sensibly pair with a table — a table-vs-paragraph similarity is 0 — so when exactly ONE free-left
        // table and ONE free-right table survive the threshold in this gap, they are the same table edited,
        // even when their cell-content Jaccard is below the generic BlockSimilarityThreshold (a heavily-edited
        // table is still ONE edited table, not a delete+insert). Pairing them as Modified feeds IrTableDiffer's
        // row/cell diff, matching WmlComparer's per-cell endnote-table revisions (WC-1750/1760). This is the
        // table analogue of the 1×1 residue below; it fires only when the table pairing is UNAMBIGUOUS (one on
        // each side), so it never competes with a better-scoring candidate (those were taken by SimilarityPair).
        var tableLeft = new List<int>();
        foreach (int li in leftoverLeft)
            if (leftBlocks[li] is IrTable)
                tableLeft.Add(li);
        var tableRight = new List<int>();
        foreach (int rj in leftoverRight)
            if (rightBlocks[rj] is IrTable)
                tableRight.Add(rj);
        if (tableLeft.Count == 1 && tableRight.Count == 1)
        {
            int li = tableLeft[0];
            int rj = tableRight[0];
            leftKind[li] = IrAlignmentKind.Modified;
            rightKind[rj] = IrAlignmentKind.Modified;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
            leftoverLeft.Remove(li);
            leftoverRight.Remove(rj);
        }

        // Unambiguous 1×1 residue → Modified regardless of score. When exactly ONE free left and ONE free
        // right survive the threshold, there is no competing candidate to disambiguate: classifying the
        // lone pair as "the same block, edited" is the only sensible reading (and is what M2.1's positional
        // pairing did for an isolated edit). The BlockSimilarityThreshold exists to choose AMONG candidates
        // and to reject leftovers when there is a surplus on one side — not to demote a solitary in-place
        // edit (e.g. "beta" → "BETA-edited") to Delete+Insert. A genuine cross-gap relocation never reaches
        // here as a 1×1 gap residue (it occupies DIFFERENT gaps, handled by DetectCrossGapMoves), so this
        // does not manufacture false in-place edits out of moves.
        if (leftoverLeft.Count == 1 && leftoverRight.Count == 1)
        {
            int li = leftoverLeft[0];
            int rj = leftoverRight[0];
            leftKind[li] = IrAlignmentKind.Modified;
            rightKind[rj] = IrAlignmentKind.Modified;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
            return;
        }

        // Otherwise the leftovers fall out as Deleted / Inserted (a surplus on one side, or a multi-block
        // residue where below-threshold pairs are deliberately split rather than positionally guessed).
        foreach (int li in leftoverLeft)
            leftKind[li] = IrAlignmentKind.Deleted;
        foreach (int rj in leftoverRight)
            rightKind[rj] = IrAlignmentKind.Inserted;
    }

    /// <summary>
    /// Greedy best-score one-to-one pairing of <paramref name="freeLeft"/> × <paramref name="freeRight"/>
    /// as <see cref="IrAlignmentKind.Modified"/>: repeatedly pick the highest-scoring still-free pair with
    /// score ≥ <paramref name="threshold"/> (ties: smallest left index, then smallest right index),
    /// consume both, repeat until no qualifying pair remains. Leaves the unpaired blocks for the caller to
    /// classify Deleted/Inserted. Deterministic: the pick is a pure function of the score grid + index
    /// tie-break, and scoring is cached so the grid is cheap to rescan each round.
    /// </summary>
    private static void SimilarityPair(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        List<int> freeLeft, List<int> freeRight,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        IrBlockSimilarity similarity, double threshold)
    {
        while (true)
        {
            double bestScore = threshold;
            int bestLeft = -1, bestRight = -1;
            bool found = false;
            foreach (int li in freeLeft)
            {
                if (leftMatch[li] != -1)
                    continue;
                foreach (int rj in freeRight)
                {
                    if (rightMatch[rj] != -1)
                        continue;
                    double score = similarity.Score(leftBlocks[li], rightBlocks[rj]);
                    // Strictly-greater wins; on a tie keep the first seen (freeLeft / freeRight are in
                    // ascending index order), which is exactly "smallest left, then smallest right".
                    if (score > bestScore || (!found && score >= threshold))
                    {
                        bestScore = score;
                        bestLeft = li;
                        bestRight = rj;
                        found = true;
                    }
                }
            }

            if (!found)
                return;

            leftKind[bestLeft] = IrAlignmentKind.Modified;
            rightKind[bestRight] = IrAlignmentKind.Modified;
            leftMatch[bestLeft] = bestRight;
            rightMatch[bestRight] = bestLeft;
        }
    }

    /// <summary>
    /// In-order first-to-first matching within a gap. For each free right block (in order), pair it
    /// with the FIRST still-free left block (in order) whose key matches under this pass's gate
    /// (content-equal, plus format-equal for Unchanged / format-differ for FormatOnly). This is the
    /// greedy first-to-first matching the plan specifies — it resolves repeated-boilerplate gaps
    /// (identical content+format) into one-to-one Unchanged pairs with the surplus falling out as
    /// Deleted/Inserted, with zero Moved/Modified. It is O(gap²) in the worst case — a single
    /// all-distinct-content gap of size G costs ~G²/2 comparisons, i.e. ~2M at G≈2000 (sub-ms) —
    /// but the dominant boilerplate case (a single shared key) is effectively linear; gaps are
    /// bounded by the spacing between unique anchors, so this never reintroduces a global O(n²).
    /// Scale-guard fixtures (Task 3) should size inputs against that G²/2 bound deliberately.
    /// </summary>
    private static void InOrderRefine(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        List<int> freeLeft, List<int> freeRight,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        bool requireFormatEqual, IrAlignmentKind kind, IrDiffSettings settings)
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
                bool formatEqual = FormatEqual(leftBlocks[candLeft], rightBlocks[rj], settings);
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

    // ------------------------------------------------------------------ cross-gap fuzzy moves

    /// <summary>
    /// M2.2 Task 3 cross-gap fuzzy move detection. After ALL gap fill, the only remaining Deleted (left)
    /// and Inserted (right) blocks are content that found no in-place counterpart. A relocated-and-edited
    /// block lands here as a Deleted at its old position + an Inserted at its new position. We re-pair such
    /// blocks: among the global leftover Deleted × Inserted candidates, a pair with ≥
    /// <see cref="IrDiffSettings.MoveMinimumTokenCount"/> Word tokens on BOTH sides and similarity ≥
    /// <see cref="IrDiffSettings.MoveSimilarityThreshold"/> becomes a move.
    /// </summary>
    /// <remarks>
    /// <para><b>Greedy + deterministic.</b> Same discipline as in-gap pairing: repeatedly take the
    /// highest-scoring qualifying pair (ties: smallest left index, then smallest right index), consume
    /// both, repeat. Each block is consumed at most once.</para>
    /// <para><b>Move vs MovedModified.</b> A qualifying pair is normally <see cref="IrAlignmentKind.MovedModified"/>
    /// — the edit script re-token-diffs it (move + nested edits, the capability WmlComparer cannot
    /// express). A score of exactly 1.0 means the token multisets are identical; if additionally the
    /// ContentHashes are equal the blocks are exact-content relocations, which must classify as plain
    /// <see cref="IrAlignmentKind.Moved"/> (a MovedModified with an all-Equal token diff would be a lie
    /// about there being an edit). In practice exact-content moves are already caught by off-spine
    /// anchoring and never reach here, but the guard makes the classification correct regardless.</para>
    /// <para><b>Cost.</b> Bounded by the global leftover counts D × I, not the document size: the
    /// dominant boilerplate / clean-edit cases leave few leftovers. Tokenization is cached (shared with
    /// in-gap pairing), so each leftover block is tokenized at most once across both passes.</para>
    /// </remarks>
    private static void DetectCrossGapMoves(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        IrBlockSimilarity similarity, IrDiffSettings settings)
    {
        // Collect global leftovers in ascending index order (drives the deterministic tie-break).
        var deleted = new List<int>();
        for (int i = 0; i < leftBlocks.Count; i++)
            if (leftKind[i] == IrAlignmentKind.Deleted)
                deleted.Add(i);
        var inserted = new List<int>();
        for (int j = 0; j < rightBlocks.Count; j++)
            if (rightKind[j] == IrAlignmentKind.Inserted)
                inserted.Add(j);

        if (deleted.Count == 0 || inserted.Count == 0)
            return;

        double threshold = settings.MoveSimilarityThreshold;
        int minTokens = settings.MoveMinimumTokenCount;

        while (true)
        {
            double bestScore = threshold;
            int bestLeft = -1, bestRight = -1;
            bool found = false;
            foreach (int li in deleted)
            {
                if (leftMatch[li] != -1)
                    continue;
                if (similarity.WordCount(leftBlocks[li]) < minTokens)
                    continue; // too short to be a reliable move (mirrors MoveMinimumWordCount)
                foreach (int rj in inserted)
                {
                    if (rightMatch[rj] != -1)
                        continue;
                    if (similarity.WordCount(rightBlocks[rj]) < minTokens)
                        continue;
                    double score = similarity.Score(leftBlocks[li], rightBlocks[rj]);
                    if (score > bestScore || (!found && score >= threshold))
                    {
                        bestScore = score;
                        bestLeft = li;
                        bestRight = rj;
                        found = true;
                    }
                }
            }

            if (!found)
                return;

            // Exact-content relocation (score 1.0 + equal ContentHash) is plain Moved, not MovedModified:
            // there is genuinely no edit to re-diff. Everything else is MovedModified (the edit script
            // re-token-diffs it).
            bool exact = bestScore >= 1.0 &&
                leftBlocks[bestLeft].ContentHash.Equals(rightBlocks[bestRight].ContentHash);
            var kind = exact ? IrAlignmentKind.Moved : IrAlignmentKind.MovedModified;

            leftKind[bestLeft] = kind;
            rightKind[bestRight] = kind;
            leftMatch[bestLeft] = bestRight;
            rightMatch[bestRight] = bestLeft;
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

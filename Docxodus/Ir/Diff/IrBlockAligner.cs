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

        // M2.6: the gap fill records every fired split (one left → N right) and merge (N left → one
        // right) group here; EmitEntries consumes them to emit the single Split/Merge entry per group.
        var splitGroups = new List<(int SingularIndex, List<int> PluralIndexes)>();
        var mergeGroups = new List<(int SingularIndex, List<int> PluralIndexes)>();

        FillGaps(leftBlocks, rightBlocks, spinePairs, leftKind, rightKind, leftMatch, rightMatch,
            similarity, settings, splitGroups, mergeGroups);

        // --- Cross-gap fuzzy moves: over the GLOBAL leftover Deleted × Inserted sets (after all gap
        // fill), re-pair similar blocks as Moved / MovedModified. Runs AFTER gap fill so it sees the
        // final Deleted/Inserted leftovers, never blocks already consumed in-place.
        DetectCrossGapMoves(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);

        // --- Emit in right order with left-anchored deletion interleave.
        var entries = EmitEntries(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch,
            splitGroups, mergeGroups);
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
        IrBlockSimilarity similarity, IrDiffSettings settings,
        List<(int SingularIndex, List<int> PluralIndexes)> splitGroups,
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, sl, prevRight + 1, sr,
                leftKind, rightKind, leftMatch, rightMatch, similarity, settings, splitGroups, mergeGroups);
            prevLeft = sl;
            prevRight = sr;
        }
        // Tail gap (after the last spine pair, or the whole document if there were no spine pairs).
        FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, leftBlocks.Count, prevRight + 1, rightBlocks.Count,
            leftKind, rightKind, leftMatch, rightMatch, similarity, settings, splitGroups, mergeGroups);
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
        IrBlockSimilarity similarity, IrDiffSettings settings,
        List<(int SingularIndex, List<int> PluralIndexes)> splitGroups,
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups)
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

        // M2.6 1:N split / N:1 merge containment scan (gated; default OFF during the build-out).
        //
        // PLACEMENT RATIONALE. The scan runs AFTER SimilarityPair (and the table residue) so that a
        // better 1:1 pairing always wins first — a clean Modified pair is never torn into a
        // speculative split; the scan only PROMOTES an existing this-gap Modified pairing when the
        // run-containment evidence says the partner is one segment of a multi-paragraph split. It
        // runs BEFORE the 1×1-residue rule and the surplus classification because without it a 1:N
        // split's members fall through to surplus Inserted/Deleted — exactly the WC-1450/WC-1830
        // corpus deviation this pass exists to fix — and the residue must still be re-classifiable
        // when the scan sees it.
        //
        // The split scan runs BEFORE the merge scan; every block a split group consumes gets its
        // match slot stamped, so the merge scan can never reuse it (F2.2 overlap ceiling — no block
        // is ever a member of two groups).
        if (settings.DetectSplitMerge)
        {
            DetectOneToManyInGap(leftBlocks, rightBlocks, leftFrom, leftTo, rightFrom, rightTo,
                leftKind, rightKind, leftMatch, rightMatch, leftoverLeft, leftoverRight,
                IrAlignmentKind.Split, splitGroups, settings);
            DetectOneToManyInGap(rightBlocks, leftBlocks, rightFrom, rightTo, leftFrom, leftTo,
                rightKind, leftKind, rightMatch, leftMatch, leftoverRight, leftoverLeft,
                IrAlignmentKind.Merge, mergeGroups, settings);
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

    // ------------------------------------------------------------------ split/merge detection (M2.6)

    /// <summary>
    /// One-directional 1:N containment scan over a gap, side-parameterized so the SAME worker serves
    /// both directions: for a SPLIT, singular = left / plural = right (one left paragraph whose
    /// content migrated across N adjacent right paragraphs); for a MERGE the call mirrors the sides
    /// (singular = right / plural = left) and stamps <see cref="IrAlignmentKind.Merge"/>.
    /// </summary>
    /// <remarks>
    /// <para><b>Candidates (F4.2).</b> A singular-side gap block qualifies only if it is an
    /// <see cref="IrParagraph"/> that is either still FREE or was Modified-paired BY THIS GAP'S
    /// SimilarityPair to a plural-side paragraph inside this gap. Unchanged/FormatOnly/Moved blocks
    /// are NEVER candidates: an identity-reserved (WC022) or content-anchored pair is
    /// ContentHash-equal, so its singular side has ZERO unmatched tail — promoting it could only
    /// manufacture a false split out of a genuinely-new neighbor (review finding F4.2; regression
    /// test <c>Detection_never_promotes_an_identity_reserved_unchanged_pair</c>).</para>
    /// <para><b>Consumption (F2.2).</b> Scan order is ascending singular index; the first qualifying
    /// window per candidate wins; fired members get their match slots stamped immediately, and a
    /// candidate window never admits an already-consumed (non-free, non-partner) index by
    /// construction — so no block can belong to two groups.</para>
    /// <para><b>Determinism.</b> Pure index-ascending scans; no dictionary enumeration feeds output.</para>
    /// </remarks>
    private static void DetectOneToManyInGap(
        IrNodeList<IrBlock> singularBlocks, IrNodeList<IrBlock> pluralBlocks,
        int singularFrom, int singularTo, int pluralFrom, int pluralTo,
        IrAlignmentKind?[] singularKind, IrAlignmentKind?[] pluralKind,
        int[] singularMatch, int[] pluralMatch,
        List<int> leftoverSingular, List<int> leftoverPlural,
        IrAlignmentKind kind,
        List<(int SingularIndex, List<int> PluralIndexes)> groups,
        IrDiffSettings settings)
    {
        // O(1)-prefilter content-token counts, computed lazily once per gap (-1 = not yet counted).
        // The thresholds imply HARD length bounds a window must satisfy before any LCS scoring is
        // worth running: matched ≤ min(singularContent, windowContent), so coverage ≥ T needs
        // windowContent ≥ T·singularContent, and slack ≤ S needs windowContent ≤ singularContent/(1−S)
        // (unmatched ≥ windowContent − singularContent). Without this, a fully-rewritten G×G gap pays
        // G²·LCS for windows that cannot possibly qualify (the adversarial 200×200 scale fixture).
        var pluralContent = new int[pluralTo - pluralFrom];
        Array.Fill(pluralContent, -1);
        int PluralContent(int pj)
        {
            int idx = pj - pluralFrom;
            if (pluralContent[idx] < 0)
                pluralContent[idx] = pluralBlocks[pj] is IrParagraph p ? ContentTokenCount(p, settings) : 0;
            return pluralContent[idx];
        }

        for (int si = singularFrom; si < singularTo; si++)
        {
            if (singularBlocks[si] is not IrParagraph singularPara)
                continue;

            int partner = -1;
            if (singularMatch[si] != -1)
            {
                // F4.2: only a this-gap Modified pairing may be promoted (see remarks).
                if (singularKind[si] != IrAlignmentKind.Modified)
                    continue;
                partner = singularMatch[si];
                if (partner < pluralFrom || partner >= pluralTo)
                    continue;
            }

            var run = FindQualifyingRun(singularPara, partner, pluralBlocks, pluralFrom, pluralTo,
                pluralMatch, settings, PluralContent);
            if (run is null)
                continue;

            // The gate guarantees a paired candidate's partner is inside the run, so the partner's
            // prior Modified stamp is overwritten as a member in the loop below — no dissolve step.
            singularKind[si] = kind;
            singularMatch[si] = run[0];
            foreach (int pj in run)
            {
                pluralKind[pj] = kind;
                pluralMatch[pj] = si;
            }

            groups.Add((si, run));

            // Remove the consumed indices so the 1×1-residue rule and the surplus classification
            // only see what genuinely remains in the gap.
            leftoverSingular.Remove(si);
            foreach (int pj in run)
                leftoverPlural.Remove(pj);
        }
    }

    /// <summary>
    /// Enumerate candidate windows of ADJACENT eligible plural-side indices (free paragraphs, or the
    /// candidate's own Modified partner) and return the first — smallest (start, end), both scanned
    /// ascending — that passes ALL gates after edge trimming, or null. Window length is capped at
    /// <see cref="IrDiffSettings.SplitMaxRunLength"/>. Shortest-qualifying-first is deliberate:
    /// the smallest window that already clears the coverage bar absorbs the least foreign content,
    /// so the group claims no more blocks than the evidence supports (a longer window can only add
    /// slack, never coverage the smaller one lacked at the same start).
    /// </summary>
    private static List<int>? FindQualifyingRun(
        IrParagraph singular, int partner,
        IrNodeList<IrBlock> pluralBlocks, int pluralFrom, int pluralTo,
        int[] pluralMatch, IrDiffSettings settings, Func<int, int> pluralContent)
    {
        bool Eligible(int pj) =>
            pluralBlocks[pj] is IrParagraph && (pluralMatch[pj] == -1 || pj == partner);

        // O(1) length prefilter bounds (see DetectOneToManyInGap): a window whose content-token total
        // falls outside [coverage·singular, singular/(1−slack)] cannot clear the thresholds, so the
        // LCS scorer never runs on it. The lower bound uses the UNTRIMMED window (trimming only
        // removes zero-match members, which cannot raise coverage); the upper bound is checked after
        // a hypothetical best-case trim is unknowable cheaply, so it is applied to the raw window —
        // a window that only passes POST-trim is re-admitted because the trimmed window is itself
        // enumerated as a smaller (a,b) candidate by the ascending scan.
        int singularContent = ContentTokenCount(singular, settings);
        double maxWindowContent = singularContent / (1.0 - settings.SplitForeignSlack);
        double minWindowContent = settings.SplitCoverageThreshold * singularContent;

        for (int a = pluralFrom; a < pluralTo; a++)
        {
            if (!Eligible(a))
                continue;
            int windowContent = pluralContent(a);
            for (int b = a + 1; b < pluralTo && b - a + 1 <= settings.SplitMaxRunLength; b++)
            {
                if (!Eligible(b))
                    break; // adjacency requirement: the window must be a contiguous eligible run
                windowContent += pluralContent(b);
                if (windowContent > maxWindowContent)
                    break; // adding members only grows content — no longer window from this start qualifies
                if (windowContent < minWindowContent)
                    continue; // too little content to cover the singular side yet — extend the window
                var trimmed = TrimAndGate(singular, partner, a, b, pluralBlocks, pluralMatch, settings);
                if (trimmed is not null)
                    return trimmed;
            }
        }

        return null;
    }

    /// <summary>Content-token count of a paragraph (non-Separator, non-Textbox — the scoring rule).</summary>
    private static int ContentTokenCount(IrParagraph p, IrDiffSettings settings)
    {
        int n = 0;
        foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
            if (t.Kind is not (IrDiffTokenKind.Separator or IrDiffTokenKind.Textbox))
                n++;
        return n;
    }

    /// <summary>
    /// Score one window, apply the R2 edge trim, and check the firing gates. Returns the trimmed
    /// member index list when the window qualifies, else null.
    /// </summary>
    /// <remarks>
    /// <b>R2 edge trim (false-positive guard).</b> Leading and trailing members with ZERO
    /// LCS-matched content tokens are dropped before gating: an unrelated edge insert (net-new
    /// neighbor paragraph) and an edge empty carrier (an empty paragraph has no content tokens, so
    /// it can never match) must not ride along in a split group just because they are adjacent.
    /// INTERIOR zero-match members — WC-1830's net-new math paragraph between the two halves — are
    /// deliberately KEPT (absorbed): they sit between matched segments, so excluding them would
    /// break the run's adjacency, and their foreign content is already priced by the slack gate.
    /// </remarks>
    private static List<int>? TrimAndGate(
        IrParagraph singular, int partner, int a, int b,
        IrNodeList<IrBlock> pluralBlocks, int[] pluralMatch, IrDiffSettings settings)
    {
        var window = new List<int>(b - a + 1);
        for (int pj = a; pj <= b; pj++)
            window.Add(pj);
        var paras = window.Select(pj => (IrParagraph)pluralBlocks[pj]).ToList();
        var score = IrSplitSegmenter.Score(singular, paras, settings);

        // R2 edge trim (see remarks).
        int lo = 0, hi = window.Count - 1;
        while (lo <= hi && score.MemberMatchedContent[lo] == 0)
            lo++;
        while (hi >= lo && score.MemberMatchedContent[hi] == 0)
            hi--;
        if (hi - lo + 1 < 2)
            return null;
        if (lo != 0 || hi != window.Count - 1)
        {
            window = window.GetRange(lo, hi - lo + 1);
            paras = paras.GetRange(lo, hi - lo + 1);
            score = IrSplitSegmenter.Score(singular, paras, settings);
        }

        // Gate 1: ≥2 members carrying at least one content token each. Interior empties are absorbed
        // but do not count toward N — a "split" whose other member is an empty carrier is not a split.
        int contentMembers = paras.Count(p => HasContentTokens(p, settings));
        if (contentMembers < 2)
            return null;

        // Gate 2 (paired candidate only): the partner must survive the trim, and at least one OTHER
        // member must be free — otherwise nothing new is being claimed beyond the existing pairing.
        if (partner != -1)
        {
            // The window is a contiguous index range — an O(1) bounds check, not a List.Contains scan.
            if (partner < window[0] || partner > window[window.Count - 1])
                return null;
            bool anyFree = false;
            foreach (int pj in window)
                if (pj != partner && pluralMatch[pj] == -1)
                    anyFree = true;
            if (!anyFree)
                return null;
        }

        // Gates 3+4: containment thresholds on the trimmed run.
        if (score.Coverage < settings.SplitCoverageThreshold)
            return null;
        if (score.ForeignSlack > settings.SplitForeignSlack)
            return null;

        return window;
    }

    /// <summary>True iff the paragraph tokenizes to at least one content (non-Separator, non-Textbox)
    /// token — the same content rule <see cref="IrSplitSegmenter"/> scores by.</summary>
    private static bool HasContentTokens(IrParagraph p, IrDiffSettings settings)
    {
        foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
            if (t.Kind is not (IrDiffTokenKind.Separator or IrDiffTokenKind.Textbox))
                return true;
        return false;
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
        // Phase 1 — SAME-UNID identity reservation (M2.6 Task 2). Before any first-fit, pair every free right
        // block with a free left block that shares BOTH its key (ContentHash + this pass's format gate) AND its
        // persisted unid. The unid is the IR's stable per-element identity (an unchanged paragraph keeps it
        // across the two documents), so an identity-keyed pair is the genuinely-unchanged correspondence. Doing
        // this FIRST stops the plain first-fit below from stealing an identity-matched left for a DIFFERENT-unid
        // right that happens to be scanned earlier — the WC022 crossing: two adjacent empty paragraphs where a
        // bare empty (kept identity) was consumed by an earlier different-identity empty, forcing the leftover
        // to cross document order and reconstruct swapped on reject. Reserving identities keeps the pairing
        // monotonic. Pure deterministic tie-break: it only changes WHICH equal-key left fills an equal-key
        // right (same kind, same accept/reject content), never which blocks pair overall.
        foreach (int rj in freeRight)
        {
            if (rightMatch[rj] != -1)
                continue;
            foreach (int candLeft in freeLeft)
            {
                if (leftMatch[candLeft] != -1)
                    continue;
                if (!string.Equals(leftBlocks[candLeft].Anchor.Unid, rightBlocks[rj].Anchor.Unid,
                        StringComparison.Ordinal))
                    continue;
                if (!leftBlocks[candLeft].ContentHash.Equals(rightBlocks[rj].ContentHash))
                    continue;
                if (requireFormatEqual != FormatEqual(leftBlocks[candLeft], rightBlocks[rj], settings))
                    continue;

                leftKind[candLeft] = kind;
                rightKind[rj] = kind;
                leftMatch[candLeft] = rj;
                rightMatch[rj] = candLeft;
                break;
            }
        }

        // Phase 2 — first-to-first in document order over whatever identity reservation left free.
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
    /// M2.6: a split group emits ONE <see cref="IrAlignmentKind.Split"/> entry at its FIRST member's
    /// right position (the other members emit nothing); a merge group emits its
    /// <see cref="IrAlignmentKind.Merge"/> entry at the singular right block's position.
    /// </summary>
    private static List<IrAlignedBlock> EmitEntries(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        List<(int SingularIndex, List<int> PluralIndexes)> splitGroups,
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups)
    {
        // O(1) lookups for the right-walk below (lookup only — never enumerated, so determinism
        // rests purely on the index-ascending walk).
        // Split group: singular = the one LEFT index, plural = the N right member indexes.
        // Merge group: singular = the one RIGHT index, plural = the N left member indexes.
        var splitByFirstMember = new Dictionary<int, (int SingularIndex, List<int> PluralIndexes)>();
        foreach (var g in splitGroups)
            splitByFirstMember[g.PluralIndexes[0]] = g;
        var mergeByRight = new Dictionary<int, (int SingularIndex, List<int> PluralIndexes)>();
        foreach (var g in mergeGroups)
            mergeByRight[g.SingularIndex] = g;

        // Group deleted left indices by the left index of the nearest preceding PAIRED left block.
        // anchorLeftIndex = the left index whose right-side entry a deletion trails; -1 = emit at front.
        // Split/merge participants have leftMatch set (the split singular; every merge member), so they
        // correctly act as lastPairedLeft anchors here; their buckets are flushed by the explicit
        // EmitDeletions calls in the walk below — each bucket exactly once.
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
            else if (leftMatch[i] != -1 &&
                     leftKind[i] is not (IrAlignmentKind.Moved or IrAlignmentKind.MovedModified))
            {
                // Only an IN-PLACE paired left anchors trailing deletions. A MOVED left's entry is
                // emitted at its destination's RIGHT position — anchoring a deletion to it would make
                // the deleted block restore at the move DESTINATION on reject instead of in its left
                // neighborhood (a reject-order corruption surfaced by the M2.6 fuzz reshuffle, seed 16:
                // [deleted, moved-away, deleted] left runs restored permuted).
                lastPairedLeft = i;
            }
        }

        var entries = new List<IrAlignedBlock>();

        // Front deletions (those preceding every paired left block).
        EmitDeletions(deletionsAfterLeft, -1, leftBlocks, entries);

        for (int j = 0; j < rightBlocks.Count; j++)
        {
            // A Split-stamped right index emits the group's ONE entry iff it is the FIRST member;
            // every other member is consumed silently (the TryGetValue miss below) — and must not
            // re-flush the group's deletion bucket, which the generic path would (every member's
            // rightMatch points at the same left index).
            if (rightKind[j] == IrAlignmentKind.Split)
            {
                if (splitByFirstMember.TryGetValue(j, out var sg))
                {
                    entries.Add(new IrAlignedBlock(IrAlignmentKind.Split, leftBlocks[sg.SingularIndex], null,
                        IrNodeList.From(sg.PluralIndexes.Select(rj => rightBlocks[rj]).ToList())));
                    EmitDeletions(deletionsAfterLeft, sg.SingularIndex, leftBlocks, entries);
                }

                continue;
            }

            if (rightKind[j] == IrAlignmentKind.Merge)
            {
                var mg = mergeByRight[j];
                entries.Add(new IrAlignedBlock(IrAlignmentKind.Merge, null, rightBlocks[j],
                    IrNodeList.From(mg.PluralIndexes.Select(mi => leftBlocks[mi]).ToList())));

                // Flush the deletion bucket of EVERY left member, in ascending left order — a
                // deletion anchored to a non-final member must still flush exactly once.
                foreach (int mi in mg.PluralIndexes)
                    EmitDeletions(deletionsAfterLeft, mi, leftBlocks, entries);
                continue;
            }

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

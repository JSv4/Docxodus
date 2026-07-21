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
/// <para><b>Granularity.</b> Tables, section breaks, block-level content controls, and opaque blocks align as
/// WHOLE units — keyed on their <c>ContentHash</c>/<c>FormatFingerprint</c> like any other block. Row/cell-level
/// table alignment is M2.2+; block content controls deliberately remain atomic so their OOXML envelope is never
/// reconstructed from independently aligned descendants.</para>
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
        => AlignBlocks(left.Body.Blocks, right.Body.Blocks, settings, markBodyFullRewriteGaps: true);

    /// <summary>
    /// Align two raw block lists (M2.2 Task 4 generalization). The public <see cref="Align"/> calls this
    /// with the bodies; <see cref="IrTableDiffer"/> calls it on a table CELL's block list to recurse the
    /// same machinery into cell contents. Identical semantics — anchoring, LIS spine, gap fill, fuzzy
    /// moves — just over an arbitrary block list rather than a document body.
    /// </summary>
    public static IrBlockAlignment AlignBlocks(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks, IrDiffSettings settings)
        => AlignBlocks(leftBlocks, rightBlocks, settings, markBodyFullRewriteGaps: false);

    /// <summary>
    /// Shared block-list implementation. Only the document-body entry point enables
    /// <paramref name="markBodyFullRewriteGaps"/>: a full lexical 1×1 rewrite in a cell, textbox,
    /// note, header, or footer keeps the normal Word-shaped seam.
    /// </summary>
    private static IrBlockAlignment AlignBlocks(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks, IrDiffSettings settings,
        bool markBodyFullRewriteGaps)
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
        // index. On-spine candidates keep their anchor kind (Unchanged/FormatOnly); off-spine become a
        // plain move when format-equal, or a move-and-modify when the content-equal pair has a format delta.
        candidates.Sort((a, b) => a.LeftIndex.CompareTo(b.LeftIndex));
        var onSpine = LongestIncreasingSubsequence(candidates);

        // A reorder can have SEVERAL longest spines (e.g. [A, table, B] → [A, B, table]: keeping {A, table}
        // or {A, B} both cost one relocation). When the arbitrary patience-sort pick relocates a heavy,
        // STRUCTURAL block (a table / section break / opaque block) that an equal-length spine could instead
        // anchor, re-pick the spine that keeps the most structural blocks anchored — so the lighter PARAGRAPH
        // is the one relocated. Beyond producing cleaner 2-way markup, this is what lets a paragraph move that
        // crosses a table boundary compose in the N-way consolidate (a spuriously-moved table contests the
        // whole block, blocking per-cell composition of another reviewer's disjoint table edit — issue #229).
        // The guard keeps the common path byte-identical: it fires ONLY when a structural block is off-spine.
        if (AnyOffSpineStructuralBlock(candidates, onSpine, leftBlocks))
            onSpine = LongestIncreasingSubsequencePreferringStructuralAnchors(candidates, leftBlocks, nRight);

        for (int c = 0; c < candidates.Count; c++)
        {
            var cand = candidates[c];
            if (onSpine.Contains(c))
            {
                leftKind[cand.LeftIndex] = cand.AnchorKind;
                rightKind[cand.RightIndex] = cand.AnchorKind;
            }
            else if (leftBlocks[cand.LeftIndex] is IrSdtBlock || rightBlocks[cand.RightIndex] is IrSdtBlock)
            {
                // A block-level content control that relocated must lower to a delete+insert pair, rather
                // than native move markup. Its envelope owns non-run metadata and can only round-trip when
                // rendered as one whole control at each side.  Release the exact-content candidate so the
                // ordinary gap logic emits the two structural operations in their proper locations.
                leftMatch[cand.LeftIndex] = -1;
                rightMatch[cand.RightIndex] = -1;
            }
            else
            {
                // Off-spine content anchor = relocated. A content-equal PARAGRAPH pair can nevertheless carry
                // a modeled format delta (for example, a paragraph moved and made bold). It needs the same
                // in-move token projection as a lexical move-and-edit so the destination can retain the right
                // formatting while rPrChange restores the left formatting on reject. Structural blocks remain
                // plain moves until they have an equivalent reversible format projection.
                var kind = cand.AnchorKind == IrAlignmentKind.FormatOnly &&
                    leftBlocks[cand.LeftIndex] is IrParagraph && rightBlocks[cand.RightIndex] is IrParagraph
                    ? IrAlignmentKind.MovedModified
                    : IrAlignmentKind.Moved;
                leftKind[cand.LeftIndex] = kind;
                rightKind[cand.RightIndex] = kind;
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

        // Renderer-only body provenance: an explicit shared id survives on the two standalone
        // Delete/Insert entries of a 1×1 full lexical rewrite. It deliberately lives beside, rather
        // than inside, the alignment kind/match arrays: the two blocks are still genuinely unpaired.
        // Nested callers use the same engine with marking disabled, so their normal seam behavior is
        // byte-for-byte unchanged — and do not allocate these per-block arrays.
        int?[]? leftBodyFullRewriteGroups = markBodyFullRewriteGaps ? new int?[nLeft] : null;
        int?[]? rightBodyFullRewriteGroups = markBodyFullRewriteGaps ? new int?[nRight] : null;
        int nextBodyFullRewriteGroupId = 1;

        FillGaps(leftBlocks, rightBlocks, spinePairs, leftKind, rightKind, leftMatch, rightMatch,
            similarity, settings, splitGroups, mergeGroups,
            leftBodyFullRewriteGroups, rightBodyFullRewriteGroups, markBodyFullRewriteGaps,
            ref nextBodyFullRewriteGroupId);

        // A renderer emits paired in-place blocks in RIGHT order.  Therefore every such pairing must be
        // monotone: crossed Modified pairs accept to the right sequence, but reject leaves their LEFT content
        // in that same right order. SimilarityPair intentionally considers correspondence by content rather
        // than position, so normalize any crossed fuzzy pair back to Delete+Insert before move detection. The
        // global move pass can then promote sufficiently strong correspondence to MovedModified; weaker cases
        // stay conservative and, crucially, reversible.
        ReleaseCrossingModifiedPairs(leftKind, rightKind, leftMatch, rightMatch);

        // --- Cross-gap fuzzy moves: over the GLOBAL leftover Deleted × Inserted sets (after all gap
        // fill), re-pair similar blocks as Moved / MovedModified. Runs AFTER gap fill so it sees the
        // final Deleted/Inserted leftovers, never blocks already consumed in-place.
        DetectCrossGapMoves(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch, similarity, settings);

        // --- Emit in right order with left-anchored deletion interleave.
        var entries = EmitEntries(leftBlocks, rightBlocks, leftKind, rightKind, leftMatch, rightMatch,
            splitGroups, mergeGroups, leftBodyFullRewriteGroups, rightBodyFullRewriteGroups);
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

    /// <summary>True iff any OFF-spine candidate anchors a non-paragraph (structural) block — a table,
    /// section break or opaque block — that the arbitrary longest-spine pick chose to relocate. This is the
    /// cheap guard that keeps the structural-anchor rebalance off the common path: it walks the candidate list
    /// once and only reports true for the rare reorder where a heavy block was picked as the mover.</summary>
    private static bool AnyOffSpineStructuralBlock(
        List<Candidate> candidates, HashSet<int> onSpine, IrNodeList<IrBlock> leftBlocks)
    {
        for (int c = 0; c < candidates.Count; c++)
            if (!onSpine.Contains(c) && leftBlocks[candidates[c].LeftIndex] is not IrParagraph)
                return true;
        return false;
    }

    /// <summary>
    /// Re-pick the spine as a maximum-WEIGHT strictly-increasing subsequence (by <see cref="Candidate.RightIndex"/>)
    /// where each candidate weighs <c>BIG + (structural ? 1 : 0)</c> and <c>BIG = candidates.Count + 1</c>
    /// dominates any achievable structural-bonus total. Because BIG dominates, the result is still a LONGEST
    /// increasing subsequence — the number of relocated blocks is unchanged — and the bonus is a pure tie-break
    /// that, among all longest spines, keeps the MOST structural (non-paragraph) blocks anchored. So an
    /// ambiguous reorder relocates a light paragraph rather than a heavy table / section break.
    /// </summary>
    /// <remarks>
    /// <para>Runs only when <see cref="AnyOffSpineStructuralBlock"/> flags a relocated structural block, so the
    /// paragraph-only common case never pays for it. O(k log k) via a Fenwick prefix-max over the right-index
    /// domain (right indices are a permutation — each used once). Deterministic: <c>&gt;</c> comparisons keep
    /// the earliest-processed (smallest left index) candidate on every tie, in both the tree updates and the
    /// endpoint pick, so two <see cref="Align"/> calls on the same inputs return sequence-equal spines.</para>
    /// </remarks>
    private static HashSet<int> LongestIncreasingSubsequencePreferringStructuralAnchors(
        List<Candidate> candidates, IrNodeList<IrBlock> leftBlocks, int nRight)
    {
        int n = candidates.Count;
        var result = new HashSet<int>();
        if (n == 0)
            return result;

        long big = n + 1; // > any structural-bonus total (at most n), so cardinality stays the primary key.

        // Fenwick prefix-MAX over right-index positions 1..nRight (right value r → position r+1). Each cell
        // holds the best (weight, endingCandidateIndex) of any subsequence ending at a right value ≤ its range.
        var treeWeight = new long[nRight + 1];
        var treeIndex = new int[nRight + 1];
        Array.Fill(treeIndex, -1);

        void Update(int pos, long weight, int candIndex)
        {
            for (; pos <= nRight; pos += pos & -pos)
                if (weight > treeWeight[pos]) // strict → keep the earliest-processed candidate on ties
                {
                    treeWeight[pos] = weight;
                    treeIndex[pos] = candIndex;
                }
        }

        (long Weight, int Index) Query(int pos)
        {
            long bestWeight = 0;
            int bestIndex = -1;
            for (; pos > 0; pos -= pos & -pos)
                if (treeWeight[pos] > bestWeight) // strict → keep the earliest-processed candidate on ties
                {
                    bestWeight = treeWeight[pos];
                    bestIndex = treeIndex[pos];
                }
            return (bestWeight, bestIndex);
        }

        var dp = new long[n];
        var parent = new int[n];

        // Candidates are already sorted ascending by left index (strictly increasing), so a left-order walk
        // with a "right value strictly smaller" predecessor query yields strictly-increasing subsequences.
        for (int c = 0; c < n; c++)
        {
            int r = candidates[c].RightIndex; // 0-based
            long weight = big + (leftBlocks[candidates[c].LeftIndex] is IrParagraph ? 0 : 1);
            var (predWeight, predIndex) = Query(r); // positions 1..r cover right values 0..r-1 (strictly smaller)
            dp[c] = weight + predWeight;
            parent[c] = predIndex;
            Update(r + 1, dp[c], c);
        }

        long globalBest = long.MinValue;
        int globalEnd = -1;
        for (int c = 0; c < n; c++)
            if (dp[c] > globalBest) // strict → smallest candidate index on ties
            {
                globalBest = dp[c];
                globalEnd = c;
            }

        for (int c = globalEnd; c != -1; c = parent[c])
            result.Add(c);
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
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups,
        int?[]? leftBodyFullRewriteGroups, int?[]? rightBodyFullRewriteGroups,
        bool markBodyFullRewriteGaps, ref int nextBodyFullRewriteGroupId)
    {
        int prevLeft = -1, prevRight = -1;
        foreach (var (sl, sr) in spinePairs)
        {
            FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, sl, prevRight + 1, sr,
                leftKind, rightKind, leftMatch, rightMatch, similarity, settings, splitGroups, mergeGroups,
                leftBodyFullRewriteGroups, rightBodyFullRewriteGroups, markBodyFullRewriteGaps,
                ref nextBodyFullRewriteGroupId);
            prevLeft = sl;
            prevRight = sr;
        }
        // Tail gap (after the last spine pair, or the whole document if there were no spine pairs).
        FillOneGap(leftBlocks, rightBlocks, prevLeft + 1, leftBlocks.Count, prevRight + 1, rightBlocks.Count,
            leftKind, rightKind, leftMatch, rightMatch, similarity, settings, splitGroups, mergeGroups,
            leftBodyFullRewriteGroups, rightBodyFullRewriteGroups, markBodyFullRewriteGaps,
            ref nextBodyFullRewriteGroupId);
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
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups,
        int?[]? leftBodyFullRewriteGroups, int?[]? rightBodyFullRewriteGroups,
        bool markBodyFullRewriteGaps, ref int nextBodyFullRewriteGroupId)
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
        // Positional extension (Word-parity): Word merges the k-th old table into the k-th new table
        // of a replace gap with per-cell del+ins interleave even when the counts differ — surplus
        // tables on either side insert/delete whole. Zip the leftover tables in document order; each
        // zipped pair is "the same table edited" and feeds IrTableDiffer's row/cell diff.
        int tablePairs = Math.Min(tableLeft.Count, tableRight.Count);
        for (int t = 0; t < tablePairs; t++)
        {
            int li = tableLeft[t];
            int rj = tableRight[t];
            leftKind[li] = IrAlignmentKind.Modified;
            rightKind[rj] = IrAlignmentKind.Modified;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
            leftoverLeft.Remove(li);
            leftoverRight.Remove(rj);
        }

        // M2.6 1:N split / N:1 merge containment scan (gated by IrDiffSettings.DetectSplitMerge — ON by default).
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

        // Word-matcher junction pairing (calibrated against the Word-compare oracle corpus). Word's
        // replace-gap arrangement pairs old/new paragraphs by ITS OWN matcher first and only then
        // merges each pair into a single mixed ins+del paragraph; unpaired paragraphs stay separate.
        // The similarity pass above (Jaccard ≥ BlockSimilarityThreshold + locality) reproduces the
        // strong pairs; this pass reproduces the WEAK ones Word still forms — e.g. the corpus-decoded
        // "Subtitle Style Demo" ↔ "Superscript Demo" (one shared word, Jaccard 0.25) and
        // "Title Style Centered Demo" ↔ "Times New Roman Font Demo" (Jaccard 0.125) — via an
        // order-preserving LCS over the remaining free paragraphs with a word-overlap predicate.
        // Word does NOT pair on zero shared words ("Support Tickets" ↔ "Test 1 – Fixed Width Table")
        // nor on stopword-grade overlap (header_no_rels: "…with just an empty p…" vs "…with extra
        // bold emphasis." stayed separate), which the word-Jaccard floor encodes.
        JunctionPair(leftBlocks, rightBlocks, leftFrom, leftTo, rightFrom, rightTo,
            leftoverLeft, leftoverRight, leftKind, rightKind, leftMatch, rightMatch, similarity);

        // Unambiguous 1×1 residue → Modified. When exactly ONE free left and ONE free right survive
        // the threshold, there is no competing candidate to disambiguate: classifying the lone pair
        // as "the same block, edited" is the natural reading (and is what M2.1's positional pairing
        // did for an isolated edit). A genuine cross-gap relocation never reaches here as a 1×1 gap
        // residue (it occupies DIFFERENT gaps, handled by DetectCrossGapMoves), so this does not
        // manufacture false in-place edits out of moves.
        //
        // CONDITION (Word-matcher calibration): a PARAGRAPH residue pair that is a full LEXICAL
        // rewrite stays separate — the Word oracle keeps those as an ins-marked + a del-marked
        // paragraph ("24" ↔ "1.5 Line Spacing Demo"; forcing them into one Modified paragraph
        // token-interleaves two unrelated texts — the corpus scored it strictly worse). The
        // evidence test (IrBlockSimilarity.ResidueForcePair) is punctuation-trimmed + case-folded
        // raw-word overlap, and an atomic-only/empty paragraph (textboxes, images — no words at
        // all) always force-pairs: there is no lexical evidence to demand, and demoting it would
        // lose the nested textbox/image diff (WC013/WC019 round-trip regressions proved it).
        // Non-paragraph (or mixed) residues keep the unconditional behavior.
        if (leftoverLeft.Count == 1 && leftoverRight.Count == 1)
        {
            int li = leftoverLeft[0];
            int rj = leftoverRight[0];
            // An SDT envelope may pair atomically only with another SDT envelope. Pairing it with a
            // paragraph/table merely because it is the lone residue would hand the renderer two incompatible
            // ownership topologies and lose the content-control wrapper on one accept/reject path.
            bool sdtTopologyMismatch = (leftBlocks[li] is IrSdtBlock) != (rightBlocks[rj] is IrSdtBlock);
            bool forcePair = !sdtTopologyMismatch &&
                (leftBlocks[li] is not IrParagraph lp || rightBlocks[rj] is not IrParagraph rp ||
                 similarity.ResidueForcePair(lp, rp));
            if (forcePair)
            {
                leftKind[li] = IrAlignmentKind.Modified;
                rightKind[rj] = IrAlignmentKind.Modified;
                leftMatch[li] = rj;
                rightMatch[rj] = li;
                return;
            }

            // A body-only full lexical 1×1 rewrite is deliberately still Delete+Insert at the
            // alignment layer. Word's paragraph-mark arrangement depends on the FIRST following
            // in-place pair: when that pair is a real body block it keeps separate marked paragraphs
            // (the interior blue-underline→bold-italic rewrite and the head title before the
            // incompatible 3→4-column table); when the only follower is the trailing section-break
            // sentinel it emits one mixed paragraph. This is explicit adjacent alignment evidence,
            // not a renderer guess based on cardinality or text.
            bool HasFollowingBodyPair()
            {
                int nextLeft = li + 1, nextRight = rj + 1;
                return nextLeft < leftBlocks.Count && nextRight < rightBlocks.Count &&
                    leftMatch[nextLeft] == nextRight && rightMatch[nextRight] == nextLeft &&
                    leftKind[nextLeft] is not (IrAlignmentKind.Moved or IrAlignmentKind.MovedModified) &&
                    rightKind[nextRight] is not (IrAlignmentKind.Moved or IrAlignmentKind.MovedModified) &&
                    leftBlocks[nextLeft] is not IrSectionBreak && rightBlocks[nextRight] is not IrSectionBreak;
            }
            if (markBodyFullRewriteGaps &&
                leftBodyFullRewriteGroups is { } leftGroups &&
                rightBodyFullRewriteGroups is { } rightGroups &&
                HasFollowingBodyPair() &&
                leftBlocks[li] is IrParagraph && rightBlocks[rj] is IrParagraph)
            {
                int groupId = nextBodyFullRewriteGroupId++;
                leftGroups[li] = groupId;
                rightGroups[rj] = groupId;
            }
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
        // Locality prior (calibrated against the Word-compare oracle corpus): Word's stream diff
        // anchors an insertion next to its matched neighbors and deletes distant old content
        // WHOLESALE — it never pairs a paragraph with a weakly-similar counterpart at the far end of
        // a large gap (the result renders as interleaved "word salad" inside an unrelated
        // paragraph). Eligibility therefore scales with the pair's relative displacement inside the
        // gap: sim ≥ threshold + λ·|relPosLeft − relPosRight|. High-similarity pairs still match
        // across the whole gap (a swapped-and-edited paragraph at displacement 1 needs
        // threshold + λ — e.g. 0.65 — which a genuine edit clears); weak pairs only form near their
        // own position. Ranking subtracts the same penalty so near pairs beat far pairs on ties.
        var positions = GapRelativePositions(freeLeft, leftMatch, freeRight, rightMatch);
        while (true)
        {
            double bestEffective = double.NegativeInfinity;
            int bestLeft = -1, bestRight = -1;
            foreach (int li in freeLeft)
            {
                if (leftMatch[li] != -1)
                    continue;
                foreach (int rj in freeRight)
                {
                    if (rightMatch[rj] != -1)
                        continue;
                    // Empty-vs-empty paragraphs score a vacuous 1.0 ("identical content") that
                    // defeats the locality prior at any displacement — an empty freed elsewhere in
                    // the gap would relocate here as a Modified pair. Word never fuzzy-pairs
                    // empties; they belong to the in-order passes (which pair them monotonically)
                    // or fall out as plain delete/insert.
                    if (leftBlocks[li] is IrParagraph leftParagraph && rightBlocks[rj] is IrParagraph rightParagraph)
                    {
                        if (similarity.WordCount(leftParagraph) == 0 && similarity.WordCount(rightParagraph) == 0)
                            continue;

                        // The general score includes spaces and punctuation because it deliberately
                        // shares the downstream token model.  At the lowered similarity threshold,
                        // two unrelated prose paragraphs can therefore clear the score merely by
                        // having the same separator skeleton (especially numbered lists).  Weak
                        // in-gap pairing needs actual lexical evidence; the dedicated 1×1 residue
                        // rule below still handles unambiguous labels and atomic-only paragraphs.
                        if (similarity.PairingWordOverlap(leftParagraph, rightParagraph).SharedWords == 0)
                            continue;
                    }
                    double score = similarity.Score(leftBlocks[li], rightBlocks[rj]);
                    double displacement = Math.Abs(positions.Left[li] - positions.Right[rj]);
                    if (score < threshold + PairLocalityPenalty * displacement)
                        continue;
                    double effective = score - PairLocalityPenalty * displacement;
                    // Strictly-greater wins; on a tie keep the first seen (freeLeft / freeRight are in
                    // ascending index order), which is exactly "smallest left, then smallest right".
                    if (effective > bestEffective)
                    {
                        bestEffective = effective;
                        bestLeft = li;
                        bestRight = rj;
                    }
                }
            }

            if (bestLeft == -1)
                return;

            leftKind[bestLeft] = IrAlignmentKind.Modified;
            rightKind[bestRight] = IrAlignmentKind.Modified;
            leftMatch[bestLeft] = bestRight;
            rightMatch[bestRight] = bestLeft;
        }
    }

    /// <summary>λ of the similarity-pair locality prior — see the comment in
    /// <see cref="SimilarityPair"/>. At 0.3 with the 0.35 base floor, a full-gap displacement demands
    /// 0.65 similarity (a genuine swapped edit clears it; corpus word-salad pairs do not).</summary>
    private const double PairLocalityPenalty = 0.3;

    /// <summary>Relative position (0..1) of each still-free index within its side's free list — the
    /// coordinate system of the locality prior. Single-element sides sit at 0 so a lone insertion
    /// measures its displacement against the gap HEAD (where Word anchors it), not the middle.</summary>
    private static (double[] Left, double[] Right) GapRelativePositions(
        List<int> freeLeft, int[] leftMatch, List<int> freeRight, int[] rightMatch)
    {
        var left = new double[leftMatch.Length];
        var right = new double[rightMatch.Length];
        var ls = freeLeft.Where(li => leftMatch[li] == -1).ToList();
        var rs = freeRight.Where(rj => rightMatch[rj] == -1).ToList();
        for (int i = 0; i < ls.Count; i++)
            left[ls[i]] = ls.Count == 1 ? 0 : (double)i / (ls.Count - 1);
        for (int j = 0; j < rs.Count; j++)
            right[rs[j]] = rs.Count == 1 ? 0 : (double)j / (rs.Count - 1);
        return (left, right);
    }

    // ------------------------------------------------------------------ junction pairing (Word matcher parity)

    // Calibrated constants (empirically fitted against the Word-compare oracle corpus, 2026-07;
    // per-variant subset means and the decoded oracle data points are recorded in the commit
    // message / CHANGELOG). Candidate predicates measured: ≥1 shared word (mean +5.13, 3 docs
    // regressed >2pts), ≥2 shared words (−0.68), shared word + hard displacement cap (+4.61,
    // 3 regressed), word-Jaccard floors 0.10/0.15/0.20 (+4.06/+3.40/+1.74), Jaccard+λ·displacement
    // (+4.34), + diagonal growth (+5.65), + conditional 1×1 (+5.71), + growth size-parity (+6.00,
    // ZERO docs regressed >2pts — the shipped configuration).

    /// <summary>A regular junction pair must share at least this many lexical WORD tokens —
    /// zero-shared paragraphs never pair (oracle: "Support Tickets" ↔ "Test 1 – Fixed Width Table"
    /// stayed separate). The labeled calendar-date bridge is a separately bounded LCS fallback.</summary>
    private const int JunctionMinSharedWords = 1;

    /// <summary>Word-token Jaccard floor of the junction LCS. 0.10 splits the decoded oracle
    /// boundary: "Title Style Centered Demo" ↔ "Times New Roman Font Demo" (0.125, Word PAIRS) vs
    /// header_no_rels's stopword-grade overlap ("…with just an empty p…" ↔ "…with bold creates the
    /// strongest…", 0.091, Word keeps separate).</summary>
    private const double JunctionMinWordJaccard = 0.10;

    /// <summary>λ of the junction locality term: the Jaccard floor grows by λ·|relative
    /// displacement|, so weak pairs only form near their own position (same discipline as
    /// <see cref="PairLocalityPenalty"/> — Word deletes distant old content wholesale rather than
    /// pairing it with a weakly-similar counterpart across the gap).</summary>
    private const double JunctionDispLambda = 0.3;

    /// <summary>Growth size-parity guard: on shared-word-only evidence a paragraph does not pair
    /// with one more than ~3× its word count (oracle: the 30-word justified body does NOT merge
    /// into the 7-word "This document demonstrates large 24pt font size." although they share
    /// "This document demonstrates").</summary>
    private const double JunctionGrowRatio = 1.0 / 3;

    /// <summary>
    /// LCS bound: the pass is skipped when the free-paragraph grid exceeds this product, keeping the
    /// aligner inside its documented G²-class gap budget (the adversarial scale guard).
    /// </summary>
    private const int JunctionPairScaleCeiling = 10000;

    /// <summary>
    /// Order-preserving junction pairing over a gap's remaining free PARAGRAPHS (Word-matcher
    /// parity — see the call-site comment in <see cref="FillOneGap"/>). Computes the maximum
    /// weighted longest-common-subsequence over the (still-ordered) leftover paragraph lists where a
    /// pair is admissible iff it (a) does not cross an already-formed non-Moved pair of this gap,
    /// (b) shares at least <see cref="JunctionMinSharedWords"/> lexical WORD tokens
    /// (punctuation/whitespace/numeric-ordinal overlap is no evidence — Word never pairs on it), or
    /// is the bounded labeled calendar-date bridge, and (c) clears the displacement-scaled
    /// word-Jaccard floor <see cref="JunctionMinWordJaccard"/> + <see cref="JunctionDispLambda"/>·disp.
    /// Chosen pairs become <see cref="IrAlignmentKind.Modified"/> (rendered as Word's single mixed
    /// ins+del paragraph); the rest fall through to Deleted/Inserted exactly as before. LCS
    /// maximizes pair COUNT first, then total word-Jaccard (deterministic index-order tie-break),
    /// and is non-crossing among its own picks by construction. A diagonal growth phase then
    /// extends pairings outward from every in-gap pair (see the inline comment).
    /// </summary>
    private static void JunctionPair(
        IrNodeList<IrBlock> leftBlocks, IrNodeList<IrBlock> rightBlocks,
        int leftFrom, int leftTo, int rightFrom, int rightTo,
        List<int> leftoverLeft, List<int> leftoverRight,
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch,
        IrBlockSimilarity similarity)
    {
        var ls = new List<int>();
        foreach (int li in leftoverLeft)
            if (leftBlocks[li] is IrParagraph)
                ls.Add(li);
        var rs = new List<int>();
        foreach (int rj in leftoverRight)
            if (rightBlocks[rj] is IrParagraph)
                rs.Add(rj);
        int m = ls.Count, n = rs.Count;
        if (m == 0 || n == 0 || (long)m * n > JunctionPairScaleCeiling)
            return;

        // Non-crossing bounds versus pairs already formed INSIDE this gap (SimilarityPair Modified
        // pairs, table zips, split/merge groups; Moved/MovedModified are exempt — long-range
        // correspondence belongs to the move detector). All in-gap partners lie inside the gap, so a
        // single ascending/descending sweep over [leftFrom, leftTo) yields, for each candidate left,
        // the window of right indexes that keeps document order reconstructible on reject.
        var maxBelow = new int[m];
        var minAbove = new int[m];
        {
            int running = int.MinValue, k = 0;
            for (int i = leftFrom; i < leftTo && k < m; i++)
            {
                if (i == ls[k]) { maxBelow[k] = running; k++; continue; }
                if (leftMatch[i] != -1 &&
                    leftKind[i] != IrAlignmentKind.Moved && leftKind[i] != IrAlignmentKind.MovedModified)
                    running = Math.Max(running, leftMatch[i]);
            }
            running = int.MaxValue; k = m - 1;
            for (int i = leftTo - 1; i >= leftFrom && k >= 0; i--)
            {
                if (i == ls[k]) { minAbove[k] = running; k--; continue; }
                if (leftMatch[i] != -1 &&
                    leftKind[i] != IrAlignmentKind.Moved && leftKind[i] != IrAlignmentKind.MovedModified)
                    running = Math.Min(running, leftMatch[i]);
            }
        }

        // PAIRING-EVIDENCE discipline (both corpus-decoded): qualifying shared lexical content is either
        // (a) at least one shared word that is not an English closed-class function word — Word
        // never pairs replace-gap paragraphs on stopword-grade overlap alone ('2.2 Numbered (with
        // nested)' does NOT merge into 'Q1: Launch v2.0 with new dashboard' on the shared "with",
        // 34pts worse when it did; 'This text will be indented.' does NOT merge into 'This
        // document contains a hyperlink to a website.' on the sentence-initial "This", 19pts
        // worse — while shared CONTENT words pair even when repeated across the gap: 'Title',
        // 'Demo', 'Q1'); or (b) CONTAINMENT — the shared words cover at least HALF of the smaller
        // side's words, in which case even function-word overlap pairs (Word merges the paragraph
        // 'a' into "A) ST_OnOff values for <w:b> on a run:", 21pts better when we do too: a
        // mostly-contained paragraph is an extension, not a replacement).
        bool HasPairingEvidence(IrParagraph lp, IrParagraph rp, int sharedWords)
        {
            int minWords = Math.Min(similarity.PairingWordCount(lp), similarity.PairingWordCount(rp));
            if (minWords > 0 && sharedWords * 2 >= minWords)
                return true;
            var a = similarity.PairingWordKeys(lp);
            var b = similarity.PairingWordKeys(rp);
            var (small, large) = a.Count <= b.Count ? (a, b) : (b, a);
            foreach (var kv in small)
                if (large.ContainsKey(kv.Key) && !FunctionWords.Contains(kv.Key))
                    return true;
            return false;
        }

        // Pair weight: 0 = inadmissible; otherwise the word-Jaccard (used only as the LCS
        // secondary objective). Computed once per candidate cell; bags are cached per Align call.
        double Weight(int i, int j)
        {
            int li = ls[i], rj = rs[j];
            if (rj <= maxBelow[i] || rj >= minAbove[i])
                return 0;
            var lp = (IrParagraph)leftBlocks[li];
            var rp = (IrParagraph)rightBlocks[rj];
            var (shared, jaccard) = similarity.PairingWordOverlap(lp, rp);
            double posL = m == 1 ? 0 : (double)i / (m - 1);
            double posR = n == 1 ? 0 : (double)j / (n - 1);
            double disp = Math.Abs(posL - posR);
            double floor = JunctionMinWordJaccard + JunctionDispLambda * disp;
            if (shared >= JunctionMinSharedWords)
            {
                if (jaccard < floor || !HasPairingEvidence(lp, rp, shared))
                    return 0;
                return jaccard;
            }

            // Word's product-roadmap oracle has a single semantic exception: a short title ending
            // in a year pairs with a labeled Date: Month day, year paragraph. Do not promote a
            // generic year into weak evidence or diagonal-growth fuel.
            if (!similarity.IsLabeledCalendarDateBridge(lp, rp))
                return 0;
            var (bridgeShared, bridgeJaccard) = similarity.JunctionWordOverlap(lp, rp);
            return bridgeShared >= JunctionMinSharedWords && bridgeJaccard >= floor ? bridgeJaccard : 0;
        }

        var weight = new double[m, n];
        for (int i = 0; i < m; i++)
            for (int j = 0; j < n; j++)
                weight[i, j] = Weight(i, j);

        // Weighted LCS DP: maximize (pair count, total weight) lexicographically. dp[i, j] covers
        // ls[0..i) × rs[0..j). Deterministic: pure function of the weight grid + fixed preference
        // order (take > skip-left > skip-right) applied identically in DP and backtrack.
        var count = new int[m + 1, n + 1];
        var total = new double[m + 1, n + 1];
        for (int i = 1; i <= m; i++)
            for (int j = 1; j <= n; j++)
            {
                int bc = count[i - 1, j];
                double bt = total[i - 1, j];
                if (count[i, j - 1] > bc || (count[i, j - 1] == bc && total[i, j - 1] > bt))
                {
                    bc = count[i, j - 1];
                    bt = total[i, j - 1];
                }
                double w = weight[i - 1, j - 1];
                if (w > 0)
                {
                    int tc = count[i - 1, j - 1] + 1;
                    double tt = total[i - 1, j - 1] + w;
                    if (tc > bc || (tc == bc && tt > bt))
                    {
                        bc = tc;
                        bt = tt;
                    }
                }
                count[i, j] = bc;
                total[i, j] = bt;
            }

        // Backtrack (mirrors the DP's preference order).
        var pairs = new List<(int Li, int Rj)>();
        {
            int i = m, j = n;
            while (i > 0 && j > 0)
            {
                double w = weight[i - 1, j - 1];
                if (w > 0 && count[i, j] == count[i - 1, j - 1] + 1 &&
                    total[i, j] == total[i - 1, j - 1] + w)
                {
                    pairs.Add((ls[i - 1], rs[j - 1]));
                    i--; j--;
                }
                else if (count[i, j] == count[i - 1, j] && total[i, j] == total[i - 1, j])
                    i--;
                else
                    j--;
            }
        }

        foreach (var (li, rj) in pairs)
        {
            leftKind[li] = IrAlignmentKind.Modified;
            rightKind[rj] = IrAlignmentKind.Modified;
            leftMatch[li] = rj;
            rightMatch[rj] = li;
            leftoverLeft.Remove(li);
            leftoverRight.Remove(rj);
        }

        // Diagonal GROWTH (patience-style anchor extension): a free paragraph DIAGONALLY ADJACENT to
        // an already-formed pair of this gap pairs on ≥1 shared word alone — the neighbor pair is the
        // positional evidence the Jaccard floor otherwise demands. This reproduces Word pairing e.g.
        // "Demonstrating Heading 3 paragraph style." ↔ "Heading 4 style with right alignment and
        // italic formatting." (word-Jaccard 0.077 — below any defensible floor, but sitting right
        // under the paired demo titles). Unsupported stopword-grade pairs (header_no_rels) still
        // never form: growth only steps ±1 from a real pair, and each step needs a shared word.
        // Same scale ceiling as the LCS, measured on the FULL gap (growth scans it for seeds).
        if ((long)(leftTo - leftFrom) * (rightTo - rightFrom) <= JunctionPairScaleCeiling)
        {
            bool CrossesAny(int nl, int nr)
            {
                for (int i = leftFrom; i < leftTo; i++)
                {
                    if (leftMatch[i] == -1 ||
                        leftKind[i] == IrAlignmentKind.Moved || leftKind[i] == IrAlignmentKind.MovedModified)
                        continue;
                    if ((i < nl && leftMatch[i] > nr) || (i > nl && leftMatch[i] < nr))
                        return true;
                }
                return false;
            }

            var queue = new Queue<(int L, int R)>();
            for (int i = leftFrom; i < leftTo; i++)
                if (leftMatch[i] != -1 &&
                    leftKind[i] != IrAlignmentKind.Moved && leftKind[i] != IrAlignmentKind.MovedModified)
                    queue.Enqueue((i, leftMatch[i]));

            while (queue.Count > 0)
            {
                var (l, r) = queue.Dequeue();
                for (int d = -1; d <= 1; d += 2)
                {
                    int nl = l + d, nr = r + d;
                    if (nl < leftFrom || nl >= leftTo || nr < rightFrom || nr >= rightTo)
                        continue;
                    if (leftMatch[nl] != -1 || rightMatch[nr] != -1)
                        continue;
                    if (leftBlocks[nl] is not IrParagraph lp || rightBlocks[nr] is not IrParagraph rp)
                        continue;
                    var (shared, _) = similarity.PairingWordOverlap(lp, rp);
                    if (shared < JunctionMinSharedWords)
                        continue;
                    // Size-parity guard: on shared-word-only evidence a paragraph does not pair
                    // with one several times its word count (see JunctionGrowRatio).
                    int wl = similarity.PairingWordCount(lp), wr = similarity.PairingWordCount(rp);
                    if (Math.Min(wl, wr) < JunctionGrowRatio * Math.Max(wl, wr))
                        continue;
                    // Pairing-evidence discipline (same as the LCS): adjacency to a pair plus
                    // a shared function word is still no evidence — see HasPairingEvidence.
                    if (!HasPairingEvidence(lp, rp, shared))
                        continue;
                    if (CrossesAny(nl, nr))
                        continue;
                    leftKind[nl] = IrAlignmentKind.Modified;
                    rightKind[nr] = IrAlignmentKind.Modified;
                    leftMatch[nl] = nr;
                    rightMatch[nr] = nl;
                    leftoverLeft.Remove(nl);
                    leftoverRight.Remove(nr);
                    queue.Enqueue((nl, nr));
                }
            }
        }
    }

    /// <summary>
    /// English closed-class function words (case-insensitive) — words that carry no pairing
    /// evidence for the junction discipline (see <c>HasContentSharedWord</c> in
    /// <see cref="JunctionPair"/>). Non-English text simply finds no members here, so the
    /// discipline degrades to "any shared word" for such corpora (documented scope).
    /// </summary>
    private static readonly HashSet<string> FunctionWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "a", "an", "the", "and", "or", "but", "nor", "so", "yet",
        "of", "in", "on", "at", "by", "for", "with", "to", "from", "as",
        "into", "over", "under", "up", "down", "out", "off", "about", "after",
        "before", "between", "during", "through", "per", "via",
        "is", "are", "was", "were", "be", "been", "being", "am",
        "do", "does", "did", "have", "has", "had",
        "will", "would", "can", "could", "shall", "should", "may", "might", "must",
        "this", "that", "these", "those", "it", "its",
        "he", "she", "they", "them", "his", "her", "their",
        "we", "us", "our", "you", "your", "i", "me", "my",
        "not", "no", "if", "then", "than", "there", "here",
        "when", "where", "which", "who", "whom", "what", "why", "how",
        "all", "each", "both", "some", "any", "such", "same", "other", "another",
        "more", "most", "only", "just", "also", "too", "very", "own",
    };

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

            // A split/merge GROUP needs ≥ 2 members by definition — a 1-member "run" is just an
            // ordinary 1:1 pairing wearing a costume, and admitting it here would bypass the
            // similarity-pair bar entirely (the edge-trimmed containment gates are far laxer than
            // the Jaccard-with-locality bar; the corpus word-salad pairs slipped through exactly
            // this way). A true 1:1 edit either cleared SimilarityPair already or should lower to
            // Delete+Insert.
            if (run.Count < 2)
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

        // A split needs two phrase-sized inherited fragments, not merely two isolated retained
        // tokens. Without this guard `Video. Click.` → `Video.` + inserted math + `Click` is
        // indistinguishable to the raw LCS from a true split, yet Word treats it as one deleted
        // paragraph plus one coalesced inserted region (WC-1840). The genuine corpus splits and
        // the public split model both carry multiple inherited Word tokens in each textual half.
        const int MinMatchedWordsPerSplitFragment = 2;
        int phraseMembers = score.MemberMatchedWords.Count(n => n >= MinMatchedWordsPerSplitFragment);
        if (phraseMembers < 2)
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
        //
        // In-gap pairing is ORDER-PRESERVING: a candidate that crosses any already-formed non-Moved
        // pair is rejected outright (long-range correspondence belongs exclusively to the move
        // detector). Both phases enforce it — unids are content-derived and collide across distinct
        // blocks, so even the phase-1 "identity" reservation can propose a crossing pair. Without
        // the guard, content-equal empty paragraphs pair across intervening tables and the left's
        // empties silently RELOCATE (reject then reproduces the wrong block order).
        // Candidate (l, r) is order-safe iff maxJBelow[l] < r < minJAbove[l]; the bounds are
        // rebuilt after each accepted pair (O(n) per acceptance — within the documented G² budget)
        // so the per-candidate check stays O(1) and the scale guard holds.
        // Bounds are LAZY: allocated and filled on the first actual crossing check. InOrderRefine
        // runs for every gap and most gaps never reach a check (candidates fail the content-hash
        // gate first) — an eager O(n) rebuild per gap invocation compounds to O(n²) across a
        // document's gaps and trips the aligner's scale guard.
        int n = leftMatch.Length;
        int[]? maxJBelow = null;
        int[]? minJAbove = null;
        void RebuildBounds()
        {
            maxJBelow ??= new int[n];
            minJAbove ??= new int[n];
            int running = int.MinValue;
            for (int i = 0; i < n; i++)
            {
                maxJBelow[i] = running;
                if (leftMatch[i] != -1 && leftKind[i] != IrAlignmentKind.Moved)
                    running = Math.Max(running, leftMatch[i]);
            }
            running = int.MaxValue;
            for (int i = n - 1; i >= 0; i--)
            {
                minJAbove[i] = running;
                if (leftMatch[i] != -1 && leftKind[i] != IrAlignmentKind.Moved)
                    running = Math.Min(running, leftMatch[i]);
            }
        }
        bool Crosses(int l, int r)
        {
            if (maxJBelow is null)
                RebuildBounds();
            return r <= maxJBelow![l] || r >= minJAbove![l];
        }
        // Incremental bound propagation with early termination: each position's maxJBelow only
        // ever increases (minJAbove only decreases), so total propagation work across all
        // acceptances is O(n + inversions) — amortized linear for the monotone boilerplate case
        // the scale guard times, never worse than O(n) per acceptance.
        void NoteAccepted(int l0, int r0)
        {
            if (maxJBelow is null)
                return; // bounds never materialized — next Crosses() builds them fresh
            for (int i = l0 + 1; i < n && maxJBelow[i] < r0; i++) maxJBelow[i] = r0;
            for (int i = l0 - 1; i >= 0 && minJAbove![i] > r0; i--) minJAbove[i] = r0;
        }

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
                if (Crosses(candLeft, rj))
                    continue;

                leftKind[candLeft] = kind;
                rightKind[rj] = kind;
                leftMatch[candLeft] = rj;
                rightMatch[rj] = candLeft;
                NoteAccepted(candLeft, rj);
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
                if (Crosses(candLeft, rj))
                    continue;

                leftKind[candLeft] = kind;
                rightKind[rj] = kind;
                leftMatch[candLeft] = rj;
                rightMatch[rj] = candLeft;
                NoteAccepted(candLeft, rj);
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
    /// ContentHashes are equal <em>and their formats compare equal</em> the blocks are exact relocations,
    /// which classify as plain <see cref="IrAlignmentKind.Moved"/>. A content-equal paragraph format delta
    /// still needs <see cref="IrAlignmentKind.MovedModified"/> so its FormatChanged token spans are rendered.
    /// In practice exact moves are already caught by off-spine anchoring, but the guard keeps this fallback
    /// consistent.</para>
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
                // A block-level content control owns an OOXML envelope that cannot be expressed with
                // native move markup. Even if a future similarity model assigns it text/word counts,
                // keep it as the reversible delete+insert pair established above.
                if (leftBlocks[li] is IrSdtBlock)
                    continue;
                if (similarity.WordCount(leftBlocks[li]) < minTokens)
                    continue; // too short to be a reliable move (mirrors MoveMinimumWordCount)
                foreach (int rj in inserted)
                {
                    if (rightMatch[rj] != -1)
                        continue;
                    if (rightBlocks[rj] is IrSdtBlock)
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

            // A paragraph plain move needs BOTH exact content and format equality. Exact paragraph text with a
            // formatting delta still needs a token diff (FormatChanged spans) so the markup can carry rPrChange
            // at the moved destination and restore the old formatting on reject. Structural blocks do not yet
            // have an equivalent in-move format projection, so retain their existing plain-move behavior.
            var leftBlock = leftBlocks[bestLeft];
            var rightBlock = rightBlocks[bestRight];
            bool needsParagraphFormatProjection = leftBlock is IrParagraph && rightBlock is IrParagraph &&
                !FormatEqual(leftBlock, rightBlock, settings);
            bool exact = bestScore >= 1.0 && leftBlock.ContentHash.Equals(rightBlock.ContentHash) &&
                !needsParagraphFormatProjection;
            var kind = exact ? IrAlignmentKind.Moved : IrAlignmentKind.MovedModified;

            leftKind[bestLeft] = kind;
            rightKind[bestRight] = kind;
            leftMatch[bestLeft] = bestRight;
            rightMatch[bestRight] = bestLeft;
        }
    }

    /// <summary>
    /// Release every in-place <see cref="IrAlignmentKind.Modified"/> pair that participates in a crossing.
    /// Exact anchors and the order-preserving refinement passes are already monotone; only greedy similarity
    /// pairing can produce an inversion. Both endpoints of an inversion must be released — retaining just one
    /// crossed edit would still put its old content at the wrong physical paragraph when revisions are rejected.
    /// The prefix/suffix extrema identify every modified pair in an inversion in O(n), including permutations
    /// larger than a simple adjacent swap.
    /// </summary>
    private static void ReleaseCrossingModifiedPairs(
        IrAlignmentKind?[] leftKind, IrAlignmentKind?[] rightKind,
        int[] leftMatch, int[] rightMatch)
    {
        int n = leftMatch.Length;
        var maxRightBefore = new int[n];
        int maxRight = int.MinValue;
        for (int li = 0; li < n; li++)
        {
            maxRightBefore[li] = maxRight;
            if (IsOneToOneInPlace(leftKind[li], leftMatch[li]))
                maxRight = Math.Max(maxRight, leftMatch[li]);
        }

        var minRightAfter = new int[n];
        int minRight = int.MaxValue;
        for (int li = n - 1; li >= 0; li--)
        {
            minRightAfter[li] = minRight;
            if (IsOneToOneInPlace(leftKind[li], leftMatch[li]))
                minRight = Math.Min(minRight, leftMatch[li]);
        }

        var release = new List<int>();
        for (int li = 0; li < n; li++)
        {
            int rj = leftMatch[li];
            if (leftKind[li] != IrAlignmentKind.Modified || rj < 0)
                continue;
            if (rj < maxRightBefore[li] || rj > minRightAfter[li])
                release.Add(li);
        }

        foreach (int li in release)
        {
            int rj = leftMatch[li];
            // A Modified pair is necessarily one-to-one. Be defensive in case a future alignment kind reuses
            // this normalization path with a partially released counterpart.
            if (rj < 0 || rightMatch[rj] != li || rightKind[rj] != IrAlignmentKind.Modified)
                continue;
            leftKind[li] = IrAlignmentKind.Deleted;
            rightKind[rj] = IrAlignmentKind.Inserted;
            leftMatch[li] = -1;
            rightMatch[rj] = -1;
        }

        static bool IsOneToOneInPlace(IrAlignmentKind? kind, int rightIndex) =>
            rightIndex >= 0 && kind is IrAlignmentKind.Unchanged or IrAlignmentKind.FormatOnly or IrAlignmentKind.Modified;
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
        List<(int SingularIndex, List<int> PluralIndexes)> mergeGroups,
        int?[]? leftBodyFullRewriteGroups, int?[]? rightBodyFullRewriteGroups)
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
        EmitDeletions(deletionsAfterLeft, -1, leftBlocks, entries, leftBodyFullRewriteGroups);

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
                    EmitDeletions(deletionsAfterLeft, sg.SingularIndex, leftBlocks, entries,
                        leftBodyFullRewriteGroups);
                }

                continue;
            }

            if (rightKind[j] == IrAlignmentKind.Merge)
            {
                // A Merge-stamped right block always has exactly one recorded group (one right per merge),
                // so a miss is an engine bug — fail loud with a clear message rather than a bare
                // KeyNotFoundException (and never a silent skip, unlike Split's expected non-first-member miss).
                if (!mergeByRight.TryGetValue(j, out var mg))
                    throw new System.InvalidOperationException(
                        $"IrBlockAligner: right block {j} is Merge-stamped but no merge group was recorded for it.");
                entries.Add(new IrAlignedBlock(IrAlignmentKind.Merge, null, rightBlocks[j],
                    IrNodeList.From(mg.PluralIndexes.Select(mi => leftBlocks[mi]).ToList())));

                // Flush the deletion bucket of EVERY left member, in ascending left order — a
                // deletion anchored to a non-final member must still flush exactly once.
                foreach (int mi in mg.PluralIndexes)
                    EmitDeletions(deletionsAfterLeft, mi, leftBlocks, entries, leftBodyFullRewriteGroups);
                continue;
            }

            var kind = rightKind[j] ?? IrAlignmentKind.Inserted;
            int li = rightMatch[j];
            IrBlock? leftBlock = li != -1 ? leftBlocks[li] : null;
            entries.Add(new IrAlignedBlock(kind, leftBlock, rightBlocks[j],
                BodyFullRewriteGroupId: kind == IrAlignmentKind.Inserted && rightBodyFullRewriteGroups is { } rightGroups
                    ? rightGroups[j] : null));

            // After emitting a paired right block, flush deletions anchored to its left partner.
            if (li != -1)
                EmitDeletions(deletionsAfterLeft, li, leftBlocks, entries, leftBodyFullRewriteGroups);
        }

        return entries;
    }

    private static void EmitDeletions(
        Dictionary<int, List<int>> deletionsAfterLeft, int anchorLeftIndex,
        IrNodeList<IrBlock> leftBlocks, List<IrAlignedBlock> entries,
        int?[]? leftBodyFullRewriteGroups)
    {
        if (!deletionsAfterLeft.TryGetValue(anchorLeftIndex, out var list))
            return;
        foreach (int li in list) // already in ascending left order
            entries.Add(new IrAlignedBlock(IrAlignmentKind.Deleted, leftBlocks[li], null,
                BodyFullRewriteGroupId: leftBodyFullRewriteGroups is { } leftGroups ? leftGroups[li] : null));
    }
}

#nullable enable

using System;
using System.Collections.Generic;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Intra-block token differ (M2.2 Task 1): sequence-diffs two paragraph token lists by
/// <see cref="IrDiffToken.MatchKey"/>, then runs a format post-pass that splits content-equal runs
/// into <see cref="IrTokenOpKind.Equal"/> and <see cref="IrTokenOpKind.FormatChanged"/> spans by
/// per-token <see cref="IrRunFormat"/> record equality.
/// </summary>
/// <remarks>
/// <para><b>Algorithm.</b> Content tokens (everything but whitespace separators) are aligned by a
/// character-weighted LCS (<see cref="CharWeightedLcs"/>), and the anchor pairs it finds partition both
/// streams into segments emitted as whitespace-trimmed delete+insert (see <see cref="AnchoredSpans"/>). The
/// LCS table is O(n·m) in time and memory, so it only runs on regions within <see cref="LcsCellCap"/> cells;
/// a longer paragraph pair (a pasted transcript, a data dump) is first cut into such regions by
/// linear-memory anchoring (<see cref="BoundedAnchors"/>: common prefix/suffix, unique-key patience
/// anchors, rarest shared key). Ordinary paragraphs never reach the cap and get the exact table result.</para>
/// <para><b>Determinism.</b> The table's back-walk prefers advancing the left side on ties, and the
/// over-cap anchoring visits candidates in left order, so the op sequence is a pure function of the two
/// token lists. The format post-pass is a deterministic linear scan. Two <see cref="Diff"/> calls on the
/// same inputs return record-equal results.</para>
/// <para><b>Coalescing.</b> The per-token edit stream is coalesced into maximal same-kind spans
/// (Equal/Insert/Delete) before the format post-pass.</para>
/// </remarks>
internal static class IrTokenDiffer
{
    /// <summary>
    /// Diff <paramref name="left"/> against <paramref name="right"/> by <see cref="IrDiffToken.MatchKey"/>,
    /// producing format-refined token ops. <paramref name="endsAtRetainedMark"/> declares whether the END of
    /// both token streams abuts a RETAINED paragraph-mark pair (true for an ordinary Modified pair — Task 3
    /// treats every Modify as a same-mark edit — and for the final member of a split/merge or the Equal-mark
    /// cell of a cross-paragraph run; false for a slice that closes on a ¶INS/¶DEL mark).
    /// <paramref name="suppressLonePunctuation"/> gates the lone-punctuation anchor rule (see
    /// <see cref="SuppressUnanchoredPunctuation"/>): callers must pass false when either source paragraph
    /// carries anything but direct text runs — dropping an anchor inside a transparent field/hyperlink
    /// region re-tiles that region into del+ins, and the renderer then emits the container's zero-width
    /// plumbing for BOTH sides (a REF field surviving Accept twice — same containment discipline as the
    /// renderer's whitespace re-anchoring gate).
    /// </summary>
    public static IrTokenDiff Diff(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right, IrDiffSettings settings,
        bool endsAtRetainedMark = true, bool suppressLonePunctuation = true)
    {
        // 1. Raw token-grain edits from the content-anchored alignment, already coalesced into same-kind
        // spans. (MatchKey/Format were precomputed by the tokenizer under these settings; the alignment
        // keys on MatchKey.)
        var spans = AnchoredSpans(left, right, endsAtRetainedMark, suppressLonePunctuation);

        // 2. Format post-pass: split each Equal span into Equal / FormatChanged sub-spans. The
        // FormatComparison policy (M2.2 Task 4) decides whether unmodeled rPr noise (lang/bCs/iCs/…)
        // raises a FormatChanged span — ModeledOnly (default) ignores it.
        var ops = new List<IrTokenOp>(spans.Count);
        foreach (var span in spans)
        {
            if (span.Kind == IrTokenOpKind.Equal)
                SplitEqualByFormat(left, right, span, ops, settings.FormatComparison);
            else
                ops.Add(span);
        }

        return new IrTokenDiff(IrNodeList.From(ops));
    }

    // ------------------------------------------------------------------ content-anchored two-level diff

    /// <summary>
    /// Content-anchored token diff, returning coalesced same-kind <see cref="IrTokenOp"/> spans
    /// (Equal/Insert/Delete only — the format pass runs later). Insert spans carry an empty left span
    /// at the anchor index; Delete spans an empty right span.
    /// </summary>
    /// <remarks>
    /// A single all-token LCS keyed on <see cref="IrDiffToken.MatchKey"/> mis-anchors on whitespace:
    /// every separator shares the key <c>" "</c>, so with many identical spaces the LCS spends its
    /// budget matching spaces and DROPS interior shared CONTENT words (delete+re-insert them). Word
    /// anchors on content words, not whitespace. We do the same in two levels:
    /// <list type="number">
    /// <item>Run the LCS over the subsequence of NON-connective (content) tokens only — a token is
    /// connective iff it is a whitespace-only <see cref="IrDiffTokenKind.Separator"/>; Words, punctuation
    /// separators, and the atomic kinds all count as content and CAN anchor. This yields ordered Equal
    /// content-anchor pairs mapped back to full-stream indices.</item>
    /// <item>Partition both full streams at the anchors and emit, per segment, a common WHITESPACE prefix
    /// and suffix as Equal with the middle as Delete(all left)+Insert(all right) — no nested all-token
    /// LCS (that would reintroduce the whitespace crowding).</item>
    /// </list>
    /// The forward per-token edit stream feeds the shared <see cref="Coalesce"/> so anchors merge with
    /// adjacent whitespace Equal and consecutive Delete/Insert merge into maximal spans.
    /// </remarks>
    private static List<IrTokenOp> AnchoredSpans(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right,
        bool endsAtRetainedMark, bool suppressLonePunctuation)
    {
        int n = left.Count, m = right.Count;
        var spans = new List<IrTokenOp>();

        // Degenerate sides: a single Delete (whole left) and/or Insert (whole right).
        if (n == 0 && m == 0)
            return spans;
        if (n == 0)
        {
            spans.Add(new IrTokenOp(IrTokenOpKind.Insert, 0, 0, 0, m));
            return spans;
        }
        if (m == 0)
        {
            spans.Add(new IrTokenOp(IrTokenOpKind.Delete, 0, n, 0, 0));
            return spans;
        }

        // 1. Content-token anchors (full-stream index pairs, strictly increasing on both sides).
        // Atomic tokens (note refs, images, tabs, breaks, opaque, textboxes) share coarse MatchKeys
        // (every footnote ref keys "fn"), so anchoring the content LCS on them can pair the WRONG
        // occurrence and mis-attribute which one was inserted — breaking note/definition reject
        // round-trips. When either side carries an atomic token, fall back to all-token anchoring
        // (the pre-content-anchor behavior: whitespace participates, so an atomic token is paired
        // in its full surrounding context). Pure text/word/space paragraphs — where whitespace
        // crowding is the problem this pass exists to fix — take the content-anchored path.
        bool anchorAll = HasAtomic(left) || HasAtomic(right);
        var anchors = ContentAnchors(left, right, anchorAll);

        // 1b. Lone-punctuation rule (decoded 2026-07-27 from reference compare output): an isolated
        // matched PUNCTUATION token does not stand as an anchor — Word duplicates it into both the
        // ins and del regions ("margin." / "italic." each keep their own period) — UNLESS the match is
        // ANCHORED: contiguous (bridging equal whitespace) with a matched word pair, with another
        // anchored punctuation pair, or with the retained paragraph-mark pair at the end of both
        // streams. Skipped on the atomic (all-token) path, which pairs tokens in full context already,
        // and when the caller declares a non-plain source paragraph (transparent containers).
        if (!anchorAll && suppressLonePunctuation)
            SuppressUnanchoredPunctuation(anchors, left, right, endsAtRetainedMark);

        // 2. Partition at anchors: for each anchor, emit the segment before it, then the anchor as Equal.
        var edits = new List<(IrTokenOpKind Kind, int Left, int Right)>();
        int li = 0, ri = 0;
        foreach (var (al, ar) in anchors)
        {
            EmitSegment(left, right, li, al, ri, ar, edits);
            edits.Add((IrTokenOpKind.Equal, al, ar));
            li = al + 1;
            ri = ar + 1;
        }

        // Trailing segment after the last anchor.
        EmitSegment(left, right, li, n, ri, m, edits);

        // 3. Coalesce the forward per-token edit stream into maximal same-kind spans.
        Coalesce(edits, spans);
        return spans;
    }

    /// <summary>True for a connective token: a whitespace-only <see cref="IrDiffTokenKind.Separator"/>.
    /// These are the tokens that MUST NOT anchor the diff (else the LCS crowds on abundant spaces).</summary>
    private static bool IsConnective(IrDiffToken t) =>
        t.Kind == IrDiffTokenKind.Separator && string.IsNullOrWhiteSpace(t.Text);

    /// <summary>True if any token is atomic (not a Word or Separator) — a note ref, image, tab,
    /// break, opaque inline, or textbox. These carry coarse MatchKeys and must be paired in full
    /// context (all-token anchoring), never content-anchored, or note/definition reject can break.</summary>
    private static bool HasAtomic(IReadOnlyList<IrDiffToken> tokens)
    {
        foreach (var t in tokens)
            if (t.Kind is not (IrDiffTokenKind.Word or IrDiffTokenKind.Separator))
                return true;
        return false;
    }

    /// <summary>
    /// Compute the ordered content-anchor pairs: the LCS (<see cref="BoundedAnchors"/>) over the non-connective (content) token
    /// subsequences of <paramref name="left"/> and <paramref name="right"/>, keyed on MatchKey, mapped
    /// back to full-stream indices. Strictly increasing on both sides.
    /// </summary>
    private static List<(int Left, int Right)> ContentAnchors(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right, bool anchorAll)
    {
        var leftContent = new List<int>();
        for (int i = 0; i < left.Count; i++)
            if (anchorAll || !IsConnective(left[i]))
                leftContent.Add(i);

        var rightContent = new List<int>();
        for (int j = 0; j < right.Count; j++)
            if (anchorAll || !IsConnective(right[j]))
                rightContent.Add(j);

        var leftKeys = new string[leftContent.Count];
        var weights = new int[leftContent.Count];
        for (int a = 0; a < leftKeys.Length; a++)
        {
            leftKeys[a] = left[leftContent[a]].MatchKey;
            weights[a] = left[leftContent[a]].Text.Length;
        }
        var rightKeys = new string[rightContent.Count];
        for (int b = 0; b < rightKeys.Length; b++)
            rightKeys[b] = right[rightContent[b]].MatchKey;

        var pairs = BoundedAnchors(leftKeys, rightKeys, weights);

        var anchors = new List<(int, int)>(pairs.Count);
        foreach (var (a, b) in pairs)
            anchors.Add((leftContent[a], rightContent[b]));
        return anchors;
    }

    /// <summary>
    /// Drop matched punctuation anchors that are NOT anchored to solid context, mutating
    /// <paramref name="anchors"/> in place. Decoded from Word's compare output (2026-07-27): word matches
    /// always stand, but an isolated punctuation match ("margin." vs "italic." sharing only the period) is
    /// duplicated into both changed regions. A punctuation anchor STANDS iff — walking over pairwise
    /// key-equal whitespace on both sides — it reaches (a) a matched WORD anchor, (b) another standing
    /// punctuation anchor, or (c) the end of BOTH streams when <paramref name="endsAtRetainedMark"/>
    /// (the retained paragraph-mark pair is the anchor there: "size 24." vs "size 18 point text." retains
    /// the final period because both periods abut the retained pilcrow). Chains of punctuation resolve by
    /// fixed-point propagation from the solid ends; a pure punctuation island with changed words on both
    /// sides and a marked/absent terminal never stands.
    /// </summary>
    private static void SuppressUnanchoredPunctuation(
        List<(int Left, int Right)> anchors,
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right,
        bool endsAtRetainedMark)
    {
        if (anchors.Count == 0)
            return;

        // Any punctuation anchors at all? (Words are the overwhelming majority; exit cheap.)
        bool anyPunct = false;
        foreach (var (al, _) in anchors)
            anyPunct |= left[al].Kind == IrDiffTokenKind.Separator;
        if (!anyPunct)
            return;

        int n = left.Count, m = right.Count;
        var solid = new bool[anchors.Count];             // words are solid from the start
        for (int i = 0; i < anchors.Count; i++)
            solid[i] = left[anchors[i].Left].Kind == IrDiffTokenKind.Word;

        var pairIndex = new Dictionary<(int, int), int>(anchors.Count);
        for (int i = 0; i < anchors.Count; i++)
            pairIndex[anchors[i]] = i;

        // Walk from (l, r) one step at a time over pairwise key-equal connective whitespace in the given
        // direction; return the first non-connective position pair (or (-1,-1) when the sides desync).
        (int L, int R) BridgeFrom(int l, int r, int dir)
        {
            while (true)
            {
                l += dir;
                r += dir;
                if (l < 0 && r < 0)
                    return (l, r);                       // start of both streams
                if (l >= n && r >= m)
                    return (l, r);                       // end of both streams
                if (l < 0 || r < 0 || l >= n || r >= m)
                    return (-1, -1);                     // one side ran out — no bridge
                if (!IsConnective(left[l]) || !IsConnective(right[r]))
                    return (l, r);
                if (left[l].MatchKey != right[r].MatchKey)
                    return (-1, -1);
            }
        }

        bool changed = true;
        while (changed)
        {
            changed = false;
            for (int i = 0; i < anchors.Count; i++)
            {
                if (solid[i] || left[anchors[i].Left].Kind == IrDiffTokenKind.Word)
                    continue;

                var back = BridgeFrom(anchors[i].Left, anchors[i].Right, -1);
                bool anchoredBack = back != (-1, -1) &&
                    pairIndex.TryGetValue(back, out var bi) && solid[bi];

                var fwd = BridgeFrom(anchors[i].Left, anchors[i].Right, +1);
                bool anchoredFwd = fwd != (-1, -1) &&
                    ((fwd.L >= n && fwd.R >= m && endsAtRetainedMark) ||
                     (pairIndex.TryGetValue(fwd, out var fi) && solid[fi]));

                if (anchoredBack || anchoredFwd)
                {
                    solid[i] = true;
                    changed = true;
                }
            }
        }

        for (int i = anchors.Count - 1; i >= 0; i--)
            if (!solid[i])
                anchors.RemoveAt(i);
    }

    /// <summary>
    /// Largest anchor region (left·right content tokens) aligned by the exact
    /// <see cref="CharWeightedLcs"/> table: 4M cells, a 16 MB table, about 2,000 words a side. The table is
    /// O(n·m) in time and memory — a 20k-word paragraph pair would need 1.6 GB — so larger regions are first
    /// cut by <see cref="BoundedAnchors"/>. Set above the longest paragraph pair in the test corpus (a
    /// ~1,370-word clause, 1.9M cells), so real paragraphs keep exactly the table's anchors.
    /// </summary>
    private const long LcsCellCap = 4_000_000;

    /// <summary>
    /// A key repeating more often than this on either side of an over-cap region is too ambiguous to
    /// split it on when the region holds no key that is unique on both sides.
    /// </summary>
    private const int MaxSplitKeyOccurrences = 64;

    /// <summary>
    /// Ordered anchor pairs <c>(a, b)</c> (strictly increasing on both sides) between the key sequences
    /// <paramref name="leftKeys"/> and <paramref name="rightKeys"/>. A pair within <see cref="LcsCellCap"/>
    /// gets the exact character-weighted LCS. A larger pair is cut down with linear-memory steps until each
    /// remaining region fits under the cap: matching the common prefix and suffix, then anchoring on the
    /// longest increasing run of keys that occur exactly once on each side (patience diff), or — when no key
    /// is unique on both sides — on the first occurrence of the rarest shared key. Every step strictly
    /// shrinks the region, and the total scanning work is bounded by a multiple of the input length; a region
    /// left over once that budget is spent stays unanchored (deleted and reinserted), which is the correct
    /// output for text that shares no rare words anyway.
    /// </summary>
    private static List<(int A, int B)> BoundedAnchors(string[] leftKeys, string[] rightKeys, int[] weights)
    {
        int n = leftKeys.Length, m = rightKeys.Length;
        var matches = new List<(int A, int B)>();
        if (n == 0 || m == 0)
            return matches;
        int[]? table = null;
        if ((long)n * m <= LcsCellCap)
        {
            CharWeightedLcs(leftKeys, 0, n, rightKeys, 0, m, weights, matches, ref table);
            return matches;
        }

        long scanBudget = 64L * (n + m);
        var regions = new Stack<(int Ls, int Le, int Rs, int Re)>();
        regions.Push((0, n, 0, m));
        while (regions.Count > 0)
        {
            var (ls, le, rs, re) = regions.Pop();
            while (ls < le && rs < re && leftKeys[ls] == rightKeys[rs])
                matches.Add((ls++, rs++));
            while (ls < le && rs < re && leftKeys[le - 1] == rightKeys[re - 1])
                matches.Add((--le, --re));
            if (ls == le || rs == re)
                continue;
            if ((long)(le - ls) * (re - rs) <= LcsCellCap)
            {
                CharWeightedLcs(leftKeys, ls, le, rightKeys, rs, re, weights, matches, ref table);
                continue;
            }

            scanBudget -= (le - ls) + (re - rs);
            if (scanBudget < 0)
                continue;

            int pl = ls, pr = rs;
            foreach (var (a, b) in SplitAnchors(leftKeys, ls, le, rightKeys, rs, re))
            {
                matches.Add((a, b));
                regions.Push((pl, a, pr, b));
                pl = a + 1;
                pr = b + 1;
            }
            if (pl != ls)
                regions.Push((pl, le, pr, re));
        }

        matches.Sort();
        return matches;
    }

    /// <summary>
    /// Anchors that cut the region <c>left[ls..le) × right[rs..re)</c>: the longest increasing run (by right
    /// position, in left order) of the keys occurring exactly once on each side, else the first occurrences
    /// of the rarest key shared by both sides (at most <see cref="MaxSplitKeyOccurrences"/> per side), else
    /// none. Deterministic: candidates are visited in left order and ties keep the earliest.
    /// </summary>
    private static List<(int A, int B)> SplitAnchors(
        string[] leftKeys, int ls, int le, string[] rightKeys, int rs, int re)
    {
        var leftSeen = new Dictionary<string, (int Count, int First)>(StringComparer.Ordinal);
        for (int i = ls; i < le; i++)
            leftSeen[leftKeys[i]] = leftSeen.TryGetValue(leftKeys[i], out var e) ? (e.Count + 1, e.First) : (1, i);
        var rightSeen = new Dictionary<string, (int Count, int First)>(StringComparer.Ordinal);
        for (int j = rs; j < re; j++)
            rightSeen[rightKeys[j]] = rightSeen.TryGetValue(rightKeys[j], out var e) ? (e.Count + 1, e.First) : (1, j);

        var unique = new List<(int A, int B)>();
        (int A, int B) rarest = (-1, -1);
        int rarestCount = int.MaxValue;
        for (int i = ls; i < le; i++)
        {
            var (leftCount, leftFirst) = leftSeen[leftKeys[i]];
            if (leftFirst != i || !rightSeen.TryGetValue(leftKeys[i], out var right))
                continue;
            if (leftCount == 1 && right.Count == 1)
                unique.Add((i, right.First));
            else if (leftCount <= MaxSplitKeyOccurrences && right.Count <= MaxSplitKeyOccurrences &&
                     leftCount + right.Count < rarestCount)
            {
                rarest = (i, right.First);
                rarestCount = leftCount + right.Count;
            }
        }

        if (unique.Count > 0)
            return LongestIncreasingRun(unique);
        return rarest.A < 0 ? new List<(int A, int B)>() : new List<(int A, int B)> { rarest };
    }

    /// <summary>
    /// The longest subsequence of <paramref name="pairs"/> (ordered by A) whose B values strictly increase —
    /// patience sorting, O(k log k); ties keep the earliest pile tops, so the result is deterministic.
    /// </summary>
    private static List<(int A, int B)> LongestIncreasingRun(List<(int A, int B)> pairs)
    {
        var pileTops = new List<int>();          // index into pairs of each pile's top
        var previous = new int[pairs.Count];     // back-link to the top of the pile on the left
        for (int k = 0; k < pairs.Count; k++)
        {
            int lo = 0, hi = pileTops.Count;
            while (lo < hi)
            {
                int mid = (lo + hi) / 2;
                if (pairs[pileTops[mid]].B < pairs[k].B)
                    lo = mid + 1;
                else
                    hi = mid;
            }
            previous[k] = lo > 0 ? pileTops[lo - 1] : -1;
            if (lo == pileTops.Count)
                pileTops.Add(k);
            else
                pileTops[lo] = k;
        }

        var run = new List<(int A, int B)>(pileTops.Count);
        for (int k = pileTops[^1]; k >= 0; k = previous[k])
            run.Add(pairs[k]);
        run.Reverse();
        return run;
    }

    /// <summary>
    /// Common-subsequence match over <c>left[ls..le) × right[rs..re)</c> that maximizes total matched
    /// CHARACTER length (each match contributes <paramref name="weights"/>[a]) rather than token COUNT — a
    /// hypothesis for Word's anchor tie-break: among equal-length subsequences Word keeps the one covering
    /// more characters (a distinctive "strikethrough"/13 over an incidental "text"/4; a contiguous phrase
    /// over a scattered pair). O(n·m) DP with a deterministic prefer-left back-walk; falls back to token
    /// count when all weights are 1. Appends absolute index pairs to <paramref name="matches"/>. Callers keep
    /// the region within <see cref="LcsCellCap"/>; <paramref name="table"/> is grown as needed and reused.
    /// </summary>
    private static void CharWeightedLcs(
        string[] leftKeys, int ls, int le, string[] rightKeys, int rs, int re, int[] weights,
        List<(int A, int B)> matches, ref int[]? table)
    {
        int n = le - ls, m = re - rs;
        if (n == 0 || m == 0)
            return;

        // Row-major (n+1)×(m+1) table, reused across the regions of one alignment. Only the last row and
        // column are read before being written, so they are the only cells cleared.
        int stride = m + 1;
        int cells = (n + 1) * stride;
        if (table == null || table.Length < cells)
            table = new int[cells];
        var dp = table;
        Array.Clear(dp, n * stride, stride);
        for (int i = 0; i < n; i++)
            dp[i * stride + m] = 0;

        for (int i = n - 1; i >= 0; i--)
        {
            string key = leftKeys[ls + i];
            int weight = Math.Max(1, weights[ls + i]);
            int row = i * stride, below = row + stride;
            for (int j = m - 1; j >= 0; j--)
                dp[row + j] = key == rightKeys[rs + j]
                    ? dp[below + j + 1] + weight
                    : Math.Max(dp[below + j], dp[row + j + 1]);
        }
        for (int i = 0, j = 0; i < n && j < m;)
        {
            int row = i * stride, below = row + stride;
            if (leftKeys[ls + i] == rightKeys[rs + j] && dp[row + j] == dp[below + j + 1] + Math.Max(1, weights[ls + i]))
            {
                matches.Add((ls + i, rs + j)); i++; j++;
            }
            else if (dp[below + j] >= dp[row + j + 1]) i++;
            else j++;
        }
    }

    /// <summary>
    /// Emit the forward per-token edits for one anchor-free segment <c>left[ls..le) × right[rs..re)</c>
    /// (no content-anchor matches inside by construction): retain a common WHITESPACE prefix and suffix
    /// as Equal, and emit the middle as Delete(all remaining left) then Insert(all remaining right).
    /// A nested all-token LCS is deliberately NOT run here — it would reintroduce whitespace crowding.
    /// </summary>
    private static void EmitSegment(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right,
        int ls, int le, int rs, int re,
        List<(IrTokenOpKind Kind, int Left, int Right)> edits)
    {
        int leftLen = le - ls;
        int rightLen = re - rs;

        // Common whitespace prefix: grow while both sides share a connective token.
        int p = 0;
        while (p < leftLen && p < rightLen &&
               left[ls + p].MatchKey == right[rs + p].MatchKey &&
               IsConnective(left[ls + p]))
            p++;

        // Common whitespace suffix from the ends, not overlapping the prefix on either side.
        int sfx = 0;
        while (p + sfx < leftLen && p + sfx < rightLen &&
               left[le - 1 - sfx].MatchKey == right[re - 1 - sfx].MatchKey &&
               IsConnective(left[le - 1 - sfx]))
            sfx++;

        // Equal prefix.
        for (int t = 0; t < p; t++)
            edits.Add((IrTokenOpKind.Equal, ls + t, rs + t));

        // Delete middle-left.
        for (int t = ls + p; t < le - sfx; t++)
            edits.Add((IrTokenOpKind.Delete, t, -1));

        // Insert middle-right.
        for (int t = rs + p; t < re - sfx; t++)
            edits.Add((IrTokenOpKind.Insert, -1, t));

        // Equal suffix.
        for (int t = 0; t < sfx; t++)
            edits.Add((IrTokenOpKind.Equal, le - sfx + t, re - sfx + t));
    }

    /// <summary>
    /// Coalesce a forward-ordered per-token edit stream into maximal same-kind
    /// <see cref="IrTokenOp"/> spans. Insert spans get an empty left span at the running left cursor;
    /// Delete spans an empty right span at the running right cursor.
    /// </summary>
    private static void Coalesce(
        List<(IrTokenOpKind Kind, int Left, int Right)> edits, List<IrTokenOp> spans)
    {
        int i = 0;
        int leftCursor = 0, rightCursor = 0;
        while (i < edits.Count)
        {
            var kind = edits[i].Kind;
            int j = i;
            while (j < edits.Count && edits[j].Kind == kind)
                j++;
            int len = j - i;

            switch (kind)
            {
                case IrTokenOpKind.Equal:
                    spans.Add(new IrTokenOp(IrTokenOpKind.Equal,
                        leftCursor, leftCursor + len, rightCursor, rightCursor + len));
                    leftCursor += len;
                    rightCursor += len;
                    break;
                case IrTokenOpKind.Delete:
                    spans.Add(new IrTokenOp(IrTokenOpKind.Delete,
                        leftCursor, leftCursor + len, rightCursor, rightCursor));
                    leftCursor += len;
                    break;
                case IrTokenOpKind.Insert:
                    spans.Add(new IrTokenOp(IrTokenOpKind.Insert,
                        leftCursor, leftCursor, rightCursor, rightCursor + len));
                    rightCursor += len;
                    break;
            }

            i = j;
        }
    }

    // ------------------------------------------------------------------ format post-pass

    /// <summary>
    /// Split one content-equal span into alternating <see cref="IrTokenOpKind.Equal"/> and
    /// <see cref="IrTokenOpKind.FormatChanged"/> sub-spans by per-token <see cref="IrRunFormat"/>
    /// record equality. A position whose left/right Format records differ is FormatChanged; consecutive
    /// such positions merge into one FormatChanged span; equal-format positions stay Equal. This makes
    /// every position inside an emitted FormatChanged span pairwise format-UNEQUAL by construction.
    /// </summary>
    private static void SplitEqualByFormat(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right,
        IrTokenOp span, List<IrTokenOp> ops, IrFormatComparison comparison)
    {
        int len = span.LeftLength;
        int i = 0;
        while (i < len)
        {
            bool changed = FormatDiffers(left[span.LeftStart + i].Format, right[span.RightStart + i].Format, comparison);
            int j = i + 1;
            while (j < len &&
                   FormatDiffers(left[span.LeftStart + j].Format, right[span.RightStart + j].Format, comparison) == changed)
                j++;

            ops.Add(new IrTokenOp(
                changed ? IrTokenOpKind.FormatChanged : IrTokenOpKind.Equal,
                span.LeftStart + i, span.LeftStart + j,
                span.RightStart + i, span.RightStart + j));

            i = j;
        }
    }

    /// <summary>
    /// Per-token format comparison under the <paramref name="comparison"/> policy: modeled-only field
    /// equality (default) or full record equality (byte-fidelity). Two nulls are equal (non-run kinds —
    /// tab/break/etc. — carry null and never trip a format change); a null vs non-null pair differs only
    /// when the non-null side carries some modeled formatting (under ModeledOnly, a run whose rPr is
    /// entirely unmodeled keys equal to null).
    /// </summary>
    private static bool FormatDiffers(IrRunFormat? a, IrRunFormat? b, IrFormatComparison comparison) =>
        !IrModeledFormat.RunFormatEqual(a, b, comparison);
}

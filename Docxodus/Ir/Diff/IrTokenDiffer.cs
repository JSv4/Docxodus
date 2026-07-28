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
/// <para><b>Algorithm choice — Myers' O(ND) greedy diff (forward, Eugene W. Myers, "An O(ND)
/// Difference Algorithm and Its Variations", Algorithmica 1986, §2).</b> We use the simple forward
/// greedy variant (no linear-space middle-snake refinement). Inputs here are word-grain tokens of a
/// single paragraph pair — tens to a few hundred tokens each — so D (the edit distance) and N (the
/// summed length) are small; the O(ND) time and the O((N+M)·D)-bounded V-trace memory are negligible
/// at this scale, and the greedy form is the clearest correct implementation. We deliberately do NOT
/// use the LCS DP table (O(N·M) memory) — Myers is strictly better here and avoids the quadratic
/// allocation that a 200×200 token table would impose per Modified pair across a large corpus.</para>
/// <para><b>Determinism.</b> Myers' greedy walk is deterministic for fixed inputs; the standard
/// tie-break (prefer moving "down"/insert before "right"/delete when the furthest-reaching D-paths
/// tie, i.e. <c>k == -d || (k != d &amp;&amp; V[k-1] &lt; V[k+1])</c>) fixes the trace, so the op
/// sequence is a pure function of the two token lists. The format post-pass is a deterministic linear
/// scan. Two <see cref="Diff"/> calls on the same inputs return record-equal results.</para>
/// <para><b>Coalescing.</b> The raw Myers backtrace yields per-token edits; adjacent edits of the
/// same kind (Equal/Insert/Delete) are coalesced into maximal spans before the format post-pass.</para>
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
        // 1. Raw token-grain edits via Myers, already coalesced into same-kind spans. (MatchKey/Format
        // were precomputed by the tokenizer under these settings; the Myers walk keys on MatchKey.)
        var spans = MyersSpans(left, right, endsAtRetainedMark, suppressLonePunctuation);

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
    /// A single all-token Myers keyed on <see cref="IrDiffToken.MatchKey"/> mis-anchors on whitespace:
    /// every separator shares the key <c>" "</c>, so with many identical spaces Myers spends its LCS
    /// budget matching spaces and DROPS interior shared CONTENT words (delete+re-insert them). Word
    /// anchors on content words, not whitespace. We do the same in two levels:
    /// <list type="number">
    /// <item>Run Myers' LCS over the subsequence of NON-connective (content) tokens only — a token is
    /// connective iff it is a whitespace-only <see cref="IrDiffTokenKind.Separator"/>; Words, punctuation
    /// separators, and the atomic kinds all count as content and CAN anchor. This yields ordered Equal
    /// content-anchor pairs mapped back to full-stream indices.</item>
    /// <item>Partition both full streams at the anchors and emit, per segment, a common WHITESPACE prefix
    /// and suffix as Equal with the middle as Delete(all left)+Insert(all right) — no nested all-token
    /// Myers (that would reintroduce the whitespace crowding).</item>
    /// </list>
    /// The forward per-token edit stream feeds the shared <see cref="Coalesce"/> so anchors merge with
    /// adjacent whitespace Equal and consecutive Delete/Insert merge into maximal spans.
    /// </remarks>
    private static List<IrTokenOp> MyersSpans(
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
    /// These are the tokens that MUST NOT anchor the diff (else Myers crowds on abundant spaces).</summary>
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
    /// Compute the ordered content-anchor pairs: Myers' LCS over the non-connective (content) token
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

        var pairs = CharWeightedLcs(
            leftContent.Count, rightContent.Count,
            (a, b) => left[leftContent[a]].MatchKey == right[rightContent[b]].MatchKey,
            a => left[leftContent[a]].Text.Length);

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
    /// Common-subsequence match that maximizes total matched CHARACTER length (each match contributes
    /// <paramref name="weight"/>(a)) rather than token COUNT — a hypothesis for Word's anchor tie-break:
    /// among equal-length subsequences Word keeps the one covering more characters (a distinctive
    /// "strikethrough"/13 over an incidental "text"/4; a contiguous phrase over a scattered pair). O(n·m)
    /// DP with a deterministic prefer-left back-walk; falls back to token count when all weights are 1.
    /// </summary>
    private static List<(int A, int B)> CharWeightedLcs(int n, int m, Func<int, int, bool> eq, Func<int, int> weight)
    {
        var matches = new List<(int, int)>();
        if (n == 0 || m == 0)
            return matches;
        var dp = new int[n + 1, m + 1];
        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
                dp[i, j] = eq(i, j)
                    ? dp[i + 1, j + 1] + Math.Max(1, weight(i))
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);
        for (int i = 0, j = 0; i < n && j < m;)
        {
            if (eq(i, j) && dp[i, j] == dp[i + 1, j + 1] + Math.Max(1, weight(i)))
            {
                matches.Add((i, j)); i++; j++;
            }
            else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
            else j++;
        }
        return matches;
    }

    /// <summary>
    /// Forward greedy Myers O(ND) diff (Myers §2/§4) over two length-only sequences with a supplied
    /// equality predicate, returning the LCS as ordered matched index pairs <c>(a, b)</c> (a in
    /// <c>[0,n)</c>, b in <c>[0,m)</c>, both strictly increasing). Deterministic for fixed inputs via
    /// the standard prefer-down tie-break. Ignores non-matching positions (this level needs only the
    /// anchor matches; the segment pass handles the rest).
    /// </summary>
    private static List<(int A, int B)> MyersMatches(int n, int m, Func<int, int, bool> eq)
    {
        var matches = new List<(int, int)>();
        if (n == 0 || m == 0)
            return matches;

        int max = n + m;
        // V is indexed by diagonal k in [-max, max]; offset by `max`. trace[d] snapshots V before the
        // d-th round so we can backtrace the actual edit path (Myers §4 "recording the trace").
        int offset = max;
        var v = new int[2 * max + 1];
        var trace = new List<int[]>();

        bool reached = false;
        for (int d = 0; d <= max && !reached; d++)
        {
            trace.Add((int[])v.Clone());

            for (int k = -d; k <= d; k += 2)
            {
                // Prefer down (insert): k == -d, or (k != d and the up neighbour reaches further). This
                // fixes the path deterministically.
                int x;
                if (k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1]))
                    x = v[offset + k + 1];          // down: consume a right item (x unchanged)
                else
                    x = v[offset + k - 1] + 1;      // right: consume a left item (x advances)

                int y = x - k;

                // Follow the snake (matching diagonal) as far as the predicate agrees.
                while (x < n && y < m && eq(x, y))
                {
                    x++;
                    y++;
                }

                v[offset + k] = x;

                if (x >= n && y >= m)
                {
                    reached = true;
                    break;
                }
            }
        }

        // Backtrace from (n,m) to (0,0), collecting the diagonal (match) steps in reverse.
        int curX = n, curY = m;
        for (int d = trace.Count - 1; d > 0; d--)
        {
            var vv = trace[d];
            int k = curX - curY;

            int prevK;
            if (k == -d || (k != d && vv[offset + k - 1] < vv[offset + k + 1]))
                prevK = k + 1; // came from down (insert)
            else
                prevK = k - 1; // came from right (delete)

            int prevX = vv[offset + prevK];
            int prevY = prevX - prevK;

            while (curX > prevX && curY > prevY)
            {
                curX--;
                curY--;
                matches.Add((curX, curY));
            }

            if (curX == prevX)
                curY--; // insert (unmatched right)
            else
                curX--; // delete (unmatched left)
        }

        // Any remaining snake down to (0,0) at d == 0 is all matches.
        while (curX > 0 && curY > 0)
        {
            curX--;
            curY--;
            matches.Add((curX, curY));
        }

        matches.Reverse();
        return matches;
    }

    /// <summary>
    /// Emit the forward per-token edits for one anchor-free segment <c>left[ls..le) × right[rs..re)</c>
    /// (no content-anchor matches inside by construction): retain a common WHITESPACE prefix and suffix
    /// as Equal, and emit the middle as Delete(all remaining left) then Insert(all remaining right).
    /// A nested all-token Myers is deliberately NOT run here — it would reintroduce whitespace crowding.
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

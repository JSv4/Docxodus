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
    /// producing format-refined token ops.
    /// </summary>
    public static IrTokenDiff Diff(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right, IrDiffSettings settings)
    {
        // 1. Raw token-grain edits via Myers, already coalesced into same-kind spans. (MatchKey/Format
        // were precomputed by the tokenizer under these settings; the Myers walk keys on MatchKey.)
        var spans = MyersSpans(left, right);

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

    // ------------------------------------------------------------------ Myers O(ND)

    /// <summary>
    /// Forward greedy Myers O(ND) diff over MatchKeys, returning coalesced same-kind
    /// <see cref="IrTokenOp"/> spans (Equal/Insert/Delete only — the format pass runs later). Insert
    /// spans carry an empty left span at the anchor index; Delete spans an empty right span.
    /// </summary>
    private static List<IrTokenOp> MyersSpans(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right)
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

        int max = n + m;
        // V is indexed by diagonal k in [-max, max]; offset by `max`. trace[d] snapshots V after the
        // d-th round so we can backtrace the actual edit path (Myers §4 "recording the trace").
        int offset = max;
        var v = new int[2 * max + 1];
        var trace = new List<int[]>();

        bool reached = false;
        for (int d = 0; d <= max && !reached; d++)
        {
            var snapshot = (int[])v.Clone();
            trace.Add(snapshot);

            for (int k = -d; k <= d; k += 2)
            {
                // Choose to come from the diagonal above (insert/down) or left (delete/right).
                // Tie-break: prefer down (insert) — k == -d, or (k != d and the up neighbour reaches
                // further). This fixes the path deterministically.
                int x;
                if (k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1]))
                    x = v[offset + k + 1];          // down: insert a right token (x unchanged)
                else
                    x = v[offset + k - 1] + 1;      // right: delete a left token (x advances)

                int y = x - k;

                // Follow the snake (matching diagonal) as far as MatchKeys agree.
                while (x < n && y < m && left[x].MatchKey == right[y].MatchKey)
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

        Backtrace(left, right, trace, offset, n, m, spans);
        return spans;
    }

    /// <summary>
    /// Walk the recorded V-snapshots back from (n,m) to (0,0), emitting per-token edits in reverse,
    /// then reverse and coalesce them into maximal same-kind spans.
    /// </summary>
    private static void Backtrace(
        IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right,
        List<int[]> trace, int offset, int n, int m, List<IrTokenOp> spans)
    {
        // Reverse-order per-token edits as (kind, leftIndex, rightIndex). For Equal we record the
        // matched (x-1, y-1); for Delete the consumed left index (x-1); for Insert the right index (y-1).
        var rev = new List<(IrTokenOpKind Kind, int Left, int Right)>();

        int curX = n, curY = m;
        for (int d = trace.Count - 1; d > 0; d--)
        {
            var v = trace[d];
            int k = curX - curY;

            int prevK;
            if (k == -d || (k != d && v[offset + k - 1] < v[offset + k + 1]))
                prevK = k + 1; // came from down (insert)
            else
                prevK = k - 1; // came from right (delete)

            int prevX = v[offset + prevK];
            int prevY = prevX - prevK;

            // Snake (diagonal) steps before the edit are Equal matches.
            while (curX > prevX && curY > prevY)
            {
                curX--;
                curY--;
                rev.Add((IrTokenOpKind.Equal, curX, curY));
            }

            // The single non-diagonal edit at this D-step (d > 0 throughout this loop, so an edit always exists).
            if (curX == prevX)
            {
                // Insert: right token at prevY..curY-1 (one token, curY-1).
                curY--;
                rev.Add((IrTokenOpKind.Insert, -1, curY));
            }
            else
            {
                // Delete: left token curX-1.
                curX--;
                rev.Add((IrTokenOpKind.Delete, curX, -1));
            }
        }

        // Any remaining snake down to (0,0) at d == 0 is Equal.
        while (curX > 0 && curY > 0)
        {
            curX--;
            curY--;
            rev.Add((IrTokenOpKind.Equal, curX, curY));
        }

        // Coalesce the reversed edits (which are right-to-left) into forward maximal spans.
        rev.Reverse();
        Coalesce(rev, spans);
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

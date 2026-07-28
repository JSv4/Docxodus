#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// M2.6 split/merge segmentation: scores a (singular paragraph, candidate run of paragraphs) pair via an
/// in-order LCS over token MatchKeys (coverage + foreign slack, spec §2.2), and computes the
/// per-segment token diffs by slicing the singular side's token stream at the LCS assignment boundaries
/// and re-running <see cref="IrTokenDiffer.Diff"/> per slice — which guarantees every segment diff carries
/// the full IrTokenDiff invariants over (slice, member block), making slice boundaries IMPLICIT in the
/// diff ops (review F3.3: slice i's left length = Σ non-Insert left-span lengths; the slices tile the
/// singular token stream exactly, in order).
/// </summary>
/// <remarks>
/// <para><b>Determinism.</b> Standard O(n·m) LCS DP with a fixed back-walk tie-break (prefer the
/// singular-side advance on equal subproblem values — the same discipline as
/// <c>IrEditScriptBuilder.LongestCommonSubsequence</c>). Unmatched singular-side tokens attach to the
/// segment of the nearest PRECEDING matched token (leading unmatched → segment 0) — a total, documented
/// rule. No dictionary enumeration feeds output.</para>
/// <para><b>Content tokens.</b> Coverage and slack count only non-Separator, non-Textbox tokens
/// (separators are connective; a masked textbox placeholder is not content) — the same rule
/// <c>IrRevisionRenderer.CountContent</c> applies. The LCS itself runs over ALL tokens so boundary
/// assignment has separator context, but only content-token matches score.</para>
/// <para><b>Cost.</b> O(|singular|·|run|) DP per call, gap-bounded and capped by
/// <see cref="IrDiffSettings.SplitMaxRunLength"/> at the detection site. NOTE:
/// <see cref="Score"/> and <see cref="ComputeSegmentDiffs"/> each re-tokenize and re-run the LCS
/// independently, so a caller that scores k candidate runs and then computes segment diffs for the
/// winner pays k+1 DP passes over the singular stream — at the ≤8-member cap that is at most 9
/// passes per gap, which stays inside the aligner's documented G²-class budget but should be
/// factored into any future threshold-sweep cost analysis.</para>
/// </remarks>
internal static class IrSplitSegmenter
{
    /// <summary>One candidate's score: in-order coverage of the singular side's content tokens, the
    /// run's foreign-content fraction, the per-member matched-content counts (used by detection to
    /// trim net-new edge members — the R2 false-positive guard), and per-member matched Word-token
    /// counts (used to require two meaningful phrase fragments rather than isolated token moves).</summary>
    internal sealed record SplitScore(
        double Coverage,
        double ForeignSlack,
        IReadOnlyList<int> MemberMatchedContent,
        IReadOnlyList<int> MemberMatchedWords);

    /// <summary>Score one candidate: the singular paragraph vs the concatenated run, per spec §2.2.</summary>
    public static SplitScore Score(IrParagraph singular, IReadOnlyList<IrParagraph> run, IrDiffSettings settings)
    {
        var single = IrDiffTokenizer.Tokenize(singular, settings);
        var memberTokens = run.Select(p => IrDiffTokenizer.Tokenize(p, settings)).ToList();
        var flat = new List<IrDiffToken>();
        var memberOfFlat = new List<int>();
        for (int m = 0; m < memberTokens.Count; m++)
            foreach (var t in memberTokens[m])
            {
                flat.Add(t);
                memberOfFlat.Add(m);
            }

        var partner = LcsMatch(single, flat, out var flatMatched);

        int singleContent = CountContent(single);
        int matchedContent = 0;
        var memberMatched = new int[run.Count];
        var memberMatchedWords = new int[run.Count];
        for (int i = 0; i < single.Count; i++)
        {
            if (partner[i] < 0 || !IsContent(single[i]))
                continue;
            matchedContent++;
            memberMatched[memberOfFlat[partner[i]]]++;
            if (single[i].Kind == IrDiffTokenKind.Word)
                memberMatchedWords[memberOfFlat[partner[i]]]++;
        }

        int runContent = CountContent(flat);
        int runMatchedContent = 0;
        for (int j = 0; j < flat.Count; j++)
            if (flatMatched[j] && IsContent(flat[j]))
                runMatchedContent++;

        double coverage = singleContent == 0 ? 0.0 : (double)matchedContent / singleContent;
        double slack = runContent == 0 ? 0.0 : (double)(runContent - runMatchedContent) / runContent;
        return new SplitScore(coverage, slack, memberMatched, memberMatchedWords);
    }

    /// <summary>
    /// Per-segment diffs for a confirmed split (singular = LEFT, <paramref name="singularIsLeft"/>
    /// true) — or, with the arguments mirrored by the caller, a merge (singular = RIGHT,
    /// <paramref name="singularIsLeft"/> false; the caller then flips each diff via
    /// <see cref="MirrorDiff"/>). Returns one COMPLETE IrTokenDiff per run member: slice-local
    /// singular-side spans, full member-side spans.
    /// </summary>
    public static IrNodeList<IrTokenDiff> ComputeSegmentDiffs(
        IrParagraph singular, IReadOnlyList<IrParagraph> run, IrDiffSettings settings,
        bool singularIsLeft = true)
    {
        var single = IrDiffTokenizer.Tokenize(singular, settings);
        var memberTokens = run.Select(p => IrDiffTokenizer.Tokenize(p, settings)).ToList();
        var flat = new List<IrDiffToken>();
        var memberOfFlat = new List<int>();
        for (int m = 0; m < memberTokens.Count; m++)
            foreach (var t in memberTokens[m])
            {
                flat.Add(t);
                memberOfFlat.Add(m);
            }

        var partner = LcsMatchWeighted(single, flat);

        // Assign every singular token to a member segment (decoded 2026-07-27 from reference compare
        // output — the same merged-stream expansion arrangement the cross-paragraph segmenter renders:
        // between two anchors the right side's one-sided items emit BEFORE the boundary event and the
        // left side's AFTER it). A segment is defined by SOLID matches only: matched content tokens,
        // plus matched separators CHAINED to a solid neighbor by partner adjacency (the cross-paragraph
        // flanking-separator law — "ddd. " keeps its period and trailing space with "ddd"). An
        // UNCHAINED separator match must not drag the boundary (it is connective; counting it pulled
        // "which" into the following member). Unmatched singular tokens — and unchained separator
        // matches — are the residue, and the residue's SIDE decides its member:
        //  - singular = RIGHT (merge): the residue renders as INSERTED text, which the expansion emits
        //    directly after the preceding anchor — nearest PRECEDING solid match's member (leading → 0);
        //  - singular = LEFT (split): the residue renders as DELETED text, which the expansion emits
        //    directly before the FOLLOWING anchor, past any ¶INS boundary — nearest FOLLOWING solid
        //    match's member (trailing → the LAST member, whose retained mark closes the stream).
        var solid = new bool[single.Count];
        for (int i = 0; i < single.Count; i++)
            solid[i] = partner[i] >= 0 && IsContent(single[i]);
        for (int i = 1; i < single.Count; i++)
            if (!solid[i] && partner[i] >= 0 && solid[i - 1] && partner[i] == partner[i - 1] + 1)
                solid[i] = true;
        for (int i = single.Count - 2; i >= 0; i--)
            if (!solid[i] && partner[i] >= 0 && solid[i + 1] && partner[i] == partner[i + 1] - 1)
                solid[i] = true;

        var segmentOf = new int[single.Count];
        if (singularIsLeft)
        {
            int next = run.Count - 1;
            for (int i = single.Count - 1; i >= 0; i--)
            {
                if (solid[i])
                    next = memberOfFlat[partner[i]];
                segmentOf[i] = next;
            }
        }
        else
        {
            // TRAILING inserted residue (beyond the LAST solid match) goes to the FINAL member. Every
            // reference-output example of trailing insert residue has its preceding solid match in the
            // final member already, so "preceding" and "final member" are indistinguishable on the
            // decoded evidence — and when the aligner's merge slot disagrees with the reference's (the
            // ¶DEL-slot family), the final-member branch keeps the trailing insert beside the surviving
            // mark, which measures strictly better. Interior residue follows the expansion law: nearest
            // PRECEDING solid member (leading → 0).
            int lastSolid = -1;
            for (int i = single.Count - 1; i >= 0 && lastSolid < 0; i--)
                if (solid[i])
                    lastSolid = i;
            int current = 0;
            for (int i = 0; i < single.Count; i++)
            {
                if (solid[i])
                    current = memberOfFlat[partner[i]];
                segmentOf[i] = lastSolid >= 0 && i > lastSolid ? run.Count - 1 : current;
            }
        }
        // The LCS is in-order, so segments are non-decreasing; enforce defensively anyway.
        for (int i = 1; i < single.Count; i++)
            if (segmentOf[i] < segmentOf[i - 1])
                segmentOf[i] = segmentOf[i - 1];

        // Separator attachment at segment boundaries (decoded 2026-07-27 from reference compare output):
        // the whitespace separating an UNMATCHED word from a following cross-boundary matched word rides
        // with the MATCH — Word retains " and " (leading space included) in the later output paragraph,
        // while our nearest-preceding rule dumped that separator into the earlier slice's ins/del run
        // ("…outcomes ", " and" split). Move the maximal run of connective singular tokens directly
        // before a segment-opening matched token into that token's segment, but only while the member
        // side has pairwise key-equal connectives directly before the partner (re-pairable — the slice
        // re-diff below then retains them via its whitespace prefix/suffix rule) and the singular token
        // preceding the run is NOT a matched content token (trailing attachment to a matched word wins
        // over leading attachment — the decoded precedence).
        for (int i = 1; i < single.Count; i++)
        {
            if (segmentOf[i] == segmentOf[i - 1] || partner[i] < 0)
                continue;
            int seg = segmentOf[i];
            if (memberOfFlat[partner[i]] != seg)
                continue;                                // defensively bumped segment — leave untouched
            int j = i;                                   // tokens [j..i) move into segment `seg`
            int k = partner[i];                          // flat counterpart cursor (exclusive)
            while (j - 1 >= 0 && segmentOf[j - 1] < seg &&
                   IsConnective(single[j - 1]) &&
                   k - 1 >= 0 && memberOfFlat[k - 1] == memberOfFlat[partner[i]] &&
                   IsConnective(flat[k - 1]) &&
                   single[j - 1].MatchKey == flat[k - 1].MatchKey)
            {
                j--;
                k--;
            }
            if (j == i)
                continue;
            // Precedence guard: leave the run in place when it trails a matched content token.
            if (j - 1 >= 0 && partner[j - 1] >= 0 && IsContent(single[j - 1]))
                continue;
            for (int t = j; t < i; t++)
                segmentOf[t] = seg;
        }

        // Lone-punctuation suppression precondition: every source paragraph is plain direct text
        // (the split/merge lowering already excludes field carriers, but hyperlinks/atomics pass it).
        bool plain = IrCrossParagraphSegmenter.IsStreamable(singular);
        foreach (var member in run)
            plain &= IrCrossParagraphSegmenter.IsStreamable(member);

        var diffs = new List<IrTokenDiff>(run.Count);
        int cursor = 0;
        for (int m = 0; m < run.Count; m++)
        {
            int start = cursor;
            while (cursor < single.Count && segmentOf[cursor] == m)
                cursor++;
            diffs.Add(IrTokenDiffer.Diff(Slice(single, start, cursor), memberTokens[m], settings,
                endsAtRetainedMark: m == run.Count - 1, suppressLonePunctuation: plain));
        }
        Debug.Assert(cursor == single.Count,
            "segment assignment must consume the whole singular token stream");
        return IrNodeList.From(diffs);
    }

    /// <summary>Swap a token diff's left/right sides (Insert↔Delete, spans mirrored) — turns a
    /// split-shaped (slice vs member) diff into the merge-shaped (member vs slice) orientation.</summary>
    public static IrTokenDiff MirrorDiff(IrTokenDiff diff)
    {
        var ops = diff.Ops.Select(o => new IrTokenOp(
            o.Kind switch
            {
                IrTokenOpKind.Insert => IrTokenOpKind.Delete,
                IrTokenOpKind.Delete => IrTokenOpKind.Insert,
                var k => k,
            },
            o.RightStart, o.RightEnd, o.LeftStart, o.LeftEnd)).ToList();
        return new IrTokenDiff(IrNodeList.From(ops));
    }

    // ------------------------------------------------------------------ internals

    private static IReadOnlyList<IrDiffToken> Slice(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        var list = new List<IrDiffToken>(end - start);
        for (int i = start; i < end; i++)
            list.Add(tokens[i]);
        return list;
    }

    private static bool IsContent(IrDiffToken t) =>
        t.Kind is not (IrDiffTokenKind.Separator or IrDiffTokenKind.Textbox);

    /// <summary>Connective whitespace — the tokens that re-pair across a moved segment boundary
    /// (same definition as <c>IrTokenDiffer.IsConnective</c>).</summary>
    private static bool IsConnective(IrDiffToken t) =>
        t.Kind == IrDiffTokenKind.Separator && string.IsNullOrWhiteSpace(t.Text);

    private static int CountContent(IReadOnlyList<IrDiffToken> tokens)
    {
        int n = 0;
        foreach (var t in tokens)
            if (IsContent(t))
                n++;
        return n;
    }

    /// <summary>
    /// CHAR-WEIGHTED LCS over MatchKeys for the SEGMENTATION pass (decoded 2026-07-27): a word match
    /// weighs its visible characters and a separator match 1, the same matched-character objective as
    /// the ordinary differ and the cross-paragraph segmenter. The token-count LCS (<see cref="LcsMatch"/>,
    /// kept verbatim for <see cref="Score"/> so DETECTION thresholds are untouched) ties between "match
    /// the crossing word" and "match one more separator", and its singular-side-advance tie-break
    /// dropped the word — the reference output matches the word ("for" retained in the following member
    /// while our slice inserted "for document" wholesale). Front-walk prefers the match on equal dp.
    /// </summary>
    private static int[] LcsMatchWeighted(IReadOnlyList<IrDiffToken> a, IReadOnlyList<IrDiffToken> b)
    {
        int n = a.Count, m = b.Count;
        var partner = new int[n];
        Array.Fill(partner, -1);
        if (n == 0 || m == 0)
            return partner;

        int Weight(int i, int j) =>
            a[i].MatchKey != b[j].MatchKey ? 0
            : a[i].Kind == IrDiffTokenKind.Word ? Math.Max(1, a[i].Text.Length)
            : 1;

        var dp = new int[n + 1, m + 1];
        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
            {
                int w = Weight(i, j);
                int best = Math.Max(dp[i + 1, j], dp[i, j + 1]);
                if (w > 0)
                    best = Math.Max(best, dp[i + 1, j + 1] + w);
                dp[i, j] = best;
            }
        for (int i = 0, j = 0; i < n && j < m;)
        {
            int w = Weight(i, j);
            if (w > 0 && dp[i, j] == dp[i + 1, j + 1] + w)
            {
                partner[i] = j;
                i++;
                j++;
            }
            else if (dp[i + 1, j] >= dp[i, j + 1])
            {
                i++;
            }
            else
            {
                j++;
            }
        }
        return partner;
    }

    /// <summary>Standard LCS DP over MatchKeys. Returns the per-singular-index partner flat index
    /// (or -1); <paramref name="flatMatched"/> marks consumed flat indices. Same DP fill + back-walk
    /// tie-break as <see cref="IrEditScriptBuilder"/>'s private <c>LongestCommonSubsequence</c> —
    /// kept separate because the output shapes differ (partner array + matched bitmap here vs index
    /// pairs there); keep the tie-break rules in sync if either changes.</summary>
    private static int[] LcsMatch(
        IReadOnlyList<IrDiffToken> a, IReadOnlyList<IrDiffToken> b, out bool[] flatMatched)
    {
        int n = a.Count, m = b.Count;
        var dp = new int[n + 1, m + 1];
        for (int i = n - 1; i >= 0; i--)
            for (int j = m - 1; j >= 0; j--)
                dp[i, j] = a[i].MatchKey == b[j].MatchKey
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        var partner = new int[n];
        Array.Fill(partner, -1);
        flatMatched = new bool[m];
        for (int i = 0, j = 0; i < n && j < m;)
        {
            if (a[i].MatchKey == b[j].MatchKey) { partner[i] = j; flatMatched[j] = true; i++; j++; }
            else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
            else j++;
        }
        return partner;
    }
}

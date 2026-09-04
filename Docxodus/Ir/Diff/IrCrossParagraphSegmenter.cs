#nullable enable

using System;
using System.Collections.Generic;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Cross-paragraph token-stream segmenter (2026-07-25). Given a run of ≥2 ADJACENT word-matched
/// Modified paragraph pairs — K left paragraphs paired 1:1 with K right paragraphs by the aligner's
/// in-gap word matching — factors the run's combined token alignment back into output-paragraph
/// <see cref="IrCrossParagraphCell"/>s, reproducing the within-run flat word+pilcrow stream decoded
/// from Word's compare output: retained words from base paragraph k can land in output paragraph k±1
/// (crossing the pilcrow), pilcrows are ¶INS/¶DEL stream tokens by side, and the output paragraph
/// count follows the token-level interleave — a 2×2 word-matched region can emit 3 paragraphs. No
/// per-paragraph model reproduces this.
/// </summary>
/// <remarks>
/// <para><b>Algorithm (the two-pass decode).</b> Decoded case-by-case from Word's compare output over
/// word-matched runs. A single free-for-all LCS over the concatenated streams (pilcrows as weight-1
/// sentinel anchors) is provably NOT the observed behavior: measured against Word's own compare output
/// it BOTH misses the crossings (the sentinel pair outcompetes them, reproducing the per-pair result)
/// AND invents crossings the reference output never makes (stealing a pair's interior words into the
/// next paragraph when the pair's own trailing match seals them in). The decoded behavior:</para>
/// <para>(1) <b>Anchor units.</b> Content anchors are WORDS and COMPOUNDS (words joined by
/// non-whitespace separators with no interior space — "right-aligned" anchors as one unit, so a plain
/// "right" elsewhere cannot poach it). Whitespace AND punctuation separators are connective — they
/// never anchor; they ride with the cells and re-pair during the per-cell re-diff.</para>
/// <para>(2) <b>Pass 1 — per-pair anchors, char-weighted, never discarded.</b> Each pair (Lk, Rk) is
/// matched independently by a CHAR-WEIGHTED LCS over its units (the same matched-character objective
/// as the ordinary per-pair differ), and the matches become hard anchors — a residue trapped BETWEEN
/// two in-pair anchors can never cross a boundary. In-slot anchors always stand: Word matches even an
/// isolated single word within its own slot, so the cross-run pass sees ONLY words with no in-slot
/// counterpart. (An earlier density floor discarded sparse pairs' anchors here; that flooded the
/// cross-windows with words that had in-slot homes and made the fusion rearrange content Word keeps
/// per-slot — measured as two large corpus regressions. A later matched-char coverage floor failed the
/// same way: the retain-in-place and yield-to-cross classes overlap on coverage, so no threshold
/// separates them.) The exception is the STORY-FINAL pair of a story-ending run: it anchors only its
/// common unit prefix, and its other same-ordinal recoveries stand in pass 2 only as adjacent bigrams —
/// the decoded retained-in-place signatures — so isolated recoveries cannot seal the window against a
/// genuine crossing chain.</para>
/// <para>(3) <b>Pass 2 — residues flow across the run.</b> Between consecutive anchors the unmatched
/// units of both sides (which may span several paragraphs — pilcrows are not tokens at this level) are
/// re-matched by a CHAR-WEIGHTED LCS, with a compound additionally able to match a plain word equal to
/// its first/last member ("right-aligned" ↔ "aligned"; the cell re-diff then renders the shared member
/// and strikes the rest — the observed sub-word shape). New matches are the boundary-CROSSING retained
/// words: a pair's unmatched tail matching the next pair's right content, or a right tail matching a
/// following left paragraph's head (the backward flow).</para>
/// <para>(4) <b>Pilcrows pair structurally, never by content.</b> Boundary b (¶Lb, ¶Rb) keeps a shared
/// retained pilcrow iff no matched token lies between the two pilcrows in the merged stream
/// (equivalently: the same number of matched tokens precedes each on its own side). Otherwise the
/// boundary splits: ¶Rb renders as an INSERTED paragraph mark and ¶Lb as a DELETED one, each at its own
/// side's stream position. The run's FINAL boundary always pairs (both final pilcrows follow every
/// match).</para>
/// <para>(5) <b>Expansion.</b> Anchors (matched tokens + paired boundaries) are walked in stream order;
/// between two anchors the RIGHT side's one-sided items (inserted tokens, ¶INS boundaries) emit before
/// the LEFT side's (deleted tokens, ¶DEL boundaries) — the decoded new-content-first arrangement. Every
/// boundary event closes one output cell; each cell's left/right slices are then RE-diffed
/// slice-relative with the ordinary <see cref="IrTokenDiffer"/> so the shared markup renderer produces
/// ins-before-del + rPrChange for free, exactly as for a split/merge segment.</para>
/// <para><b>Safety.</b> By the pilcrow-boundary construction each cell references AT MOST one left and
/// one right paragraph; the walk asserts this and the within-paragraph slice contiguity, returning
/// <c>null</c> (→ the caller falls back to the ordinary per-pair path) if EITHER is ever violated, if
/// the last cell would not close on a retained boundary, or if a window's LCS would exceed the DP size
/// cap. The caller additionally restricts every run member to <see cref="IsStreamable"/> paragraphs
/// (plain text, no structural carriers).</para>
/// <para><b>Run + trailing gap (2026-07-27).</b> A word-matched run whose FOLLOWING blocks are a
/// story-final replace gap (or a story-tail split/merge surplus) extends the same stream with ONE-SIDED
/// trailing members: deleted paragraphs join the left member list, inserted ones the right, after the
/// paired prefix. Decoded from Word's compare output ("justified body → large font"): the last pair's
/// long left tail crosses FORWARD into the paragraph carrying the trailing insert, sharing retained
/// words with it, and each side's story-final pilcrow is the same immutable stream token — so the two
/// final pilcrows pair structurally and the tail-fusion construct moves from the last PAIR to the gap.
/// One-sided↔one-sided content never matches at the stream level (interior gap paragraphs stay fully
/// marked, per the replace-gap grammar); with zero strictly-crossing matches the whole construct falls
/// back (the ordinary replace-gap / split / merge grammar already renders that shape).</para>
/// <para><b>Round-trip.</b> ACCEPT keeps Ins pilcrow marks + drops Del marks (fusing each Del cell into
/// the next paragraph) + keeps right slices + drops left → the right member paragraphs; REJECT the
/// inverse → the left members. A single-sided cell fuses only its own (empty on the removing side)
/// content. Verified by <c>DocxDiffCrossParagraphTests</c> and the fusion-ON corpus/fuzz batteries.</para>
/// </remarks>
internal static class IrCrossParagraphSegmenter
{
    /// <summary>Cap on any single LCS DP table (units² per window); larger windows bail to the
    /// per-pair path rather than risk pathological cost.</summary>
    private const long LcsCellCap = 1_000_000;

    /// <summary>
    /// A paragraph is streamable iff it is plain, sliceable text with no structural carrier: every inline is
    /// an <see cref="IrTextRun"/> (so a slice boundary can only ever fall between plain text tokens, which the
    /// renderer's <c>SourceRunModel</c> splits cleanly) and it carries no inline section transition. Hyperlinks,
    /// fields, note refs, images, opaque inlines, textboxes, tabs and breaks are all EXCLUDED — they are
    /// zero-width or atomic and a mid-word / boundary slice could double-emit or drop them across cells, so the
    /// caller falls back to the ordinary per-pair path when any run paragraph is not streamable. (The caller
    /// ALSO excludes structural-envelope/field-carrier paragraphs via its own digest gate — an inline SDT's
    /// runs read as plain <see cref="IrTextRun"/>s here but its envelope cannot be sliced.)
    /// </summary>
    public static bool IsStreamable(IrParagraph p)
    {
        if (p.InlineSectionBreakAnchor is not null || p.InlineSectionFormat is not null)
            return false;
        foreach (var inline in p.Inlines)
            if (inline is not IrTextRun)
                return false;
        return true;
    }

    /// <summary>One content-anchor unit: a word, or a compound (words joined by non-whitespace
    /// separators). <see cref="Start"/>/<see cref="Length"/> address the side's FLAT token stream;
    /// <see cref="Key"/> is the concatenated member MatchKey; <see cref="Chars"/> the visible text
    /// length; <see cref="FirstWordKey"/>/<see cref="LastWordKey"/> support the partial
    /// compound↔word endpoint match (identical to <see cref="Key"/> for a plain word).</summary>
    private readonly record struct AnchorUnit(
        int Start, int Length, string Key, int Chars, string FirstWordKey, string LastWordKey);

    /// <summary>One chain anchor of the expansion: a matched token pair or a PAIRED boundary
    /// (<see cref="BoundaryLeft"/>/<see cref="BoundaryRight"/> are the per-side boundary ordinals —
    /// equal for a prefix pair, and (kl−1, kr−1) for the story-final structural pair).</summary>
    private readonly record struct Anchor(
        int LeftPos, int RightPos, bool IsBoundary, int BoundaryLeft, int BoundaryRight);

    private enum ItemKind { Paired, Ins, Del, BoundaryEqual, BoundaryIns, BoundaryDel }

    /// <summary>One merged-stream item. For token items the fields are FLAT token indices; for boundary
    /// items they carry the boundary ordinal (on the side(s) the boundary owns, −1 otherwise).</summary>
    private readonly record struct MergedItem(ItemKind Kind, int Left, int Right);

    /// <summary>
    /// Segment a word-matched run — optionally extended by a trailing STORY-FINAL one-sided tail — into
    /// output-paragraph cells, or return <c>null</c> to signal the caller must fall back to the ordinary
    /// per-pair path. The first <paramref name="pairedCount"/> members of each side are 1:1 pairs (the
    /// word-matched run); any members beyond that prefix are ONE-SIDED — trailing deleted paragraphs on
    /// the left, trailing inserted paragraphs on the right (a replace gap or split/merge surplus the
    /// builder absorbed; only valid when the combined range reaches the story end, so the final boundary
    /// of EACH side is the story's final pilcrow and pairs structurally). <paramref name="pairedCount"/>
    /// = −1 (the default) means "all paired" and requires equal counts. All members must be
    /// <see cref="IsStreamable"/> (the caller enforces this, plus its structural-carrier digest gate).
    /// </summary>
    public static List<IrCrossParagraphCell>? Segment(
        IReadOnlyList<IrParagraph> left, IReadOnlyList<IrParagraph> right, IrDiffSettings settings,
        bool runEndsStory = false, int pairedCount = -1)
    {
        int kl = left.Count, kr = right.Count;
        int p = pairedCount >= 0 ? pairedCount : (kl == kr ? kl : -1);
        if (p < 1 || p > kl || p > kr)
            return null;
        var prefixPairs = new List<(int L, int R)>(p);
        for (int i = 0; i < p; i++)
            prefixPairs.Add((i, i));
        return SegmentRegion(left, right, prefixPairs, settings, runEndsStory);
    }

    /// <summary>
    /// The general REGION form (2026-07-27): members are the region's paragraphs in document order per
    /// side, and <paramref name="pairs"/> lists the word-matched pairs as (left index, right index),
    /// strictly ascending on both coordinates; members not covered by a pair are ONE-SIDED (a deleted
    /// paragraph on the left, an inserted one on the right) at ANY position — leading, interior, or
    /// trailing. <see cref="Segment"/>'s prefix form maps to pairs [(0,0)..(p−1,p−1)] and keeps its
    /// exact behavior. New configurations decoded from the reference compare output: a ZERO-pair
    /// story-final region still streams when its one-sided paragraphs share words (function words
    /// included), and LEADING one-sided members join a following pair's stream (their words landing
    /// retained in later output paragraphs). Match bans: with pairs present the member-ordinal
    /// forward rule and the trailing↔trailing one-sided ban stand exactly as before; a zero-pair
    /// region's matches are constrained only by stream order.
    /// </summary>
    public static List<IrCrossParagraphCell>? SegmentRegion(
        IReadOnlyList<IrParagraph> left, IReadOnlyList<IrParagraph> right,
        IReadOnlyList<(int L, int R)> pairs, IrDiffSettings settings, bool runEndsStory)
    {
        int kl = left.Count, kr = right.Count;
        if (kl == 0 || kr == 0)
            return null;
        var leftPairIdx = new int[kl];
        var rightPairIdx = new int[kr];
        Array.Fill(leftPairIdx, -1);
        Array.Fill(rightPairIdx, -1);
        for (int i = 0; i < pairs.Count; i++)
        {
            var (li, ri) = pairs[i];
            if (li < 0 || li >= kl || ri < 0 || ri >= kr)
                return null;
            if (i > 0 && (li <= pairs[i - 1].L || ri <= pairs[i - 1].R))
                return null;
            leftPairIdx[li] = i;
            rightPairIdx[ri] = i;
        }
        // One-sided members (the run+gap construct, 2026-07-27, generalized to any position). With
        // any one-sided member the story-final pilcrows of the two sides pair structurally regardless
        // of member ordinals, and the decoded fusion construct lives on the one-sided matter, so
        // every PAIR keeps its in-slot anchors.
        bool hasTail = pairs.Count < kl || pairs.Count < kr;
        int lastPairL = pairs.Count > 0 ? pairs[^1].L : -1;
        int lastPairR = pairs.Count > 0 ? pairs[^1].R : -1;

        // Tokenize every paragraph; build flat coordinate maps. Pilcrows are NOT tokens — they are the
        // paragraph boundaries themselves, handled structurally below. Tokenizing under the SAME settings
        // the renderer uses keeps the slices it re-derives byte-identical to these.
        var leftReal = new List<IrDiffToken>[kl];
        var rightReal = new List<IrDiffToken>[kr];
        var offL = new int[kl + 1];
        var offR = new int[kr + 1];
        for (int i = 0; i < kl; i++)
        {
            leftReal[i] = new List<IrDiffToken>(IrDiffTokenizer.Tokenize(left[i], settings));
            offL[i + 1] = offL[i] + leftReal[i].Count;
        }
        for (int i = 0; i < kr; i++)
        {
            rightReal[i] = new List<IrDiffToken>(IrDiffTokenizer.Tokenize(right[i], settings));
            offR[i + 1] = offR[i] + rightReal[i].Count;
        }
        int totalL = offL[kl], totalR = offR[kr];
        var flatL = new IrDiffToken[totalL];
        var flatR = new IrDiffToken[totalR];
        for (int i = 0; i < kl; i++)
            leftReal[i].CopyTo(flatL, offL[i]);
        for (int i = 0; i < kr; i++)
            rightReal[i].CopyTo(flatR, offR[i]);

        // ---- Pass 1: per-pair anchor units, token-count LCS (earliest tie-break), density gate.
        var all = new List<(int Lf, int Rf)>(); // matched TOKEN pairs, stream order (both sides monotone)
        var pass1 = new List<(int Lf, int Rf)>();
        for (int pi = 0; pi < pairs.Count; pi++)
        {
            var (li, ri) = pairs[pi];
            // The STORY-FINAL pair is the tail-FUSION construct — decoded from Word's compare
            // output, its content is mostly replaced and full in-slot anchoring would both invent
            // retentions Word does not make there and seal the window the legitimate cross-flow
            // needs. It anchors ONLY its common unit PREFIX in place ("Highlighted" staying put
            // while the rest of the line is rewritten, decoded law refinement 2026-07-27);
            // everything beyond the prefix retains only via the pass-2 window — the bigram valve
            // there, or cross-flow from an earlier pair's residue ("is"/"centered" landing in the
            // final paragraph). With one-sided members the fusion construct is the one-sided
            // matter itself, so every pair — the last included — anchors in-slot fully.
            if (runEndsStory && !hasTail && pi == pairs.Count - 1)
            {
                var pUnitsL = BuildUnits(flatL, offL[li], offL[li + 1]);
                var pUnitsR = BuildUnits(flatR, offR[ri], offR[ri + 1]);
                int np = Math.Min(pUnitsL.Count, pUnitsR.Count);
                for (int k = 0; k < np && pUnitsL[k].Key == pUnitsR[k].Key; k++)
                    AddUnitMatchTokens(pUnitsL[k], pUnitsR[k], flatL, flatR, pass1);
                continue;
            }
            var unitsL = BuildUnits(flatL, offL[li], offL[li + 1]);
            var unitsR = BuildUnits(flatR, offR[ri], offR[ri + 1]);
            if (unitsL.Count == 0 || unitsR.Count == 0)
                continue;
            if ((long)unitsL.Count * unitsR.Count > LcsCellCap)
                return null;
            // Char-weighted, like the ordinary per-pair differ (Word's matched-character tie-break):
            // in-slot anchors ALWAYS stand for interior pairs — Word matches even an isolated single
            // word within its own slot — so the cross-run pass below sees only words with NO in-slot
            // counterpart. Discarding sparse pairs' anchors (an earlier density floor) flooded the
            // cross-windows with words that had in-slot homes and made the fusion rearrange what
            // Word keeps per-slot.
            var matches = UnitLcs(unitsL, unitsR, charWeighted: true, allowPartial: false);
            foreach (var (a, b) in matches)
                AddUnitMatchTokens(unitsL[a], unitsR[b], flatL, flatR, pass1);
        }

        // ---- Pass 2: residues between consecutive pass-1 anchors re-match across the whole run
        // (char-weighted, compound-endpoint partials allowed). Windows span paragraph boundaries, so
        // new matches here are exactly the boundary-crossing words.
        bool bail = false;
        int crossUnitMatches = 0; // pass-2 unit matches (words; separator extensions excluded)
        string? firstCrossKey = null; // the first pass-2 match's unit key (the c=1 construct evidence)
        int MemberOfL(int flat)
        {
            int m = 0;
            while (m + 1 < kl + 1 && offL[m + 1] <= flat) m++;
            return m;
        }
        int MemberOfR(int flat)
        {
            int m = 0;
            while (m + 1 < kr + 1 && offR[m + 1] <= flat) m++;
            return m;
        }
        // NB (decoded boundary, resolved 2026-07-27): a same-ordinal pass-2 match INSIDE the sealed
        // story-final pair can outweigh a genuine crossing chain in the window LCS and seal the run
        // ("centered bold" class), yet UNCONDITIONALLY weight-zeroing those pairs was measured
        // net-negative (a near-equal final pair's recoveries are the valve that keeps its shared
        // content retained — "yellow highlight"/"italic underline" classes). The resolution, decoded
        // pairwise from the reference outputs, is STRUCTURAL, not statistical: the final pair retains
        // in place exactly (a) its common unit prefix (anchored in pass 1 above) and (b) same-ordinal
        // runs containing an adjacent BIGRAM; isolated interior recoveries never stand, so a genuine
        // crossing chain wins the window instead. Coverage floors over the same evidence could not
        // separate the classes (retain at 0.300 vs cross at 0.286 matched-char coverage).
        void CrossWindow(int lFrom, int lTo, int rFrom, int rTo)
        {
            if (bail || lTo <= lFrom || rTo <= rFrom)
                return;
            var unitsL = BuildUnits(flatL, lFrom, lTo);
            var unitsR = BuildUnits(flatR, rFrom, rTo);
            if (unitsL.Count == 0 || unitsR.Count == 0)
                return;
            if ((long)unitsL.Count * unitsR.Count > LcsCellCap)
            {
                bail = true;
                return;
            }
            var picked = UnitLcs(unitsL, unitsR, charWeighted: true, allowPartial: true);
            if (runEndsStory && !hasTail && pairs.Count > 0 &&
                lTo > offL[lastPairL] && rTo > offR[lastPairR])
            {
                // Two-round decode for a window touching the story-final pair. Round B (the valve):
                // same-ordinal final-pair matches allowed. They STAND only when they include an
                // adjacent bigram — two consecutive matches on adjacent units of the final pair on
                // BOTH sides ("text is" staying put) — the decoded retained-in-place signature.
                // Otherwise round A re-runs the LCS with those matches suppressed, so an earlier
                // pair's crossing chain can win the window instead of being sealed by isolated
                // final-pair recoveries.
                bool InFinal(AnchorUnit ua, AnchorUnit ub) =>
                    MemberOfL(ua.Start) == lastPairL && MemberOfR(ub.Start) == lastPairR;
                bool bigram = false;
                for (int m = 1; m < picked.Count && !bigram; m++)
                {
                    var (a0, b0) = picked[m - 1];
                    var (a1, b1) = picked[m];
                    bigram = a1 == a0 + 1 && b1 == b0 + 1 &&
                             InFinal(unitsL[a0], unitsR[b0]) && InFinal(unitsL[a1], unitsR[b1]);
                }
                bool anyFinalSame = false;
                foreach (var (a2, b2) in picked)
                    anyFinalSame |= InFinal(unitsL[a2], unitsR[b2]);
                if (anyFinalSame && !bigram)
                    picked = UnitLcs(unitsL, unitsR, charWeighted: true, allowPartial: true,
                        (ua, ub) => InFinal(ua, ub));
            }
            foreach (var (a, b) in picked)
            {
                // FORWARD-only cross-flow (regions WITH pairs): every crossing Word's compare output
                // makes carries an EARLIER base paragraph's residue into a LATER output paragraph
                // (the stream flows forward past deleted pilcrows). A backward match — a later base
                // paragraph's word poaching an earlier right paragraph's slot — never occurs in the
                // decoded output; admitting them is exactly what rearranged content Word keeps
                // per-slot. A ZERO-pair region has no slots to poach — its matches are constrained
                // only by stream order (the decoded no-pair gap stream retains a later base
                // paragraph's function word against the first right paragraph's head).
                int ml = MemberOfL(unitsL[a].Start), mr = MemberOfR(unitsR[b].Start);
                if (pairs.Count > 0)
                {
                    if (ml > mr)
                        continue;
                    // TRAILING one-sided ↔ TRAILING one-sided never matches by content: an interior
                    // wordful↔wordful replace stays fully separate marked paragraphs (the decoded
                    // gap grammar — the aligner already declined to pair them). The story-final
                    // del+ins still FUSE into one cell structurally, and that cell's re-diff retains
                    // any shared tokens, so banning the stream-level match loses nothing there.
                    // LEADING one-sided members are the fusion construct itself (decoded 2026-07-27)
                    // and match freely.
                    if (leftPairIdx[ml] < 0 && rightPairIdx[mr] < 0 && ml > lastPairL && mr > lastPairR)
                        continue;
                }
                crossUnitMatches++;
                firstCrossKey ??= unitsL[a].Key;
                AddUnitMatchTokens(unitsL[a], unitsR[b], flatL, flatR, all);
            }
        }

        int prevL = 0, prevR = 0;
        foreach (var (lf, rf) in pass1)
        {
            CrossWindow(prevL, lf, prevR, rf);
            all.Add((lf, rf));
            prevL = lf + 1;
            prevR = rf + 1;
        }
        CrossWindow(prevL, totalL, prevR, totalR);
        if (bail)
            return null;
        // `all` is sorted by left AND right flat index: pass-1 anchors are monotone, and every pass-2
        // window is a disjoint interval on both sides between two consecutive pass-1 anchors.

        // ---- Separator-pair extension (decoded 2026-07-27 from reference compare output): the
        // whitespace DIRECTLY flanking a matched pair pairs too when both sides carry a key-equal
        // connective there — Word retains "text " (trailing space) in the ¶INS paragraph and " and "
        // (leading space) in the fused tail, i.e. the shared separator rides WITH the match across a
        // cell boundary instead of leaking into the previous cell's one-sided ins run. Stream order
        // (earlier pair's trailing extension first) realizes the decoded precedence: with one base
        // space between two matched words, the trailing attachment wins and the next side's extra
        // space stays inserted. Same-member constraints keep every extension flush with its pair, so
        // boundary pairing counts (cl/cr), the crossing gate, and cell-slice contiguity are all
        // preserved by construction — only the slice extents change.
        ExtendMatchesWithFlankingSeparators(all, flatL, flatR, offL, offR, kl, kr);

        // ---- Pilcrow pairing (structural): prefix boundary b retains a shared pilcrow iff the same
        // number of matched tokens precedes ¶Lb and ¶Rb on their own sides (no matched token between
        // them). The STORY-FINAL boundary of each side always pairs with the other side's — both
        // final pilcrows follow every match. With a one-sided tail the junction boundary (final on
        // exactly ONE side) never pairs as a prefix boundary: the story-final structural pair owns
        // that side's final pilcrow, so the other side's ¶ becomes a marked stream token.
        // The FINAL structural pair applies when the region ends the story (both story-final
        // pilcrows follow every match) or when the region's last members are its last word-matched
        // PAIR on both sides (a pure run, possibly with leading one-sided members — the run's
        // boundary to the following block is never touched). An INTERIOR region with trailing
        // one-sided members has no forced final pair: its side-final pilcrows are ordinary stream
        // tokens (a trailing inserted paragraph keeps its ¶INS before the following anchor block).
        bool forceFinal = runEndsStory ||
            (pairs.Count > 0 && lastPairL == kl - 1 && lastPairR == kr - 1);

        var pairedL = new bool[kl];
        var pairedR = new bool[kr];
        var pairedBounds = new List<(int Bl, int Br)>();
        // A boundary is a pairing CANDIDATE only between two members that are paired TOGETHER (the
        // pair's own terminating pilcrows — at their OWN ordinals, which may be skewed when
        // one-sided members precede the pair on one side only); a boundary owned by a one-sided
        // member is a marked stream token (¶INS/¶DEL) unless it is the final structural pair.
        foreach (var (li, ri) in pairs)
        {
            bool finalL = li == kl - 1, finalR = ri == kr - 1;
            if (forceFinal && (finalL || finalR))
                continue; // owned by (or junction to) the final structural pair appended below.
            int posL = offL[li + 1], posR = offR[ri + 1];
            int cl = 0, cr = 0;
            foreach (var (lf, rf) in all)
            {
                if (lf < posL) cl++;
                if (rf < posR) cr++;
            }
            if (cl == cr)
            {
                pairedL[li] = true;
                pairedR[ri] = true;
                pairedBounds.Add((li, ri));
            }
        }

        // Skewed COUNT-EQUAL boundary constructs (decoded 2026-07-27, ZERO-pair regions only): when
        // the same number of matched tokens (≥1) precedes a left and a right pilcrow — no matched
        // token between them in the merged stream — those pilcrows pair into a retained-¶ construct
        // even at DIFFERENT member ordinals (a single shared word mid-region materializes the
        // word-matched-pair shape without any aligner pair: the reference output retains " and " /
        // "document" under a shared pilcrow there). One construct per count value, earliest
        // boundary of each side, monotone by construction (counts are nondecreasing in position);
        // count 0 never pairs (leading one-sided paragraphs keep their marked pilcrows, per the
        // replace-gap grammar). Regions WITH pairs never enter — their one-sided boundary behavior
        // is pinned by the run+tail construct.
        int constructPairs = 0;
        if (pairs.Count == 0)
        {
            int CountL(int b)
            {
                int c = 0;
                foreach (var (lf, _) in all)
                    if (lf < offL[b + 1])
                        c++;
                return c;
            }
            int CountR(int b)
            {
                int c = 0;
                foreach (var (_, rf) in all)
                    if (rf < offR[b + 1])
                        c++;
                return c;
            }
            bool EligibleL(int b) => !pairedL[b] && (!forceFinal || b < kl - 1);
            bool EligibleR(int b) => !pairedR[b] && (!forceFinal || b < kr - 1);
            var countValues = new SortedSet<int>();
            for (int b = 0; b < kl; b++)
                if (EligibleL(b) && CountL(b) >= 1)
                    countValues.Add(CountL(b));
            foreach (int c in countValues)
            {
                int bl = -1, br = -1;
                for (int b = 0; b < kl && bl < 0; b++)
                    if (EligibleL(b) && CountL(b) == c)
                        bl = b;
                for (int b = 0; b < kr && br < 0; b++)
                    if (EligibleR(b) && CountR(b) == c)
                        br = b;
                if (bl < 0 || br < 0)
                    continue;
                // A construct supported by a SINGLE function-word match is positional scaffolding,
                // not correspondence — the decoded reference output keeps such regions with the
                // replace-gap grammar ("This"↔"This" heading two first paragraphs never pairs
                // them; a lone shared "and" mid-region renders as plain ins/del). Only a CONTENT
                // word ("document") forms a lone-match construct. (The count is UNIT matches —
                // boundary counts include separator extensions and cannot key this test.)
                if (crossUnitMatches == 1 && firstCrossKey is not null &&
                    IrBlockAligner.IsFunctionWordKey(firstCrossKey))
                    continue;
                pairedL[bl] = true;
                pairedR[br] = true;
                pairedBounds.Add((bl, br));
                constructPairs++;
                // AT MOST ONE construct per region, and the stream's MATCHING phase ends at it —
                // decoded from the reference output: after the construct's shared pilcrow, the
                // remaining region renders as the plain replace-gap arrangement (own-¶INS inserted
                // paragraphs, the last right's runs fused into the following ¶DEL cell, plain
                // deletions to the story end); a later scattered shared word ("and" again, four
                // paragraphs down) is never retained. Suppress every match past either construct
                // boundary and stop scanning count values.
                var kept = new List<(int Lf, int Rf)>(all.Count);
                foreach (var m in all)
                    if (m.Lf < offL[bl + 1] && m.Rf < offR[br + 1])
                        kept.Add(m);
                all.Clear();
                all.AddRange(kept);
                break;
            }
        }

        if (forceFinal)
        {
            pairedL[kl - 1] = true;
            pairedR[kr - 1] = true;
            pairedBounds.Add((kl - 1, kr - 1));
        }
        pairedBounds.Sort();

        // STRUCTURE gate: the fused stream ships only when it CHANGES the paragraph structure.
        // Without a tail: at least one interior boundary left unpaired (a ¶INS/¶DEL pilcrow, the
        // decoded signature of a genuine cross-boundary flow; e.g. a 3↔3 word-matched run emitting
        // 4 paragraphs). When every boundary pairs, the fusion would only redistribute within-pair
        // anchors versus the ordinary per-pair differ — all shape churn, no structural gain — so
        // the caller keeps the per-pair path. With a tail: at least one STRICTLY-crossing match
        // (an earlier member's residue landing in a later member — the decoded run+gap signature);
        // one-sided absorption with zero crossings falls back, because the ordinary replace-gap
        // grammar already renders that exact shape.
        if (pairs.Count == 0)
        {
            // Zero-pair region ship gate (decoded): ≥2 matched WORD units (the residue-pair
            // interleave threshold — a lone shared word with no construct renders via the
            // replace-gap grammar with no retention, a two-word crossing chain streams) OR a
            // count-equal boundary construct (a lone "and"/"document" match whose pilcrow counts
            // balance pairs the pilcrows and streams the region). Punctuation and separator
            // extensions never count as units. An INTERIOR region ships only on a construct.
            // The ≥2 floor was A/B-tested against the reference corpus (issue #699): relaxing it to
            // ≥1 ships 16 more regions for ZERO wins and one loss, and 15 of those 16 matched only
            // on function words or separators — exactly the positional scaffolding the construct
            // gate below refuses. It is the empirically right floor, not a tuning knob.
            // Either way the stream must CHANGE structure — ≥1 interior boundary left unpaired.
            bool anyInteriorUnpaired = false;
            for (int b = 0; b < kl - 1 && !anyInteriorUnpaired; b++)
                anyInteriorUnpaired = !pairedL[b];
            for (int b = 0; b < kr - 1 && !anyInteriorUnpaired; b++)
                anyInteriorUnpaired = !pairedR[b];
            bool ships = runEndsStory
                ? crossUnitMatches >= 2 || constructPairs > 0
                : constructPairs > 0;
            if (!ships || !anyInteriorUnpaired)
                return null;
        }
        else if (hasTail)
        {
            // A CROSSING match joins members that are not the same pair — one-sided matter flowing
            // into a pair, or one pair's residue into another. A pair's own in-slot matches never
            // count (member ordinals would misread a SKEWED pair's in-slot anchors as crossings
            // when one-sided members precede it on one side only). STORY-final mixed regions
            // require a genuine crossing (the pinned run+tail law); an INTERIOR region with pairs
            // ships on ANY match — the decoded interior construct hoists the pair's leading
            // ins-residue into the preceding ¶DEL cell and retains the pair's in-slot anchors,
            // with no cross-member match required.
            bool crossing = false;
            foreach (var (lf, rf) in all)
            {
                int pl = leftPairIdx[MemberOfL(lf)], pr = rightPairIdx[MemberOfR(rf)];
                if (pl < 0 || pr < 0 || pl != pr)
                {
                    crossing = true;
                    break;
                }
            }
            if (!crossing && (runEndsStory || all.Count == 0))
                return null;
        }
        else
        {
            bool anyUnpaired = false;
            for (int b = 0; b < kl - 1; b++)
                if (!pairedL[b])
                {
                    anyUnpaired = true;
                    break;
                }
            if (!anyUnpaired)
                return null;
        }

        // ---- Master anchor chain: matched tokens + paired boundaries, in stream order. A paired
        // boundary at (offL[bl+1], offR[br+1]) precedes any match at or beyond those positions; the
        // same-segment property makes this ordering monotone in both coordinates, and the story-final
        // pair follows every match by construction.
        var chain = new List<Anchor>(all.Count + pairedBounds.Count);
        {
            int mi = 0;
            foreach (var (bl, br) in pairedBounds)
            {
                int posL = offL[bl + 1];
                while (mi < all.Count && all[mi].Lf < posL)
                {
                    chain.Add(new Anchor(all[mi].Lf, all[mi].Rf, IsBoundary: false, 0, 0));
                    mi++;
                }
                chain.Add(new Anchor(posL, offR[br + 1], IsBoundary: true, bl, br));
            }
            while (mi < all.Count)
            {
                chain.Add(new Anchor(all[mi].Lf, all[mi].Rf, IsBoundary: false, 0, 0));
                mi++;
            }
        }

        // Unpaired boundaries, in ordinal (= position) order per side.
        var unpairedL = new List<int>();
        for (int b = 0; b < kl; b++)
            if (!pairedL[b])
                unpairedL.Add(b);
        var unpairedR = new List<int>();
        for (int b = 0; b < kr; b++)
            if (!pairedR[b])
                unpairedR.Add(b);

        // ---- Expansion: walk the chain; per window emit the RIGHT side's one-sided items (inserted
        // tokens + ¶INS boundaries, in right order) BEFORE the LEFT side's (deleted tokens + ¶DEL
        // boundaries, in left order) — the decoded new-content-first arrangement.
        var merged = new List<MergedItem>();
        int lc = 0, rc = 0;   // flat token cursors
        int ubL = 0, ubR = 0; // cursors into unpairedBounds per side

        void EmitOneSided(int lTo, int rTo)
        {
            while (true)
            {
                if (ubR < unpairedR.Count &&
                    offR[unpairedR[ubR] + 1] <= rc && offR[unpairedR[ubR] + 1] <= rTo)
                {
                    merged.Add(new MergedItem(ItemKind.BoundaryIns, -1, unpairedR[ubR]));
                    ubR++;
                }
                else if (rc < rTo)
                {
                    merged.Add(new MergedItem(ItemKind.Ins, -1, rc));
                    rc++;
                }
                else
                {
                    break;
                }
            }
            while (true)
            {
                if (ubL < unpairedL.Count &&
                    offL[unpairedL[ubL] + 1] <= lc && offL[unpairedL[ubL] + 1] <= lTo)
                {
                    merged.Add(new MergedItem(ItemKind.BoundaryDel, unpairedL[ubL], -1));
                    ubL++;
                }
                else if (lc < lTo)
                {
                    merged.Add(new MergedItem(ItemKind.Del, lc, -1));
                    lc++;
                }
                else
                {
                    break;
                }
            }
        }

        foreach (var a in chain)
        {
            EmitOneSided(a.LeftPos, a.RightPos);
            if (a.IsBoundary)
            {
                merged.Add(new MergedItem(ItemKind.BoundaryEqual, a.BoundaryLeft, a.BoundaryRight));
            }
            else
            {
                merged.Add(new MergedItem(ItemKind.Paired, a.LeftPos, a.RightPos));
                lc = a.LeftPos + 1;
                rc = a.RightPos + 1;
            }
        }
        EmitOneSided(totalL, totalR); // defensive: the final paired boundary is the last chain anchor.

        // ---- Factor the merged stream into cells: every boundary event closes one output paragraph.
        var cells = new List<IrCrossParagraphCell>();
        int curLeftPara = -1, curLeftStart = 0, curLeftLen = 0;
        int curRightPara = -1, curRightStart = 0, curRightLen = 0;

        bool AccumulateLeft(int lf)
        {
            int para = ParaOf(offL, kl, lf);
            int idx = lf - offL[para];
            if (curLeftPara < 0) { curLeftPara = para; curLeftStart = idx; curLeftLen = 1; }
            else if (para != curLeftPara || idx != curLeftStart + curLeftLen)
                return false; // ≥2 left paragraphs in one cell, or a non-contiguous slice — bail (defensive).
            else curLeftLen++;
            return true;
        }

        bool AccumulateRight(int rf)
        {
            int para = ParaOf(offR, kr, rf);
            int idx = rf - offR[para];
            if (curRightPara < 0) { curRightPara = para; curRightStart = idx; curRightLen = 1; }
            else if (para != curRightPara || idx != curRightStart + curRightLen)
                return false;
            else curRightLen++;
            return true;
        }

        void CloseCell(IrCrossParagraphMark mark, int leftParaIfEmpty, int rightParaIfEmpty)
        {
            if (curLeftPara < 0) curLeftPara = leftParaIfEmpty;
            if (curRightPara < 0) curRightPara = rightParaIfEmpty;
            var leftSlice = curLeftPara < 0
                ? (IReadOnlyList<IrDiffToken>)Array.Empty<IrDiffToken>()
                : SubTokens(leftReal[curLeftPara], curLeftStart, curLeftLen);
            var rightSlice = curRightPara < 0
                ? (IReadOnlyList<IrDiffToken>)Array.Empty<IrDiffToken>()
                : SubTokens(rightReal[curRightPara], curRightStart, curRightLen);
            cells.Add(new IrCrossParagraphCell(
                curLeftPara < 0 ? null : left[curLeftPara].Anchor.ToString(), curLeftStart, curLeftLen,
                curRightPara < 0 ? null : right[curRightPara].Anchor.ToString(), curRightStart, curRightLen,
                IrTokenDiffer.Diff(leftSlice, rightSlice, settings,
                    endsAtRetainedMark: mark == IrCrossParagraphMark.Equal),
                mark));
            curLeftPara = -1; curLeftStart = 0; curLeftLen = 0;
            curRightPara = -1; curRightStart = 0; curRightLen = 0;
        }

        foreach (var item in merged)
        {
            switch (item.Kind)
            {
                case ItemKind.Paired:
                    if (!AccumulateLeft(item.Left) || !AccumulateRight(item.Right))
                        return null;
                    break;
                case ItemKind.Ins:
                    if (!AccumulateRight(item.Right))
                        return null;
                    break;
                case ItemKind.Del:
                    if (!AccumulateLeft(item.Left))
                        return null;
                    break;
                case ItemKind.BoundaryEqual:
                    CloseCell(IrCrossParagraphMark.Equal, item.Left, item.Right);
                    break;
                case ItemKind.BoundaryIns:
                    CloseCell(IrCrossParagraphMark.Inserted, -1, item.Right);
                    break;
                case ItemKind.BoundaryDel:
                    CloseCell(IrCrossParagraphMark.Deleted, item.Left, -1);
                    break;
            }
        }

        // Nothing may trail the final boundary. A story-ending region's final cell must be a
        // retained (Equal) pilcrow — the run's boundary to the following block is never touched.
        // An INTERIOR region may instead end on a PURE one-sided marked cell (¶INS with no left
        // slice / ¶DEL with no right slice): removing that side drops the whole cell, so nothing
        // can ever fuse across the region boundary into the following anchor block. A trailing
        // marked cell CARRYING the other side's content (a hoist in a final ¶DEL cell) would fuse
        // into the anchor on accept/reject — bail.
        if (curLeftLen != 0 || curRightLen != 0 || curLeftPara >= 0 || curRightPara >= 0)
            return null;
        if (cells.Count == 0)
            return null;
        if (forceFinal)
        {
            if (cells[^1].Mark != IrCrossParagraphMark.Equal)
                return null;
        }
        else
        {
            var last = cells[^1];
            bool safeTail = last.Mark == IrCrossParagraphMark.Equal ||
                (last.Mark == IrCrossParagraphMark.Inserted && last.LeftAnchor is null) ||
                (last.Mark == IrCrossParagraphMark.Deleted && last.RightAnchor is null);
            if (!safeTail)
                return null;
        }

        return cells;
    }

    /// <summary>
    /// Extend the matched-token list with FLANKING SEPARATOR pairs: for each match (in stream order),
    /// pair the connective whitespace token directly before it — and directly after it — when both
    /// sides carry one with an equal key, it is unclaimed, and it lives in the same member paragraph
    /// as the match's own side (Word's stream has pilcrow tokens between paragraphs, so a separator is
    /// never adjacent to a match across a boundary). Claim tracking makes the pass order-safe: an
    /// earlier pair's trailing extension takes the only shared separator, and the later pair's leading
    /// candidate finds it claimed — the decoded trailing-attachment precedence. Monotonicity on both
    /// coordinates is preserved: an extension is flush with its pair, and any collision with the next
    /// pair (or its extension) is excluded by the claim check.
    /// </summary>
    private static void ExtendMatchesWithFlankingSeparators(
        List<(int Lf, int Rf)> all, IrDiffToken[] flatL, IrDiffToken[] flatR,
        int[] offL, int[] offR, int kl, int kr)
    {
        if (all.Count == 0)
            return;

        var claimedL = new bool[flatL.Length];
        var claimedR = new bool[flatR.Length];
        foreach (var (lf, rf) in all)
        {
            claimedL[lf] = true;
            claimedR[rf] = true;
        }

        static bool Connective(IrDiffToken t) =>
            t.Kind == IrDiffTokenKind.Separator && string.IsNullOrWhiteSpace(t.Text);

        bool Pairable(int l, int r, int lAnchor, int rAnchor) =>
            l >= 0 && r >= 0 && l < flatL.Length && r < flatR.Length &&
            !claimedL[l] && !claimedR[r] &&
            Connective(flatL[l]) && Connective(flatR[r]) &&
            flatL[l].MatchKey == flatR[r].MatchKey &&
            ParaOf(offL, kl, l) == ParaOf(offL, kl, lAnchor) &&
            ParaOf(offR, kr, r) == ParaOf(offR, kr, rAnchor);

        var extended = new List<(int Lf, int Rf)>(all.Count + 8);
        foreach (var (lf, rf) in all)
        {
            if (Pairable(lf - 1, rf - 1, lf, rf))
            {
                claimedL[lf - 1] = true;
                claimedR[rf - 1] = true;
                extended.Add((lf - 1, rf - 1));
            }
            extended.Add((lf, rf));
            if (Pairable(lf + 1, rf + 1, lf, rf))
            {
                claimedL[lf + 1] = true;
                claimedR[rf + 1] = true;
                extended.Add((lf + 1, rf + 1));
            }
        }
        all.Clear();
        all.AddRange(extended);
    }

    /// <summary>Build the content-anchor units over <c>flat[from..to)</c>: plain words, and compounds —
    /// maximal word(sep)word… chains whose interior separators are single non-whitespace tokens
    /// ("right-aligned", "A/B"). Whitespace and free-standing punctuation are connective (no unit).</summary>
    private static List<AnchorUnit> BuildUnits(IrDiffToken[] flat, int from, int to)
    {
        var units = new List<AnchorUnit>();
        int i = from;
        while (i < to)
        {
            if (flat[i].Kind != IrDiffTokenKind.Word)
            {
                i++;
                continue;
            }
            int start = i;
            var key = new System.Text.StringBuilder(flat[i].MatchKey);
            int chars = flat[i].Text.Length;
            string firstWord = flat[i].MatchKey;
            string lastWord = flat[i].MatchKey;
            i++;
            while (i + 1 < to &&
                   flat[i].Kind == IrDiffTokenKind.Separator && !string.IsNullOrWhiteSpace(flat[i].Text) &&
                   flat[i + 1].Kind == IrDiffTokenKind.Word)
            {
                key.Append(flat[i].MatchKey);
                chars += flat[i].Text.Length;
                key.Append(flat[i + 1].MatchKey);
                chars += flat[i + 1].Text.Length;
                lastWord = flat[i + 1].MatchKey;
                i += 2;
            }
            units.Add(new AnchorUnit(start, i - start, key.ToString(), chars, firstWord, lastWord));
        }
        return units;
    }

    /// <summary>LCS over anchor units. <paramref name="charWeighted"/> false = token-count objective
    /// (each match weighs 1 — the pass-1 shape, whose front-walk realizes the earliest-left tie-break);
    /// true = matched-char objective (pass 2). <paramref name="allowPartial"/> additionally lets a
    /// compound match a plain word equal to its first or last member (weight = the shared word's chars).</summary>
    private static List<(int A, int B)> UnitLcs(
        List<AnchorUnit> unitsL, List<AnchorUnit> unitsR, bool charWeighted, bool allowPartial,
        Func<AnchorUnit, AnchorUnit, bool>? suppress = null)
    {
        int n = unitsL.Count, m = unitsR.Count;
        var matches = new List<(int, int)>();
        if (n == 0 || m == 0)
            return matches;

        int Weight(int a, int b)
        {
            var ua = unitsL[a];
            var ub = unitsR[b];
            if (suppress is not null && suppress(ua, ub))
                return 0;
            if (ua.Key == ub.Key)
                return charWeighted ? Math.Max(1, Math.Min(ua.Chars, ub.Chars)) : 1;
            if (!allowPartial)
                return 0;
            // Compound ↔ plain word endpoint match (sub-word shape; the cell re-diff renders it).
            if (ua.Length > 1 && ub.Length == 1 && (ua.FirstWordKey == ub.Key || ua.LastWordKey == ub.Key))
                return charWeighted ? Math.Max(1, ub.Chars) : 1;
            if (ub.Length > 1 && ua.Length == 1 && (ub.FirstWordKey == ua.Key || ub.LastWordKey == ua.Key))
                return charWeighted ? Math.Max(1, ua.Chars) : 1;
            return 0;
        }

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
                matches.Add((i, j));
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
        return matches;
    }

    /// <summary>Expand one matched unit pair into matched TOKEN pairs. Exact-key units pair member by
    /// member (equal keys ⇒ identical member shape). A partial compound↔word match pairs ONLY the shared
    /// endpoint word token — the compound's other members stay one-sided residues.</summary>
    private static void AddUnitMatchTokens(
        AnchorUnit ua, AnchorUnit ub, IrDiffToken[] flatL, IrDiffToken[] flatR, List<(int Lf, int Rf)> sink)
    {
        if (ua.Key == ub.Key && ua.Length == ub.Length)
        {
            for (int t = 0; t < ua.Length; t++)
                sink.Add((ua.Start + t, ub.Start + t));
            return;
        }
        if (ua.Length > 1 && ub.Length == 1)
        {
            sink.Add((ua.FirstWordKey == ub.Key ? ua.Start : ua.Start + ua.Length - 1, ub.Start));
            return;
        }
        if (ub.Length > 1 && ua.Length == 1)
        {
            sink.Add((ua.Start, ub.FirstWordKey == ua.Key ? ub.Start : ub.Start + ub.Length - 1));
            return;
        }
        // Same key but different member count cannot occur (keys encode the member sequence); pair the
        // first tokens defensively so monotonicity is preserved.
        sink.Add((ua.Start, ub.Start));
    }

    private static int ParaOf(int[] off, int k, int flat)
    {
        // k is tiny (a run of adjacent pairs); the linear scan is clearest.
        for (int i = 0; i < k; i++)
            if (flat < off[i + 1])
                return i;
        return k - 1;
    }

    private static List<IrDiffToken> SubTokens(IReadOnlyList<IrDiffToken> tokens, int offset, int len)
    {
        var list = new List<IrDiffToken>(len);
        for (int i = offset; i < offset + len && i < tokens.Count; i++)
            list.Add(tokens[i]);
        return list;
    }
}

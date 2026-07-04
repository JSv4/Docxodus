#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.3 Task 2 — the differential harness. Runs BOTH revision engines head-to-head over the WC corpus
/// (<see cref="WcCorpus.BuildPairs"/>), both directions, and classifies each (pair, direction) by SEMANTIC
/// agreement of their revision sets. This is the first head-to-head measurement against the standing
/// program directive of full WmlComparer-test parity; its triage table is the M2.4 roadmap.
///
/// <para><b>OLD engine</b> (the shipped <see cref="WmlComparer"/>): <c>Compare(left, right, settings)</c>
/// produces a tracked-revisions document, then <c>GetRevisions(compared, settings)</c> reads it back to a
/// <c>List&lt;WmlComparerRevision&gt;</c>. <c>Compare</c> mutates nothing but is heavyweight (LCS over the
/// whole document); 92 pairs × 2 directions takes minutes — acceptable for a Trait=Differential test. The
/// old engine can THROW on some pairs (it predates the IR rewrite); we catch and record OLD_ERROR and keep
/// going.</para>
///
/// <para><b>NEW engine</b> (the IR revisions surface): <see cref="IrReader"/> ×2 (RetainSources=false,
/// Accept) → <see cref="IrEditScriptBuilder.Build"/> → <see cref="IrRevisionRenderer.Render"/>. A throw on
/// the new side is a REGRESSION, so the test asserts zero NEW_ERROR.</para>
/// </summary>
/// <remarks>
/// <para><b>Semantic normalization (the comparison contract).</b> The two engines atomize edits at
/// different granularities (WmlComparer can report a deletion run as one revision or split it across run
/// boundaries; the IR renderer emits one revision per token-op span). So we compare SEMANTICALLY, never
/// structurally:</para>
/// <list type="number">
/// <item>Each engine's revisions map to comparable atoms <c>(kind, normtext)</c>. <c>normtext</c> collapses
/// every run of whitespace to a single space, trims, and drops the atom if empty after normalization.
/// <b>Case is preserved</b> (a case flip is a real content change). Revision kinds map 1:1
/// (Inserted/Deleted/Moved/FormatChanged ↔ <see cref="IrRevisionType"/>).</item>
/// <item>Per kind we build a MULTISET of <c>normtext</c> (counts matter — two identical insertions are two
/// atoms).</item>
/// <item>For Inserted+Deleted we ALSO build a granularity-independent CONCATENATED CHAR BAG: the sorted
/// multiset of every character across all atoms of that kind. If one engine reports "ab" as one revision
/// and the other reports "a","b" as two, the multisets differ but the char bags are equal.</item>
/// </list>
///
/// <para><b>Classification</b> (per pair, per direction):</para>
/// <list type="bullet">
/// <item><b>MATCH</b> — per-kind <c>normtext</c> multisets are equal for all four kinds.</item>
/// <item><b>GRANULARITY</b> — multisets differ, but the COMBINED Inserted+Deleted char bag is equal AND the
/// Moved/FormatChanged multisets match: same content, different atomization. Deliberately loose triage
/// (within-kind reordering and ins↔del crossings are maskable here); advisory only — no correctness
/// assertion rides on this classification.</item>
/// <item><b>DIVERGENT</b> — anything else. Writes a per-pair detail file (gitignored artifacts dir) with the
/// per-kind set differences and both engines' counts, and is sorted into a known-cause sub-bucket where one
/// is mechanically detectable (see <see cref="DivergenceCause"/>).</item>
/// </list>
///
/// <para>The test asserts the totality invariants (zero NEW_ERROR, every pair classified, every DIVERGENT
/// pair has a detail file) AND a <b>gated semantic ratchet</b> over the adjudicated baseline distribution:
/// a Match-rate floor, a Divergent ceiling, and per-cause ceilings on the genuine-fidelity-loss buckets
/// (ScopeGapNewEmpty / OldEmpty / OpaqueGap / Unclassified). A quality regression that still PARSES — the
/// IR engine silently starting to over- or under-report vs WmlComparer — moves a case out of Match (or into
/// a fidelity-loss bucket) and turns the test RED. The benign documented-granularity buckets
/// (TokenSpanGranularity / PunctuationBoundary / MoveSemantics / FormatOnly / SpecialChars) are bounded only
/// by the Divergent ceiling so legitimate render-grain tuning is not over-constrained. The floors/ceilings
/// are a RATCHET: when the engine improves, tighten them to the new numbers (they may only move the good
/// way). The full triage table is still emitted for the controller to adjudicate.</para>
/// </remarks>
[Trait("Category", "Differential")]
public class IrVsWmlComparerTests
{
    private static readonly IrDiffSettings NewDiff = new();

    /// <summary>
    /// The render-time WmlComparer-compatible projection (M2.4 Task 2). The differential harness runs the
    /// SAME edit script through this granularity as a SECOND classification pass: under compatible-mode
    /// coalescing/trim/dedup the TokenSpanGranularity bucket (the dominant DIVERGENT family — same content,
    /// finer atomization) collapses, which the secondary triage table makes visible. The primary (Fine) pass
    /// keeps its role as the engine-truth measurement; the assertions ride on it.
    /// </summary>
    private static readonly IrDiffSettings CompatDiff = new() { RevisionGranularity = RevisionGranularity.WmlComparerCompatible };

    /// <summary>Same author on both engines so the author field never contributes a spurious difference.</summary>
    private const string Author = "Open-Xml-PowerTools";

    private readonly ITestOutputHelper _out;

    public IrVsWmlComparerTests(ITestOutputHelper output) => _out = output;

    // ---------------------------------------------------------------------- classification model

    private enum Classification { Match, Granularity, Divergent, OldError }

    /// <summary>
    /// Known-cause sub-buckets for a DIVERGENT pair, in priority order (first that applies wins). Each is a
    /// MECHANICAL heuristic — a hypothesis the report's root-cause read confirms against detail files, not a
    /// proof. <see cref="Unclassified"/> is the residual.
    /// </summary>
    /// <remarks>
    /// The char bags used below are WHITESPACE-FREE multisets (see <see cref="RevisionBags.InsDelCharBag"/>):
    /// inter-word spaces are excluded so that one engine reporting <c>"a b c"</c> as a single atom and the
    /// other reporting <c>"a","b","c"</c> as three are NOT penalized for the boundary spaces neither's normtext
    /// retains internally. That whitespace-independence is the whole point of the granularity comparison.
    /// </remarks>
    private enum DivergenceCause
    {
        /// <summary>
        /// The IR engine surfaced NOTHING (new total = 0) while the old engine reported revisions. The
        /// dominant cause across the corpus: the edit lives in a part/scope the IR diff path does not reach —
        /// textboxes, footnotes, endnotes (WC036/WC037/WC044–WC051/WC065–WC067). A clear M2.4 work item:
        /// extend IR reading/diffing to non-body scopes.
        /// </summary>
        ScopeGapNewEmpty,

        /// <summary>
        /// The OLD engine surfaced NOTHING (old total = 0) while the IR engine reported revisions. The
        /// heuristic flags these for the controller to confirm whether WmlComparer under-reported (IR more
        /// correct) or the IR over-reported (IR bug).
        /// <para><b>Correction (2026-06-11):</b> the original WC055/WC056 ("French apostrophe") example here was
        /// a MISDIAGNOSIS. Those pairs are a pure space→NBSP edit; under
        /// <see cref="IrDiffSettings.ConflateBreakingAndNonbreakingSpaces"/> that is NOT a content change, so
        /// WmlComparer's 0 was CORRECT and the IR's revisions were a tokenizer bug (NBSP folded only in the
        /// post-split match key, splitting token boundaries differently on the two sides). Fixed in
        /// <c>IrDiffTokenizer</c> (NBSP is a separator at split time when conflating); WC055/WC056 now MATCH at
        /// 0 and are NO LONGER an <c>OldEmpty</c> pair. The bucket itself remains valid for genuine
        /// WmlComparer under-reports.</para>
        /// </summary>
        OldEmpty,

        /// <summary>
        /// The new engine reports Moved where the old reports an Inserted+Deleted of the same content. Detected
        /// by re-bagging the new side's Moved atoms into the ins/del char bags and re-checking GRANULARITY: if
        /// folding moves away makes the two sides granularity-equal, the sole disagreement is move semantics.
        /// </summary>
        MoveSemantics,

        /// <summary>The Inserted+Deleted multisets agree and the ONLY disagreement is in the FormatChanged kind.</summary>
        FormatOnly,

        /// <summary>
        /// The symmetric-difference of the ins/del char bags consists ONLY of characters WmlComparer is known
        /// to drop: non-breaking hyphen (U+2011), soft hyphen (U+00AD), and the Private Use Area F000–F0FF
        /// (symbol-font glyph mappings). These are documented WmlComparer omissions.
        /// </summary>
        SpecialChars,

        /// <summary>
        /// The residual ins/del char-bag difference is ONLY trailing/word-adjacent PUNCTUATION (e.g. the
        /// sentence-final <c>.</c> or a <c>,</c> that one engine attaches to the changed word and the other
        /// splits off). The two engines agree on the changed LETTERS but disagree on how much adjacent
        /// punctuation belongs to the revision. WC001/WC043 are the canonical examples. The punctuation
        /// characters checked are exactly those NOT in <see cref="IrDiffSettings.DefaultWordSeparators"/> that
        /// nonetheless commonly bound a word: <c>. ! ? : " ' ’ …</c> (the separators themselves never reach
        /// here — they already split tokens identically on both sides).
        /// </summary>
        PunctuationBoundary,

        /// <summary>
        /// The pair's IR edit script contains a <c>ModifyBlock</c> whose block is opaque or a section break —
        /// the new engine's documented silent path (the renderer emits NOTHING for a Modified opaque/sectPr
        /// block), so the old engine may report content there that the new side cannot.
        /// </summary>
        OpaqueGap,

        /// <summary>
        /// The two engines agree on the changed LETTERS modulo a strict SUBSET relationship: ignoring
        /// whitespace and boundary punctuation, one side's ins+del letter bag strictly CONTAINS the other's.
        /// This is the dominant residual family — a block/phrase-pairing granularity disagreement. WmlComparer
        /// reports the MINIMAL changed words (e.g. del <c>"it","too"</c>), while the IR engine, when a
        /// paragraph pair falls under <see cref="IrDiffSettings.BlockSimilarityThreshold"/> (or a separator
        /// boundary lands differently), reports a WIDER span as edited and re-inserts the surrounding context
        /// it deleted (e.g. del <c>"after it.","before too."</c> + ins <c>"after.","before."</c>). Same edit,
        /// different amount of surrounding context attributed to it. WC002/WC007/WC042/WC052 are canonical.
        /// </summary>
        TokenSpanGranularity,

        /// <summary>Residual — none of the above mechanical heuristics fired.</summary>
        Unclassified,
    }

    private sealed record PairResult(
        string Label,
        Classification Class,
        DivergenceCause? Cause,
        int OldTotal,
        int NewTotal);

    // ---------------------------------------------------------------------- the harness

    [Fact]
    public void Ir_vs_WmlComparer_over_WC_corpus_both_directions()
    {
        var pairs = WcCorpus.BuildPairs();
        Assert.True(pairs.Count >= 30, $"Expected a substantial WC pair list; got {pairs.Count}.");

        var sw = Stopwatch.StartNew();
        var artifactsDir = ArtifactsDir();
        Directory.CreateDirectory(artifactsDir);
        ClearStaleDetails(artifactsDir);

        var results = new List<PairResult>();
        var compatResults = new List<PairResult>(); // M2.4 Task 2 — same edit script, compatible-mode render
        int newErrors = 0;
        var newErrorLabels = new List<string>();

        foreach (var (baseName, variantName) in pairs)
        {
            // Each WC file is read once per engine per direction. The OLD engine consumes WmlDocument bytes
            // off disk; the NEW engine consumes the IR. Both directions: (base→variant) and (variant→base).
            RunDirection(baseName, variantName, artifactsDir, results, compatResults, ref newErrors, newErrorLabels);
            RunDirection(variantName, baseName, artifactsDir, results, compatResults, ref newErrors, newErrorLabels);
        }

        sw.Stop();

        // ----- corpus report ------------------------------------------------------------------------
        _out.WriteLine($"Differential harness: {pairs.Count} WC pairs × 2 directions = {results.Count} comparisons");
        _out.WriteLine($"Wall time: {sw.Elapsed.TotalSeconds:F1}s");
        _out.WriteLine("");

        foreach (var r in results)
        {
            string bucket = r.Class == Classification.Divergent ? $"DIVERGENT/{r.Cause}" : r.Class.ToString().ToUpperInvariant();
            _out.WriteLine($"  {bucket,-22} old={r.OldTotal,-4} new={r.NewTotal,-4} {r.Label}");
        }
        _out.WriteLine("");

        _out.WriteLine("===== TRIAGE TABLE =====");
        foreach (var g in results.GroupBy(r => r.Class).OrderBy(g => (int)g.Key))
            _out.WriteLine($"  {g.Key,-12} {g.Count()}");
        _out.WriteLine("");
        _out.WriteLine("DIVERGENT sub-buckets:");
        foreach (var g in results.Where(r => r.Class == Classification.Divergent)
                     .GroupBy(r => r.Cause!.Value).OrderBy(g => (int)g.Key))
        {
            _out.WriteLine($"  {g.Key,-16} {g.Count()}");
            foreach (var ex in g.Take(5))
                _out.WriteLine($"      e.g. old={ex.OldTotal} new={ex.NewTotal}  {ex.Label}");
        }
        _out.WriteLine("");
        _out.WriteLine($"Detail files for DIVERGENT pairs written to: {artifactsDir}");

        // ----- compatible-mode differential (M2.4 Task 2) -------------------------------------------------
        // The SAME engine edit script, rendered under WmlComparer-compatible granularity, classified against
        // the SAME old-engine bags. The point of the second table: the TokenSpanGranularity bucket — pairs
        // that are "same content, finer atomization" under Fine mode — collapses into MATCH/GRANULARITY when
        // the renderer coalesces to WmlComparer's grain. Reported side by side so the collapse is visible.
        _out.WriteLine("");
        _out.WriteLine("===== TRIAGE TABLE — Fine vs WmlComparerCompatible (both render modes) =====");
        _out.WriteLine($"  {"Classification",-14} {"Fine",6} {"Compat",8}");
        foreach (var cls in Enum.GetValues<Classification>())
            _out.WriteLine($"  {cls,-14} {results.Count(r => r.Class == cls),6} {compatResults.Count(r => r.Class == cls),8}");
        _out.WriteLine("");
        _out.WriteLine("DIVERGENT sub-buckets — Fine vs Compat:");
        foreach (var cause in Enum.GetValues<DivergenceCause>())
        {
            int fine = results.Count(r => r.Class == Classification.Divergent && r.Cause == cause);
            int compat = compatResults.Count(r => r.Class == Classification.Divergent && r.Cause == cause);
            if (fine > 0 || compat > 0)
                _out.WriteLine($"  {cause,-22} {fine,6} {compat,8}");
        }

        // ----- assertions: totality invariants -----------------------------------------------------------
        Assert.Equal(2 * pairs.Count, results.Count); // totality: every pair classified, both directions
        Assert.Equal(results.Count, compatResults.Count); // the compatible-mode pass classifies every pair too
        Assert.True(newErrors == 0,
            $"NEW engine threw on {newErrors} comparison(s) — a regression: {string.Join(", ", newErrorLabels)}");
        foreach (var r in results.Where(r => r.Class == Classification.Divergent))
            Assert.True(File.Exists(Path.Combine(artifactsDir, DetailFileName(r.Label))),
                $"DIVERGENT pair missing its detail file: {r.Label}");

        // ----- gated semantic ratchet -------------------------------------------------------------------
        // Totality alone stays green even if the engine silently begins to over- or under-report (a quality
        // regression that still parses). These floors/ceilings — captured from the adjudicated baseline of
        // 184 comparisons (92 WC pairs × 2 directions) — turn such a regression RED: any case that LEAVES the
        // Match bucket drops the Match floor, and any case that worsens INTO a genuine-fidelity-loss bucket
        // breaks a per-cause ceiling. They are a RATCHET — when the engine improves, tighten to the new
        // numbers; they may only move in the good direction (floors up, ceilings down).
        //
        // SCOPE / KNOWN GAP (be honest about what this catches): the gates are bucket-COUNT gates, so they catch
        // a regression that MOVES a case out of Match (MatchFloor) or into divergence (DivergentCeiling / the
        // per-cause ceilings). They do NOT catch a regression that worsens WITHIN an already-Divergent case while
        // it stays in the unceilinged TokenSpanGranularity bucket (e.g. a PARTIAL under-report on a mixed
        // body+note pair that keeps newB.Total > 0 and a clean letter-bag containment). That specific note-/scope-
        // content gap is closed elsewhere — positively, by content round-trips: DocxDiffScenarioTests (synthetic
        // note edits) and DocxDiffBookmarkRealDocTests (the DD001 real-doc fixture's right side now edits footnote
        // AND endnote CONTENT, so accept≡right / reject≡left on referenced-note TEXT exercises the note-diff path).
        // So the perturbation guarantee here is: an over/under-report that shifts a previously-MATCH pair fails the
        // Match floor; within-bucket note/scope fidelity is guarded by those content round-trips.
        //
        // The absolute numbers are COUPLED to the current TestFiles/WC corpus snapshot (WcCorpus globs it at
        // runtime). Adding/removing a WC file legitimately re-shifts the distribution — re-baseline these
        // constants then (a louder red here than a silent miss, by design).
        int match     = results.Count(r => r.Class == Classification.Match);
        int divergent = results.Count(r => r.Class == Classification.Divergent);
        int oldError  = results.Count(r => r.Class == Classification.OldError);
        int Cause(DivergenceCause c) => results.Count(r => r.Class == Classification.Divergent && r.Cause == c);

        const int MatchFloor = 96;        // semantic agreement (Fine mode); may only RISE
        const int DivergentCeiling = 66;  // total semantic divergence; may only FALL
        const int OldErrorCeiling = 2;    // WC-BodyBookmarks both directions — the legacy engine throws. A CEILING
                                          // (not exact): if a WmlComparer fix stops it throwing, fewer is fine.

        Assert.True(match >= MatchFloor,
            $"Match rate REGRESSED: {match} < floor {MatchFloor}. A pair that previously agreed left the Match " +
            $"bucket — the IR engine is now over/under-reporting vs WmlComparer. See the triage table above.");
        Assert.True(divergent <= DivergentCeiling,
            $"Divergence GREW: {divergent} > ceiling {DivergentCeiling}. New semantic disagreement introduced.");
        Assert.True(oldError <= OldErrorCeiling,
            $"OldError GREW: {oldError} > ceiling {OldErrorCeiling}. More pairs now throw in the legacy WmlComparer " +
            "than the baseline — likely a corpus change; re-baseline after confirming it is not an IR-side regression.");

        // Genuine-fidelity-loss causes — these may not grow (baseline: all 0 except Unclassified=2).
        Assert.True(Cause(DivergenceCause.ScopeGapNewEmpty) <= 0,
            $"ScopeGapNewEmpty rose to {Cause(DivergenceCause.ScopeGapNewEmpty)}: IR surfaced NOTHING where WmlComparer reported revisions (an under-report / scope miss).");
        Assert.True(Cause(DivergenceCause.OldEmpty) <= 0,
            $"OldEmpty rose to {Cause(DivergenceCause.OldEmpty)}: IR reported revisions where WmlComparer surfaced none (an over-report).");
        Assert.True(Cause(DivergenceCause.OpaqueGap) <= 0,
            $"OpaqueGap rose to {Cause(DivergenceCause.OpaqueGap)}: IR missed an opaque/section sub-block the old engine covered.");
        Assert.True(Cause(DivergenceCause.Unclassified) <= 2,
            $"Unclassified rose to {Cause(DivergenceCause.Unclassified)}: an unexplained divergence none of the mechanical heuristics can bucket — inspect its detail file.");

        // The WmlComparer-compatible projection is the headline parity grain; its Match is the parity number
        // and must not regress either (baseline 150/184).
        const int CompatMatchFloor = 150;
        int compatMatch = compatResults.Count(r => r.Class == Classification.Match);
        Assert.True(compatMatch >= CompatMatchFloor,
            $"WmlComparer-compatible Match REGRESSED: {compatMatch} < floor {CompatMatchFloor}.");
    }

    private void RunDirection(
        string leftName, string rightName, string artifactsDir,
        List<PairResult> results, List<PairResult> compatResults, ref int newErrors, List<string> newErrorLabels)
    {
        string label = $"{Stem(leftName)} -> {Stem(rightName)}";

        // ----- OLD engine: Compare then GetRevisions. May throw (legacy engine) → OLD_ERROR. ------------
        RevisionBags? oldBags;
        try
        {
            oldBags = RunOldEngine(leftName, rightName);
        }
        catch (Exception ex)
        {
            results.Add(new PairResult(label, Classification.OldError, null, -1, -1));
            compatResults.Add(new PairResult(label, Classification.OldError, null, -1, -1));
            WriteOldErrorDetail(artifactsDir, label, ex);
            return;
        }

        // ----- NEW engine: IR pipeline. A throw here is a REGRESSION → recorded + asserted zero. --------
        RevisionBags newBags;
        IrEditScript script;
        IrDocument irLeft, irRight;
        try
        {
            irLeft = WcCorpus.ReadWc(leftName);
            irRight = WcCorpus.ReadWc(rightName);
            script = IrEditScriptBuilder.Build(irLeft, irRight, NewDiff);
            var revs = IrRevisionRenderer.Render(script, irLeft, irRight, NewDiff);
            // Header/footer-scope revisions are filtered from the DIFFERENTIAL comparison (2026-07-03
            // campaign): WmlComparer structurally cannot report them (it never diffs header/footer
            // scopes), so a Fine-mode hdr/ftr revision is capability the oracle lacks, not disagreement
            // with it — exactly the rationale of the OldError bucket. Without the filter the
            // WC004-Large↔-Mod pair (the corpus' one differing footer) leaves Match in both directions
            // for reporting a REAL footer change the oracle silently drops. Compat mode already excludes
            // them by definition (its contract is "the oracle's revision set"). The hdr/ftr surface is
            // gated positively elsewhere: the renderer battery's story round-trip (AssertRoundTrip) and
            // IrHeaderFooterDiffTests.
            newBags = RevisionBags.FromIr(revs.Where(r => !IsHeaderFooterScoped(r)));
        }
        catch (Exception ex)
        {
            newErrors++;
            newErrorLabels.Add(label);
            results.Add(new PairResult(label, Classification.Divergent, DivergenceCause.Unclassified, oldBags.Total, -1));
            compatResults.Add(new PairResult(label, Classification.Divergent, DivergenceCause.Unclassified, oldBags.Total, -1));
            WriteNewErrorDetail(artifactsDir, label, ex);
            return;
        }

        var (cls, cause) = Classify(oldBags, newBags, script, irLeft);
        if (cls == Classification.Divergent)
            WriteDivergenceDetail(artifactsDir, label, oldBags, newBags, cause!.Value);

        results.Add(new PairResult(label, cls, cls == Classification.Divergent ? cause : null,
            oldBags.Total, newBags.Total));

        // Compatible-mode pass: render the SAME edit script at WmlComparer-compatible granularity and classify
        // against the SAME old bags. No new detail files (the Fine pass owns those); totals-only here.
        var compatRevs = IrRevisionRenderer.Render(script, irLeft, irRight, CompatDiff);
        var compatBags = RevisionBags.FromIr(compatRevs);
        var (compatCls, compatCause) = Classify(oldBags, compatBags, script, irLeft);
        compatResults.Add(new PairResult(label, compatCls, compatCls == Classification.Divergent ? compatCause : null,
            oldBags.Total, compatBags.Total));
    }

    /// <summary>True when a revision anchors in a header/footer scope (<c>kind:hdrN:unid</c> /
    /// <c>kind:ftrN:unid</c>) — capability the legacy oracle lacks; see the filter above.</summary>
    private static bool IsHeaderFooterScoped(IrRevision r)
    {
        static bool Scoped(string? anchor)
        {
            if (anchor is null) return false;
            int colon = anchor.IndexOf(':');
            if (colon < 0) return false;
            var scope = anchor.AsSpan(colon + 1);
            return scope.StartsWith("hdr", StringComparison.Ordinal) ||
                   scope.StartsWith("ftr", StringComparison.Ordinal);
        }
        return Scoped(r.LeftAnchor) || Scoped(r.RightAnchor);
    }

    private static RevisionBags RunOldEngine(string leftName, string rightName)
    {
        var left = LoadWml(leftName);
        var right = LoadWml(rightName);
        var settings = new WmlComparerSettings { AuthorForRevisions = Author };
        var compared = WmlComparer.Compare(left, right, settings);
        var revs = WmlComparer.GetRevisions(compared, settings);
        return RevisionBags.FromWmlComparer(revs);
    }

    // ---------------------------------------------------------------------- classification

    private static (Classification, DivergenceCause?) Classify(
        RevisionBags oldB, RevisionBags newB, IrEditScript script, IrDocument left)
    {
        if (oldB.MultisetsEqual(newB))
            return (Classification.Match, null);

        // GRANULARITY: ins+del CHAR BAGS equal AND Moved + FormatChanged multisets match. Same content, the
        // two engines just atomized it differently (one revision vs many).
        if (oldB.IsGranularityVariant(newB))
            return (Classification.Granularity, null);

        // DIVERGENT — assign the first applicable known cause (priority order).
        return (Classification.Divergent, DiagnoseCause(oldB, newB, script, left));
    }

    private static DivergenceCause DiagnoseCause(
        RevisionBags oldB, RevisionBags newB, IrEditScript script, IrDocument left)
    {
        // SCOPE_GAP_NEW_EMPTY: the IR engine produced no revisions at all but the old engine did — the edit is
        // in a part/scope (textbox, footnote, endnote) the IR diff path doesn't reach. Checked FIRST: it is the
        // dominant, most actionable bucket and subsumes any incidental char-bag coincidence.
        if (newB.Total == 0 && oldB.Total > 0)
            return DivergenceCause.ScopeGapNewEmpty;

        // OLD_EMPTY: the old engine produced nothing while the IR engine reported a real edit (WmlComparer
        // under-report). The IR side is arguably the correct one here.
        if (oldB.Total == 0 && newB.Total > 0)
            return DivergenceCause.OldEmpty;

        // MOVE_SEMANTICS: fold each side's Moved atoms back into its ins/del char bags. If that makes the two
        // sides granularity-equal, the sole disagreement is whether content was reported as a move or as
        // ins+del of the same text.
        if (oldB.IsGranularityVariantFoldingMoves(newB))
            return DivergenceCause.MoveSemantics;

        // FORMAT_ONLY: the ins/del multisets agree; the only disagreement is FormatChanged.
        if (MultisetEqual(oldB.Multiset(IrRevisionType.Inserted), newB.Multiset(IrRevisionType.Inserted)) &&
            MultisetEqual(oldB.Multiset(IrRevisionType.Deleted), newB.Multiset(IrRevisionType.Deleted)) &&
            MultisetEqual(oldB.Multiset(IrRevisionType.Moved), newB.Multiset(IrRevisionType.Moved)) &&
            !MultisetEqual(oldB.Multiset(IrRevisionType.FormatChanged), newB.Multiset(IrRevisionType.FormatChanged)))
            return DivergenceCause.FormatOnly;

        // SPECIAL_CHARS: the (whitespace-free) symmetric difference of the ins+del char bags is entirely
        // WmlComparer's known dropped characters (U+2011, U+00AD, PUA F000–F0FF).
        if (OnlyDiffCharsSatisfy(oldB, newB, IsKnownDroppedChar))
            return DivergenceCause.SpecialChars;

        // PUNCTUATION_BOUNDARY: the residual char-bag difference is only word-adjacent punctuation that one
        // engine attaches to the changed word and the other splits off — the engines agree on the letters.
        if (OnlyDiffCharsSatisfy(oldB, newB, IsBoundaryPunctuation))
            return DivergenceCause.PunctuationBoundary;

        // OPAQUE_GAP: the IR script has a ModifyBlock over an opaque / section-break block — the renderer's
        // documented silent path (it emits no revision there), so content the old engine reports has no new
        // counterpart. (Distinct from SCOPE_GAP_NEW_EMPTY: here the IR side reports SOME revisions but misses
        // an opaque sub-block.)
        if (ScriptHasOpaqueOrSectionModify(script, left))
            return DivergenceCause.OpaqueGap;

        // TOKEN_SPAN_GRANULARITY: ignoring whitespace AND boundary punctuation, one side's ins+del LETTER bag
        // strictly contains the other's — the engines agree on the changed letters but disagree on how much
        // surrounding context is attributed to the edit (a block/phrase-pairing granularity difference).
        if (LetterBagStrictlyContains(oldB, newB) || LetterBagStrictlyContains(newB, oldB))
            return DivergenceCause.TokenSpanGranularity;

        return DivergenceCause.Unclassified;
    }

    /// <summary>
    /// True iff <paramref name="bigger"/>'s ins+del LETTER bag (whitespace and boundary punctuation removed)
    /// is a strict super-multiset of <paramref name="smaller"/>'s: every letter of the smaller appears in the
    /// bigger with at least equal count, and the bigger has strictly more letters overall. Equal letter bags
    /// are NOT containment (they would have been caught as PUNCTUATION_BOUNDARY).
    /// </summary>
    private static bool LetterBagStrictlyContains(RevisionBags bigger, RevisionBags smaller)
    {
        var big = bigger.InsDelLetterBag();
        var small = smaller.InsDelLetterBag();
        int bigCount = big.Values.Sum();
        int smallCount = small.Values.Sum();
        if (bigCount <= smallCount)
            return false;
        foreach (var (c, n) in small)
            if (!big.TryGetValue(c, out var m) || m < n)
                return false;
        return true;
    }

    /// <summary>
    /// True iff the (whitespace-free) symmetric difference of the two sides' ins+del char bags is non-empty
    /// and every differing character satisfies <paramref name="predicate"/>. Used for the SPECIAL_CHARS and
    /// PUNCTUATION_BOUNDARY buckets, which differ only in which characters they accept.
    /// </summary>
    private static bool OnlyDiffCharsSatisfy(RevisionBags oldB, RevisionBags newB, Func<char, bool> predicate)
    {
        var diff = SymmetricCharDifference(oldB.InsDelCharBag(), newB.InsDelCharBag());
        return diff.Count > 0 && diff.All(predicate);
    }

    /// <summary>
    /// Word-adjacent punctuation that bounds a token but is NOT a
    /// <see cref="IrDiffSettings.DefaultWordSeparators"/> member (separators already split identically on both
    /// engines, so they never reach this bucket).
    /// </summary>
    private static bool IsBoundaryPunctuation(char c) =>
        c is '.' or '!' or '?' or ':' or '"' or '\'' or '’' or '‘'
          or '“' or '”' or '…' or '*' or '/' or '[' or ']' or '{' or '}'
          or '<' or '>' or '=' or '+' or '%' or '&' or '@' or '#' or '|' or '\\' or '~' or '`' or '$' or '^';

    private static bool IsKnownDroppedChar(char c) =>
        c == '\u2011' ||                          // non-breaking hyphen
        c == '\u00AD' ||                          // soft hyphen
        (c >= '\uF000' && c <= '\uF0FF');        // Private Use Area (symbol-font glyph mappings)

    private static bool ScriptHasOpaqueOrSectionModify(IrEditScript script, IrDocument left)
    {
        foreach (var op in script.Operations)
        {
            if (op.Kind != IrEditOpKind.ModifyBlock || op.LeftAnchor is not { } anchor)
                continue;
            if (left.AnchorIndex.TryGetValue(anchor, out var block) &&
                block is IrOpaqueBlock or IrSectionBreak)
                return true;
        }
        return false;
    }

    // ---------------------------------------------------------------------- revision bags

    /// <summary>
    /// Per-kind multisets of normalized revision text plus the granularity-independent ins/del char bag.
    /// Both engines map into this shape; classification is pure set algebra over it.
    /// </summary>
    private sealed class RevisionBags
    {
        // kind -> (normtext -> count)
        private readonly Dictionary<IrRevisionType, Dictionary<string, int>> _byKind = new();

        // Move atoms kept as a flat normtext bag (source and dest text both contribute) for the
        // move-folding heuristic; Moved multiset above is the same content but per-atom.
        private readonly Dictionary<string, int> _moveTextBag = new();

        public int Total { get; private set; }

        private RevisionBags()
        {
            foreach (var k in Enum.GetValues<IrRevisionType>())
                _byKind[k] = new Dictionary<string, int>(StringComparer.Ordinal);
        }

        public static RevisionBags FromIr(IEnumerable<IrRevision> revs)
        {
            var b = new RevisionBags();
            foreach (var r in revs)
                b.Add(r.Type, r.Text);
            return b;
        }

        public static RevisionBags FromWmlComparer(IEnumerable<WmlComparer.WmlComparerRevision> revs)
        {
            var b = new RevisionBags();
            foreach (var r in revs)
                b.Add(MapKind(r.RevisionType), r.Text);
            return b;
        }

        private static IrRevisionType MapKind(WmlComparer.WmlComparerRevisionType t) => t switch
        {
            WmlComparer.WmlComparerRevisionType.Inserted => IrRevisionType.Inserted,
            WmlComparer.WmlComparerRevisionType.Deleted => IrRevisionType.Deleted,
            WmlComparer.WmlComparerRevisionType.Moved => IrRevisionType.Moved,
            WmlComparer.WmlComparerRevisionType.FormatChanged => IrRevisionType.FormatChanged,
            _ => throw new ArgumentOutOfRangeException(nameof(t), t, "Unknown WmlComparerRevisionType"),
        };

        private void Add(IrRevisionType kind, string? rawText)
        {
            string norm = Normalize(rawText);
            if (norm.Length == 0)
                return; // empty-after-normalization atoms are dropped on both sides
            Bump(_byKind[kind], norm);
            if (kind == IrRevisionType.Moved)
                Bump(_moveTextBag, norm);
            Total++;
        }

        public IReadOnlyDictionary<string, int> Multiset(IrRevisionType kind) => _byKind[kind];

        public bool MultisetsEqual(RevisionBags other) =>
            Enum.GetValues<IrRevisionType>().All(k => MultisetEqual(_byKind[k], other._byKind[k]));

        /// <summary>
        /// GRANULARITY: ins+del CHAR BAGS equal AND Moved + FormatChanged multisets match. Captures
        /// "same content, different atomization" for the text kinds while holding moves/format exact.
        /// </summary>
        public bool IsGranularityVariant(RevisionBags other) =>
            CharBagsEqual(InsDelCharBag(), other.InsDelCharBag()) &&
            MultisetEqual(_byKind[IrRevisionType.Moved], other._byKind[IrRevisionType.Moved]) &&
            MultisetEqual(_byKind[IrRevisionType.FormatChanged], other._byKind[IrRevisionType.FormatChanged]);

        /// <summary>
        /// Move-folding variant of <see cref="IsGranularityVariant"/>: each side's Moved text is folded into
        /// its ins/del char bag (a move = an insert at the destination + a delete at the source of the same
        /// text), and FormatChanged is still required to match. If THIS makes them granularity-equal, the move
        /// vs ins+del framing is the sole disagreement.
        /// </summary>
        public bool IsGranularityVariantFoldingMoves(RevisionBags other) =>
            CharBagsEqual(InsDelCharBagFoldingMoves(), other.InsDelCharBagFoldingMoves()) &&
            MultisetEqual(_byKind[IrRevisionType.FormatChanged], other._byKind[IrRevisionType.FormatChanged]);

        /// <summary>
        /// WHITESPACE-FREE char multiset across all Inserted + Deleted atoms (granularity-independent). Spaces
        /// are excluded so that one engine reporting <c>"a b c"</c> as a single atom and another reporting
        /// <c>"a","b","c"</c> as three compare EQUAL — the inter-word boundary spaces are an artifact of
        /// atomization, not of content, and including them would spuriously fail the granularity check (the
        /// WC053 finding). Letter-for-letter agreement modulo whitespace is exactly "same content".
        /// </summary>
        public Dictionary<char, int> InsDelCharBag()
        {
            var bag = new Dictionary<char, int>();
            AddCharsOf(_byKind[IrRevisionType.Inserted], bag);
            AddCharsOf(_byKind[IrRevisionType.Deleted], bag);
            return bag;
        }

        /// <summary>
        /// Like <see cref="InsDelCharBag"/> but ALSO drops boundary punctuation, leaving only the "letters"
        /// (everything that is neither whitespace nor <see cref="IsBoundaryPunctuation"/>). Used by the
        /// TOKEN_SPAN_GRANULARITY containment heuristic so that punctuation-attachment differences don't break
        /// an otherwise clean subset relationship between the two engines' changed-letter sets.
        /// </summary>
        public Dictionary<char, int> InsDelLetterBag()
        {
            var bag = new Dictionary<char, int>();
            foreach (var (c, n) in InsDelCharBag())
                if (!IsBoundaryPunctuation(c))
                    bag[c] = n;
            return bag;
        }

        private Dictionary<char, int> InsDelCharBagFoldingMoves()
        {
            var bag = InsDelCharBag();
            // A move contributes the same text once on each side (delete at source + insert at dest), so fold
            // each moved atom in twice to mirror the ins+del an unfolded engine would report.
            AddCharsOf(_moveTextBag, bag);
            AddCharsOf(_moveTextBag, bag);
            return bag;
        }

        private static void AddCharsOf(Dictionary<string, int> multiset, Dictionary<char, int> bag)
        {
            foreach (var (text, count) in multiset)
                for (int i = 0; i < count; i++)
                    foreach (char c in text)
                    {
                        if (char.IsWhiteSpace(c))
                            continue; // whitespace-free: boundary spaces are an atomization artifact, not content
                        bag[c] = bag.TryGetValue(c, out var n) ? n + 1 : 1;
                    }
        }

        private static void Bump(Dictionary<string, int> d, string key) =>
            d[key] = d.TryGetValue(key, out var n) ? n + 1 : 1;

        public string Dump(IrRevisionType kind)
        {
            var d = _byKind[kind];
            if (d.Count == 0)
                return "(none)";
            return string.Join(", ", d.OrderBy(kv => kv.Key, StringComparer.Ordinal)
                .Select(kv => kv.Value == 1 ? $"\"{kv.Key}\"" : $"\"{kv.Key}\"×{kv.Value}"));
        }
    }

    private static bool MultisetEqual(IReadOnlyDictionary<string, int> a, IReadOnlyDictionary<string, int> b)
    {
        if (a.Count != b.Count)
            return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var w) || w != v)
                return false;
        return true;
    }

    private static bool CharBagsEqual(Dictionary<char, int> a, Dictionary<char, int> b)
    {
        if (a.Count != b.Count)
            return false;
        foreach (var (k, v) in a)
            if (!b.TryGetValue(k, out var w) || w != v)
                return false;
        return true;
    }

    private static List<char> SymmetricCharDifference(Dictionary<char, int> a, Dictionary<char, int> b)
    {
        var diff = new List<char>();
        foreach (var c in a.Keys.Union(b.Keys))
        {
            int delta = (a.TryGetValue(c, out var av) ? av : 0) - (b.TryGetValue(c, out var bv) ? bv : 0);
            for (int i = 0; i < Math.Abs(delta); i++)
                diff.Add(c);
        }
        return diff;
    }

    // ---------------------------------------------------------------------- text normalization

    /// <summary>
    /// The comparison normalization, applied identically to both engines' revision text: every run of
    /// whitespace (spaces, tabs, newlines, NBSP-class) collapses to a single ASCII space, the result is
    /// trimmed, and an empty result yields the empty string (the caller drops empties). <b>Case is
    /// preserved</b> — a case change is a real content edit. This is the precise contract by which the two
    /// engines' differently-atomized revisions are made comparable.
    /// </summary>
    private static string Normalize(string? text)
    {
        if (string.IsNullOrEmpty(text))
            return string.Empty;
        var sb = new StringBuilder(text!.Length);
        bool pendingSpace = false;
        bool sawNonSpace = false;
        foreach (char c in text)
        {
            if (char.IsWhiteSpace(c) || c == ' ')
            {
                pendingSpace = sawNonSpace; // suppress leading whitespace
                continue;
            }
            if (pendingSpace)
            {
                sb.Append(' ');
                pendingSpace = false;
            }
            sb.Append(c);
            sawNonSpace = true;
        }
        return sb.ToString();
    }

    // ---------------------------------------------------------------------- artifacts

    private static WmlDocument LoadWml(string fileName)
    {
        var fi = new FileInfo(Path.Combine(WcCorpus.WcDir.FullName, fileName));
        Assert.True(fi.Exists, $"Missing WC test file: {fi.FullName}");
        return new WmlDocument(fi.FullName);
    }

    private void WriteDivergenceDetail(
        string dir, string label, RevisionBags oldB, RevisionBags newB, DivergenceCause cause)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"DIVERGENT  cause={cause}");
        sb.AppendLine($"pair: {label}");
        sb.AppendLine($"old total revisions = {oldB.Total}   new total revisions = {newB.Total}");
        sb.AppendLine();
        foreach (var kind in Enum.GetValues<IrRevisionType>())
        {
            sb.AppendLine($"[{kind}]");
            sb.AppendLine($"  OLD: {oldB.Dump(kind)}");
            sb.AppendLine($"  NEW: {newB.Dump(kind)}");
            sb.AppendLine($"  only-in-OLD: {DiffDump(oldB.Multiset(kind), newB.Multiset(kind))}");
            sb.AppendLine($"  only-in-NEW: {DiffDump(newB.Multiset(kind), oldB.Multiset(kind))}");
            sb.AppendLine();
        }
        File.WriteAllText(Path.Combine(dir, DetailFileName(label)), sb.ToString());
    }

    private static string DiffDump(IReadOnlyDictionary<string, int> a, IReadOnlyDictionary<string, int> b)
    {
        var onlyA = new List<string>();
        foreach (var (k, v) in a.OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            int delta = v - (b.TryGetValue(k, out var w) ? w : 0);
            if (delta > 0)
                onlyA.Add(delta == 1 ? $"\"{k}\"" : $"\"{k}\"×{delta}");
        }
        return onlyA.Count == 0 ? "(none)" : string.Join(", ", onlyA.Take(40));
    }

    private static void WriteNewErrorDetail(string dir, string label, Exception ex) =>
        File.WriteAllText(Path.Combine(dir, DetailFileName(label)),
            $"NEW_ERROR (regression) for pair: {label}\n\n{ex}");

    private void WriteOldErrorDetail(string dir, string label, Exception ex)
    {
        _out.WriteLine($"  OLD_ERROR  {label}: {ex.GetType().Name}: {ex.Message}");
        File.WriteAllText(Path.Combine(dir, DetailFileName(label)),
            $"OLD_ERROR (legacy engine threw) for pair: {label}\n\n{ex}");
    }

    private static string DetailFileName(string label)
    {
        string safe = label.Replace(" -> ", "__").Replace(' ', '_');
        foreach (char c in Path.GetInvalidFileNameChars())
            safe = safe.Replace(c, '_');
        return safe + ".diff";
    }

    private static void ClearStaleDetails(string dir)
    {
        foreach (var f in Directory.GetFiles(dir, "*.diff"))
            File.Delete(f);
    }

    private static string Stem(string fileName) =>
        fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ? fileName[..^5] : fileName;

    private static string ArtifactsDir([CallerFilePath] string thisFile = "") =>
        Path.Combine(Path.GetDirectoryName(thisFile)!, "DifferentialArtifacts");
}

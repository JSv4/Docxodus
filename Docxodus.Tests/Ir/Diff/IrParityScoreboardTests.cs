#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.3 Task 4 — the WmlComparer PARITY SCOREBOARD (the standing USER-DIRECTIVE deliverable). For every
/// RUNNABLE_NOW row of the <c>Docxodus.Tests/WmlComparer*</c> suite — the cases whose original assertion
/// rides on <c>GetRevisions</c> counts / types / texts / move semantics rather than on produced OOXML
/// markup — this test re-expresses that exact assertion against the IR engine via
/// <see cref="IrWmlComparerAdapter"/> and records PASS/FAIL.
///
/// <para><b>Measurement, not a gate.</b> The scoreboard's job is to establish the baseline M2.4 must drive
/// to 100%. EXPECTED FAILURES ARE FINE here: each ported case is scored with a SOFT ASSERT (a try/catch that
/// records the outcome instead of throwing), and the test asserts ONLY totality — every case ran, nothing
/// crashed the harness. The per-case PASS/FAIL table + totals are emitted to test output; the controller
/// reads them into the M2.3 Outcome scoreboard. Original tests are NOT modified and NOT weakened.</para>
///
/// <para><b>What is ported.</b> Each ported case carries its ORIGINAL test id and the assertion data copied
/// EXACTLY from the source test:</para>
/// <list type="bullet">
/// <item><b>WC003_Compare</b> (WmlComparerTests.cs, 105 rows) — <c>revisionCount == GetRevisions().Count()</c>.
/// File-based WC/RC/CA corpus pairs. Category C.</item>
/// <item><b>WC004_Compare_To_Self</b> (56 rows) — comparing a document to itself yields ZERO revisions.
/// Category C (the semantic content of the original's structural self-compare).</item>
/// <item><b>WC005_Compare_CaseInsensitive</b> (1 row) — count==2 under <c>CaseInsensitive</c>. Category C+G.</item>
/// <item><b>FormatChange GetRevisions</b> (WmlComparerFormatChangeTests.cs, 3 cases) — FormatChanged type
/// present, <c>ChangedPropertyNames</c> contains "bold", text contains "sample". Category E.</item>
/// <item><b>MoveDetection</b> (WmlComparerMoveDetectionTests.cs, 14 GetRevisions cases) — Moved counts,
/// MoveGroupId pairing, IsMoveSource, threshold/min-word/case settings. Category D.</item>
/// </list>
///
/// <para><b>What is NOT ported (MARKUP_BLOCKED / CONSOLIDATE / NOT_APPLICABLE — see the M2.3 Outcome table).</b>
/// Cases asserting on the produced document (validation, accept/reject round-trip, native
/// w:ins/w:del/moveFrom/moveTo/rPrChange elements, revision-id uniqueness) need OOXML markup the IR engine
/// does not emit until M2.4. Consolidate is out of v1 scope. Settings-default assertions on the
/// <c>WmlComparerSettings</c> object test the old engine's own type, not behavior, and have no IR analogue.</para>
/// </summary>
[Trait("Category", "Parity")]
public class IrParityScoreboardTests
{
    private readonly ITestOutputHelper _out;
    public IrParityScoreboardTests(ITestOutputHelper output) => _out = output;

    private static readonly DirectoryInfo SourceDir = new("../../../../TestFiles/");

    // ---------------------------------------------------------------------- the scoreboard run

    [Fact]
    public void Parity_scoreboard_over_runnable_now_WmlComparer_cases()
    {
        var board = new Scoreboard(DocumentedDeviations);

        foreach (var (id, left, right, expected) in WC003_Compare_Rows())
            board.Score(id, "C", () => Wc003(left, right, expected));

        foreach (var (id, name) in WC004_CompareToSelf_Rows())
            board.Score(id, "C", () => Wc004(name));

        board.Score("WCI-1000", "C+G", () => Wc005("WC/WC040-Case-Before.docx", "WC/WC040-Case-After.docx", 2));

        ScoreFormatChangeCases(board);
        ScoreMoveDetectionCases(board);

        board.Report(_out);

        // Totality: every scored case ran, none threw out of the soft-assert harness. Each case lands in one
        // of THREE states: PASS (count-exact to WmlComparer's GetRevisions), DEVIATION (a documented,
        // adjudicated expected-difference — see DocumentedDeviations; it is VISIBLE in the report with its
        // reason and counts toward the floor), or FAIL (an undocumented regression). The floor is a RATCHET on
        // PASS + DEVIATION: it may only go up.
        //
        // M2.4 Task 2 raised the floor from 133 to 179 (the full runnable set) by render-time WmlComparer-
        // compatible granularity (contiguous-region coalescing, word-boundary common-affix trim, zero-width
        // prune, Choice/Fallback textbox dedup) + the DetectMoves render switch. The residual 20 cases that
        // render-time projection cannot reconcile WITHOUT changing the engine (the binding adjudication forbids
        // touching alignment / the edit script's grain) are DOCUMENTED deviations, not failures — see the
        // catalog below for each one's root cause and why it is engine-level.
        // M2.4b Workstream C drove the GENUINE-PASS count from 161 to 173 by closing eight more deviation
        // rows: WC-1210/1420/1430/1440/1840 (adjacent-block insert/delete coalescing), WC-1770 (textbox-
        // interior compat coarsening), WC-1750/1760 (table-aware block similarity + note-scope empty-mark
        // prune). Workstream D then closed WC-1900 (174) by making the Choice/Fallback textbox dedup robust to
        // NON-ADJACENT (nested) branch interleaving via content-signature occurrence parity — the oracle
        // MC-resolves AlternateContent to one branch (PreProcessMarkup, ProcessAllParts/Office2007), so it
        // counts one copy; the reader walks both for projection parity, and the dedup now mirrors the oracle
        // wherever the Fallback copy lands. The five surviving deviations are all genuine engine-grain cases:
        // WC-1450/1830 (block-vs-atom paragraph granularity inside a cell), WC-1710/1720 (the IR's id-less
        // note-ref tokenization is coarser than the oracle at a verified mid-word note-ref relocation —
        // ORACLE CORRECT), WC-1920 (tokenizer punctuation-attachment grain in a textbox-nested table — the
        // duplicate half is now fixed; the residual -1 is the `test`/`test!` word-boundary difference). All
        // five are the same deferred-tokenizer / sub-paragraph-alignment class, deferred to M2.5.
        const int ParityFloor = 179;     // PASS + documented deviation = full runnable set (179/179)
        const int GenuinePassFloor = 174; // PASS-only ratchet (M2.4b WS-D: +WC-1900): may only rise
        Assert.True(board.Total > 0, "Scoreboard scored no cases.");
        Assert.Equal(board.Total, board.Pass + board.Deviation + board.Fail);
        Assert.True(board.Pass + board.Deviation >= ParityFloor,
            $"PARITY REGRESSION: {board.Pass} PASS + {board.Deviation} DEVIATION = {board.Pass + board.Deviation} " +
            $"< ratchet floor {ParityFloor}. Undocumented FAILs: " +
            string.Join(", ", board.FailingIds) + ". The scoreboard may only improve, and any new shortfall must " +
            "be either fixed at render time or moved to DocumentedDeviations with an adjudicated reason.");
        Assert.True(board.Pass >= GenuinePassFloor,
            $"GENUINE-PASS REGRESSION: {board.Pass} PASS < ratchet floor {GenuinePassFloor}. A row that was a " +
            "count-exact PASS has regressed to a deviation/fail — the genuine-pass ratchet may only rise.");
    }

    /// <summary>
    /// The adjudicated PASS_WITH_DOCUMENTED_DEVIATION catalog (M2.4 Task 2): scoreboard rows whose IR
    /// revision COUNT differs from WmlComparer's for a reason that render-time granularity compatibility
    /// CANNOT reconcile without changing the engine — and the binding adjudication makes the engine alignment
    /// and the edit script's grain untouchable. Each entry's value is the human-readable deviation reason
    /// shown in the report. A row here that actually PASSES is flagged (a stale deviation to remove); a row
    /// here that FAILS counts toward the floor as a DEVIATION, not a regression.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> DocumentedDeviations = new Dictionary<string, string>
    {
        // ---- M2.4b Workstream B CLOSED three rows of the "coincidental Equal island" family — now genuine
        // PASSES, no catalog entry:
        //   * WC-1170 (WC007-Longest-At-End: `Video provides.` → a 36-token paragraph): the 1×1 gap-residue
        //     aligner pairs the near-rewrite as Modified, then Myers credits the COINCIDENTAL `Video` (source
        //     `Video` vs the after-text's `Online Video`) as an Equal island, splitting one rewrite into 3.
        //   * WC-1950 (WC053-Text-in-cell: `from the beginning…` → `in accordance with the procedure…`): only
        //     coincidental function words (`the`/`of`/`each`) shared; Myers split the rewrite into 4.
        //   * WC-1190 (WC007-Moved-into-Table): the spurious 3rd revision was a zero-width Inserted over the
        //     EMPTY cell paragraph (a bare paragraph mark) the moved-into-table block left behind.
        // FIXES (render-time, WmlComparerCompatible only; Fine grain untouched — IrRevisionRenderer): (1) a
        // low-coverage coarsening — when a Modified pair's Equal+FormatChanged CONTENT-token coverage (larger
        // side) is below 0.67 AND the larger side has ≥8 content tokens (so a near-rewrite, not a short edit),
        // bridge the word-bearing Equal islands too and emit ONE del+ins (the affix trim still keeps any wholly
        // common edge); (2) an empty-paragraph-mark prune — a whole-block Insert/Delete of a paragraph with
        // ZERO content tokens surfaces no WmlComparer revision. The 0.67/≥8 thresholds were swept over the
        // corpus (stable plateau floor 0.55–0.72 × min 6–10); both fixes are size/coverage-gated so no passing
        // row regresses (verified: WC-1930's 3-token `designs that compleme` short edit stays fine-grained; the
        // standalone math/SmartArt/image block inserts WC-1320/1340/1350/1550 keep their revisions — only an
        // EMPTY-mark paragraph is pruned, not a math/image one).
        //
        // ---- Engine token-differ: coincidental sub-word matches / wider span attribution (TokenSpanGranularity).
        // WmlComparer's whole-document LCS reports the minimal changed phrase as ONE del + ONE ins; the IR
        // token differ (Myers over word tokens, per Modified pair) finds a coincidental interior token match
        // that splits the change into more revisions. The split is in the ENGINE'S edit script (the grain is
        // untouchable); render-time coalescing merges separator-bridged regions and (M2.4b) low-coverage
        // near-rewrites, but cannot UN-match a genuine equal token inside a HIGH-coverage in-place edit without
        // re-running the diff at a coarser grain — and coarsening THOSE would UNDER-report (regress WC-1930).
        // The residual rows below are all high-coverage edits or adjacent-block-coalescing/structural cases that
        // the low-coverage coarsening deliberately does NOT touch.
        // ---- M2.4b Workstream C CLOSED five more rows via render-time adjacent-block insert/delete
        // COALESCING (IrRevisionRenderer.RenderBlockOpList) — now genuine PASSES, no catalog entry:
        //   * WC-1210 (Para-before-table): the `Abcde` | empty-table | `fghij` insert run coalesces with the
        //     table as a sub-region START boundary, giving ins `Abcde` + ins `<table>\nfghij` = 2 (+ the del) = 3.
        //   * WC-1420/1430 (Math-heavy paragraph): the math/run-boundary InsertBlock fragments that previously
        //     surfaced one-per-block now coalesce into the adjacent text region, recovering WmlComparer's count.
        //   * WC-1440 (Image+math+para): the two CONSECUTIVE inserted body paragraphs (`Jean-Antoine…` +
        //     `Watteau…` + the math-only para) coalesce into ONE Inserted region; `End` stays separate (Equal
        //     blocks bound it). 10 == 10.
        //   * WC-1840 (Table-5 cell): the cell's `Video.` + math-only + `Click` consecutive inserts coalesce to
        //     ONE region; del `Video. Click.` stays = 2.
        // The coalescing is GATED so WC-1550 (two STANDALONE math-only paragraph inserts, zero word content)
        // and WC-1320/1340/1350 (standalone image/SmartArt inserts) keep one revision per block — a sub-region
        // with NO Word token is left un-coalesced (those are what WmlComparer counts individually). Verified
        // zero regression on the math/image corpus rows.
        ["WC-1450"] = "Table-4-row-image: +1 (8 vs 7). The cell's blocks aligned with a spurious InsertBlock `Video provides…` paired AGAINST a same-text ModifyBlock (the aligner anchored the wrong one of two identical `Video provides…` cell paragraphs), so the IR reports one extra whole-cell-paragraph insert WmlComparer folds into its contiguous w:ins region. This is a cell-INTERNAL block-alignment grain difference (which of two identical paragraphs anchors), not render-coalescible without re-pairing the cell — engine alignment. Kept as a deviation (the adjacent-block coalescing closes the inter-block runs; this residual is an intra-cell anchor ambiguity).",
        // WC-1940 (WC052-SmartArt-Same vs -Mod): CLOSED in M2.4b Workstream A — now a genuine PASS (IR 2 ==
        // WmlComparer 2). The two spurious empty-text revisions were over UNCHANGED pure-SmartArt paragraphs
        // whose diagram drawing-object id (wp:docPr/@id, 1 vs 2) and diagram rel ids differed side-to-side,
        // so their opaque content hashes diverged and the aligner paired them del+ins. IrHasher.Canonicalize
        // now strips the renumber-prone wp:docPr/@id and resolves relationship attributes to stable
        // content-identity tokens (matching WmlComparer's CloneBlockLevelContentForHashing), so the unchanged
        // paragraphs hash equal and align as Equal. No catalog entry — the row PASSES.
        // WC-1950 (WC053-Text-in-cell): CLOSED in M2.4b Workstream B — now a genuine PASS (see the WS-B note
        // above). The cell rewrite shared only coincidental function words; the low-coverage coarsening
        // (max-side content coverage 0.41 < 0.67, 19 content tokens ≥ 8) collapses it to one del+ins.

        // ---- WC-1710/1720 (WC034-Endnotes-Before vs -After3): IR 6 vs WmlComparer 7 (-1).
        //
        // WC034 'Video' INVESTIGATION (M2.4b WS-C — CORRECTS the prior "oracle spurious" misdiagnosis):
        // examined the raw OOXML of both sides' body paragraph. BEFORE the body run is `Video `[en-ref id=1]
        // `provides…` — `Video` is one contiguous text run. In After3 the endnote-reference id=1 has been
        // RELOCATED INTO THE MIDDLE of the word: the runs are `Vi`[en-ref id=1]`deo`` `[en-ref id=2]`provides…`
        // (verified run-by-run). So the word `Video` genuinely changed its run/atom structure — a note
        // reference was inserted between `Vi` and `deo`. WmlComparer's atom LCS faithfully reports `Video`
        // del + `Video` ins (the old contiguous atom is gone; the reconstituted `Vi`+`deo` is new). This is
        // CORRECT oracle behavior, NOT a fault — there is a real structural edit to the word.
        //
        // M2.5 Task 1 — TOKENIZER FIX LANDED (intra-word note-ref interruption). IrDiffTokenizer now marks a
        // word that a zero-width content atom (NoteRef/Image/Opaque/Textbox) splits WITHOUT a separator on
        // either side: the flanking words' MatchKeys carry an interruption marker keyed on the interrupting
        // atoms, so `Vi`⟨ref⟩`deo` is NOT word-equal to a contiguous `Video`. The Fine-mode edit script now
        // CORRECTLY reports the body change (verified: ModifyBlock token diff = Delete `Vi` / Insert
        // `Vi`+ref+`deo`, the `Video` word no longer Equal). A ref BETWEEN words (separator-adjacent — the
        // overwhelmingly common case) is untouched: zero corpus blast radius (full scoreboard unchanged at
        // 174 PASS + 5 DEV + 0 FAIL).
        //
        // RESIDUAL (why this is still a -1 deviation in WmlComparerCompatible COUNT): two SEPARATE,
        // render/note-subsystem effects the tokenizer fix does NOT address, isolated by the M2.5 oracle
        // breakdown (oracle body 3: del `Video` + ins `Video` + ins ``; oracle endnotes 4: en correspondence):
        //   (a) The WmlComparerCompatible common-affix trim cancels the body del `Video` / ins `Vi`+`deo`
        //       because the inserted letters spell the SAME `Video` (the relocated zero-width ref carries no
        //       text), so the body del+ins collapses to zero in the COUNT projection. (Fine mode keeps it.)
        //   (b) The note-store diff aligns notes BY w:id, while the oracle aligns by CONTENT/reference-order
        //       (it remaps note ids): After3 inserts a new en#1 and pushes Before's en#1 content to en#2, so
        //       by-id matching over-reports the note scope by +1. A note-store CONTENT-CORRESPONDENCE aligner
        //       (notes matched by content/reference-order, not raw id) is required — a broad note-subsystem
        //       change (it restructures IrNoteDiff + the note markup renderer + body↔note id remapping, and a
        //       trial regressed WC-1660/1670/1750/1760), the same class as deferred item #5 (note-store
        //       cross-part conversion). Effects (a)+(b) net to the -1 today; closing them is a follow-on, not
        //       part of the note-ref-within-word tokenization.
        // Kept as a deviation: the note-ref-within-word TOKENIZATION is FIXED (engine/Fine correct); the
        // residual count gap is note-store-correspondence-alignment (deferred).",
        ["WC-1710"] = "Endnote-After3: IR 6 vs WmlComparer 7 (-1). The note-ref-within-word TOKENIZATION is now FIXED (M2.5 Task 1): the Fine-mode edit script correctly reports `Video` del + `Vi`+ref+`deo` ins for the mid-word endnote-reference relocation (`Vi`[en-ref]`deo` vs Before's contiguous `Video `[en-ref], verified run-by-run). The residual -1 in the WmlComparerCompatible COUNT is TWO separate effects the tokenizer fix does not touch: (a) the compat common-affix trim cancels the text-identical body del/ins (`Video` == `Vi`+`deo`; the relocated ref is textless), and (b) the note-store diff aligns notes by w:id while the oracle aligns by content/reference-order (After3 renumbers en#1→en#2 and inserts a new en#1), over-reporting the note scope by +1 — a note-store CONTENT-CORRESPONDENCE aligner is required (broad note-subsystem change, deferred item #5 class). ORACLE CORRECT; tokenization FIXED; residual is note-store-correspondence-alignment.",
        ["WC-1720"] = "Reverse of WC-1710 (After3 → Before): same root as WC-1710. The mid-word note-ref interruption tokenization is FIXED (Fine-mode body diff correct); the residual -1 in the compat count is the same (a) text-identical-affix-trim + (b) note-store-by-id-vs-content-correspondence over-report. ORACLE CORRECT; tokenization FIXED; residual deferred (note-store correspondence).",

        // ---- Reader: textbox VML/DrawingML duplication NOT collapsed by the adjacent-pair dedup.
        // Word emits one logical textbox as a DrawingML mc:Choice + a VML mc:Fallback. The render-time dedup
        // collapses the pair when both land as ADJACENT textbox diffs in ONE paragraph (the common case,
        // fixed). When the two branches land in SEPARATE IR paragraphs/cells (textbox-in-cell), they are not
        // adjacent and the dedup cannot pair them without the reader's MC-preprocessing (WmlComparer's
        // approach) — an engine/reader change outside render scope.
        // ---- M2.4b Workstream C CLOSED WC-1770 via textbox-interior compat coarsening
        // (IrRevisionRenderer.RenderTextboxInnerOp) — now a genuine PASS, no catalog entry. WmlComparer never
        // descends w:txbxContent (it treats w:drawing as one OPAQUE comparison atom, WmlComparer.cs ~L8225/8673),
        // so a changed textbox surfaces as one whole-paragraph del+ins. The IR reader models the textbox interior
        // (for the markdown projection) and its token diff split `In1`→`In` to a lone Deleted `1`. In compatible
        // mode a textbox-interior Modified paragraph now renders as whole-block Deleted (left text) + Inserted
        // (right text), matching the oracle's coarser grain (2 == 2). Validated against WC-1890/2080 (interior
        // token diff already whole-paragraph) and WC-2090/2092 (interior insert/delete) — all keep passing.
        ["WC-1830"] = "Table-5 cell, SUB-PARAGRAPH content migration (+1, 3 vs 2). Before cell p0 = `Video provides…point. When you click…add.` (one paragraph); after the SAME text is SPLIT across after-p0 `Video provides…point. ` + a math paragraph + after-p2 `When you click…add.` (three paragraphs). WmlComparer's whole-document atom LCS reports this as one contiguous del+ins region (2). The IR aligns at PARAGRAPH grain — it cannot split before-p0's content across two after-paragraphs — so the block aligner's similarity pass pairs before-p0 with after-p0 (shared prefix, Modify), inserts the math paragraph, and matches after-p2 against the deleted before-p1 tail, surfacing one extra whole-paragraph revision. This is a genuine block-vs-atom GRANULARITY difference (WmlComparer's sub-paragraph LCS is finer than the IR's per-block token diff here), not a render-coalescible inter-block run and not an oracle fault — the IR's per-block account is internally consistent, just coarser at the paragraph boundary. engine alignment grain (sub-paragraph content migration).",
        // WC-1900 (WC048-Text-Box-in-Cell): CLOSED in M2.4b Workstream D — now a genuine PASS (IR 6 ==
        // WmlComparer 6). WmlComparer's PreProcessMarkup opens with MarkupCompatibilityProcessMode.ProcessAllParts
        // (Office2007), which MC-RESOLVES each mc:AlternateContent to a single branch and discards the other, so
        // the oracle sees ONE copy of every logical textbox and counts its change once (evidence: the Choice
        // declares Requires="wps", an Office2010 namespace the Office2007 processor cannot satisfy, so the SDK
        // keeps the VML Fallback — either way, ONE branch survives). The IR reader walks BOTH branches (M1.4
        // markdown-projection parity), so a changed textbox renders to two value-equal revision batches. The
        // render-time Choice/Fallback dedup collapsed only the ADJACENT-pair case; WC048's `Textbox3` is a NESTED
        // textbox whose two branches are interleaved with an empty wrapper body ([Textbox3, ⌀, Textbox3, ⌀], not
        // [Textbox3, Textbox3]), so the i/i+1 pair-walk never matched the pair and the duplicate leaked (+2).
        // Fixed by switching the dedup to CONTENT-SIGNATURE OCCURRENCE PARITY (emit odd occurrences, drop even):
        // each AlternateContent body appears an even number of times within a carrier paragraph, so the Fallback
        // copy is dropped wherever it lands. No catalog entry — the row PASSES. (Reader output is unchanged: the
        // fix is render-time/count-only, so the markdown-projection equivalence suite stays green and the markup
        // round-trip path — which keeps both branches because both LEFT and RIGHT carry both — is untouched.)
        ["WC-1920"] = "Table-in-textbox: IR 7 vs WmlComparer 8 (-1). The Choice/Fallback DUPLICATE is now correctly collapsed (the WS-D signature-parity dedup), so the residual -1 is NOT duplication — it is a tokenizer/affix-trim GRAIN difference inside the textbox's nested table. The cell text `This is a test` → `This is a test!` (one `!` appended). WmlComparer's word-atom LCS pairs the whole changed word as del `test` + ins `test!` (2 revisions); the IR tokenizes `!` as its own punctuation token, so compatible-mode common-affix trim reports just ins `!` (1 revision). Matching the oracle here needs the IR tokenizer to attach trailing punctuation to the preceding word the way WmlComparer's atomizer does — a corpus-wide tokenizer word-boundary change (Fine-mode + whole-corpus blast radius), the same deferred-tokenizer class as WC-1710/1720's note-ref-within-word and WC-1830's sub-paragraph grain. ORACLE follows its word-atom grain; IR finer at the punctuation-attachment boundary — deferred to M2.5.",

        // ---- M2.4b Workstream C CLOSED WC-1750/1760 (endnote-with-table) — now genuine PASSES, no catalog
        // entry. Two fixes combined: (1) table-aware block similarity + an unambiguous table-residue rule in
        // IrBlockAligner.FillOneGap pair the two endnote tables as Modified (was whole-table del+ins), feeding
        // IrTableDiffer's per-cell diff (Aaa→Aaa1, Eee→Eee1, deleted Ggg/Hhh/Iii row); (2) the empty-mark
        // prune is scoped to BODY paragraphs, so the inserted/deleted empty paragraph the oracle surfaces in
        // a note scope (`\n`) is no longer suppressed. 6 == 6 both directions.

        // NOTE — WC-1970/WC-1980 (WC055/WC056 French "l'article 1" → "l'article 1", a pure
        // space→NBSP edit) were FORMERLY catalogued here as a WmlComparer "oracle under-report". That was a
        // MISDIAGNOSIS: WmlComparer's 0 revisions is CORRECT — under ConflateBreakingAndNonbreakingSpaces an
        // NBSP↔space swap is not a content change. The IR engine's spurious 2 revisions were a real tokenizer
        // BUG: it folded NBSP→space only in the post-split match key, so the NBSP side glued "l'article 1"
        // into ONE word while the space side split it into three tokens — different boundaries, spurious diff.
        // Fixed in IrDiffTokenizer (NBSP is now a separator at SPLIT time when conflating); both rows now
        // genuinely PASS (0 == 0) and are no longer deviations. See IrDiffTokenizerTests
        // Nbsp_conflation_on_* and the dated correction in the M2.3 plan Outcome.
    };

    // ---------------------------------------------------------------------- WC003: revisionCount parity

    private void Wc003(string name1, string name2, int expected)
    {
        var revs = AdapterRevisions(name1, name2, new WmlComparerSettings());
        SoftEqual(expected, revs.Count, "revisionCount");
    }

    // ---------------------------------------------------------------------- WC004: compare-to-self ⇒ 0

    private void Wc004(string name)
    {
        var revs = AdapterRevisions(name, name, new WmlComparerSettings());
        SoftEqual(0, revs.Count, "self-compare revisionCount");
    }

    // ---------------------------------------------------------------------- WC005: case-insensitive count

    private void Wc005(string name1, string name2, int expected)
    {
        var settings = new WmlComparerSettings
        {
            CaseInsensitive = true,
            CultureInfo = System.Globalization.CultureInfo.CurrentCulture,
        };
        var revs = AdapterRevisions(name1, name2, settings);
        SoftEqual(expected, revs.Count, "case-insensitive revisionCount");
    }

    private static List<IrRevision> AdapterRevisions(string name1, string name2, WmlComparerSettings settings)
    {
        var left = new WmlDocument(Path.Combine(SourceDir.FullName, name1));
        var right = new WmlDocument(Path.Combine(SourceDir.FullName, name2));
        return IrWmlComparerAdapter.GetRevisions(left, right, settings);
    }

    // ---------------------------------------------------------------------- format-change cases (E)

    private void ScoreFormatChangeCases(Scoreboard board)
    {
        // GetRevisions_FormatChange_ShouldReturnFormatChangedType
        board.Score("FC-ReturnFormatChangedType", "E", () =>
        {
            var revs = FormatChangeRevisions(Para("This is some sample text."), BoldPara("This is some sample text."));
            SoftTrue(revs.Any(r => r.Type == IrRevisionType.FormatChanged), "has FormatChanged revision");
        });

        // GetRevisions_FormatChange_ShouldHaveFormatChangeDetails — ChangedPropertyNames contains "bold"
        board.Score("FC-HaveFormatChangeDetails", "E", () =>
        {
            var revs = FormatChangeRevisions(Para("This is some sample text."), BoldPara("This is some sample text."));
            var fc = revs.FirstOrDefault(r => r.Type == IrRevisionType.FormatChanged);
            SoftTrue(fc is not null, "has FormatChanged revision");
            SoftTrue(fc?.FormatChange is not null, "FormatChange details present");
            SoftTrue(fc?.FormatChange?.ChangedPropertyNames.Contains("bold") == true, "ChangedPropertyNames contains 'bold'");
        });

        // GetRevisions_FormatChange_ShouldHaveCorrectText — Text contains "sample"
        board.Score("FC-HaveCorrectText", "E", () =>
        {
            var revs = FormatChangeRevisions(Para("This is some sample text."), BoldPara("This is some sample text."));
            var fc = revs.FirstOrDefault(r => r.Type == IrRevisionType.FormatChanged);
            SoftTrue(fc is not null, "has FormatChanged revision");
            SoftTrue(fc?.Text?.Contains("sample") == true, "Text contains 'sample'");
        });

        // FormatChange_WithTextChange_ShouldTrackBoth — plain "Hello world" → bold "Hello there":
        // >0 revisions, with at least one Inserted/Deleted (text change). Category C.
        board.Score("FormatChange_WithTextChange_ShouldTrackBoth", "C", () =>
        {
            var revs = FormatChangeRevisions(Para("Hello world"), BoldPara("Hello there"));
            SoftTrue(revs.Count > 0, "has revisions");
            SoftTrue(revs.Any(r => r.Type is IrRevisionType.Inserted or IrRevisionType.Deleted), "has ins/del");
        });
    }

    private static List<IrRevision> FormatChangeRevisions(WmlDocument doc1, WmlDocument doc2) =>
        IrWmlComparerAdapter.GetRevisions(doc1, doc2, new WmlComparerSettings { DetectFormatChanges = true });

    // ---------------------------------------------------------------------- move-detection cases (D)

    private void ScoreMoveDetectionCases(Scoreboard board)
    {
        var moveSettings = new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.8, MoveMinimumWordCount = 3 };

        // MoveDetection_IdenticalText_ShouldMarkAsMove — ≥2 Moved, each group has source+dest
        board.Score("MoveDetection_IdenticalText_ShouldMarkAsMove", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "This is paragraph A with enough words for move detection.",
                        "This is paragraph B with sufficient content here.",
                        "This is paragraph C with more words added." },
                new[] { "This is paragraph B with sufficient content here.",
                        "This is paragraph A with enough words for move detection.",
                        "This is paragraph C with more words added." }, moveSettings);
            var moved = revs.Where(r => r.Type == IrRevisionType.Moved).ToList();
            SoftTrue(moved.Count >= 2, $"≥2 Moved (got {moved.Count})");
            foreach (var g in moved.GroupBy(r => r.MoveGroupId))
            {
                SoftTrue(g.Key is not null, "MoveGroupId set");
                SoftTrue(g.Any(r => r.IsMoveSource == true), "group has source");
                SoftTrue(g.Any(r => r.IsMoveSource == false), "group has destination");
            }
        });

        // MoveDetection_SimilarText_AboveThreshold_ShouldMarkAsMove — ≥2 Moved
        board.Score("MoveDetection_SimilarText_AboveThreshold_ShouldMarkAsMove", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "The quick brown fox jumps over the lazy dog today.", "Another paragraph here." },
                new[] { "Another paragraph here.", "The quick brown fox jumps over the lazy dog now." }, moveSettings);
            SoftTrue(revs.Count(r => r.Type == IrRevisionType.Moved) >= 2, "≥2 Moved for similar text");
        });

        // MoveDetection_DissimilarText_BelowThreshold_ShouldRemainInsertedDeleted — no Moved; has ins/del
        board.Score("MoveDetection_DissimilarText_BelowThreshold_ShouldRemainInsertedDeleted", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "The quick brown fox jumps over the lazy dog.", "Another paragraph here." },
                new[] { "Another paragraph here.", "A completely different sentence with new words entirely." }, moveSettings);
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved), "no Moved");
            SoftTrue(revs.Any(r => r.Type is IrRevisionType.Inserted or IrRevisionType.Deleted), "has ins/del");
        });

        // MoveDetection_ShortText_BelowMinimum_ShouldRemainInsertedDeleted — no Moved containing Hello/world
        board.Score("MoveDetection_ShortText_BelowMinimum_ShouldRemainInsertedDeleted", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "Hello world", "Another paragraph here with more content." },
                new[] { "Another paragraph here with more content.", "Hello world" }, moveSettings);
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved &&
                (r.Text?.Contains("Hello") == true || r.Text?.Contains("world") == true)), "short text not moved");
        });

        // MoveDetection_Disabled_ShouldNotDetectMoves — DetectMoves=false ⇒ no Moved
        board.Score("MoveDetection_Disabled_ShouldNotDetectMoves", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "This is paragraph A with enough words for move detection.",
                        "This is paragraph B with sufficient content here." },
                new[] { "This is paragraph B with sufficient content here.",
                        "This is paragraph A with enough words for move detection." },
                new WmlComparerSettings { DetectMoves = false });
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved), "no Moved when disabled");
        });

        // MoveDetection_CustomThreshold_ShouldRespectSetting — movesLow ≥ movesHigh
        board.Score("MoveDetection_CustomThreshold_ShouldRespectSetting", "D", () =>
        {
            var left = new[] { "The quick brown fox jumps over the lazy dog in the park.", "Another paragraph here." };
            var right = new[] { "Another paragraph here.", "The quick brown cat runs under the sleepy dog in the yard." };
            int low = MoveRevs(left, right, new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.5, MoveMinimumWordCount = 3 })
                .Count(r => r.Type == IrRevisionType.Moved);
            int high = MoveRevs(left, right, new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.9, MoveMinimumWordCount = 3 })
                .Count(r => r.Type == IrRevisionType.Moved);
            SoftTrue(low >= high, $"low({low}) >= high({high})");
        });

        // MoveDetection_CustomMinWordCount_ShouldRespectSetting — min3 ≥ min5 (for "Four..." text)
        board.Score("MoveDetection_CustomMinWordCount_ShouldRespectSetting", "D", () =>
        {
            var left = new[] { "Four word sentence here.", "Another paragraph with more content for testing purposes." };
            var right = new[] { "Another paragraph with more content for testing purposes.", "Four word sentence here." };
            int min3 = MoveRevs(left, right, new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.8, MoveMinimumWordCount = 3 })
                .Count(r => r.Type == IrRevisionType.Moved && r.Text?.Contains("Four") == true);
            int min5 = MoveRevs(left, right, new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.8, MoveMinimumWordCount = 5 })
                .Count(r => r.Type == IrRevisionType.Moved && r.Text?.Contains("Four") == true);
            SoftTrue(min3 >= min5, $"min3({min3}) >= min5({min5})");
        });

        // MoveDetection_CaseInsensitive_ShouldMatchIgnoringCase — ≥2 Moved
        board.Score("MoveDetection_CaseInsensitive_ShouldMatchIgnoringCase", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "THE QUICK BROWN FOX JUMPS OVER THE LAZY DOG.", "Another paragraph here." },
                new[] { "Another paragraph here.", "the quick brown fox jumps over the lazy dog." },
                new WmlComparerSettings { DetectMoves = true, MoveSimilarityThreshold = 0.8, MoveMinimumWordCount = 3, CaseInsensitive = true });
            SoftTrue(revs.Count(r => r.Type == IrRevisionType.Moved) >= 2, "case-insensitive ≥2 Moved");
        });

        // MoveDetection_MultipleMoves_ShouldMatchCorrectly — each group ≥2, has source+dest
        board.Score("MoveDetection_MultipleMoves_ShouldMatchCorrectly", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "First paragraph with content alpha beta gamma.",
                        "Second paragraph with content delta epsilon zeta.",
                        "Third paragraph with content eta theta iota.",
                        "Fourth paragraph with content kappa lambda mu." },
                new[] { "Third paragraph with content eta theta iota.",
                        "Fourth paragraph with content kappa lambda mu.",
                        "First paragraph with content alpha beta gamma.",
                        "Second paragraph with content delta epsilon zeta." }, moveSettings);
            var moved = revs.Where(r => r.Type == IrRevisionType.Moved).ToList();
            foreach (var gid in moved.Where(r => r.MoveGroupId.HasValue).Select(r => r.MoveGroupId!.Value).Distinct())
            {
                var grp = moved.Where(r => r.MoveGroupId == gid).ToList();
                SoftTrue(grp.Count >= 2, $"group {gid} ≥2 revisions");
                SoftTrue(grp.Any(r => r.IsMoveSource == true), $"group {gid} has source");
                SoftTrue(grp.Any(r => r.IsMoveSource == false), $"group {gid} has destination");
            }
        });

        // MoveDetection_EmptyDocument_ShouldNotThrow — no Moved
        board.Score("MoveDetection_EmptyDocument_ShouldNotThrow", "D", () =>
        {
            var revs = MoveRevs(Array.Empty<string>(),
                new[] { "New content added here with several words." }, new WmlComparerSettings { DetectMoves = true });
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved), "no Moved from empty");
        });

        // MoveDetection_IdenticalDocuments_ShouldHaveNoRevisions — empty
        board.Score("MoveDetection_IdenticalDocuments_ShouldHaveNoRevisions", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "Same content in both documents with enough words." },
                new[] { "Same content in both documents with enough words." }, new WmlComparerSettings { DetectMoves = true });
            SoftEqual(0, revs.Count, "identical docs ⇒ 0 revisions");
        });

        // MoveDetection_OnlyDeletions_ShouldNotCreateMoves — no Moved; has Deleted
        board.Score("MoveDetection_OnlyDeletions_ShouldNotCreateMoves", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "First paragraph that will be deleted.", "Second paragraph that stays here." },
                new[] { "Second paragraph that stays here." }, new WmlComparerSettings { DetectMoves = true });
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved), "no Moved");
            SoftTrue(revs.Any(r => r.Type == IrRevisionType.Deleted), "has Deleted");
        });

        // MoveDetection_OnlyInsertions_ShouldNotCreateMoves — no Moved; has Inserted
        board.Score("MoveDetection_OnlyInsertions_ShouldNotCreateMoves", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "First paragraph that stays here." },
                new[] { "First paragraph that stays here.", "Second paragraph that is newly added." },
                new WmlComparerSettings { DetectMoves = true });
            SoftTrue(!revs.Any(r => r.Type == IrRevisionType.Moved), "no Moved");
            SoftTrue(revs.Any(r => r.Type == IrRevisionType.Inserted), "has Inserted");
        });

        // MoveDetection_RevisionProperties_ShouldBeCorrect — every Moved has MoveGroupId>0 + IsMoveSource set
        board.Score("MoveDetection_RevisionProperties_ShouldBeCorrect", "D", () =>
        {
            var revs = MoveRevs(
                new[] { "Paragraph to be moved with enough words for detection.",
                        "Static paragraph that does not change here." },
                new[] { "Static paragraph that does not change here.",
                        "Paragraph to be moved with enough words for detection." }, moveSettings);
            var moved = revs.Where(r => r.Type == IrRevisionType.Moved).ToList();
            foreach (var rev in moved)
            {
                SoftTrue(rev.MoveGroupId is not null && rev.MoveGroupId > 0, "MoveGroupId > 0");
                SoftTrue(rev.IsMoveSource is not null, "IsMoveSource set");
            }
        });
    }

    private static List<IrRevision> MoveRevs(string[] left, string[] right, WmlComparerSettings settings) =>
        IrWmlComparerAdapter.GetRevisions(Doc(left), Doc(right), settings);

    // ---------------------------------------------------------------------- soft-assert plumbing

    /// <summary>Thrown by a soft assertion; caught by <see cref="Scoreboard.Score"/> to record a FAIL
    /// without aborting the run.</summary>
    private sealed class SoftAssertException : Exception
    {
        public SoftAssertException(string message) : base(message) { }
    }

    private static void SoftTrue(bool condition, string what)
    {
        if (!condition)
            throw new SoftAssertException($"expected {what}");
    }

    private static void SoftEqual<T>(T expected, T actual, string what)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new SoftAssertException($"{what}: expected {expected}, got {actual}");
    }

    /// <summary>One of the three scoreboard outcomes for a scored case.</summary>
    private enum RowState { Pass, Deviation, Fail }

    /// <summary>
    /// Per-case PASS / DEVIATION / FAIL accumulator that emits the parity table and totals. A case that the
    /// soft asserts mark failing but whose id is in the documented-deviation catalog is recorded as DEVIATION
    /// (an adjudicated expected-difference that counts toward the floor), NOT FAIL. A documented-deviation id
    /// that nonetheless PASSES is flagged STALE so the catalog stays honest.
    /// </summary>
    private sealed class Scoreboard
    {
        private readonly List<(string Id, string Category, RowState State, string Detail)> _rows = new();
        private readonly IReadOnlyDictionary<string, string> _deviations;

        public Scoreboard(IReadOnlyDictionary<string, string> documentedDeviations) =>
            _deviations = documentedDeviations;

        public int Pass { get; private set; }
        public int Deviation { get; private set; }
        public int Fail { get; private set; }
        public int Total => _rows.Count;
        public IEnumerable<string> FailingIds => _rows.Where(r => r.State == RowState.Fail).Select(r => r.Id);

        public void Score(string id, string category, Action body)
        {
            string? failDetail = null;
            try
            {
                body();
            }
            catch (SoftAssertException ex)
            {
                failDetail = ex.Message;
            }
            catch (Exception ex)
            {
                // An unexpected throw (e.g. the adapter blew up) is a FAIL with the exception type, not a
                // harness crash — the scoreboard measures it like any other failing case.
                failDetail = $"{ex.GetType().Name}: {ex.Message}";
            }

            if (failDetail is null)
            {
                // Passed. If it is ALSO listed as a documented deviation, that listing is now STALE — surface
                // it as a FAIL so the catalog gets pruned (a deviation must describe a real, current divergence).
                if (_deviations.ContainsKey(id))
                {
                    _rows.Add((id, category, RowState.Fail,
                        "STALE DEVIATION: this case now PASSES — remove it from DocumentedDeviations."));
                    Fail++;
                }
                else
                {
                    _rows.Add((id, category, RowState.Pass, ""));
                    Pass++;
                }
                return;
            }

            // Failed the count assert. A documented, adjudicated deviation counts toward the floor; anything
            // else is a real regression.
            if (_deviations.TryGetValue(id, out var reason))
            {
                _rows.Add((id, category, RowState.Deviation, $"{failDetail}  —  {reason}"));
                Deviation++;
            }
            else
            {
                _rows.Add((id, category, RowState.Fail, failDetail));
                Fail++;
            }
        }

        public void Report(ITestOutputHelper o)
        {
            o.WriteLine("===== IR PARITY SCOREBOARD (RUNNABLE_NOW cases) =====");
            o.WriteLine($"Total: {Total}   PASS: {Pass}   DEVIATION: {Deviation}   FAIL: {Fail}   " +
                        $"({100.0 * (Pass + Deviation) / Math.Max(1, Total):F1}% pass-or-deviation)");
            o.WriteLine("");
            foreach (var g in _rows.GroupBy(r => r.Category).OrderBy(g => g.Key))
                o.WriteLine($"  [{g.Key,-4}] {g.Count(r => r.State == RowState.Pass)} pass + " +
                            $"{g.Count(r => r.State == RowState.Deviation)} deviation / {g.Count()}");
            o.WriteLine("");
            o.WriteLine("FAILING cases (undocumented regressions — must be empty for the floor to hold):");
            foreach (var r in _rows.Where(r => r.State == RowState.Fail))
                o.WriteLine($"  FAIL  {r.Id,-60} {r.Detail}");
            o.WriteLine("");
            o.WriteLine("DOCUMENTED DEVIATIONS (PASS_WITH_DOCUMENTED_DEVIATION — visible, counts toward floor):");
            foreach (var r in _rows.Where(r => r.State == RowState.Deviation))
                o.WriteLine($"  DEV   {r.Id,-12} {r.Detail}");
            o.WriteLine("");
            o.WriteLine("PASSING cases:");
            foreach (var r in _rows.Where(r => r.State == RowState.Pass))
                o.WriteLine($"  PASS  {r.Id}");
        }
    }

    // ---------------------------------------------------------------------- minimal doc builders
    // Mirror the programmatic builders in WmlComparerMoveDetectionTests / WmlComparerFormatChangeTests so the
    // ported assertions run over byte-identical fixtures.

    private static WmlDocument Doc(params string[] paragraphs)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            AddDefaults(mainPart);
            doc.Save();
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    private static WmlDocument Para(string text) => Doc(text);

    private static WmlDocument BoldPara(string text)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new RunProperties(new Bold()), new Text(text)))));
            AddDefaults(mainPart);
            doc.Save();
        }
        return new WmlDocument("test.docx", stream.ToArray());
    }

    private static void AddDefaults(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.Styles = new Styles(new DocDefaults(
            new RunPropertiesDefault(new RunPropertiesBaseStyle(
                new RunFonts { Ascii = "Calibri" }, new FontSize { Val = "22" })),
            new ParagraphPropertiesDefault()));
        mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
    }

    // ---------------------------------------------------------------------- WC003 rows (exact copies)

    /// <summary>The 105 live WC003_Compare InlineData rows (id, left, right, expectedRevisionCount), copied
    /// verbatim from WmlComparerTests.cs.</summary>
    private static IEnumerable<(string Id, string Left, string Right, int Expected)> WC003_Compare_Rows() => new[]
    {
        ("WC-1000", "CA/CA001-Plain.docx", "CA/CA001-Plain-Mod.docx", 1),
        ("WC-1010", "WC/WC001-Digits.docx", "WC/WC001-Digits-Mod.docx", 4),
        ("WC-1020", "WC/WC001-Digits.docx", "WC/WC001-Digits-Deleted-Paragraph.docx", 1),
        ("WC-1030", "WC/WC001-Digits-Deleted-Paragraph.docx", "WC/WC001-Digits.docx", 1),
        ("WC-1040", "WC/WC002-Unmodified.docx", "WC/WC002-DiffInMiddle.docx", 2),
        ("WC-1050", "WC/WC002-Unmodified.docx", "WC/WC002-DiffAtBeginning.docx", 2),
        ("WC-1060", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtBeginning.docx", 1),
        ("WC-1070", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtBeginning.docx", 1),
        ("WC-1080", "WC/WC002-Unmodified.docx", "WC/WC002-InsertAtEnd.docx", 1),
        ("WC-1090", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteAtEnd.docx", 1),
        ("WC-1100", "WC/WC002-Unmodified.docx", "WC/WC002-DeleteInMiddle.docx", 1),
        ("WC-1110", "WC/WC002-Unmodified.docx", "WC/WC002-InsertInMiddle.docx", 1),
        ("WC-1120", "WC/WC002-DeleteInMiddle.docx", "WC/WC002-Unmodified.docx", 1),
        ("WC-1140", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Row.docx", 1),
        ("WC-1150", "WC/WC006-Table-Delete-Row.docx", "WC/WC006-Table.docx", 1),
        ("WC-1160", "WC/WC006-Table.docx", "WC/WC006-Table-Delete-Contests-of-Row.docx", 2),
        ("WC-1170", "WC/WC007-Unmodified.docx", "WC/WC007-Longest-At-End.docx", 2),
        ("WC-1180", "WC/WC007-Unmodified.docx", "WC/WC007-Deleted-at-Beginning-of-Para.docx", 1),
        ("WC-1190", "WC/WC007-Unmodified.docx", "WC/WC007-Moved-into-Table.docx", 2),
        ("WC-1200", "WC/WC009-Table-Unmodified.docx", "WC/WC009-Table-Cell-1-1-Mod.docx", 1),
        ("WC-1210", "WC/WC010-Para-Before-Table-Unmodified.docx", "WC/WC010-Para-Before-Table-Mod.docx", 3),
        ("WC-1220", "WC/WC011-Before.docx", "WC/WC011-After.docx", 2),
        ("WC-1230", "WC/WC012-Math-Before.docx", "WC/WC012-Math-After.docx", 2),
        ("WC-1240", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After.docx", 2),
        ("WC-1250", "WC/WC013-Image-Before.docx", "WC/WC013-Image-After2.docx", 2),
        ("WC-1260", "WC/WC013-Image-Before2.docx", "WC/WC013-Image-After2.docx", 2),
        ("WC-1270", "WC/WC014-SmartArt-Before.docx", "WC/WC014-SmartArt-After.docx", 2),
        ("WC-1280", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-After.docx", 2),
        ("WC-1310", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After.docx", 3),
        ("WC-1320", "WC/WC014-SmartArt-With-Image-Before.docx", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx", 1),
        ("WC-1330", "WC/WC015-Three-Paragraphs.docx", "WC/WC015-Three-Paragraphs-After.docx", 3),
        ("WC-1340", "WC/WC016-Para-Image-Para.docx", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx", 1),
        ("WC-1350", "WC/WC017-Image.docx", "WC/WC017-Image-After.docx", 3),
        ("WC-1360", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-1.docx", 2),
        ("WC-1370", "WC/WC018-Field-Simple-Before.docx", "WC/WC018-Field-Simple-After-2.docx", 3),
        ("WC-1380", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-1.docx", 3),
        ("WC-1390", "WC/WC019-Hyperlink-Before.docx", "WC/WC019-Hyperlink-After-2.docx", 5),
        ("WC-1400", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-1.docx", 3),
        ("WC-1410", "WC/WC020-FootNote-Before.docx", "WC/WC020-FootNote-After-2.docx", 5),
        ("WC-1420", "WC/WC021-Math-Before-1.docx", "WC/WC021-Math-After-1.docx", 9),
        ("WC-1430", "WC/WC021-Math-Before-2.docx", "WC/WC021-Math-After-2.docx", 6),
        ("WC-1440", "WC/WC022-Image-Math-Para-Before.docx", "WC/WC022-Image-Math-Para-After.docx", 10),
        ("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx", "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx", 7),
        ("WC-1460", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After.docx", 1),
        ("WC-1470", "WC/WC024-Table-Before.docx", "WC/WC024-Table-After2.docx", 7),
        ("WC-1480", "WC/WC025-Simple-Table-Before.docx", "WC/WC025-Simple-Table-After.docx", 4),
        ("WC-1500", "WC/WC026-Long-Table-Before.docx", "WC/WC026-Long-Table-After-1.docx", 2),
        ("WC-1510", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-1.docx", 2),
        ("WC-1520", "WC/WC027-Twenty-Paras-After-1.docx", "WC/WC027-Twenty-Paras-Before.docx", 2),
        ("WC-1530", "WC/WC027-Twenty-Paras-Before.docx", "WC/WC027-Twenty-Paras-After-2.docx", 4),
        ("WC-1540", "WC/WC030-Image-Math-Before.docx", "WC/WC030-Image-Math-After.docx", 2),
        ("WC-1550", "WC/WC031-Two-Maths-Before.docx", "WC/WC031-Two-Maths-After.docx", 4),
        ("WC-1560", "WC/WC032-Para-with-Para-Props.docx", "WC/WC032-Para-with-Para-Props-After.docx", 3),
        ("WC-1570", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After1.docx", 2),
        ("WC-1580", "WC/WC033-Merged-Cells-Before.docx", "WC/WC033-Merged-Cells-After2.docx", 4),
        ("WC-1600", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After1.docx", 1),
        ("WC-1610", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After2.docx", 4),
        ("WC-1620", "WC/WC034-Footnotes-Before.docx", "WC/WC034-Footnotes-After3.docx", 3),
        ("WC-1630", "WC/WC034-Footnotes-After3.docx", "WC/WC034-Footnotes-Before.docx", 3),
        ("WC-1640", "WC/WC035-Footnote-Before.docx", "WC/WC035-Footnote-After.docx", 2),
        ("WC-1650", "WC/WC035-Footnote-After.docx", "WC/WC035-Footnote-Before.docx", 2),
        ("WC-1660", "WC/WC036-Footnote-With-Table-Before.docx", "WC/WC036-Footnote-With-Table-After.docx", 5),
        ("WC-1670", "WC/WC036-Footnote-With-Table-After.docx", "WC/WC036-Footnote-With-Table-Before.docx", 5),
        ("WC-1680", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After1.docx", 1),
        ("WC-1700", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After2.docx", 4),
        ("WC-1710", "WC/WC034-Endnotes-Before.docx", "WC/WC034-Endnotes-After3.docx", 7),
        ("WC-1720", "WC/WC034-Endnotes-After3.docx", "WC/WC034-Endnotes-Before.docx", 7),
        ("WC-1730", "WC/WC035-Endnote-Before.docx", "WC/WC035-Endnote-After.docx", 2),
        ("WC-1740", "WC/WC035-Endnote-After.docx", "WC/WC035-Endnote-Before.docx", 2),
        ("WC-1750", "WC/WC036-Endnote-With-Table-Before.docx", "WC/WC036-Endnote-With-Table-After.docx", 6),
        ("WC-1760", "WC/WC036-Endnote-With-Table-After.docx", "WC/WC036-Endnote-With-Table-Before.docx", 6),
        ("WC-1770", "WC/WC037-Textbox-Before.docx", "WC/WC037-Textbox-After1.docx", 2),
        ("WC-1780", "WC/WC038-Document-With-BR-Before.docx", "WC/WC038-Document-With-BR-After.docx", 2),
        ("WC-1800", "RC/RC001-Before.docx", "RC/RC001-After1.docx", 2),
        ("WC-1810", "RC/RC002-Image.docx", "RC/RC002-Image-After1.docx", 1),
        ("WC-1820", "WC/WC039-Break-In-Row.docx", "WC/WC039-Break-In-Row-After1.docx", 1),
        ("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx", 2),
        ("WC-1840", "WC/WC042-Table-5.docx", "WC/WC042-Table-5-Mod.docx", 2),
        ("WC-1850", "WC/WC043-Nested-Table.docx", "WC/WC043-Nested-Table-Mod.docx", 2),
        ("WC-1860", "WC/WC044-Text-Box.docx", "WC/WC044-Text-Box-Mod.docx", 2),
        ("WC-1870", "WC/WC045-Text-Box.docx", "WC/WC045-Text-Box-Mod.docx", 2),
        ("WC-1880", "WC/WC046-Two-Text-Box.docx", "WC/WC046-Two-Text-Box-Mod.docx", 2),
        ("WC-1890", "WC/WC047-Two-Text-Box.docx", "WC/WC047-Two-Text-Box-Mod.docx", 2),
        ("WC-1900", "WC/WC048-Text-Box-in-Cell.docx", "WC/WC048-Text-Box-in-Cell-Mod.docx", 6),
        ("WC-1910", "WC/WC049-Text-Box-in-Cell.docx", "WC/WC049-Text-Box-in-Cell-Mod.docx", 5),
        ("WC-1920", "WC/WC050-Table-in-Text-Box.docx", "WC/WC050-Table-in-Text-Box-Mod.docx", 8),
        ("WC-1930", "WC/WC051-Table-in-Text-Box.docx", "WC/WC051-Table-in-Text-Box-Mod.docx", 9),
        ("WC-1940", "WC/WC052-SmartArt-Same.docx", "WC/WC052-SmartArt-Same-Mod.docx", 2),
        ("WC-1950", "WC/WC053-Text-in-Cell.docx", "WC/WC053-Text-in-Cell-Mod.docx", 2),
        ("WC-1960", "WC/WC054-Text-in-Cell.docx", "WC/WC054-Text-in-Cell-Mod.docx", 0),
        ("WC-1970", "WC/WC055-French.docx", "WC/WC055-French-Mod.docx", 0),
        ("WC-1980", "WC/WC056-French.docx", "WC/WC056-French-Mod.docx", 0),
        ("WC-1990", "WC/WC057-Table-Merged-Cell.docx", "WC/WC057-Table-Merged-Cell-Mod.docx", 4),
        ("WC-2000", "WC/WC058-Table-Merged-Cell.docx", "WC/WC058-Table-Merged-Cell-Mod.docx", 6),
        ("WC-2010", "WC/WC059-Footnote.docx", "WC/WC059-Footnote-Mod.docx", 5),
        ("WC-2020", "WC/WC060-Endnote.docx", "WC/WC060-Endnote-Mod.docx", 3),
        ("WC-2030", "WC/WC061-Style-Added.docx", "WC/WC061-Style-Added-Mod.docx", 1),
        ("WC-2040", "WC/WC062-New-Char-Style-Added.docx", "WC/WC062-New-Char-Style-Added-Mod.docx", 3),
        ("WC-2050", "WC/WC063-Footnote.docx", "WC/WC063-Footnote-Mod.docx", 1),
        ("WC-2060", "WC/WC063-Footnote-Mod.docx", "WC/WC063-Footnote.docx", 1),
        ("WC-2070", "WC/WC064-Footnote.docx", "WC/WC064-Footnote-Mod.docx", 0),
        ("WC-2080", "WC/WC065-Textbox.docx", "WC/WC065-Textbox-Mod.docx", 2),
        ("WC-2090", "WC/WC066-Textbox-Before-Ins.docx", "WC/WC066-Textbox-Before-Ins-Mod.docx", 1),
        ("WC-2092", "WC/WC066-Textbox-Before-Ins-Mod.docx", "WC/WC066-Textbox-Before-Ins.docx", 1),
    };

    /// <summary>The 56 live WC004_Compare_To_Self InlineData rows (id, file), copied verbatim from
    /// WmlComparerTests.cs. Self-compare must yield ZERO revisions.</summary>
    private static IEnumerable<(string Id, string Name)> WC004_CompareToSelf_Rows() => new[]
    {
        ("WCS-1000", "WC/WC001-Digits.docx"),
        ("WCS-1010", "WC/WC001-Digits-Deleted-Paragraph.docx"),
        ("WCS-1020", "WC/WC001-Digits-Mod.docx"),
        ("WCS-1030", "WC/WC002-DeleteAtBeginning.docx"),
        ("WCS-1040", "WC/WC002-DeleteAtEnd.docx"),
        ("WCS-1050", "WC/WC002-DeleteInMiddle.docx"),
        ("WCS-1060", "WC/WC002-DiffAtBeginning.docx"),
        ("WCS-1070", "WC/WC002-DiffInMiddle.docx"),
        ("WCS-1080", "WC/WC002-InsertAtBeginning.docx"),
        ("WCS-1090", "WC/WC002-InsertAtEnd.docx"),
        ("WCS-1100", "WC/WC002-InsertInMiddle.docx"),
        ("WCS-1110", "WC/WC002-Unmodified.docx"),
        ("WCS-1140", "WC/WC006-Table.docx"),
        ("WCS-1150", "WC/WC006-Table-Delete-Contests-of-Row.docx"),
        ("WCS-1160", "WC/WC006-Table-Delete-Row.docx"),
        ("WCS-1170", "WC/WC007-Deleted-at-Beginning-of-Para.docx"),
        ("WCS-1180", "WC/WC007-Longest-At-End.docx"),
        ("WCS-1190", "WC/WC007-Moved-into-Table.docx"),
        ("WCS-1200", "WC/WC007-Unmodified.docx"),
        ("WCS-1210", "WC/WC009-Table-Cell-1-1-Mod.docx"),
        ("WCS-1220", "WC/WC009-Table-Unmodified.docx"),
        ("WCS-1230", "WC/WC010-Para-Before-Table-Mod.docx"),
        ("WCS-1240", "WC/WC010-Para-Before-Table-Unmodified.docx"),
        ("WCS-1250", "WC/WC011-After.docx"),
        ("WCS-1260", "WC/WC011-Before.docx"),
        ("WCS-1270", "WC/WC012-Math-After.docx"),
        ("WCS-1280", "WC/WC012-Math-Before.docx"),
        ("WCS-1290", "WC/WC013-Image-After.docx"),
        ("WCS-1300", "WC/WC013-Image-After2.docx"),
        ("WCS-1310", "WC/WC013-Image-Before.docx"),
        ("WCS-1320", "WC/WC013-Image-Before2.docx"),
        ("WCS-1330", "WC/WC014-SmartArt-After.docx"),
        ("WCS-1340", "WC/WC014-SmartArt-Before.docx"),
        ("WCS-1350", "WC/WC014-SmartArt-With-Image-After.docx"),
        ("WCS-1360", "WC/WC014-SmartArt-With-Image-Before.docx"),
        ("WCS-1370", "WC/WC014-SmartArt-With-Image-Deleted-After.docx"),
        ("WCS-1380", "WC/WC014-SmartArt-With-Image-Deleted-After2.docx"),
        ("WCS-1390", "WC/WC015-Three-Paragraphs.docx"),
        ("WCS-1400", "WC/WC015-Three-Paragraphs-After.docx"),
        ("WCS-1410", "WC/WC016-Para-Image-Para.docx"),
        ("WCS-1420", "WC/WC016-Para-Image-Para-w-Deleted-Image.docx"),
        ("WCS-1430", "WC/WC017-Image.docx"),
        ("WCS-1440", "WC/WC017-Image-After.docx"),
        ("WCS-1450", "WC/WC018-Field-Simple-After-1.docx"),
        ("WCS-1460", "WC/WC018-Field-Simple-After-2.docx"),
        ("WCS-1470", "WC/WC018-Field-Simple-Before.docx"),
        ("WCS-1480", "WC/WC019-Hyperlink-After-1.docx"),
        ("WCS-1490", "WC/WC019-Hyperlink-After-2.docx"),
        ("WCS-1500", "WC/WC019-Hyperlink-Before.docx"),
        ("WCS-1510", "WC/WC020-FootNote-After-1.docx"),
        ("WCS-1520", "WC/WC020-FootNote-After-2.docx"),
        ("WCS-1530", "WC/WC020-FootNote-Before.docx"),
        ("WCS-1540", "WC/WC021-Math-After-1.docx"),
        ("WCS-1550", "WC/WC021-Math-Before-1.docx"),
        ("WCS-1560", "WC/WC022-Image-Math-Para-After.docx"),
        ("WCS-1570", "WC/WC022-Image-Math-Para-Before.docx"),
    };
}

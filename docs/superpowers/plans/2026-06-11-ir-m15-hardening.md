# Document IR — M1.5 Hardening (pre-Phase-2)

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Close the Phase 1 gate's deferred quality items before the diff engine builds on the IR: model textbox bodies (the dominant equivalence gap AND a `ContentHash` blind spot), make provenance retention optional (memory 11×→~2-4×), buy back perf headroom (1.90×→≤1.6×), and sweep the small remaining divergence fixtures.

**Baseline:** `feat/ir-m15-hardening` @ 5b6838d (Phase 1 merged; equivalence 608/668; perf 1.90×; memory ~11× XML retained).

**Targets:** equivalence ≥ 643/668 (textbox closes ~35) with only the 13 accepted oracle-bug divergences plus whatever the sweep can't close remaining; wall ≤ 1.6× oracle; memory ≤ 4× XML with provenance retained, ~2× with retention off. The 13 oracle-bug fixtures stay open BY DESIGN (closing them means changing shipped-converter output — bundled with cutover D3, not this milestone).

## Task 1: Textbox bodies in the IR

New inline node `IrTextbox(IrNodeList<IrBlock> Blocks)` for `w:txbxContent` found inside `w:drawing` (`wps:txbx`) and VML `w:pict`/`v:textbox`. Inner blocks walked by the normal block walker (own anchors, in AnchorIndex, ContentHash/FormatFingerprint as usual; depth-capped recursion — textboxes nest). ContentHash contribution: new sentinel `SentinelTextbox = 0x0B` + each inner block's ContentHash appended (spec §6.1 addendum — update spec). The drawing's image promotion (blip) takes precedence as today; a drawing with BOTH blip and textbox: pick the oracle-parity behavior (read the oracle). Diagnostic JSON branch + completeness guard keeps passing; snapshot regen reviewed (new textbox nodes; pre-existing anchors/hashes stable except fixtures that genuinely contain textboxes). Emitter: match the oracle's rendering/indexing of textbox inner paragraphs (read `BuildAnchorIndex`/emit walk precisely — descendants order); `ScopeHasContent` for header/footer scopes now sees textbox text → the 5 detection fixtures close. Corpus equivalence expected 608 → ~643.

## Task 2: Optional provenance retention (memory)

`IrReaderOptions.RetainSources` (default `true` — current behavior). When `false`: `IrDocument.Sources` empty, `IrProvenance.Element` null everywhere (PartUri facts survive — promote the part URI to `IrScope.PartUri` (additive) and use it in the emitter/index instead of per-block provenance). Re-measure memory with the existing M1.4 methodology (largest fixture): record retained-bytes ratios with retention on and off in the task report + gate-report addendum. Equivalence/corpus/suite must be unaffected in default mode; add a retention-off corpus smoke (read all 668 with RetainSources=false, totality + spot anchor/hash equality vs retained mode).

## Task 3: Perf pass

Profile first (any lightweight approach — Stopwatch段 sampling or dotnet-trace if available), THEN optimize the top items. Known suspects: `WmlToMarkdownConverter.KindFor`/`IsListItem` per-paragraph XML style-chain walks (careful: oracle-parity — any replacement must be proven output-identical corpus-wide), `ListItemRetriever` numbering resolution per paragraph, registry construction, canonicalization allocations. Rules: no equivalence regression (corpus stat must not drop), IR suite green, every optimization measured (before/after numbers in report). Target ≤1.6× oracle wall on the full corpus benchmark (DOCXODUS_RUN_PERF=1); update the perf test threshold only if comfortably under.

## Task 4: Small-fixture sweep + milestone close

Attempt the remaining non-accepted divergences: CC/revision spacing (4), heading-style-numPr display (2), other (~6: blank-line placement, TOC field text, heading ordering). Timebox per category; what doesn't fall cleanly gets an updated triage entry with findings. Close out: regenerate gate-report addendum (`## M1.5 Outcome` in this file + program-plan decision log note), CHANGELOG, final corpus/suite/Release verification.

## Out of scope

- The 13 accepted oracle-bug divergences (special-char drops, hyperlink splits) — D3/cutover work.
- Any public API surface; WASM/npm/python ripple (still none).

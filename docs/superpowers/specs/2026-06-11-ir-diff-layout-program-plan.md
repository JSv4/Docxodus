# Document IR / Diff Engine / Layout Engine — Program Plan

**Date:** 2026-06-11
**Status:** Approved direction. Phase 1 ready to start; Phase 3 explicitly deferred.
**Companion spec:** [`2026-06-11-document-ir-spec.md`](./2026-06-11-document-ir-spec.md) (detailed IR design)

## Vision

Today every major Docxodus module re-derives its own private, throwaway view of
the document from raw OOXML: `WmlComparer` builds `ComparisonUnitAtom` lists,
`WmlToMarkdownConverter` builds its projection, `OpenContractExporter` builds a
text/structure view, `FormattingAssembler` resolves the style cascade
destructively on the XML. Each re-implements normalization with subtly
different rules, which is where "module A and module B disagree about whether
these paragraphs are equal" bugs come from.

This program replaces those private views with **one shared semantic IR**
(intermediate representation) and then builds the next-generation consumers on
top of it:

```
                         ┌────────────────────────────┐
        OOXML (.docx) ──►│   Document IR (Phase 1)    │
                         │  typed, normalized,        │
                         │  anchor-identified,        │
                         │  immutable snapshot        │
                         └─────┬──────┬──────┬────────┘
                               │      │      │
              ┌────────────────┘      │      └───────────────────┐
              ▼                       ▼                          ▼
   Markdown projection        Diff engine (Phase 2)      Layout engine (Phase 3,
   (ported consumer,          edit script keyed by       DEFERRED) box tree keyed
   validates the IR)          anchors → renderers:       by anchors → paginated
                              native OOXML revisions,    browser rendering
                              revisions JSON, HTML
```

The strategic payoffs, in order of importance:

1. **Diff-as-data.** The current `WmlComparer` has no intermediate
   representation — the mutated document *is* the diff. A diff engine built on
   the IR emits a first-class edit script addressed by the same anchors that
   `DocxSession` and the markdown projection already use, which is the missing
   read-side complement to the agentic editing pipeline.
2. **Moves and format changes in the alignment**, not bolted on as post-hoc
   Jaccard/equal-atom passes.
3. **Determinism and thread safety by construction** (immutable snapshots; no
   `s_MaxId`-style static state).
4. **A foundation the deferred layout engine can consume** without a third
   rewrite of document parsing.

## Operating principles

- **Strangler, not rewrite-in-place.** Existing modules (`WmlComparer`,
  `WmlToHtmlConverter`, the browser pagination stack) keep working untouched
  the entire time. New engines ship behind settings/flags with the old path as
  default until burn-in completes.
- **Oracle-driven development.** Every phase has a machine-checkable
  correctness oracle before feature work starts: golden snapshots over
  `TestFiles/` for the IR; differential testing against `WmlComparer` plus the
  accept/reject round-trip invariant for the diff engine.
- **Every milestone ships working, testable software.** No milestone ends in
  "scaffolding done, nothing runs."
- **Plans are layered.** This document is the program plan. At the start of
  each milestone, a bite-sized task-level implementation plan (per
  `superpowers:writing-plans` format) is authored into
  `docs/superpowers/plans/` from this document plus the IR spec.
- **Follow CLAUDE.md feature workflow** per milestone: CHANGELOG entries,
  architecture docs, and the four-layer ripple (WASM / npm / python) — noting
  that Phase 1 is `internal` and has **no** cross-layer ripple at all.

---

## Phase 1 — Document IR (~6–8 weeks)

**Goal:** an immutable, typed, anchor-identified, normalized in-memory model
of a DOCX, validated by porting the markdown projection onto it with
output-identical results over the corpus.

**Scope guardrails:** read-only (no IR→OOXML writer), `internal` visibility
(no public API, no WASM/npm/python ripple), lossy-tolerant (unmodeled content
becomes `Opaque` nodes — see spec §4.4).

### M1.1 — IR core types + reader skeleton (week 1–2)

- `Docxodus/Ir/` namespace with the type model from the spec: `IrDocument`,
  scopes, `IrParagraph`/`IrTable`/`IrOpaqueBlock`, inline nodes, format
  records, `IrAnchor`, `IrHash`.
- `IrReader.Read(WmlDocument, IrReaderOptions)` covering body paragraphs,
  runs, tables, breaks/tabs, with everything else landing as `Opaque` nodes
  (correct-by-construction fallback, not an error).
- Diagnostic JSON serialization (`ToDiagnosticJson()`) — the substrate for
  snapshot tests.
- **Exit:** reader runs over every file in `TestFiles/` without throwing;
  golden-snapshot test infrastructure in place with initial snapshots
  committed.

### M1.2 — Normalization + hashing (week 2–3)

- Implement normalization rules N1–N15 from the spec (rsid stripping, run
  coalescing, field handling, revision view, etc.).
- `ContentHash` / `FormatFingerprint` / opaque canonical hashing per spec §6,
  including the unmodeled-format digest.
- Invariant tests: hash stability across re-reads of the same bytes;
  hash sensitivity tests (change one char → content hash changes, bold a run →
  fingerprint changes, content hash doesn't).
- **Exit:** documented equality semantics (the spec's normalization table is
  the source of truth) with a test per rule.

### M1.3 — Effective formatting + registries (week 3–5)

- Style registry, numbering registry, theme font resolution; lazy
  cascade-resolved `EffectiveParaFormat`/`EffectiveRunFormat` (reusing
  `FormattingAssembler` logic non-destructively, not calling it).
- List facts (`numId`/`abstractNumId`/`ilvl`/format/start-override/from-style)
  matching what `GetListMembership` reports today — assert parity in tests.
- Remaining scopes: headers, footers, footnotes, endnotes, comments store.
- **Exit:** effective-format parity spot-checks against
  `FormattingAssembler` output on corpus fixtures.

### M1.4 — Markdown projection port (week 5–8) — **PHASE GATE**

- Reimplement `WmlToMarkdownConverter` as an IR consumer (new internal code
  path; the shipped converter is untouched until the gate passes).
- Run both implementations over the full `TestFiles/` corpus and diff outputs.
- Triage every difference: IR bug, old-converter bug (file an issue, accept
  the new output), or intentional. Goal is byte-identical modulo triaged
  accepted diffs.
- Anchors emitted by the IR path must be **identical** to the current
  projection's anchors (same deterministic Unid pipeline) — this is
  non-negotiable, since `DocxSession` clients hold these ids.
- **Phase exit criteria:**
  1. Projection equivalence across the corpus (triaged).
  2. Perf budget: IR build + projection ≤ 2× current converter wall time on
     the corpus; memory ≤ 3× the document XML size.
  3. Architecture doc `docs/architecture/document_ir.md` written.
  4. Decision recorded: cut the shipped converter over to the IR path now, or
     defer the cutover to ride along with Phase 2.

### Stretch (only if M1.4 lands early)

- Port `OpenContractExporter`'s text-extraction layer onto the IR as a second
  consumer proof.

### Phase 1 risks

| Risk | Mitigation |
|---|---|
| IR shape doesn't fit real consumers (abstraction built on speculation) | M1.4 port is *in the phase*, not after it — the gate forces the fit test |
| Anchor drift vs shipped projection | Reuse the exact Unid assignment code path; parity assertion in CI |
| Normalization disagreements surface as snapshot churn | Every rule numbered in the spec, one test per rule, snapshot diffs reviewed not regenerated blindly |

---

## Phase 2 — Diff engine (~3–4 months)

**Goal:** a from-scratch comparison engine on the IR producing a first-class
edit script, with renderers for (a) native OOXML tracked-changes markup,
(b) the `GetRevisions()`-style JSON surface, shipped behind a setting with
`WmlComparer` remaining the default.

**Out of scope for v1:** `Consolidate()` (multi-reviewer merge), textbox-body
diffing, header/footer diffing (compare body + footnotes/endnotes first;
opaque-hash everything else so it still reports changed/unchanged correctly).

### M2.1 — Tokenization + block alignment (month 1)

- Diff-time tokenizer: IR runs → word tokens (honoring `WordSeparators`,
  culture, case folding as *diff settings*, not IR facts).
- Block-level alignment over `ContentHash`/`FormatFingerprint` pairs using
  unique-hash anchoring (histogram-diff style) with **move detection
  integrated into alignment** — a block appearing once on each side in
  different positions is a move candidate by construction, no Jaccard pass.
- **Exit:** alignment unit tests incl. adversarial fixtures (500 near-identical
  paragraphs; boilerplate-heavy contracts) with complexity assertions.

### M2.2 — Edit script + intra-block diff (month 1–2)

- `IrEditScript`: ordered operations (insert/delete/equal/move/format-change)
  addressed by anchor + token span, with move pairs linked and **re-diffing
  within matched move pairs** (moved-and-edited renders as move + nested
  edits — the case the current engine structurally cannot express).
- Format-change detection falls out of the `ContentHash`-equal /
  `FormatFingerprint`-different case plus token-level fingerprint comparison.
- **Exit:** edit-script invariants (apply(script, a) reconstructs b's IR;
  script round-trips through JSON).

### M2.3 — Revisions surface + differential harness (month 2)

- Renderer: edit script → `WmlComparerRevision`-shaped output.
- **Differential harness:** run old and new engines over the corpus and over a
  generative fuzzer (random paragraph/table/run mutations); compare revision
  sets semantically; triage every divergence.
- **Exit:** divergence rate quantified and triaged; fuzzer in CI.

### M2.4 — Native OOXML revision renderer (month 2–4) — **GO/NO-GO GATE**

This is the program's risk concentration: emitting `w:ins`/`w:del`/
`w:moveFrom`/`w:moveTo`/`w:rPrChange` markup that Word opens without repair
and that round-trips. Deleted paragraph marks, move range elements crossing
block boundaries, table row/cell revisions, footnote refs inside deleted runs,
numbering on inserted paragraphs, trailing `sectPr`.

- Build against the **accept/reject invariant fuzz harness** from day one:
  `accept(compare(a,b)) ≈ normalize(b)`, `reject(compare(a,b)) ≈ normalize(a)`,
  checked via IR hashes.
- **Gate criterion (per the standing USER DIRECTIVE) is now THE SCOREBOARD AT
  100%.** M2.3 Task 4 built `IrParityScoreboardTests` (Trait `Category=Parity`),
  the definitive measurement of WmlComparer-suite parity. M2.4's job is to drive
  it to 100% by (a) closing the RUNNABLE_NOW failures (currently 129/179 = 72.1%
  pass — granularity, footnote/endnote/textbox scope gaps, the exact-move-via-
  anchoring divergence; see the [M2.3 Outcome scoreboard](../plans/2026-06-11-diff-m23-revisions-surface-differential.md#m23-outcome))
  and (b) emitting native OOXML markup so the MARKUP_BLOCKED rows (the A/B
  categories — produced-document validation, accept/reject round-trip, native
  `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` elements, revision-id
  uniqueness) become portable and pass. As markup lands, those rows move from the
  MARKUP_BLOCKED bucket into the scoreboard and must pass too. The accept/reject
  invariant fuzzer (above) is the oracle for the round-trip rows.
- **Gate (held ~6 weeks in):** if the scoreboard is not on a clear trajectory to
  100% by then, stop and re-scope — fall back to shipping the edit
  script + revisions surface only (still valuable for the agentic pipeline)
  and keep `WmlComparer` for document production.

### M2.5 — Productization (month 4)

- Public surface (working name `DocxDiff`) + `WmlComparerSettings`-style
  options; old engine remains the default; new engine opt-in via setting.
- Four-layer ripple per CLAUDE.md: `DocxSessionOps` (if session-exposed),
  WASM bridge, `npm/src/types.ts` + `index.ts`, python host + `docx_scalpel`.
- Manual validation: open outputs in Word and LibreOffice; accept-all /
  reject-all by hand on a sample; redline a real contract set.
- Docs: `docs/architecture/ir_diff_engine.md`; update
  `wml_comparer_gaps.md` (already stale — fix the move-markup/format-change
  claims while there) and `comparison_engine.md`.

### Phase 2 risks

| Risk | Mitigation |
|---|---|
| Word rejects/repairs generated revision markup (unknown unknowns) | M2.4 gate; invariant fuzzer from day one; existing test suite as spec; LibreOffice + Word manual passes |
| Semantic divergence from `WmlComparer` interpreted as regression by users | Differential harness with explicit triage log; opt-in flag; document intentional behavior differences |
| Scope creep into `Consolidate`/textboxes | Explicitly out of v1; opaque-hash fallback keeps correctness ("changed") without depth |

---

## Phase 3 — Layout engine (DEFERRED — sketch only)

Not scheduled. Entry criteria to revisit: (a) Phase 1 IR hardened by two
consumers in production, (b) Phase 2 shipped, (c) confirmed product need for
page-faithful rendering (e.g. page-number citations, paginated redline view)
that the current browser-measured pagination (`npm/src/pagination.ts`) cannot
meet.

Direction of record when revisited:

- Layout core in C# consuming the IR, compiled to WASM (shared with
  server-side rendering), running in the existing worker.
- Output: box tree keyed by IR anchors → absolutely-positioned DOM (selection,
  a11y, annotations nearly free); canvas later only for thumbnails/virtualized
  scroll.
- Font metrics via metric-compatible substitutes (Carlito/Caladea/Liberation)
  + DOCX-embedded fonts; shaping via HarfBuzz-WASM or OpenType metric parsing.
- Oracle: LibreOffice headless reference renders over the corpus; compare page
  breaks/line counts automatically.
- Honest fidelity target: plausible, stable pagination — **not** pixel parity
  with Word (compat-flag line breaking makes that a non-goal).
- Tier 3 (editing surface) remains out of scope permanently;
  editing stays "DocxSession mutates, engine re-lays-out."

---

## Cross-cutting workstreams

- **Corpus & oracle infra (starts in Phase 1):** corpus runner over
  `TestFiles/`, golden-snapshot tooling with reviewed regeneration, generative
  DOCX mutator (shared by IR hash tests and the Phase 2 fuzzer).
- **Determinism:** IR build and diff output must be byte-stable for identical
  inputs (fixed revision dates behind a `Deterministic` option; ids assigned
  in a single stable pass). No static mutable state anywhere in new code.
- **Performance:** budgets asserted in tests from M1.4 on; BenchmarkDotNet
  project added at M2.1 with the adversarial fixtures.
- **Code standards:** all new files `#nullable enable`; must compile under
  `WASM_BUILD` (no SkiaSharp references); Release build is warnings-as-errors.

## Decision log / gates

| # | Decision/gate | When | Status |
|---|---|---|---|
| D1 | IR first, diff second, layout deferred | 2026-06-10 | **Decided** |
| D2 | IR is `internal`, read-only, lossy-tolerant | 2026-06-11 (spec) | **Decided** |
| G1 | Phase 1 gate: projection equivalence + perf budget | end M1.4 | **PASSED** (2026-06-11) — 608/668 byte-equal corpus + fully-triaged remainder (accepted oracle-bug diffs / deferred textbox work); perf 1.90× ≤ 2.0× budget; memory ≈11× XML retained on the largest-body fixture (reported, not gated); arch doc `docs/architecture/document_ir.md` written. Full report in the [M1.4 plan Outcome section](../plans/2026-06-11-ir-m14-markdown-projection-port.md#outcome-phase-1-gate-report). |
| M1.5 | Pre-Phase-2 hardening (textboxes, memory, perf, sweep) | post-G1 | **COMPLETE** (2026-06-11) — equivalence **608 → 648/668** (textbox bodies modeled, closing the `ContentHash` blind spot + the dominant gap; sweep closed heading-numPr layout + unterminated-field/TOC result); perf **1.90× → 1.16×** (gate tightened 2.0× → 1.5×); memory **11.08× → 2.73×** XML with retention off (`RetainSources`); revision-skip scan made provably sound (all parts + complete element set + set-drift guard). The residual 20 divergences are all accepted oracle-bug-family (closing them = changing shipped-converter output, bundled with D3). **Phase 2 entry criteria met:** IR is textbox-complete, hash-sound, memory/perf within budget, equivalence fully triaged. Full report in the [M1.5 plan Outcome section](../plans/2026-06-11-ir-m15-hardening.md#m15-outcome). |
| D3 | Cut shipped markdown converter over to IR path | at G1 | **DEFERRED to Phase 2 / M2.5** (2026-06-11) — emitter parity gaps are exactly the accepted-divergence set (special-char drops, hyperlink/run splits, customXml-range CC acceptance — all cases where the IR is *more* correct), and cutover buys nothing until the diff engine exists; keep oracle as the shipped path and the IR path as a CI-validated alternative. Revisit at M2.5. |
| M2.1 | Tokenization + block alignment (Phase 2 open) | month 1 | **COMPLETE** (2026-06-11) — `Docxodus/Ir/Diff/`: `IrDiffSettings`/`IrDiffTokenizer` (word/separator/atomic tokens, hyperlink-target-in-key, field transparency) + `IrBlockAligner` (unique-hash `(ContentHash,FormatFingerprint)` anchoring → LIS spine → in-order gap fill; moves fall out off-spine by construction, no Jaccard pass; `MovedModified` reserved for M2.2). **Exit criteria met:** unit tests (tokenizer + 18 aligner cases) plus a **92-pair WC corpus smoke** (161/163 files; invariants hold forward AND reversed, per-pair kind histograms logged), adversarial fixtures (500 near-identical → 499 Unchanged + 1 Modified, 0 Moved; 500 identical − 1 → 499 Unchanged + 1 Deleted, 0 Moved/Modified; 200×200 full rewrite → 200 Modified no throw; contiguous 10-of-300 block move → exactly 10 Moved, LIS drops the smaller side off the spine as designed), and a **scale guard** (500→2000 para = 1.4→6.6 ms, **4.7× for 4× input** ≤ 8× anti-O(n²) bound). Full report in the [M2.1 plan Outcome section](../plans/2026-06-11-diff-m21-tokenizer-block-alignment.md#m21-outcome). Carried to M2.2: cross-gap move+edit → Del+Ins (exact-hash only), MovedModified, row/cell-granular table alignment, similarity gap pairing. |
| M2.2 | Edit script + intra-block diff (token diff, fuzzy moves, table granularity) | month 1–2 | **COMPLETE** (2026-06-11) — `Docxodus/Ir/Diff/`: `IrTokenDiffer` (Myers O(ND), Equal/Insert/Delete/FormatChanged), the anchor-addressed `IrEditScript`/`IrEditScriptBuilder`/`IrEditScriptJson` (EqualBlock/FormatOnlyBlock/ModifyBlock/Insert/Delete/MoveBlock/MoveModifyBlock; moves as source+destination pairs), similarity-based in-gap pairing + cross-gap fuzzy moves (`MovedModified` now reachable — relocated-AND-edited as move + nested edits, the case `WmlComparer` cannot express), nested **table row/cell diffs** (`IrTableDiff`/`IrRowOp`/`IrCellOp` via `IrTableDiffer` — a cell-text edit surfaces as a token diff inside that cell, not a whole-table blob), and the **`FormatComparison = ModeledOnly (default) | Full`** diff-time policy resolving the M2.1 FormatFingerprint run-boundary-noise finding (boundary-normalized modeled-only signature; no IR snapshot churn). **Exit criteria met:** apply-verification (apply(script,left) reconstructs right at text level, incl. nested table reconstruction + row/cell anchor validation) + JSON round-trip green over all synthetic cases AND the full 92-pair WC corpus both directions; 102 `Ir.Diff` tests. Final corpus (forward): Unchanged=2220, FormatOnly=50 (was 1714 — noise collapsed), Modified=1419, Moved=3, Inserted=970, Deleted=104; 18 tables → nested diffs (53 row ops, 25 cells with token diffs). Full report in the [M2.2 plan Outcome section](../plans/2026-06-11-diff-m22-edit-script.md#m22-outcome). Carried to M2.3+: grid-aware cell pairing (gridSpan/vMerge), fuzzy row moves, the revisions-API renderer + differential harness. |
| M2.3 | Revisions surface + differential harness + parity scoreboard | month 2 | **COMPLETE** (2026-06-11) — Task 1 `IrRevisionRenderer` (`WmlComparerRevision`-shaped output off the edit script: Inserted/Deleted/Moved/FormatChanged with text/author/date/MoveGroupId/IsMoveSource/FormatChange details); Task 2 differential harness `IrVsWmlComparerTests` (both engines over the 92-pair WC corpus × 2 directions, semantic combined-char-bag comparison, MATCH/GRANULARITY/DIVERGENT triage with 8 mechanical cause buckets — dominant DIVERGENT cause is `ScopeGapNewEmpty`: footnote/endnote/textbox edits the IR body-only path doesn't reach; zero NEW_ERROR); Task 3 deterministic seeded fuzzer `IrDiffFuzzTests` (own-oracle alignment+apply+JSON invariants always; cross-engine differential for the comparable mutation classes — 50/500-seed runs green, zero new-empty regressions). **Task 4 — the USER-DIRECTIVE deliverable — the PARITY SCOREBOARD:** inventoried all 8 `WmlComparer*` test files (live cases: WmlComparerTests 246 InlineData [WC003 105 / WC004 56 / WC005 1 + WC001/WC002 Consolidate 84], WmlComparerTests2 4 [CZ001 only — rest `#if false`], MoveDetection 32 [14 GetRevisions / 16 markup / 2 settings-default], FormatChange 13, LegalNumbering 5, BodyLevelElements 5, BodyLevelBookmark 1, ParallelRace 1). Built `IrWmlComparerAdapter` (test-side `GetRevisions` over the IR pipeline, `WmlComparerSettings → IrDiffSettings` mapping) + `IrParityScoreboardTests` (Trait `Category=Parity`, soft-assert per case → PASS/FAIL table + totals, asserts only totality). **Scoreboard baseline: 179 RUNNABLE_NOW cases ported, 129 PASS / 50 FAIL = 72.1%** (C 113/161, C+G 1/1, D 12/14, E 3/3). Failures: WC003 count granularity (off-by-1/2 TokenSpanGranularity) + footnote/endnote/textbox scope gaps (got 0) + 2 over-report table-cell cases; 2 move cases (exact relocation still caught by aligner anchoring under `DetectMoves=false`/below-min). Full scoreboard + burn-down in the [M2.3 plan Outcome section](../plans/2026-06-11-diff-m23-revisions-surface-differential.md#m23-outcome). 130 `Ir.Diff` tests; Release green. **M2.4 gate is now THE SCOREBOARD AT 100%.** |
| G2 | Phase 2 go/no-go: native markup renderer viability | ~6 wks into M2.4 | Open |
| D4 | New diff engine becomes default | post-M2.5 burn-in | Open |

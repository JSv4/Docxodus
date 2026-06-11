# Document IR — M1.4 Markdown Projection Port (PHASE 1 GATE)

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Reimplement the markdown projection as an IR consumer (`IrMarkdownEmitter`, internal, shipped converter untouched) and prove equivalence: corpus-wide output equality vs `WmlToMarkdownConverter` (markdown string + anchor index), anchor parity, and the perf budget. This is the Phase 1 exit gate from the program plan.

**Baseline:** `feat/document-ir` @ da2a34d — IR complete through M1.3 (all scopes, registries, list facts, effective formats), IR suite 135 green, corpus 668/668.

**Gate criteria (program plan M1.4):**
1. Projection equivalence across `TestFiles/` (byte-identical modulo controller-triaged accepted diffs).
2. Perf: IR build + emit ≤ 2× shipped converter wall time corpus-wide; memory ≤ 3× document XML size.
3. `docs/architecture/document_ir.md` written.
4. Cutover decision recorded (cut the shipped converter over now vs. defer to Phase 2).

**Method rules:**
- The emitter consumes the IR. When the IR lacks a fact the projection needs, PREFER extending the IR additively (new field, documented); peeking at `Source` provenance is a last resort and every instance must be reported — each one is evidence of IR incompleteness for the gate report.
- The shipped `WmlToMarkdownConverter` is the oracle and must remain byte-untouched (except already-landed visibility changes).
- Equivalence harness reports per-fixture stats and writes differing samples to an artifacts dir for controller triage; it is the loop driver, not just a pass/fail gate.

## Task 1: Emitter scaffold + equivalence harness + simple-fixture equality

`Docxodus/Ir/IrMarkdownEmitter.cs`: `internal static class` with `Emit(IrDocument, WmlToMarkdownConverterSettings) → MarkdownProjection`-shaped result (markdown string + anchor index entries). Port the body-paragraph path: headings/plain/list-item lines with `{#kind:scope:unid}` anchors, inline formatting subset (bold/italic/links/etc. — read EmitMarkdown for the exact rules incl. escaping), default settings only (FullUnid rendering, default EmptyParagraphs). Harness (`IrMarkdownEquivalenceTests.cs`, Trait "Corpus"): run shipped converter + IR path over every TestFiles fixture; compare markdown + anchor-index; write mismatches to `Docxodus.Tests/Ir/EquivalenceArtifacts/` (gitignored); assert-equal on a curated must-pass list that grows per task; report corpus-wide stats via ITestOutputHelper. Exit: simple paragraph/heading/list fixtures equal; stats baseline recorded.

## Task 2: Tables, images, opaque blocks, section breaks, settings modes

GFM-vs-opaque table rendering, image lines (`![alt](docxodus://img/…){#img:…}` — extend IrInlineImage with whatever id the URL scheme needs if missing), opaque anchor blocks with fenced summaries, `sec` thematic breaks, EmptyParagraphs modes, AnchorIdRendering modes (FullUnid/abbreviated/sequential — port the AnchorIdMap behavior). Exit: corpus equal-count strictly increasing; all table/image fixtures on the must-pass list.

## Task 3: Multipart + auto-number prefixes + corpus closure

Headers/footers/footnotes/endnotes/comments sections (multipart namespacing headings), auto-number prefix computation (numbering counters — port the projection's counter walk onto IR list facts), boilerplate-note parity, TextPreview/AnchorTarget index fields. Drive the corpus to full equivalence; every remaining diff goes into a triage table (fixture, category, root cause, proposed disposition) in the task report for CONTROLLER adjudication — do not self-accept diffs. Exit: 100% equal OR complete triage table.

## Task 4: Perf budget + architecture doc + gate report

BenchmarkDotNet not required: a corpus-wall-time comparison test (Trait "Perf", tolerant threshold ≤2×, ITestOutputHelper numbers) + a memory spot-check on the largest fixture (≤3× XML size). `docs/architecture/document_ir.md`: IR overview, type model, normalization table pointer, hashing, registries, effective formats, scopes, the emitter, evolution policy — written for the repo's architecture-doc conventions. Gate report appended to this plan as `## Outcome`: criteria pass/fail, provenance-peek inventory, triage decisions, cutover recommendation for D3 (controller decides).

## Out of scope

- Actually cutting `WmlToMarkdownConverter.Convert` over to the IR path (that's decision D3, recorded not executed).
- Public API changes, WASM/npm/python ripple (none — everything stays internal).

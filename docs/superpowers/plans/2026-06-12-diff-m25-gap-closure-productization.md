# Diff Engine — M2.5 Final Gap Closure + Productization

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** Close the five M2.4b-deferred gaps (targeting deviations 5→0 and allowlist 5→0, or evidence-retained per the method rule), then productize: public API surface, D3/D4 decisions, docs, cross-layer ripple.

**USER METHOD RULE (still binding):** WmlComparer presumed correct per gap; IR fixed to match unless oracle fault established with concrete evidence. Fine-mode grain stays available; engine fixes that improve correctness are allowed; render-time policies preferred where they suffice. NO merges to main; LOCAL ONLY; granular commits; updates to the user per workstream.

**Baseline:** `feat/diff-m24` @ a5e3f3f — GetRevisions 174 PASS + 5 DEV (floors 179/174), markup floor 39, allowlist 5 fixtures, full suite 1935/0.

## Task 1: Note-ref-within-word tokenization (WC-1710/1720 + markup WC034 ×2)

The established evidence (M2.4b): a note ref relocated INTO a word (`Vi[ref]deo` vs `Video [ref]`) genuinely changes the word's structure; the oracle reports del+ins; the IR under-reports. Fix at the tokenizer/differ level so a ref-split word is NOT equal to its contiguous form — design options to evaluate (pick with evidence, measure blast radius corpus-wide): (a) word tokens carry an adjacent-ref-boundary marker in MatchKey when a NoteRef interrupts mid-word (no separator between text runs around the ref); (b) the ref atom's MatchKey position-sensitivity. Constraint: a ref BETWEEN words (the common case) must stay non-disruptive — only intra-word interruption changes keys. Verify: WC-1710/1720 → genuine PASS; WC034 markup pair round-trips clean (allowlist −2); corpus/differential/fuzz no regressions; deviations catalog updated.

## Task 2: Sub-paragraph grain (WC-1450 anchor ambiguity, WC-1830 content migration)

- WC-1450: two IDENTICAL cell paragraphs; the aligner anchors the "wrong" one (+1 vs oracle). Investigate: is a deterministic positional preference (prefer the pairing preserving relative order among equal-content candidates) achievable in the anchor/gap machinery without breaking the boilerplate guarantees? If yes, engine fix; if it requires arbitrary tie-breaking the oracle gets right only by accident of its LCS order — evidence + retain.
- WC-1830: one before-paragraph's content migrates across TWO after-paragraphs (split). The oracle's atom-level LCS handles splits naturally; the IR's block model pairs 1:1. Investigate a bounded split/merge detection in gap fill (left block's token multiset ≈ union of 2 adjacent right blocks → render as the oracle does). Timebox honestly: if the principled fix is a real sub-paragraph alignment model, sketch the design, retain with evidence, and propose as a Phase-2-follow-on item rather than hacking.

## Task 3: Remaining markup leftovers (WC-1920, WC019 residual, WC-BodyBookmarks)

- WC-1920 punctuation-attachment grain inside textbox-nested table: diagnose precisely; likely related to separator/punctuation tokenization vs the oracle's atomization — fix at the narrowest correct layer.
- WC019 residual: rejecting `w:del`/`w:ins` nested inside `w:hyperlink` — the oracle sidesteps by stripping hyperlinks pre-compare. Options: (a) markup renderer emits hyperlink-edit markup the way the ORACLE's output shape does (mirror its produced shape — study WmlComparer's output for hyperlink edits); (b) extend shared RevisionProcessor (touches the oracle's accept path — only with extreme care + full old-engine suite + explicit report flagging). Prefer (a) per the presumption rule.
- WC-BodyBookmarks: endnote→footnote note-store conversion (a note CHANGES KIND between docs). Investigate what the oracle produces; match or evidence.

## Task 4: Productization — public surface + decisions + docs

- Public API (the first public surface of the program — naming decision: propose `DocxDiff` static facade + `DocxDiffSettings`/`DocxDiffRevision` etc., wrapping the internal pipeline; old `WmlComparer` API untouched and remains the DEFAULT engine — D4: recommend old-default until burn-in, IR engine opt-in). Surface: Compare(left,right,settings)→WmlDocument (native markup), GetRevisions(left,right,settings), GetEditScriptJson(left,right,settings) (the diff-as-data differentiator). `#nullable enable`, XML-doc'd, multi-author-friendly (consolidate-compatible). Record D3 (markdown cutover) + D4 recommendations in the program plan decision log (controller/user ratify).
- `docs/architecture/ir_diff_engine.md` (repo conventions); update `wml_comparer_gaps.md` (stale claims about move markup etc. — fix while there per the old plan note); CHANGELOG consolidation.

## Task 5: Cross-layer ripple (per CLAUDE.md table — new public surface)

WASM bridge (`DocxodusWasm` JSExport for the new compare surface), npm (`src/types.ts` + `index.ts` wrapper), python host dispatcher + `docx_scalpel` types/session per `python_docxodus.md` patterns — SCOPED MINIMAL: the three Compare/GetRevisions/EditScriptJson entry points only; follow each layer's established patterns; build all layers (`build-wasm.sh`, `npm run build`, tsc) + smoke tests where each layer's test harness allows cheaply. If python wheel infra makes the python slice impractical in-session, deliver the dispatcher + types and document the remainder.

## Exit criteria

Deviations + allowlist each → 0 or evidence-retained (verdict per row); ratchets raised to final values; public surface shipped with docs + ripple; D3/D4 recorded; full verification (suite, corpus, fuzz, projection equivalence, Release, WASM build) green; everything on the feature branch.

## M2.5 Outcome

**Status: COMPLETE (2026-06-12).** All five tasks landed on `feat/diff-m24`, local only, no merges to main.

### Per-task summary

- **T1 — note-ref-within-word tokenization.** Intra-word note-ref interruption (`Vi`⟨ref⟩`deo` ≠ contiguous `Video`) modeled as a genuine word-structure change in `IrDiffTokenizer` (sentinel-framed interruption marker in the flanking words' `MatchKey`), paired with note-store reference-order correspondence in `IrEditScriptBuilder` (align notes by body-reference order, not raw `w:id`) and a structural-only affix-trim guard. **WC-1710/1720 + WC-1620/1630 genuine PASSes; genuine-pass ratchet → 176.** A ref between words (the common case) is byte-untouched.
- **T2 — sub-paragraph grain.** WC-1450 AND WC-1830 re-diagnosed as ONE root cause: a 1:N paragraph SPLIT (one before-paragraph's content migrates across two after-paragraphs). PROVED not closable in the strictly-1:1 `IrEditOp` model and not render-coalescible. Sketched as the engine-level `IrSplitBlockOp`/`IrMergeBlockOp` capability and **deferred to M2.6** (design sketch `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md`); both retained as evidence-backed deviations, floors unchanged.
- **T3 — markup leftovers.** Affix-trim word boundary now mirrors `WmlComparer.GetComparisonUnitList` exactly (`IsOracleSplitChar`) — **WC-1920 genuine PASS, ratchet → 177.** `RevisionProcessor` reject reversal rules extended to `w:hyperlink` parent + empty-hyperlink-shell drop — **WC019 closed, round-trip allowlist 5→4.**
- **T4 — public surface + decisions + docs.** The `DocxDiff` static facade (`Docxodus/DocxDiff.cs`) — `Compare`/`GetRevisions`/`GetEditScriptJson` + `DocxDiffSettings`/`DocxDiffRevision`/`DocxDiffFormatChange`, `#nullable enable`, fully XML-doc'd, anchor-addressed (`LeftAnchor`/`RightAnchor`), multi-author/consolidate-compatible, internal `IrDiffSettings` kept internal. 15 public-surface smoke tests. D3 (markdown cutover — defer, recommended at M2.5) + D4 (default-engine swap — recommendation recorded, ratification post-burn-in) in the program-plan decision log. `docs/architecture/ir_diff_engine.md` written; `wml_comparer_gaps.md` stale claims corrected; CLAUDE.md + CHANGELOG updated.
- **T5 — cross-layer ripple.** The three entry points exposed through every shipping layer, all routing through one shared core facade **`Docxodus/Internal/DocxDiffOps.cs`** — the single owner of the settings-in (JSON object mirroring `DocxDiffSettings`) and revisions-out (`{"revisions":[…]}`, hand-built trim-safe JSON) wire shapes, the same single-owner pattern as `HtmlConversionOps`. Both bridges are thin passthroughs:
  - **WASM** — `wasm/DocxodusWasm/DocxDiffBridge.cs` (`[JSExport] Compare` bytes→bytes, `GetRevisionsJson`, `GetEditScriptJson`), revision/settings DTOs in `JsonContext.cs`.
  - **npm** — `DocxDiffSettings`/`DocxDiffRevision` + `DocxDiffRevisionGranularity`/`DocxDiffFormatComparison` enums + the `DocxDiffBridge` slice on `DocxodusWasmExports` (`npm/src/types.ts`); `docxDiffCompare`/`docxDiffGetRevisions`/`docxDiffGetEditScript` wrappers (`npm/src/index.ts`); a 4-test in-browser Playwright spec (`npm/tests/docx-diff.spec.ts`) over the WC001 fixtures.
  - **Python** — the stdio host gains `docx_diff_compare`/`docx_diff_get_revisions`/`docx_diff_get_edit_script` ops (`tools/python-host/Dispatcher.cs`); `docx-scalpel` ships the matching module functions + frozen `DocxDiffSettings`/`DocxDiffRevision`/`DocxDiffFormatChange` dataclasses + `DocxDiffRevisionType`/`DocxDiffRevisionGranularity`/`DocxDiffFormatComparison` enums (`python/src/docx_scalpel/{session,types,enums}.py`).
  - All three are **stateless** (two DOCX blobs in, no session handle), since `DocxDiff` is a pure two-document compare — they sit as module-level functions alongside `convert_docx_to_html`, not on the session class.

### Final scoreboard / ratchets

- GetRevisions genuine-pass ratchet **177**; PASS-or-deviation floor **179/179** (0 FAIL).
- Produced-markup floor **39**.
- Round-trip allowlist **4 fixtures**.
- **2 evidence-retained GetRevisions deviations** — WC-1450 and WC-1830, both the single 1:N split root cause, deferred to M2.6 with design sketch.

### Public surface + ripple state

`DocxDiff` is the program's first public comparison surface, live in all four layers (.NET core, WASM/npm, python-host/`docx-scalpel`). `WmlComparer` remains the default/blessed engine; `DocxDiff` ships as a production-candidate (D4 swap deferred post-burn-in).

### M2.6 sketch

Engine-level 1:N paragraph split/merge alignment: detect in gap fill via in-order containment of `bag(L)` by the union of an adjacent right-block run; represent as `IrSplitBlockOp`/`IrMergeBlockOp`; ripple through the apply-verifier, markup renderer, JSON writer/reader, and fuzzer. A real capability, not a patch — explicitly out of the M2.5 timebox. Full sketch in `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md`.

### Verification

Full .NET suite **1954 passed / 0 failed / 1 skipped**; `scripts/build-wasm.sh` green (DocxodusWasm.wasm rebuilt with the new bridge); `npm run build` end-to-end + `npx tsc --noEmit` clean; pyhost `dotnet build` clean; `docx-scalpel` import smoke + `mypy` clean (7 source files, no issues); the new DocxDiff Playwright spec **4/4 green** in-browser against the real WASM bridge. NOT separately re-run this task (unchanged by T5, green at T1–T4): the IR.Diff corpus/fuzz/projection-equivalence harnesses and the Release warnings-as-errors build — T5 added only additive bridge/wrapper code with no core-engine change.

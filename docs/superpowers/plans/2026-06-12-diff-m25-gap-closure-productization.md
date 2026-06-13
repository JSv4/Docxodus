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

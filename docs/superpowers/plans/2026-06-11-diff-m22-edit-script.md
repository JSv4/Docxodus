# Diff Engine — M2.2 Intra-Block Diff + Edit Script

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** The diff-as-data product: token-level diff inside Modified pairs, the `IrEditScript` (anchor-addressed, JSON-round-trippable, apply-verifiable), similarity-based gap pairing + fuzzy moves (`MovedModified` becomes reachable), row/cell table granularity, and resolution of the FormatFingerprint run-boundary-noise finding.

**Baseline:** `feat/diff-m22` @ 2c4783a (M2.1 merged: tokenizer + aligner, 92 WC pairs covered, invariants test infrastructure in `IrAlignmentAsserts`).

**Program-plan contract (M2.2):** ordered operations (insert/delete/equal/move/format-change) addressed by anchor + token span; move pairs linked; re-diffing within matched move pairs; format changes from ContentHash-equal/fingerprint-different + token-level comparison; exit invariants — apply(script, left) reconstructs right at text level; script round-trips through JSON.

**Carry-list items owed (M2.1 Outcome):** cross-gap move+edit, MovedModified, table row/cell granularity, similarity gap pairing, FormatFingerprint run-boundary over-sensitivity (diagnose root noise in WC-BodyBookmarks FIRST, then choose: new IR normalization rule vs diff-time modeled-only format comparison — decision recorded with evidence).

**Layout:** `Docxodus/Ir/Diff/` continues; all internal, `#nullable enable`, WASM-safe, no new deps; deterministic everywhere.

## Task 1: Intra-block token diff

`IrTokenDiffer.Diff(IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right, IrDiffSettings) → IrTokenDiff`. Sequence diff over `MatchKey` (Myers O(ND) or LCS with the standard middle-snake/linear-space variant — pick, justify, cite; inputs are word-grain so sizes are modest). Post-pass: within Equal runs, tokens whose `IrRunFormat` records differ → `FormatChanged` spans (per-token comparison, record equality). Op model:

```csharp
internal enum IrTokenOpKind { Equal, Insert, Delete, FormatChanged }
internal sealed record IrTokenOp(IrTokenOpKind Kind, int LeftStart, int LeftEnd, int RightStart, int RightEnd); // token-index half-open spans; char spans derivable from tokens
internal sealed record IrTokenDiff(IrNodeList<IrTokenOp> Ops);  // ordered, covering both sides exactly once
```

Invariants (asserted in tests + a reusable asserts helper): ops cover left tokens exactly once ascending, right tokens exactly once ascending; Equal/FormatChanged spans have equal lengths + equal MatchKeys pairwise; FormatChanged ⇒ some format record differs pairwise; determinism. Tests: single word change, prefix/suffix edits, all-changed, all-equal, separator-only change, format-only run (bold word: Equal content → FormatChanged span), empty sides, adversarial repeated words ("the the the …"), determinism.

## Task 2: IrEditScript + apply-verification + JSON

```csharp
internal enum IrEditOpKind { EqualBlock, FormatOnlyBlock, ModifyBlock, InsertBlock, DeleteBlock, MoveBlock, MoveModifyBlock }
internal sealed record IrEditOp(IrEditOpKind Kind, string? LeftAnchor, string? RightAnchor,
                                IrTokenDiff? TokenDiff, int? MoveGroupId, bool? IsMoveSource);
internal sealed record IrEditScript(IrNodeList<IrEditOp> Operations);
```

`IrEditScriptBuilder.Build(IrDocument left, IrDocument right, IrDiffSettings) → IrEditScript`: runs the aligner, tokenizes + token-diffs every Modified pair (paragraphs; non-paragraph Modified blocks — tables in M2.2 Task 4, others stay TokenDiff=null), links move pairs (one MoveBlock per side, shared MoveGroupId, source+destination both present in document order). Script ordering mirrors the alignment's entry order; anchors are the blocks' anchor strings.
- **Apply-verification** (the exit invariant): a test-side `IrEditScriptVerifier` that reconstructs the RIGHT body's per-block token text sequence from the LEFT IR + script (Equal/FormatChanged copy left tokens; Insert takes right tokens — the verifier may consult the right doc for inserted content, that's fine: the invariant is structural consistency, not self-containment; document) and asserts block-by-block text equality against the actual right IR. Run it over every WC corpus pair (extend the corpus test) + all synthetic cases.
- **JSON**: `IrEditScriptJson.Write/Read` round-trip (hand-written writer/reader, deterministic; same style as IrDiagnosticJson). Round-trip equality test (script == Read(Write(script))).

## Task 3: Similarity gap pairing + fuzzy moves (MovedModified)

- **In-gap similarity pairing:** replace blind positional Modified pairing inside gaps with similarity-scored pairing: score = containment/Jaccard over token MatchKey multisets (threshold via `IrDiffSettings.BlockSimilarityThreshold`, default ~0.5 — justify; ties broken by position; unpaired → Insert/Delete as today). Must stay deterministic + bounded (gaps only — document the cost model vs the G²/2 note).
- **Cross-gap fuzzy moves:** after gap fill, among leftover Deleted×Inserted blocks: similarity ≥ `MoveSimilarityThreshold` (default 0.8 — mirror WmlComparerSettings' default, cite) + minimum token count (`MoveMinimumWordCount`-equivalent, default 3, cite) → MovedModified pair (re-token-diffed in the edit script: move + nested edits — THE capability WmlComparer can't express). Deterministic candidate ordering (best score, then left position).
- Aligner kinds: MovedModified now produced; update IrAlignmentAsserts invariants accordingly (MovedModified ⇒ both non-null, ContentHash NOT required equal).
- Tests: moved-and-edited paragraph → MovedModified with nested token ops (headline); cross-gap edit now pairs as Modified instead of Del+Ins (the M2.1 limitation test gets upgraded); thresholds respected (below-threshold stays Del+Ins); boilerplate still zero false moves; WC corpus re-run — report histogram drift vs M2.1 (expect some Del+Ins converting; invariants still hold).

## Task 4: Table row/cell granularity + fingerprint-noise resolution + close

- **Tables:** Modified table pairs get structural treatment: align rows by row ContentHash (same anchor/LIS/gap machinery generalized or a simpler unique-hash + positional scheme — justify scope), cells positionally within paired rows; paragraph cells token-diff. Edit script gains nested table ops (`ModifyTable { RowOps[...] }` — design minimally, document). A cell-text edit must surface as a token diff in that cell, not a whole-table Modified blob.
- **Fingerprint noise:** diagnose WC-BodyBookmarks' actual rPr noise (dump a few content-equal/fingerprint-diff paragraph pairs' rPr leftovers). Decide WITH EVIDENCE: (a) new IR normalization rule(s) for genuine noise (e.g. w:lang/w:noProof leftovers — snapshot churn, reviewed), and/or (b) diff-time format comparison policy (`IrDiffSettings.FormatComparison = ModeledOnly | Full`, default justify) for the remainder. Record the decision + evidence in the plan outcome; FormatOnly count on WC-BodyBookmarks must drop dramatically (report before/after).
- Close: `## M2.2 Outcome` here (incl. WC corpus histogram before/after across M2.2), program-plan M2.2 row, CHANGELOG, full verification.

## Out of scope (M2.3+)

Revisions-API renderer, differential harness vs WmlComparer, native OOXML markup, any public surface.

## M2.2 Outcome

**Status: COMPLETE (2026-06-11).** All four tasks landed; exit criteria met. The diff layer
(`Docxodus/Ir/Diff/`, all `internal`, `#nullable enable`, WASM-safe, no new dependencies) now holds the
intra-block token differ (Task 1), the anchor-addressed `IrEditScript` with apply-verification + JSON
round-trip (Task 2), similarity gap pairing + cross-gap fuzzy moves (Task 3), and — this task — row/cell
table granularity, the diff-time `FormatComparison` policy, and the close.

### Headline capabilities (M2.2 product)

- **Token-level diff inside Modified paragraph pairs** (`IrTokenDiff`: Equal / Insert / Delete /
  FormatChanged), Myers O(ND).
- **`IrEditScript`** — ordered, anchor-addressed (`kind:scope:unid`), JSON-round-trippable,
  apply-verifiable (apply(script,left) reconstructs right at text level over the whole WC corpus + all
  synthetic cases).
- **Fuzzy moves** — `MovedModified` reachable: a relocated-AND-edited paragraph emits a move pair with a
  nested token diff (the capability WmlComparer cannot express).
- **Nested table diffs** — a Modified table pair carries `IrTableDiff(rowOps[])`; a cell-text edit
  surfaces as a token diff *inside that cell*, not a whole-table blob.
- **Diff-time format policy** — `FormatComparison = ModeledOnly (default) | Full`, resolving the M2.1
  FormatFingerprint run-boundary-noise finding without any IR snapshot churn.

### Sub-task A — table row/cell granularity (design)

- **Row alignment (self-contained, `IrTableDiffer`).** Rows carry a `ContentHash` but no
  `FormatFingerprint`, and `IrRow` is not an `IrBlock`, so the body aligner's IrBlock/fingerprint-keyed
  machinery does not apply directly. A focused row aligner mirrors the SAME design at row grain:
  unique-`ContentHash` anchoring → LIS spine (on-spine = `EqualRow`, off-spine exact-hash = `MovedRow`)
  → positional gap fill (paired = `ModifyRow`, surplus left = `DeleteRow`, surplus right = `InsertRow`).
  Justification for not generalizing the block aligner: that would mean refactoring it around a
  hash-provider interface for little reuse; the ~120-line row aligner is cleaner and self-contained.
  Row kinds reduce to Unchanged/Modify/Insert/Delete + optional Moved (free off-spine exact only; no
  fuzzy row moves — documented limitation).
- **Cell pairing — positional.** Within a `ModifyRow`, cell *i* pairs with cell *i* (grid-aware
  gridSpan/vMerge pairing is M2.3+). A content-equal cell carries no recursion; a differing cell's
  paragraph blocks are aligned by the SHARED block aligner (`IrBlockAligner.AlignBlocks`, the new
  generalized entry point) and projected through the SHARED `IrEditScriptBuilder.ProjectAlignment`, so a
  cell-text edit lands as a `ModifyBlock` carrying a token diff — recursion reuses the exact body
  machinery one level down.
- **Edit-script shapes (`IrEditScript.cs`).** `IrEditOp` gained `IrTableDiff? TableDiff`;
  `IrTableDiff(IrNodeList<IrRowOp>)`, `IrRowOp(kind, leftRowAnchor?, rightRowAnchor?, cellOps?,
  moveGroupId?, isMoveSource?)`, `IrCellOp(leftCellAnchor?, rightCellAnchor?, blockOps?)`. JSON
  writer/reader extended (`tableDiff`/`rowOps`/`cellOps`/`blockOps`); apply-verifier extended to
  reconstruct table content row-by-row/cell-by-cell and to validate row + cell anchors against the
  actual tables (row/cell anchors are NOT in `AnchorIndex` — only cell-child blocks are — so the
  verifier resolves them against the table structure directly).

### Sub-task B — fingerprint-noise diagnosis + resolution

**Diagnosis (with evidence, `FingerprintNoiseDiagnosticTests`, `Category=Diagnostic`).** Over the
WC-BodyBookmarks pair (the sole source of the corpus' 1,714 FormatOnly entries), all 1,714 FormatOnly
paragraph pairs are ContentHash-equal with every MODELED run-format field byte-identical — the ONLY
difference is the unmodeled-rPr `UnmodeledDigest`. Leftover (unmodeled) rPr child inventory across all
pairs:

| element | occurrences | nature |
|---|---:|---|
| `w:lang` | 4597 | locale (nb-NO) — genuine IR fact, pure diff noise |
| `w:iCs` | 1328 | complex-script italic toggle — diff noise |
| `w:bCs` | 550 | complex-script bold toggle — diff noise |
| `w:rFonts` (hAnsi/cs faces; ascii is modeled) | 33 | mostly diff noise |
| `w:noProof` | 4 | already dropped by N2 in `Canonicalize` — does NOT reach `UnmodeledDigest`; the diagnostic's leftover reconstruction over-reports it |
| `w:rtl` | 4 | borderline |
| `w:szCs` | 3 | complex-script size — diff noise |

**Decision (with evidence).** Implement `IrDiffSettings.FormatComparison = ModeledOnly | Full`, default
**ModeledOnly**. Do NOT add IR-level normalization rules. Rationale:

- A `w:rPrChange`-grade format-change report can only ever DESCRIBE modeled fields, so a FormatOnly
  classification driven by an undescribable unmodeled-digest flip is a pure false positive — exactly the
  1,714-entry population. ModeledOnly is therefore the correct default.
- These elements (lang/bCs/iCs/szCs/rtl/rFonts-cs) are LEGITIMATE IR facts that byte-fidelity consumers
  want; an N-rule strip would cause snapshot churn for ALL consumers. `Full` preserves the M2.1 behavior
  for them. So this is purely a **diff-time policy**: the IR's stored hashes never change.
- `w:noProof` is already dropped by N2 (it never reaches the digest), so no new normalization is owed.

**Layering (documented in `IrDiffSettings`/`IrModeledFormat`).** ModeledOnly compares `IrRunFormat`
records EXCLUDING `UnmodeledDigest` at the token level (the differ's FormatChanged post-pass), and at the
BLOCK level recomputes a boundary-normalized modeled-only signature at diff time — the per-token
`(MatchKey, modeled-format key)` sequence, boundary-independent by construction — instead of trusting the
reader's stored block `FormatFingerprint` (which folds in the digest AND is run-boundary-sensitive). The
IR's stored hashes DO NOT change.

**Measurement (default ModeledOnly).** WC-BodyBookmarks pair: **FormatOnly 1714 → 50**, Unchanged
290 → 1664, Modified 1374 → 1310 (forward). Corpus-wide aligner totals: **FormatOnly 1714 → 50**,
**Unchanged 556 → 2220**, Modified 1488 → 1419, Moved 3 (unchanged), Inserted 901 → 970, Deleted
35 → 104. (Modified/Inserted/Deleted drift reflects content-equal pairs that previously anchored as
FormatOnly now anchoring as Unchanged and freeing the gaps differently — net: the noise collapsed and
real content classifications stayed coherent; all invariants hold both directions.)

### Final corpus histograms

- **Aligner (`IrAlignerCorpusTests`, forward):** Unchanged=2220, FormatOnly=50, Modified=1419, Moved=3,
  MovedModified=0, Inserted=970, Deleted=104.
- **Edit script (`IrEditScriptCorpusTests`, forward op kinds):** EqualBlock=2220, FormatOnlyBlock=50,
  ModifyBlock=1419, InsertBlock=970, DeleteBlock=104, MoveBlock=6, MoveModifyBlock=0. (Moved=3 alignment
  entries → 6 MoveBlock ops, one source + one destination each.)
- **Table-diff stats (forward):** 18 ModifyBlocks carry a nested `IrTableDiff`; 53 row ops (Equal=26,
  Modify=20, Insert=1, Delete=6, Moved=0); 39 cell ops; **25 cells carry a block-level token diff** —
  i.e. 25 cell-text edits that M2.1 would have buried in a whole-table Modified blob now surface as
  in-cell token diffs.

### Verification

- `Ir.Diff` suite: **102 passed** (was 92 at M2.1 close; +10 across table-diff, format-comparison, and
  the diagnostic). Full IR suite: 327 passed. Full test suite: 1,868 passed, 1 skipped — the lone
  `Scale_guard` red under full-suite CPU contention is timing flakiness (passes isolated at ~4.4×, well
  within the 8× bound; not a regression). Release build (`-c Release`, warnings-as-errors) of the
  solution succeeds.

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

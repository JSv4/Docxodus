# Diff Engine — M2.3 Revisions Surface + Differential Harness + Parity Scoreboard

> **For agentic workers:** REQUIRED SUB-SKILL: superpowers:subagent-driven-development. Executed by the controller as sequenced subagent tasks.

**Goal:** The new engine becomes comparable to the old one. A `WmlComparerRevision`-shaped revisions surface over the edit script; a differential harness running BOTH engines head-to-head on the WC corpus with triaged divergences; a deterministic generative fuzzer; and — per the standing USER DIRECTIVE — a parity scoreboard over the entire `Docxodus.Tests/WmlComparer*` suite establishing exactly what M2.4 must close to 100%.

**Baseline:** `feat/diff-m23` @ M2.2 merge (`17166ee`): IrEditScript apply-verified over 92 WC pairs, fuzzy moves, table granularity, ModeledOnly format policy.

**USER DIRECTIVE (binding):** the ultimate bar is the IR engine passing ALL the same tests as the original WmlComparer.cs. This milestone builds the measurement; M2.4 drives it.

**Layout:** `Docxodus/Ir/Diff/` continues; internal; deterministic; no WmlComparer.cs changes whatsoever.

## Task 1: Revisions surface — `IrRevisionRenderer`

`Render(IrEditScript, IrDocument left, IrDocument right, IrDiffSettings) → IrNodeList<IrRevision>` where `IrRevision` mirrors the public `WmlComparerRevision`'s consumer-relevant shape (study it in WmlComparer.cs ~line 3300 + CLAUDE.md: RevisionType Inserted/Deleted/Moved/FormatChanged; Text; Author/Date from settings — add `AuthorForRevisions`/`DateTimeForRevisions` to IrDiffSettings with WmlComparer-compatible defaults + a `Deterministic` flag pinning the date; MoveGroupId/IsMoveSource; FormatChange details record for FormatChanged: old/new modeled fields + changed-property names — derive from the paired token formats). Mapping: InsertBlock/DeleteBlock → block-level Inserted/Deleted revisions (text = concatenated block text); ModifyBlock → one revision per token-op span (Insert spans → Inserted with span text; Delete spans → Deleted; FormatChanged spans → FormatChanged with details); MoveBlock/MoveModifyBlock → Moved pairs with group ids (MoveModify destination ALSO yields nested Inserted/Deleted revisions for its token edits — decide ordering, document); FormatOnlyBlock → FormatChanged revision(s) from the per-token modeled diffs; EqualBlock → nothing; tables → recurse nested ops. Ordering: script order, deterministic. Tests per mapping + JSON-ish dump determinism + the move/format details.

## Task 2: Differential harness vs WmlComparer

`IrVsWmlComparerTests.cs` (Trait Category=Differential): for every WC corpus pair (reuse WcCorpus): old engine = `WmlComparer.Compare(left,right,settings)` + `WmlComparer.GetRevisions(...)`; new = IR pipeline → IrRevisionRenderer. Compare SEMANTICALLY, not structurally: normalized multisets of (kind, normalized-text) where normalized-text strips whitespace runs (the two engines atomize differently — define the normalization precisely, document); Moved compared by paired source/dest text. Classify each pair: MATCH / SUBSET (one side reports more granular ops whose union equals the other) / DIVERGENT. Output per-pair classification + corpus totals; write DIVERGENT details to artifacts dir. NO pass/fail threshold this milestone — the test asserts totality + writes the triage table to the task report; the controller adjudicates. Include known-cause buckets where identifiable (e.g. WmlComparer's special-char drops, sectPr handling, whitespace atomization). Both directions.

## Task 3: Generative fuzzer

`IrDiffFuzzTests.cs` (Trait Category=Fuzz, deterministic seeds): a small mutation engine over programmatic docs (seeded Random, N seeds × M mutations; mutations: edit word, insert/delete paragraph, relocate paragraph, bold a word, edit table cell, insert/delete row). For each case: (a) IR pipeline invariants — alignment invariants + apply-verifier + JSON round-trip (the strongest oracle we own); (b) differential check vs WmlComparer's GetRevisions under the Task-2 normalization where the mutation class is comparable (text edits/inserts/deletes — moves/format excluded from cross-engine equivalence, WmlComparer semantics differ; document exclusions). Seed corpus committed small (e.g. 50 seeds CI, 500 nightly-style opt-in env var). Failures must dump the seed + minimized repro info.

## Task 4: WmlComparer parity scoreboard + close

The USER-DIRECTIVE deliverable. Inventory EVERY test in Docxodus.Tests/WmlComparer*.cs (Theories expand: count InlineData rows): for each, categorize what it asserts (compare-output-document content / accept-reject round-trip / GetRevisions counts+types / move detection semantics / format-change details / consolidate / settings behavior / thread safety) and what the IR engine needs to satisfy it (revisions-surface-only [runnable NOW — run it via an adapter where feasible and record pass/fail] vs native-OOXML-markup [M2.4] vs consolidate [out of v1 scope — flag for user decision] vs internal-to-old-engine [not applicable — justify each]). Produce the scoreboard table in the plan outcome: total tests, runnable-now pass/fail, M2.4-blocked, out-of-scope-pending-user. An `IrWmlComparerAdapter` (test-side) that exposes the IR pipeline through a WmlComparer-shaped GetRevisions API for the runnable-now rows. Close: `## M2.3 Outcome`, program plan row, CHANGELOG, full verification.

## Out of scope (M2.4+)

Native OOXML markup generation (w:ins/w:del/moveFrom/moveTo/rPrChange in a produced document), accept/reject round-trip invariants on produced documents, Compare()-shaped byte output, consolidate.

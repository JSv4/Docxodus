# IR Diff Engine

> **Status:** Public surface shipped (M2.5). The engine ships as `DocxDiff` (`Docxodus/DocxDiff.cs`) — a **production-candidate**, NOT yet the blessed default. `WmlComparer` remains the default comparison API. The IR engine becomes the default only after the Word manual-verification checklist clears and a burn-in period (decision **D4**, still open). See the decision log in `docs/superpowers/specs/2026-06-11-ir-diff-layout-program-plan.md`.

The IR diff engine is a structure-aware DOCX comparison engine built on Docxodus' intermediate document representation (IR). It is the write-side analogue of the read-only IR pipeline that backs the markdown projection: it reads two documents into anchor-addressed IR snapshots, computes an **edit script** between them, and renders that script three ways — native tracked-changes markup, a consumer revision list, or the script itself as JSON (diff-as-data).

It is a sibling to `WmlComparer` in the comparison family. The differences that motivate it:

1. **Anchor-addressed revisions.** Every revision carries the stable block anchor(s) (`kind:scope:unid`) it derives from — the same anchor grammar as the markdown projection and `DocxSession`. A revision can be located in the projection or fed straight to a `DocxSession` mutation. `WmlComparer.WmlComparerRevision` has no anchors.
2. **Diff-as-data.** The edit script serializes to stable JSON, so the diff is storable, transportable to non-.NET consumers, and auditable. `WmlComparer` only produces an OOXML document or an in-memory revision list.
3. **A modeled IR.** Comparison runs over the IR's typed blocks/runs/format records rather than raw atom streams, which makes table row/cell-precise diffs, footnote/endnote scope diffs, and modeled-format-change detection first-class.

## Public surface

`public static class DocxDiff` (`Docxodus/DocxDiff.cs`):

| Method | Returns | Purpose |
|---|---|---|
| `Compare(left, right, settings?)` | `WmlDocument` | Tracked-changes document with native `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` markup. Satisfies the WmlComparer contract: `AcceptRevisions(result) ≡ right`, `RejectRevisions(result) ≡ left` at the per-block text level. |
| `GetRevisions(left, right, settings?)` | `IReadOnlyList<DocxDiffRevision>` | The consumer revision list, rendered directly off the edit script (no produce-then-reparse round-trip). |
| `GetEditScriptJson(left, right, settings?)` | `string` | The edit script as indented JSON — the diff-as-data differentiator. |

Supporting public types: `DocxDiffSettings`, `DocxDiffRevision`, `DocxDiffRevisionType`, `DocxDiffFormatChange`, `DocxDiffRevisionGranularity`, `DocxDiffFormatComparison`. All `#nullable enable`, fully XML-documented, no static or process-global state (multi-author / consolidate-compatible — author flows per call via `DocxDiffSettings.AuthorForRevisions`).

### Anchor grammar and DocxSession interop

`DocxDiffRevision.LeftAnchor` / `RightAnchor` are block anchors of the form `kind:scope:unid` — e.g. `p:body:a1b2c3d4` (a body paragraph), `li:body:…` (list item), `tbl:body:…` (table), `p:fn3:…` (a paragraph in footnote 3). The `kind`/`scope` match the markdown projection's and `DocxSession`'s anchor grammar, so:

- A revision resolves to a location in the markdown projection (review UIs, blame).
- A revision can be passed straight to a `DocxSession` call (`ReplaceText`, `GetBlockMetadata`, `AddAnnotation`, …) on the corresponding document.

A left anchor resolves against the `left` document's IR; a right anchor against `right`. Anchor presence by revision type: Inserted → right only; Deleted → left only; FormatChanged → both; Moved source → left, Moved destination → right. A token-level revision inside a modified/moved-and-edited block carries the enclosing block's anchor(s).

## Pipeline

```
                    settings (DocxDiffSettings.ToIrDiffSettings → IrDiffSettings)
                         │
left  ─ IrReader.Read ──▶ IrDocument ─┐
                                      ├─▶ IrEditScriptBuilder.Build ─▶ IrEditScript ─┬─▶ IrMarkupRenderer.Render   ─▶ WmlDocument
right ─ IrReader.Read ──▶ IrDocument ─┘                                              ├─▶ IrRevisionRenderer.Render ─▶ revisions
                                                                                     └─▶ IrEditScriptJson.Write    ─▶ JSON
```

Internal stages (all `internal`, under `Docxodus/Ir/Diff/`):

- **`IrReader`** — reads a `WmlDocument` to an `IrDocument` (anchor-indexed blocks; accepted-revision view; provenance off for the diff path). Shared with the markdown projection.
- **`IrDiffTokenizer`** — splits IR runs into word/separator/atomic tokens with match keys (case folding, NBSP conflation, hyperlink-target-in-key, field transparency). The diff's tokenization, NOT an IR fact.
- **`IrBlockAligner`** — unique-hash `(ContentHash, FormatFingerprint)` anchoring → LIS spine → in-order gap fill; relocations fall off the spine as moves; similarity-based in-gap pairing + cross-gap fuzzy moves.
- **`IrTokenDiffer`** — Myers O(ND) token diff inside a paired block (Equal/Insert/Delete/FormatChanged).
- **`IrTableDiffer`** — nested table row/cell diffs (a cell-text edit surfaces as a token diff inside that cell, not a whole-table blob).
- **`IrEditScriptBuilder`** — assembles the `IrEditScript` from the alignment + token/table diffs, including footnote/endnote scope ops.
- **`IrMarkupRenderer` / `IrRevisionRenderer` / `IrEditScriptJson`** — the three renderers above.

## Edit script

The `IrEditScript` is an ordered list of block operations (`IrEditOpKind`), plus a parallel `noteOps` list for footnote/endnote scopes:

| Kind | Meaning | Anchors |
|---|---|---|
| `EqualBlock` | Both sides identical | both |
| `FormatOnlyBlock` | Text-equal, modeled format differs (`w:rPrChange`-grade) | both |
| `ModifyBlock` | Same block, edited (carries a nested token/table diff) | both |
| `InsertBlock` | Right-only block | right only |
| `DeleteBlock` | Left-only block | left only |
| `MoveBlock` | Relocated block (source + destination ops share a `moveGroupId`) | source: left, dest: right |
| `MoveModifyBlock` | Relocated AND edited (the case `WmlComparer` cannot express as a move) | source: left, dest: right |
| `SplitBlock` | One left paragraph split across N≥2 right paragraphs (M2.6) | left + `splitMergeAnchors` (the N rights) |
| `MergeBlock` | N≥2 left paragraphs fused into one right paragraph (M2.6) | right + `splitMergeAnchors` (the N lefts) |

**Footnote/endnote structural fidelity.** The `noteOps` carry per-note block diffs that the markup renderer applies *inside* the produced footnotes/endnotes part, after which `IrMarkupRenderer.RenumberNoteIds` re-sequences every reference + definition into body-reference document order (mirroring `WmlComparer`'s `ChangeFootnoteEndnoteReferencesToUniqueRange`). The invariants this path guarantees — verified by `DocxDiffScenarioTests.Scenario_PreservesFootnoteStructure` (every edit×feature scenario) and `DocxDiffFootnoteRobustnessTests`, with a headless-LibreOffice load backstop:
- **Definition ids are unique** and every body reference resolves to exactly one definition. A `continuationNotice` (or any typed) reserved note at a *positive* id keeps its id and leads the part; real notes renumber to a range disjoint from the kept reserved ids (the counter seeds above the highest positive reserved id), so a reserved note never collides with a renumbered real note.
- **A reference is never dropped.** A footnote/endnote reference is a zero-width inline; the token-diff run rebuilder (`SourceRunModel.Slice`) attributes a boundary zero-width to exactly the op whose token range owns it (`ZeroWidthBoundaries`), so a reference at the tail of an edited paragraph survives and is never double-counted.
- **A reference-deleted note's definition is emitted once.** When an edit removes a reference but the definition lingers in both stores, the builder reconciles the orphan into a single matched pair (identity = `w:id`) rather than a delete + an insert of the same id; the renumber pass links a `w:del` reference to its preserved definition so it stays resolvable on reject even when its id ≠ its reference ordinal.
- **A note-in-note reference is renumbered with its target.** A footnote/endnote whose body cites another note keeps a reference inside its definition body; the body-only renumber walk never visits it, so `RenumberNoteIds` records each definition's old→new id and remaps same-kind references inside the note part afterwards — otherwise a nested reference to a renumbered note dangles on accept/reject (cross-kind nesting — an endnote ref inside a footnote body — is a documented v1 gap).

**Comment fidelity (Covered).** Comments (`w:commentRangeStart`/`End`/`w:commentReference`) are dropped from IR paragraphs (reader rule N15 records them as `IrCommentStore` char-offset spans for the markdown projection only), but they are carried through the fine token-diff path as `AlwaysKeep` zero-width markers — exactly like bookmarks — so an edited commented paragraph gets **fine per-word `w:ins`/`w:del`** markup with its comment anchors intact (no whole-block bail). Three passes guarantee integrity, the comment analogue of `NormalizeBookmarks`:

- `MergeRightCommentDefinitions` copies any RIGHT-only `w:comment` the emitted right-sourced content references into the LEFT-based comments part (creating it if the left had none), plus the referenced `commentsExtended` (`w15:commentEx` `paraIdParent` reply links) and `commentsIds` entries — so a right-**added** comment never dangles.
- `NormalizeComments` reconciles the body so every `commentReference` resolves to exactly one `w:comment`, every `commentRangeStart` is unique + pairs 1:1 with a `commentRangeEnd`, and an unchanged comment survives both accept and reject: **(A)** a common comment with a bare survivor collapses to a single bare range; **(A2)** a right-added / left-deleted comment's bare markers are wrapped in `w:ins`/`w:del` so they toggle with their side; **(B)** a wholly-rewritten comment's del/ins copies renumber the deleted copy to a fresh id + a cloned definition — the comment-dedup analogue of the bookmark renumber-collision — with its OWN fresh `w14:paraId` + a cloned `commentsExtended` entry so a reject-side threaded reply keeps its parent link; **(C)** orphaned markers are paired/dropped so the output is always schema-valid + fully resolvable.

DocxDiff is strictly ahead of the blessed `WmlComparer` oracle, which drops comments entirely on any edit. Verified by `DocxDiffCommentStructureTests` (9-shape synthetic corpus) + `DocxDiffCommentRealDocTests` (vendored `TestFiles/DD/DD002-DenseComments.docx`) under OpenXmlValidator schema validity, a comment-structure round-trip, and a headless-LibreOffice comment oracle (`tools/diffharness/lo/lo_comment_check.py`). See `docs/ooxml_corner_cases.md` (comment threading is keyed on `w14:paraId`; the validator's paraId-uniqueness blind spot).

A `ModifyBlock` over a paragraph carries a `tokenDiff`; over a table, a `tableDiff` (row ops with nested cell ops); a textbox-bearing block carries `textboxDiffs`. A `SplitBlock`/`MergeBlock` carries `splitMergeAnchors` (the plural side, document order) plus `segmentDiffs` — one COMPLETE token diff per member, whose singular-side spans are slice-local; the slices tile the singular side's token stream exactly, in order (the partition invariant: slice *i*'s length = Σ of segment *i*'s non-Insert left-span lengths for a split, non-Delete right-span lengths for a merge — boundaries are implicit in the diff ops, never stored). N:M (plural on both sides) is physically representable by the nullable fields but rejected by the test-side `AssertSplitMergePairing` (a `SplitBlock` must carry a null `rightAnchor`, a `MergeBlock` a null `leftAnchor`, no anchor may appear in two ops' `splitMergeAnchors`) and never emitted by the builder — the pairing assert is the load-bearing scope ceiling. The JSON is a faithful serialization of this structure (top-level `operations` + optional `noteOps`; `splitMergeAnchors`/`segmentDiffs` appear only on split/merge ops, so pre-M2.6 scripts serialize byte-identically), and is deterministic for identical inputs.

## Settings

`DocxDiffSettings` is the public mirror of the internal `IrDiffSettings`; it exposes the consumer-relevant subset and maps onto it in `ToIrDiffSettings()`.

| Public setting | Default | Maps to | Notes |
|---|---|---|---|
| `AuthorForRevisions` | `"Open-Xml-PowerTools"` | `IrDiffSettings.AuthorForRevisions` | matches `WmlComparerSettings` |
| `Deterministic` | `true` | `IrDiffSettings.Deterministic` | **deviation from `WmlComparerSettings`** (which is wall-clock by default) |
| `DateTimeForRevisions` | `null` → epoch or `DateTime.Now` | `IrDiffSettings.DateTimeForRevisions` | explicit value always wins |
| `CaseInsensitive` / `Culture` | `false` / `null` | `CaseInsensitive` / `Culture` | |
| `ConflateBreakingAndNonbreakingSpaces` | `true` | same | |
| `WordSeparators` | `null` → default set | `WordSeparators` | |
| `DetectMoves` | `true` | `RenderMoves` | render-time relabel: the engine always ALIGNS a relocation as a move; this controls whether it is REPORTED as one |
| `MoveSimilarityThreshold` | `0.8` | same | |
| `MoveMinimumWordCount` | `3` | `MoveMinimumTokenCount` | |
| `RevisionGranularity` | `Fine` | `RevisionGranularity` | `Fine` = engine-native one-revision-per-token-span (byte-stable); `WmlComparerCompatible` = coalesce/trim/prune to match the legacy comparer's coarser revision set |
| `FormatComparison` | `ModeledOnly` | `IrFormatComparison` | `ModeledOnly` reports only modeled-field deltas (false-negative on unmodeled rPr); `Full` sees every rPr difference |
| `PreAcceptInputRevisions` | `false` | — (a pre-pass, not an `IrDiffSettings` field) | when true, runs `RevisionProcessor.AcceptRevisions` on EACH input before diffing — the first-class "accept-all both sides, then compare" wrapper. See [Inputs that already carry tracked changes](#inputs-that-already-carry-tracked-changes-rule-n13--preacceptinputrevisions). .NET-only in v1 (no bridge ripple). |

### Two honest defaults that deviate from `WmlComparerSettings`

1. **Deterministic dates.** `WmlComparerSettings.DateTimeForRevisions` defaults to `DateTime.Now` — the same compare twice yields different dates. `DocxDiff` pins a fixed epoch by default so output is reproducible. Opt into wall-clock via `Deterministic = false`.
2. **`FormatComparison = ModeledOnly`.** A `w:rPrChange`-grade report can only DESCRIBE modeled fields, so a format change driven by an undescribable unmodeled-only rPr flip (`w:lang`, `w:bCs`, complex-script toggles) is noise. `ModeledOnly` collapses that noise; the trade-off is a false negative on a visible-but-unmodeled change (e.g. `w:shd` run shading). `Full` restores byte-fidelity comparison.

### Inputs that already carry tracked changes (rule N13 + `PreAcceptInputRevisions`)

DocxDiff is frequently asked to diff a document that is itself a redline — its body, notes, headers, or comments already carry un-accepted `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange`. The handling is **explicit and pinned**, not incidental. Characterization tests: `Docxodus.Tests/Ir/Diff/RevisionsInInputDefaultTests.cs` (default) and `PreAcceptInputRevisionsTests.cs` (the flag).

**The default — diff the ACCEPTED VIEW, carry non-body markup through verbatim.**

- **Rule N13 (the IR is a revision-free view).** `IrReader` resolves revisions with `RevisionView.Accept` on a working copy *before* building the IR (the original bytes are untouched). So the edit script — and therefore `GetRevisions`/`GetEditScriptJson` and the body of `Compare` — is computed over the **accepted view** of each side: an input's own `w:ins`/`w:del` never surface as their own diff, and the produced body carries only THIS diff's revisions, attributed to `AuthorForRevisions`. At the body level the round-trip already holds against the accepted view (`reject(result) ≡ accept-view(left)`, `accept(result) ≡ accept-view(right)`).
- **The carry-over leak.** `IrMarkupRenderer` assembles the output on a **clone of the LEFT package** (so styles/numbering/theme/settings/section/media carry over by reuse — what WmlComparer does), and only the **body** (plus *changed* footnotes/endnotes) is rebuilt from accept-clean source. Everything else — **headers/footers, UNCHANGED footnotes/endnotes, styles, the comments part** — is passed through *verbatim from the original left input*. Any pre-existing revision markup there **survives into the result**, attributed to its ORIGINAL author. This is the documented limitation the inspector's `revisionsInInput` entry flags.
- **Why the leak matters.** A leaked pre-existing `w:ins` in, say, a header is then *rejected* by `RejectRevisions` (it strips the insertion) — so under the default the round-trip does **not** hold in the carried-over scopes (`reject(result)` drops header text the accepted view of the left actually contains). Pinned by `Default_leak_breaks_the_header_round_trip`.

**The opt-in — `PreAcceptInputRevisions` (default `false`).** When set, every input is run through `RevisionProcessor.AcceptRevisions` *before* it enters the pipeline, so both the IR read and the cloned output package are revision-free on both sides. It is, by construction, **exactly** `Compare(AcceptRevisions(left), AcceptRevisions(right))` — byte-for-byte identical to that wrapper (oracle: `Flag_on_is_byte_identical_to_accept_all_then_compare_wrapper`). The effect:

- every `w:ins`/`w:del`/`w:moveFrom` in the result is attributable to THIS diff (no stale input revision passed through) in the body, header/footer, note, and style scopes;
- the round-trip holds against the accepted view in those scopes;
- the consumer revision list (`GetRevisions`) is unchanged — the flag only additionally cleans the rendered package's carried-over parts.

**The flag's coverage is exactly what `RevisionProcessor.AcceptRevisions` processes** — the body, headers, footers, footnotes, endnotes, and styles part. Carried-over parts it does **not** process keep their pre-existing revisions: notably the `WordprocessingCommentsPart` (a tracked change inside a COMMENT *definition* survives — pinned by `PreAcceptInputRevisions_does_not_flatten_a_revision_inside_a_comment_definition`) and the `GlossaryDocumentPart` (building-blocks / AutoText entries; the renderer clones the whole left package, including the glossary part, verbatim). These are the honest boundaries of the accept-all pre-pass — resolve revisions in those parts separately if they matter. (The narrower `NumberingDefinitionsPart` carries only style-grade `pPrChange`/`rPrChange`, never `w:ins`/`w:del`, so it is outside the inspector's tracked-changes scope.)

Applied uniformly to all seven entry points (`Compare`/`GetRevisions`/`GetEditScriptJson` and the four consolidate-family methods, accepting the base and every reviewer). **.NET-only in v1** — not yet surfaced on the WASM/npm/python bridges (deferred).

**Two honest costs of accept-all (read before enabling).** Accept-all is a lossy, opinionated pre-flatten:

1. **It flattens pre-existing authorship and change boundaries.** Accepting collapses each input's tracked changes into final text, so *who* made a prior edit and *where* the prior change boundaries were are lost — the result's authorship reflects only this diff.
2. **"Accept all" is itself a policy.** It overrides any change a prior reviewer had left unaccepted — including one they had effectively rejected by leaving in tracked form — materializing every insertion and dropping every deletion. To preserve or re-adjudicate the inputs' in-flight revisions, resolve them by your own policy first, then diff. See `docs/ooxml_corner_cases.md` → "DocxDiff: `PreAcceptInputRevisions` accept-all flattens prior authorship".

## N-way composite / Consolidate

The IR engine merges **N reviewers' edits** — each an independently revised copy of ONE shared base — into a single tracked-changes document, an attributed revision list, a composite edit-script-as-data, and a structured conflict report. This is the IR-native answer to the last `WmlComparer` capability the engine had not addressed: `WmlComparer.Consolidate` (the 84 `CONSOLIDATE` cases in the M2.3 parity inventory, deferred there as "out of v1 scope").

### Public surface

Four entry points on `public static class DocxDiff` (`Docxodus/DocxDiff.cs`), with the supporting types in `Docxodus/DocxDiffConsolidate.cs`:

| Method | Returns | Purpose |
|---|---|---|
| `Consolidate(base, reviewers, settings?)` | `WmlDocument` | One multi-author tracked-changes document — each reviewer's edits stamped with that reviewer's own author name (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange`). The N-way, shared-base counterpart to `Compare`. |
| `GetConsolidatedRevisions(base, reviewers, settings?)` | `IReadOnlyList<DocxDiffConsolidatedRevision>` | The attributed revision list — `DocxDiffRevision`'s shape plus the contributing reviewer's `Author` and, on a conflict winner, a `ConflictId`. |
| `GetConsolidatedEditScriptJson(base, reviewers, settings?)` | `string` | The composite edit script as data: every op additively carries `author`/`sourceReviewer`, a `conflictId` when it won a conflict, and (for a composed paragraph) `authoredTokens` + `sourceRightAnchors`; the document gains a top-level `conflicts` array. |
| `GetConflicts(base, reviewers, settings?)` | `IReadOnlyList<DocxDiffConflict>` | The inspect-before-merge view — the same merge run, surfacing only the conflict list so a caller can review disagreements (and pick a policy) before committing to an output. |

Supporting public types (all `#nullable enable`, XML-doc'd, no static state — author flows per reviewer): `DocxDiffReviewer { Document, Author }`, `DocxDiffConsolidateSettings { Diff, ConflictResolution }` (composes — does not inherit — the `sealed DocxDiffSettings` via `Diff`), `enum ConflictResolution { BaseWins, FirstReviewerWins, StackAll }`, `DocxDiffConsolidatedRevision` (= `DocxDiffRevision` + `ConflictId`), `DocxDiffConflict { Id, BaseAnchor, TokenStart, TokenEnd, AppliedPolicy, Competitors }`, `DocxDiffConflictCompetitor { Author, ResultText }`. N reviewers, no cap; reviewer LIST ORDER is significant (it determines competitor order and policy tie-breaking). Zero reviewers returns the base unchanged / empty lists. The four surfaces are exposed through every shipping layer — WASM (`DocxDiffBridge.Consolidate`/`GetConflictsJson`/`GetConsolidatedRevisionsJson`/`GetConsolidatedEditScriptJson`), npm (`docxDiffConsolidate`/`docxDiffGetConflicts`/`docxDiffGetConsolidatedRevisions`/`docxDiffGetConsolidatedEditScript`), docx-scalpel (`docx_diff_consolidate`/`docx_diff_get_conflicts`/`docx_diff_get_consolidated_revisions`/`docx_diff_get_consolidated_edit_script`) — all routing through `DocxDiffOps`, so the wire shapes live in one place.

### The merge algorithm

The source of truth is `IrCompositeMerger.Merge` (`Docxodus/Ir/Diff/IrCompositeMerger.cs`). The merge builds **N pairwise edit scripts** — `IrEditScriptBuilder.Build(baseIr, reviewer_i)` — which all share the base's anchor space AND, within a paired block, the same base token coordinate system. That shared coordinate system is what makes the merge exact rather than heuristic. It then walks the base document in block order; per base block:

- **Untouched** (no reviewer op, or all `EqualBlock`) → one base-sourced passthrough.
- **One reviewer touched it** → that reviewer's op verbatim, authored to that reviewer.
- **≥2 reviewers, all producing the SAME right result** → consensus: a single op, authored to the first reviewer. A set of consensus deletes collapses to one delete.
- **≥2 reviewers, all paragraph `ModifyBlock` token edits with UNCHANGED paragraph properties** → **token-span composition** (`ComposeTokenSpans`): the per-reviewer token diffs, all expressed over the same base token stream `[0, baseTokenCount)`, compose into one merged authored token-op list. Non-overlapping span edits each land inline under their own author; overlapping spans become a conflict resolved by the policy. The emitted spans tile the base token stream exactly once (a **runtime** totality invariant via `IrCompositeMerger.Invariant` — enforced in Release too, not a `Debug.Assert` that the shipped/CI Release build would strip).
- **≥2 reviewers, all table `ModifyBlock` edits (no `MovedRow`, column structure stable)** → **per-cell table composition** (`ComposeTableDiffs`): rows align by base row anchor, cells pair positionally, and DISJOINT cross-reviewer cell edits compose inline (each cell authored to its reviewer). A cell edited by ≥2 reviewers RECURSES into the SAME body block/token composition over the cell's paragraph mini-body (`MergeBlockStream` over the per-reviewer cell `BlockOps`) — so disjoint words inside one cell paragraph fuse, and same-word edits become a cell-paragraph-anchored conflict resolved by the policy. The op's `Op.TableDiff` is the merged apply/JSON truth; an additive `AuthoredRows` carries the renderer/revision attribution view. Authored rows/cells tile the base table exactly once (`AssertTilesBaseTable`, enforced at **runtime** via `IrCompositeMerger.Invariant`, not a Release-stripped `Debug.Assert`). STOP boundaries fall back to a whole-table block conflict (see the v1 limitations below).
- **Anything else** (delete-vs-modify; a reviewer who changed both the paragraph's text AND its `w:pPr`; a table with a `MovedRow` or a column add/remove by some reviewer; mixed kinds) → a **block-level conflict** resolved by the policy. (The pPr gate is deliberate: the compose path clones the BASE paragraph's pPr, so a reviewer who also changed pPr would have that change silently dropped — routing such an op to the conflict path preserves and surfaces the reviewer's full edit instead. A table with a row move or a column-count change routes here too because the count-stable positional per-cell render cannot compose those — see the v1 limitations below.)

**Block-level inserts never conflict.** Every reviewer's right-only inserted block appears, slotted immediately after the shared base anchor it follows, ordered by reviewer index — two reviewers both inserting after the same paragraph both appear, attributed. A NATIVE move DESTINATION (see below) is right-positioned content too (null left anchor) and is routed the same way.

**Native move composition.** Before the by-base grouping, `PlanMoves` decides, per reviewer move group, NATIVE vs LOWER: a move group (reviewer R, source base anchor S) renders as a native `w:moveFrom`/`w:moveTo` iff `RenderMoves` is on AND R is the only reviewer that touches base block S (so the move does not collide with another reviewer's edit/move on S). Native groups are assigned **globally-namespaced** move-group ids (one deterministic counter, reviewers in list order then ascending local gid) so two reviewers' independent moves never share a `w:name`; `ApplyMovePlan` keeps the native move ops (rewriting the gid) and lowers everything else. The native move SOURCE rides its base anchor through the by-base path (sole-toucher → emitted verbatim, authored to the mover); the native move DESTINATION rides the preceding-anchor (right-positioned) path. COLLIDING moves — move-vs-edit on S, or two reviewers moving the same block — are NOT native: both LOWER to del/ins and record a conflict (see the contested-relocation note below). The markup/revision/JSON renderers already handle native moves with a per-op author override, so no renderer change was needed.

### Conflict model + the three policies

A conflict is a base span (a token span, or a whole block) edited DIFFERENTLY by two or more reviewers. The configured `ConflictResolution` decides what lands in the OUTPUT document; the conflict is **always recorded in the data regardless of policy** (`GetConflicts` / the `conflicts` JSON array / a `ConflictId` on the winning revision):

| Policy | Output at the conflicted span |
|---|---|
| `BaseWins` (default) | The base text is kept; every competitor is recorded. |
| `FirstReviewerWins` | The first reviewer (list order) is applied inline; the others are recorded. |
| `StackAll` | Each competing edit is emitted in reviewer order; all are recorded. |

A `DocxDiffConflict` carries `BaseAnchor` + the base **token span** `[TokenStart, TokenEnd)` (an empty interval = a block-level conflict), the `AppliedPolicy`, and the per-reviewer `Competitors` (each with `Author` and the `ResultText` that reviewer's edit would have produced — `""` for a deletion). Its `Id` matches the `ConflictId` on the winning `DocxDiffConsolidatedRevision`, so conflicts correlate to the revision actually placed in the document.

### Multi-author rendering + round-trip

The same renderer backs `Compare` and `Consolidate`: `IrMarkupRenderer` was extended with a per-op author override + per-reviewer source selection (each reviewer's composed INSERT token spans index THAT reviewer's right-token list, so the renderer carries each contributing reviewer's own right paragraph anchor). Single-document `Compare` output is **byte-unchanged** by this extension. The round-trip invariant generalizes `Compare`'s accept ≡ right / reject ≡ left: `RejectRevisions(Consolidate(...))` content-equals the **base** (rejecting every reviewer restores the base), and `AcceptRevisions(...)` content-equals the **policy-resolved composite** (e.g. base text at conflicted spans under `BaseWins`, the first reviewer's text under `FirstReviewerWins`).

### Before / after intuition

- **Two reviewers editing DIFFERENT sentences of one paragraph.** Base: *"The cat sat. The dog ran."* Alice edits the first sentence, Bob the second. Their two token diffs share the base token stream, so token-span composition fuses them into **ONE merged paragraph** carrying Alice's edit on sentence 1 and Bob's on sentence 2, each `w:ins`/`w:del` attributed to its own author — directly consumable in Word's reviewing pane. (Legacy `WmlComparer.Consolidate` would instead append two stacked, labeled, colored single-cell boxes after the original — a side-by-side juxtaposition for a human to eyeball, never an inline merge.)
- **Two reviewers editing the SAME word.** Base: *"the quick fox"*. Alice changes *quick*→*brown*, Bob *quick*→*slow*. The token spans overlap, so this is a recorded conflict at `[BaseAnchor, TokenStart..TokenEnd)`. Under the default `BaseWins`, the output keeps *quick* and `GetConflicts` returns one `DocxDiffConflict` with both competitors (Alice→*brown*, Bob→*slow*); under `FirstReviewerWins` the output reads *brown* (Alice inline) with Bob still recorded; under `StackAll` both edits emit in order. The conflict data is identical across all three policies — only the document body differs.

### Parity outcome — and why the deviation catalog is empty

`ConsolidateParityScoreboardTests` (`Docxodus.Tests/Ir/Diff/ConsolidateParityScoreboardTests.cs`) scores all **84** legacy `WmlComparer.Consolidate` corpus cases (WC001's 10 multi-reviewer rows + WC002's 74 single-reviewer rows): **84/84 reproduce-PASS, 0 deviations, 0 fails**, with the genuine-pass floor ratcheted at 84.

The headline finding the scoreboard records: **legacy `WmlComparer.Consolidate` is a juxtaposition/triage tool, not a merge.** It keeps the original document intact and, for every changed block, APPENDS that reviewer's labeled revised copy (wrapped in a colored single-cell table under the default `ConsolidateWithTable`) right after the original — even for a single reviewer. Its accepted body is therefore, per changed block, `[revisor label][reviewer's block][original block]`: a side-by-side juxtaposition with the revisor name as literal body text. The IR-native engine instead produces a true inline merge. Because the two engines emit categorically different document SHAPES by design, raw accepted-body-text equality is the wrong relation (it never holds — that mismatch IS the supersession). Parity is therefore measured by a **sound-semantics** metric over normalized body plaintext:

- **Single reviewer** → accept ≡ that reviewer's document, **char-exact** (the same accept ≡ right contract `Compare` obeys; holds for all 74 WC002 rows).
- **Multi-reviewer, no conflict** → every reviewer's ADDED tokens appear in the accepted body (added-token containment, not whole-body subsequence — composition fuses adjacent edited tokens).

The whole-corpus juxtaposition-vs-inline-merge shape divergence is the deliberate supersession, quantified once in the scoreboard footer rather than catalogued as 84 identical per-row entries. A row is catalogued as a per-row **deviation** only when `GetConflicts` reports a TRUE cross-reviewer token-overlap conflict (two reviewers editing the same span differently) — and **no legacy-corpus row produces one** (every corpus edit is single-reviewer or disjoint-span), so the per-row catalog is empty for this corpus. The conflict-supersession path is instead exercised by the unit suites (`DocxDiffOpsConsolidateTests` / `DocxDiffConsolidateApiTests`) and the K-way composite fuzzer (`CompositeFuzzTests`: round-trip reject ≡ base + the `IrCompositeVerifier` apply-verifier over 3/4/5-way seeds).

### v1 limitations (honest)

- **Note scopes are not merged.** The merger does not build composite note-scope (footnote/endnote) ops; `IrCompositeScript.NoteOps` is always null. A reviewer's footnote/endnote edit is not yet consolidated — a follow-on task. To stop this being a *silent* loss, the merger **fails fast** (`NotSupportedException`) when an **N≥2** consolidate has a reviewer who edited note content (a non-`EqualBlock` note op — a pure id-shift from renumbering is ignored): there is no single-call fallback for multi-reviewer note merging, so a loud error beats dropping the edit. A **single-reviewer** consolidate is NOT hard-failed (it degrades to a body-level merge that omits the note edit; use `DocxDiff.Compare` for full single-reviewer note fidelity — the established single-reviewer body-parity corpus exercises exactly this shape). The `IrCompositeMarkupRenderer` `NoteOps`-empty tripwire is now a **runtime** guard (was a Release-stripped `Debug.Assert`) so a future change that populates `NoteOps` without renderer support also fails loudly.
- **Multi-reviewer table edits compose per-cell, with three documented fallbacks.** DISJOINT cross-reviewer table-cell edits now COMPOSE inline (Alice edits cell(0,0), Bob edits cell(1,2) → both land, attributed; disjoint words inside one cell paragraph fuse via the recursion); only edits to the SAME cell by ≥2 reviewers become a recorded conflict resolved by the policy. Three STOP boundaries fall back to a whole-table **block conflict** (no silent loss — the base table is kept under `BaseWins` and the disagreement is recorded under every policy): (1) a reviewer **`MovedRow`** (a row relocated off the spine — not composable per-cell in v1), (2) a **column-count change** (a column add/remove surfaces as an unpaired cell op; the positional per-cell render is count-stable in v1), (3) cell-shell-property changes the IR does not model (a pure `w:tcPr` width / gridSpan / vMerge change leaves every IR hash identical, so the reviewer's table reads as `EqualBlock` — no touch — and never reaches the compose branch; such a change is invisible rather than dropped). A reviewer who ALSO changes a cell's text composes normally. Note: a 2-way column ADD is not yet marked as a tracked insertion (a pre-existing whole-table-renderer gap), so a column-add fallback under `FirstReviewerWins`/`StackAll` rejects to the reviewer's column-changed table rather than to base — independent of the consolidate path.
- **Non-colliding reviewer moves render natively; colliding moves are lowered to del/ins as a recorded conflict. Splits/merges still lower.** A SINGLE reviewer's `MoveBlock`/`MoveModifyBlock` whose source base block is touched by ONLY that reviewer renders as a native `w:moveFrom`/`w:moveTo` (`PlanMoves`/`ApplyMovePlan`, globally-namespaced move-group ids; see "Native move composition" above) — authored to the mover, with `Moved` source+destination revisions and a `MoveBlock` pair in the edit-script JSON. A COLLIDING move (a move-vs-edit on the same base block, or two reviewers moving the same block) is LOWERED: `LowerStructuralOps`/`LowerOneStructuralOp` rewrites the move SOURCE → `DeleteBlock` (retaining `MoveGroupId`/`IsMoveSource` as a relocation marker, stripped before emission) and the move DEST → `InsertBlock`. `SplitBlock`/`MergeBlock` are always lowered (a split → `DeleteBlock` + N ordered `InsertBlock`s; a merge → N ordered `DeleteBlock`s + an `InsertBlock`), preserving op order. Lowering is required for correctness: a null-left-anchor structural op (a move destination or a merge) reached neither the by-base grouping nor the preceding-anchor insert routing and was previously **silently dropped, losing content on accept** (B5). A lowered move/split/merge shows as `w:del`/`w:ins` — **content is fully preserved and round-trips** (accept shows the moved/merged text, reject ≡ base). Native cross-reviewer split/merge composition (preserving split/merge markup across reviewers) is a follow-on. Two reviewers relocating the SAME base block to different places collide on the lowered source-delete → a recorded **placement conflict** (the consensus removal is emitted once for `reject ≡ base`; each reviewer's relocating insert survives independently — no loss; this conflict is NOT policy-resolved into the op stream, so a `BaseWins` flip never wrongly restores a block both reviewers removed).
- **A reviewer who changed BOTH a paragraph's text AND its paragraph properties (pPr)** routes that block to the conflict path (so the pPr change is preserved/surfaced by the policy, never silently dropped) rather than into token-span composition.
- **Conflict spans are base TOKEN indices** (`TokenStart`/`TokenEnd`) + `BaseAnchor`, not character offsets — suitable for machine consumers inspecting the edit script.

## Parity status

The engine was developed against `WmlComparer` as the oracle under a binding method rule: WmlComparer presumed correct per gap; the IR is fixed to match unless an oracle fault is established with concrete evidence. As of M2.6 (this surface):

- **`GetRevisions` parity: 179/179 genuine count-exact PASSES — the documented-deviation catalog is EMPTY.** M2.6 closed the last two deviations (WC-1450/WC-1830, the 1:N sub-paragraph split — see the section below); the scoreboard asserts `Deviation == 0` and ratchets the genuine-pass floor at 179.
- **Produced-markup parity:** floor 39 fixtures round-trip clean (accept ≡ right, reject ≡ left, schema-valid); the round-trip allowlist is **1 fixture** (WC-BodyBookmarks endnote→footnote whole-note-store conversion, on which the WmlComparer oracle itself throws — there is no oracle behaviour to match). M2.6 Task 2 closed WC022 (the `InOrderRefine` same-unid identity reservation); the split/merge markup (below) round-trips on both corpus split fixtures and the synthetic split/merge shapes.
- The only remaining artifact anywhere is that single oracle-crashes allowlist fixture.

### 1:N paragraph split / N:1 merge (M2.6) — the implemented algorithm

One before-paragraph whose content migrates across N after-paragraphs (the user pressed Enter mid-paragraph), or the reverse merge, is a first-class engine capability. Design: `docs/superpowers/specs/2026-06-12-subparagraph-split-merge-design.md` (DESIGN-RESOLVED + adversarial review); this section records what the code DOES — the source of truth is `IrBlockAligner.DetectOneToManyInGap`/`FindQualifyingRun`/`TrimAndGate` + `IrSplitSegmenter`.

**Detection (in `FillOneGap`, after the unambiguous-table-residue rule, before the 1×1-residue rule; gated by `IrDiffSettings.DetectSplitMerge`, default ON).** One side-parameterized worker runs twice per gap — split (singular = left) first, then merge (singular = right); a block consumed by a split group is never reconsidered by the merge scan.

1. **Candidates.** A gap paragraph on the singular side qualifies iff it is still FREE, or was Modified-paired by THIS gap's `SimilarityPair` to a plural-side paragraph inside the gap (the pairing is *promoted* if a window qualifies). `Unchanged`/`FormatOnly`/`Moved` blocks are NEVER candidates: a content-equal pair has zero unmatched tail, so promoting one could only manufacture a false split — this is what preserves the WC022 identity-reservation reject-order invariant (review finding F4.2; regression-tested both directions with detection on).
2. **Window enumeration.** For each candidate, ascending start × ascending end over maximal CONTIGUOUS runs of eligible plural indices (free paragraphs, or the candidate's own partner), capped at `SplitMaxRunLength` (8). The first window that passes all gates wins — shortest-first is deliberate: the smallest window clearing the coverage bar absorbs the least foreign content.
3. **O(1) length prefilter.** Before any scoring, a window is skipped unless its cached content-token total lies in `[SplitCoverageThreshold × singularContent, singularContent / (1 − SplitForeignSlack)]` — the thresholds' arithmetic implications. Growing a window only adds content, so exceeding the upper bound breaks out of the end-loop. This is what keeps a fully-rewritten G×G gap at G²·O(1) instead of G²·O(LCS) (the adversarial 200×200 fixture's 5-second bound).
4. **Scoring (`IrSplitSegmenter.Score`).** One in-order LCS (standard DP, deterministic back-walk tie-break) of the singular paragraph's full token stream against the window's concatenated streams. `Coverage` = LCS-matched singular CONTENT tokens / singular content tokens; `ForeignSlack` = unmatched window content tokens / window content tokens (content = non-Separator, non-Textbox; separators participate in the LCS for boundary context but never score).
5. **Edge trim (the false-positive guard, review R2).** Leading and trailing members with ZERO matched content are dropped, then the trimmed window is re-scored. This excludes an unrelated net-new edge neighbor and edge empty carriers, while keeping INTERIOR net-new members (WC-1830's inserted math paragraph between the two halves) — their foreign content is priced by the slack gate.
6. **Fire gates (on the trimmed window).** ≥2 members carrying at least one content token; a paired candidate's partner inside the window plus ≥1 other free member; `Coverage ≥ SplitCoverageThreshold` (0.90); `ForeignSlack ≤ SplitForeignSlack` (0.34). The thresholds are corpus-swept (`IrSplitThresholdSweepTests`): the shipped pair sits on a plateau at the grid maximum with ≥1 full grid step of margin on every axis (the F4.1 gate, re-asserted on every run).
7. **On fire.** Any prior Modified pairing is overwritten; every member's kind/match slots are stamped immediately (no window may reuse a consumed block — the F2.2 overlap ceiling); the group is recorded and its indices leave the leftover lists, so the 1×1 rule and surplus classification see only what remains. `EmitEntries` collapses each group to ONE alignment entry (`IrAlignmentKind.Split` at the FIRST member's right position carrying `MultiBlocks`; `Merge` at the right block's position), with deletion buckets flushed exactly once per anchored left.

**Segmentation (`IrSplitSegmenter.ComputeSegmentDiffs`, at projection time).** The same LCS assigns every singular token to a member: a matched token goes to its partner's member, an unmatched token to the nearest PRECEDING matched token's member (leading unmatched → member 0) — a total, monotone rule. Each slice is re-diffed against its member with the ordinary Myers differ, so every segment diff carries the full token-diff invariant battery over (slice, member) and the partition invariant holds structurally. For a merge the segmenter runs with singular = right, and the builder mirrors each diff (`MirrorDiff`: Insert↔Delete + span swap) so stored diffs always read left → right.

**Surfaces.** *Apply-verifier:* a `SplitBlock` pushes one reconstructed tuple per member (the existing count/order/`ReferenceEquals` loop then proves the N rights sit contiguously at the op's position); a `MergeBlock` pushes one tuple reconstructed from the N members; the cell/note path additionally asserts the produced right-anchor SEQUENCE equals the right block list (the F3.2 strengthening, asserted corpus-wide). *Revisions:* Fine mode reports each segment's token diff plus one `Inserted "\n"` per new pilcrow (`Deleted "\n"` per removed one) — a clean split is visible but claims no content change; compatible mode reproduces the oracle's account — segment 0's inline edits (seam-whitespace-only ins/del suppressed) plus exactly ONE coalesced `Inserted` (`"\n" + Σ(memberText + "\n")`) per split — which is what lands WC-1830 at 2 and WC-1450 at 7. *Markup:* the anchored-split shape — N paragraphs, each member's pPr and segment content, with `MarkParagraphMark` inserted marks on paragraphs 0..N−2 (deleted marks for a merge, whose non-final paragraphs keep their LEFT pPr); REJECT removes the marks and RevisionProcessor re-fuses the paragraphs, reconstructing LEFT; ACCEPT yields the N right paragraphs. *Wire:* additive optional arrays only. *Fuzzer:* `SplitParagraph`/`MergeParagraphs` mutation kinds run the own-oracle battery (apply-verify + JSON round-trip + determinism) on every seed; they are excluded from the cross-engine differential class because the engines frame a clean split differently BY CONSTRUCTION (WmlComparer reports the tail as a Deleted+Inserted pair of identical text; the IR keeps it Equal and reports the structural mark — the `RelocateParagraph` precedent).

**Deltas vs the spec worth knowing.** (a) The spec's §3.2 "anchored-split" example cell (WC-1450's `Second `-prefix cell) turned out NOT to be a split at all — its before-paragraph never contained the tail (`Score` member-match probe: `[11, 0]`); the oracle's `Inserted "Second "` + `Inserted "When you click…"` is the ordinary Modify + InsertBlock account, which the edge trim correctly preserves. (b) The implemented mark placement (inserted marks on 0..N−2, last paragraph keeps the original pilcrow) differs from the spec's §3.2 oracle excerpt but satisfies the same accept ≡ right / reject ≡ left contract the spec adjudicates on (§3.3/§5.5 explicitly allow this). (c) There is no `IrSegmentDiff` wrapper record (review F1.3) — `segmentDiffs` is a plain token-diff list. (d) Scope ceilings: N:M and cross-gap splits never fire (one singular side, one gap, by construction); `DetectSplitMerge = false` restores strict 1:1 op semantics.

## Relationship to WmlComparer

| | `WmlComparer` | `DocxDiff` (IR engine) |
|---|---|---|
| Status | Default / blessed | Production-candidate (D4 open) |
| Comparison substrate | Atom streams | Modeled IR (blocks/runs/format records) |
| Revisions | `WmlComparerRevision` (OOXML members, no anchors) | `DocxDiffRevision` (anchor-addressed; no OOXML members) |
| Move markup | `GetRevisions`-only post-process; **native `w:moveFrom`/`w:moveTo` IS produced** by the IR markup renderer | native `w:moveFrom`/`w:moveTo` |
| Format change | detected + described (`w:rPrChange`) | detected + described (modeled-only by default) |
| Diff-as-data | none | edit-script JSON |
| Determinism | wall-clock dates by default | deterministic by default |

> **Note for readers of `wml_comparer_gaps.md`:** that document's older "native move markup is not generated" / "format change detection is a gap" claims were stale (both shipped in the v6.x line and are produced by `DocxDiff`'s markup renderer here). The gaps doc has been corrected and points here.

## Cross-layer ripple

The four-layer ripple (WASM bridge → npm/TypeScript → python host/`docx_scalpel`) for these three entry points is tracked as M2.5 Task 5 (see the program plan). This document covers the .NET public surface.

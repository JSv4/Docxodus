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

A `ModifyBlock` over a paragraph carries a `tokenDiff`; over a table, a `tableDiff` (row ops with nested cell ops); a textbox-bearing block carries `textboxDiffs`. The JSON is a faithful serialization of this structure (top-level `operations` + optional `noteOps`), and is deterministic for identical inputs.

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

### Two honest defaults that deviate from `WmlComparerSettings`

1. **Deterministic dates.** `WmlComparerSettings.DateTimeForRevisions` defaults to `DateTime.Now` — the same compare twice yields different dates. `DocxDiff` pins a fixed epoch by default so output is reproducible. Opt into wall-clock via `Deterministic = false`.
2. **`FormatComparison = ModeledOnly`.** A `w:rPrChange`-grade report can only DESCRIBE modeled fields, so a format change driven by an undescribable unmodeled-only rPr flip (`w:lang`, `w:bCs`, complex-script toggles) is noise. `ModeledOnly` collapses that noise; the trade-off is a false negative on a visible-but-unmodeled change (e.g. `w:shd` run shading). `Full` restores byte-fidelity comparison.

## Parity status

The engine was developed against `WmlComparer` as the oracle under a binding method rule: WmlComparer presumed correct per gap; the IR is fixed to match unless an oracle fault is established with concrete evidence. As of M2.5 (this surface):

- **`GetRevisions` parity:** 174 PASS + 5 documented deviations (floor 179/179). Each surviving deviation carries established root-cause evidence (engine-alignment-grain, tokenizer-grain, shared `RevisionProcessor`, note-store cross-part conversion); four of the five explicitly establish the ORACLE correct.
- **Produced-markup parity:** floor 39 fixtures round-trip clean (accept ≡ right, reject ≡ left, schema-valid); the round-trip allowlist holds at 5 fixtures, all with established reader/aligner-level root causes (none a renderer-markup fault).

The end-state on the deviation/allowlist sets is **evidence-retained, not zero**: the remaining items are cases where closing them would require either changing shipped-converter output or a real engine capability (the 1:N sub-paragraph split — see below), not a patch. They are catalogued in the M2.4b/M2.5 plan outcomes and the deviations catalog.

### Deferred follow-on (M2.6)

The dominant retained deviation (WC-1450 / WC-1830) is a **1:N paragraph split**: one before-paragraph's content migrates across two after-paragraphs. The oracle's flat atom LCS credits both halves as Equal against the single before-paragraph; the IR's `IrEditOp` is strictly 1:1, so it surfaces the second half as a whole-paragraph revision (+1). This is not closable in the 1:1 model and not render-coalescible; the correct fix is an engine-level 1:N split/merge op with apply-verifier/markup/JSON/fuzzer ripple — a real capability, sketched and deferred. See `docs/superpowers/specs/2026-06-12-subparagraph-split-alignment-sketch.md`.

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

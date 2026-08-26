# PR #355 — "Add draggable editor blocks with tracked moves" — smoke test

Branch `investigate/block-drag-handles` @ `a2e7232` (== PR head, merge-base == `origin/main` `f290c23`).
Fixture: NVCA Model Certificate of Incorporation (10-1-2025) — 234 top-level body blocks,
392 bookmarks, 94 footnotes, 4 `w:sectPr` (3 inline section breaks at block 23/217/221), 0 tables.
Full `npm run build` + `npm run pretest` from PR source; served on a fresh port; Chromium.

## Verdict

The model layer is solid — every move round-trips exactly. Two defects worth fixing before merge:
one **correctness** bug in tracked mode (B1) and one **confirmed-direction** UX gap (B2).

## STATUS: all findings closed on this branch

| Finding | Resolution |
|---|---|
| B1 bookmarks | `DocxSession.RenumberClonedBookmarks` — fresh ids on the clone, names kept (mirrors `NormalizeBookmarks` (B)) |
| B1 comments | `CommentOps.CloneCommentsForMoveSource` — fresh id + cloned definition for the move source, threading parts kept consistent (mirrors `NormalizeComments` (B)) |
| B1 footnote refs | Deliberately left duplicated — not uniqueness-constrained, faithful to a pending move, resolves to one citation. Documented. |
| B2 | `DocxSession.ValidMoveTargets` + editor gating: no indicator on a doomed target, menu items disabled, no handle on an unmovable block, **"move to top/bottom" now works within the source's region**, friendly announcement. Reports each SIDE separately — verified live on NVCA, where 2 of a block's 3 reachable targets are legal on only one side |
| B3 | `MarkParagraphContentAndMark` — cell content marked, and `w:moveFrom` runs now carry `w:delText` |
| B4 | `RevisionProcessor` coalesce keeps the surviving paragraph's attributes (and so its `pt:Unid`) |
| Autoscroll / silent drop | Registered on the real scroll container; an outside drop is announced |
| Stale doc header | Design doc flipped to "shipped" + a deviations/follow-ups section |

Re-measured after the fixes (same NVCA fixture, same script that found them):

```
BASELINE (pristine NVCA) validation errors = 65
tracked paragraph move (bookmarks only): errors 65 (delta 0) PASS   was 77
tracked paragraph move WITH a comment  : errors 65 (delta 0) PASS   was 80
REJECT === ORIGINAL order: True      ACCEPT === expected moved: True
duplicate bookmark / commentRange / commentRef ids: 0 / 0 / 0
```

.NET suite 3367 passed / 0 failed; Release build 0 errors; LibreOffice renders both outputs.

## What works (verified end-to-end)

| Check | Result |
|---|---|
| Direct-mode paragraph drag | Correct order; **DOM node identity preserved**, `lastReconcileFallback: null` (true incremental reconcile) |
| Direct-mode latency | ~835 ms incl. simulated human drag |
| Undo / redo | Exact restore + reapply; rejected moves consume **no** undo slot |
| Save → order | Pure reorder, 234 blocks, multiset identical, **zero** revision markup |
| Part / reference fidelity | 0 parts lost, 392/392 bookmarks, 94/94 footnote refs |
| LibreOffice | Opens + renders the new order correctly |
| Tracked paragraph move | Named pair `move1001`, paired ranges, author honored, `w:trackRevisions` set; read-only move-from + editable move-to |
| **Accept ≡ right / reject ≡ left** | `REJECT === ORIGINAL order: True`, `ACCEPT === expected moved order: True`; exactly one copy each |
| Tracked table move | Source rows `trPr/del`, dest rows `trPr/ins`; accept → destination only, reject → original position; **0 new validation errors** |
| Cell hover → whole table | Correct (`srcTag: TABLE`); no cell-level handle |
| Text selection | Unaffected — 56-char in-block selection, drag not hijacked |
| Escape on move menu | Closes, document unchanged, focus restored to handle |
| Footnote/endnote blocks | Not in `bodyUnitNodes()` → no handle (correct) |
| Paginated mode | Handle not created at all (correctly gated) |
| Console | **Zero real errors** across the whole session |
| PR's own specs | `editor-block-drag.spec.ts` + `editor-reconcile-unit.spec.ts` → **20/20 pass** locally |
| **Selective** accept/reject (paragraph) | `ListRevisions()` → **one** entry `type=move`; `AcceptRevision` resolves **both** sides (moveFrom/moveTo/ranges → 0, 1 copy at index 9, 234 blocks, 0 revisions left); `RejectRevision` → index 5. Exactly the design-doc "one group" requirement |
| **Selective** accept/reject (table) | 2 entries (`delete` + `insert`) — as designed (direction #8: Word presents these separately); accepting both → 1 table at index 21 |
| **Undo/redo after a TRACKED move** | Exact and idempotent: 235/16 moveFrom/16 moveTo → undo → 234/0/0 → redo → 235/16/16 → undo → 234/0/0; no residual markup |

## B1 — Tracked move clones identity-bearing children (correctness)

`DocxSession.MoveBlock` tracked path does:

```csharp
var destination = new XElement(source);
foreach (var el in destination.DescendantsAndSelf())
    el.Attributes(PtOpenXml.Unid).Remove();   // strips Unid only
```

Everything else in the clone is duplicated verbatim. Measured on one paragraph move:

- **6 duplicate `w:bookmarkStart/@w:id`** (4–9) → **12 new OpenXmlValidator errors** (65 → 77).
  Schema requires the id to be unique.
- **6 duplicate bookmark `@w:name`** (`_DV_M7`, `_DV_C8`, `_DV_M8`, `_DV_C10`, `_DV_M9`, `_DV_M10`)
  → cross-references resolving to those names become ambiguous.
- **`w:footnoteReference w:id="2"` duplicated** → body footnote refs 94 → 95, so note numbering
  is wrong while the revision is pending.

**Comments duplicate too — measured**, not inferred. `AddComment` on a paragraph, then a tracked
`MoveBlock` of it, then save:

```
commentRangeStart ids = [1,1]   duplicates=1
commentRangeEnd   ids = [1,1]
commentReference  ids = [1,1]
```

So one comment ends up with two ranges and two references — it appears twice in Word's Reviewing
pane and anchors to both the source and the destination. `BlockMoveSafetyError` only rejects
*cross-block* ranges, so an intra-block comment sails straight through.

Scope: it corrupts only while the revision is **pending** — accept and reject both return to the
document's own 65-error baseline. But pending is exactly the state you send to a counterparty.
LibreOffice tolerates it; Word's stricter handling is **unverified here**.

Why CI missed it: every test in `editor-block-drag.spec.ts` uses a blank document with three
typed paragraphs — no bookmarks, comments, or notes. The `.NET` `DocxSessionMoveBlockTests` are
likewise synthetic. A fixture with bookmarks would have caught this immediately.

Fix direction: on clone, renumber/strip `bookmarkStart/End` ids + names, comment range ids, and
decide the footnote-citation policy — same single-owner treatment the Unid strip already gets.

## B2 — Invalid drop targets still show a drop indicator (confirmed direction #7 unmet)

Design doc, confirmed v1 direction #7: *"Invalid targets should not display a drop indicator … hide
invalid target indicators rather than allowing a doomed drop."*

`refreshBlockDropTargets()`'s `canDrop` only checks
`source.data.type === BLOCK_DRAG_TYPE && sourceAnchorId !== targetAnchorId`. It never consults
section-break or cross-block-range validity, and `onDragEnter`/`onDrag` call `showDropIndicator`
unconditionally.

Measured: dragging block 18 → block 30 (across the section break at 23) **showed the indicator**,
then the drop failed. The model correctly refused (order unchanged) — this is UX, not corruption.

Blast radius on this document (3 inline section breaks → 4 drag islands):

- **"Move to top" fails for 211 of 234 blocks** (blocks 23–233 rejected; 1–22 succeed; block 0 is a
  no-op). "Move to bottom" fails for 222 of 234 (blocks 0–221).
- All four menu items are **always enabled** (`disabled: false`) regardless of validity.
- The section-break paragraph itself reports `isMovableBlockUnit: true`, so it gets a handle
  it can never successfully use.
- The announcement leaks the raw engine string: `"cannot move a block across a section-break paragraph"`.

## B3 — Tracked table move marks rows but not cell content (spec divergence, minor)

Design doc §Whole tables: *"every source row **and its content** is deleted, and every destination
row **and its content** is inserted."*

`MarkTableRowsAsTrackedRevision` only adds `w:ins`/`w:del` to each `w:trPr`. Measured on the saved
file: `content del=0 ins=0 delText=0` for both tables. The deleted table's text is therefore not
struck through — the editor's own renderer does distinguish it (red row background), and
accept/reject round-trips correctly, so this is fidelity, not correctness.

## B4 — Accepted-view render plan carries one phantom unit after a tracked move (latent)

Reopen the saved tracked document with `trackedChanges: 0` (accepted view) in the live editor:

| Source | Body units |
|---|---|
| DOM (`bodyUnitNodes()`) | **233** |
| `ListRenderedBlocks(handle, false)` | **234** |
| `ListRenderedBlocks(handle, true)` | 235 |
| Model `RevisionProcessor.AcceptRevisions` block count | **234** |

Exactly one plan unit (`p:body:5a67f4a5…`, plan index 5) has no DOM node. The plan agrees with the
model; the accepted-view **renderer** is the one producing a block fewer, so `IsRemovedInAcceptedRevisionView`
(new in this PR) is not the party that's wrong — the two views of "what survives accept" disagree by one.

No visible breakage observed: a following structural op still reconciled incrementally
(`lastReconcileFallback: null`, units tracked +1). But the reconciler is permanently diffing against
a plan with one unit the DOM can never contain. Worth a look before this ships; not a blocker on the
evidence I have.

## Minor

- `autoScrollForElements({ element: this.editRoot })` registers on `.docx-body-flow`, which is not
  the scroller (the real one is `.dxr-scroll`, `clientH 595 / scrollH 32834`). Pragmatic logs
  *"Auto scrolling has been attached to an element that appears not to be scrollable"* on **every**
  document open. Scrolling during drag still works via the window-level/native path (measured:
  scrollTop 24.8 → 682 during a dwell, drop succeeded), so this is a dead registration + log noise.
- A drop released outside any valid target is a **silent** no-op — no announcement at all.
- Tracked moves cost **~5.2–6.2 s** on this document because the tracked path always calls
  `remount()`. Direct moves are ~0.8 s incremental.
- No ribbon UI toggles track changes; `trackedChanges`/`revisionAuthor` are mount-time options only.
- `docs/architecture/editor_block_drag_handles.md` still says
  *"Status: investigation / design … (no feature implementation yet)"* — stale now that it shipped.

## Not a regression (checked against main)

- Ctrl+A selects one block, not all. The PR touches no keyboard path
  (`git diff f290c23..HEAD -- npm/src/editor.ts` has no `selectAllBlocks`/`onKeydown` hits, and
  `selectAllBlocks` exists on neither main nor PR head). Pre-existing, out of scope.

## Artifacts

`.playwright-mcp/`: `nvca-direct-move.docx`, `nvca-tracked-move.docx`, `nvca-tracked-table.docx`,
`nvca-accepted.docx`, `nvca-rejected.docx`, `tracked-move.png`, `tracked-table-move.png`,
`lo-out/*.pdf`. Verifier: `/home/jman/.cache/docxodus-smoke-verify` (outside the repo).

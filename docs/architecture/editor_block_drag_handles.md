# Editor block drag handles — design and shipped behaviour

Status: **shipped** — direct and tracked moves in continuous view, with the
accessible move menu. Paginated dragging remains deferred.

Branch: `investigate/block-drag-handles`, based on `origin/main` `f290c23`
(2026-08-05). The sections below are the design as built; "Shipped deviations
and follow-ups" at the end records where the implementation went further than,
or stopped short of, this plan.

## Verdict

Adding an Atlassian-style handle for reordering editor blocks is feasible without
changing the editor's model-of-record. The UI half is a contained addition to
`DocxEditor`; the missing prerequisite is an **atomic, anchor-addressed `MoveBlock`
mutation on `DocxSession`**. Because moves must produce native Word revisions when track
changes is enabled, the editor's currently accepted-only render/reconcile path also needs
to become revision-view-aware.

The feature must not reorder only the DOM. The DOM is a projection; `save()` serializes
the live OOXML session, so a DOM-only move would disappear on save and undo. The correct
flow is:

```text
handle drag / accessible move command
  -> source anchor + target anchor + before|after
  -> DocxSession.MoveBlock (one OOXML mutation, one undo snapshot)
  -> ordered-unit reconcile (or paginated/list fallback remount)
  -> focus restored to the moved block
```

This is a medium-sized editor feature rather than a rewrite. Recommended scope and
effort:

| Deliverable | Scope | Estimate |
|---|---|---:|
| Interaction spike | Pointer-only, continuous view, paragraphs, remount after drop | 1–2 engineering days |
| Direct-move foundation | Paragraphs/headings/list items + whole tables, continuous view, keyboard/menu alternative, undo/redo, autoscroll, incremental reorder, tests | 5–8 engineering days |
| Native-revision integration | Review-aware editor projection, named paragraph moves, tracked whole-table relocation, author/date, accept/reject and save/reopen tests | +4–7 engineering days |
| Confirmed production scope | Direct moves plus native revisions in continuous view | **9–15 engineering days total** |
| Paginated hardening | Cross-page targeting, scroll, full reflow/remount, focus restoration | +2–4 engineering days |

The estimates assume one engineer familiar with this repository and include the required
bridge/client ripple and browser tests. Cross-block OOXML ranges may add time if v1 must
preserve their exact membership rather than rejecting unsafe moves (see “OOXML safety”).

## Confirmed v1 direction

1. Reorder **top-level body render units only**: `p`, `h`, `li`, and a whole `tbl`.
   Never expose a separate handle for a table-cell paragraph; hovering/focusing inside a
   table targets the table unit.
2. Ship **continuous view first**. Paginated dragging is follow-up scope and should do a
   full session-attached remount after drop until scoped repagination exists (roadmap M4).
3. Follow Word editing semantics: with track changes off, directly reorder the existing
   element; with track changes on, emit native Word revision markup with the configured
   author/date. Paragraph-like blocks use a named `w:moveFrom`/`w:moveTo` pair. A whole
   table uses Word's table-row insertion/deletion revision vocabulary unless a Word-made
   fixture demonstrates a conformant named whole-table move shape (details below).
4. Keep the framework-agnostic `DocxEditor` backward compatible with a new option such as
   `blockDrag?: boolean` (default `false`); the full `mountRibbon` surface can default it
   to `true`.
5. Use the handle as both a drag source and a button. Click/keyboard opens a small move
   menu (`Move before…`, `Move after…`, `Move to top`, `Move to bottom`) and completion is
   announced through an `aria-live` region. Pointer dragging is not an accessible
   substitute for those commands.
6. Use Atlassian's Pragmatic drag-and-drop runtime plus its optional autoscroll package;
   Docxodus continues to own the handle, drop indicator, menu, and live-region UI.
7. Reject moves that would change or invert pre-existing cross-block comment, bookmark,
   permission, or native-move ranges. Invalid targets should not display a drop indicator.
8. For whole tables in track-changes mode, accept Word-native source deletion plus
   destination insertion even when Word presents them as separate revisions rather than
   one named “Moved” revision.

## What already exists

The current architecture supplies most of the feature:

- Every rendered paragraph/heading/list item/table has `data-anchor`; `DocxEditor` maps it
  to the session's full `kind:scope:unid` address.
- `DocxSession` owns lossless OOXML mutation, bounded undo/redo, and transactional
  pre-operation snapshots.
- `ListBlocks` already returns body units in document order, flattening block-level
  content controls to mirror the renderer.
- `DocxEditor.reconcile()` already diffs the DOM's ordered units against `ListBlocks`,
  batch-renders changed units, and preserves untouched DOM nodes.
- `bodyUnitNodes()` already collapses cell paragraphs into one table unit—the exact list
  the drag UI needs.
- Continuous mode can patch structurally; paginated mode already falls back to a full
  remount, which is correct for cross-page reflow.
- The editor already owns document-level mouse handlers for cross-contenteditable text
  selection. A handle outside the editable block avoids conflicting with that path.

The notable gaps are:

- no `DocxSession.MoveBlock` / bridge command;
- the unit diff classifies a pure reorder as remove+add/substitution, and
  `applyBodyDiff()` explicitly rejects kept nodes appearing in a different order;
- no handle/drop-indicator UI, drag lifecycle, autoscroll, or accessible alternative;
- no move-specific OOXML safety policy;
- `DocxEditor` hard-codes an accepted revision view while `ListBlocks()` enumerates raw
  XML blocks. A tracked move creates source and destination blocks, so the HTML and render
  plan would diverge after the first move unless both become revision-view-aware.

## Model mutation

Add this primitive at the core, then expose the same operation through every session
transport:

```csharp
public EditResult MoveBlock(
    string sourceAnchorId,
    string targetAnchorId,
    Position position); // Before | After
```

### Contract

- Source and target must resolve, be block-level kinds (`p`, `h`, `li`, `tbl`), live in
  the same package part, and share a direct XML parent.
- Moving a source relative to itself, or to the position it already occupies, is a
  successful no-op and records no undo snapshot.
- A real move records exactly one pre-op snapshot. In accepted/direct mode it detaches
  the existing `XElement` and inserts that same element before/after the target. It must
  not clone/reparse XML in this mode.
- In direct mode, the existing element and all descendant Unids remain unchanged.
  Hyperlink, image, comment, note, and other relationship IDs remain valid because v1
  only moves within one part.
- In tracked mode, preserve a revision-marked source at the old location and add a
  revision-marked destination at the new location. The two live elements need distinct
  anchor identities while both revisions are visible; `EditResult` must return the
  destination anchor so focus follows the logical moved block.
- Invalidate the projection/index cache and return the current source and target anchors
  in `EditResult.Modified`. Returning both lets the client recognize list/contextual
  repaint requirements.
- On failure, restore the snapshot and return the existing typed error envelope.
- A paragraph carrying an inline `pPr/sectPr` should be rejected in v1. It represents a
  section boundary, not an ordinary visual block; moving it has surprising header,
  footer, margin, and page-numbering effects.

The same-parent rule intentionally rejects a move into/out of a block-level `w:sdt` even
though the HTML renderer visually flattens its contents. Moving across that boundary
would change content-control semantics. The UI can initially surface the core error; a
polished version should add container identity/movability to `ListBlocks` so it never
draws an invalid drop target.

### Required layer ripple

Per the repository's single-owner session architecture, the operation should be added to:

1. `Docxodus/DocxSession.cs` and `Docxodus.Tests/DocxSessionTests.cs`.
2. `Docxodus/Internal/DocxSessionOps.cs`.
3. `wasm/DocxodusWasm/DocxSessionBridge.cs`.
4. `tools/python-host/Dispatcher.cs`.
5. `npm/src/types.ts`, `npm/src/session.ts`, and `npm/src/editor.ts`'s narrow bridge type.
6. `python/src/docx_scalpel/session.py`.
7. `tools/mcp-server/ToolCatalog.cs` and `tools/mcp-server/Dispatcher.cs` if the generic
   agent edit surface is intended to keep parity (recommended).
8. `docs/architecture/docx_mutation_api.md` and the public npm/Python API docs.

No new response type is required; the existing `EditResult` wire shape is sufficient.

## Native Word revision semantics

`DocxSession` already has most of the revision plumbing this operation needs:

- mutable `TrackedChangeMode` and revision-author session settings;
- monotonic revision IDs and author/date envelopes;
- automatic `w:trackRevisions` document setting;
- native paragraph move emission in `IrMarkupRenderer`;
- grouped accept/reject behavior for named `moveFrom`/`moveTo` pairs; and
- whole-table insert/delete marking that round-trips through accept/reject.

The new mutation should reuse or extract those helpers rather than create a second OOXML
revision dialect.

### Paragraphs, headings, and list items

When `TrackedChangeMode.RenderInline` is active, retain a move-source paragraph at the old
position and create a move-destination paragraph at the requested position. Use one named
move group with paired `w:moveFromRangeStart`/`End` and
`w:moveToRangeStart`/`End`; mark both run content and paragraph marks as moved. This is the
same native shape the existing IR diff renderer emits and lets selective accept/reject
resolve the pair as one logical move.

The OOXML definition is explicitly paragraph/run based: Microsoft's SDK documentation
calls `w:moveFrom` in paragraph properties a “Move Source Paragraph,” and the move-range
documentation describes the moved pieces as inline content and paragraphs:

- [MoveFrom (Move Source Paragraph)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movefrom?view=openxml-3.0.1)
- [MoveToRangeStart (named move destination container)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.movetorangestart?view=openxml-3.0.1)

### Whole tables

WordprocessingML has row/cell insertion and deletion revisions, but no corresponding
table-existence `moveFrom`/`moveTo` marker. The existing `IrMarkupRenderer` therefore
deliberately lowers a non-paragraph `MoveBlock` to a whole-block delete plus insert:
every source row and its content is deleted, and every destination row and its content is
inserted. Accept keeps the destination and removes the source; reject does the inverse.
This is native, conformant Word revision markup, but Word may present it as a deletion and
an insertion rather than one named “Moved” revision.

That lowering is the recommended v1 behavior for a dragged table. Do not invent a named
whole-table move by merely bracketing a `w:tbl`: although move range markers are legal in
several table containers, a range marker alone does not revision-track the existence of
the table or its rows. A different representation should only ship after capturing a
whole-table move made by desktop Word and pinning Word/Open XML validation plus
accept/reject golden tests. Relevant schema semantics:

- [Table (`w:tbl`) content model](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.table?view=openxml-2.20.0)
- [Inserted table row (`w:trPr/w:ins`)](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.inserted?view=openxml-3.0.1)

### Review-aware editor projection

This is the largest addition caused by the confirmed track-changes requirement:

1. Add editor options for the mutation mode and revision author, and expose the existing
   `SetTrackedChanges`/`SetRevisionAuthor` bridge methods through `DocxEditor`'s narrow
   bridge type. Track-changes-off remains the backward-compatible default.
2. In inline-review mode, pass `renderTrackedChanges: true`,
   `showDeletedContent: true`, and `renderMoveOperations: true` consistently through
   first paint, `RenderHtml`, `RenderBlocksHtml`, and fallback rendering. The converter
   already knows how to render move sources/destinations and inserted/deleted table rows.
3. Make `ListBlocks` follow the selected render view. Today it lists raw `w:p`/`w:tbl`
   siblings even when the accepted HTML renderer removes deleted/move-source content.
   Either add a render-view parameter or expose a separate `ListRenderedBlocks` plan so
   the plan and DOM always contain the same units.
4. Treat source and destination as separate visible units in inline-review mode. Reconcile
   can render both; focus and the post-drop flash follow the destination. In accepted view
   only the destination belongs in the render plan.
5. Pin changing modes with pre-existing revisions: switching mutation mode does not
   accept/reject old revisions, so render view and mutation mode should be separate editor
   concepts even if the first UI toggles them together.

## Ordered-unit reconcile

`diffUnits()` should represent exact-token reorders explicitly:

```ts
interface UnitDiff {
  keep: Map<number, number>;
  removed: number[];
  added: number[];
  substituted: Array<{ oldIndex: number; newIndex: number }>;
  moved: Array<{ oldIndex: number; newIndex: number }>;
}
```

Pair identical `unid|contentSignature` entries found in both `removed` and `added` as
`moved` before the existing same-Unid substitution pass. A substitution means “same slot,
fresh HTML”; a move means “same live DOM node, new slot.” Keeping those concepts separate
is necessary for undo/redo as well as drag-and-drop.

`applyBodyDiff()` can then relocate the existing unit wrapper in new-plan order, subject
to the same conservative guards used today:

- move the wrapper returned by `unitWrapperOf()` so table alignment wrappers stay intact;
- require source and insertion neighbors to share a provable parent;
- fall back to a full remount for multi-child border groups or any ambiguous wrapper;
- remount when a moved/source/target unit is a list item, because numbering is
  position-derived and sibling markers may all change;
- keep the churn threshold and universal exception fallback.

The drag UI may optimistically move an unambiguous wrapper on drop, but the session op
must run immediately and reconcile remains authoritative. If the model rejects the move,
put the DOM node back (or remount) and announce the error.

## Interaction design

### Handle ownership

Use **one floating handle** owned by `DocxEditor`, positioned beside the current top-level
unit on pointer hover or focus. Do not inject a control into every converted block.

Why:

- wrapping blocks changes converter layout, margins, contextual spacing, and border
  grouping;
- a child inside `contenteditable` can enter text serialization and offset calculations;
- a child inside a paginated page is clipped by the page box's `overflow:hidden`;
- single-block re-render and reconcile replace nodes, requiring per-node control cleanup;
- one overlay is cheaper on a 346-block document and works for tables without touching
  their internal DOM.

The handle should be at least 24×24 CSS pixels for touch, use a six-dot/grip affordance,
show on hover/focus in continuous mode, remain visible while focused/dragging, and expose
an accessible name containing a short block preview.

### Drag lifecycle

1. Resolve the hovered/focused node through `bodyUnitNodes()`; a descendant cell resolves
   to its containing table.
2. On drag start, commit any dirty active block, capture source anchor/index and focus,
   dim the source to roughly 40%, and create a compact drag preview.
3. Register top-level compatible units as drop targets. Pointer position relative to the
   target midpoint selects `before` or `after`.
4. Draw a 2px insertion line in an editor-owned overlay so page/table overflow cannot
   clip it. Ignore self/adjacent no-op placements.
5. Autoscroll the nearest editor scroll container near its top/bottom edge.
6. On drop, call `MoveBlock`, reconcile/remount as required, flash the moved unit, restore
   focus, call `onMove`/`onEdit`, and announce old/new position.
7. On cancel/error/close/remount, remove transient styles, indicators, registrations, and
   pending animation frames.

The handle must be the only draggable element. Marking the entire paragraph `draggable`
would prevent normal text selection, which is the editor's primary interaction.

### Dependency choice

Recommended: Atlassian's framework-agnostic **Pragmatic drag and drop** core, with only
the optional pieces actually needed (autoscroll is the most useful). It is headless,
incremental, works with native DOM rather than React, and is designed around a separate
drag handle/drop targets. Docxodus would render its own handle, indicator, menu, and live
region so no Atlassian design-system or React dependencies are pulled in.

This will be the npm package's first ordinary runtime dependency. Keep it narrowly scoped
to Pragmatic drag-and-drop core and the optional autoscroll package; do not pull in React
or Atlassian design-system components. Docxodus remains responsible for its own visual
and accessible UI.

Relevant first-party guidance:

- [Pragmatic drag and drop repository](https://github.com/atlassian/pragmatic-drag-and-drop)
- [Atlassian drag-and-drop design guidelines](https://atlassian.design/components/pragmatic-drag-and-drop/design-guidelines/)
- [Atlassian accessibility guidelines](https://atlassian.design/components/pragmatic-drag-and-drop/accessibility-guidelines)
- [Autoscroll package](https://atlassian.design/components/pragmatic-drag-and-drop/optional-packages/auto-scroll)

## Continuous vs paginated view

### Continuous

This is the straightforward path. Body units appear as one ordered flow. Most plain
direct-mode paragraph/table moves can preserve the existing DOM node and finish in the
same latency class as the current structural reconciler. Tracked moves add separately
rendered source/destination units; lists and border groups can remount conservatively.

### Paginated

Blocks live under different page-content parents, page boxes clip overflow, and changing
order can move every following block to a different page. The model operation is still
correct, but the visible projection needs repagination.

For the follow-up paginated hardening:

- place the handle/indicator outside page boxes;
- allow autoscroll across pages;
- call `MoveBlock` using document-order anchors;
- perform the existing full session-attached paginated remount after a successful drop;
- restore focus by moved anchor (preferred) or new document index after remount.

Do not optimistically transplant a block between page-content elements and treat that as
finished; it leaves page heights and every subsequent page stale. Scoped repagination
from the earliest affected page is a later optimization shared with roadmap M4.

## OOXML safety and edge cases

Moving the existing element within one XML part is lossless for the block itself, but
document order carries semantics beyond visible text:

| Case | Recommended v1 behavior |
|---|---|
| Paragraph / heading | Supported |
| List item | Supported with full remount so markers recalculate |
| Whole table | Supported as one unit; direct reorder when tracking is off, tracked delete+insert when on |
| Table cell paragraph / row | No handle; target the whole table |
| Header/footer/note/comment story | Not in v1; body only |
| Inline section-break paragraph (`pPr/sectPr`) | Reject / no handle |
| Trailing body `sectPr` | Not addressable/rendered; never a source |
| Block inside `w:sdt` | Only reorder against a sibling with the same XML parent |
| Floating drawing/image relationships | Safe within the same part; relationship IDs stay valid |
| Footnote/endnote citations | Block move is safe; note ordering/chrome must reconcile |
| Bookmarks/comments/permissions spanning blocks | Needs an explicit policy and tests |
| Existing tracked revisions | Preserve or reject unsafe nesting; test both accepted and inline-review views |
| Tracked paragraph-like move | Named `moveFrom`/`moveTo` pair with one accept/reject group |
| Tracked whole-table move | Native whole-table delete+insert; see the presentation caveat above |

The hardest fidelity edge is a range whose start and end markers are in different blocks.
A literal sibling move can move content into/out of that range or move one endpoint past
the other. The confirmed v1 policy is the safe restriction:

- detect cross-block range membership and reject moves that would change or invert it;
- hide invalid target indicators rather than allowing a doomed drop; and
- return a typed core error if a caller bypasses the UI restriction.

At minimum cover comment ranges, bookmarks, permission ranges, and native move ranges.

## Test plan

### .NET session tests

- move first/middle/last paragraph before and after another paragraph;
- same-source/target and already-adjacent no-ops do not consume undo;
- one undo restores the old order; redo reapplies it;
- direct-mode source and descendant Unids remain unchanged;
- save/reopen retains order and content;
- whole table move preserves cell anchors, relationships, and validation;
- tracked paragraph/heading/list move emits paired named ranges with unique revision IDs,
  configured author/date, and `w:trackRevisions` enabled;
- tracked whole-table move emits source row/content deletions and destination
  row/content insertions with no duplicate anchors;
- accept tracked move yields the requested order exactly once; reject restores the
  original order exactly once, for both paragraph-like and table units;
- selective accept/reject treats a named paragraph move as one group; document-wide
  accept/reject handles the lowered table pair cleanly;
- list item move preserves `numPr` and produces correct projection order;
- reject cross-part, different-parent/content-control, cell paragraph, and section-break
  cases;
- failure restores the exact pre-op document;
- policy tests for cross-block range markers and existing revisions.

### Pure TypeScript reconcile tests

- exact-token reorder is `moved`, not `substituted`;
- forward/backward moves preserve node identity;
- duplicate-content/unid collision cases remain deterministic;
- wrapper ambiguity and list moves request remount;
- mixed move + insertion/deletion produces the requested final order;
- accepted and inline revision-view plans exactly match their rendered DOM units.

### Playwright editor tests

- real mouse drag uses only the handle and persists through save/reopen;
- dragging paragraph text still selects text and never starts block dragging;
- whole-table drag; no cell-level handle;
- track changes off produces no revision markup; track changes on shows a moved-from and
  moved-to paragraph and persists native markup through save/reopen;
- tracked whole-table drag shows deleted/inserted table rows, then Word-compatible
  accept/reject yields the destination/original respectively;
- undo/redo after move, including untouched-node identity in continuous view;
- autoscroll in a long document;
- accessible menu move, focus restoration, and live-region announcement;
- cancel/Escape/outside drop leaves model and DOM unchanged;
- list numbering and footnote chrome after move;
- content-control/section-break invalid targets;
- compact ribbon and multiple editors on one page;
- paginated cross-page move remounts into the correct order;
- Chromium plus targeted Firefox coverage for the interaction lifecycle.

Add move timing to `npm/tests/editor-latency-bench.spec.ts`; a continuous plain-paragraph
drop should target sub-150ms model+reconcile time on the standing HC031 fixture, excluding
the human drag duration.

## Suggested implementation sequence

1. Land direct-mode `DocxSession.MoveBlock` + full transport/client ripple and .NET tests.
2. Add tracked paragraph/table emission by extracting the existing revision helpers;
   pin validation and accept/reject before wiring the editor.
3. Make full/block rendering and `ListBlocks` use one explicit revision view.
4. Teach `diffUnits` / `applyBodyDiff` about moves; pin undo/redo and node identity.
5. Add `DocxEditor.moveBlock()` and accessible non-pointer commands before drag UI.
6. Add the floating handle, drop indicator, pointer drag, and cleanup lifecycle in
   continuous mode.
7. Add list/table/content-control/range safety and browser coverage.
8. Enable paginated-mode full-remount behavior and cross-page autoscroll.
9. Update `editor_ui_surface.md`, README/API docs, CHANGELOG, and screenshots when the
   feature—not this investigation—ships.

## Product decisions

All investigation-stage decisions are resolved. Production v1 is whole-table capable,
continuous-view first, emits native Word revisions when track changes is active, uses
Pragmatic drag-and-drop plus autoscroll, and rejects unsafe cross-block range moves.

## Shipped deviations and follow-ups

Found by smoke-testing the shipped surface against the NVCA Model Certificate of
Incorporation (234 body blocks, 392 bookmarks, 94 footnotes, 3 inline section breaks)
and fixed on this branch.

### Identity-bearing markers in the tracked clone

A tracked move keeps the source and the destination live simultaneously, so every
id-bearing marker the clone copies is a second live copy:

| Marker | Policy | Owner |
|---|---|---|
| `w:bookmarkStart`/`w:bookmarkEnd` | Destination clone gets fresh document-unique ids; **both copies keep the NAME**, so each survives its own resolution and every `REF`/`PAGEREF`/`HYPERLINK \l` still resolves | `DocxSession.RenumberClonedBookmarks`, mirroring `IrMarkupRenderer.NormalizeBookmarks` (B) |
| `w:commentRangeStart`/`End`/`Reference` | The **move source** takes a fresh comment id + a cloned definition (fresh `w14:paraId`, threading entries in both metadata parts, cloned replies re-pointed at cloned parents), leaving the destination on the original comment and its thread | `CommentOps.CloneCommentsForMoveSource`, mirroring `IrMarkupRenderer.NormalizeComments` (B) |
| `w:footnoteReference`/`w:endnoteReference` | **Deliberately duplicated.** A note cited at both the old and the new position is a faithful depiction of a pending move, it is not uniqueness-constrained, and exactly one citation survives either resolution | — |

Without the first two the pending redline is schema-invalid (`Sem_UniqueAttributeValue`)
and the comment shows twice in Word's Reviewing pane. Both resolve to a valid single-copy
document on accept and on reject. One residue is accepted: the resolved-away copy's comment
*definition* stays in `comments.xml` unreferenced — `RevisionProcessor` prunes no orphaned
definition, for any resolved comment, and Word ignores an unanchored one.

### Whole-block revisions mark content, not just structure

`w:moveFrom` is a deletion, so its runs carry `w:delText`/`w:delInstrText` — as Word writes
them and as `RevisionProcessor`'s reject path swaps back. The whole-table lowering marks
**every cell paragraph's content and mark**, not only `w:trPr`; row marks alone left a
moved-away table's text rendering as ordinary body text inside a row Word believes is
deleted. Both go through one owner, `DocxSession.MarkParagraphContentAndMark`.

### Move validity is a query, not a failed attempt

`DocxSession.ValidMoveTargets(sourceAnchorId)` returns the anchors a block may legally move
next to, sharing `MoveBlock`'s own guards (`MoveSourceRejection` + `BlockMoveSafetyError`)
so the UI and the engine cannot disagree. The editor asks once per drag source and once per
menu open, and uses the answer to:

- register only valid blocks as drop targets, so no indicator is drawn over a doomed drop
  (confirmed direction #7);
- disable move-menu items with no destination;
- withhold the handle entirely from a block that can move nowhere (a section-break paragraph);
- resolve **"move to top/bottom" within the source's own region**. On a document with N
  section breaks the body is partitioned into N+1 move regions; targeting the document ends
  meant those two commands could never succeed anywhere but the first and last region.

The engine remains authoritative — the UI gate is advisory, and a caller bypassing it still
gets the typed error.

### Editor projection

`ListBlocks(renderTrackedChanges)` mirrors the rendered DOM in both views. The accepted view
previously disagreed by one unit after a tracked move: `RevisionProcessor` built the
paragraph it coalesces for a deleted/moved-away paragraph mark **without its attributes**, so
the surviving block lost its `pt:Unid` and rendered with no `data-anchor` — unaddressable,
and invisible to the reconciler. Identity now comes from the same member the properties do.

### Still open

- Paginated dragging (handle is not mounted in paginated mode).
- Tracked moves repaint via full remount (~5–6 s on the NVCA charter) where a direct move
  reconciles incrementally (~0.8 s).
- No ribbon control toggles track changes; `trackedChanges`/`revisionAuthor` are mount-time
  options on `DocxEditor`/`mountRibbon`.

# Editor block-drag investigation

Date: 2026-08-07
Branch: `investigate-editor-block-drag`
Baseline: `origin/main` at `fd3a5b02dbe02707f044c55b2a03a4262f05d462` (`v9.4.0`)

## Executive summary

The reported symptoms are real, but they are not one defect. Four mechanisms combine to make
the drag surface look broken:

1. **Firefox hides the handle during a native drag.** Firefox emits `pointercancel` and a chain of
   `pointerleave` events after `dragstart`. The editor clears its pointer-down guard on
   `pointercancel`; the container's `pointerleave` handler then sets the handle to `display:none`
   before Pragmatic Drag and Drop's deferred `onDragStart` marks the editor as dragging.
2. **The insertion line has large hitbox holes.** Only each block element's border box is a drop
   target. Word paragraph margins are outside those elements. In the public demo document, 50 of
   60 adjacent block pairs have a positive uncovered gap (4–26.7 px). The line is shown inside a
   block and immediately hidden in the gap where a user naturally aims to insert between blocks.
3. **The overlay does not observe layout changes.** The handle is repositioned on editor
   `pointermove`, document `scroll`, and window `resize`, but not when fonts, images, edits, or
   surrounding layout move the source block. A controlled reflow moved a block by 142 px while
   its handle stayed at the old coordinate; a synthetic resize snapped it back.
4. **The public web surfaces still run the slower 9.3 implementation.** The GitHub Pages app,
   player, and landing-page embed are pinned to `docxodus@9.3.0`. The substantial block-move
   latency work is in 9.4.0/latest main, so public-surface testing does not exercise the code that
   the current repository benchmark measures.

The OOXML move primitive and the basic pointer drop path are working. A simple three-paragraph
move completed in both Chromium and Firefox, and all five existing block-drag tests passed when
the suite was forced to run in Firefox. Those tests pass despite the Firefox handle being a 0×0,
`display:none` box during the gesture, because they assert only the final order.

## Finding 1 — Firefox drag-start race hides the handle

Severity: **high**

The relevant state transitions are in `DocxEditor.setupBlockDrag()`:

- handle `pointerdown` sets `blockDragPointerDown = true`;
- document `pointercancel` and `pointerup` both clear it;
- container `pointerleave` calls `hideBlockHandle()`;
- `hideBlockHandle()` protects only `blockDragging` and the open menu;
- the PDD `onDragStart` callback later sets `blockDragging = true`.

Low-level Playwright mouse steps produced this Firefox event/state sequence on latest main:

```text
dragstart      dragging=false  pointerDown=true   handle=flex
pointercancel  dragging=false  pointerDown=false  handle=flex
pointerleave   dragging=false  pointerDown=false  handle=flex
pointerleave   dragging=false  pointerDown=false  handle=none
dragenter      dragging=true   pointerDown=false  handle=none
```

During that same gesture:

| Browser | Handle box during drag | Computed display | Drop line | Final reorder |
|---|---:|---|---|---|
| Chromium | 28×30 px | `flex` | visible over target | succeeded |
| Firefox | 0×0 px | `none` | visible over target | succeeded |

This explains why Firefox feels as though drag never started even when native drag events and the
model mutation eventually complete. It also explains the intermittent quality: the race depends
on browser event ordering, not document content.

Recommended correction:

- establish a native-drag guard synchronously on the handle's `dragstart`, before Firefox's
  `pointercancel`/`pointerleave` sequence;
- do not allow `pointercancel` to create a window in which neither the pointer-down guard nor the
  drag-active guard is true;
- explicitly keep the handle visible in the PDD start callback as a defensive invariant;
- add a Firefox assertion that the handle remains connected, displayed, and non-zero throughout
  a low-level mouse gesture.

## Finding 2 — the insertion line exists, but not between many blocks

Severity: **high**

The editor creates a fixed, 3 px blue `.docx-block-drop-indicator`. `showDropIndicator()` positions
it at the top or bottom edge of a target, so the styling itself is not missing. The issue is target
coverage:

- `refreshBlockDropTargets()` registers `dropTargetForElements({ element: unit })` for the block
  element only;
- CSS margins are not part of an element's pointer hitbox;
- leaving a registered element triggers `onDragLeave`, which hides the line;
- releasing without a registered target cancels the move.

Measured against `docs/demo/docxodus-demo-guide.docx` on latest main:

- 61 movable body units;
- all 60 other units are legal targets for every source, so engine safety restrictions are not the
  explanation in this document;
- 50 of 60 adjacent pairs have uncovered vertical space;
- gaps range from 4 px to 26.7 px.

A native gesture over a 10 px demo-document gap produced:

```text
pointer 2 px inside the next block: indicator display = block
pointer centered in its top margin: indicator display = none
elementFromPoint in the margin: DIV (no registered block target)
```

This is the inverse of the expected interaction: users commonly aim at the whitespace between
blocks, which is exactly where feedback and dropping disappear. Documents with illegal move
regions add another no-line case by design. When a drop is cancelled, the only explanation is
written to `.docx-block-live-region`, which is clipped to 1×1 px and therefore invisible to a
sighted pointer user.

The implementation also stops short of two feedback items in the shipped design document: it
does not dim the source block and does not create a compact custom drag preview. Only the handle
receives a dragging class, and on Firefox that handle is hidden by Finding 1.

Recommended correction:

- model insertion positions as continuous vertical hit regions, not only block border boxes;
- either register wrapper/edge hitboxes that own the margins or use one drag monitor that resolves
  the nearest legal before/after boundary from pointer Y;
- keep the last valid indicator stable while crossing the whitespace belonging to that boundary;
- preserve the engine's per-side legality rules when resolving the nearest boundary;
- give pointer users visible invalid/cancelled feedback rather than relying only on the live region;
- dim the actual source unit and provide a browser-stable custom preview.

## Finding 3 — handle coordinates go stale after reflow

Severity: **medium**

The handle is a fixed overlay positioned from `unit.getBoundingClientRect()`. It is refreshed by:

- editor `pointermove`;
- focus entering a block;
- captured document `scroll`;
- window `resize`.

There is no `ResizeObserver`, mutation/layout observer, or frame-coalesced position controller for
the active source. A diagnostic changed the height of an earlier paragraph without moving the
pointer or viewport:

```text
before reflow:       handle top 205.875, source top 205.875, gap   0 px
after reflow:        handle top 205.875, source top 347.875, gap -142 px
after resize event:  handle top 347.875, source top 347.875, gap   0 px
```

Real equivalents include web-font settlement, image sizing, nearby edits, host layout changes,
and any asynchronous chrome that changes the document flow. During a drag the editor deliberately
ignores editor `pointermove`, making scroll events the only continuous correction path.

Recommended correction:

- observe the active unit (and, where needed, the editor/scroller) with `ResizeObserver`;
- coalesce pointer, scroll, resize, and observed-layout updates through one
  `requestAnimationFrame` position writer;
- prefer transform-based overlay movement to repeated `top`/`left` layout writes;
- position both handle and indicator through the same controller so they cannot represent
  different layout frames;
- test reflow and autoscroll alignment in the full ribbon scroller.

## Finding 4 — performance differs sharply by deployed version

Severity: **high for the public demo, low-to-medium on latest main**

`docs/demo/app.html`, `docs/demo/player.html`, and the landing-page embed default to
`docxodus@9.3.0`. Version 9.3 performs a synchronous `ValidMoveTargets` bridge query and
re-registers one drop listener per body block whenever hover crosses to a new block.

The 9.4 performance change (`006dd66`) moved the legality query off the immediate hover path,
memoized it, made the engine sweep linear, changed DOM block resolution from a document scan to
an ancestor climb, and deferred drop-target registration until drag start. The historical
bookmark-dense charter measurements recorded by that change were:

| Operation | 9.3 | optimized implementation |
|---|---:|---:|
| Hover across block boundary | 624 ms | 8 ms |
| Drag-start legality query | 4.3 s | 20–30 ms |
| Move-menu open | 4.4 s | 12 ms |
| Direct move | 780 ms | 150 ms |

The latest-main benchmark was rerun during this investigation on the same 234-unit NVCA charter:

| Operation | Measured latest main |
|---|---:|
| Hover boundary (12-event average) | 18.5 ms |
| `ValidMoveTargets` | 20.8 ms |
| Drop-target registration (234 units) | 2.0 ms |
| Direct move | 144.2 ms |
| Undo after move | 107.0 ms |
| Tracked move | 391.4 ms |

A focused warm-path probe measured `showBlockHandle`/dispatched pointer work near 0.07 ms but
captured a 21 ms forced-layout spike on one geometry read. This suggests that the 18.5 ms aggregate
is variable browser/layout work rather than the old multi-hundred-millisecond engine query. The
main-branch latency is vastly improved, but a direct move still visibly occupies roughly 100–150
ms and overlay geometry can miss an occasional frame.

Recommended correction:

- update all public web-surface pins together after release validation so the demo actually runs
  9.4.0 or the intended newer release;
- retain exact pins, but add a single source of truth or release check so app/player/landing-page
  versions cannot silently lag;
- keep the current charter benchmark and add frame/long-task observations around real pointer
  gestures and overlay layout, not only synchronous method timing.

## Coverage gaps that allowed this to ship

1. `editor-block-drag.spec.ts` runs only in the Chromium project. The only Firefox project is
   restricted to `editor-multiblock-format.spec.ts`.
2. The positive drag test uses Playwright `locator.dragTo()` and checks only final document order.
   It never samples the handle or indicator during the gesture.
3. The invalid-target test observes that the indicator never shows, but no positive test requires
   it to show over a valid target.
4. Tests use tightly stacked synthetic paragraphs; they do not drag across Word-derived margins.
5. No test changes surrounding layout after the handle appears.
6. The block-drag tests mount the minimal `DocxEditor`, not the actual ribbon scroller used by all
   public web surfaces.

Minimum regression matrix for a fix:

- Chromium and Firefox low-level mouse gesture, with `dragstart`, `dragover`, `drop`, and `dragend`;
- handle visible and non-zero from pointerdown until drop;
- source block visibly marked as moving;
- indicator visible inside targets and throughout inter-block margins;
- before/after changes as the pointer crosses a target midpoint;
- invalid per-side destinations never become droppable, with visible explanation;
- handle/source alignment after reflow, host resize, manual scroll, and drag autoscroll;
- final model order, save/reopen order, undo, and tracked-move behavior unchanged;
- the same positive smoke test through `mountRibbon`, not only the minimal harness.

## Verification performed

- Fetched and verified `origin/main`; local main, remote main, and `FETCH_HEAD` all resolved to
  `fd3a5b02dbe02707f044c55b2a03a4262f05d462` before branching.
- Existing Chromium block-drag suite: **5 passed**.
- Existing block-drag suite forced through bundled Firefox: **5 passed**.
- Low-level Chromium gesture: native lifecycle observed, handle 28×30, indicator visible, reorder
  succeeded.
- Low-level Firefox gesture: native lifecycle observed, handle `display:none`/0×0, indicator
  visible inside target, reorder succeeded.
- Public demo document target sweep: 61 units, every source had all 60 destinations available.
- Public demo margin sweep: 50 positive gaps, maximum 26.7 px; line cleared in a measured 10 px gap.
- Reflow geometry probe: reproduced 142 px handle/source separation and resize recovery.
- Latest-main NVCA latency benchmark: passed all standing budgets with no reconcile fallback.

No product fix is included in this investigation branch; the changes above are recommendations
derived from reproduced current-state behavior.

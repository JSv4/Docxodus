# The Editor Surface

What the in-browser editor exposes, what each control actually does, and what it costs.

This is the **UI reference**. For *why* the editor is built the way it is (Option B: the live
`DocxSession` is the model of record and the IR/anchor system is the addressing overlay), read
[`ir_editor_feasibility.md`](ir_editor_feasibility.md). For the underlying edit contract — anchor
lifecycle, error catalog, supported markdown subset — read
[`docx_mutation_api.md`](docx_mutation_api.md).

Three layers are described here, and they are not the same thing:

| Layer | Path | What it is |
|-------|------|------------|
| `DocxEditor` | `npm/src/editor.ts` | The framework-agnostic editor **engine**: the live `DocxSession`, block wiring, every command. Deliberately no chrome. A consumer who wants their own UI builds on this. |
| `mountRibbon` | `npm/src/ribbon.ts` (+ `ribbon-chrome.ts`) | The **surface**: tabbed ribbon, anchor rail, table picker, loading overlay, responsive layout. Shipped in the npm package; it is what the screenshots below show. |
| `createRibbonEditor` | `npm/src/embed.ts` | The one-call CDN form of the surface — boots WASM, scopes the document CSS, mounts the ribbon, narrates the wait. |

`mountRibbon` used to be three hand-written copies (`npm/examples/editor.html`, the GitHub Pages
landing page, the compact player), which drifted until the demo advertised a smaller editor than
the one that shipped. It now has one owner, and the hosts differ only in how they obtain the WASM
exports and how much chrome they turn on:

| Host | Obtains exports | Chrome |
|------|-----------------|--------|
| `npm/examples/editor.html` | boots `./_framework/dotnet.js` itself | full, bands on |
| `docs/demo/app.html` | `createRibbonEditor` from jsDelivr | `auto`, full-bleed |
| `docs/demo/index.html` | `createRibbonEditor` from jsDelivr | `auto`, inside the landing page's frame |
| `docs/demo/player.html` | `createRibbonEditor` from jsDelivr, on tap | pinned `compact`, no hint |

Every screenshot is the [NVCA Model Certificate of Incorporation](https://nvca.org/model-legal-documents/)
(346 blocks, 94 footnote citations, 4 sections, 48 rendered pages) opened unmodified.

---

## 1. Anatomy

![The editor with a document open](../images/editor/editor-overview.png)

Four regions, top to bottom:

1. **Document strip** — `New` / `Open` / `Save`, then `Undo` / `Redo`. Never behind a tab; these are
   used constantly and hiding them costs more than the space they take.
2. **Tab strip** — `Home`, `Insert`, `Layout`, and a contextual `Table` tab that exists only while the
   caret is inside a table.
3. **Anchor rail** — live engine state (§6).
4. **Document** — the rendered DOCX. Blocks are individually `contenteditable`; the page sheet is the
   body flow, with header/footer bands docked around it (§4).

A block shows a focus outline when it is the active edit target. In the shot above the caret is in a
body paragraph with a sub-range selected — note that the ribbon's `I` is lit (the paragraph is
italic) and the size box reads `11`, both derived from the selection rather than from editor state.

---

## 2. Home

![Home tab](../images/editor/ribbon-home.png)

| Control | `DocxEditor` command | Session op |
|---------|---------------------|-----------|
| **B** / *I* / <u>U</u> / ~~S~~ / `</>` / x² / x₂ | `format(key)` | `ApplyFormat` |
| Size box | `setFontSize(pts)` | `ApplyFormat` (`FontSizePts` → `w:sz`/`w:szCs`) |
| Font dropdown | `setFontFamily(name)` | `ApplyFormat` (`FontFamily` → `w:rFonts`) |
| Align left / center / right / justify | `setAlignment(a)` | `SetParagraphFormat` |
| Decrease / increase indent | `indent(±720)` | `SetParagraphFormat`, **or `SetListLevel` when the block is a list item** — outdent/indent on a list means "change level", which is what the user means and what Word does |
| • List / 1. List | `toggleList(kind)` | `GetListMembership` to detect the current kind, then `ApplyListFormat` with `none` to toggle off |
| Page break | `pageBreakBefore(true)` | `SetParagraphFormat` |
| Style dropdown | `setParagraphStyle(id)` | `SetParagraphStyle` |
| Delete block | `deleteBlock()` | `DeleteBlock` — inert inside a table and when it is the only editable block |

The inline-format buttons apply to a **selected sub-range**, not the whole block, and every
paragraph-level command applies across a **multi-block selection**, reconciled as N single-block
swaps with the cross-block selection restored.

Each document block remains an independent `contenteditable` host so edits can be committed through
its stable OOXML anchor without replacing the surrounding document. Browsers normally fence a
physical mouse drag at that host boundary, so `DocxEditor` takes over only after the pointer enters
another editable block in the same OOXML story and exposes the gesture as one normalized DOM
`Range`. Cross-block updates are coalesced into the next animation frame, after the browser's own
mousemove selection update but before paint; this keeps Firefox's highlight live instead of
snapping to the complete range only on mouseup. Intra-block selection remains native, and
body/header/footer story boundaries remain uncrossable.

Size and font controls cache the last real selection, because a combobox steals focus when clicked.
Single-block selections are bookmarked as one anchor plus a span; multi-block selections use two
stable anchors plus content offsets. Without those bookmarks a sub-range selection would be lost
before the command ran.

---

## 3. Insert

![Insert tab](../images/editor/ribbon-insert.png)

| Control | `DocxEditor` command | Session op |
|---------|---------------------|-----------|
| Table | `insertTable(rows, cols, opts)` | `InsertTable` |
| Single / Thick / Double rule | `insertHorizontalRule(weight, style, position)` | `InsertHorizontalRule` (an empty bottom-bordered paragraph) |
| Below block / Above block | *(modifies the rule position argument)* | — |
| Clear | `clearParagraphBorders()` | `SetParagraphFormat` (`clearBorders`) |
| Footnote / Endnote | `insertFootnote(md?)` / `insertEndnote(md?)` | `InsertFootnote` / `InsertEndnote` |

**Table** opens a size picker rather than a text prompt:

![Table size picker](../images/editor/table-picker.png)

Hover picks the dimensions; the footer carries cell alignment and a borderless toggle (borderless is
the default because a layout table is the common case in legal documents).

**Footnote / Endnote** cite a new note at the caret. Body blocks only — Word does not allow a note
reference inside a header/footer story or inside another note, so those are rejected client-side
rather than round-tripping to an `AnchorWrongKind`. The caret offset is captured *before* the block
is synced (syncing re-renders the block and would drop the live selection), so the citation lands
mid-word if that is where the caret was.

---

## 4. Layout

![Layout tab](../images/editor/ribbon-layout.png)

| Control | `DocxEditor` command | Session op |
|---------|---------------------|-----------|
| Page view | `setPaginated(bool)` | *(render mode; re-renders the live session, edits survive)* |
| Header & footer bands | *(re-opens with `headerFooter`)* | — |
| Format / Start at | `setPageNumbering({format, start})` | `SetPageNumbering` → `w:pgNumType` |
| Clear | `clearPageNumbering()` | `ClearPageNumbering` |
| Page number / Total pages | `insertPageNumber(field)` | `InsertPageNumberField` |

Page numbering is **per section**, resolved from the section that owns the caret's block. The same
setting is surfaced on both header/footer bands; all three read the live session, so changing it in
one place updates the others.

The field buttons live here rather than under Insert because the field lands in the **footer story**,
while every Insert control acts at the caret. Grouping them with the section's numbering keeps that
readable.

---

## 5. Table (contextual)

![Contextual Table tab](../images/editor/ribbon-table-contextual.png)

Appears only while the caret is inside a table; selecting away from the table hides it and falls back
to Home.

| Control | `DocxEditor` command | Session op |
|---------|---------------------|-----------|
| Insert above / below | `insertTableRow("above"\|"below")` | `InsertTableRow` |
| Insert left / right | `insertTableColumn("left"\|"right")` | `InsertTableColumn` |
| Delete row / column | `deleteTableRow()` / `deleteTableColumn()` | `DeleteTableRow` / `DeleteTableColumn` |

This replaced a **floating** table toolbar, whose absolute positioning had to be corrected twice — it
covered the first row, and then it covered the content below the table. A docked tab cannot overlap
the cell being edited, so the whole class of bug is gone by construction. Deleting the last row or
column removes the table. v1 assumes a rectangular grid (no `w:gridSpan`).

---

## 6. The anchor rail

```
ANCHOR  p:body:09612b1c13…   BLOCKS 346   SESSION #1   LAST OP  page format 651 ms
```

| Cell | Meaning |
|------|---------|
| `anchor` | The focused block's `kind:scope:unid`. `kind` ∈ `p`/`h`/`li`; `scope` ∈ `body`/`hdr`/`ftr`/`fn`/`en` |
| `blocks` | Addressable blocks currently rendered |
| `session` | The live `DocxSession` handle in WASM |
| `last op` | The last command and how long it actually took |

The rail is not decoration. Anchor addressing *is* the architecture — every edit is routed by anchor,
not by DOM position — so the demo states the anchor rather than hiding it in devtools. It is also the
fastest way to confirm scope resolution is right: clicking a footnote body shows `p:fn:…`, a header
line shows `p:hdr:…`.

`last op` exists because operation cost here is not uniform (§9), and a surface that hides its repaint
cost invites you to design as if it were free — it is how the old six-second structural remount was
caught and driven down to a few hundred milliseconds.

---

## 7. Header/footer bands

![Header band showing the first-page story](../images/editor/header-band.png)

Opt-in (`headerFooter: true`). Header/footer stories live in their own OOXML parts outside the body,
so they dock as their own regions rather than joining the body flow.

Each band carries a `Default` / `First page` / `Even pages` kind selector, a page-number control, and
the section's page-number format/start. Band paragraphs are wired by the same code path as body
blocks, so **the entire ribbon works inside a band** with no band-specific command code.

The shot above shows the document's real first-page header after switching the kind to `First page`,
along with the inline warning the band raises: `w:titlePg` makes page 1 use its own first-page header
*and footer*, so an empty first-page footer silently leaves page 1 with no footer at all. A kind whose
story is inherited from an earlier section is shown, marked inherited, rather than offering to create
a redundant part.

Bands compose per story paragraph via the session-attached `RenderBlockHtml`, not from the body
render: the full render stamps anchors only in the main document part, and paginated mode clones one
header node onto every page — so a page-margin overlay could never be uniquely addressable.

---

## 8. Notes and pagination

Footnotes and endnotes render inline and are **ordinary editable blocks** — not opt-in, because they
are document content:

![Footnotes section](../images/editor/footnotes.png)

The citation marker and the `↩` backref are converter-generated chrome: they are excluded from the
content-offset space and are not editable. Without that, offsets drift, or the rendered display
number gets committed as literal text and destroys the citation run.

`Page view` flows blocks into real page boxes, with notes at the foot of the page that cites them and
per-page number substitution:

![Paginated view](../images/editor/paginated.png)

Page numbers here are *computed per page*, not the field's cached result — a header is authored once
and cloned onto every page, so the cached value would read the same number throughout. The footer
above shows `Last Updated October 2025 i` on page 1 of a section formatted `lowerRoman`.

---

## 8a. Page geometry: the column is the document's, the zoom is the view's

Both views lay the document out at ONE width — the text column its `w:sectPr` defines (page size
minus margins), which the converter stamps on every section wrapper (`data-content-width`,
`data-page-width`, the four margins) in **every** render mode. `DocumentViewport`
(`npm/src/viewport.ts`) applies it: each section wrapper gets the authored column plus the
authored margins as gutters, so `.docx-body-flow` measures exactly one page wide and *is* the
sheet the chrome paints.

A window narrower than that page is handled the way a word processor handles it — by zooming the
page to fit, never by reflowing the column:

| | continuous | paginated |
|---|---|---|
| Column width | section's `contentWidth` (pt) | page boxes, already page-sized by `pagination.ts` |
| Fit | zoom on `.docx-body-flow` | zoom on `#pagination-container` |
| Recomputed | `ResizeObserver` on the editor container | same |

`DocxEditor.zoom` reports the applied scale, and the viewport publishes the page's on-screen
width as `--docx-sheet-width` so chrome docked *outside* the zoomed sheet — the header/footer
bands — lines up with it at any zoom. `fitToWidth: false` opts out (an oversized page then
overflows and the surface scrolls); `columnWidth: "fluid"` restores the pre-geometry behavior of
sizing the column from the host.

Why it matters beyond aesthetics: with a device-sized column, line breaking differed on every
screen and matched neither Word nor the editor's own page view, and content the document sizes in
absolute units — a full-width table above all — could not fit at all. A table box never shrinks
below its content's minimum, so it overflowed the sheet and was clipped by the window; enlarging a
run's font size made that minimum grow, so the damage peaked exactly while the user was editing.
The converter's side of the same fix (Word's table layout and word-breaking rules, which CSS's
defaults invert) is in `docs/ooxml_corner_cases.md`.

---

## 9. What operations cost

Measured on a real document (`HC031-Complicated-Document.docx`, Chromium, WASM, warm) by
`npm/tests/editor-latency-bench.spec.ts` — the standing latency instrument; run it before and
after touching any hot path. Values are single-sample and machine-dependent; treat them as
ratios, not contracts:

| Operation | Cost | Before the 2026-08 latency pass | Why |
|-----------|------|--------------------------------|-----|
| Open + first render | ~3 s | ~3 s | Full document conversion (one-time; M3 worker offload is the open item) |
| Text edit (commit on blur) | ~30 ms | ~100 ms | Session op + single-block re-render through the persistent shell |
| Inline format (bold, size) | ~35–50 ms | ~90–110 ms | Same, plus selection restore |
| Paragraph format (align) | ~40 ms | ~85 ms | Same |
| Enter (split) | ~40 ms | ~135 ms | Both halves render in ONE batched `RenderBlocksHtml` call |
| Backspace (merge) | ~25 ms | ~80 ms | Session op + one block render |
| Insert table / row | ~85–130 ms | ~1.2–1.5 s (remount on real docs) | Incremental reconcile |
| Delete block | ~25 ms | ~1.2 s (remount on real docs) | Incremental reconcile |
| Undo / redo | ~55–125 ms | ~1.1–1.2 s (remount on real docs) | Snapshot restore + incremental reconcile |
| `save()` | ~60–80 ms | ~60 ms | Lossless serialize |

The "before" column's structural-op numbers deserve a note: the incremental reconcile existed,
but on any document containing a block-level `w:sdt` (a TOC — i.e. most real documents) the
render plan missed the sdt's content blocks, the plan/DOM diff read as 100 % churn, and every
structural op silently fell back to the multi-second full remount. `ListBlocks` now flattens
`w:sdtContent` exactly as the renderer does, so the diff actually engages.

### The move surface

Block drag/reorder has its own instrument, `npm/tests/editor-move-latency-bench.spec.ts`, run
against `TestFiles/NVCA-Model-COI.docx` — 234 body blocks, **392 bookmarks**, 94 footnotes, 3
section breaks. The fixture is chosen, not incidental: the move guards are driven by cross-block
ranges and section breaks, and HC031's six bookmarks cannot exercise them at all.

| Operation | Cost | Before the 2026-08 move pass | Why |
|-----------|------|------------------------------|-----|
| Hover (pointer crosses a block) | ~8 ms | ~620 ms | Hover no longer asks the engine anything |
| Drag start / `ValidMoveTargets` | ~20–30 ms | ~4.3 s | One precomputed context per sweep, not per candidate |
| Move menu open | ~12 ms | ~4.4 s | Same query, plus memoization |
| Move a block (tracking off) | ~150 ms | ~780 ms | Incremental reconcile; undo restores the XML cache, not the package |
| Undo a move | ~165 ms | ~410 ms | Same |
| Move a block (review mode) | ~390–640 ms | ~8.4 s | Reconciles instead of remounting |

What made the move surface slow was not the move. `ValidMoveTargets` rebuilt the whole block
sequence and re-scanned the story for every marker pair, **per candidate and per side** — 2N
passes over the document, each materializing a member set per cross-block range — and the drag
UI called it (plus a re-registration of one drop listener per block) every time the pointer
crossed a paragraph boundary. On this charter that was seconds of work to move the mouse.

The guard now precomputes what is a property of the CONTAINER — block order, which blocks own a
section break (as a prefix sum), and each cross-block range as an index pair — once per sweep,
and answers each candidate with index arithmetic. `DocxSessionMoveBlockTests` pins the rewrite
against the original set-membership definition for every (source, target, side) triple, because
"faster" is only interesting if it decides identically. On the UI side, hovering consumes a
memoized answer and schedules the ask for idle time (so the handle withdraws from an immovable
block a beat later rather than the pointer paying for the query).

#### Drag feedback, and why drop position is geometry

The handle floats in the page MARGIN, 32 px left of the block's text column. So the natural
gesture — press it and pull straight down — keeps the pointer in the gutter, where it never
crosses a paragraph box. Resolving the drop target by asking which element the pointer is over
therefore found nothing for the entire gesture: no drop line, and a release that silently did
nothing. Drops only worked if the user happened to steer into the text.

There is now **one** drop target, on the document flow, and `DocxEditor.resolveDropAt` decides
where a release lands from the pointer's vertical position:

- every movable block is measured once at drag start (`captureDropZones`, ~1.4 ms on the
  charter; nothing reflows mid-drag, and scrolling — including drag autoscroll — moves the flow
  rigidly, which `dropZoneShift` corrects with one subtraction);
- the nearest block by vertical distance is the target, the half the pointer is in picks the
  side, snapped to the other side when only that one is legal;
- when neither side is legal — the pointer is in a region this block cannot reach, e.g. across a
  section break — there is no drop and nothing is drawn. Snapping to a distant legal boundary
  would move the block somewhere the user never pointed at.

Resolution costs **0.0 ms** per `dragover` on the charter, and the flow no longer carries one
registered drop listener per block (234 of them, re-registered at every drag start).

Three signals, each answering a different question:

| Signal | Answers | Mechanism |
|--------|---------|-----------|
| Source block dimmed to 38 % | *what* is moving | `.docx-block-drag-source`, added after the preview snapshot so the chip is not dimmed too |
| Preview chip naming the block | *what* is under the cursor | `setCustomNativeDragPreview`; the browser would otherwise ghost the 26 px grip |
| 2 px accent line with a leading dot | *where* it will land | `.docx-block-drop-indicator`, positioned by `transform` (no layout) with a one-shot 110 ms fade on each appearance |

`editor-block-drag.spec.ts` drags down the gutter without ever entering the text column and
asserts all three, sampled per pointer step rather than only at the end.

A review-mode move used to force a full remount, on the reasoning that the move-from/move-to
pair had to stay canonical. It does not: the render plan signs every unit with a content hash,
so the source — which keeps its unid while gaining the `w:moveFrom` wrapper — diffs as an
ordinary in-place substitution. Both bench and `editor-block-drag.spec.ts` assert
`lastReconcileFallback === null` after a tracked move, since a silent return to remounting would
show up only as a number.

Open on this document is ~10–12 s and is dominated by the full HTML conversion (~9 s of it),
which no part of this pass touches — see the M3 worker offload item.

Structural operations reconcile: `DocxEditor.reconcile()` diffs the DOM's top-level unit
sequence against the session's render plan (`ListBlocks` — LCS over `unid|contentHash` tokens),
keeps unchanged units' DOM nodes, renders changed/created units in one batched WASM call
(`RenderBlocksHtml`, with real sibling context and true list-marker numbers), and renumbers
footnote/endnote marker chrome positionally from `ListNotes`. Substituted units pair by unid
first, positionally as fallback. A full remount survives as the universal **fallback** —
paginated mode, pure list-item insert/remove (sibling numbers shift without sibling XML
changing), border-`div` regrouping (`insertHorizontalRule`, `clearBorders`, list toggles), or
any inconsistency — so correctness never depends on the diff; the reconciled DOM is pinned
equal to a remounted DOM by `npm/tests/editor-reconcile.spec.ts`. When an op reads slow in the
rail, `editor['lastReconcileFallback']` says why it fell back.

Single-block re-renders go through a **persistent render shell**: the session keeps an open
throwaway document holding the formatting parts, and each render replaces only its body
(`HtmlConversionOps.RenderTargetsFromShell`), so the package open, styles/numbering parse, and
the converter's style-resolution caches are paid once per formatting-signature change instead
of per keystroke.

---

## 10. Driving the surface from tests

- Blocks are addressable in the DOM as `#editor [data-anchor]`; editable ones carry
  `contenteditable="true"`.
- `window.__demo` exposes `{ ribbon, exports, openDoc(bytes, name), getEditor() }` — and only once
  the engine is usable, so its appearance is the "surface is live" signal.
- `window.__selectTab(name)` activates a ribbon tab without pointer geometry; `window.__ribbon` is
  the `RibbonEditor` itself (`selectTab`, `setChrome`, `loader`, `open`, `destroy`).

**A control on a non-active tab is `display:none` and therefore not clickable.** A spec that touches
one must activate its tab first — `npm/tests/editor-demo-grid.spec.ts` calls `__selectTab('insert')`
before clicking `#table`.

Every control carries `data-dxr="<name>"`, and the surface *also* stamps `id = idPrefix + name`.
With no explicit `idPrefix` it uses bare ids when they are free on the page and generates
`dxr<N>-` the moment any of them is taken — so the historical spec selectors keep working and a
second ribbon on the same page cannot collide with the first. Stable ids the specs bind to:
`#editor`, `#fontsize`, `#new`, `#save`, `#undo`, `#redo`, `#table`, `#gridpicker`, `#gridcells`,
`#gridalign`, `#paginated`, `#loader`. Behavioural attributes are the delegation contract:
`data-cmd`, `data-align`, `data-indent`, `data-list`, `data-tt`.

The root reports its own state: `.dxr[data-state]` is `idle` | `loading` | `ready` | `error`, and
`.dxr[data-chrome]` is `full` | `compact`.

Serving the demo: `npm run build`, then copy `examples/editor.html` + the bundles into `dist/wasm/`
(this is what `pretest` does) and serve that directory. After a WASM rebuild, serve on a **new port** —
a warm browser blocks the new payload with an SRI integrity error that looks like a build failure.

---

## 11. Responsive behaviour

![The compact layout on a phone](../images/editor/ribbon-compact.png)

Density is measured from the **root element**, not the viewport, via a `ResizeObserver` — a narrow
embed inside a wide desktop page is narrow, and a viewport media query gets that wrong. Below
`compactBreakpoint` (default 720 px) the surface switches to `compact`:

| | `full` | `compact` |
|---|---|---|
| Ribbon panel | multi-row groups with labels | one horizontally-scrolling strip, labels dropped |
| Anchor rail | shown | hidden |
| Editing hint | shown | hidden |
| Title bar | brand + doc name + status | brand + doc name, status dropped |
| Table picker | popover under its button | docked to the bottom edge, within thumb reach |
| Sheet padding | 56 px vertical | 22 px vertical (gutters come from `w:sectPr`, not from chrome) |
| Touch targets | 30 px | 40 px on coarse pointers |

**Compact trims the chrome, never the page.** The document's text column is the width its
`w:sectPr` defines at every density; a window too narrow for it is handled by the viewport's
fit-to-width zoom (§8a), not by reflowing the column. A phone therefore shows a whole smaller
page rather than a narrower one whose lines break where Word's never would.

**No command is removed in compact** — the strip scrolls instead. That is the one rule the previous
hand-rolled mini toolbar broke, and the reason the compact player kept falling behind the editor.

`chrome: "compact"` or `"full"` pins the density (what `player.html` does, since a host site may
size its iframe wide enough to trip the roomy layout); `"auto"` is the default.

Options the host controls: `chrome`, `compactBreakpoint`, `rail`, `fileActions`, `hint`, `loader`,
`documentName`, `idPrefix`, `onSave`, `onOpen`, `onStatus`, `onCommand`, plus every
`DocxEditorOptions` field.

---

## 12. The loading overlay

Opening a document in the browser means streaming a trimmed .NET runtime. The wait is real, so the
surface spends it explaining what is being built rather than showing a spinner: a staged narrative
(engine → document → wiring → ready) with a progress bar, and a rotating capability card.

The overlay paints **before** any runtime exists, which is what lets a host that boots its own
runtime narrate the gap: `mountRibbon` returns immediately with `ribbon.loader`, and the host calls
`stage(n)` / `progress(pct, label)` / `done()` / `fail(err)` as it goes. `createRibbonEditor` drives
those four stages itself. `fail()` shows the error with a Retry button instead of leaving a dead
surface; `loader: false` removes the overlay entirely and every method becomes a no-op.

It covers the whole instrument, so half-built chrome never flashes, and drops `pointer-events` the
instant the fade starts — the surface underneath is already live by then.

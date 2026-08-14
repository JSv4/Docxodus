# IR-Powered DOCX Editor — Roadmap

Companion to `ir_editor_feasibility.md` (which records the verdict, architecture, and
PoC results). This is the **sequenced, prioritized** plan for turning the proven
foundation + MVP into a complete editor. Supersedes the scattered "Still Plan 2" notes.

Status (branch `feat/ir-editor-feasibility-poc`, PR #234): **foundation + MVP shipped and
proven; M1 (rich in-block editing), M2 (structural editing), M5 + M5b (formatting controls:
bold/italic/underline/strike/code, super/sub, alignment, indent, page break, paragraph style,
undo/redo) done; the full ribbon now SHIPS as `mountRibbon` (`npm/src/ribbon.ts`) /
`createRibbonEditor` (`docxodus/embed`), hosted by `npm/examples/editor.html` (`npm run
demo`) and by all three GitHub Pages demo pages; Mlists (bullets/numbered) done — all 7
requested controls shipped.** M3 (worker
offload) / M4 (re-paginate-on-edit) are next.

## Architecture invariants (do not break)

1. **Model-of-record = the live OOXML in `DocxSession`** (lossless `Save`). The IR is
   read-only and has no IR→OOXML writer; never make the IR or the DOM the source of truth.
2. **Addressing = the shared `{#kind:scope:unid}` anchor system.** `convertDocxToHtml`
   (stampAnchors) ↔ `DocxSession` ↔ `RenderBlock` all use one Unid scheme; keep it that way.
3. **Render is a projection; patch incrementally.** An edit goes through `DocxSession` by
   anchor, then only the changed block re-renders (`RenderBlockHtml`). Never round-trip the
   whole doc through `convertDocxToHtml` per edit.
4. **Untouched content stays byte-faithful on save.** Edits may simplify the *edited* block
   (within the markdown subset), but must never degrade blocks the user didn't touch.

## Shipped (foundation + MVP)

- C#: `WmlToHtmlConverterSettings.StampAnchors`; `HtmlConversionOps.RenderBlockHtml`
  (stateless + session-attached); `DocxSession.LiveDocument`.
- WASM/npm: `RenderBlockHtml`, `stampAnchors`, `renderBlockHtml()`, `DocxSession.renderBlock()`.
- `DocxEditor` (pure TS): faithful render → editable paragraphs/headings → commit via
  `DocxSession` → incremental re-render → lossless save; `{ paginated: true }` page boxes.
- Tests: C# `HCO050`/`HCO052`; browser `render-block.spec.ts`, `editor.spec.ts`.

## Milestones (priority order = impact)

### M1 — Rich in-block editing (preserve inline formatting)  · effort M–L · ✅ **DONE**
**Problem:** `commitBlock` replaced an edited block from `el.textContent` (plain text), so
editing a formatted paragraph destroyed its bold/italic/links — the biggest correctness trap.
**Shipped:** `serializeInlineMarkdown(block)` (exported from `npm/src/editor.ts`) walks the
edited block's DOM and emits the projector's markdown subset, detecting emphasis via
`getComputedStyle` (font-weight/font-style) and links via `href`, merging adjacent
same-format runs; `commitBlock` sends that markdown to `ReplaceText` instead of plain text.
Test `editor.spec.ts` "M1: editing preserves inline formatting" edits a block with
bold/italic/link and confirms `**…**` / `*…*` / `[…](…)` survive save+reopen. Formatting the
markdown subset cannot express (size/color) is still dropped on an *edited* block; a future
pass can use finer-grained `ReplaceTextAtSpan`/`ApplyFormat`. **Applying** new formatting
(toolbar) is M5.

### M2 — Structural editing via keyboard  · effort M · ✅ **DONE**
**Problem:** no way to add/split/merge blocks from the UI; ops existed in `DocxSession` but
weren't wired.
**Shipped:** Enter at the caret → `SplitParagraph(anchor, offset)` (offset from a Selection
range); Backspace at block start → `MergeParagraphs(prev, this)`. A `keydown` handler on each
block intercepts both, flushes any uncommitted typing first (`syncBlock`), applies the op,
and reconciles the DOM from `EditResult.modified/created/removed` (re-render the affected
block(s), insert/remove nodes, update the `unid → fullId` map, place the caret). Test
`editor.spec.ts` "M2: split and merge" splits a block (+1), merges the halves back (−1, text
restored exactly), and round-trips through save. Insert-at-doc-start and block delete/reorder
remain follow-ups (Enter-split + Backspace-merge cover the core authoring loop).

### M2.8 — Block drag handles / reorder  · effort L · ○ **DESIGNED**
**Problem:** blocks can be split, merged, inserted, and deleted, but the session has no atomic
move primitive and the editor has no direct manipulation affordance for reordering them.
**Plan:** add one undoable `DocxSession.MoveBlock(source, target, before|after)` operation,
ripple it through the session transports, teach the ordered-unit reconciler to relocate an
existing node, and add a floating (non-contenteditable) drag handle plus accessible move menu.
Confirmed product scope is whole-table support, continuous view first, and native Word revisions
when track changes is active. Paragraph-like blocks use named `moveFrom`/`moveTo` pairs; whole
tables use Word-native row/content deletion plus insertion unless a desktop-Word fixture proves a
conformant named whole-table move representation. This also requires revision-view-aware rendering
and render plans. Pragmatic drag-and-drop core plus autoscroll is approved. Content controls and
section breaks remain immovable; moves that would disrupt cross-block ranges are rejected. Full
investigation and test matrix:
[`editor_block_drag_handles.md`](editor_block_drag_handles.md).

### M2.7 — Latency pass: persistent render shell + reconcile on real documents  · effort M · ✅ **DONE**
**Problem:** interactive ops were still visibly laggy. Two causes: (1) every single-block
re-render re-opened the render shell from bytes — package open + styles/numbering parse +
converter style-cache rebuild on every keystroke commit (~50–100 ms of the ~55–135 ms per
op); (2) on any document containing a block-level `w:sdt` (a TOC — most real documents),
`ListBlocks` missed the sdt's content blocks, the plan/DOM diff read as 100 % churn, and
M2.6's reconcile silently fell back to the multi-second remount for EVERY structural op.
**Shipped:** the shell `WordprocessingDocument` stays open on the session and each render
replaces only the main part's body document (annotation caches on the formatting-part
XDocuments persist; `FormattingAssembler` caches its style indexes on the styles XDocument
behind a style-count guard); `ListBlocks` flattens `w:sdtContent` exactly as the renderer
does; reconcile substitutions pair by unid first; sig-less (single-block-swapped) wrapped
paragraphs may leaf-swap in place; `SetParagraphFormat`/`SplitParagraph` no longer rebuild
the whole anchor index per op; the per-op index-only rebuild no longer flushes parts (the
whole-main-part XML write per keystroke that only external save paths need); Enter renders
both split halves in one batched call. Measured (HC031, warm): text commit 102 → 30 ms,
Enter 137 → 41 ms, bold 108 → 52 ms, insert row 1.25 s → 85 ms, delete block 1.2 s → 24 ms,
undo/redo ~1.2 s → 124/54 ms. Standing instrument:
`npm/tests/editor-latency-bench.spec.ts`.

### M2.6 — Incremental STRUCTURAL repaint (reconcile replaces remount)  · effort L · ✅ **DONE**
**Problem:** every structural op (insert table/row/col, footnote/endnote, delete block,
undo/redo) paid a full remount — ~5–6 s on a 346-block document — and the engine ops
themselves carried ~1 s of per-op projection overhead (patch generation + full-projection
anchor lookup + unpruned Unid pass).
**Shipped (two tracks):** engine — `EmitMarkdownPatch` opt-out, index-only anchor lookup,
pruned deterministic Unid pass + conditional part flush (per-op rebuild 74 → 2 ms native);
client — `DocxEditor.reconcile()`: LCS diff of DOM unit sequence vs `ListBlocks` render
plan (`unid|contentHash` tokens; content hashes because in-session unids survive edits),
batched `RenderBlocksHtml` (one throwaway doc; real sibling context so contextualSpacing
resolves; live `ListItemRetriever` annotations transplanted so an isolated list item
renders its TRUE number — the M9 numbering-continuation gap, closed for swaps), and
positional note-chrome renumbering from `ListNotes`. Full remount remains the universal
fallback (paginated mode, pure li insert/remove, border-div regrouping, any error), with
`lastReconcileFallback` recording why. NVCA measured: insert table 6.1 s → 225 ms, insert
footnote 6.2 s → 210 ms, undo/redo 5.6 s → 200–360 ms, delete block 93 ms.
Pinned by `editor-reconcile-unit.spec.ts` + `editor-reconcile.spec.ts` (node identity,
remount equivalence, chrome renumber, save/reopen).

### M2.5 — Incremental multi-block ops + session-attached remount  · effort S · ✅ **DONE**
**Problem:** every multi-block ribbon action (`format`/`setAlignment`/`setFontSize`/… over a
multi-paragraph selection) fell back to `remount()` — a full-document convert (~1–2.5 s) per
click — while the single-block path swapped one block in ~10 ms; and `remount()` itself
marshaled the saved bytes WASM→JS→WASM (two multi-MB copies) before converting.
**Shipped:** the multi-block paths now apply each block's op and swap each edited block via
the session-attached `RenderBlockHtml` (exactly N single-block swaps — fidelity-identical to
the single-block path by construction), restoring the cross-block selection so consecutive
ribbon actions keep working. Full remount is kept only where whole-document context is real:
list-touching results (numbering), `clearBorders` (border-div regrouping), paginated mode
(reflow, until M4). `remount()` now renders through a new session-attached
`DocxSessionBridge.RenderHtml` (same option profile, byte-identical output, old-bundle
fallback to Save+Convert). Pinned by `npm/tests/editor-perf-incremental.spec.ts`:
node-identity proof that untouched blocks survive (no remount), selection restore,
save/reopen fidelity, and RenderHtml ≡ bytes-path parity.

### M3 — Worker offload  · effort M–L
**Problem:** the initial full convert (~0.7–2.4 s) and session ops run on the main thread →
the UI freezes on open and on big docs. (Per-edit is already ~10 ms.)
**Approach:** extend the Web Worker surface (`docxodus.worker.ts` / `worker-proxy.ts`) to
carry session open/edit/render-block/save, transfer bytes zero-copy; the main thread holds
only the DOM. Keep the synchronous `DocxEditor` API working by awaiting worker round-trips.
**Acceptance:** opening and editing a large doc never blocks the main thread > ~16 ms;
existing editor tests pass through the worker path.

### M4 — Re-paginate on edit  · effort M
**Problem:** in paginated mode an edited block can overflow its page box (the MVP patches in
place without reflowing).
**Approach:** after a commit in paginated mode, re-run pagination from the affected page
forward (staging originals are retained, so a scoped reflow is feasible); debounce.
**Acceptance:** an edit that grows a block past a page boundary reflows to a new page.

### M5 — Formatting controls + ribbon + undo/redo  · effort S–M · ✅ **DONE**
**Shipped:** `DocxEditor` command methods `format(key, value?)` (bold/italic/underline/
strike/code on the selection span via `ApplyFormat`, toggling off computed state),
`setParagraphStyle(styleId)` (via `SetParagraphStyle`), `undo()`/`redo()` (via
`DocxSession.Undo/Redo` + full re-render), and `queryFormatState()` for button highlighting.
Keyboard shortcuts Ctrl/Cmd+B/I/U and Ctrl+Z / Ctrl+Shift+Z (redo). The demo
(`examples/editor.html`) has a ribbon (B/I/U/S/code, style dropdown, undo/redo) that
preserves the editor selection via `mousedown`-preventDefault. Formatting routes through
DocxSession (lossless, supports underline/color, not just markdown). **Note:** the editor
now defaults to `fabricateClasses: false` (inline styles) so per-block re-renders stay
self-contained — fabricated class names are per-conversion and have no page stylesheet.
Test `editor.spec.ts` "M5" applies bold to a selection (survives save), sets Heading1
(+1 h1), and undoes it; verified live in the browser.

### M5b — Extended formatting controls (super/sub, alignment, indent, page break)  · effort M · ✅ **DONE**
**Shipped (new C# ops, rippled through the 8 layers):**
- **Superscript / subscript** — added `string? VertAlign` to `FormatOp`; `ApplyFormatToRun` emits
  `w:vertAlign` (super/sub/baseline). Auto-rides the existing `ApplyFormat` JSON path (no new
  bridge method). `editor.format('superscript'|'subscript')` toggles via `w:vertAlign`.
- **Alignment / indent / page-break** — new `DocxSession.SetParagraphFormat(anchor, ParagraphFormatOp{Alignment?, IndentDelta?, PageBreakBefore?})`
  writing `w:jc` / `w:ind/@w:left` (twips delta, clamped, sibling-preserving) / `w:pageBreakBefore`,
  with a CT_PPr `SetPPrChildInOrder` schema-ordering helper. Rippled: DocxSessionOps →
  DocxSessionJson (`ParseParagraphFormatOp`) → `DocxSessionBridge.SetParagraphFormat` → types.ts →
  session.ts (`setParagraphFormat`) → editor.ts (`setAlignment`/`indent`/`pageBreakBefore`) → ribbon.
- Demo ribbon gained x²/x₂, L/C/R/J, indent ⇤/⇥, and page-break buttons.
- Tests: C# `DS200`–`DS202` (vertAlign set/clear, jc, pageBreakBefore + accumulating indent);
  browser `M5b` (center renders `text-align:center`, indent → margin, superscript → `<sup>`).
  Verified live. **Note:** the editor uses inline styles (`fabricateClasses:false`), so the
  converter renders super/sub as `<sup>`/`<sub>`.

### Mlists — Bullets & numbered lists (promote plain paragraph → list item)  · effort L · ✅ **DONE**
`SetListLevel`/`RemoveListMembership` only work on *existing* list items. **Shipped:** new
`DocxSession.ApplyListFormat(anchor, ListFormat.None|Bullet|Decimal)` + `Internal/NumberingFactory`
that ensures the `NumberingDefinitionsPart` exists and **find-or-creates** a spec-valid 9-level
bullet/decimal `w:abstractNum` + `w:num` tagged by a fixed marker `w:nsid` (idempotent across
calls/save/reopen/undo — no cache needed), then sets/replaces the paragraph's `w:numPr` (ilvl
preserved, p→li flip via re-projection). The factory flushes the numbering part itself
(`PutXDocument`) since the session's `Save` only persists projected parts. Rippled through all 8
layers; editor `toggleList('bullet'|'decimal')` toggles via `GetListMembership`; demo ribbon has
•/1. buttons. Tests: C# `DS210`–`DS212` (promote+reuse, decimal→none, save/reopen round-trip);
browser `Mlists` (bridge promote + membership + remove, editor toggle re-renders **with a
visible bullet marker + hanging indent**). The marker renders correctly in both the full and the
single-block (incremental) paths — the session-attached render copies the numbering part, so the
converter's `ListItemRetriever` resolves the marker; C# `DS213` asserts the Symbol bullet (U+F0B7)
+ `text-indent`. Raw was confirmed NOT a shortcut (can't reach the numbering part). Remaining
nuance: per-item numbering *continuation* for a block rendered in isolation is whole-doc context
(M9), but the marker glyph itself shows.

### M6 — Tracked-changes / review mode  · effort M
**Approach:** open the session with `TrackedChanges = RenderInline`; render `ins`/`del` with
author colors; serve the redline/review use case.
**Acceptance:** edits land as `w:ins`/`w:del` with author attribution, visible in the editor.

### M7 — Table-cell & table-structure editing  · effort M · ✅ **DONE**
**Shipped (resolving the S-1 smoke-test gaps):** cell text edits/round-trips; **Enter inside a
cell** splits the cell paragraph in place (stacked lines — value over label, multi-line
addresses); first-class row/column ops `DocxSession.{InsertTableRow,InsertTableColumn,
DeleteTableRow,DeleteTableColumn}` (by a canonical `tc` anchor; deleting the last row/col removes
the table) surfaced through the bridge + `DocxEditor` + a floating table toolbar; explicit
`tbl`/`tr`/`col`/`tc` metadata and anchor ↔ grid-coordinate resolution; deterministic structural
anchor mappings; span-aware merge/unmerge and ragged-grid CRUD; per-column
`TableInsertOptions.ColumnWidths`; a visual table grid picker in the demo. Tests: C#
`DocxSessionTableEditTests`, `DocxSessionTableAddressingTests` DT250–DT257, and
`DocxSessionS1FeaturesTests` DS214/DS215; browser `editor-cell-multiparagraph` /
`editor-table-edit` / `editor-table-colwidths` / `editor-demo-grid`.
**Remaining:** drag-to-resize columns; programmatic widths are first-class.

### M8 — React wrapper  · effort S
**Approach:** `useDocxEditor` hook + `<DocxEditor>` component over the pure-TS core, in
`npm/src/react.ts`.
**Acceptance:** a React app mounts the editor with one component.

### M9 — Single-block render fidelity  · effort M · ◐ list numbering DONE
**Approach:** copy image parts into the throwaway render doc; ~~resolve list-numbering
continuation for a block rendered in isolation~~ (done in M2.6: the batch render
transplants the live document's `ListItemRetriever` annotations onto the clones, so an
isolated list item renders its true number; `HCO081` pins element-identical output vs
the full render across li/p/h units).
**Remaining:** image parts (an inline image still loses its uncopied part).
**Acceptance:** re-rendering an image-bearing block matches the full render.

## Recommended sequencing

**M1 + M2 done** — "make editing real" is complete (edits preserve formatting; Enter/Backspace
split/merge). A runnable demo exists (`npm run demo` → `editor.html`). **M3 next** (worker
offload) for responsiveness on large docs. M4–M9 sequence by target use case: authoring favors
M4/M5; review favors M6; broad fidelity favors M7/M9. M8 (React) any time.

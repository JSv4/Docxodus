# IR-Powered DOCX Editor — Roadmap

Companion to `ir_editor_feasibility.md` (which records the verdict, architecture, and
PoC results). This is the **sequenced, prioritized** plan for turning the proven
foundation + MVP into a complete editor. Supersedes the scattered "Still Plan 2" notes.

Status (branch `feat/ir-editor-feasibility-poc`, PR #234): **foundation + MVP shipped and
proven; M1 (rich in-block editing), M2 (structural editing), M5 + M5b (formatting controls:
bold/italic/underline/strike/code, super/sub, alignment, indent, page break, paragraph style,
undo/redo) done; runnable demo with a full ribbon at `npm/examples/editor.html` (`npm run
demo`); Mlists (bullets/numbered) done — all 7 requested controls shipped.** M3 (worker
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

### M7 — Table-cell & table-structure editing  · effort M · ✅ **DONE** (except cell-merge)
**Shipped (resolving the S-1 smoke-test gaps):** cell text edits/round-trips; **Enter inside a
cell** splits the cell paragraph in place (stacked lines — value over label, multi-line
addresses); first-class row/column ops `DocxSession.{InsertTableRow,InsertTableColumn,
DeleteTableRow,DeleteTableColumn}` (by a cell-paragraph anchor; deleting the last row/col removes
the table) surfaced through the bridge + `DocxEditor` + a floating table toolbar; per-column
`TableInsertOptions.ColumnWidths`; a visual table grid picker in the demo. v1 assumes a
rectangular grid (no `w:gridSpan`). Tests: C# `DocxSessionTableEditTests` DT201–DT207 +
`DocxSessionS1FeaturesTests` DS214/DS215; browser `editor-cell-multiparagraph` /
`editor-table-edit` / `editor-table-colwidths` / `editor-demo-grid`.
**Remaining:** horizontal/vertical **cell merge** (`w:gridSpan`/`w:vMerge`) and drag-to-resize
columns — still via `session.Raw.*` for now.

### M8 — React wrapper  · effort S
**Approach:** `useDocxEditor` hook + `<DocxEditor>` component over the pure-TS core, in
`npm/src/react.ts`.
**Acceptance:** a React app mounts the editor with one component.

### M9 — Single-block render fidelity  · effort M
**Approach:** copy image parts into the throwaway render doc; resolve list-numbering
continuation for a block rendered in isolation.
**Acceptance:** re-rendering an image-bearing or list-item block matches the full render.

## Recommended sequencing

**M1 + M2 done** — "make editing real" is complete (edits preserve formatting; Enter/Backspace
split/merge). A runnable demo exists (`npm run demo` → `editor.html`). **M3 next** (worker
offload) for responsiveness on large docs. M4–M9 sequence by target use case: authoring favors
M4/M5; review favors M6; broad fidelity favors M7/M9. M8 (React) any time.

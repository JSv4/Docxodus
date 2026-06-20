# Multi-Block Selection & Actions (browser editor)

How the in-browser `DocxEditor` (`npm/src/editor.ts`) lets a user select across several body
blocks and apply actions to the whole selection. This is the write-side counterpart to the
single-block editing the editor already did; the design spec is
`docs/superpowers/specs/2026-06-19-multi-block-selection-design.md`.

## Problem it solves

The editor used to render every block as its own `contenteditable="true"` element. Browsers confine
a native selection to one editing host, so a real user could not drag, Shift+Click, or Shift+Arrow
across block boundaries — only one block could be selected and acted upon. The multi-block *action*
plumbing already existed but was only reachable via a programmatically-built `Range`.

## Architecture: controlled single root

The editing host is now a **single `contenteditable` surface** (`makeRootEditable`), so a native
selection genuinely spans blocks. Body blocks are **not** their own editing hosts (nested same-value
`contenteditable` would re-introduce the boundary); instead each is:

- marked `data-editable="1"` (the canonical `EDITABLE_SELECTOR`, decoupled from the `contenteditable`
  attribute), and
- made programmatically focusable with `tabindex="-1"` (so `block.focus()` still enters it and
  `document.activeElement` can be the block — `tabindex` is not an editing host, so it does **not**
  create a selection boundary), with `outline: none` to suppress the focus ring.

**Tables are a selection boundary (v1).** A table is wrapped `contenteditable="false"` and each cell
paragraph stays its own `contenteditable="true"` island, so a body selection can't cross into/through
a table and cell editing is unchanged. The `SelectionModel` also filters out any block inside a
table.

### Event model (single root)

| Event | Where | Why |
|---|---|---|
| `focus` | per block | `block.focus()` sets the active target synchronously (`activeBlock`) |
| `keydown` | root | focus stays on the host during editing, so keydown fires on the root; the handler resolves the target block via `keydownBlock` (event target → caret → activeBlock) |
| `beforeinput` | root | classify + route structural/cross-block input (see below) |
| `selectionchange` | document | track `activeBlock` from the selection start; **commit the block a collapsed caret just left** (the primary between-block commit, since focus never leaves the host) |
| `blur` (capture) | root | commit on focus leaving the editor; **ignores transient internal focus shifts** (`relatedTarget` inside the editor — e.g. setting a selection moves focus block→host) so it doesn't spuriously re-render a block mid-interaction |

The root listeners are attached once (`data-editor-root-wired`) so a non-paginated remount (which
reuses the same container) doesn't double-attach them.

### `beforeinput` routing (`editor-input.ts`)

`classifyBeforeInput(inputType, data, isMultiBlock)` is pure (no DOM). Single-block input stays
**native** (the browser edits the text node; the change commits on blur / selection-leave). Only
structural / cross-block operations are intercepted:

| inputType | single-block | multi-block selection |
|---|---|---|
| `insertText` / `insertReplacementText` / IME | native | `typeOver` |
| `insertParagraph` (Enter) | native (onKeydown splits) | `splitAtSelection` |
| `insertLineBreak` / `deleteContent*` / `deleteWord*` / `deleteByCut` | native | `deleteSelection` |
| `insertFromPaste` | native | `paste` (plain text, v1) |
| `formatBold/Italic/…` | `format()` | `format()` across blocks |
| `deleteByDrag` / `insertFromDrop` | blocked (v1) | blocked (v1) |

`onKeydown` lets **Enter** fall through to `beforeinput` when the selection is multi-block (Backspace
/Delete already do, because their handlers require a collapsed caret), so the compound handler runs
instead of splitting the active block.

## Selection model (`editor-selection.ts`)

`readSelection(editRoot, deps)` maps the native selection to `MultiBlockSelection`: the ordered
covered **body** blocks (table cells excluded) with each block's clipped content span and a role
(`only` / `first` / `middle` / `last`). The offset helpers (`contentOffsetOf`, `blockContentText`,
`selectionSpanIn`) are injected from `editor.ts` (which owns the content-offset space — list markers
and bidi marks excluded) to avoid an import cycle.

`selectedBlocks()` returns all covered body blocks for a multi-block selection (otherwise the active
block), feeding the existing `applyInlineOpAcrossBlocks` / `applyParagraphOpAcrossBlocks`.

## Compound cross-block edits

`deleteSelectionInner(model)` performs, by anchor (re-threading ids across each op since a text edit
re-hashes a block's unid):

1. trim the first block's selected tail (`ReplaceTextAtSpan`),
2. trim the last block's selected head (`ReplaceTextAtSpan`),
3. delete whole middle blocks (`DeleteBlock`),
4. merge the trimmed last block into the first (`MergeParagraphs`).

`MergeParagraphs` deliberately inserts a sentence-joining space (right for a Backspace-join, wrong
for a delete-selection), so step 5 **removes that one join space** when its exact condition held
(both join sides end/start non-whitespace) — the delete joins seamlessly (`Al`+`arlie` → `Alarlie`)
while keeping the moved runs' formatting.

- `deleteSelection` = collapse + remount + caret at the join.
- `typeOverSelection` = collapse → remount → place caret → native `insertText` → commit.
- `splitAtSelection` = collapse → remount → place caret → `splitAtCaret`.
- `handlePaste` = `typeOverSelection` with the clipboard's plain text.

All run inside one `group()` (see below), so each compound edit is one atomic undo. The insert/split
reuse the proven single-block native path rather than fragile post-merge span math.

## Atomic undo (client-side grouping)

The editor is the sole driver of the session's Undo/Redo, so atomicity lives in `editor.ts` with no
C# change. `parseEdit` (used only for mutation results) records each successful mutation: ungrouped →
a group of 1 (single-edit UX unchanged); inside `group(fn)` → folded into one unit. `undo()`/`redo()`
reverse/replay a whole group (N session ops); `redoGroups` mirrors and is cleared by any new edit.
Multi-block format/paragraph ops and the compound edits all run in a `group()`.

## Invariants preserved

1. **Model-of-record = live OOXML in `DocxSession`.** All edits route through session ops by anchor.
2. **One `{#kind:scope:unid}` anchor scheme.** The `data-editable` marker is orthogonal to addressing.
3. **Render is a projection, patched incrementally.** Typing keeps `RenderBlockHtml`; structural ops
   remount (as the editor's structural commands already did).
4. **Untouched content stays byte-faithful on save.** Compound ops only touch the selected blocks; a
   test asserts untouched outer paragraphs survive a cross-block delete.

## v1 limits

- Tables are a selection boundary (no cross-cell selection).
- Plain-text paste only; rich/HTML paste is future work.
- Cross-block drag-and-drop of content is disabled.
- Tested in Chromium (the existing Playwright harness); Firefox/Safari `beforeinput` parity is future.

## Tests

- `editor-multiblock-reach.spec.ts` — real mouse-drag, Shift+Click, Shift+ArrowDown reach a
  multi-block selection that commands act on.
- `editor-multiblock-format.spec.ts` — multi-block formatting (alignment + bold) round-trips.
- `editor-multiblock-delete.spec.ts` — cross-block Delete/Backspace (trim/delete/merge, seamless
  join, untouched outer paragraphs intact).
- `editor-multiblock-edit.spec.ts` — type-over, Enter-split, and one-undo round-trip to pre-edit.
- `editor-undo-group.spec.ts` — a multi-block format is one undo.
- `editor-multiblock-table-boundary.spec.ts` — a selection into a table only formats body blocks.

# Full-fidelity inline formatting on block edit

**Date:** 2026-06-18
**Component:** `npm/src/editor.ts` (`DocxEditor`)
**Status:** approved design, pending implementation plan

## Problem

When a block is edited in `DocxEditor`, the commit path re-serializes the block to
the markdown projection subset and rebuilds every run:

```
commitBlock / syncBlock → serializeInlineMarkdown(el) → DocxSession.ReplaceText(anchor, markdown)
```

The markdown subset expresses only **bold / italic / links**, so editing a block
**drops** underline, strikethrough, color, font size, font family, highlight, and
super/subscript from that block. The header comment in `editor.ts` documents this as
an MVP limitation. We want edits to preserve **all** run properties (full rPr
fidelity), not just the markdown-expressible ones.

## Why a re-serialize approach can't reach full fidelity

Reconstructing runs from the rendered DOM's *computed CSS* (Approach B, rejected)
cannot round-trip themed colors, exact half-point `w:sz`, `w:lang`, kerning/spacing,
complex-script properties, etc. The only way to guarantee **all** rPr survives is to
**never rebuild untouched runs**.

## Approach (chosen): minimal prefix/suffix-diff via `ReplaceTextAtSpan`

`DocxSession.ReplaceTextAtSpan(anchor, spanStart, spanLength, replace)` already:

- rewrites only the runs *inside* the span; runs outside are byte-identical
  (exact rPr preserved);
- drops the replacement text into the **first** spanned run with its `rPr` untouched
  (the "formatting-preservation contract" in `ApplyFragmentReplacement`), so typed
  text inherits the boundary run's formatting like Word / contenteditable;
- returns the **same** anchor (edits in place — no unid churn).

It is already exposed through every bridge layer: `DocxSessionBridge.ReplaceTextAtSpan`,
`DocxSessionOps.ReplaceTextAtSpan`, and npm `types.ts` (`ReplaceTextAtSpan` on
`DocxodusWasmExports`). **No .NET / WASM / npm-bridge changes are required** — the
change is contained to `editor.ts`.

### Commit algorithm

Replaces the `serializeInlineMarkdown → ReplaceText` call in **both** `commitBlock`
and `syncBlock`:

```
old = el.dataset.committedText      // content text: list markers + bidi marks excluded
new = blockContentText(el)          // same offset space
if (old === new) return             // no-op

P = longest common prefix length (old, new)
S = longest common suffix length (old, new), capped so P + S ≤ min(old.length, new.length)
changedOldLen = old.length - P - S
newMiddle     = new.slice(P, new.length - S)

ReplaceTextAtSpan(handle, fullId, P, changedOldLen, newMiddle)
```

### Offset-alignment invariant

`blockContentText(el)` (list markers + injected bidi marks excluded — see the bidi fix
in commit `7a53439`) equals the session's flat run-text space, and `committedText` is
re-derived from the session-rendered HTML after each commit. Therefore the diff offsets
map directly onto `ReplaceTextAtSpan`'s span coordinates. This is the **same** invariant
`splitAtCaret` / `mergeWithPrevious` already depend on (`caretOffsetIn → SplitParagraph`),
so it is already assumed and tested. `committedText` is standardized to **content text**
(not raw `textContent`) at every site that sets it: `wireBlock`, `commitBlock`,
`syncBlock`, and the list-item branch of `commitBlock`.

### Edge cases

| Case | Handling |
|------|----------|
| **Pure insertion** (`changedOldLen === 0`) | Never pass a zero-length span (resolves to no runs → fails). Anchor a neighbor char: if `P > 0`, use span `(P-1, 1)` with `newMiddle = old[P-1] + newMiddle` (inherits the **left** run); if `P === 0`, use span `(0, 1)` with `newMiddle = newMiddle + old[0]` (inherits the first run). |
| **Empty paragraph → text** (`old.length === 0`) | No char to anchor → fall back to `ReplaceText(new)`. An empty block has no inline formatting to preserve, so zero fidelity loss. |
| **Select-all replace** (P = S = 0, whole span) | New text inherits the first run's rPr; old runs' formatting is gone — expected, since no text is untouched. |
| **Pure deletion** (`newMiddle === ""`) | Non-zero span, `replace = ""` — supported. |
| **List items / bidi paragraphs** | Unchanged: markers and bidi marks are already excluded from `blockContentText`, so they never enter the diff. |

### Re-render & focus behavior

Unchanged. `commitBlock` still skips re-render for list items (a re-render during the
blur cancels the browser's in-flight focus transfer) and re-renders plain blocks from
the live session for canonical HTML. Only the underlying session op changes. Because
`ReplaceTextAtSpan` returns the same anchor (no unid churn), `syncBlock` keeps the id
stable across the flush-before-structural-op path.

### Error handling

If `ReplaceTextAtSpan` returns a non-success result, revert the DOM to `committedText`
(identical to today's `ReplaceText`-failure path), keeping the view in sync with truth.

### Disposition of `serializeInlineMarkdown`

Left exported (removing an export would break npm consumers) but no longer on the commit
path. M1's "preserve bold/italic/link on edit" behavior is subsumed and strictly improved:
all run formatting now survives structurally, not just the markdown-expressible subset.

## Ripple

`editor.ts` only:

1. Add `ReplaceTextAtSpan: (handle, anchor, spanStart, spanLength, replace) => string` to
   the `DocxEditorExports.DocxSessionBridge` interface (it is already present on the real
   bridge and on `DocxodusWasmExports`).
2. Rewrite `commitBlock` and `syncBlock` per the algorithm above; standardize the
   `committedText` baseline to content text.

## Testing

New `editor.spec.ts` cases against a rich-formatted fixture (`HC031-Complicated-Document.docx`
carries bold / colored / sized runs):

1. **Untouched-run fidelity:** type into the middle of a formatted run, blur, `save()`,
   reopen → the untouched runs' formatting survives the round-trip.
2. **Boundary inheritance:** type at a run boundary → the new text carries the expected
   neighbor formatting.
3. **Cross-run delete:** delete a selection spanning two differently-formatted runs →
   the surviving text keeps its formatting.

`ReplaceTextAtSpan`'s rPr-preservation contract already has .NET coverage; no new .NET
tests are required, though one may be added if a gap surfaces.

## Risks

- **Offset divergence:** if `blockContentText` ever diverges from the session's flat
  run-text for some character, the diff span maps to the wrong runs → visible corruption.
  Mitigated by the by-construction invariant (`committedText` is derived from
  session-rendered HTML), the DOM-revert safety net on failure, and the new round-trip
  tests. This is the same invariant the existing split/merge already trust.
- **Non-contiguous multi-region edits in one blur** widen the replaced span (Approach C
  territory). Correct, but loses formatting on any untouched runs *between* the scattered
  edits. Real block-edit commits are contiguous; deferred unless it bites.

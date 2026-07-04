# S-1 cover-page smoke test — editor feature build

Goal: draft an SEC **Form S-1 cover page** entirely through the in-browser
`DocxEditor` (model-of-record = `DocxSession`), then render via the converter
(the faithfulness oracle) and confirm it matches the target.

## Gaps found by smoke-testing the existing editor

| # | S-1 element | Capability needed | Status before |
|---|-------------|-------------------|---------------|
| F0 | Drafting from scratch | New blank document | ❌ editor only opens existing bytes |
| F1 | "FORM S-1", company name (large) | run **font size** (`w:sz`) | ❌ not in `FormatOp` |
| F2 | ~6 thick horizontal rules | paragraph **borders / HR** (`w:pBdr`) | ❌ not in `ParagraphFormatOp` |
| F3 | 3-col value/label rows; right-aligned filing header; law-firm columns | **table insertion** (`w:tbl`) | ❌ no `InsertTable` (cells editable, can't create) |

Already present and used: center/right/justify alignment, bold/italic/underline,
indent, page-break, paragraph styles, split/merge, undo/redo, lossless save.

## Layers each feature ripples through (single-owner facade first)

`DocxSession.cs` (engine) → `Internal/DocxSessionOps.cs` (facade) →
`Internal/DocxSessionJson.cs` (hand-written wire parsers — WASM AOT disables
reflection JSON) → `wasm/DocxodusWasm/DocxSessionBridge.cs` (`[JSExport]`) →
`npm/src/types.ts` → `npm/src/session.ts` + `npm/src/editor.ts` →
`npm/examples/editor.html` (demo toolbar).

## Engine surface added

- `FormatOp.FontSizePts` (double, points → `w:sz`/`w:szCs` half-points).
- `ParagraphBorderEdge` record + `ParagraphFormatOp.{TopBorder,BottomBorder,ClearBorders}`.
- `DocxSession.InsertHorizontalRule(anchor, pos, ParagraphBorderEdge?)` — empty `w:p` with a bottom border.
- `DocxSession.InsertTable(anchor, pos, rows, cols, TableInsertOptions?)` — `w:tbl`, borderless option, row-major `CellContents`, optional `CellAlignment`.
- `DocxSession.CreateBlankDocxBytes()` (static) — complete blank DOCX (Normal style, settings, US-Letter sectPr).

Tests: `Docxodus.Tests/DocxSessionS1FeaturesTests.cs` (DS201–DS210), incl. DS210
which builds an S-1-style page with all four features and asserts the OOXML is
schema-valid via `OpenXmlValidator` (zero errors).

## Smoke-test result

Drafted the full SpaceX S-1 cover page **end-to-end through the editing surface**
(blank doc → `InsertTable`/`InsertParagraph`/`ApplyFormat`/`SetParagraphFormat`/
`InsertHorizontalRule`), saved, and rendered via the converter (the faithfulness
oracle). The render matches the target: justified filing header, double top rule,
centered bold headings, large "FORM S-1", bold-italic registration statement, large
company name, the 3-column registrant-facts table, centered address/agent blocks,
italic "With copies to:", and the 3-column counsel table.

Also verified through the **live `DocxEditor`** (not just the bridge): "New" blank
doc → type a heading → `setFontSize(22)` + `format('bold')` + `setAlignment('center')`
(live single-block re-render) → `insertHorizontalRule` → `insertTable` (borderless,
seeded, centered) → edit a table cell → `save()` → re-open the saved bytes as a fresh
session (lossless round-trip; the cell edit and all content survive).

Build/run the demo: `cd npm && npm run demo` → http://localhost:8088/editor.html
(New / Size / rule / table buttons in the ribbon).

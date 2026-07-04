# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Important Instructions

- **Never credit yourself in commits.** Do not add "Generated with Claude Code" or "Co-Authored-By: Claude" to commit messages.

## Coding Standards

### Nullable Reference Types

The project has `<Nullable>disable</Nullable>` globally due to ~9,000 warnings in legacy code. However, **new code should use nullable annotations**:

- **New files**: Add `#nullable enable` at the top of the file
- **Substantial refactors**: When significantly modifying an existing file, consider adding `#nullable enable` and fixing warnings in that file
- **Use proper annotations**: Mark nullable parameters/returns with `?`, use null checks or `!` where appropriate

```csharp
#nullable enable

namespace Docxodus;

public class MyNewClass
{
    public string Name { get; set; } = string.Empty;  // Non-nullable with default
    public string? Description { get; set; }           // Explicitly nullable

    public string? FindItem(string key)                // May return null
    {
        // ...
    }
}
```

See [Issue #13](https://github.com/JSv4/Docxodus/issues/13) for the full nullable migration plan.

## Feature Development Workflow

When implementing new features or significant changes, follow this workflow:

### 1. Documentation Updates

- **CHANGELOG.md** - Add entry under `[Unreleased]` section describing the feature/fix
- **CLAUDE.md** - Update if the feature adds new settings, modules, or changes architecture
- **docs/architecture/** - Create or update architecture docs for significant features (e.g., `comment_rendering.md`, `comparison_engine.md`)
- **docs/ooxml_corner_cases.md** - Document any OOXML edge cases where Word's behavior differs from spec or our implementation (see below)

### 2. Test Updates

- Add tests to the appropriate test file in `Docxodus.Tests/`:
  - `HtmlConverterTests.cs` - WmlToHtmlConverter features
  - `WmlComparerTests.cs` - Document comparison features
  - `DocumentBuilderTests.cs` - Document merging/splitting
  - Use existing test files from `TestFiles/` when possible
  - When creating programmatic test documents, ensure all required parts exist (StyleDefinitionsPart, DocumentSettingsPart, etc.)

### 3. WASM/npm Wrapper Updates

Update these when adding new settings or methods to the .NET API:

- **wasm/DocxodusWasm/DocumentConverter.cs** - Add new JSExport methods or parameters
- **wasm/DocxodusWasm/DocumentComparer.cs** - For comparison-related changes
- **npm/src/types.ts** - Add TypeScript types, enums, and update `DocxodusWasmExports` interface
- **npm/src/index.ts** - Update wrapper functions to use new WASM methods

Build and verify with:
```bash
npm run build          # Builds WASM and TypeScript
dotnet test            # Run .NET tests
```

### 4. When to Update Each Layer

| Change Type | .NET | Tests | WASM | npm/TS | Docs |
|-------------|------|-------|------|--------|------|
| New converter setting | ✓ | ✓ | ✓ | ✓ | ✓ |
| Bug fix | ✓ | ✓ | - | - | CHANGELOG |
| New public enum | ✓ | ✓ | ✓ | ✓ | ✓ |
| Internal refactor | ✓ | ✓ | - | - | - |
| New module | ✓ | ✓ | ✓ | ✓ | ✓ |

## Release Process

Releases are **CHANGELOG + annotated git tag only** — no version bump in
`Docxodus.csproj` (`<Version>` is intentionally left at `1.0.0`) or
`npm/package.json`; those are not tied to the release tags.

Versioning is **semver** on tags of the form `vMAJOR.MINOR.PATCH` (e.g.
`v6.3.0`). Pick the bump from what landed in `[Unreleased]` since the last tag:

| Bump | When |
|------|------|
| Patch (`v6.2.0` → `v6.2.1`) | only `### Fixed` entries (bug fixes) |
| Minor (`v6.2.0` → `v6.3.0`) | any `### Added`/`### Changed` (new feature/surface, no breaking change) |
| Major (`v6.x` → `v7.0.0`) | a breaking public-API change |

To cut a release from an up-to-date `main`:

1. Edit `CHANGELOG.md`: insert `## [X.Y.Z] - YYYY-MM-DD` immediately under the
   `## [Unreleased]` header, leaving the accumulated `### Added`/`### Fixed`/etc.
   entries beneath it as the released section and `## [Unreleased]` empty above.
   (Do not add per-release version commits to csproj/package.json.)
2. Commit changelog-only with message `docs(changelog): cut vX.Y.0 release notes`.
3. Create an **annotated** tag whose message is the version string:
   `git tag -a vX.Y.Z -m vX.Y.Z`.
4. `git push origin main` then `git push origin vX.Y.Z`.

Prior release-cut commits (`#206`, `#209`) and tags (`v6.1.0`, `v6.2.0`) are the
reference for the exact diff shape.

## Build Commands

```bash
# Build the entire solution
dotnet build Docxodus.sln

# Build specific project
dotnet build Docxodus/Docxodus.csproj

# Release build — warnings are errors (Directory.Build.props)
dotnet build -c Release Docxodus.sln

# Build the WASM target (sets WASM_BUILD; excludes SkiaSharp)
./scripts/build-wasm.sh

# Build the npm package end-to-end (runs build-wasm.sh + tsc + esbuild bundles)
cd npm && npm run build
```

`TreatWarningsAsErrors=true` is set for Release config only (`Directory.Build.props`). Debug builds tolerate the ~9,000 legacy nullable warnings; do not regress Release.

## Test Commands

```bash
# Run all .NET tests
dotnet test Docxodus.Tests/Docxodus.Tests.csproj

# Run a specific test by name (test IDs are prefixed by feature, e.g. WC001, DB001)
dotnet test --filter "FullyQualifiedName~DB001_DocumentBuilderKeepSections"

# Run tests for a specific test class
dotnet test --filter "FullyQualifiedName~DbTests"

# Browser/WASM tests (Playwright) — must rebuild npm package first
cd npm
npm install                       # first time only
npx playwright install chromium   # first time only
npm run build                     # produces dist/ which the harness loads
npm test                          # run all Playwright specs
npx playwright test --grep "Document Structure"  # single test by name
npx playwright test --headed      # see the browser
npx tsc --noEmit                  # TS type-check only
```

Playwright tests serve from `npm/dist/wasm/` — if you edit C#, .ts, or the harness HTML, re-run `npm run build` (or at minimum the relevant `build:*` script) before re-running tests, or you will test stale artifacts.

## Architecture Overview

Docxodus is a library for manipulating Open XML documents (DOCX, XLSX, PPTX) built on top of the Open XML SDK. It is a fork of OpenXmlPowerTools upgraded to .NET 10.0. All code is in the `Docxodus` namespace.

### Repository Layout

This repo is not just a .NET library — it ships a four-layer stack. Changes to public surface usually need to ripple through all of them:

| Layer | Path | Purpose |
|-------|------|---------|
| Core library | `Docxodus/` | The .NET library — all OOXML logic lives here. NuGet package `Docxodus`. |
| Bridge core | `Docxodus/Internal/{SessionRegistry,DocxSessionOps,DocxSessionJson}.cs` | Shared handle pool + per-op session-lookup-and-serialize facade + JSON helpers. Both the WASM bridge and the stdio host route through these — wire shapes live in exactly one place. |
| Unit tests | `Docxodus.Tests/` | xUnit tests for the core library (~1,000+ tests). |
| CLI tools | `tools/redline/`, `tools/docx2html/`, `tools/docx2oc/` | Thin `dotnet tool`-installable wrappers over the library. |
| WASM bridge | `wasm/DocxodusWasm/` | `[JSExport]` shells (`DocumentConverter.cs`, `DocumentComparer.cs`, `DocxSessionBridge.cs`) exposing the library to JS via .NET WASM. `DocxSessionBridge` is now a thin passthrough to `DocxSessionOps`. |
| Stdio host | `tools/python-host/` | .NET 10 console binary (`docxodus-pyhost`) that reads NDJSON requests on stdin and dispatches to `DocxSessionOps`. The upcoming python-docxodus pip package will subprocess this. |
| npm/TypeScript | `npm/` | Wrapper around the WASM bridge — `src/index.ts` is the public API, `src/react.ts` is the React hook layer, `src/docxodus.worker.ts`/`worker-proxy.ts` run WASM off the main thread. |
| Web demo | `web/DocxodusWeb/` | Blazor/web demo app (separate workflow). |

When the core library changes a public method or setting on `DocxSession`, update **`Docxodus/Internal/DocxSessionOps.cs` first** — both bridges and both clients pick up the change automatically. Then ripple through: tests, the WASM `[JSExport]` shell in `DocxSessionBridge.cs`, the stdio dispatcher in `tools/python-host/Dispatcher.cs`, `npm/src/types.ts` + `npm/src/index.ts`, `python/src/docx_scalpel/types.py` + `python/src/docx_scalpel/session.py`. The table in "Feature Development Workflow" below summarizes when each is required.

The same single-owner-facade pattern applies to the **stateless** surfaces (no session handle): `HtmlConversionOps` owns DOCX→HTML, and `DocxDiffOps` (`Docxodus/Internal/DocxDiffOps.cs`) owns the public `DocxDiff` engine (Compare / GetRevisions / GetEditScriptJson, plus the byte→byte `AcceptRevisions`/`RejectRevisions` primitive clients use to verify a redline's round-trip contract — accept ≡ right, reject ≡ left). Both the WASM bridge (`DocxDiffBridge.cs`) and the stdio dispatcher route through `DocxDiffOps`, so the settings-in / revisions-out JSON wire shapes — and any future change to them — live in exactly one place. The corresponding clients are `npm/src/index.ts`'s `docxDiff*` wrappers (`DocxDiffBridge` on `DocxodusWasmExports`) and `docx-scalpel`'s `docx_diff_*` module functions (`python/src/docx_scalpel/session.py`, with the `DocxDiff*` types/enums in `types.py`/`enums.py`). When a stateless surface changes, update its `*Ops` facade first, then ripple the two bridges + two clients exactly as for `DocxSession`.

### WASM Conditional Compilation

The core library compiles in two modes controlled by the `WASM_BUILD` MSBuild property (set by `scripts/build-wasm.sh`):

- **Default build**: includes `SkiaSharp` + `SkiaSharp.NativeAssets.Linux.NoDependencies` for image/font work.
- **`WASM_BUILD=true`**: defines the `WASM_BUILD` constant, excludes SkiaSharp (no native deps in the browser). Code that needs SkiaSharp must be guarded with `#if !WASM_BUILD` or routed through a no-op fallback. See `docs/architecture/skiasharp-removal-plan.md`.

When touching image/font/color code, check whether your change compiles under `WASM_BUILD` before shipping — the npm build will fail loudly if it doesn't.

**WASM-mode output isolation:** the WASM-mode `Docxodus.dll` (no `SkiaSharp`, no `ImageInfo.SaveImage`) builds into its own `Docxodus/bin/wasm/` + `Docxodus/obj/wasm/` paths (see the `WASM_BUILD` PropertyGroup in `Docxodus.csproj`), so neither `scripts/build-wasm.sh` nor a solution build (which compiles Docxodus twice — `DocxodusWasm` references it with `WASM_BUILD=true`) can clobber the default-mode assembly. If you ever see `error CS1061: 'ImageInfo' does not contain a definition for 'SaveImage'` again, a stale pre-isolation artifact is lingering — delete `Docxodus*/bin` + `Docxodus*/obj` once; no recurring `dotnet clean` ritual is needed.

### Document Wrapper Classes

The library uses in-memory byte array wrappers for documents:
- `DocxodusDocument` - Base class holding `DocumentByteArray` and `FileName`
- `WmlDocument` - Word documents (.docx)
- `SmlDocument` - Spreadsheet documents (.xlsx)
- `PmlDocument` - Presentation documents (.pptx)

These allow immutable-style document manipulation via `OpenXmlMemoryStreamDocument` pattern:
```csharp
using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
{
    using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
    {
        // modify document
    }
    return streamDoc.GetModifiedWmlDocument();
}
```

### Core Modules

**DocumentBuilder.cs** - Merge/split DOCX files. Uses `Source` objects to specify document ranges:
```csharp
var sources = new List<Source> { new Source(wmlDoc, keepSections: true) };
DocumentBuilder.BuildDocument(sources, outputPath);
```

**WmlComparer.cs** - Compare two DOCX files, producing a document with tracked revisions. Supports nested tables and text boxes. Key settings in `WmlComparerSettings`:
- `AuthorForRevisions` - Author name for tracked changes
- `DetailThreshold` - 0.0-1.0, lower = more detailed comparison (default: 0.15)
- `CaseInsensitive` - Case-insensitive comparison
- `DetectMoves` - Enable move detection in `GetRevisions()` (default: true)
- `SimplifyMoveMarkup` - Convert move markup to del/ins (default: false)
- `MoveSimilarityThreshold` - Jaccard similarity threshold for moves (default: 0.8)
- `MoveMinimumWordCount` - Minimum words for move detection (default: 3)
- `DetectFormatChanges` - Enable format change detection (default: true)

Move detection produces **native Word move markup** (`w:moveFrom`/`w:moveTo`) when `DetectMoves` is enabled:
- The comparer analyzes deleted/inserted content blocks for similarity after LCS comparison
- Matching pairs (≥80% Jaccard similarity by default) are converted to move markup
- The output document contains `w:moveFromRangeStart`/`w:moveFromRangeEnd` and `w:moveToRangeStart`/`w:moveToRangeEnd` elements
- Move pairs are linked via the `w:name` attribute (e.g., "move1")
- `GetRevisions()` recognizes this native markup and returns `WmlComparerRevisionType.Moved` revisions
- `WmlComparerRevision.MoveGroupId` links source and destination revisions
- `WmlComparerRevision.IsMoveSource` - true = moved FROM here, false = moved TO here

Format change detection produces **native Word format change markup** (`w:rPrChange`) when `DetectFormatChanges` is enabled:
- The comparer analyzes Equal atoms (same text content) for run property differences after LCS comparison
- When text is identical but formatting differs (bold, italic, font size, etc.), atoms are marked as FormatChanged
- The output document contains `w:rPrChange` elements inside `w:rPr` with the old formatting properties
- `GetRevisions()` recognizes this native markup and returns `WmlComparerRevisionType.FormatChanged` revisions
- `WmlComparerRevision.FormatChange` contains details about what changed (old/new properties, changed property names)

**DocxDiff.cs** - Public facade over the IR diff engine — a structure-aware, anchor-addressed DOCX comparison engine and the diff-side counterpart to `WmlToMarkdownConverter`/`DocxSession`. The NEW engine; `WmlComparer` remains the default/blessed comparison API (`DocxDiff` ships as a production-candidate pending Word manual-verification + burn-in — decision D4):
- `Compare(left, right, settings?)` → `WmlDocument` — native tracked-changes markup (`w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange`); satisfies the WmlComparer contract (accept ≡ right, reject ≡ left)
- `GetRevisions(left, right, settings?)` → `IReadOnlyList<DocxDiffRevision>` — consumer revisions rendered off the edit script
- `GetEditScriptJson(left, right, settings?)` → `string` — the edit script as data (the differentiator vs `WmlComparer`)
- Accept/reject primitive (the round-trip verifier, surfaced for clients via `DocxDiffOps.AcceptRevisions`/`RejectRevisions` → WASM `DocxDiffBridge.AcceptRevisions`/`RejectRevisions` + npm `docxDiffAcceptRevisions`/`docxDiffRejectRevisions`, stdio `docx_diff_accept_revisions`/`docx_diff_reject_revisions` + docx-scalpel `docx_diff_accept_revisions`/`docx_diff_reject_revisions`): byte→byte materialize the right (accept) / left (reject) side of a redline, so a client can prove `accept(Compare(left,right))` ≡ `right` and `reject` ≡ `left` at the per-block text level — wraps `RevisionProcessor`
- N-way composite / consolidate (closes the last WmlComparer gap): `Consolidate(base, reviewers, settings?)` → `WmlDocument`, plus `GetConsolidatedRevisions`/`GetConsolidatedEditScriptJson`/`GetConflicts` — merge N `DocxDiffReviewer{Document,Author}` (each diffed against ONE shared base) into one multi-author tracked-changes document with per-reviewer attribution, token-granular sub-block merge, and a structured `DocxDiffConflict` report; `DocxDiffConsolidateSettings.ConflictResolution` = `BaseWins`(default)/`FirstReviewerWins`/`StackAll`. Round-trip: reject ≡ base, accept ≡ the policy-resolved composite. Structurally complete: note-scope (footnote/endnote) diffs merge across reviewers (compose/consensus/conflict per block, inserted notes under fresh ids, N-reviewer-aware renumbering), table column add/remove composes with native `w:cellIns`/`w:cellDel` markup, cell-shell (`w:tcPr`) edits are visible and composable, uncontested split/merge/move/row-move ops render natively (colliding ones lower to del/ins with a recorded conflict — never a silent drop), and **reviewers' paragraph-property (`w:pPr`) changes MERGE (sub-project B1: `ComposePPr` by `IrParagraph.PPrDigest` — consensus/conflict, native `w:pPrChange` authored to the winner; table-shell/section composite merge is B2, still a ceiling)**. Wraps `Docxodus/DocxDiffConsolidate.cs` + `Docxodus/Ir/Diff/IrCompositeMerger.cs`; surfaced in WASM/npm (`docxDiffConsolidate*`)/docx-scalpel (`docx_diff_consolidate*`)
- `DocxDiffSettings` mirrors `WmlComparerSettings` defaults (two honest deviations: `Deterministic` revision dates default true; `FormatComparison` defaults `ModeledOnly`). `DocxDiffRevision` adds `LeftAnchor`/`RightAnchor` (`kind:scope:unid`, interoperable with `DocxSession`/markdown projection)
- Header/footer stories are compared like Word Compare's default-on "Headers and footers" granularity (`CompareHeadersFooters`, default true) — stories pair per section ordinal × kind with Word's inheritance rule, changed stories get native markup inside their parts (accept ≡ right / reject ≡ left extends to those scopes), Fine revisions carry `hdr`/`ftr` anchors (compatible mode excludes them), JSON gains `headerFooterOps`. WmlComparer ignores headers/footers entirely; Consolidate doesn't merge them in v1 (forced off, pinned)
- Paragraph-and-above formatting changes are tracked as native markup — the block-format-change family (closes the last "Word compares Formatting, we don't" gap): `w:pPrChange` (paragraph: jc/indent/spacing/style/**numbering**, + `w:pPr/w:rPr/w:rPrChange` for a changed mark), `w:tcPrChange`/`w:trPrChange`/`w:tblPrChange`/`w:tblGridChange`/`w:tblPrExChange` (table shells — the per-element digests `IrCell.ShellDigest`/`IrRow.TrPrShellDigest`/`TrPrExDigest`/`IrTable.TblPrDigest`/`TblGridDigest` drive attribution; makes the #250 cell-shell edits *tracked*, not just visible), and `w:sectPrChange` (trailing section AND mid-document inline `w:pPr/w:sectPr` via `IrParagraph.InlineSectionFormat`). Detected via `FormatComparison` for paragraphs (canonical for shells/section); accept ≡ right / reject ≡ left holds at the property-byte level for every detected change. Note/header/footer-scope `w:pPrChange` works via the shared `RenderBlockOp` dispatch (no per-scope gate). `DocxDiffRevision.FormatChange.Scope` (`DocxDiffFormatChangeScope`: Run/Paragraph/TableCell/TableRow/Table/Section) — `WmlComparerCompatible` excludes non-Run scopes (oracle produces none); additive `scope` on the revisions wire. **`TrackBlockFormatChanges` is a public opt-out** (default true; wire `trackBlockFormatChanges`). **Two v1 ceilings: (a) Consolidate forced off, pinned** (`IrCompositeMerger` sets `TrackBlockFormatChanges=false` — sub-project B is the N-way merge); **(b) split/merge members don't emit pPrChange** (deliberate decline — members are new paragraphs already tracked by the pilcrow mark). Rode-along consume-side fix: `RevisionProcessor` no longer drops header/footer refs (sectPrChange) or an inline sectPr (pPrChange) on reject — CT_*Base inners exclude them (see `docs/ooxml_corner_cases.md`)
- No static state — `AuthorForRevisions` flows per call (multi-author / consolidate-compatible)
- Wraps the internal `Docxodus/Ir/Diff/` pipeline (`IrReader → IrEditScriptBuilder → IrMarkupRenderer/IrRevisionRenderer/IrEditScriptJson`); see `docs/architecture/ir_diff_engine.md`

**WmlToHtmlConverter.cs / HtmlToWmlConverter.cs** - Bidirectional DOCX ↔ HTML conversion. Key settings in `WmlToHtmlConverterSettings`:
- `RenderTrackedChanges` - Render insertions/deletions as `<ins>`/`<del>` instead of accepting them
- `RenderMoveOperations` - Distinguish move operations from regular insert/delete
- `RenderFootnotesAndEndnotes` - Include footnotes/endnotes sections in HTML output
- `RenderHeadersAndFooters` - Include document headers/footers in HTML output
- `RenderComments` - Render document comments in HTML output
- `CommentRenderMode` - How to render comments: `EndnoteStyle` (default), `Inline`, or `Margin`
- `AuthorColors` - Dictionary mapping author names to CSS colors for styling
- `StampAnchors` - Stamp `data-anchor="<unid>"` on block elements (`p`/`h1`-`h6`/`li`/`table`) so DOM blocks are addressable by the `kind:scope:unid` anchor system (powers the browser editor's incremental re-render). Default false.

See `docs/architecture/comment_rendering.md` for detailed comment rendering documentation.

**Single-block render (`HtmlConversionOps.RenderBlockHtml`)** - Renders ONE block (addressed by a `kind:scope:unid` anchor, or a bare unid) to faithful HTML, for incremental editor re-render. Overloads: `(byte[] bytes, …)` (stateless), `(DocxSession, …)` / `(int handle, …)` (session-attached — resolves against the live document with no byte re-open / whole-doc Unid pass, ~2.5× faster). Builds a throwaway document copying the source's styles/numbering/theme/font/settings parts. The full-document render is the faithfulness oracle. Surfaced in WASM/npm (`renderBlockHtml`, `DocxSession.renderBlock`). See `docs/architecture/ir_editor_feasibility.md`.

**DocxEditor (npm, `npm/src/editor.ts`)** - Framework-agnostic, pure-TypeScript in-browser block editor (the write-side editor counterpart to the read-side projection). Renders a faithful document with `data-anchor` blocks, makes projection-addressable paragraphs/headings `contenteditable`, and on commit edits via `DocxSession` then re-renders only the changed block. `{ paginated: true }` flows blocks into real page boxes via `pagination.ts`. Lossless `save()`. Commands: `format`/`setFontSize`/`setFontFamily`/`setAlignment`/`indent`/`setParagraphStyle`/`toggleList`/`pageBreakBefore` (all apply across a multi-block selection, not just the focused block), `clearParagraphBorders` (remove an HR/paragraph border), `deleteBlock` (remove the active block — inert inside a table or when it is the only editable block), structural `insertHorizontalRule(weight, style, position?)` (`position` = `"above"`|`"below"`, default below; `"above"` puts the rule before the active block — e.g. the S-1 top bar)/`insertTable`, table editing `insertTableRow("above"|"below")`/`insertTableColumn("left"|"right")`/`deleteTableRow`/`deleteTableColumn`, `undo`/`redo`, and the `DocxEditor.openBlank(container, exports, options?)` "New document" factory. `setFontSize`/`setFontFamily` cache the last real selection so a focus-stealing combobox/dropdown still applies to a sub-range. `insertTable` on an empty paragraph inserts the table *before* it so that line becomes the editable paragraph below the table (no stray line above). Enter inside a table cell splits the cell paragraph in place (stacked lines); Enter inside an empty horizontal rule does NOT inherit the rule's border; Shift+Enter inserts a real line break (`w:br`). The demo (`examples/editor.html`) ships a visual table grid picker (with an L/C/R cell-alignment selector), a floating table toolbar, an editable font-size combobox, a curated font-family dropdown, an Above/Below rule-position toggle, double-rule and clear-rule buttons, and a delete-block button. The model-of-record is the live OOXML in `DocxSession`; the IR/anchor system is the addressing overlay (the IR itself is read-only, no IR→OOXML writer). See `docs/architecture/ir_editor_feasibility.md`; the S-1 cover-page feature build is `docs/architecture/s1_smoke_test_features.md`.

**DocumentAssembler.cs** - Template population from XML data using content controls.

**PresentationBuilder.cs** - Merge/split PPTX files.

**SpreadsheetWriter.cs** - Simplified XLSX creation API with streaming support for large files.

**OpenXmlRegex.cs** - Search/replace in DOCX/PPTX using regular expressions.

**RevisionAccepter.cs / RevisionProcessor.cs** - Handle tracked revisions.

**FormattingAssembler.cs** - Resolve and flatten document formatting.

**MetricsGetter.cs** - Extract document metrics (styles, fonts, languages).

**OpenContractExporter.cs** - Export documents to OpenContracts format for interoperability:
- `Export(WmlDocument)` / `Export(WordprocessingDocument)` - Export to `OpenContractDocExport`
- Complete text extraction (paragraphs, tables, headers, footers, footnotes, endnotes)
- PAWLS-format page layout with token positions
- Structural annotations (sections, paragraphs, tables) with relationships
- See `docs/architecture/opencontracts_export.md` for detailed documentation

**WmlToMarkdownConverter.cs** - Anchor-addressed markdown projection of a Word document. A stable text view of a DOCX with stable IDs, suitable for LLM editing pipelines, structured search indexers, and diff/review UIs:
- `Convert(WmlDocument, WmlToMarkdownConverterSettings)` / `Convert(WordprocessingDocument, ...)` - returns `MarkdownProjection` (markdown text + anchor index)
- Anchors have the form `{#kind:scope:unid}` (e.g. `{#p:body:a1b2c3d4}`), derived from Docxodus' existing Unid system
- See `docs/architecture/markdown_projection.md` for the projection spec

**DocxSession.cs** - Stateful in-memory DOCX editing API keyed by markdown-projection anchor ids. The write-side counterpart to `WmlToMarkdownConverter` for agentic editing pipelines:
- `new DocxSession(byte[] bytes, DocxSessionSettings? settings = null)` - open a session over in-memory DOCX bytes
- Tier A (text CRUD): `ReplaceText(anchor, markdown)`, `DeleteBlock(anchor)`
- Tier B (structural): `InsertParagraph(anchor, Position, markdown)`, `SplitParagraph(anchor, offset)` (splits correctly inside a table cell too — the new `w:p` stays in the `w:tc`; splitting an EMPTY bordered paragraph does NOT copy its `w:pBdr` to the new paragraph, so Enter on a horizontal rule doesn't stack rules), `MergeParagraphs(first, second)`, `InsertHorizontalRule(anchor, Position, ParagraphBorderEdge?)` (empty bottom-bordered paragraph; `Style` supports `single`/`double`/`thick`), `InsertTable(anchor, Position, rows, cols, TableInsertOptions?)` (borderless option, row-major `CellContents`, `CellAlignment`, per-column `ColumnWidths` (twips); returns created cell anchors; always keeps a trailing `w:p` after the table so an end-of-body table has an editable paragraph below it), and post-insert table editing addressed by a cell-paragraph anchor: `InsertTableRow(cellAnchor, Position)`, `InsertTableColumn(cellAnchor, Position)`, `DeleteTableRow(cellAnchor)`, `DeleteTableColumn(cellAnchor)` (deleting the last row/column removes the table; v1 assumes a rectangular grid, no `w:gridSpan`). Intra-paragraph newlines in the markdown subset (the GFM hard break `"  \n"`) round-trip as a real `w:br`.
- Tier C (formatting): `ApplyFormat(anchor, CharSpan?, FormatOp)` (`FormatOp.FontSizePts` → `w:sz`/`w:szCs`; `FormatOp.FontFamily` → `w:rFonts` ascii/hAnsi/cs, `""` clears), `SetParagraphStyle(anchor, styleId)`, `SetParagraphFormat(anchor, ParagraphFormatOp)` (alignment/indent/page-break + `TopBorder`/`BottomBorder`/`ClearBorders` → `w:pBdr`), `SetListLevel(anchor, delta)`, `RemoveListMembership(anchor)`
- Tier D (advanced): `ReplaceCellContent(cellAnchor, markdown)`; `Settings.TrackedChanges = RenderInline` makes all mutations land as `w:ins`/`w:del`
- Factory: `DocxSession.CreateBlankDocxBytes()` (static) — mint a complete blank DOCX ("New document" seed: Normal style, US-Letter section); WASM/npm `createBlankDocx()` + `DocxEditor.openBlank(...)`
- Tier E (annotations): `AddAnnotation(anchorId, span, DocumentAnnotation)`,
  `RemoveAnnotation(id)`, `UpdateAnnotation(id, AnnotationUpdate)`,
  `MoveAnnotation(id, newAnchorId, newSpan)` — anchor-addressed annotation
  CRUD that mutates the live session document. `EditResult.AnnotationId`
  carries the affected id on success.
- Inspection: `GetBlockMetadata(anchor)`, `GetBlockMetadatas(anchors)`,
  `GetListMembership(anchor)`, `GetSectionInfo(anchor)` — read-only
  block-level metadata (style id/name, outline level, list facts:
  numId/abstractNumId/ilvl/format/start-override/from-style,
  sectPr page setup). Returns null for unknown anchors.
- Raw OOXML escape hatch: `session.Raw.GetXml(anchor)`, `Raw.InsertXml(anchor, Position, xml)`, `Raw.ReplaceXml(anchor, xml)` for content the markdown subset can't express
- Bounded snapshot `Undo()`/`Redo()` (configurable depth via `Settings.UndoDepth`)
- Every mutation returns a typed `EditResult` envelope: `Success`, `EditError(EditErrorCode, message, anchorId)`, `Created`/`Removed`/`Modified` anchor lists, and a `MarkdownPatch` for the affected scope
- Available in .NET, WASM (`DocxSessionBridge`), and npm TypeScript (`openDocxSession`, `DocxSession`)
- See `docs/architecture/docx_mutation_api.md` for the full surface contract, anchor lifecycle table, error catalog, and supported markdown subset

**ExternalAnnotationProjector.cs** - Incremental annotation overlay API (Issue #106). Decouples annotation projection from DOCX conversion for dramatically better performance when annotations change:
- `ProjectAnnotationsOntoHtml(html, set, settings)` - Project a full annotation set onto pre-converted HTML (~56ms vs ~892ms for full re-conversion, 15.9x faster)
- `AddAnnotationToHtml(html, annotation, label, settings)` - Add a single annotation (~0.3ms, 2972x faster than full re-conversion)
- `RemoveAnnotationFromHtml(html, annotationId, cssPrefix)` - Remove a single annotation by ID (~18ms)
- `GenerateVisibilityCss(hiddenLabelIds, cssPrefix)` - Generate CSS to hide/show annotations by label (instant toggling)
- `GenerateAnnotationCssString(labels, settings)` - Generate annotation CSS independently
- Works by building a text map of the HTML, finding annotation text via string search, and wrapping matches with styled `<span>` elements
- `GetTextNodes` skips already-projected annotation wrappers to prevent offset drift from label text
- Available in .NET, WASM (JSExport), and npm TypeScript wrapper
- See `docs/architecture/incremental_annotation_overlay.md` for detailed documentation

### Target Frameworks

Library targets: `net10.0`
Tests target: `net10.0`

### Dependencies

- **DocumentFormat.OpenXml**: 3.4.1 (Open XML SDK)
- **SkiaSharp**: 2.88.9 (cross-platform graphics, replaces System.Drawing)

### Test Data

Test files are in `TestFiles/` directory with prefixes indicating their purpose:
- `DB*` - DocumentBuilder tests
- `DA*` - DocumentAssembler tests
- `HC*` - HTML Converter tests
- `WC/` - WmlComparer tests
- `SH*` - Spreadsheet tests
- `CU*` - Chart Updater tests

## Legacy Migration Notes

Docxodus is a fork of OpenXmlPowerTools, upgraded from net45/net46/netstandard2.0 → .NET 8.0 → .NET 10.0 and from Open XML SDK 2.8.1 → 3.x. A few artifacts of that migration are worth knowing when reading code:

- **`GetPackage()` extension in `PtOpenXmlUtil.cs`** — Open XML SDK 3.x made the internal `Package` private; we access it via reflection. Use this extension rather than reaching for `OpenXmlPackage.Package` directly.
- **`PartTypeInfo` pattern** — replaces SDK 2.x's `FontPartType`/`ImagePartType` enums when adding parts.
- **`Dispose()` not `.Close()`** — SDK 3.x dropped `Close()`; always use `using` blocks or `Dispose()`.
- **SkiaSharp replaces System.Drawing** — `SKColor`/`SKBitmap`/`SKTypeface`/`SKEncodedImageFormat`. Helpers in `SkiaSharpHelpers.cs` (notably `ColorHelper` for color name mapping). Remember the WASM build excludes SkiaSharp entirely — see WASM Conditional Compilation above.
- **Rebranded namespaces** — everything is `Docxodus`; old `OpenXmlPowerTools*` types are `Docxodus*` (e.g. `DocxodusDocument`, `DocxodusException`). Legacy example projects live in `archived-examples/` (not in the solution).
- **Preprocessor cleanup pending** — `NET35` and `ELIDE_XUNIT_TESTS` directives still appear in some files; safe to remove when you touch a file (Phase 4 of the migration plan).

For specific bugfix history (e.g. relationship copying in `DocumentBuilder`, footnote/endnote Unid assignment, LCS-based table row matching), use `git log` rather than maintaining a list here.

## Architecture Documentation

Detailed design docs for the major subsystems live in `docs/architecture/`. Read the relevant doc before making non-trivial changes to:

- `comparison_engine.md`, `wml_comparer_gaps.md`, `native_move_markup.md`, `move_detection_implementation_plan.md`, `format_change_detection.md`, `tracked_changes.md` — WmlComparer internals
- `docx_converter.md`, `comment_rendering.md`, `paginated_headers_footers.md`, `custom_annotations.md`, `unsupported_content_placeholders.md`, `wml_to_html_converter_gaps.md` — WmlToHtmlConverter internals
- `opencontracts_export.md` — OpenContractExporter format
- `markdown_projection.md` — WmlToMarkdownConverter design
- `ir_diff_engine.md` — DocxDiff (IR diff engine) public surface, pipeline, edit script, settings, parity status, relationship to WmlComparer
- `docx_mutation_api.md` — DocxSession surface, anchor lifecycle, error catalog, supported markdown subset
- `ir_editor_feasibility.md` — IR-powered browser DOCX editor: architecture (Option B — DocxSession is model-of-record, IR/anchors are addressing), RenderBlockHtml + DocxEditor surface, measured results, findings
- `ir_editor_roadmap.md` — sequenced, impact-ordered roadmap for the editor (M1 rich in-block editing → M9 render fidelity); architecture invariants to preserve
- `python_docxodus.md` — planned Python wrapper for DocxSession; wire protocol, type mapping, distribution
- `skiasharp-removal-plan.md`, `wasm-optimization-plan.md`, `ui_responsiveness.md`, `profiling-results.md` — WASM/browser work

## OOXML Corner Cases

When investigating bugs where our output differs from Word/LibreOffice rendering, **always document findings** in `docs/ooxml_corner_cases.md`. This is critical because:

1. **Word doesn't always follow the spec** - Microsoft Word sometimes implements undocumented behavior or interprets ambiguous spec sections differently than expected
2. **Future reference** - These edge cases are hard to rediscover; documenting them saves hours of debugging later
3. **Test coverage** - Each documented case should eventually have a corresponding test

### What to Document

- Any case where Word renders differently than a literal reading of the OOXML spec would suggest
- Behaviors that differ between Word, LibreOffice, and our implementation
- Numbering/list formatting edge cases (especially legal numbering, multi-level formats)
- Style inheritance quirks
- Table layout anomalies
- Character/paragraph property interactions

### Documentation Format

For each corner case, include:
1. **Minimal XML reproducer** - The smallest XML snippet that demonstrates the issue
2. **Renderer comparison table** - What Word, LibreOffice, and Docxodus each produce
3. **Analysis** - Your hypothesis about why the difference exists
4. **Relevant code** - Which Docxodus files/functions are involved
5. **Proposed fix** - If known, how to align with Word's behavior

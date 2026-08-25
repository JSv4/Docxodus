# CLAUDE.md

Guidance for Claude Code (claude.ai/code) working in this repository.

**This file is a map, not a manual.** It covers what you cannot discover quickly from the
code: repository layout, which surfaces a change must ripple through, build gotchas, and
where the real documentation lives. Per-module design detail belongs in
`docs/architecture/` — see the index below. Do not re-document a subsystem here; a
duplicated description is one that will go stale.

## Important Instructions

- **Never credit yourself in commits.** Do not add "Generated with Claude Code" or
  "Co-Authored-By: Claude" to commit messages.
- **Write PR descriptions for a human reviewer, not for the next agent.** Explain, in plain
  language, why the change exists, how it works at a mechanism level, and how it was
  validated. Don't invent terminology, lean on internal issue-number shorthand, or use
  jargon a reader outside the change wouldn't recognize — a reviewer who hasn't read the
  code should be able to follow the description on its own.

## Coding Standards

### Nullable Reference Types

The project sets `<Nullable>disable</Nullable>` globally: enabling it today produces
**~4,800 warnings** in the legacy core. New code should still be annotated.

- **New files**: add `#nullable enable` at the top.
- **Substantial refactors**: consider adding `#nullable enable` and fixing that file's warnings.
- **Use proper annotations**: mark nullable parameters/returns with `?`; use null checks or `!`.

About 60% of files in `Docxodus/` already carry `#nullable enable` — everything written for
this fork. The un-annotated remainder is the inherited OpenXmlPowerTools core
(`WmlComparer`, `WmlToHtmlConverter`, the `HtmlToWml*` family, `FormattingAssembler`,
`DocumentBuilder`, `RevisionProcessor`, `PtOpenXmlUtil`, `PtUtil`, `ListItemRetriever`).

See [Issue #13](https://github.com/JSv4/Docxodus/issues/13) for the migration plan.

### Warnings

`Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release — but
**`Docxodus.csproj` and `Docxodus.Tests.csproj` both override it to `false`**, so the core
library and the test project do *not* fail on warnings. The CLI tools, MCP server,
python-host and WASM project do inherit it. Current baseline: the library builds with
**115 warnings**, the test project with **~618** (mostly StyleCop `SA1633` file headers and
xUnit analyzer suggestions). Don't add to either baseline.

## Repository Layout

This is not just a .NET library — it ships a multi-layer stack. Public-surface changes
usually ripple through all of it.

| Layer | Path | Purpose |
|-------|------|---------|
| Core library | `Docxodus/` | All OOXML logic. NuGet package `Docxodus`. |
| Shared facades | `Docxodus/Internal/{DocxSessionOps,DocxDiffOps,HtmlConversionOps,SessionRegistry,DocxSessionJson}.cs` | Single-owner op + wire-shape layer every transport routes through. |
| Unit tests | `Docxodus.Tests/` | xUnit, **~3,440 tests**, ~4 min. |
| CLI tools | `tools/redline/`, `tools/docx2html/`, `tools/docx2oc/` | Thin `dotnet tool` wrappers. |
| WASM bridge | `wasm/DocxodusWasm/` | `[JSExport]` shells over the facades. |
| Stdio host | `tools/python-host/` | NDJSON-over-stdin host (`docxodus-pyhost`) that the `docx-scalpel` pip package subprocesses. |
| Agent server | `tools/mcp-server/` | JSON-RPC 2.0 / MCP stdio server (`docxodus-mcp`): 3 lifecycle tools + 12 grouped-intent tools. See `docs/architecture/docx_agent_server.md`. |
| Python client | `python/` | `docx-scalpel` on PyPI. |
| npm/TypeScript | `npm/` | Browser package: `src/index.ts` (API), `src/editor.ts` (block editor), `src/ribbon.ts` (shipped UI surface), `src/embed.ts` (CDN entries), `src/react.ts`, worker proxy. |
| Workflow evals | `eval/` + `Docxodus.Tests/Eval/` | Deterministic document-workflow scenarios scored on task completion, target precision, and collateral change. Corpus and contract in `eval/README.md`. |
| Pages demo | `docs/demo/` | Static pages hosting the **shipped** surface via `createRibbonEditor`. They contain no editor UI of their own — change `npm/src/ribbon.ts`, not these files. |

### The single-owner rule

When a public `DocxSession` method or setting changes, **update `Docxodus/Internal/DocxSessionOps.cs`
first** — both bridges and both clients pick the change up from there. Then ripple outward.

The same pattern governs the stateless surfaces: `HtmlConversionOps` owns DOCX→HTML,
`DocxDiffOps` owns the `DocxDiff` engine, and `DocxCompare` owns the
`WmlComparer`-vs-`DocxDiff` engine selection. Change the facade first, then the two bridges
and two clients.

### Ripple checklist

| Change type | Core | Tests | `*Ops` facade | WASM | npm/TS | stdio + Python | MCP | Docs |
|-------------|------|-------|---------------|------|--------|----------------|-----|------|
| New setting or method | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| New public enum | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| Bug fix | ✓ | ✓ | – | – | – | – | – | CHANGELOG |
| Internal refactor | ✓ | ✓ | – | – | – | – | – | – |

Concretely, a new session op touches: `DocxSession.cs` → `DocxSessionOps.cs` →
`DocxSessionJson.cs` → `wasm/DocxodusWasm/DocxSessionBridge.cs` → `npm/src/types.ts` +
`npm/src/index.ts` → `tools/python-host/Dispatcher.cs` → `python/src/docx_scalpel/{types,session}.py`
→ `tools/mcp-server/{ToolCatalog,Dispatcher}.cs`.

**The generated-PDF benchmark is a ripple site for PageMap shape.**
`npm/tests/visual-parity/pdf-result.ts` validates the supported `convertDocxToPdf` envelope with
*exact* key matching on `pageMap`, each page, and each fragment, so adding an optional field to any
of them fails the benchmark from a test file none of the tables above name. Update it in the same
change.

**Defaults are declared once.** Read them from the settings object rather than repeating
literals per surface — `DocxSessionJson.ParseSettings` and the MCP dispatcher had each
hardcoded their own copy of `undoDepth`, and it drifted.

## Build Commands

```bash
dotnet build Docxodus.sln                       # needs the wasm-tools workload
dotnet build Docxodus/Docxodus.csproj           # library only, no workload needed
dotnet build -c Release Docxodus.sln
./scripts/build-wasm.sh                         # WASM target (sets WASM_BUILD, excludes SkiaSharp)
cd npm && npm run build                         # build-wasm.sh + tsc + esbuild bundles
```

A full-solution build requires `dotnet workload install wasm-tools`. Without it, build
`Docxodus/Docxodus.csproj` and `Docxodus.Tests/Docxodus.Tests.csproj` directly.

## Test Commands

```bash
dotnet test Docxodus.Tests/Docxodus.Tests.csproj
dotnet test --filter "FullyQualifiedName~DB001_DocumentBuilderKeepSections"   # test ids are feature-prefixed
dotnet test --filter "FullyQualifiedName~DbTests"

cd npm
npm install && npx playwright install chromium   # first time only
npm run build                                    # produces dist/ that the harness loads
npm test                                         # all Playwright specs
npx playwright test --grep "Document Structure"
npx tsc --noEmit
```

Playwright serves from `npm/dist/wasm/`. **If you edit C#, `.ts`, or the harness HTML,
re-run `npm run build` first** or you will test stale artifacts.

**Running the .NET suite dirties the working tree.** The `SH*` spreadsheet tests write back
to their own committed fixtures under `TestFiles/`, leaving ~47 modified `.xlsx` files in
`git status`. They are test-run noise, not your changes — `git checkout -- TestFiles/` before
committing.

## WASM Conditional Compilation

The library compiles in two modes, controlled by the `WASM_BUILD` MSBuild property
(set by `scripts/build-wasm.sh`):

- **Default**: includes `SkiaSharp` + `SkiaSharp.NativeAssets.Linux.NoDependencies`.
- **`WASM_BUILD=true`**: defines the `WASM_BUILD` constant and excludes SkiaSharp (no native
  deps in the browser). Code needing SkiaSharp must be guarded with `#if !WASM_BUILD` or routed
  through a no-op fallback.

When touching image/font/color code, check that it compiles under `WASM_BUILD` — the npm
build fails loudly if it doesn't.

**Output isolation:** the WASM-mode assembly builds into `Docxodus/bin/wasm/` +
`Docxodus/obj/wasm/` (see the `WASM_BUILD` PropertyGroup in `Docxodus.csproj`) so it can never
clobber the default-mode assembly — a solution build compiles Docxodus twice. If you ever see
`error CS1061: 'ImageInfo' does not contain a definition for 'SaveImage'`, a stale
pre-isolation artifact is lingering: delete `Docxodus*/bin` + `Docxodus*/obj` once.

## Core Concepts

**Document wrappers.** `DocxodusDocument` (base, holds `DocumentByteArray` + `FileName`) with
`WmlDocument` / `SmlDocument` / `PmlDocument` for Word / Excel / PowerPoint. Immutable-style
manipulation goes through `OpenXmlMemoryStreamDocument`:

```csharp
using (var streamDoc = new OpenXmlMemoryStreamDocument(doc))
{
    using (var document = streamDoc.GetWordprocessingDocument()) { /* modify */ }
    return streamDoc.GetModifiedWmlDocument();
}
```

**Anchors.** `kind:scope:unid` (e.g. `p:body:a1b2c3d4`) is the addressing system shared by the
markdown projection, `DocxSession`, `DocxDiff` revisions, and the editor's `data-anchor`
attributes. It is an addressing overlay over the live OOXML, which remains the model of record
— there is no IR→OOXML writer.

**Comparison engines.** `DocxDiff` is the default since v8.0.0; `WmlComparer` is the older
engine, feature-frozen and available by explicit selection. Note that `DocxDiff` still depends
on `WmlComparer`'s *types* (`WmlComparerSettings`, `WmlComparerRevision`), so the older engine
is not removable.

## Module Map

Read the linked doc before non-trivial changes. These docs are the source of truth for design
detail; this file deliberately does not restate them.

| Module | What it does | Design doc |
|--------|--------------|------------|
| `DocxSession.cs` | Stateful anchor-addressed editing API (text, structural, formatting, tables, notes, comments, revisions, annotations, raw XML, undo/redo) | `docx_mutation_api.md` |
| `DocxDiff.cs` + `Ir/Diff/` | Structure-aware comparison → native tracked changes; edit script as data; N-way consolidate | `ir_diff_engine.md` |
| `WmlComparer.cs` | Legacy comparison engine (frozen) | `comparison_engine.md`, `wml_comparer_gaps.md`, `native_move_markup.md`, `format_change_detection.md` |
| `WmlToHtmlConverter.cs` | DOCX → HTML, the render fidelity oracle | `docx_converter.md`, `comment_rendering.md`, `paginated_headers_footers.md`, `wml_to_html_converter_gaps.md` |
| `HtmlToWmlConverter.cs` | HTML → DOCX | — |
| `WmlToMarkdownConverter.cs` | Anchor-addressed markdown projection | `markdown_projection.md` |
| `npm/src/editor.ts` + `ribbon.ts` | In-browser block editor and its shipped UI surface | `ir_editor_feasibility.md`, `ir_editor_roadmap.md`, `editor_ui_surface.md` |
| `ExternalAnnotationProjector.cs` | Incremental annotation overlay on pre-converted HTML | `incremental_annotation_overlay.md`, `custom_annotations.md` |
| `OpenContractExporter.cs` | Export to OpenContracts format | `opencontracts_export.md` |
| `DocumentBuilder.cs` | Merge / split DOCX | — |
| `DocumentAssembler.cs` | Template population from XML via content controls | — |
| `PresentationBuilder.cs` | Merge / split PPTX | — |
| `SpreadsheetWriter.cs` | Simplified XLSX creation, streaming | — |
| `OpenXmlRegex.cs` | Regex search/replace in DOCX/PPTX | — |
| `RevisionProcessor.cs` | Accept/reject tracked revisions | `tracked_changes.md` |
| `FormattingAssembler.cs` | Resolve and flatten formatting | — |
| `MetricsGetter.cs` | Document metrics (styles, fonts, languages) | — |

WASM/browser work: `wasm-packaging.md` (trimming, Brotli, size budget, measured payload),
`wasm-optimization-plan.md`, `skiasharp-removal-plan.md`, `ui_responsiveness.md`,
`profiling-results.md`. Python wrapper: `python_docxodus.md`.

### Coverage gaps worth knowing

The DOCX path is exhaustively tested. The inherited spreadsheet/presentation modules are not:
`WorksheetAccessor` and `TextReplacer` have **zero** references in `Docxodus.Tests/`, and
`PresentationBuilder`, `SpreadsheetWriter`, `ChartUpdater` and `SmlToHtmlConverter` have one
test file each. Treat changes there as unguarded.

## Feature Development Workflow

1. **CHANGELOG.md** — add an entry under `[Unreleased]`.
2. **Tests** — add to the matching file in `Docxodus.Tests/`. Reuse `TestFiles/` fixtures where
   possible; programmatic fixtures need all required parts (`StyleDefinitionsPart`,
   `DocumentSettingsPart`, …).
3. **Ripple** — follow the checklist above.
4. **Docs** — update `docs/architecture/` for significant features, and
   `docs/ooxml_corner_cases.md` for any Word-behaviour discovery (see below).

## OOXML Corner Cases

When our output differs from Word or LibreOffice, **document the finding** in
`docs/ooxml_corner_cases.md`. Word does not always follow the spec, these cases are expensive
to rediscover, and each one should eventually get a test.

Record: a minimal XML reproducer, a Word / LibreOffice / Docxodus comparison table, your
analysis of why they differ, the Docxodus code involved, and the fix if known.

## Release Process

A release is **a CHANGELOG section + an annotated tag + a GitHub Release**. Versions are
injected from the tag at publish time — `Docxodus.csproj` stays at `1.0.0` and
`npm/package.json` at `0.0.0` deliberately; do not bump them.

Semver on `vMAJOR.MINOR.PATCH` tags, chosen from what accumulated in `[Unreleased]`:

| Bump | When |
|------|------|
| Patch | only `### Fixed` |
| Minor | any `### Added` / `### Changed`, no breaking change |
| Major | a breaking public-API change |

From an up-to-date `main`:

1. In `CHANGELOG.md`, insert `## [X.Y.Z] - YYYY-MM-DD` under `## [Unreleased]`, leaving the
   accumulated entries beneath it and `[Unreleased]` empty above.
2. Commit changelog-only: `docs(changelog): cut vX.Y.Z release notes`.
3. Annotated tag whose message is the version: `git tag -a vX.Y.Z -m vX.Y.Z`.
4. `git push origin main && git push origin vX.Y.Z`.
5. `gh release create vX.Y.Z --title vX.Y.Z --notes-file <body.md> --latest --verify-tag`.
   Every tag back to `v5.x` has a Release; a tag without one is an incomplete release.

Release body opens with a one-line lead linking the CHANGELOG anchor
(`…/CHANGELOG.md#XYZ---YYYY-MM-DD`, digits only). Patch/minor: the lead plus the
`### Added`/`### Changed`/`### Fixed` sections verbatim (see `v7.1.0`). Major: the lead plus
`### Highlights` and `### Breaking changes` as a *summary* — a major's accumulated entries are
far too long to dump (see `v7.0.0`, `v8.0.0`). `### Breaking changes` must say what silently
changes for a caller who passes nothing, and how to pin the old behaviour.

Reference commits: `#206`, `#209`; reference tags: `v6.1.0`, `v6.2.0`.

## Dependencies

- **DocumentFormat.OpenXml** 3.5.1
- **SkiaSharp** 4.148.0 (+ `SkiaSharp.NativeAssets.Linux.NoDependencies`), excluded under `WASM_BUILD`

Target framework: `net10.0` for both library and tests.

## Legacy Migration Notes

Docxodus forks OpenXmlPowerTools, upgraded net45/net46/netstandard2.0 → .NET 10 and Open XML
SDK 2.8.1 → 3.x. Artifacts of that migration worth knowing:

- **`GetPackage()` in `PtOpenXmlUtil.cs`** — SDK 3.x made the internal `Package` private; we
  reach it by reflection. Use this extension, not `OpenXmlPackage.Package`.
- **`PartTypeInfo`** replaces SDK 2.x's `FontPartType`/`ImagePartType` enums when adding parts.
- **`Dispose()`, not `.Close()`** — SDK 3.x dropped `Close()`.
- **SkiaSharp replaces System.Drawing** — `SKColor`/`SKBitmap`/`SKTypeface`/`SKEncodedImageFormat`,
  helpers in `SkiaSharpHelpers.cs` (notably `ColorHelper`). Remember the WASM build excludes it.
- **Preprocessor cleanup pending** — `NET35` and `ELIDE_XUNIT_TESTS` directives remain in some
  files; safe to remove when you touch one.
- Legacy example projects live in `archived-examples/` and are not in the solution.

For bugfix history, use `git log` rather than maintaining a list here.

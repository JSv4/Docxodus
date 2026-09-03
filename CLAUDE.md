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

`Docxodus.csproj` sets `<Nullable>enable</Nullable>` (issue #13): every file is
nullable-checked by default, so a new file needs no directive. The inherited
OpenXmlPowerTools core has fully migrated (issue #645): no file under `Docxodus/`
carries a `#nullable disable` header anymore. `grep -l "^#nullable disable" Docxodus/*.cs`
should return nothing — reintroducing the header on a new or refactored file is a
regression, not a shortcut. `Docxodus.csproj` no longer has a `NoWarn` property at all
(issue #651) — `CS8632`, `CS8073` and `CA2200` are gone along with the last opted-out
file that could fire them.

### Warnings

`Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release — but
**`Docxodus.csproj` and `Docxodus.Tests.csproj` both override it to `false`**, so the core
library and the test project do *not* fail on warnings. The CLI tools, MCP server,
python-host and WASM project do inherit it. Current baseline: the library builds with
**113 warnings**, the test project with **707** (mostly StyleCop `SA1633`/`SA1636` file
headers and `SA1206` using-order). Don't add to either baseline. Measure with
`--no-incremental` — a warm incremental build reports zero because nothing recompiles.

The one movement that is not a regression: no file in the library carries a StyleCop file
header, so `SA1633` fires once per file and **every new `.cs` file under `Docxodus/` moves
both counts by exactly one** (the test build compiles the library through its project
reference). A change that adds one file and one warning is at baseline; anything more is not.

Update the two numbers here in the same commit rather than leaving them stale.

## Repository Layout

This is not just a .NET library — it ships a multi-layer stack. Public-surface changes
usually ripple through all of it.

| Layer | Path | Purpose |
|-------|------|---------|
| Core library | `Docxodus/` | All WordprocessingML logic. NuGet package `Docxodus`. |
| Shared facades | `Docxodus/Internal/{DocxSessionOps,DocxDiffOps,HtmlConversionOps,SessionRegistry,DocxSessionJson}.cs` | Single-owner op + wire-shape layer every transport routes through. |
| Delivery | `Docxodus/Delivery/` | Delivery bundles: manifest, publisher, revision policy, host renderer. See `delivery_bundle.md`. |
| Verification | `Docxodus/Verification/` | Package manifests, semantic diff, redline-reversibility proof, deliverable inspection. See `deliverable_verification.md`, `package_manifests.md`, `semantic_diff.md`, `redline_reversibility_proof.md`. |
| Unit tests | `Docxodus.Tests/` | xUnit, **~3,900 tests**, ~4 min. |
| CLI tools | `tools/redline/`, `tools/docx2html/`, `tools/docx2oc/` | Thin `dotnet tool` wrappers. |
| WASM bridge | `wasm/DocxodusWasm/` | `[JSExport]` shells over the facades. |
| Stdio host | `tools/python-host/` | NDJSON-over-stdin host (`docxodus-pyhost`) that the `docx-scalpel` pip package subprocesses. |
| Agent server | `tools/mcp-server/` | JSON-RPC 2.0 / MCP stdio server (`docxodus-mcp`): lifecycle, grouped-intent, and sessionless tools. See `docs/architecture/docx_agent_server.md`. |
| Python client | `python/` | `docx-scalpel` on PyPI. |
| npm/TypeScript | `npm/` | Browser package: `src/index.ts` (API), `src/editor.ts` (block editor), `src/ribbon.ts` (shipped UI surface), `src/embed.ts` (CDN entries), `src/react.ts`, worker proxy. |
| Export package | `npm-export/` | `@docxodus/export` — deterministic standalone HTML + PDF through a pinned Chromium. Has its own `dist/`, tests, and `docxodus doctor` preflight. See `standalone_paginated_export.md`. |
| Workflow evals | `eval/` + `Docxodus.Tests/Eval/` | Deterministic document-workflow scenarios scored on task completion, target precision, and collateral change. Corpus and contract in `eval/README.md`. |
| Pages demo | `docs/demo/` | Static pages hosting the **shipped** surface via `createRibbonEditor`. They contain no editor UI of their own — change `npm/src/ribbon.ts`, not these files. |
| Out-of-solution tools | `tools/diffharness/`, `tools/manifest-fuzz/`, `tools/screenshots/`, `benchmarks/` | Not in `Docxodus.sln`; build them by path. LibreOffice-backed diff verification, AFL manifest fuzzing, README screenshot capture, a form-document edit benchmark, and a `DocxDiff` perf/output-parity stress harness. |

### The single-owner rule

When a public `DocxSession` method or setting changes, **update `Docxodus/Internal/DocxSessionOps.cs`
first** — both bridges and both clients pick the change up from there. Then ripple outward.

The same pattern governs the stateless surfaces: `HtmlConversionOps` owns DOCX→HTML,
`DocxDiffOps` owns the `DocxDiff` engine, and `DocxCompare` owns the
the shared comparison front door (its input-revision policy). Change the facade first, then the two bridges
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
./scripts/record-aot-profile.sh                 # re-record wasm/DocxodusWasm/docxodus.aotprofile (~4 min)
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

The .NET suite no longer dirties `TestFiles/` — the `SH*` spreadsheet tests that used to
write back to their own committed `.xlsx` fixtures were removed with the SpreadsheetML
modules. If `git status` shows modified fixtures after a run, a test is writing where it
should be using a temp copy; fix the test rather than reflexively `git checkout -- TestFiles/`.

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
`WmlDocument` for Word. Immutable-style manipulation goes through
`OpenXmlMemoryStreamDocument`:

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

**Comparison engine.** `DocxDiff` is the only comparison engine. `WmlComparer` was removed in
v11.0.0; its cross-package merge helpers survive as `PackageMerge.cs` (styles, numbering and
related-part copying — never comparison logic), which the markup renderers call. The parity
evidence gathered before removal is frozen under
`docs/architecture/wmlcomparer_parity_baseline/`, and `DocxDiffCorpusBaselineTests` is the live
regression net that replaced the differential harness.

**`DocxCompare` vs `DocxDiff`.** `DocxCompare.Compare` is the front door every transport routes
through; it always applies `PreAcceptInputRevisions` + `PreserveInputRevisions`, which the raw
`DocxDiff` API leaves opt-in. Calling `DocxDiff.Compare` with fresh settings is NOT equivalent —
on revision-bearing inputs it emits whole-document churn.

## Module Map

Read the linked doc before non-trivial changes. These docs are the source of truth for design
detail; this file deliberately does not restate them.

| Module | What it does | Design doc |
|--------|--------------|------------|
| `DocxSession.cs` | Stateful anchor-addressed editing API (text, structural, formatting, tables, notes, comments, revisions, annotations, raw XML, undo/redo) | `docx_mutation_api.md` |
| `DocxDiff.cs` + `Ir/Diff/` | Structure-aware comparison → native tracked changes; edit script as data; N-way consolidate | `ir_diff_engine.md` |
| `WmlToHtmlConverter.cs` | DOCX → HTML, the render fidelity oracle | `docx_converter.md`, `comment_rendering.md`, `paginated_headers_footers.md`, `wml_to_html_converter_gaps.md` |
| `HtmlToWmlConverter.cs` | HTML → DOCX | — |
| `WmlToMarkdownConverter.cs` | Anchor-addressed markdown projection | `markdown_projection.md` |
| `npm/src/editor.ts` + `ribbon.ts` | In-browser block editor and its shipped UI surface | `ir_editor_feasibility.md`, `ir_editor_roadmap.md`, `editor_ui_surface.md` |
| `ExternalAnnotationProjector.cs` | Incremental annotation overlay on pre-converted HTML | `incremental_annotation_overlay.md`, `custom_annotations.md` |
| `OpenContractExporter.cs` | Export to OpenContracts format | `opencontracts_export.md` |
| `DocumentBuilder.cs` | Merge / split DOCX | — |
| `OpenXmlRegex.cs` | Regex search/replace in DOCX | — |
| `RevisionProcessor.cs` | Accept/reject tracked revisions | `tracked_changes.md` |
| `FormattingAssembler.cs` | Resolve and flatten formatting | — |
| `MetricsGetter.cs` | Document metrics (styles, fonts, languages) | — |
| `PageMap.cs` + `npm/src/pagination.ts` | Page geometry and the paginated view | `page_map.md`, `paginated_headers_footers.md` |
| `Ir/` | The structure-aware IR the diff and markdown projection read | `document_ir.md` |
| `Delivery/` | Delivery bundles and change receipts | `delivery_bundle.md`, `delivery_change_receipt.md` |
| `Verification/` | Package manifests, semantic change sets, reversibility proof | `deliverable_verification.md`, `package_manifests.md`, `semantic_diff.md`, `redline_reversibility_proof.md` |
| `npm-export/` | Standalone HTML/PDF export via pinned Chromium | `standalone_paginated_export.md` |

Editor internals: `editor_block_drag_handles.md`, `editor_inline_formatting_on_edit.md`,
`native_content_controls.md`, `native_images.md`, `unsupported_content_placeholders.md`,
`browser_llm_demo.md`. Diff/comparison detail: `docxdiff_libreoffice_findings.md`,
`move_detection_implementation_plan.md`. Smoke-test contracts: `s1_smoke_test_features.md`,
`epic_435_acceptance_smoke.md`.

WASM/browser work: `wasm-packaging.md` (trimming, profile-guided AOT, Brotli, size budget,
measured payload and speed frontier),
`wasm-optimization-plan.md`, `skiasharp-removal-plan.md`, `ui_responsiveness.md`,
`profiling-results.md`. Python wrapper: `python_docxodus.md`.

### Scope: the DOCX toolchain, nothing else

Docxodus handles `.docx`/`.docm`/`.dotx`/`.dotm`, and every module in `Docxodus/` earns its
place by serving `DocxSession`, `DocxDiff`, or the render/projection paths.
Two rounds of the inherited OpenXmlPowerTools fork were removed on that rule — see the
`### Removed` entry in `CHANGELOG.md` for the full list:

- **Non-Wordprocessing formats.** `SpreadsheetWriter`, `WorksheetAccessor`,
  `SmlToHtmlConverter`, `XlsxTables`, `ChartUpdater`, `PresentationBuilder`, `TextReplacer`,
  and the `SmlDocument`/`PmlDocument` wrappers. `GetDocumentType()` still recognises XLSX and
  PPTX packages so that feeding one in throws a clear `PowerToolsDocumentException` instead of
  failing deeper in the stack.
- **DOCX code with no callers.** `WmlToXml` (DOCX → custom XML), `DocumentAssembler`
  (content-control templating), `ReferenceAdder` (TOC/TOF/TOA fields, plus the `WmlDocument`
  partial that exposed them), `PowerToolsBlock`, `PowerToolsBlockExtensions`,
  `StronglyTypedBlock`, and `OxPtHelpers`.

Don't reintroduce either category. Before adding a module here, ask which of the three
engines it serves; if the answer is "none", it belongs somewhere else.

The `TestFiles/DA*.docx` fixtures outlived `DocumentAssembler` deliberately — they are
content-control-heavy real Word documents, and the `Ir*` tests glob `TestFiles/**/*.docx` as
a corpus. Deleting an unreferenced `.docx` silently shrinks that corpus; check the globs
before you do.

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

### What the tag actually publishes

Pushing `vX.Y.Z` fires `publish.yml`, which derives the version from the tag and publishes
**four NuGet packages** (`Docxodus`, `Redline`, `Docx2Html`, `Docx2OC`), **two npm packages**
(`docxodus`, then `@docxodus/export`), and twelve self-contained CLI binaries (three tools ×
four RIDs) as workflow artifacts. Release assets stay empty — the binaries are artifacts, not
attachments, and have
been for every release; an empty asset list is not a failed run.

**`docx-scalpel` does not ship from a `vX.Y.Z` tag.** PyPI is driven by
`python-publish.yml` on a separate `docx-scalpel-v<PEP440>` tag. That decoupling is
deliberate (see the header comment in that workflow): a Python-only point release should not
drag core/npm/binaries along, and a core release should not force a PyPI bump. Cut a
`docx-scalpel-v*` tag when the wheel needs to move — not as part of every release.

**The run can partially succeed, and the order matters.** Each registry is published by a
different job. NuGet goes first and independently; `docxodus` and `@docxodus/export` publish
in that order in one job, because the companion peer-depends on the exact matching
`docxodus` version, so it *cannot* go first. A failure in the companion step therefore
leaves NuGet and `docxodus` published while `@docxodus/export` is not. Re-run with
`gh workflow run publish.yml --ref main -f version=X.Y.Z` after fixing: NuGet pushes use
`--skip-duplicate`, and an already-published npm version fails loudly rather than silently
overwriting.

**One-time bootstrap for `@docxodus/export`.** The companion publishes through npm OIDC
trusted publishing, which is configured *per package* and therefore cannot be configured for
a package that does not exist yet. Its first-ever publish fails with
`404 Not Found - PUT https://registry.npmjs.org/@docxodus%2fexport`, which reads like a
permissions bug and is not one. Bootstrap it once by hand — create the `@docxodus` scope,
publish one version with a token, then configure trusted publishing on the package — after
which CI owns it. Until that is done, expect every release to publish everything except the
companion.

### Post-release: re-pin the demos

`docs/demo/` loads the library from jsDelivr, so its `docxodus@X.Y.Z` pins can only move
*after* npm publishes and the CDN serves the new bundle. Confirm with a real fetch
(`curl -I https://cdn.jsdelivr.net/npm/docxodus@X.Y.Z/dist/embed.bundle.js`), then update the
demo pages, `docs/demo/README.md`, `docs/npm-package.md`, `npm/README.md`,
`npm/examples/embed.html`, and the `RELEASE_ENGINE` constant in
`npm/tests/social-demo.spec.ts` — that spec is the guard that proves the demos load the pin
rather than a 404, so run it. One reference in `docs/demo/README.md` is prose recounting a
past pin-ahead-of-release; leave it alone.

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
- The upstream `archived-examples/` console projects were removed with the SpreadsheetML and
  PresentationML modules — they exercised
  the spreadsheet/presentation modules and were never in the solution. `git log` has them.

For bugfix history, use `git log` rather than maintaining a list here.

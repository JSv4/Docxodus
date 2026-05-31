# Changelog

All notable changes to this project will be documented in this file.

## [Unreleased]

## [6.4.0] - 2026-05-30

### Added
- **npm `DocxSession` find-by surface** (issue #171). The TypeScript wrapper at
  `npm/src/session.ts` now exposes the six `DocxSession` methods whose bridge
  shells landed in #168 but were unreachable from the typed API: `exists`,
  `findByText`, `findAllByText`, `findByRegex`, and `findByKind` (`replaceMatch`
  was already present). New `FindOptions` type in `npm/src/types.ts`
  (`ignoreCase` / `ignoreWhitespace` / `kindFilter` / `scopes` / `scopeFilter`,
  matching the .NET `FindOptions` record), and the corresponding `Exists` /
  `FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` signatures added
  to the `DocxSessionBridge` exports interface. Wire shapes are byte-identical to
  what `tools/python-host` consumes, preserving the cross-transport parity
  invariant. Tests: `npm/tests/find-by.spec.ts` exercises the *typed* wrapper
  (via a new `window.Docxodus.openTypedSession` harness helper backed by an
  IIFE-bundled `session.ts`), covering case-sensitive/insensitive text search,
  broad-pattern regex, kind+scope filtering, existence probes, and the
  `grep → replaceMatch` round-trip.

### Fixed
- **Python `FindOptions` scope filtering was a silent no-op.** In `docx_scalpel`, `FindOptions` exposed a single `scope_filter: ProjectionScopes` field and serialized it as an **int** under the `scopeFilter` wire key. But the stdio host (and WASM bridge) parse `scopeFilter` as a **string** (a single named part like `"hdr1"`) and read the coarse `ProjectionScopes` flag set from a separate `scopes` (number) key — which the wrapper never emitted. The net effect: `find_by_text` / `find_all_by_text` / `find_by_regex` ignored any scope restriction passed from Python and always searched all scopes. `FindOptions` now mirrors the .NET record's two distinct controls: `scopes: ProjectionScopes | None` (coarse category flag set → wire `scopes`, int) and `scope_filter: str | None` (fine named-part post-filter → wire `scopeFilter`, string). **API change:** `scope_filter` is now a `str` (was `ProjectionScopes`); callers that want category filtering should use the new `scopes` field (e.g. `FindOptions(scopes=ProjectionScopes.HEADERS | ProjectionScopes.FOOTERS)`). Tests: `python/tests/test_find_options_scopes.py` pins the wire mapping and verifies `scopes=ProjectionScopes.BODY` actually drops a footnote-scope hit on HC031.

## [6.3.0] - 2026-05-30

### Fixed
- **Percentage table/cell widths no longer crash `WmlToHtmlConverter`** (issue #210). `convertDocxToHtml` threw `FormatException` ("Format_InvalidStringWithValue, 100%") whenever a table-level (`w:tblW`) or cell-level (`w:tcW`) width used `w:type="pct"` with a percent-suffixed value such as `w:w="100%"` / `w:w="50%"`. This is the form the `docx` npm library emits for `WidthType.PERCENTAGE`, and it is valid per the OOXML `ST_TblWidth` / `ST_MeasurementOrPercent` schema (which permits either a plain integer in fiftieths-of-a-percent **or** a `"<number>%"` string). The converter cast the attribute straight to `int`, which throws on `"100%"`; DXA (twips) widths were unaffected because they are always plain integers. Width parsing for `w:tblW`/`w:tcW` now goes through a single `ParseTblWidthValue` helper that tolerates the percent-suffixed form: an explicit `"100%"` is treated as a literal percentage, while a bare integer under `pct` is still interpreted as fiftieths of a percent (`5000` → `100%`). Non-numeric/garbage widths are ignored gracefully instead of throwing. Tests: `HcTablePercentageWidthTests` in `Docxodus.Tests/HtmlConverterTablePercentageWidthTests.cs`.

### Added
- **Python DOCX→HTML conversion** — `convert_docx_to_html(data, options)` and
  `DocxSession.to_html(options)` in `docx_scalpel`, backed by a new shared
  `HtmlConversionOps` core renderer that the WASM bridge now also delegates to.
  New `HtmlOptions` dataclass mirrors the existing WASM/npm conversion options.
  New stdio-host ops: `convert_to_html`, `session_to_html`.

## [6.2.0] - 2026-05-28

### Added
- **`WorkerDocxodus.prepare()` — optional comparison-path warmup** (consumer issue JSv4/crowdsourced-redlines-js#2). `createWorkerDocxodus()` warms the .NET WASM runtime but does **not** exercise the comparison engine, so the first `compareDocuments()` pays a one-time warmup cost — comparison-assembly initialization plus JIT of the diff/XML stack — on top of the diff work (~2x the steady-state latency). Consumers worked around this by shipping seed `.docx` fixtures and running a throwaway compare for the side effect. The new optional `prepare(): Promise<void>` pays the cost up front with no caller IO: it runs a complete comparison inside the worker against two tiny seed documents constructed in-memory on the .NET side (no seed fixtures to ship), forcing the full compare path to resolve and JIT. After `await prepare()`, the first real `compareDocuments()` / `compareDocumentsToHtml()` runs at steady-state speed and triggers no further `.wasm` fetches. `prepare()` is never called automatically — skip it and the first compare absorbs the warmup as before. It is idempotent (repeated/concurrent calls share one in-flight warmup and resolve immediately once complete) and concurrent-safe (a `compareDocuments()` issued while a `prepare()` is in flight does not double-load assemblies). Implemented as a new `Warmup()` `[JSExport]` on `DocumentComparer`, a `"prepare"` worker message, and the `WorkerDocxodus.prepare()` proxy method. Tests: `npm/tests/worker-prepare.spec.ts` verifies (via page-level `.wasm` request monitoring + in-worker timing) that after `prepare()` a real compare fetches no additional `.wasm`, a warmed first compare runs ~2x faster than a cold one (~758ms vs ~1504ms), a second `prepare()` resolves in <50ms, and concurrent prepare+compare never double-loads.

## [6.1.0] - 2026-05-28

### Changed
- **`DocxSession.DeleteRange` / `DeleteSection` honor `TrackedChangeMode.RenderInline`** (issue #177). The bulk-delete primitives now produce native Word tracked-deletion markup instead of silently performing a structural delete in tracked mode. Each removed paragraph has every direct-child run wrapped in `w:del` (reusing `WrapRunsInDel`) and the paragraph-mark marked deleted via `w:pPr/w:rPr/w:del` — the combination Word interprets as "this entire paragraph is a tracked deletion", so accepting the change actually removes the block (the old `DeleteBlock`-tracked path left empty paragraphs behind). Tables get `w:trPr/w:del` on every row (Word's row-deletion convention — there is no table-level "delete" markup) plus the same run/paragraph-mark wrapping inside every cell; nested tables recurse. Anchors stay live in the document tree, so the top-level block anchors are reported via `EditResult.Modified` instead of `Removed` — matching `DeleteBlock`'s existing tracked-mode contract. Block kinds outside `w:p`/`w:tbl` (e.g. `w:sdt` content controls appearing mid-range) still fall back to structural removal in tracked mode. No wire-shape changes — the WASM bridge, npm wrapper, and Python stdio host pick up the new behavior automatically through `DocxSessionOps`. Tests: `DS271`–`DS273`.

### Added
- **`DocxSession` block-metadata read surface.** New methods
  `GetBlockMetadata` / `GetBlockMetadatas` / `GetListMembership` /
  `GetSectionInfo` expose paragraph style id+name, outline level, list
  membership (`numId`/`abstractNumId`/`ilvl`/format/start-override/
  inherited-from-style flag), and the enclosing `w:sectPr` (page
  size/orientation/margins/columns/header/footer parts). New types
  `BlockMetadata`, `ListMembership`, `SectionInfo`, and the
  `NumberFormat` enum (`Decimal`/`UpperLetter`/`LowerLetter`/
  `UpperRoman`/`LowerRoman`/`Bullet`). Surfaced in WASM
  (`DocxSessionBridge`), npm (`DocxSession.getBlockMetadata` etc.), and
  Python (`docx_scalpel.session.DocxSession.get_block_metadata` etc.).
- **Markdown projection: list-item classification now follows `pStyle`
  chain.** `WmlToMarkdownConverter` previously labeled a paragraph as
  a list item only when it carried inline `w:numPr`. Now it also walks
  the `pStyle → basedOn` chain (16-level cycle guard) and labels the
  paragraph as a list item if any ancestor style contributes `w:numPr`.
  Brings the projector into agreement with `GetListMembership` for
  style-inherited list items.
- Annotation write surface on `DocxSession` (`AddAnnotation`,
  `RemoveAnnotation`, `UpdateAnnotation`, `MoveAnnotation`) exposed across
  .NET, WASM (`@docxodus/wasm`), and Python (`docx-scalpel`). New
  `EditErrorCode` values: `DuplicateAnnotationId`, `AnnotationNotFound`,
  `EmptyAnnotationSpan`. `EditResult` gained an `AnnotationId` field.
  `AnnotationUpdate` is the new partial-update payload for
  `UpdateAnnotation`. `listAnnotations` now surfaces the `metadata` bag in
  its JSON output (previously omitted).
- **`FillOptions.CoalesceWhitespaceAroundEmptyFill`** (issue #188). New opt-in flag on `DocxSession.FillPlaceholders` that smooths over the canonical template-filling cosmetic foot-gun: returning `""` from the picker (the standard "drop this optional clause entirely" signal) deletes the brackets exactly, leaving surrounding whitespace verbatim. The NVCA Model COI repro `"… on March 14, 2024 [under the name [_______________]]."` with the outer wrapper dropped becomes `"… on March 14, 2024 ."` (note the stray space before the period). When `CoalesceWhitespaceAroundEmptyFill = true`, an empty fill (after `$`-prefix preservation has been applied) absorbs adjacent chars based on the immediate neighbors of the span: whitespace on both sides → collapse to one space; whitespace before + clause-terminating punctuation (`. , ; : ! ?`) after → drop the leading space; matched open/close brackets (`() [] {}`) on either side → drop both. NBSP / narrow NBSP / thin space are folded to ASCII space. Default `false` (preserve current literal-delete behavior). The .NET implementation reads the live flat text of the enclosing block so the rules work regardless of `Boundary` setting; the npm TS implementation uses the match's already-populated `contextBefore` / `contextAfter` (so callers combining `boundary: ContextBoundary.Bracket` with this option won't see the bracket-coalesce rule fire on the JS side — leave `boundary` at the default `Char`). Tests: `DS247a`–`DS247f`.
- **`DocxSession.GetDiff(DiffFormat.Unified | SideBySide)` — line-based diff formats** (issue #178). The two enum values previously reserved as v2-deferred (`NotSupportedException`) now produce real output. `Unified` returns a `patch(1)`-compatible unified diff over the initial-vs-current markdown projections (`--- initial` / `+++ current` headers, 3 lines of context per hunk, hand-rolled `O(n*m)` LCS over `\n`-split lines; empty string when nothing has changed). `SideBySide` returns a `diff -y`-style two-column rendering — the initial projection padded to 72 chars on the left, a one-character marker (`' '` unchanged, `'|'` modified, `'<'` only-initial, `'>'` only-current), then the current projection on the right. Adjacent `Delete + Insert` pairs collapse to a single `|` "modified" row. Decision: hand-rolled LCS over a pulled-in dependency to keep the WASM build AOT-friendly (no reflection-based serializers) and avoid adding a NuGet edge case. The npm wrapper's `getDiff` is now overloaded — `getDiff()` / `getDiff(DiffFormat.Json)` returns `DiffEntry[]`, `getDiff(DiffFormat.Unified | SideBySide)` returns `string`. Tests: `DS289`–`DS289e` (replacing the prior `NotSupportedException` assertion with positive-case coverage of unified hunk shape, side-by-side marker column alignment, insert/delete/modify marker detection, and an out-of-range enum guard); Playwright `edit-summary-and-diff.spec.ts` extended with two end-to-end cases over the WASM bridge.
- **`docx_scalpel.DocxSession.fill_placeholders(picker, options?)` — Python wrapper for the C# template-fill loop** (issue #192, item 1). Mirrors the C# `DocxSession.FillPlaceholders` and the TypeScript `session.fillPlaceholders`: bundles reverse-offset ordering within a paragraph, `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass iteration for nested AlternativeClause brackets — the three foot-guns every template-fill agent was previously re-implementing in ~25 lines of subtle correctness. New `FillOptions` (kinds / scope / max_passes / preserve_dollar_prefix / context_chars / boundary) and `BulkEditResult` (filled / skipped / passes / still_present / unfilled / errors) dataclasses on the public surface. Runs entirely in Python over the existing `find_placeholders` + `replace_match` primitives (no new wire op), matching the TS wrapper's design. The `still_present` field is a post-loop `find_placeholders` count — mirrors the C# `BulkEditResult.StillPresent` field added in #191 for the trustworthy single-call "is the template done?" check. Tests: `python/tests/test_fill_placeholders.py` covering BlankFill replacement, picker-returning-None skip dedup, dollar-prefix on/off, nested-clause multi-pass convergence, max_passes validation, the AlternativeClause-visiting default kinds, and the `still_present == 0` convergence assertion.
- **`BulkEditResult.StillPresent` — trustworthy "is the template done?" metric on `FillPlaceholders`** (issue #189). `BulkEditResult.Skipped` counts placeholders the picker returned `null` for in the first pass that saw them, deduplicated across passes — but this stays `> 0` even when later passes finish the job (e.g. a nested-outer wrapper becomes fillable once its inner is stripped, or a structural delete removes the placeholder entirely). Agents reading `Skipped > 0` after a clean fill ended up cross-checking against `GetEditSummary().RemainingPlaceholders.Count` to know the truth. New additive field `BulkEditResult.StillPresent` is a post-loop `FindPlaceholders(opts.Kinds, opts.Scope).Count` — the metric to assert on for the single-call check. `Skipped > 0 && StillPresent == 0` now correctly reads as "picker skipped on first sight but later passes resolved it." `Skipped` retained for back-compat with its docstring sharpened to direct callers to `StillPresent`. Mirrored on the npm wrapper (`BulkEditResult.stillPresent`) — the TS-side multi-pass loop in `npm/src/session.ts` recomputes via `findPlaceholders`. Tests: `DS247` (multi-pass convergence: `Skipped > 0`, `StillPresent == 0`), `DS248` (picker returns null everywhere: `StillPresent` equals remaining count).
- **`FindOptions.Scopes` (`ProjectionScopes` flag set) + `session.AnchorsByScope`.** The `FindBy*` helpers previously had to default to body and use a string `ScopeFilter` to widen — surveying headers/footers/footnotes meant either passing a magic string like `"hdr1"` (which only matches one part) or walking `Project().AnchorIndex` and filtering by scope name manually. The new `FindOptions.Scopes` field is typed and composable: `Scopes = ProjectionScopes.Headers | ProjectionScopes.Footers` searches every header and every footer in one call. Defaults to `All` so existing callers see no behavior change. The string `ScopeFilter` remains for the rare case of pinning one specific named part (e.g. `"hdr1"` only); it now applies as a finer post-filter on top of `Scopes`. `session.AnchorsByScope(scopes)` is the search-free convenience for the common "enumerate every anchor in scope X" pattern. A new `ProjectionScopesExtensions.IncludesScope(scopeName)` helper exposes the scope-name → flag mapping (`hdr*` → `Headers`, `ftr*` → `Footers`, etc.) for callers that want it directly. Wire shape: `FindOptions` JSON now reads optional `scopes` (number); WASM/Python bridges pick it up automatically. Tests: `DS290`–`DS294`.
- **`DocxSession.CompactRuns(scopes?)` — remove formatting-only run residue.** Public, transactional, scope-aware primitive that removes every `w:r` whose only content is a `w:rPr` (no text, no tabs, no breaks, no field/footnote/comment references). Useful after any workflow that deletes inline content and leaves behind styled-but-empty runs — accepting tracked changes, removing footnotes/comments, run-text refactors. One pre-op snapshot is taken so a single `Undo()` rolls every removal back together; block-level anchors are unaffected because run-level Unids aren't part of the `AnchorIndex`. Defaults to `ProjectionScopes.All` so a call after a body edit also tidies header/footer/footnote/endnote/comment parts; callers that only want body cleanup can pass `ProjectionScopes.Body`. Returns a `CompactResult { RunsRemoved }` so callers can detect "did anything change" without a separate projection round-trip. Tests: `DS295`–`DS298`.
- **`AnchorTarget.AutoNumberPrefix` + `FullText`, mirrored on `AnchorInfo`.** Paragraphs / headings / list items in the body that carry numbering (inline `w:numPr` or numbering inherited from a style) now expose Word's resolved numbering label — `"1."`, `"1.1"`, `"First"`, etc. — as `AutoNumberPrefix` on the projection's `AnchorTarget` and on the `AnchorInfo` returned by `GetAnchorInfo` / `GetAnchorInfos`. `FullText` is a derived convenience that joins prefix + `TextPreview` with a space when a prefix is present. Closes the foot-gun where a caller could see `"# First The total…"` in the markdown projection but a `Grep`/`FindByText` for `"First"` would silently miss it (run text contains only `"The total…"`). The prefix is *not* added to `TextPreview` and is *not* searchable via `Grep` — `Grep` continues to walk run text only — but callers iterating `AnchorIndex` for previews or building search facets now have the rendered label available without re-resolving numbering. Mirrored on the WASM bridge (`MarkdownAnchorTargetDto`, `AnchorInfo` serializers) and the npm wrapper types. Body-only in v1 — header/footer numbering paths aren't routed through `ListItemRetriever` yet. Tests: `DS222`, `DS222a`, `DS222b`.

### Changed
- **Deterministic content-addressable Unids in the markdown projector.** `WmlToMarkdownConverter` now assigns `PtOpenXml.Unid` values via a content-addressable hash (`UnidHelper.AssignToAllElementsDeterministic`) rather than `Guid.NewGuid()`. The Unid is SHA-256(`parent_unid : tag : content_sig : dup_index`) truncated to 32 hex. Properties: same bytes → same Unids across sessions; editing a paragraph's text changes only that paragraph's Unid; inserting a unique-content paragraph anywhere doesn't shift any sibling's Unid; inserting/editing a duplicate-content paragraph shifts `dup_index` of later duplicates only. Closes the cross-session non-determinism foot-gun where a CLI script capturing anchor ids in one run would find them unresolvable in a follow-up run over the same bytes (without `PersistAnchorIds = true`). `WmlComparer` intentionally keeps the random-Guid path (`UnidHelper.AssignToAllElements`) — its matching heuristics expect content-independent Unids, and making them content-addressable inflates the detected revision count on fixtures with same-tag-but-distinct-content elements (verified against `WC003_Compare` on `WC022-Image-Math-Para`). Container elements (those that have block-level descendants) collapse to a tag-name-only signature so editing one block doesn't ripple through the parent's Unid into sibling blocks. Tests: `DS300`–`DS304`.
- **`FillOptions.Kinds` default → `PlaceholderKinds.All`.** The prior default (`BlankFill | Instruction`) silently excluded `AlternativeClause` placeholders, so a picker with `[two]` → `"two"` style bracket-stripping rules would appear to do nothing on those matches — confusing for any caller that wrote a single picker covering every kind it might see. The new default invokes the picker for every kind in the doc; pickers should return `null` for placeholders they don't recognize (the long-standing skip contract). Callers that relied on the prior filter behavior can set `Kinds = PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction` explicitly.

### Changed
- **`docx_scalpel` — `from_wire` decoders renamed to `_from_wire`** (issue #192, item 4). `AnchorTarget`, `TextMatch`, `EditResult`, `TemplatePlaceholder`, and every other value type's JSON-deserializer classmethod is now leading-underscore. The decoders are the stdio transport's JSON-to-dataclass adapter and were never intended for caller use — but their public name made them show up in `dir(ds.AnchorTarget)`, IDE autocomplete, and `help(...)` output, drawing new users into asking what they're for. The public surface is now exactly the dataclass fields plus the user-facing helpers. No back-compat shim — the package is `0.1.0a*` and nothing outside the wrapper itself should have been calling `from_wire`.

### Fixed
- **`DocxSession.GetDiff` — `ArgumentException` on duplicate Unid keys** (issue #187). `ComputeDiff` built its initial/current lookup dictionaries via `AnchorIndex.Values.ToDictionary(t => t.Unid, …)`, which threw the moment two `AnchorTarget`s shared a raw Unid. Two ways this happened in practice: (1) under non-`FullUnid` rendering the `AnchorIndex` is dual-keyed (each target is reachable via its full Unid and its rendered alias), so `Values` enumerated every target twice; (2) the deterministic Unid scheme seeds each scope's root with the root element's local name (`"hdr"` for every header part, `"ftr"` for every footer part), so structurally-identical first paragraphs across multiple header/footer parts hashed to the same raw Unid in different scopes — reproduced on the public NVCA Model COI on the default settings. `ComputeDiff` now keys by `(Anchor.Scope, Unid)` and dedupes the `Values` enumeration with `DistinctBy`, which fixes both paths without changing diff semantics: scope is stable across mutations and the kind-flip case (`p`→`h` via `SetParagraphStyle`) still resolves to the same composite key. Tests: `DS289a` (cross-scope collision smoke), `DS289b` (mutation isolates to the edited scope), `DS289c` (Abbreviated rendering doesn't crash).
- **`tools/python-host/pyhost.csproj` — suppress StyleCop SA1633/SA1636 file-header rules** (issue #173). `dotnet build -c Release tools/python-host/pyhost.csproj` was failing because `Directory.Build.props` sets `TreatWarningsAsErrors=true` for Release and the python-host project inherited the StyleCop ruleset without suppressing the file-header warnings on `Dispatcher.cs` and `Program.cs`. Added `<NoWarn>$(NoWarn);SA1633;SA1636</NoWarn>` to the csproj, matching the existing convention in `wasm/DocxodusWasm/DocxodusWasm.csproj` for tooling/wasm subprojects.
- **`DocxSession.GetDiff` JSON serialization in WASM** (issue #166). `SerializeDiff` originally called `System.Text.Json.JsonSerializer.Serialize(string)` for `anchorId` / `before` / `after` escaping, which uses the reflection-based serializer that the WASM build explicitly disables. Browser callers got `JsonSerializerIsReflectionDisabled` thrown for any non-empty diff (empty `"[]"` short-circuited). Replaced with a hand-rolled `AppendJsonString` helper that mirrors `DocxSessionJson.JsonString`'s escape table. The .NET-side `DS285`/`DS286`/`DS287` tests passed because the standard runtime allows reflection; the Playwright spec from Unit E uncovered the WASM-side breakage.

### Added
- **`DocxSession.ProjectAnchor(anchorId, depth?)`** — project a slice of the document keyed by anchor (one paragraph, a subtree, or a whole heading section) instead of paying the cost of projecting the entire document each time. `ProjectionDepth.SelfOnly` returns just the addressed block, `Subtree` adds descendants, and the default `SubtreeAndFollowingSiblings` extends headings forward through the section bounded by the next same-or-higher heading. The returned `MarkdownProjection.AnchorIndex` is filtered to the scoped Unids only. Useful for showing an LLM one section at a time. Shared core (`DocxSessionOps.ProjectAnchor`) wires the WASM bridge and the Python stdio host; npm wrapper: `session.projectAnchor(anchorId, depth?)` with the `ProjectionDepth` const re-exported. (#167)
- **`WmlToMarkdownConverterSettings.AnchorIdRendering`** — new projection setting controlling how anchor ids appear in `{#…}` tokens. `FullUnid` (default, legacy) keeps the 32-hex-char Unid; `Abbreviated` trims each Unid to the shortest unique prefix per `(kind, scope)` bucket (4-char floor) saving ~5-10% of token budget; `Sequential` replaces Unids with 1-based per-bucket counters in document order — maximally token-efficient for one-shot LLM contexts. The returned `MarkdownProjection.AnchorIndex` is **dual-keyed** in non-`FullUnid` modes: lookups by either full Unid or rendered id resolve to the same `AnchorTarget`, so callers can roundtrip rendered ids straight back to anchor-addressed methods (`DocxSession.ProjectAnchor`, `ReplaceText`, …) without an explicit translation step. `Anchor.Token` continues to return the canonical full-Unid form regardless of rendering mode. Plumbed through the WASM bridge (`MarkdownProjectionSettingsDto.AnchorIdRendering`) and the npm wrapper (`MarkdownProjectionSettings.anchorIdRendering` + exported `AnchorIdRendering` enum). (#167)
- **`GetEditSummary` + `GetDiff` for edit-state introspection** (issue #166). `DocxSession.GetEditSummary()` returns a single `EditSummary` record composing existing primitives — `RemainingPlaceholders` (from `FindPlaceholders`), `BareUnderscoreRuns` (from `Grep`), `TotalAnchors`, `FootnoteCount`, `InlineFootnoteRefCount`, `CommentCount`. Lets verification logic at the end of an edit pipeline be declarative (`Assert.Empty(summary.RemainingPlaceholders)`) instead of a regex zoo. `DocxSession.GetDiff(DiffFormat = Json)` compares the projection captured at session construction time against the current projection and returns an anchor-keyed JSON array of `DiffEntry` records (`op: delete | insert | modify`, `anchorId`, optional `before` / `after`). Gated by new `DocxSessionSettings.CaptureInitialProjection` (default `true`; set `false` to skip the ~200ms upfront cost when you don't plan to diff). `DiffFormat.Unified` and `DiffFormat.SideBySide` are reserved enum values that throw `NotSupportedException` in v1 — see issue #178 for the line-based diff follow-up. `DocxSession.RemainingPlaceholders(kinds)` is a thin discoverability alias for `FindPlaceholders`. Shared core (`DocxSessionOps.GetEditSummary` / `RemainingPlaceholders` / `GetDiff`) propagates to both the WASM bridge and the Python stdio NDJSON host (`get_edit_summary`, `remaining_placeholders`, `get_diff` ops). npm wrapper: `session.getEditSummary()`, `session.remainingPlaceholders(kinds?)`, `session.getDiff(format?)` with `DiffFormat`/`EditSummary`/`DiffEntry` types re-exported. Tests: `DS280`–`DS289`, Playwright `edit-summary-and-diff.spec.ts`.
- **`DeleteRange` and `DeleteSection` for bulk block removal** (issue #165). `DocxSession.DeleteRange(fromAnchorId, toAnchorIdExclusive)` deletes every top-level block-level sibling between two anchors in one call, with one transactional `Undo()` snapshot. Both anchors must share a direct parent and live in the same package part — anchors in different parts return `AnchorsNotAdjacent`, anchors with different parents (e.g. one inside a table cell) also return `AnchorsNotAdjacent`, and `from` not preceding `to` in document order returns `InvalidPosition`. `DocxSession.DeleteSection(headingAnchorId)` is a thin convenience: resolves the heading's level via `WmlToMarkdownConverter.HeadingLevel`, scans forward siblings for the next heading at the same or higher level, and delegates to a shared internal helper. If the target is the last heading in its parent, the section extends to the end. Tracked-change mode is documented as "v1 does structural delete regardless" — wrapping every run across many blocks in `w:del` is deferred until a consumer needs it. Shared core (`DocxSessionOps.DeleteRange` / `DeleteSection`) propagates to both the WASM bridge and the Python stdio NDJSON host (`delete_range`, `delete_section` ops). npm wrapper: `session.deleteRange(fromId, toIdExclusive)`, `session.deleteSection(headingAnchorId)`. Refactor: `WmlToMarkdownConverter.IsHeading` and `HeadingLevel` promoted `private static → internal static` so `DocxSession` can reuse them without duplication. Tests: `DS260`–`DS270`, Playwright `delete-range-section.spec.ts`.
- **`ContextBoundary` enum + widened default `contextChars`** (issue #164). `DocxSession.Grep`, `GrepCrossBlock`, and `FindPlaceholders` now accept a `ContextBoundary` parameter that controls where the context-computation walker stops: `Char` (default, legacy truncate-at-N behavior), `Bracket` (stop at `[`/`]` — the dominant template-fill case for unambiguous per-placeholder context), `Sentence` (stop at `.!?:;`), `Comma` (stop at `,`). Default `contextChars` widened from 40 → 80 across all three methods so plain `.Contains` checks have enough text to disambiguate without the agent dropping into boundary mode. `FillOptions` gains `ContextChars` + `Boundary` fields threaded into the internal `FindPlaceholders` calls. Shared core (`DocxSessionOps.Grep` / `GrepCrossBlock` / `FindPlaceholders`) propagates the new param so both the WASM bridge and the Python stdio NDJSON host pick it up (npm `GrepOptions.boundary`, exported `ContextBoundary` const). Tests: `DS250`–`DS255`, Playwright `context-boundary.spec.ts`.
- **Template-fill convenience — `FillPlaceholders`, `ReplaceInner`, `AlternativeKinds`** (issue #163). `DocxSession.FillPlaceholders(picker, options?)` bundles the three foot-guns every template-fill agent re-implements: reverse-offset ordering across matches within a paragraph, `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass iteration for nested AlternativeClause brackets. Returns a `BulkEditResult` with `Filled` / `Skipped` / `Passes` counts plus per-failure error and unfilled-placeholder lists. New `DocxSession.ReplaceInner(match, newInner)` overload replaces only the bracketed portion of a match, preserving any prefix or suffix outside it — the canonical use case for `$[___]` matches where the regex `\$?\[…\]` captured a leading `$`. `TemplatePlaceholder.AlternativeKinds` is a new additive field listing secondary classifications when the primary `Kind` is borderline (e.g. a long bracketed clause containing a `_______` blank: primary `Kind` stays `BlankFill` for back-compat, with `AlternativeClause` in `AlternativeKinds`). Shared core: `DocxSessionOps.ReplaceInner` (used by both the WASM bridge and the Python stdio host, so `replace_inner` is also exposed via the NDJSON dispatcher); `DocxSessionJson.SerializePlaceholders` emits `alternativeKinds`. npm wrapper: `session.replaceInner(match, newInner)`, `session.fillPlaceholders(picker, options?)` (TS-side mirror of the .NET control loop), new `FillOptions` and `BulkEditResult` types. Tests: `DS230`–`DS233` (ReplaceInner + AlternativeKinds), `DS240`–`DS246` (FillPlaceholders incl. MaxPasses validation + Passes-counter semantics), Playwright `fill-placeholders.spec.ts`.
- **`Docxodus.Internal.{SessionRegistry, DocxSessionOps, DocxSessionJson}` — shared bridge core for `DocxSession` transports.** Lifts the integer-handle pool, the per-op session-lookup + serialization facade, and the StringBuilder JSON helpers that previously lived inside `wasm/DocxodusWasm/DocxSessionBridge.cs` into the core library under `Docxodus/Internal/`. The WASM bridge is now a thin `[JSExport]`-attributed shell over `DocxSessionOps`; a new stdio NDJSON host at `tools/python-host/` (assembly `docxodus-pyhost`) consumes the same facade, so the WASM/TypeScript and stdio/Python clients see byte-for-byte identical JSON wire shapes. `InternalsVisibleTo` for `DocxodusWasm`, `docxodus-pyhost`, and `Docxodus.Tests`. Pure refactor — all 1411 existing tests pass unchanged.
- **`tools/python-host/` — .NET 8 console host for the upcoming python-docxodus wrapper.** Reads NDJSON requests on stdin, dispatches to `DocxSessionOps`, writes NDJSON responses on stdout (diagnostics on stderr). One host process serves many concurrent sessions via the shared handle pool, so an agentic Python pipeline pays the .NET startup cost once and gets µs-to-low-ms per-op latency thereafter. Distinguishes transport-level failures (`ok: false` envelope) from business `EditResult.Success = false` outcomes (`ok: true` envelope carrying the `EditError`). Built for self-contained single-file `dotnet publish` so the eventual pip wheel ships with zero system dependencies.
- **`Exists` / `GetAnchorInfo` / `GetAnchorInfos` / `FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` exposed on both bridges.** Closes the remaining gap where these public `DocxSession` methods existed in the .NET API but had no wire serializer (so they were unreachable from any non-.NET client). Lands them once in `DocxSessionOps`; the WASM `[JSExport]` shell and the stdio NDJSON dispatcher pick them up automatically. `FindOptions { IgnoreCase, IgnoreWhitespace, KindFilter, ScopeFilter }` on the wire as `{ignoreCase?, ignoreWhitespace?, kindFilter?, scopeFilter?}`. `GetAnchorInfos` bulk lookup follows issue #162's design: `{anchorIds: string[]} → {id: AnchorInfo | null}`. `ReplaceMatch(TextMatch)` is intentionally **not** a wire op — `ReplaceTextAtSpan(anchor, span.start, span.length, replace)` already exposes the underlying primitive; client wrappers implement `replaceMatch(match, replace)` as a 1-line helper rather than ship an 80-line `TextMatch` parser on every transport.
- **Anchor introspection ergonomics — `TextPreview` on `AnchorTarget`, boilerplate footnote filter, `GetAnchorInfos` bulk lookup** (issue #162). `WmlToMarkdownConverter` now computes the first ~80 chars of each block element's flat text during projection and exposes it as `AnchorTarget.TextPreview` — agents no longer need an N-anchor walk via `session.GetAnchorInfo` to surface previews when iterating the `AnchorIndex`. Word-reserved `w:footnote`/`w:endnote` separators (`type="separator"` / `type="continuationSeparator"`) no longer appear in the projection's `AnchorIndex` (they were internal Word plumbing surfaced as un-deletable `fn:fn:*` anchors). New `DocxSession.GetAnchorInfos(IEnumerable<string>)` returns a dictionary mapping each requested id to its `AnchorInfo?` in a single pass; unknown ids map to `null`. WASM bridge surfaces `textPreview` on `MarkdownAnchorTargetDto`, on session `Project()` responses, and on `FindBy*` results; adds `GetAnchorInfo` / `GetAnchorInfos` JSExports. npm wrapper: new `textPreview` field on `MarkdownAnchorTarget`, `AnchorTargetRef`, and `DocxSessionProjection.anchorIndex`; `session.getAnchorInfo(id)` and `session.getAnchorInfos(ids[])` methods. Tests: `MD005` (anchor TextPreview), `MD006` (boilerplate filter), `DS220`–`DS221` (bulk lookup), Playwright `anchor-introspection.spec.ts`.

## [6.0.0] - 2026-05-25

### Fixed
- **`WmlComparer` — defensive null/empty guards on three sibling consumer sites flagged by issue #128.** Follow-up to PR #124, which guarded `FindIndexOfNextParaMark`. The same `cul`-can-contain-`ComparisonUnitGroup` (and empty-descendant) hazard existed in three more places that would have crashed with `NullReferenceException` or `InvalidOperationException` (`.Last()` on empty) had the inputs reached them: `FindCommonAtBeginningAndEnd` (boundary atom dereference), `SplitAtParagraphMark` (paragraph-mark search), and `DoLcsAlgorithm` (last-atom lookup). The producer (`CreateComparisonUnitAtomListRecurse` + `ElementsToThrowAway`) already correctly filters body-level `w:bookmarkStart`/`w:bookmarkEnd`/`w:permStart`/`w:permEnd`/`w:proofErr`, so these guards are belt-and-braces. Adds `WmlComparerBodyLevelElementsTests` with five small programmatic fixtures (bookmarks, perm markers, proof-error markers at body level) that assert `Compare` succeeds — replacing the original 4 MB binary pair's weaker "no NRE" assertion for the body-level case.

### Added
- **`DocxSession.GrepCrossBlock` — text search that may span adjacent paragraphs** (issue #146). Extends `Grep` (#143) so a single match can cross block boundaries among adjacent block-level siblings (paragraphs/headings/list items) under the same direct parent. Returns `CrossBlockMatch` records, each carrying `EnclosingAnchors` (every block the match touches, in doc order) and a per-block `Slices[]` breakdown — every slice has its own `SpanInBlock`, `Fragments`, and `Anchor`, so callers can preserve per-fragment formatting when rewriting. Block boundaries appear in the concatenated text as a single `\n`, so `^`/`$` with `RegexOptions.Multiline` anchor at boundaries and `.` doesn't cross unless `Singleline` is set. Matches are scoped strictly: they never cross OOXML package parts (body → footnote), container boundaries (body → table cell), or non-paragraph siblings (a `w:tbl` between two paragraphs breaks the run). Superset of `Grep`: single-block matches still appear with one `Slice` — callers wanting only cross-block hits can filter `Slices.Count > 1`. Edit semantics deferred (per-slice vs merge vs boundary-preserve has no obviously-right default; file follow-up when a consumer needs it). WASM bridge (`GrepCrossBlock` JSExport) + npm wrapper (`session.grepCrossBlock(pattern, options?)`) + new TS types (`CrossBlockMatch`, `BlockSlice`). Tests: `DS200`–`DS209` + Playwright spec.
- **`DocxSession.FindByAnnotation` / `FindByLabel` / `FindByBookmark` / `ListAnnotations` — annotation-based anchor discovery** (issue #132). Bridges the read-side annotation API (`AnnotationManager`, which persists user labels as `_Docxodus_Ann_{id}` bookmarks + a custom XML metadata part) to the write-side session, so an agent told to "edit the indemnification clause" looks up the annotation by id and immediately gets the `AnchorTarget`s to hand to `ReplaceText` / `Raw.GetXml`. v1 returns every block-level anchor (paragraph/heading/list-item/cell/row/table) whose subtree overlaps the bookmark range, sorted in document order; callers filter by `kind` if they only want text-bearing blocks. `FindByLabel` keys by annotation id so multiple regions sharing one label stay disambiguated. `FindByBookmark` accepts any bookmark name (managed or user-authored) as an escape hatch. Long-lived sessions read annotations directly off the open `WordprocessingDocument` (new `AnnotationManager.GetAnnotations(WordprocessingDocument)` overload) — no byte-level save/reopen per query. WASM bridge (`FindByAnnotation`/`FindByLabel`/`FindByBookmark`/`ListAnnotations` JSExports) + npm wrapper (`session.findByAnnotation/findByLabel/findByBookmark/listAnnotations`) + new TS types (`AnchorTargetRef`, `DocumentAnnotation`). Tests: `DS180`–`DS187`.
- **`DocxSession.FindByText` / `FindAllByText` / `FindByRegex` / `FindByKind` helpers** (issue #137). Thin wrappers over `Grep` and the `AnchorIndex` for the workflows every consumer was reimplementing. \`FindOptions { IgnoreCase, IgnoreWhitespace, KindFilter, ScopeFilter }\` lets one call cover the common variants. \`IgnoreWhitespace\` flows down to \`Grep\`'s \`WhitespaceMode.Normalize\` so a needle with regular spaces hits NBSP-using text. \`FindByKind\` reads the projection's \`AnchorIndex\` directly (no text scan) for "enumerate every heading in the body." Tests: \`DS160\`–\`DS166\`.
- **\`DocxSession.Grep\` accepts \`WhitespaceMode\` for NBSP-tolerant matching** (issue #136). New \`WhitespaceMode { Preserve (default), Normalize }\` enum + \`whitespace\` parameter on \`Grep\`. In \`Normalize\` mode the match runs against a flat text where U+00A0 (NBSP), U+202F (narrow NBSP), and U+2009 (thin space) are folded to ASCII space; substitutions are 1:1 character-for-character so fragment \`Span\` offsets returned in the \`TextMatch\` still address the original positions. A follow-up \`ReplaceMatch\` lands in the right place even though the match was discovered via normalized text. Plumbed through the WASM bridge (\`GrepOptions.whitespace\` numeric flag) + npm wrapper. Tests: \`DS150\`–\`DS152\`.
- **\`DocxSessionSettings.SmartQuotes\`** (issue #140). When true, \`ReplaceText\` / \`ReplaceTextRange\` / \`ReplaceTextAtSpan\` (and \`ReplaceMatch\` by extension) payloads have ASCII \`"\` and \`'\` converted to typographic curly quotes (U+201C/U+201D and U+2018/U+2019) based on context: open quote at the start of the string, after whitespace, or after an open-bracket-like character; close quote elsewhere. Avoids the cosmetic regression where a Bluth-Co fill landed as \`"foo"\` adjacent to surrounding already-curly \`"foo"\` text. Plumbed through the WASM bridge (\`DocxSessionSettings.smartQuotes\` JSON flag) + npm wrapper. Tests: \`DS170\`–\`DS174\`.
- **`DocxSession.ApplyFormatToSubstring(anchor, substring, op)` + `ApplyFormat(TextMatch, op)`** (issue #138). Substring overload finds the first occurrence of the visible text in the anchor's flat text and converts to a `CharSpan` internally — eliminates the offset-arithmetic trap where an auto-number prefix shifts visible-text indices vs run-text indices. The `TextMatch` overload pairs naturally with `Grep`/`ReplaceMatch` for "format the exact match I just found." Distinct from the existing `ApplyFormat(string, CharSpan?, FormatOp)` to keep `(anchor, null, op)` whole-paragraph calls unambiguous to the C# overload resolver. WASM bridge: `ApplyFormatBySubstring` JSExport. npm wrapper: `session.applyFormatBySubstring(anchor, substring, op)` and `session.applyFormatToMatch(match, op)`. Tests: `DS130`–`DS132`.
- **`WmlToMarkdownConverterSettings.EmptyParagraphs` setting — render-mode toggle for empty paragraphs** (issue #135). New `EmptyParagraphMode` enum: `AnchorOnly` (default — bare `{#p:body:UNID}` line, current behavior), `MarkedEmpty` (appends `∅` sentinel so agents can pattern-match), `Suppress` (drops the paragraph entirely + removes it from `AnchorIndex`). Plumbed through the WASM bridge (`MarkdownProjectionSettingsDto.EmptyParagraphs`) and the npm wrapper (`MarkdownProjectionSettings.emptyParagraphs` + exported `EmptyParagraphMode` enum). Tests: `MD030`–`MD032`.
- **`DocxSession.FindPlaceholders` — typed enumeration of template slots** (issue #142). Built on `Grep` (#143); classifies bracketed regions into three kinds an agent treats differently:
  - `BlankFill` — `[___]` or `$[___]` value slots
  - `AlternativeClause` — `[entire clause text in brackets]` optional clauses to keep/strip
  - `Instruction` — `[insert X]`, `[specify Y]`, `[*italicized hint*]` — drafter hints; the inner text is exposed as `Hint` with surrounding asterisks stripped
  Returns `TemplatePlaceholder` records wrapping the underlying `TextMatch` so the caller has anchor, span, fragment list, and surrounding context for each match without a second pass. `PlaceholderKinds` flag enum lets callers narrow (e.g. just `BlankFill`). The complete template-fill workflow now collapses to: `foreach (var p in session.FindPlaceholders(PlaceholderKinds.BlankFill).OrderByDescending(p => p.Match.Span.Start)) session.ReplaceMatch(p.Match, value);` — the 200-line Bluth-Co fill script replaced by five lines. WASM bridge (`FindPlaceholders` JSExport) + npm wrapper (`session.findPlaceholders()`, `PlaceholderKinds` flag exports). Tests: `DS120`–`DS126`. Architecture: see `docs/architecture/docx_mutation_api.md` (FindPlaceholders section).
- **`DocxSession.ReplaceTextRange` — surgical text replacement that preserves run formatting** (issue #139). Built on `Grep` (#143). Three public surfaces:
  - `ReplaceTextRange(anchorId, find, replace, options?)` — finds every literal occurrence of `find` in the anchor's flat text and replaces each with `replace`. Returns one `EditResult` per attempted match.
  - `ReplaceMatch(TextMatch, replace)` — convenience for `Grep` results.
  - `ReplaceTextAtSpan(anchorId, spanStart, spanLength, replace)` — exact-span variant for the template-fill case where five identical `[___]` placeholders in the same paragraph each need a different value (the spans disambiguate; the literal text would not).
  `ReplaceOptions { IgnoreCase, MaxReplacements }`. Replacement text inherits the formatting of the FIRST run the match spanned; middle/trailing runs keep their `w:rPr` but lose the slice of text the match consumed (so the bold formatting on a phrase that got partially overwritten survives for any surviving text). Matches are applied in reverse document order so multi-match-per-paragraph cases don't invalidate each other's offsets, and the whole call records a single undo snapshot. WASM bridge (`ReplaceTextRange` + `ReplaceTextAtSpan` JSExports) + npm wrapper (`session.replaceTextRange()`, `session.replaceMatch(match, replace)`). Tests: `DS110`–`DS119`. Architecture: see `docs/architecture/docx_mutation_api.md` (ReplaceTextRange section).
- **`DocxSession.Grep` — cross-run text search with run-fragment breakdown** (issue #143). The foundational primitive `FindByText`/`ReplaceTextRange`/`FindRegexSpans` (#137/#139/#142) will build on. Searches the flat text of every paragraph/heading/list-item in scope and returns matches in document order, each with the `<w:r>` runs the match spans plus per-fragment formatting (bold/italic/strike/underline/code/color/hyperlink/runStyle). Lets callers rewrite a match in place while preserving each fragment's formatting — the format-preservation problem that the Bluth-Co smoke-test fill hit when collapsing runs. `Grep` accepts standard `RegexOptions` and a `ProjectionScopes` filter (defaults to body), with configurable surrounding-context length. Shared text-map+offset-map helper at `Docxodus/Internal/RunTextMap.cs` so future search/replace work doesn't reinvent the run walker. Public surface: `TextMatch`, `RunFragment`, `RunFormatting`. WASM bridge (`Grep` JSExport) + npm wrapper (`session.grep(pattern, options?)`). Tests: `DS100`–`DS108`. Architecture: see `docs/architecture/docx_mutation_api.md` (Grep section).
- **`DocxSession` — stateful in-memory DOCX mutation API** — The write-side counterpart to `WmlToMarkdownConverter` for agentic editing pipelines. Spec at `docs/architecture/docx_mutation_api.md`. Mutations are keyed by markdown-projection anchor ids; every method returns a typed `EditResult` envelope (no exceptions across the API boundary). Surface:
  - Lifecycle: `new DocxSession(bytes, settings?)`, `Project()`, `Save()`, `Exists()`, `GetAnchorInfo()`, `Undo()`/`Redo()`, `Dispose()`
  - Tier A (text CRUD): `ReplaceText`, `DeleteBlock`
  - Tier B (structural): `InsertParagraph`, `SplitParagraph`, `MergeParagraphs`
  - Tier C (formatting): `ApplyFormat` (whole-paragraph or `CharSpan`), `SetParagraphStyle`, `SetListLevel`, `RemoveListMembership`
  - Tier D (advanced): `ReplaceCellContent`; `Settings.TrackedChanges = RenderInline` makes mutations land as `w:ins`/`w:del`
  - Raw OOXML escape hatch: `session.Raw.GetXml/InsertXml/ReplaceXml` for content the markdown subset can't express (complex tables, math, content controls); optional `Settings.ValidateRawOps` runs `OpenXmlValidator` post-apply with rollback on failure
  - Bounded snapshot undo/redo (default depth 50) over per-part XML clones
  - Markdown payload parser (`Internal/MarkdownPayloadParser`) accepts the projector-symmetric subset (paragraphs, headings, lists, blockquotes, fenced code; bold/italic/code/strike/hyperlinks, escapes) and rejects out-of-subset syntax with typed `EditErrorCode`s (e.g. `TableInsertNotSupported`, `FootnoteRefNotSupported`)
  - WASM `[JSExport]` bridge at `wasm/DocxodusWasm/DocxSessionBridge.cs` with explicit session handles (no JS-side GC observability)
  - npm wrapper at `npm/src/session.ts` exposing `openDocxSession()` and the `DocxSession` class with `Symbol.dispose` support; full type surface in `npm/src/types.ts` (snake_case `EditErrorCode` union, `EditResult`, `AnchorRef`, `CharSpan`, `FormatOp`, `DocxSessionSettings`)
- **Full `WmlToMarkdownConverter` implementation** — Replaces the v5.5.4 scaffold with the complete anchor-addressed Markdown projection described in `docs/architecture/markdown_projection.md`. Covers:
  - Paragraphs and headings (Heading 1–6 + Title/Subtitle, with `HeadingLevelOffset`)
  - Inline runs: bold, italic, code (rStyle/monospace heuristic), strikethrough, hyperlinks (internal + external), Markdown metacharacter escaping
  - Lists with `ListItemRetriever`-resolved numbering ("1.", "1.2.", "a.", bullet); 2-space indent per level; `ResolveNumbering=false` falls back to "-" markers
  - Tables: GFM pipe tables when the shape is simple (no `gridSpan>1` / `vMerge` / nested tables / oversized cells); opaque fenced ` ```table` blocks otherwise; addressable per-cell via `{#tc:body:UNID}` anchors
  - Multipart scopes: `# Headers`/`## hdrN`, `# Footers`/`## ftrN`, GFM-style `[^fn-XXXX]`/`[^en-XXXX]` footnote and endnote references and definitions, `# Comments` list with author/date
  - Tracked-change modes: `Accept` (default), `RenderInline` (`{+ins+}`/`{-del-}`), `StripDeletions`
  - Per-element anchor index reachable via `MarkdownProjection.AnchorIndex` and `AnchorTarget.Resolve(WordprocessingDocument)`
  - WASM `[JSExport] ConvertWmlToMarkdown` and npm `convertWmlToMarkdown` wrapper with TypeScript enums for `ProjectionScopes`, `AnchorRenderMode`, `TableRenderMode`, `TrackedChangeMode`

### Changed
- **`UnidHelper`** — Extracted the `PtOpenXml.Unid` assignment logic out of `WmlComparer` into an internal shared helper so the same code paths are used by both `WmlComparer` and `WmlToMarkdownConverter`. Added `AssignToSelfAndDescendants(XElement)` overload that assigns a Unid to the root unconditionally — used by `DocxSession` when inserting freshly-built block elements that need to be addressable on the next projection.
- **`DocxSession.MergeParagraphs` now inserts a single-space separator** at the seam when both sides end/start with non-whitespace, so merged sentences no longer jam together (`"First." + "Second."` → `"First. Second."` instead of `"First.Second."`). Behavior change for callers that relied on raw concatenation. Regression test: `DS085_MergeParagraphs_InsertsSeparator_WhenBothEndsAreNonWhitespace`.

### Fixed
- **`DocxSession.DeleteBlock` now accepts footnote/endnote/comment anchors and cleans up their in-body references** (issue #133). \`DeleteBlock(footnoteAnchor)\` previously failed with \`AnchorWrongKind\`; the workaround was \`Raw.ReplaceXml\` on each footnote, which left orphan \`<w:footnoteReference w:id=\"X\"/>\` markers in the body and rendered as broken superscript in Word. The op now removes the definition AND every cross-reference pointing at its id, across every projected part of the package — \`w:footnoteReference\` / \`w:endnoteReference\` for fn/en, plus the \`w:commentReference\` + \`w:commentRangeStart\` + \`w:commentRangeEnd\` triple for comments. Empty wrapper runs (a \`<w:r>\` whose only meaningful child was the removed reference) are also stripped to avoid leaving styled-empty spans. Word-reserved fn/en kinds (\`type=\"separator\"\` and \`type=\"continuationSeparator\"\`) are refused with a typed error. Smoke on the NVCA Model COI: 95 of 97 footnotes stripped (2 separators correctly refused), zero orphan references in body, output 25% smaller than input. Tests: \`DS140\`–\`DS143\`.
- **\`DocumentSnapshot\` now captures every projected part, not just MainDocumentPart.** Required for the cross-part \`DeleteBlock\` (fn/en/cmt) above so undo restores both the definition AND the in-body references in one shot. Also fixes a previously-latent bug where \`Save()\` stripping Unids from non-main parts failed to restore them via the snapshot, leaving subsequent ops in the session unable to resolve anchors in headers/footers/footnotes/etc. No public-API change.
- **Markdown projector now resolves style-inherited numbering on headings, matching `WmlToHtmlConverter`** (issue #141). The projector's `ListNumberResolver` was guarding on inline `w:numPr` and short-circuiting for paragraphs whose numbering came from their style (e.g. an NVCA Heading1 with `<w:pStyle val="Heading1"/>` where the Heading1 style declares numPr). `ListItemRetriever` handles both inline and style-level numPr; removing the guard lets it. The NVCA Model COI's "First Article" / "Second Article" headings now project as `# 1. That the name…` / `# 2. That the Board…` instead of `# That the name…` / `# That the Board…`, lining up with the HTML converter's `1.` / `2.` rendering of the same paragraphs. Regression test: `MD033_HeadingNumberPrefix_ResolvesFromStyleLevelNumPr`.
- **`DocxSession.ReplaceText` no longer doubles auto-numbered heading prefixes.** The markdown projector emits resolved numbering inline (`## Fourth The total number…`) so a numbered heading reads as a human would see it. An agent that echoed the visible heading back as its `ReplaceText` payload caused Word to render the prefix twice (`"Fourth Fourth: …"`) because the auto-number from `w:numPr` was still being applied to the new run text. `ReplaceText` now resolves the paragraph's auto-number via the shared `Internal.ListNumberResolver` and strips a matching leading prefix (plus one optional separator: ASCII space, tab, or NBSP) from the payload before parsing. Idempotent when the prefix isn't present. Regression test: `DS091`/`DS091b`.
- **`DocxSession.Raw.ReplaceXml` no longer reports the same anchor in both `Created` and `Removed`.** The documented `Raw.GetXml → mutate → Raw.ReplaceXml` round-trip preserves Unids, but the prior impl unconditionally put `target.Anchor` in `Removed` and re-added the (same-Unid) element to `Created` — so callers pattern-matching on the lifecycle lists saw a phantom delete-then-recreate. Classification is now by Unid set intersection: overlap → `Modified`, old-only → `Removed`, new-only → `Created`. Regression tests: `DS092` (round-trip preserves Unid → `Modified`) and `DS092b` (fresh XML with new Unids → `Removed`/`Created`).
- **`DocxSession.Save` strips the internal `PtOpenXml:Unid` attribute from every part by default.** The projector assigns a Unid to every descendant of every projected scope (\~14 k attributes on the NVCA Model COI), and the prior impl serialized them all — turning a 148 KB input into a 588 KB output (4× bloat). The attribute is internal to the projector and not in the OOXML schema; stripping it on save is the correct default. The escape hatch for callers that need anchor-id stability across save/reopen is `DocxSessionSettings.PersistAnchorIds = true`. Regression tests: `DS093` (default strips), `DS094` (opt-in preserves).

- **`DocxSession` Tier B/C ops now walk `<w:hyperlink>` / `<w:sdt>` / `<w:fldSimple>` / `<w:smartTag>` containers when computing offsets and iterating runs.** The prior implementation iterated only `Elements(W.r)` (direct paragraph children), which caused four interlocking bugs uncovered by smoke-testing the NVCA Model COI:
  - `SplitParagraph` left hyperlinks stuck to the first half regardless of the split offset (so a split at offset 5 of `"Mix of bold ... [link]."` produced `"Mix olink"` + `"f bold ... ."`). Containers crossing the boundary are now split into two siblings sharing the same `r:id`/attributes.
  - `MergeParagraphs` silently discarded hyperlinks / bookmarks / sdts in the second paragraph (only direct `<w:r>` children were moved before `secondEl.Remove()`). All non-`pPr` children are now moved.
  - `ApplyFormat` skipped runs inside hyperlinks and used `ParagraphText` (direct-runs-only) for span validation, while `GetAnchorInfo.TextPreview` summed descendant text — so an agent computing offsets from the markdown projection got `OffsetOutOfRange` on valid spans. Both now share the descendant-walking `InlineRuns` helper, and hyperlink-internal runs are formatted.
  - `ReplaceText` discarded bookmarkStart/End, comment range markers, perm markers, and proofErr because `RemoveNodes()` cleared everything but `pPr`. These markers are now preserved across the replace (pre-content markers wrap before the new runs, post-content markers after). Regression tests: `DS080`-`DS088`.
- **`DocxSession.PromoteHyperlinkRelationships` dedupes by URL.** Each `ReplaceText`/`InsertParagraph` previously called `AddHyperlinkRelationship` unconditionally, so repeated edits with the same link accumulated orphan rIds in `document.xml.rels`. Same-URL ops now reuse the existing relationship. Regression test: `DS089`.
- **`DocxSession.InsertParagraph` reports a `Created[i].Kind` consistent with the next projection.** A bullet payload (`- item`) previously returned `Kind = "li"` even when no `<w:numPr>` was injected, so the returned anchor id (`li:body:…`) never appeared in the projection (`p:body:…` did). Bullet/ordered-item payloads now inherit `<w:numPr>` from a nearest-sibling list item when one exists; the reported kind is computed via the same predicate the projector uses. Regression test: `DS090`.
- **`WmlToMarkdownConverter` projection fidelity** — Surfaced and fixed during smoketesting against the NVCA Model Certificate of Incorporation (a heading-heavy legal document):
  - **Numbered headings keep their auto-number.** A `Heading{1..9}` paragraph that also carries `w:numPr` (the standard legal-doc convention for `FIRST: …` / `1.1 …` clause numbering) now prepends the resolved number to the heading text. Previously the auto-number was silently dropped, leaving headings like `## : The name of this corporation is …`.
  - **`w:sectPr` emits `---` thematic break with anchor.** Section breaks inside a paragraph's `pPr` now produce a `{#sec:scope:UNID}\n---` pair so callers can navigate sections; the trailing top-level `sectPr` (metadata only) is still suppressed in output but registered in `AnchorIndex` for editing.
  - **Inter-scope `---` separators.** A horizontal rule is emitted between adjacent non-empty scope sections (`# Document` / `# Headers` / `# Footers` / `# Footnotes` / `# Endnotes` / `# Comments`) so downstream parsers can split per scope without inspecting heading text.
  - **Heading7-9 preserve depth.** Word styles `Heading7`/`8`/`9` now emit 7/8/9 hashes instead of being silently clamped to `######`. Strict CommonMark renderers will treat 7+ hashes as literal text; LLM consumers and structured parsers can recover the original outline depth.
  - **Empty header/footer scopes are suppressed.** DOCX files commonly declare 6+ header/footer parts for first-page/even-page/default variants and leave the unused ones blank; the projection no longer emits `## hdrN` titles for scopes whose only content is whitespace.
  - **Anchor-only paragraph lines no longer carry a trailing space.** Empty paragraphs (visual spacers in Word) now render as `{#p:body:UNID}\n` instead of `{#p:body:UNID} \n`.

## [5.5.4] - 2026-05-24

### Fixed
- **NullReferenceException in `FindIndexOfNextParaMark` with body-level bookmarks (#124, thanks @papyria)** — `FindIndexOfNextParaMark` assumed all elements in the comparison-unit array were `ComparisonUnitWord`, but documents with `bookmarkStart`/`bookmarkEnd` as direct children of `w:body` produce other `ComparisonUnit` types. Now handles any `ComparisonUnit` with `Contents` (including `ComparisonUnitGroup`) and adds a null guard for the `LastOrDefault()` call.

### Added
- **`WmlToMarkdownConverter` scaffold (#127)** — Public surface for an anchor-addressed markdown projection of Word documents. `Convert(WmlDocument, WmlToMarkdownConverterSettings)` / `Convert(WordprocessingDocument, ...)` return a `MarkdownProjection` (markdown text + anchor index) with anchors of the form `{#kind:scope:unid}` derived from Docxodus' existing Unid system. **Scaffold only** — projection logic ships in subsequent phases. See `docs/architecture/markdown_projection.md` for the spec.

### Maintenance
- **Bump `Microsoft.NET.Test.Sdk` from 18.4.0 to 18.5.1 (#125)**

## [Unreleased] - .NET 8 / Open XML SDK 3.x Migration

### Fixed (npm)
- **TypeScript subpath exports not resolving under `moduleResolution: "node"` (Issue #113)** - Added `typesVersions` fallback to npm package.json so `docxodus/react` and `docxodus/worker` subpath imports resolve types correctly under all TypeScript module resolution modes. Also reordered export conditions to put `types` before `import` per TypeScript requirements.

### Added
- **Incremental annotation overlay API (Issue #106)** - Decouple HTML conversion from annotation projection to avoid full WASM re-conversion
  - `ProjectAnnotationsOntoHtml()` - Project a full annotation set onto already-converted HTML
  - `AddAnnotationToHtml()` - Add a single annotation to existing HTML without re-converting the document
  - `RemoveAnnotationFromHtml()` - Remove a single annotation by ID, unwrapping spans back to plain text
  - `GenerateVisibilityCss()` - Generate CSS to hide/show annotations by label ID for instant toggling
  - `GenerateAnnotationCssString()` - Generate annotation CSS separately for independent management
  - All methods available in .NET, WASM (JSExport), and npm TypeScript wrapper
  - CSS-based label filtering enables responsive toggle without any re-rendering

### Fixed
- **NullReferenceException in FindIndexOfNextParaMark when comparing documents with body-level bookmarks** - `FindIndexOfNextParaMark` assumed all elements in the comparison unit array were `ComparisonUnitWord`, but documents with `bookmarkStart`/`bookmarkEnd` as direct children of `w:body` produce other `ComparisonUnit` types. Now handles any `ComparisonUnit` with `Contents` (including `ComparisonUnitGroup`) and adds a null guard for the `LastOrDefault()` call.
- **Paginated rendering: text clipped at page bottom + inconsistent paragraph spacing (Issue #114)**
  - Fixed `lineRule` default handling: when `w:lineRule` is absent but `w:line` is present, treat as "auto" per OOXML spec (ISO/IEC 29500). Previously the line value was ignored, causing accumulated line-height mismatches that clipped the last line on pages.
  - Fixed `contextualSpacing` handling: now suppresses both `spacingAfter` (margin-bottom) AND `spacingBefore` (margin-top) for consecutive same-style paragraphs. Previously only `spacingAfter` was suppressed, leaving inconsistent inter-paragraph gaps.
  - Fixed pagination engine bottom margin over-reservation: the last block's bottom margin is no longer counted against page space since it's invisible (clipped by `overflow: hidden`). This prevents premature page breaks where content would have been visible.
- **Annotation projection fails on sanitized HTML (Issue #110)** - `ProjectAnnotationsOntoHtml`, `AddAnnotationToHtml`, and `RemoveAnnotationFromHtml` now handle HTML fragments with multiple root elements (e.g., DOMPurify-sanitized output) and HTML named entities (`&nbsp;`, `&ndash;`, etc.)
  - Root cause: `XElement.Parse()` requires valid XML with a single root element; sanitized HTML strips `<html>`/`<body>` wrappers leaving multiple roots
  - Fix: Auto-wraps multi-root HTML in a synthetic container for parsing, unwraps on serialization; replaces common HTML entities with numeric XML equivalents
- **Table container missing top margin (Issue #108)** - Tables preceded by paragraphs with no after-spacing now get a default `margin-top: 7.5pt` for visual separation
  - Also handles floating table spacing from `w:tblpPr` (`topFromText`/`bottomFromText` attributes)
  - Tables preceded by paragraphs with explicit after-spacing correctly skip the default margin
- **Move markup Word compatibility (Issue #96)** - Documents with move operations no longer cause Word "unreadable content" warnings
  - Root cause: `FixUpRevMarkIds()` was overwriting IDs of `w:del`/`w:ins` after `FixUpRevisionIds()` had already assigned unique IDs, causing collisions with move element IDs
  - Fix: Removed redundant `FixUpRevMarkIds()` call - `FixUpRevisionIds()` already handles all revision element IDs correctly
  - Added `SimplifyMoveMarkup` setting to optionally convert move markup to simple `w:del`/`w:ins` if desired
  - Added comprehensive ID uniqueness tests to prevent regression
  - `DetectMoves` now defaults to `true` (move detection is safe to use)
- **Footnote/endnote numbering** - Fixed footnotes and endnotes displaying raw XML IDs instead of sequential display numbers
  - Per ECMA-376, `w:id` is a reference identifier, not the display number
  - Added `FootnoteNumberingTracker` class to scan document and build XML ID → display number mapping
  - Footnotes/endnotes now render with sequential numbers (1, 2, 3...) based on document order
  - Also fixed footnote ordering in the footnotes section to match document order
  - Updated both regular and paginated rendering modes
  - See `docs/ooxml_corner_cases.md` for detailed documentation
- **Legal numbering continuation pattern** - Fixed incorrect multi-level list numbering when items continue a flat sequence at different indentation levels
  - Documents with items like 1., 2., 3. at level 0 followed by item at level 1 (with start=4) now render as "4." instead of "3.4"
  - Added "continuation pattern" detection in `ListItemRetriever.cs` that recognizes when a deeper-level item continues a flat list
  - When detected, uses level 0's format string, run properties, and paragraph properties with the current counter value
  - Fixes underline appearing on continuation items when level 1's rPr has underline but level 0's doesn't
  - Fixes tab/indentation spacing to use level 0's tab stops and indentation for consistency
  - Updated `FormattingAssembler.cs` to use `GetEffectiveLevel()` in paragraph property stack and annotation functions
  - See `docs/ooxml_corner_cases.md` for detailed documentation of this edge case
- **Tab width calculation** re-enabled in `WmlToHtmlConverter` for proper tab stop positioning
  - Previously disabled due to Azure font measurement failures; now uses estimation fallback
  - `MetricsGetter._getTextWidth()` returns character-based estimation when SkiaSharp measurement fails
  - Estimation formula: `charWidth = fontSize * 0.6 / 2` per character (same as WASM builds)
  - Tab positioning now properly accounts for preceding text width
  - Works in Azure, WASM, and environments without fonts installed
  - Added Playwright visual tests for tab rendering verification
- **Thread-safety issues** in `WmlToHtmlConverter` and `FontFamilyHelper` that could cause corruption during concurrent document conversions
  - `ShadeCache` in `WmlToHtmlConverter` now uses `ConcurrentDictionary` for thread-safe shade color caching
  - `FontFamilyHelper._unknownFonts` now uses `ConcurrentDictionary` for thread-safe font tracking
  - `FontFamilyHelper.KnownFamilies` now uses `Lazy<T>` for thread-safe lazy initialization
  - Added `WmlToHtmlConverter.ClearShadeCache()` and `FontFamilyHelper.ClearUnknownFontsCache()` methods for memory management in long-running processes

### Breaking Changes
- **Target Framework**: Changed from net45/net46/netstandard2.0 to .NET 8.0
- **Open XML SDK**: Upgraded from 2.8.1 to 3.2.0
- **Graphics Library**: Replaced System.Drawing with SkiaSharp 2.88.9

### Added
- **Table Width DXA Support** - Tables with DXA (twips) widths now render correctly
  - Previously, only percentage widths were handled; DXA widths were ignored
  - Tables with `w:tblW[@w:type="dxa"]` now render with proper `width: XXpt` CSS
  - Conversion uses standard formula: `dxa / 20 = points`
  - Addresses converter gaps #1 (Table Width Calculation)
- **Borderless Table Detection** - Tables without borders now get semantic markup
  - Tables with `w:tblBorders` set to `nil`/`none` or missing get `data-borderless="true"` attribute
  - Useful for identifying layout tables vs data tables
  - Enables CSS-based styling for signature blocks and multi-column layouts
  - Addresses converter gaps #3 (Borderless Table Detection)
- **Document Language Attribute** - HTML output now includes `lang` attribute for improved accessibility
  - New `DocumentLanguage` setting to manually override the language (default: auto-detect)
  - `<html>` element now includes `lang` attribute (e.g., `<html lang="en-US">`)
  - Language is auto-detected from:
    1. `w:themeFontLang` in document settings
    2. Default paragraph style's `w:rPr/w:lang`
    3. Falls back to "en-US"
  - Foreign text spans get `lang` attribute when different from document default
  - Improves screen reader pronunciation and browser font selection
  - Addresses converter gaps #10 (Document Language Attribute) and #11 (Foreign Text Spans)
- **Improved Font Fallback** - Unknown fonts now get appropriate generic fallback, and CJK text gets language-specific font chains
  - Unknown fonts are classified by name patterns and get proper fallback:
    - Fonts with "sans" pattern → `font-family: 'FontName', sans-serif`
    - Fonts with "mono", "code", "courier" patterns → `font-family: 'FontName', monospace`
    - Other fonts default to serif fallback
  - Fixed Courier New and Lucida Console to include `monospace` fallback (was missing)
  - CJK (Chinese, Japanese, Korean) text gets language-specific font fallback chains:
    - Japanese (ja-JP): `'Noto Serif CJK JP', 'Yu Mincho', 'MS Mincho', ...`
    - Simplified Chinese (zh-hans): `'Noto Serif CJK SC', 'Microsoft YaHei', 'SimSun', ...`
    - Traditional Chinese (zh-hant): `'Noto Serif CJK TC', 'Microsoft JhengHei', 'PMingLiU', ...`
    - Korean (ko): `'Noto Serif CJK KR', 'Malgun Gothic', 'Batang', ...`
  - Addresses converter gaps #13 (Limited Font Fallback) and #14 (No CJK Font-Family Fallback Chain)
- **Theme Color Resolution** - Document theme colors are now resolved to actual RGB values
  - New `ResolveThemeColors` setting (default: true) enables theme color resolution
  - Reads color scheme from `theme1.xml` (`a:clrScheme` element)
  - Supports all 12 theme colors: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink
  - Applies `w:themeTint` (lighten toward white) and `w:themeShade` (darken toward black) modifiers
  - Resolves `w:themeColor` in run colors, paragraph shading, cell shading, and fills
  - Falls back to explicit color value if theme color not found
  - Addresses converter gap #6 (Theme Colors Not Resolved)
- **@page CSS Rule** - Optional CSS `@page` rule generation for print stylesheets
  - New `GeneratePageCss` setting (default: false) enables `@page` rule generation
  - Reads page dimensions from `w:sectPr/w:pgSz` and margins from `w:sectPr/w:pgMar`
  - Generates CSS `@page { size: Xin Yin; margin: ... }` rules
  - Supports US Letter, A4, and custom page sizes with proper inch conversions
  - Useful for print stylesheets and PDF generation
  - Addresses converter gap #1 (No Page/Document Setup CSS)
- **Unsupported Content Placeholders** - Visual indicators for content that cannot be fully converted to HTML
  - New `RenderUnsupportedContentPlaceholders` setting (default: false for backward compatibility)
  - Supports these unsupported content types:
    - **WMF/EMF images**: Legacy Windows Metafile formats display `[WMF IMAGE]` / `[EMF IMAGE]`
    - **SVG images**: Scalable Vector Graphics display `[SVG IMAGE]`
    - **Math equations (OMML)**: Office Math Markup displays `[MATH]`
    - **Form fields**: Checkboxes, text inputs, dropdowns display `[CHECKBOX]`, `[TEXT INPUT]`, `[DROPDOWN]`
    - **Ruby annotations**: East Asian text annotations display base text with `[RUBY]` marker
  - Placeholders are styled with CSS (color-coded by type) and include:
    - `data-content-type` attribute for the content type
    - `data-element-name` attribute for the XML element name
    - `title` attribute with descriptive tooltip
  - New TypeScript enum `UnsupportedContentType` for type-safe placeholder identification
  - See `docs/architecture/unsupported_content_placeholders.md` for full documentation
- **External Annotation System** (Issue #57) - Store annotations externally without modifying the DOCX file
  - New `ExternalAnnotationSet` type extends `OpenContractDocExport` with document binding:
    - `documentId`: Unique identifier for the source document
    - `documentHash`: SHA256 hash for integrity validation
    - `createdAt`, `updatedAt`: ISO 8601 timestamps
    - `textLabels`, `docLabelDefinitions`: Label definitions keyed by ID
  - `ExternalAnnotationManager` static class provides core functionality:
    - `ComputeDocumentHash()`: SHA256 hash of document bytes
    - `CreateAnnotationSet()`: Create annotation set from document (wraps OpenContractExporter)
    - `CreateAnnotation()`: Create annotation from character offsets
    - `CreateAnnotationFromSearch()`: Create annotation by text search with occurrence index
    - `FindTextOccurrences()`: Find all occurrences of text in document
    - `Validate()`: Validate annotations against document (hash check + text verification)
    - `SerializeToJson()` / `DeserializeFromJson()`: JSON serialization
  - `ExternalAnnotationProjector` for HTML projection:
    - `ProjectAnnotations()`: Post-process HTML to wrap annotated text with styled spans
    - `ConvertWithAnnotations()`: Combined conversion + projection
    - Supports annotation labels (Above, Inline, Tooltip, None modes)
    - CSS generation with customizable class prefix
  - TypeScript/npm wrapper functions:
    - `computeDocumentHash()`: Get document hash for validation
    - `createExternalAnnotationSet()`: Create annotation set from DOCX
    - `validateExternalAnnotations()`: Validate annotations against document
    - `convertDocxToHtmlWithExternalAnnotations()`: Convert with annotations projected
    - `searchTextOffsets()`: Search for text occurrences in document
    - `createAnnotation()`, `createAnnotationFromSearch()`, `findTextOccurrences()`: Client-side helpers
  - Full type definitions: `AnnotationLabel`, `ExternalAnnotationSet`, `ExternalAnnotationValidationResult`, etc.
  - 21 unit tests covering hash computation, annotation creation, validation, serialization, and projection
- **OpenContracts Export Format** (Issue #56) - Export documents to OpenContracts format for interoperability
  - New `OpenContractExporter.Export()` method for complete document export:
    - `title`: Document title from core properties
    - `content`: Complete document text (paragraphs, tables, headers, footers, footnotes, endnotes)
    - `description`: Optional document description
    - `pageCount`: Estimated page count
    - `pawlsFileContent`: PAWLS-format page layout with token positions
    - `docLabels`: Document-level labels
    - `labelledText`: Annotations including structural elements (sections, paragraphs, tables)
    - `relationships`: Parent-child relationships between annotations
  - Full text extraction ensures 100% text coverage:
    - Main body paragraphs and tables
    - Nested tables
    - Headers and footers
    - Footnotes and endnotes
    - Content controls (structured document tags)
  - PAWLS (Page-Aware Layout Segmentation) format for layout data:
    - Page boundary information (width, height, index)
    - Token positions (x, y, width, height, text)
    - Supports annotation targeting by character offset
  - Structural annotations automatically generated:
    - Section annotations with page dimensions
    - Paragraph annotations with text spans
    - Table annotations with content ranges
    - Parent-child relationships (section contains paragraphs)
  - TypeScript API: `exportToOpenContract()` function with full type definitions
  - WASM export: `DocumentConverter.ExportToOpenContract()`
  - Compatible with OpenContracts ecosystem for document analysis
  - **New CLI tool: `docx2oc`** - Command-line tool for OpenContracts export
    - Usage: `docx2oc <input.docx> [output.json]`
    - Default output: same filename with `.oc` extension
    - Installable as .NET tool: `dotnet tool install --global Docx2OC`
- **ReadyToRun and AOT Compilation** - Performance optimizations to reduce cold-start times
  - .NET library: Added `PublishReadyToRun` for pre-compiled native code during publish
  - WASM: Added `RunAOTCompilation` for Release builds to pre-compile IL to WebAssembly
  - Eliminates JIT warmup overhead (~180ms savings on first conversion in .NET)
  - Provides consistent performance with no JIT variance in WASM
- **Lightweight WASM Image Handling** - Images are now embedded as base64 data URIs without SkiaSharp native library
  - Removed SkiaSharp native WASM dependency (~15MB+ savings in bundle size when native lib excluded)
  - Images are passed through directly from DOCX using `ImageBytes` property
  - Dimensions come from document markup (EMUs), not image decoding
  - Browser natively decodes image formats (PNG, JPEG, GIF, etc.)
  - Fallback handling: If SkiaSharp decode fails, images still work via raw bytes
  - Added image handling tests for documents with embedded and hyperlinked images
- **Frame Yielding for UI Responsiveness** (Issue #44 Phase 1) - WASM operations now yield to the browser before heavy work begins
  - All async functions in the npm wrapper (`convertDocxToHtml`, `compareDocuments`, `compareDocumentsToHtml`, `getRevisions`, `addAnnotation`, `addAnnotationWithTarget`, `getDocumentStructure`) automatically yield using double-`requestAnimationFrame` pattern
  - This allows React state updates (loading spinners, progress indicators) to paint before blocking WASM execution
  - Transparent to consumers - no API changes required
  - Gracefully skipped in non-browser environments (Node.js, SSR)
- **Web Worker Support for Non-blocking Operations** (Issue #44 Phase 2) - Fully non-blocking WASM execution via Web Workers
  - New `docxodus/worker` export provides worker-based API: `import { createWorkerDocxodus } from 'docxodus/worker'`
  - Worker API mirrors main API: `convertDocxToHtml`, `compareDocuments`, `compareDocumentsToHtml`, `getRevisions`, `getVersion`
  - Main thread remains fully responsive during WASM execution - animations continue, user interactions work
  - Zero-copy transfer of document bytes via Transferable for optimal performance
  - Worker can be terminated when no longer needed
- **Document Metadata API for Lazy Loading** (Issue #44 Phase 3) - Fast metadata extraction without full HTML rendering
  - New `getDocumentMetadata()` function returns document structure information:
    - `sections`: Array of section metadata with page dimensions and content ranges
    - `totalParagraphs`, `totalTables`: Document-wide content counts
    - `hasFootnotes`, `hasEndnotes`, `hasComments`, `hasTrackedChanges`: Feature detection
    - `estimatedPageCount`: Heuristic-based page count estimation
  - Section metadata includes:
    - Page dimensions: `pageWidthPt`, `pageHeightPt`, `marginTopPt`, etc. (all values in points, 1pt = 1/72 inch)
    - Content area: `contentWidthPt`, `contentHeightPt`
    - Header/footer heights: `headerPt`, `footerPt`
    - Content tracking: `paragraphCount`, `tableCount`, `startParagraphIndex`, `endParagraphIndex`
    - Header/footer presence: `hasHeader`, `hasFooter`, `hasFirstPageHeader`, `hasEvenPageHeader`, etc.
  - Available in main API, worker API, and raw WASM: `DocumentConverter.GetDocumentMetadata()`
  - Enables efficient lazy loading for paginated document viewing
  - Security: Maximum document size limit of 100MB to prevent memory exhaustion
  - Graceful handling of malformed documents and invalid header/footer references
  - Known limitation: Section breaks inside tables or text boxes are not detected (see #51)
- **Page Range Rendering for Virtual Scrolling** (Issue #31 Phase 4) - Render specific page ranges for lazy loading
  - New `RenderPageRange()` method in `WmlToHtmlConverter` renders only specified pages
  - Page-to-block mapping uses heuristic-based estimation (paragraphs and tables per page)
  - HTML output includes pagination metadata via data attributes:
    - `data-start-page`, `data-end-page`: Requested page range
    - `data-total-pages`: Total estimated pages in document
    - `data-start-block`, `data-end-block`: Block index range for rendered content
    - `data-block-index`: Per-element block indices for tracking
  - WASM exports: `DocumentConverter.RenderPageRange()`, `DocumentConverter.RenderPageRangeFull()`
  - TypeScript wrapper: `renderPageRange()` with full options support
  - Worker proxy support: `WorkerDocxodus.renderPageRange()` for non-blocking execution
  - React components for virtual scrolling:
    - `useVirtualPagination` hook: Manages viewport-aware page loading with IntersectionObserver
    - `VirtualPaginatedDocument` component: Auto-renders visible pages plus configurable buffer
  - All existing converter options supported (tracked changes, comments, headers/footers, etc.)
  - Graceful handling of out-of-bounds page requests (internally clamped to valid range)
- **Custom Annotations** - Full support for adding, removing, and rendering custom annotations on DOCX documents
  - `AnnotationManager` class for programmatic annotation CRUD operations:
    - `AddAnnotation()`: Add annotation by text search or paragraph range
    - `RemoveAnnotation()`: Remove annotation by ID
    - `GetAnnotations()`: Retrieve all annotations from a document
    - `GetAnnotation()`: Get a specific annotation by ID
    - `HasAnnotations()`: Check if document has any annotations
  - `DocumentAnnotation` class with properties:
    - `Id`: Unique annotation identifier
    - `LabelId`: Category/type identifier for grouping
    - `Label`: Human-readable label text
    - `Color`: Highlight color in hex format (e.g., "#FFEB3B")
    - `Author`: Optional author name
    - `Created`: Optional creation timestamp
    - `Metadata`: Custom key-value pairs
  - `AnnotationRange` class for specifying annotation targets:
    - `FromSearch(text, occurrence)`: Find text by search
    - `FromParagraphs(start, end)`: Span paragraph indices
  - **Document Structure API** for element-based annotation targeting:
    - `DocumentStructureAnalyzer.Analyze()`: Returns navigable tree of document elements
    - `DocumentElement` class with path-based IDs (e.g., `doc/p-0`, `doc/tbl-0/tr-1/tc-2`)
    - Supported element types: `Document`, `Paragraph`, `Run`, `Table`, `TableRow`, `TableCell`, `TableColumn`, `Hyperlink`, `Image`
    - `TableColumnInfo` for virtual column elements (columns aren't real OOXML elements)
  - `AnnotationTarget` class with flexible targeting modes:
    - `Element(elementId)`: Target by element ID from structure analysis
    - `Paragraph(index)`, `ParagraphRange(start, end)`: Target by paragraph index
    - `Run(paragraphIndex, runIndex)`: Target specific run
    - `Table(index)`, `TableRow(tableIndex, rowIndex)`: Target tables/rows
    - `TableCell(tableIndex, rowIndex, cellIndex)`: Target specific cell
    - `TableColumn(tableIndex, columnIndex)`: Metadata-only column annotation
    - `TextSearch(text, occurrence)`: Search text globally
    - `SearchInElement(elementId, text, occurrence)`: Search within specific element
  - WASM methods: `GetDocumentStructure()`, `AddAnnotationWithTarget()`
  - TypeScript helper functions: `findElementById()`, `findElementsByType()`, `getParagraphs()`, `getTables()`, `getTableColumns()`
  - TypeScript targeting factories: `targetElement()`, `targetParagraph()`, `targetTableCell()`, etc.
  - React `useDocumentStructure` hook with structure navigation helpers
  - Annotations stored as Custom XML Part in DOCX (non-destructive)
  - Bookmark-based text range marking for precise positioning
  - HTML rendering with configurable label modes:
    - `AnnotationLabelMode.Above`: Floating label above highlight
    - `AnnotationLabelMode.Inline`: Label at start of highlight
    - `AnnotationLabelMode.Tooltip`: Label shown on hover
    - `AnnotationLabelMode.None`: Highlight only, no label
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderAnnotations`: Enable/disable annotation rendering
    - `AnnotationLabelMode`: Select label display mode
    - `AnnotationCssClassPrefix`: Customize CSS class names (default: "annot-")
    - `IncludeAnnotationMetadata`: Include metadata in HTML data attributes
  - WASM/npm support:
    - `getAnnotations()`, `addAnnotation()`, `removeAnnotation()`, `hasAnnotations()` functions
    - `Annotation`, `AddAnnotationRequest`, `AddAnnotationResponse`, `RemoveAnnotationResponse` types
    - `AnnotationLabelMode` enum
    - `ConversionOptions` extended with annotation rendering options
  - React support:
    - `useAnnotations` hook for annotation state management
    - `AnnotatedDocument` component with click/hover event handling
    - `useDocxodus` hook extended with annotation methods
  - 20 .NET unit tests and 21 Playwright browser tests for full coverage (including 11 for element-based targeting)
- **Comment Rendering in HTML Converter** - Full support for rendering Word document comments in HTML output
  - `CommentRenderMode` enum with three rendering modes:
    - `EndnoteStyle` (default): Comments rendered at end of document with bidirectional anchor links
    - `Inline`: Comments rendered as tooltips with `title` and `data-comment` attributes
    - `Margin`: Comments positioned in a flexbox-based margin column alongside content, with author/date headers and back-reference links
  - New settings in `WmlToHtmlConverterSettings`:
    - `RenderComments`: Enable/disable comment rendering
    - `CommentRenderMode`: Select rendering mode
    - `CommentCssClassPrefix`: Customize CSS class names (default: "comment-")
    - `IncludeCommentMetadata`: Include author/date in HTML output
  - Comment highlighting with configurable CSS classes
  - Full comment metadata support (author, date, initials)
  - Margin mode includes print-friendly CSS media queries
  - WASM/npm support via `commentRenderMode` parameter and TypeScript `CommentRenderMode` enum
- **WebAssembly NPM Package** (`docxodus`) - Browser-based document comparison and HTML conversion
  - `wasm/DocxodusWasm/` - .NET 8 WASM project with JSExport methods
  - `npm/` - TypeScript wrapper with React hooks
  - Full document comparison (redlining) support in the browser
  - DOCX to HTML conversion
  - React hooks: `useDocxodus`, `useConversion`, `useComparison`
  - Build script: `scripts/build-wasm.sh`
- **Native Move Markup in WmlComparer** - Produces Word-native move tracking markup (`w:moveFrom`/`w:moveTo`)
  - Compared documents now contain proper OpenXML move elements, not just `w:del`/`w:ins`
  - Move pairs linked via `w:name` attribute for Word compatibility
  - Range markers (`w:moveFromRangeStart`/`w:moveFromRangeEnd`, `w:moveToRangeStart`/`w:moveToRangeEnd`) properly paired
  - Microsoft Word shows moves in "Track Changes" panel as relocated content
  - New `Moved` value in `WmlComparerRevisionType` enum
  - New properties on `WmlComparerRevision`: `MoveGroupId` (links source/destination), `IsMoveSource` (true=from, false=to)
  - New settings in `WmlComparerSettings`:
    - `DetectMoves`: Enable/disable move detection (default: true)
    - `MoveSimilarityThreshold`: Jaccard similarity threshold 0.0-1.0 (default: 0.8)
    - `MoveMinimumWordCount`: Minimum words to consider for move (default: 3)
  - Uses word-level Jaccard similarity for accurate matching
  - Respects `CaseInsensitive` setting for similarity comparison
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.Moved` enum value
    - `isMove()`, `isMoveSource()`, `isMoveDestination()` type guards
    - `findMovePair()` function to find linked move revisions
    - `moveGroupId` and `isMoveSource` properties on `Revision` interface
- **Format Change Detection in WmlComparer** - Detects and tracks formatting-only changes (`w:rPrChange`)
  - When text content is identical but formatting changes (bold, italic, font size, etc.), produces native Word format change markup
  - Compared documents now contain `w:rPrChange` elements that Microsoft Word recognizes in Track Changes
  - New `FormatChanged` value in `WmlComparerRevisionType` enum
  - New `FormatChange` property on `WmlComparerRevision` with:
    - `OldProperties`: Dictionary of original formatting properties
    - `NewProperties`: Dictionary of new formatting properties
    - `ChangedPropertyNames`: List of what changed (e.g., "bold", "italic", "fontSize")
  - New setting in `WmlComparerSettings`:
    - `DetectFormatChanges`: Enable/disable format change detection (default: true)
  - Full WASM/npm support with new TypeScript helpers:
    - `RevisionType.FormatChanged` enum value
    - `isFormatChange()` type guard
    - `FormatChangeDetails` interface with `oldProperties`, `newProperties`, `changedPropertyNames`
    - `formatChange` property on `Revision` interface
- **Improved Revision API** - Better TypeScript support for the `getRevisions()` API
  - `RevisionType` enum with `Inserted`, `Deleted`, and `Moved` values for type-safe comparisons
  - `isInsertion()`, `isDeletion()`, `isMove()`, `isMoveSource()`, `isMoveDestination()` helper functions
  - `findMovePair()` function to find the matching revision for a move
  - Comprehensive JSDoc documentation on the `Revision` interface
  - All types are properly exported from the package
- **Paginated Headers and Footers** - Headers/footers now render correctly with pagination enabled
  - When both `RenderHeadersAndFooters` and `RenderPagination=Paginated` are enabled, headers and footers appear on each page
  - Per-section header/footer support with section index tracking
  - First page headers/footers supported (when `w:titlePg` is set in document)
  - Even page headers/footers supported for different odd/even page layouts
  - Headers/footers rendered into hidden registry for client-side cloning per-page
  - New data attributes: `data-header-height`, `data-footer-height` on section elements
  - TypeScript `PageDimensions` interface extended with `headerHeight` and `footerHeight`
  - CSS classes `.page-header` and `.page-footer` for positioning within page boxes
  - Automatic hiding of system page number when document has footer content
  - See `docs/architecture/paginated_headers_footers.md` for full architecture details
- **Per-page Footnote Rendering** - Footnotes now appear at the bottom of each page where they are referenced
  - When `RenderFootnotesAndEndnotes=true` with `RenderPagination=Paginated`, footnotes are distributed per-page
  - Footnote registry stores footnotes in a hidden container for client-side distribution
  - `data-footnote-id` attributes added to footnote references for tracking
  - Single-pass, forward-only pagination algorithm (lazy-loading compatible)
  - Pagination engine measures footnote space and includes it in page layout calculations
  - Footnotes render with separator line (`<hr>`) above them
  - **Footnote continuation**: Long footnotes that don't fit on a page are split at paragraph boundaries and continue on subsequent pages (matching Word/Office behavior)
  - **Dynamic footnote area expansion**: Footnote area can expand upward into body content space (up to 60% of page height) to fit more footnote content before splitting, reducing wasted space
  - Endnotes remain at document end (not per-page) - traditional behavior preserved
  - New TypeScript methods: `parseFootnoteRegistry()`, `extractFootnoteRefs()`, `measureFootnotesHeight()`, `addPageFootnotes()`, `splitFootnoteToFit()`, `measureContinuationHeight()`
  - New TypeScript interfaces: `FootnoteContinuation`, `PartialFootnote`
  - New TypeScript constants: `MAX_FOOTNOTE_AREA_RATIO` (0.6), `MIN_BODY_CONTENT_HEIGHT` (72pt)
  - New CSS classes: `.page-footnotes`, `.footnote-item`, `.footnote-number`, `.footnote-content`, `.footnote-continuation`
- `SkiaSharpHelpers.cs` - Color utilities for SkiaSharp compatibility
- `GetPackage()` extension method in `PtOpenXmlUtil.cs` for SDK 3.x Package access
- `SkiaSharp.NativeAssets.Linux.NoDependencies` package for Linux runtime support

### Fixed
- **React hooks loading state not rendering before WASM blocks** (Issue #45) - Fixed `isConverting`/`isComparing`/`isLoading` states in React hooks not painting before WASM execution blocks the main thread. Added `requestAnimationFrame` yielding after state updates in:
  - `useConversion`: `convert()` function
  - `useComparison`: `compare()` and `compareToHtml()` functions
  - `useAnnotations`: `reload()`, `add()`, and `remove()` functions
  - `useDocumentStructure`: `reload()` function

- **Header/footer positioning in paginated mode** - Fixed headers and footers overlapping with body content. Headers now properly constrain to the top margin area (`height: marginTop`) and footers constrain to the bottom margin area (`height: marginBottom`). Uses flexbox layout for proper content alignment within constrained areas.

- **DocumentBuilder relationship copying** - Fixed bug where relationship IDs from source documents could incorrectly match existing IDs in target header/footer parts when using InsertId functionality. This caused validation errors like "The relationship 'rIdX' referenced by attribute 'r:embed' does not exist."
  - Removed flawed early-return optimization in `CopyRelatedImage()` that skipped processing when target part had matching relationship ID
  - Fixed diagram relationship handling (`R.dm`, `R.lo`, `R.qs`, `R.cs` attributes) to properly copy parts from source documents
  - Fixed chart and user shape relationship handling
  - Fixed OLE object relationship handling
  - Fixed external relationship attribute update to use correct attribute name parameter

- **SpreadsheetWriter date handling** - Fixed date cells being written with invalid ISO 8601 string format. Dates are now properly converted to Excel serial date numbers (days since December 30, 1899) which is required for transitional OOXML format.

- **WmlComparer null Unid handling** - Fixed null reference exceptions when comparing documents with elements lacking Unid attributes.

- **WmlComparer footnote/endnote comparison** (6 tests: WC-1660, WC-1670, WC-1710, WC-1720, WC-1750, WC-1760) - Fixed `AssignUnidToAllElements` to assign Unid to footnote/endnote elements themselves, enabling proper reconstruction of multi-paragraph footnotes/endnotes by `CoalesceRecurse`.

- **WmlComparer table row comparison** (1 test: WC-1500) - Added LCS-based row matching (`ApplyLcsToTableRows`) for large tables (7+ rows) when content differs significantly, preventing cascading false differences from insertions/deletions in the middle of tables.

- **WASM CDN loading CORS issue** - Fixed cross-origin loading failures when WASM files are served from CDNs (jsDelivr, unpkg). The .NET WASM runtime uses `credentials:"same-origin"` for fetch requests, which conflicts with CDN's `Access-Control-Allow-Origin: *` wildcard header. Build script now patches `dotnet.js` to use `credentials:"omit"` for CDN compatibility.

- **Vite bundler compatibility** - Added `@vite-ignore` comment to dynamic import in `npm/src/index.ts` to prevent Vite from trying to analyze/resolve the WASM loader path during development builds.

- **Pagination content overflow** - Fixed content overflowing page boundaries in the paginated view. The issue was caused by applying CSS transform scale to the content area while using inconsistent coordinate systems for positioning. The fix applies the scale transform to the entire page box instead, ensuring proper clipping and consistent scaling of all page elements.

- **WmlComparer legal numbering preservation** ([Issue #1634](https://github.com/dotnet/Open-XML-SDK/issues/1634)) - Fixed comparison losing legal numbering (`w:isLgl`) when comparing documents with different numbering styles. The comparer now properly merges numbering definitions from the revised document into the result:
  - Copies `abstractNum` and `num` elements from revised document when missing in original
  - Reuses existing definitions when content matches (regardless of ID)
  - Remaps IDs when conflicts occur to avoid duplicates

- **WmlToHtmlConverter null rPr crash** - Fixed `InvalidOperationException` crash in `DefineRunStyle` and `GetLangAttribute` when converting runs without `w:rPr` elements. Changed `.First()` to `.FirstOrDefault()` with null checks to handle runs that have no explicit run properties gracefully.

### Changed
- Replaced `FontPartType`/`ImagePartType` with `PartTypeInfo` pattern for SDK 3.x compatibility
- Replaced `.Close()` calls with `Dispose()` pattern
- Migrated all color handling from `System.Drawing.Color` to `SKColor`
- Migrated font handling from `FontFamily`/`FontStyle` to `SKFontManager`/`SKTypeface`
- Migrated image handling from `Bitmap`/`ImageFormat` to `SKBitmap`/`SKEncodedImageFormat`

### Documentation
- Updated `docs/architecture/wml_to_html_converter_gaps.md` with comprehensive gap analysis including pagination mode limitations, DrawingML text handling, and prioritized fix recommendations

### Test Status
- 1051 passed, 0 failed, 1 skipped out of 1052 tests (~99.9% pass rate)
- Header/footer and footnote pagination changes tested via manual integration testing

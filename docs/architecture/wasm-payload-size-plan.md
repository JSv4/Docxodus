# WASM Payload Size Reduction Plan

**Created:** August 2026
**Status:** Proposed — Phase 1 levers experimentally validated in-branch (measurements below are real builds, .NET SDK 10.0.302 / wasm-tools 10.0.10 / DocumentFormat.OpenXml 3.5.1)
**Goal:** Reduce the WASM payload to **under 5 MB** without dropping functionality.

## 1. What "payload" means (and what the target should bind to)

The npm package ships the .NET browser-wasm runtime plus our assemblies as webcil `.wasm`
files under `dist/wasm/_framework/`. Three different numbers get called "the payload":

| Metric | Today (docxodus@9.0.0) | What it affects |
|---|---|---|
| Wire transfer (what the browser downloads to boot) | **16.7 MB** uncompressed as served today; ~3.9–5.3 MB if the host compresses on the fly | First-load latency — the number users feel |
| `dist/wasm/_framework` on disk | 18 MB | npm install size, CDN storage |
| npm unpacked total | 19.4 MB | `dist.unpackedSize` on npmjs.com |

We ship **no precompressed assets** and **~0.66 MB of debug artifacts** (`.map` files +
`dotnet.native.js.symbols`) in every install. Nothing in the stack compresses the payload
for the user: the `BlazorEnableCompression` property in `DocxodusWasm.csproj` has been a
no-op since .NET 8 (it was replaced by `CompressionEnabled`, which itself only works under
the `Microsoft.NET.Sdk.WebAssembly` SDK — see Phase 2).

**Recommended target interpretation:** the 5 MB budget binds to the **wire transfer**
(the user-felt number), with a hard secondary goal of cutting the on-disk/unpacked size as
far as trimming allows (~13 MB now, ~10.5 MB with Phase 4). The measurements below show
the wire target is comfortably achievable — **3.2 MB Brotli / 4.1 MB gzip** after Phase 1
— while an *uncompressed* 5 MB is not reachable without dropping functionality (the .NET
runtime floor plus the reachable Wordprocessing/Drawing schema code alone exceed 5 MB;
§3's E3 autopsy and Phase 4 explain the ceiling).

## 2. Measured baseline (docxodus@9.0.0, reproduced bit-for-bit locally)

`dist/wasm/_framework` = 18 MB: 43 webcil `.wasm` + 4 `.js` + 2 `.map` + 1 `.symbols`.

| File | KB | Notes |
|---|---|---|
| DocumentFormat.OpenXml.wasm | 7,312 | **40% of the payload.** Fully rooted (`TrimmerRootAssembly`) |
| Docxodus.wasm | 2,931 | Fully rooted |
| System.Private.CoreLib.wasm | 1,651 | BCL-trimmed already |
| dotnet.native.wasm | 1,489 | Mono runtime + ICU-invariant + timezone data |
| System.Private.Xml.wasm | 860 | XmlReader/Writer + Schema |
| DocxodusWasm.wasm | 466 | Bridge + source-gen JSON |
| DocumentFormat.OpenXml.Framework.wasm | 358 | *Not* rooted — already ships trimmed (479 KB in the nupkg) |
| System.Text.Json.wasm | 291 | |
| dotnet.runtime.js.map | 270 | **Debug artifact, shipped** |
| System.Text.RegularExpressions.wasm | 254 | |
| dotnet.native.js / dotnet.runtime.js | 214 / 194 | |
| dotnet.native.js.symbols | 182 | **Debug artifact, shipped** |
| System.Net.Http.wasm | 135 | Survives trimming (runtime HTTP glue) |
| 29 further System.* assemblies | ~700 | Already partial-trimmed |

Wire size of that payload (excluding maps/symbols): **16.70 MB** raw,
**3.96 MB** brotli-11, **5.26 MB** gzip-9.

### Why it is this big — three root causes

1. **Trimming is disabled where it matters.** `DocxodusWasm.csproj` sets
   `PublishTrimmed=true` / `TrimMode=partial` but then roots the three application
   assemblies with `TrimmerRootAssembly` — `DocxodusWasm`, `Docxodus`, **and**
   `DocumentFormat.OpenXml`. A rooted assembly is kept in its entirety, so the two
   largest files (10.2 MB combined) ship every type they contain, including whole
   modules the WASM bridge never exposes (PresentationBuilder, SpreadsheetWriter,
   spreadsheet/presentation halves of the Open XML SDK, …).
2. **No compression story.** The project uses the plain `Microsoft.NET.Sdk` +
   `RuntimeIdentifier=browser-wasm` project style ("WasmAppBuilder"), which does not
   emit `.br`/`.gz` siblings at publish; `build-wasm.sh` even prints "Brotli compressed
   files available" checks that never fire. Static hosts that don't compress
   `application/wasm` on the fly (many don't, or cap sizes) serve 16.7 MB raw.
3. **Debug artifacts ship to npm.** `dotnet.js.map`, `dotnet.runtime.js.map`,
   `dotnet.native.js.symbols` — 0.66 MB nobody's production page loads.

## 3. Validated experiments (this branch, real `dotnet publish` builds)

### E1 — un-root + trim (TrimMode=partial)

Replace the `Docxodus`/`DocumentFormat.OpenXml` `TrimmerRootAssembly` entries with
`TrimmableAssembly` (keep rooting `DocxodusWasm`, the `[JSExport]` entry assembly):

| File | Before | After | Δ |
|---|---|---|---|
| DocumentFormat.OpenXml.wasm | 7,312 KB | 5,088 KB | **−30%** |
| Docxodus.wasm | 2,931 KB | 1,764 KB | **−40%** |
| `_framework` total | 18 MB | 14 MB | −4 MB |

Note: DocumentFormat.OpenXml ≥2.19 is built with `IsTrimmable=true`
([Open-XML-SDK #1172](https://github.com/dotnet/Open-XML-SDK/issues/1172)), so under
`TrimMode=partial` un-rooting alone suffices; the `TrimmableAssembly` items are
belt-and-braces (and required for `Docxodus`, which is not annotated).

### E2 — full trim + runtime feature switches

E1 plus `TrimMode=full`, `TrimmerRemoveSymbols=true`, `InvariantTimezone=true`,
`WasmEmitSymbolMap=false`:

| File | E1 | E2 | Δ |
|---|---|---|---|
| dotnet.native.wasm | 1,489 KB | 1,250 KB | −239 KB (timezone db) |
| System.Private.CoreLib.wasm | 1,651 KB | 1,588 KB | −63 KB |
| System.Private.Xml.wasm | 860 KB | 795 KB | −65 KB |
| dotnet.native.js.symbols | 182 KB | *(gone)* | −182 KB |

**E2 result: 12.91 MB uncompressed · 3.20 MB brotli-11 · 4.14 MB gzip-9**
(vs 16.70 / 3.96 / 5.26 baseline). The wire target is met with margin under both
encodings, and unpacked size drops 18 → ~13 MB.

**Functional gate:** the full browser Playwright suite (53 spec files, 312 tests
covering conversion, redlines/diff, sessions/mutations, comments, notes,
headers/footers, and the editor) run against the E2-trimmed artifacts —
plus `dotnet test` for the .NET side (unaffected: the csproj change is WASM-only).
Gate results are recorded in §6.

### E3 — what still survives inside the trimmed OpenXml SDK (the next ceiling)

An IL-size autopsy of the linked assemblies (System.Reflection.Metadata over
`obj/.../linked/`) shows `DocumentFormat.OpenXml.dll` retains **4,308 of 5,081 types**
(IL 2,414 of 3,095 KB). Namespaces a DOCX-only pipeline never touches survive:

| Kept namespace | Types | IL KB |
|---|---|---|
| Spreadsheet | 356 | 221 |
| Presentation | 319 | 140 |
| Drawing.Charts | 354 | 130 |
| Office2010.CustomUI + Office.CustomUI | 137 | 350 |
| Office2016.Drawing.ChartDrawing | 140 | 58 |
| Office2010.Excel | 98 | 49 |
| InkML | 49 | 28 |

≈1.0–1.1 MB of IL (≈2–2.5 MB of webcil bytes) is retained by two distinct chains:

1. **The SDK's typed part/element factory graph** — part classes statically reference
   the element types of every schema family they could contain, so
   `WordprocessingDocument`'s part constellation transitively roots
   spreadsheet/presentation/chart/CustomUI code. A known upstream limitation
   ([Open-XML-SDK #1349 — "Ensure a package only references required types"](https://github.com/dotnet/Open-XML-SDK/issues/1349)).
2. **`OpenXmlValidator`, rooted by one call**: `DocxSession.CountRealValidationErrors`
   (`DocxSession.cs:4882`), used by `Raw.InsertXml`/`Raw.ReplaceXml` when
   `Settings.ValidateRawOps` is on and statically reachable from
   `DocxSessionBridge.RawInsertXml/RawReplaceXml`. The validator keeps the full typed
   Wordprocessing object model plus the validation subsystem alive (it also explains
   most of the 173 KB of `System.Xml.Schema` IL that survives in
   `System.Private.Xml`). This single call is the biggest brake on shrinking the SDK
   further — making it linker-severable (a feature switch defaulting raw-op validation
   off in WASM, with a clear error if enabled on a build without it) is the highest-value
   Phase 4 item.

By contrast `Docxodus.dll` trims cleanly to 380 types / 856 KB IL — only the modules
actually reachable from the `[JSExport]` bridge remain. The removed ~24% of library
source includes the ~17k-line HtmlToWml CSS engine, DocumentBuilder,
PresentationBuilder, SpreadsheetWriter, ChartUpdater, DocumentAssembler, TextReplacer,
OpenXmlRegex, and the bulk of MetricsGetter — none of which were ever exported to the
browser.

## 4. The plan

### Phase 0 — Guardrails first (no size change)

1. **Size report + budget check.** Extend `build-wasm.sh` to print a per-file and total
   payload table (raw + brotli) and fail if the brotli wire total exceeds a checked-in
   budget (start: 5 MB; ratchet down after Phase 1 lands to ~3.5 MB). Add the same check
   to `ci.yml`/`playwright.yml` so regressions (a dependency bump, an accidental
   re-root) are caught at PR time, not at release time.
2. **Un-suppress trim analysis selectively.** `SuppressTrimAnalysisWarnings=true`
   currently hides the exact warnings that predict trim-induced runtime failures. Flip it
   off once, catalogue the warnings (expected: the `GetPackage()` reflection in
   `PtOpenXmlUtil.cs` and friends), then either fix with `[DynamicDependency]` /
   descriptor entries or re-suppress with a comment inventorying what was reviewed.

### Phase 1 — Trim the application assemblies (validated: 18 → ~13 MB, wire 3.2 MB br)

The E1+E2 csproj change, gated on the full Playwright suite + `dotnet test`:

```xml
<TrimMode>full</TrimMode>
<TrimmerRemoveSymbols>true</TrimmerRemoveSymbols>
<InvariantTimezone>true</InvariantTimezone>
<WasmEmitSymbolMap>false</WasmEmitSymbolMap>
...
<ItemGroup>
  <TrimmerRootAssembly Include="DocxodusWasm" />
  <TrimmableAssembly Include="Docxodus" />
  <TrimmableAssembly Include="DocumentFormat.OpenXml" />
  <TrimmableAssembly Include="DocumentFormat.OpenXml.Framework" />
</ItemGroup>
```

Functionality is defined by the `[JSExport]` surface: everything reachable from
`DocumentConverter`, `DocumentComparer`, `DocxDiffBridge`, and `DocxSessionBridge` is
kept by the trimmer automatically; code that was never callable from the browser
(PresentationBuilder, SpreadsheetWriter, OpenContractExporter CLI paths, …) is what
gets removed. **No exported API is dropped.**

A code audit of the WASM-compiled tree found **no blockers**. The specifics, and what
each needs:

- *The one real reflection hotspot:* `PtOpenXmlUtil.cs:27-137` `GetPackage(this
  OpenXmlPackage)` reflects into `DocumentFormat.OpenXml.Framework` (string-based
  `Assembly.GetType("DocumentFormat.OpenXml.Features.IPackageFeature")`,
  `GetMethod("Get")` + `MakeGenericMethod`, private `_package` fields). ILLink has no
  static edge to any of these members. Empirically it survives — Framework is
  `IsTrimmable` and already ships trimmed today with this code working — because the
  SDK's own kept code uses the same members, but that's luck, not a contract. Two
  fixes, either sufficient:
  1. **Preferred: refactor the two bridge-reachable call sites**
     (`WmlComparer.cs:6215-6240`, drawing/media copying during a WmlComparer-engine
     compare) to take the `Package` from `OpenXmlMemoryStreamDocument.GetPackage()` (a
     plain field accessor) the way `IrMarkupRenderer.cs:720-724` already deliberately
     does. The reflective helper then becomes bridge-unreachable and trims away
     entirely (its remaining callers — `Consolidate`, `PresentationBuilder` — are not
     bridged).
  2. Or add a `TrimmerRootDescriptor` (`ILLink.Descriptors.xml`) preserving
     `DocumentFormat.OpenXml.Features.*` non-public fields + `OpenXmlPackage`'s
     `_package` — surgical preservation instead of rooting 7.3 MB.
- *JSON is trim-safe on every reachable path.* The bridge passes
  `DocxodusJsonContext.Default.<T>` (source-gen) at all ~44 call sites; the core
  facades (`DocxSessionJson`, `DocxDiffOps`, `HtmlConversionOps`, `IrEditScriptJson`)
  use hand-rolled `Utf8JsonWriter`/`JsonDocument` by documented design. The only
  reflection-based `JsonSerializer` calls in the library
  (`ExternalAnnotationManager.SerializeToJson/DeserializeFromJson`) have zero callers
  from the bridge and already throw `JsonSerializerIsReflectionDisabled` in browser
  builds — the trimmer removes them.
- *No other hostile patterns.* No embedded resources in either assembly (the SDK's
  Framework carries ~25 KB of exception/validation `.resources` — noise), no
  `Activator.CreateInstance`, no `XmlSerializer`/`TypeConverter` usage, all 130
  `AddNewPart<T>` sites statically generic. `DocxodusWasm` itself must **stay** a
  `TrimmerRootAssembly`: JS resolves its 148 `[JSExport]`s by name at runtime.
- *Keep `TrimMode=partial` as fallback.* If full-mode surfaces breakage that partial
  doesn't, E1 alone still gets 14 MB / ~3.5 MB brotli; decide on evidence.
- *`.map` files*: `dotnet.js.map` + `dotnet.runtime.js.map` (306 KB) are still emitted;
  exclude them in the `build-wasm.sh` copy step (they're referenced by `//#
  sourceMappingURL` comments only — browsers fetch them lazily in DevTools; npm users
  who want them can use a debug build).

### Phase 2 — Ship the compression story (wire ≤ 3.5 MB for any host)

Emitting `.br`/`.gz` siblings and making them reachable:

1. **Precompress at package build.** Add a brotli-11 + gzip-9 pass over
   `dist/wasm/_framework/**` in `build-wasm.sh` (or a small Node script — no native
   `brotli` binary needed, `zlib.brotliCompressSync` suffices). Ship the `.br` siblings
   in the npm package. Cost: package size grows by the compressed set (~3.3 MB);
   mitigate by dropping `.gz` (gzip-capable hosts can compress on the fly; `.br` is the
   one that needs precompression because brotli-11 is too slow for on-the-fly).
2. **Serving guidance (README + docs/npm-package.md).** Hosts that support
   content negotiation (nginx `brotli_static`, Caddy `precompressed`, Netlify, Vercel,
   Cloudflare Pages) serve the `.br` sibling with `Content-Encoding: br` — browser-side
   nothing changes, wire = 3.2 MB.
3. **Loader fallback for dumb static hosts (optional, npm-side).** The runtime's
   `dotnet.withResourceLoader((type, name, defaultUri, integrity, behavior) => …)` hook
   lets `npm/src/index.ts` fetch `name + '.br'` and decompress client-side. Caveat:
   `DecompressionStream` does not support brotli — a JS/wasm brotli decoder (~200 KB)
   would be needed, **or** fall back to gzip siblings via
   `DecompressionStream('gzip')` (validated: 4.14 MB gzip — still under target with
   zero extra decoder bytes). Recommend: negotiate-first, gzip-stream fallback,
   plain fetch last. Feature-detect and keep the current plain path as default so
   nothing breaks.
4. **Consider migrating the csproj to `Microsoft.NET.Sdk.WebAssembly`** (the SDK the
   current `wasmbrowser` template uses). It emits max-level `.br`/`.gz` at publish
   natively (`CompressionEnabled`, replacing the dead `BlazorEnableCompression`),
   inlines the boot config into `dotnet.js` (one fewer round trip), supports
   `WasmFingerprintAssets` (turn **off** for npm path stability) and
   `WasmBundlerFriendlyBootConfig=true` (.NET 10: emits `import`-visible assets so
   consumers' bundlers — Vite/webpack — can see and optimize the asset graph; would
   remove our `credentials:"omit"`/integrity sed-patches in `build-wasm.sh` in favor of
   supported hooks). This is the strategic fix; the sed-patched Style-A pipeline is
   the tactical one. Migration is mechanical (`Sdk` attribute + `main.js` staying on
   the same `dotnet.js` API) but needs its own Playwright pass; treat as a separate PR.

### Phase 3 — npm packaging hygiene (−1 MB unpacked, no behavior change)

- Stop copying `*.map` and `*.symbols` into `dist/wasm/_framework` (build-wasm.sh).
- Verify `docs/npm-package.md` "Bundle Size" section reflects the new reality; document
  the wire-vs-disk distinction and the serving guidance from Phase 2.
- Optional: split the WASM runtime into its own npm package (`docxodus-wasm`) so pure-TS
  consumers (agent tooling importing types only) don't pull 13 MB — orthogonal to the
  5 MB goal but cheap to do while touching packaging. Needs a major-version note.

### Phase 4 — Deeper OpenXml surgery (optional; ~13 → ~10.5 MB unpacked)

Only if the unpacked/CDN-storage number matters after Phases 1–3 (the wire goal is
already met):

- **Make the `OpenXmlValidator` call linker-severable.** One reachable call
  (`DocxSession.cs:4882`, gated on `Settings.ValidateRawOps`) roots the typed
  Wordprocessing model + validation subsystem. Introduce a feature switch (ILLink
  substitution over a static bool, the standard BCL pattern) so the WASM build can
  publish with raw-op validation compiled out — `Raw.InsertXml`/`ReplaceXml` keep
  working, just without the schema-error count guard (or return a clear
  `EditErrorCode` if a caller enables `ValidateRawOps` on a validation-free build).
  This is the highest-value single cut left in the SDK assembly.
- **ILLink substitutions/descriptors to sever the typed part factory** for schema
  families the DOCX pipeline can't reach (Spreadsheet, Presentation, Charts, CustomUI,
  InkML ≈ 2–2.5 MB webcil). Mechanism: an `ILLink.Substitutions.xml` that stubs the
  factory branches, or feature-switch work upstream. **Brittle across SDK updates** —
  each DocumentFormat.OpenXml bump must re-validate against the full suite; embedded
  XLSX chart parts inside a DOCX would throw on typed access (today they're passed
  through untyped, so this is likely safe — verify with `WC/`+chart fixtures).
- **Upstream**: comment on Open-XML-SDK #1349 with our measurements (4,308/5,081 types
  survive a Wordprocessing-only app) and ask for feature-switched part factories; that
  is the clean long-term fix.
- **Not recommended**: NativeAOT (≈2× payload), dropping SIMD (compat, not size),
  lazy-loading assemblies (all our assemblies are needed for the first conversion;
  nothing meaningful to defer).

## 5. Expected end state

| Stage | Unpacked `_framework` | Wire (brotli) | Wire (gzip) |
|---|---|---|---|
| Today (9.0.0) | 18 MB | *(no .br shipped)* raw 16.7 MB | host-dependent |
| Phase 1 | ~13 MB | **3.2 MB** | 4.1 MB |
| + Phase 3 | ~12.6 MB | 3.2 MB | 4.1 MB |
| + Phase 4 (optional) | ~10.5 MB | ~2.6 MB (est.) | ~3.4 MB (est.) |

**The 5 MB goal is met at Phase 1+2** with ~1.8 MB of headroom on brotli hosts and
~0.9 MB on gzip-only paths, with zero functionality dropped — the trimmer removes only
code the `[JSExport]` surface never exposed to the browser.

## 6. Test/rollout gates

Every phase lands only when all of:
1. `cd npm && npm run build && npm test` — full Playwright suite green against the new
   artifacts (the suite loads the real runtime in Chromium; a trimmed-away member fails
   loudly here).
2. `dotnet test Docxodus.Tests` green (guards against accidental non-WASM regressions).
3. Size report shows the expected numbers; CI budget check updated to ratchet.
4. One manual smoke of `npm run demo` (editor on `HC031-Complicated-Document.docx`).

Two spec additions the audit identified as gaps — both cover the only paths where
trimming could bite at runtime, so they should land **with** Phase 1:
- A **WmlComparer-engine (`engine: 0`) compare of documents containing
  images/drawings** — the sole bridge-reachable route through the reflective
  `GetPackage()` (media-part copying during `CoalesceRecurse`).
- A **`rawInsertXml`/`rawReplaceXml` call with `validateRawOps` enabled** — the
  `OpenXmlValidator` path.

## Appendix: methodology

- Baseline = published `docxodus@9.0.0` tarball, cross-checked bit-for-bit against a
  local `dotnet publish -c Release` of the unmodified csproj (sizes matched).
- The publish output **is** `bin/Release/net10.0/browser-wasm/AppBundle` in this SDK
  layout — `build-wasm.sh`'s first path already picks the trimmed output (an earlier
  concern that it might copy an untrimmed build-stage bundle is unfounded).
- Compression ratios: Node `zlib.brotliCompressSync` quality 11 and `gzipSync` level 9,
  per file, `.map`/`.symbols` excluded.
- IL autopsy: System.Reflection.Metadata walker over `obj/.../browser-wasm/linked/*.dll`
  (post-ILLink IL), namespace-bucketed method-body bytes.

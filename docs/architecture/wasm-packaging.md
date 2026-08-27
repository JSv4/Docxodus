# WASM Packaging

How the browser payload is built, trimmed, compressed, and kept small. This is the
reference for `wasm/DocxodusWasm/DocxodusWasm.csproj`, `scripts/build-wasm.sh`, and the
size guardrail. (The original investigation/plan doc was folded into this overview after
implementation; measured numbers below are from real builds, .NET SDK 10.0.302 /
wasm-tools 10.0.10 / DocumentFormat.OpenXml 3.5.1, 2026-08.)

## What ships

`npm run build` → `scripts/build-wasm.sh` publishes `wasm/DocxodusWasm` (Release,
browser-wasm) and copies the AppBundle's `_framework/` into `npm/dist/wasm/`:

- ~40 webcil `.wasm` assemblies + `dotnet.js` / `dotnet.runtime.js` / `dotnet.native.js`
  \+ `dotnet.native.wasm` + `dotnet.boot.js`
- a `.br` sibling for every asset (brotli quality 11), for hosts that serve
  precompressed content
- **no** `.map` / `.symbols` debug artifacts (use a Debug build when you need them)

| Metric | Before (≤ 9.0.0) | Now |
|---|---|---|
| Browser-fetched payload, uncompressed | 16.7 MB | **14.7 MB** |
| Wire transfer on a brotli-serving host | *(no .br shipped)* | **3.60 MB** |
| Wire transfer on a gzip-on-the-fly host | ~5.3 MB | **4.63 MB** |
| Largest asset: DocumentFormat.OpenXml.wasm | 7.3 MB | 5.0 MB |
| Docxodus.wasm | 2.9 MB | 3.3 MB |

The `Docxodus.wasm` row grew rather than shrank: the trimmer already removed the
SpreadsheetML/PresentationML modules from the browser payload long before they were deleted
from the source tree, so that purge moved this number by nothing. What moved it is
everything added since 9.0.0 — the session op surface, the delivery/verification subsystems,
and pagination. The guardrail that matters is the brotli wire total, which
`scripts/build-wasm.sh` prints and holds under a 4096 KB budget.

## Trimming policy

The csproj publishes with `TrimMode=full` and roots **only** `DocxodusWasm` (JS resolves
its `[JSExport]`s by name at runtime, invisibly to ILLink). `Docxodus`,
`DocumentFormat.OpenXml`, and `DocumentFormat.OpenXml.Framework` are opted into trimming
via `TrimmableAssembly` — everything not reachable from the bridge surface
(`DocumentConverter`, `DocumentComparer`, `DocxDiffBridge`, `DocxSessionBridge`) is
removed. That deletes the modules never exported to the browser (HtmlToWml,
DocumentBuilder, DocumentAssembler, WmlToXml, …) and the unreachable halves of the
Open XML SDK. No exported API changes.

Feature switches: `InvariantGlobalization` (no ICU), `InvariantTimezone` (no tz
database, −240 KB from `dotnet.native.wasm`), `TrimmerRemoveSymbols`,
`WasmEmitSymbolMap=false`, plus the usual Release switches (`DebuggerSupport=false`,
`EventSourceSupport=false`, `UseSystemResourceKeys`). AOT stays **off** — it roughly
doubles payload size for this workload.

### The two trim-sensitive paths (and their pins)

Everything in the WASM-compiled tree is statically analyzable except:

1. **`PtOpenXmlUtil.GetPackage()`** — extracts `System.IO.Packaging.Package` from an
   `OpenXmlPackage` by reflecting through the SDK 3.x features chain. Bridge-reachable
   only when a WmlComparer-engine compare copies image/media parts.
   **Pin:** `wasm/DocxodusWasm/ILLink.Descriptors.xml` preserves
   `DocumentFormat.OpenXml.Features.*` (+ `OpenXmlPackage` fields), ~19 KB of IL.
2. **`OpenXmlValidator`** — reached only through `Raw.InsertXml`/`Raw.ReplaceXml` with
   `validateRawOps` on. Statically reachable (no pin needed), but it is the reason the
   full typed Wordprocessing model survives in the SDK assembly.

Both paths have permanent browser canaries in `npm/tests/trim-validation.spec.ts`. If
either fails after an SDK bump, suspect the descriptor first.

`SuppressTrimAnalysisWarnings` stays `true` deliberately: the one reflective pattern is
pinned and canaried, and trim safety is enforced by the Playwright suite (669 tests run
against the trimmed artifacts; also `dotnet test` for the non-WASM side).

## Compression and serving

`build-wasm.sh` writes a brotli-11 `.br` sibling next to every `_framework` asset. The
loader is unchanged — compression is the host's job:

- **Hosts with content negotiation** (nginx `brotli_static on`, Caddy `precompressed`,
  Netlify, Vercel, Cloudflare Pages): serve the `.br` sibling with
  `Content-Encoding: br` + `Vary: Accept-Encoding`, keeping the original
  `Content-Type` (`application/wasm`). Wire ≈ 3.6 MB; the browser's network stack
  decompresses while streaming — **cold open is not slowed** (measured below).
- **Hosts that gzip on the fly**: wire ≈ 4.6 MB, nothing to configure.
- **Dumb static hosts**: raw ~14.7 MB. (A JS-side brotli decode fallback was evaluated
  and deliberately **not** shipped: `DecompressionStream` has no brotli support, and a
  JS/wasm decoder decompressing the whole payload single-threaded is the pattern that
  makes brotli *feel* slow. If a fallback is ever wanted, prefer gzip via
  `DecompressionStream('gzip')` through `dotnet.withResourceLoader(...)`.)

gzip siblings are intentionally not precompressed (gzip-capable hosts do it on the fly;
brotli-11 is the one too slow for that).

### Cold-open performance (measured)

Time from navigation to `window.DocxodusReady`, median of 5 cold boots (fresh browser
per boot, cold cache), Chromium 141, wire bytes verified via CDP:

| Payload | localhost | 50 Mbps + 20 ms RTT |
|---|---|---|
| old untrimmed, raw (16.9 MB wire) | 703 ms | 3,295 ms |
| **new trimmed, raw (12.9 MB wire)** | 620 ms | 2,588 ms |
| old untrimmed, brotli (4.2 MB wire) | 747 ms | 1,202 ms |
| **new trimmed, brotli (3.2 MB wire)** | 665 ms | **1,022 ms** |

Two conclusions. Native `Content-Encoding: br` decode costs ~45 ms on localhost —
noise, not the "notable slowdown" associated with brotli in the browser (that effect
comes from the JS-decoder pattern above, which is why it isn't shipped). And on a real
network the compressed payload dominates everything: at 50 Mbps, trimmed+brotli cold
open is **1.0 s vs 3.3 s for today's shipped payload — 3.2× faster**. Trimming alone
is worth ~80 ms even on localhost (less IL to parse) and ~700 ms at 50 Mbps.

## Size guardrail

`build-wasm.sh` computes the brotli wire total on every build and **fails above 4 MB**
(measured 3.60 MB). If it trips: look for a re-rooted assembly (`TrimmerRootAssembly`),
a dependency bump growing the SDK, or a new package reference. The npm CI job runs the
same script, so regressions surface at PR time.

## Future size work (not implemented)

The remaining uncompressed ceiling is `DocumentFormat.OpenXml.wasm` (5.0 MB): the SDK's
typed part factory statically roots every schema family a part *could* contain
(Spreadsheet, Presentation, Charts, CustomUI, InkML ≈ 2–2.5 MB of webcil a DOCX-only
pipeline never touches — upstream issue
[Open-XML-SDK #1349](https://github.com/dotnet/Open-XML-SDK/issues/1349)), and the one
`OpenXmlValidator` call (`DocxSession.CountRealValidationErrors`) keeps the typed
Wordprocessing model + validation subsystem. Options if unpacked size ever matters
more: a feature switch making raw-op validation linker-severable, ILLink substitutions
stubbing the factory branches (brittle across SDK bumps), or migrating the csproj to
`Microsoft.NET.Sdk.WebAssembly` (native `CompressionEnabled`, inlined boot config,
`WasmBundlerFriendlyBootConfig` for consumer bundlers — would replace the
`credentials`/integrity sed-patches in `build-wasm.sh`). The wire numbers above make
none of it urgent.

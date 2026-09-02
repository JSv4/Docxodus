# WASM Packaging

How the browser payload is built, trimmed, AOT-compiled, compressed, and kept small. This
is the reference for `wasm/DocxodusWasm/DocxodusWasm.csproj`, `scripts/build-wasm.sh`,
`scripts/record-aot-profile.sh`, and the size guardrail. (The original investigation/plan
doc was folded into this overview after implementation; measured numbers below are from
real builds, .NET SDK 10.0.301 / wasm-tools 10.0.109 / DocumentFormat.OpenXml 3.5.1,
2026-08 for trimming and 2026-09 for the AOT tier.)

## What ships

`npm run build` → `scripts/build-wasm.sh` publishes `wasm/DocxodusWasm` (Release,
browser-wasm) and copies the AppBundle's `_framework/` into `npm/dist/wasm/`:

- ~40 webcil `.wasm` assemblies + `dotnet.js` / `dotnet.runtime.js` / `dotnet.native.js`
  \+ `dotnet.native.wasm` (runtime + the profile-guided AOT code, see below) + `dotnet.boot.js`
- a `.br` sibling for every asset (brotli quality 11), for hosts that serve
  precompressed content
- **no** `.map` / `.symbols` debug artifacts (use a Debug build when you need them)

| Metric | ≤ 9.0.0 | 10.0.0 (trimmed, interpreter) | Now (+ profile-guided AOT) |
|---|---|---|---|
| Browser-fetched payload, uncompressed | 16.7 MB | 14.7 MB | **21.2 MB** |
| Wire transfer on a brotli-serving host | *(no .br shipped)* | 3.60 MB | **4.76 MB** |
| Wire transfer on a gzip-on-the-fly host | ~5.3 MB | 4.63 MB | **6.5 MB** |
| Largest assembly: DocumentFormat.OpenXml.wasm | 7.3 MB | 5.0 MB | 5.0 MB |
| Docxodus.wasm | 2.9 MB | 3.3 MB | 3.2 MB |
| dotnet.native.wasm | — | 1.28 MB | 7.93 MB |

The `Docxodus.wasm` row grew between 9.0.0 and 10.0.0 rather than shrinking: the trimmer
already removed the SpreadsheetML/PresentationML modules from the browser payload long
before they were deleted from the source tree, so that purge moved this number by nothing.
What moved it is everything added since 9.0.0 — the session op surface, the
delivery/verification subsystems, and pagination. The AOT tier then moved
`dotnet.native.wasm`: the compiled methods live there, while the assemblies keep their size
(the IL bodies of AOT-compiled methods are zeroed in place, which is why they cost nothing
after compression). The guardrail that matters is the brotli wire total, which
`scripts/build-wasm.sh` prints and holds under a 5120 KB budget.

## Trimming policy

The csproj publishes with `TrimMode=full` and roots **only** `DocxodusWasm` (JS resolves
its `[JSExport]`s by name at runtime, invisibly to ILLink). `Docxodus`,
`DocumentFormat.OpenXml`, and `DocumentFormat.OpenXml.Framework` are opted into trimming
via `TrimmableAssembly` — everything not reachable from the bridge surface
(`DocumentConverter`, `DocumentComparer`, `DocxDiffBridge`, `DocxSessionBridge`) is
removed. That deletes the modules never exported to the browser (HtmlToWml,
DocumentBuilder, OpenXmlRegex, …) and the unreachable halves of the Open XML SDK.
No exported API changes.

Feature switches: `InvariantGlobalization` (no ICU), `InvariantTimezone` (no tz
database, −240 KB from `dotnet.native.wasm`), `TrimmerRemoveSymbols`,
`WasmEmitSymbolMap=false`, plus the usual Release switches (`DebuggerSupport=false`,
`EventSourceSupport=false`, `UseSystemResourceKeys`). AOT is **profile-guided**, never
full — see the next section for why and for the measured frontier.

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

## Runtime tier: jiterpreter + profile-guided AOT

The browser build executes IL on the Mono **interpreter**, tiered by the **jiterpreter**
(hot interpreter traces compiled to WebAssembly at runtime). Both are on in the shipped
configuration — verified, not assumed: `npm/tests/wasm-steady-state.spec.ts` boots the
bundle with `--jiterpreter-stats-enabled` and asserts that traces, interp-entry and jit-call
thunks are enabled, that traces were actually compiled, and that the generated code sits well
inside the jiterpreter's 8 MB budget (a few compares use ~0.6 MB / 500 traces). Even so, the
interpreter's steady state was **5–10× slower than warm native** on the same inputs: the
jiterpreter compiles straight-line loops well but every trace ends at a call, and this code
base is calls all the way down (XLinq accessors, LINQ, virtual dispatch).

The fix (issue #652) is **profile-guided AOT**: `RunAOTCompilation=true` with
`WasmAotProfilePath` pointing at `wasm/DocxodusWasm/docxodus.aotprofile`. The AOT compiler
then runs with `profile-only,profile=…` and compiles *only the methods listed in the
profile* — the ~6k the representative workload executes — while everything else stays on the
interpreter (`AOTMode=LLVMOnlyInterp`, the SDK default). The profile is recorded with the
Mono AOT profiler over the same workload the steady-state spec times (`npm/tests/wasm-workload.ts`:
DocxDiff compare of a small pair and of a 147 KB legal form against an edited variant of
itself, DOCX→HTML on two documents, and the editor's per-mutation ReplaceText + single-block
re-render), so what is measured is by construction what is compiled.

### Measured frontier (2026-09, 8-core Linux, Playwright Chromium; medians of a warm loop)

All four columns were measured on one tree, just before the WmlComparer engine was removed
(#643); the shipped build on the current tree is ~50 KB smaller (4871 KB wire, 7.93 MB
native, 6,089 methods) with the same timings to within noise.

| Operation | Warm native (.NET 10 x64) | Interpreter + jiterpreter | **Profile-guided AOT** | Full AOT |
|---|---|---|---|---|
| DocxDiff compare, 11 KB pair | 17 ms | 126 ms (7.5×) | **30 ms (1.8×)** | 30 ms |
| DocxDiff compare, 147 KB legal form vs edited variant | 818 ms | 5.90 s (7.2×) | **1.43 s (1.7×)** | 1.53 s |
| DOCX→HTML, 42 KB (HC031) | 150 ms | 893 ms (5.9×) | **235 ms (1.6×)** | 254 ms |
| DOCX→HTML, 147 KB legal form | 902 ms | 4.55 s (5.0×) | **1.31 s (1.5×)** | 1.33 s |
| Editor refresh (ReplaceText + block re-render) | 5.6 ms | 55.9 ms (10×) | **15.3 ms (2.7×)** | 15.4 ms |
| Methods AOT-compiled | — | 0 | 6,128 | 90,743 |
| `dotnet.native.wasm` | — | 1.28 MB | 8.37 MB | 48.7 MB |
| Payload, uncompressed | — | 14.7 MB | 21.5 MB | 60.0 MB |
| Brotli wire total | — | 3702 KB | **4925 KB** | 9673 KB |
| `dotnet publish` wall clock | — | ~1.5 min | ~3 min | ~11 min |

Three conclusions. The profile buys the whole speedup: **3.5–4.3× over the interpreter,
inside 1.5–2.7× of warm native**, and full AOT is *not faster* — the interpreter is no longer
on the hot path either way, and the 85k extra compiled methods (52k of them the Open XML
SDK's typed schema) only add bytes. The cost is **+1.2 MB over the wire and +6.8 MB
uncompressed**, which is why the wire budget moved from 4096 KB to 5120 KB — a deliberate
trade of ~200 ms of first load at 50 Mbps (once; `.br` assets are cached after) for
3–4× on every operation after it; hosts serving the raw payload pay ~1.1 s more. Cold boot
on localhost (median of 5, fresh browser each, raw assets, measured the same day on both
builds) is 738 ms with the AOT tier versus 637 ms without — the extra native code is
compiled by the browser as it streams. And compiling the AOT bitcode for size
(`WasmBitcodeCompileOptimizationFlag=-Oz`) recovers nothing (−9 KB wire): the bitcode is
already optimized inside the AOT compiler and the volume is method count, not codegen.

### Re-recording the profile

```bash
./scripts/record-aot-profile.sh     # ~4 min: profiler build → browser run → shipped rebuild
```

The script publishes the **profiler flavour** (`-p:RunAOTCompilation=false
-p:WasmProfilers=aot`: AOT off, because AOT-compiled methods are invisible to the profiler),
runs `npm/tests/aot-profile-record.spec.ts` (opt-in via `DOCXODUS_RECORD_AOT_PROFILE=1`),
which drives the shared workload in `test-harness.html?aotProfile=1` and writes
`INTERNAL.aotProfileData` to `wasm/DocxodusWasm/docxodus.aotprofile`, then rebuilds the
shipped configuration so `dist/wasm` never holds the profiler flavour. Commit the profile
(1.3 MB raw, ~250 KB in git). Re-record when the hot paths move — a new engine stage, a
renamed hot class, a runtime bump. A **stale profile costs speed, never correctness**: a method
missing from it simply runs interpreted, and a method it names that no longer exists is
skipped. The steady-state spec's numbers are how you notice drift.

The Mono AOT profiler records every method the runtime compiles (there is no hotness
threshold), so the profile is exactly the code the workload touched; widen the workload in
`wasm-workload.ts` if a new user-facing path needs the tier, and expect the wire total to
follow. Three things about the recording that are not obvious from the SDK docs:

- **`WasmAotProfilePath`, not `AOTProfilePath`.** `WasmApp.Common.targets` passes both to
  the `MonoAOTCompiler` task under what is, to MSBuild, one case-insensitive parameter; the
  item form (fed by `WasmAotProfilePath`) is evaluated last, so a build that sets only the
  documented `AOTProfilePath` silently gets **full** AOT (the 9673 KB column above — that is
  how it was measured).
- **The runtime's default hand-off method does not exist in .NET 10.** `aotProfilerOptions`
  defaults `sendTo` to `Interop/Runtime::DumpAotProfileData`, but the method lives on
  `System.Runtime.InteropServices.JavaScript.JavaScriptExports`; the harness names it
  explicitly, and `ILLink.Descriptors.AotProfiler.xml` (included only when `WasmProfilers`
  contains `aot`) roots it, because nothing references it statically and `TrimMode=full`
  otherwise removes it — the symptom is a console error, not an exception.
- **`writeAt` fires when the named method is first *compiled*, once.** The harness uses
  `DocxodusWasm.DocumentComparer::Warmup`, which nothing calls during boot, so the recorder
  calls it exactly once, after the workload.

The AOT-compiled `dotnet.native.wasm` is a different binary from the interpreter build, so
the trim canaries in `trim-validation.spec.ts` and the whole Playwright suite are what prove
the tier did not change behaviour; both ran green on the AOT bundle before it shipped.

### `Jiterpreter table N is not yet initialized` on the console

The AOT bundle logs a short burst of these during boot, and only during boot:

```
MONO_WASM: Jiterpreter table 3 is not yet initialized
MONO_WASM: Jiterpreter table 12 is not yet initialized
MONO_WASM: Jiterpreter table 31 is not yet initialized
...
```

It is **cosmetic runtime noise, not a Docxodus fault and not a broken tier** — but it is worth
understanding, because a genuinely broken jiterpreter looks exactly the same in the console.

The jiterpreter reserves 38 slices of the WebAssembly indirect function table: one for compiled
traces, one for `do_jit_call` thunks, and 36 for *interp_entry* thunks — the wrappers that
handle a call crossing **from AOT-compiled native code into an interpreted method**. There is
one interp_entry table per call shape, which is what the number in the message decodes to:

| Table | Shape |
|---|---|
| 0 | compiled traces |
| 1 | `do_jit_call` thunks |
| 2–10 | static, `void` return, 0–8 arguments |
| 11–19 | static, returns a value, 0–8 arguments |
| 20–28 | instance, `void` return, 0–8 arguments |
| 29–37 | instance, returns a value, 0–8 arguments |

The runtime allocates all 38 in `jiterpreter_allocate_tables()`, which `start_runtime()` calls
on the line *after* `mono_wasm_load_runtime()` returns. Runtime startup itself crosses the
AOT→interp boundary a handful of times, and each first crossing for a given method builds its
entry thunk eagerly (`interp_create_method_pointer_llvmonly` in `interp/interp.c`, which caches
the result on `imethod->jit_entry`). Those few crossings happen while the tables are still
zeroed, so `mono_jiterp_allocate_table_entry` prints the message and returns 0 — and the caller
falls back to the runtime's generic entry wrapper, which is the pre-jiterpreter path and is
explicitly handled: *"Compiling a trampoline can fail for various reasons, so in that case we
will fall back to the pre-existing ones below."* The cost is that those specific startup methods
never get a specialized thunk. Everything after `start_runtime()` — i.e. every method the
workload touches — allocates normally.

This is why it arrived with the AOT tier and was never seen on the interpreter build: with no
AOT code there are no AOT→interp transitions to wrap. Measured on this bundle (HC031, Chromium,
2026-09):

| | Boot-time messages | During workload | Jiterpreter stats after the workload | DOCX→HTML |
|---|---|---|---|---|
| Interpreter build (`RunAOTCompilation=false`) | 0 | 0 | `567 KB jitted; 500 traces; 0 jit_calls; 0 interp_entries` | 853 ms |
| Shipped profile-guided AOT | 9 | 0 | `6.6 KB jitted; 8 traces; 6 jit_calls; 11 interp_entries` | 249 ms |

Nine is deterministic across runs, and the tables *do* get allocated: the same boot logs
`Allocated 122881 function table entries for jiterpreter`, and interp_entry thunks are compiled
normally afterwards.

**The failure mode this hides.** If `jiterpreter_allocate_tables()` ever threw — a future SDK
growing the reservation past what `WebAssembly.Table.grow` will give, say — *no* table would be
initialized, every trace and thunk would silently fall back to the interpreter, and the console
would show the same message with different numbers. `npm/tests/wasm-steady-state.spec.ts` pins
the difference: the allocation line must be present, the messages must name only interp_entry
tables (never table 0 or 1), and none may appear after boot. There is no supported knob that
reorders the two calls, so the noise itself stays; silencing it would mean
`--jiterpreter-interp-entry-enabled=0`, which turns the thunks off rather than fixing anything.

## Compression and serving

`build-wasm.sh` writes a brotli-11 `.br` sibling next to every `_framework` asset. The
loader is unchanged — compression is the host's job:

- **Hosts with content negotiation** (nginx `brotli_static on`, Caddy `precompressed`,
  Netlify, Vercel, Cloudflare Pages): serve the `.br` sibling with
  `Content-Encoding: br` + `Vary: Accept-Encoding`, keeping the original
  `Content-Type` (`application/wasm`). Wire ≈ 4.8 MB; the browser's network stack
  decompresses while streaming — **cold open is not slowed** (measured below).
- **Hosts that gzip on the fly**: wire ≈ 6.6 MB, nothing to configure.
- **Dumb static hosts**: raw ~21.2 MB. (A JS-side brotli decode fallback was evaluated
  and deliberately **not** shipped: `DecompressionStream` has no brotli support, and a
  JS/wasm decoder decompressing the whole payload single-threaded is the pattern that
  makes brotli *feel* slow. If a fallback is ever wanted, prefer gzip via
  `DecompressionStream('gzip')` through `dotnet.withResourceLoader(...)`.)

gzip siblings are intentionally not precompressed (gzip-capable hosts do it on the fly;
brotli-11 is the one too slow for that).

### Cold-open performance (measured, interpreter build, 2026-08)

Time from navigation to `window.DocxodusReady`, median of 5 cold boots (fresh browser
per boot, cold cache), Chromium 141, wire bytes verified via CDP. The AOT tier adds
~1.2 MB brotli / ~6.8 MB raw on top of these payloads (see the frontier table above for
its localhost boot cost):

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

`build-wasm.sh` computes the brotli wire total on every build and **fails above 5 MB**
(measured 4.76 MB: ~3.6 MB of trimmed IL and runtime, ~1.2 MB of profile-guided AOT
code). If it trips: look for a re-rooted assembly (`TrimmerRootAssembly`), a dependency
bump growing the SDK, a new package reference, or a re-recorded AOT profile that got much
wider. The npm CI job runs the same script, so regressions surface at PR time.

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

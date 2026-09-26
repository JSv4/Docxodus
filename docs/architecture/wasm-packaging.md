# WASM Packaging

How the browser payload is built, trimmed, AOT-compiled, compressed, and kept small. This
is the reference for `wasm/DocxodusWasm/DocxodusWasm.csproj`, `scripts/build-wasm.sh`,
`scripts/record-aot-profile.sh`, and the size guardrail. (The original investigation/plan
doc was folded into this overview after implementation; measured numbers below are from
real builds, .NET SDK 10.0.301 / wasm-tools 10.0.109 / DocumentFormat.OpenXml 3.5.1,
2026-08 for trimming and 2026-09 for the AOT tier, except where a newer toolchain is
listed explicitly.)

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
`scripts/build-wasm.sh` prints and holds under a 5888 KiB (5.75 MiB) budget.
The table above records the original AOT rollout; the expanded browser profile
below measures 5.44 MiB Brotli on SDK 10.0.301.

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
profile* — selected by the representative workload — while everything else stays on the
interpreter (`AOTMode=LLVMOnlyInterp`, the SDK default). The profile is recorded with the
Mono AOT profiler over the same workload the steady-state spec times (`npm/tests/wasm-workload.ts`:
DocxDiff compare of a small pair and of a 147 KB legal form against an edited variant of
itself, DOCX→HTML on two documents, and the editor's per-mutation ReplaceText + single-block
re-render), plus the additional browser paths described below.

The recording spec also exercises a dense terminal-style paragraph through
`RawReplaceXml` and incremental rendering: 1,200 runs with repeated formats and
explicit line and character spacing. This covers the general formatting-template
optimization and batched identity assignment without depending on any game code.
That additional coverage initially measured **5,172,053 bytes (4.93 MiB) Brotli**
on SDK 10.0.301. After merging the history bridge and building with CI's SDK
10.0.400 / wasm-tools workload set 10.0.400.1 / runtime 10.0.11, the combined
payload reached 5128 KiB and exceeded the unchanged 5 MiB wire budget.

The recorder also exercises full synchronous and asynchronous paginated editor mounts,
conversion with headers/footers, anchor stamping, pagination and notes, and external
annotation-set creation (#783). HC031 and the sections-with-headers fixture train these
paths; NVCA is an additional benchmark input for them. The controlled A/B experiment
measured roughly **2–3x faster annotation creation**, with smaller 2–8% mount gains,
for **286.5 KiB more Brotli** (5.16 → 5.44 MiB). See
[the results and reproduction steps](../../benchmarks/aot-coverage/README.md).

Release builds now add `WasmOptConfigurationFlags Include="-Oz"` to the SDK's
post-link Binaryen pass. This optimizes the complete native WASM before the SDK
generates asset hashes. AOT bitcode compilation and linking retain their default
`-O2`; the recorded profile, exported APIs, and runtime features remain intact.
The combined release framework measures **5,214,762 bytes (4.973 MiB) Brotli**
on that CI toolchain, 28,118 bytes below the same budget. The setting applies to
all document workloads. Validate code generation changes with the trim canaries,
history clients/viewers, dense-text rendering tests, and steady-state workload.

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
./scripts/record-aot-profile.sh     # profiler build → browser run → shipped rebuild
```

The script publishes the **profiler flavour** (`-p:RunAOTCompilation=false
-p:WasmProfilers=aot`: AOT off, because AOT-compiled methods are invisible to the profiler),
runs `npm/tests/aot-profile-record.spec.ts` (opt-in via `DOCXODUS_RECORD_AOT_PROFILE=1`),
which drives the shared workloads in `test-harness.html?aotProfile=1` and writes
`INTERNAL.aotProfileData` to `wasm/DocxodusWasm/docxodus.aotprofile`, then rebuilds the
shipped configuration. It builds the JavaScript harness dependencies before recording
and regenerates the export assets and staged site for the final shipping bundle.
Commit the profile with the workload change. `DOCXODUS_AOT_PROFILE_OUT` can direct an
experimental recording to a separate file; the final rebuild still uses the checked-in
shipping profile. Re-record when the hot paths move — a new engine stage, a
renamed hot class, a runtime bump. A **stale profile costs speed, never correctness**: a method
missing from it simply runs interpreted, and a method it names that no longer exists is
skipped. The steady-state spec's numbers are how you notice drift.

The Mono AOT profiler records method compilation/preparation and generic instances
(there is no hotness threshold). The profile selects whole methods, including shared
helpers; it does not rank CPU cost or limit compilation to the branches executed.
Widen `wasm-workload.ts` or the supplemental recording workload when a new user-facing
path needs the tier, and expect the wire total to follow. These details are not obvious
from the SDK docs:

- **Profile changes must invalidate AOT outputs.** On SDK 10.0.301 / runtime 10.0.9,
  the compiler's incremental check ignores the profile and can reuse old code even
  when a different profile is selected. `AotProfile.targets` hashes the selected
  profile contents before AOT. A changed or missing stamp removes generated AOT
  bitcode, objects, trimming-token files and the IL-stripped assemblies (ILStrip's own
  check is mtime-only, and stripped IL bodies must match the compiled method set), while
  preserving native runtime sources. The stamp is written only after AOT succeeds, and the
  outer publish fails if the stamp does not match the selected profile, so a renamed SDK
  hook cannot silently ship code compiled for a previous profile. This applies to direct
  `dotnet publish` and the build scripts; an unchanged profile keeps incremental
  compilation. The regression tests run with `node --test scripts/aot-profile-cache.test.mjs`.

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

### Precise interpreter stack marking is off (issues #811, #695, #696, #779)

`DocxodusWasm.csproj` sets `MONO_INTERPRETER_OPTIONS=-precise` through a
`WasmEnvironmentVariable`, which lands in `_framework/dotnet.boot.js`. Mono reads it once, when
the interpreter initializes, and clears the interpreter's precise-GC option (on by default in
.NET 10: `mono_interp_opt` reads `0x1ff` in the shipped binary, `0xff` with the setting).
`npm/tests/wasm-runtime-config.spec.ts` pins that the runtime a page loads carries it.

**What the option did.** With precise marking on, every nursery collection first walks the
thread's chain of interpreter-to-native transition records ("LMFs", *last managed frame*) and
each record's interpreter frames, to learn which interpreter stack slots cannot hold references
(`interp_mark_no_ref_slots`, inlined into `interp_mark_stack`). In this mixed AOT + interpreter
build that chain can become cyclic. A native process would fault on the bad pointer
(dotnet/runtime#123573 is the same function crashing on `lmf = 0x40`); WebAssembly does not trap
on it, so the walk never ends. The call that happened to allocate — `previewBatch`,
`proveRedlineReversibility`, `verifyDeliverable`, a first comparison — sits at 100% CPU with no
exception, and because the JS thread is inside that synchronous export no timeout can fire.
With the option off, `interp_mark_stack` scans the interpreter stack conservatively (the
.NET 8 behaviour; .NET 9 introduced the option and made it the default) and never touches the
chain.

**Evidence.** The SDK's own `dotnet.native.js.symbols` is written *before* the post-link
`wasm-opt -Oz` pass renumbers functions, so for a Release build it names the wrong functions.
Correct names came from building with `-p:WasmRunWasmOpt=false -p:WasmEmitSymbolMap=true`,
writing that map into the binary as a wasm `name` section, and re-running the SDK's post-link
command by hand (`wasm-opt --enable-simd --enable-exception-handling --enable-bulk-memory -Oz
--strip-dwarf`) with binaryen's `--symbolmap=<file>` added; the output is function-for-function
identical to a normal build. Adding `-g` to `wasm-opt` instead changes the output (27,307 functions
against 21,904), and that build did not hang in 48 runs, so it cannot stand in for the shipped one.

- Every sample of every hang — 50 of 50 across two hangs, reached from different exports — is
  at the header of the LMF loop, below `collect_nursery → pin_from_roots →
  sgen_client_scan_thread_data → interp_mark_stack`. None is in the frame or slot loops inside
  it, so it is the chain itself that cycles.
- A two-author tracked-edit workload (twelve rounds of preview, commit, verify, accept, prove,
  diff) hung in 6 of 276 runs with precise marking on and 0 of 330 with it off.
- The #695/#696 sequence — one `DocxSession` revision read, then the first comparison of the
  page — built from the commit before the comparison warm-up existed hangs every time (3 of 3,
  over 200 s) and finishes in 0.5–0.7 s every time (3 of 3) with only this setting added.

**What it replaces.** #697 read the same stack as "the cold path pays a conservative root scan
proportional to the interpreter stack" and fixed each comparison export by running a tiny seed
comparison first; #779 applied that diagnosis to the first preview. (#779's own hang no longer
reproduces from its parent commit, so it is attributed here by mechanism, not by a red-green
run.) Our inference is that the seeds worked by moving the heap and code layout off a losing
configuration, which would also explain why the hang looked non-monotone in nursery size (4m,
6m and 12m hung; 8m and 16m did not) and why the #697 sequence no longer hangs on current builds
even without its seed. Either way they could not reach an export that had already run warm —
#811's hangs came rounds into a process. The per-export warm-ups are gone;
`DocumentComparer.Warmup` (the npm worker's `prepare()`) remains, as a latency tool only
(`ComparisonEngine`).

**What it does not fix.** The cycle in the transition chain is a runtime defect; this setting
removes the one walker of it that runs on every collection. Other walkers (exception stack
traces) still exist, and none has been seen to hang. When the runtime is upgraded, re-check
dotnet/runtime#123573 before dropping the setting.

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

`build-wasm.sh` computes the brotli wire total on every build and **fails above 5.75 MiB**
(5888 KiB; measured 5.44 MiB with the expanded browser AOT profile). If it trips:
look for a re-rooted assembly (`TrimmerRootAssembly`), a dependency bump growing the SDK, a
new package reference, or a re-recorded AOT profile that got much wider. The npm CI job runs
the same script, so regressions surface at PR time.

### Why the budget moved from 5 MB to 5.25 MB

The 5 MB line was set when the payload measured 4.76 MB, and two changes consumed the
remaining margin in quick succession. The dense-text profile coverage above took main from
4981 KB to 5105 KB, leaving 15 KB. Exposing the portable-history file controls to the
browser then added 46 KB: before that change the WASM history surface was
`read`/`updates`/`create`/`list`/`get`/`export`/`materialize`/`replay`/`resolveTime`/`restore`
only, and `exportArchive`, `importArchive`, `exportDocx`, `compare`, `operations`,
`getOperation` and `exportOperationProposal` pull the whole `.docxhistory` reader/writer —
ZIP entry walking, manifest and graph validation, the change-set codec — into the bundle for
the first time. That lands at 5151 KB.

Re-recording the profile without the dense-text workload was measured as an alternative: it
recovers 28 KB (5123 KB) — still over the old budget, and it costs the formatting-template
and batched-identity coverage that workload exists to provide. Paying 31 KB of wire for a
feature that ships in every other transport was the better trade. That line kept ~225 KB
of headroom at the time.

### Why the budget moved from 5.25 MiB to 5.75 MiB

The full-mount, conversion-options and annotation workloads add 286.5 KiB compressed,
bringing the framework to 5,704,974 bytes on SDK 10.0.301. Annotation creation improves
roughly 2–3x on the measured documents, including a longer warm-up check. The broader
profile is an accepted download/performance tradeoff; 5.75 MiB leaves about 315 KiB for
toolchain variation and subsequent changes. The benchmark records the original 5.25 MiB
gate failure rather than retroactively applying the new limit to those measurements.

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

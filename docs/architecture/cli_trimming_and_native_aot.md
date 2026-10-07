# Trimming and NativeAOT for the CLI binaries

Issue #903. This covers the `redline`, `docx2html` and `docx2oc` release binaries
(`tools/redline`, `tools/docx2html`, `tools/docx2oc`) that `publish.yml` attaches to each GitHub
Release. The WASM build is out of scope. It has its own trimming setup, described in
`wasm-packaging.md`.

## Summary

- The release now publishes the three CLIs **trimmed**, still as ReadyToRun single files:
  `-p:PublishTrimmed=true` was added to the existing `publish.yml` command. On linux-x64 each binary
  shrinks from 139 MB to 52–58 MB, and cold start stays the same or gets slightly faster.
- **NativeAOT** is the bigger win: binaries of 28–32 MB, starting 3–4× faster than ReadyToRun. It
  was measured and is runtime-safe on linux-x64, but it is **not** switched on. Its Linux binaries
  need a newer glibc, and the other three release platforms were not tested. See
  [What switching to NativeAOT would take](#what-switching-to-nativeaot-would-take).
- The one change trimming needed: `docx2oc` now serializes its export through a
  source-generated `System.Text.Json` context instead of reflection. Its JSON output is
  byte-identical on all 252 corpus documents.
- No library code changed. The whole-application trim and AOT analysis reports **zero** warnings
  for all three CLIs. The warnings a trimmed publish does print come from library code that the
  trimmer removes from these binaries.
- CI now publishes the three CLIs trimmed on Linux, Windows and macOS and runs each one, so a
  trimming break fails a pull request instead of a release.

## Where the warnings come from

A trimmed or AOT publish prints two different kinds of warning. Only the second kind says
anything about the shipped binary.

1. **Compile-time analyzer warnings in `Docxodus`.** Passing `PublishTrimmed` or `PublishAot`
   turns on the Roslyn trim and AOT analyzers for every project being built, the library included.
   The library is then recompiled and the analyzers flag every risky call in it, whether or not a
   CLI can reach it. These warnings are attributed to `Docxodus.csproj`. They appear only when the
   library is actually recompiled, which is why some logs show them and others don't. The library
   otherwise builds warning-free and fails Release builds on any warning, but `Docxodus.csproj`
   turns `TreatWarningsAsErrors` off when `PublishTrimmed` or `PublishAot` is set, so these never
   fail a CLI publish.
2. **Whole-application warnings from ILLink (trimming) and ILC (NativeAOT).** These analyze only
   the code the application can reach and are printed as `Trim analysis warning` /
   `AOT analysis warning`. The CLI projects inherit `TreatWarningsAsErrors=true` in Release, and
   ILLink's `ILLinkTreatWarningsAsErrors` defaults to that value. So any warning of this kind fails
   the release publish.

### Whole-application warnings (what the binaries actually contain)

Every build used `-p:TrimmerSingleWarn=false`, so each warning is reported at its own call site.

| CLI | Before this change | After |
|-----|--------------------|-------|
| `redline` | 0 | 0 |
| `docx2html` | 0 | 0 |
| `docx2oc` | 2: `tools/docx2oc/Program.cs(79)`, IL2026 (trim) and IL3050 (AOT), on `JsonSerializer.Serialize<T>(T, JsonSerializerOptions)` | 0 |

The `docx2oc` warning is real. Trimmed and AOT apps switch off reflection-based
`System.Text.Json` by default, so the untouched AOT binary failed on every export with *"Reflection-based
serialization has been disabled for this application"*. The fix is a source-generated
`ExportJsonContext` in `tools/docx2oc/Program.cs`, built with the same options as before (indented,
camelCase, nulls omitted). It also registers the runtime types that the `object`-typed
`OpenContractsAnnotation.AnnotationJson` property can hold. Its JSON is byte-identical to the old
reflection output on the whole corpus (see [Test evidence](#test-evidence)).

### Compile-time analyzer warnings in the library

There are 20 trim warnings and 12 AOT warnings, at 20 call sites. None is reachable from the three
CLIs. To check that, the CLIs were published trimmed without single-file bundling, and the trimmed
`Docxodus.dll` was searched for each method below. None survives in any of the three, while the
untrimmed library contains all of them.

| Code path | Sites | Warnings | Why the CLIs are safe |
|-----------|-------|----------|-----------------------|
| `PtOpenXmlUtil.GetPackage()` / `ExtractPackageFromFeature()`: walks the Open XML SDK's internal features chain by reflection to reach the `System.IO.Packaging.Package` | `PtOpenXmlUtil.cs` lines 45, 48, 51, 57, 78, 101, 112, 122, 125 | IL2026 ×1, IL2060 ×1, IL2075 ×7, IL3050 ×1 (`MakeGenericMethod`) | Its only caller is `OwnedPartRelationships.RestoreExactImageTopology`, which runs when a `DocxSession` undoes or redoes an image edit. No CLI does either. The trimmer removes both methods. A test build that printed a stack trace whenever `GetPackage()` ran recorded zero calls across the 957-run corpus. |
| Delivery bundle JSON (`DeliveryBundleManifest.ToJsonBytes`, `DeliveryBundleCanonicalJson.SerializePayload`, `DeliveryBundleValidationReport`, `DeliveryBundleVerifier`, `DeliveryPackageDeltaReport`, `DocxodusExportHostRenderer`) | 7 | IL2026 + IL3050 each | Delivery bundles are an MCP/library feature. No CLI reaches them. |
| `ExternalAnnotationManager.SerializeToJson` / `DeserializeFromJson` | 2 | IL2026 + IL3050 each | Annotation sets are not used by the CLIs. |
| `DocxSessionJson` (line 1182): fallback for wire values of unknown type | 1 | IL2026 + IL3050 | Session wire protocol, used by the bridges, not the CLIs. |
| `MutationTransactions` (line 439): quotes a duplicate property name in an error message | 1 | IL2026 + IL3050 | Transaction replay, not used by the CLIs. |

Because nothing here is reachable, **no library code was changed or annotated**. Two other routes
were considered and rejected:

- **`[RequiresUnreferencedCode]` on these paths.** The attribute must be repeated on every caller,
  and through `DocxSession` that would reach a large part of the public API, adding warnings for
  every library consumer.
- **Moving the reflection walk's ILLink descriptor into the library.** The descriptor is
  `wasm/DocxodusWasm/ILLink.Descriptors.xml`, which keeps the members `GetPackage()` reaches.
  Embedding it in `Docxodus.dll` would let the library suppress its own `GetPackage()` warnings
  honestly, but the CLIs gain nothing from it, because the method is not in them. The WASM build
  needs the descriptor because its `[JSExport]` surface does reach session undo and redo.

If a future CLI feature reaches one of these paths, the whole-application analysis will report it,
and the warning will fail both the CI step and the release publish. That is the moment to fix the
path. For JSON the fix is a source-generated context. For `GetPackage()` it is the WASM descriptor,
referenced through `TrimmerRootDescriptor`.

## Measurements

The method follows #855. Each command was timed as wall-clock time over 15 runs after two warm-up
runs, and the median is reported. All variants of one command were interleaved run by run, so
background load hits them equally. The fixtures are a one-paragraph document and a one-word edit of
it. The compare is `redline a.docx b.docx out.docx`, and each convert takes the one-page document
as input. The machine is an 8-core linux-x64 desktop on .NET SDK 10.0.301 (runtime 10.0.9). Other
work was running, with a load average of 6–8. The compressed row lands within 10% of #855's
numbers (139/663/404 ms), and so do ReadyToRun's `--version` and compare (50/234 ms), so the two
machines are comparable. Only the ReadyToRun convert is clearly faster here (136 ms against 212).

| Variant | Size | `--version` | one-paragraph compare (`redline`) | convert (`docx2html`) | export (`docx2oc`) |
|---|---|---|---|---|---|
| single file, compressed (before #855) | 43.5 MB | 130 ms | 624 ms | 384 ms | 418 ms |
| ReadyToRun, no compression (#855, the release before this change) | 139.0 MB | 50 ms | 212 ms | 136 ms | 173 ms |
| **ReadyToRun + trimmed (the release from this change)** | 51.5–57.9 MB | 49 ms | 193 ms | 127 ms | 132 ms |
| trimmed, no ReadyToRun | 24.8–26.3 MB | 107 ms | 920 ms | 536 ms | 562 ms |
| NativeAOT | 27.7–31.9 MB | 23 ms | 55 ms | 46 ms | 35 ms |

Notes:

- **Size** is the single executable, as uploaded. The size ranges cover the three tools: redline
  is the largest in each variant (57.9 MB trimmed, 31.9 MB AOT) and docx2oc the smallest. The three
  ReadyToRun binaries are all about 139.0 MB, and the compressed ones about 43.5 MB.
- **`--version`** is `redline --version`. The other two tools measure within about 20 ms of it
  in every variant.
- **ReadyToRun export** is `docx2oc` before the JSON change (reflection). With the
  source-generated context, the ReadyToRun export takes 146 ms.
- **Trimmed without ReadyToRun is slow.** It is smaller, but trimming rewrites the framework
  assemblies and drops the ReadyToRun code they normally ship with, so everything is JIT-compiled at
  startup. It is slower than even the compressed variant.
- **Larger inputs.** On bigger documents (7 interleaved runs each), NativeAOT stayed ahead even
  without tiered compilation. A `docx2html` conversion of the 147 KB `NVCA-Model-COI.docx` took
  836 ms with ReadyToRun, 1033 ms trimmed and 776 ms with AOT. A `docx2oc` export of the same file
  took 401, 428 and 248 ms.

## Test evidence

There is no CLI test project. The `Docxodus.Tests` suite calls the library directly, and CI only
builds the CLIs. So the binaries were checked end to end with a corpus run, which every variant ran
the same way:

- **`redline`**: 99 pairs. Inside each `TestFiles/WC/WCnnn-*` group, the first file was compared
  with each of the others. Each pair ran twice, plain and with `--detect-moves`, always with a fixed
  `--date-time` so the output is deterministic. Three extra pairs were added: the image document
  `HC042-Image-Png.docx` against `HC006-Test-01.docx` in both directions, and the 2 MB
  `WC-BodyBookmarks` pair. That makes 201 runs.
- **`docx2html`**: all 250 top-level `TestFiles/*.docx` plus the 2 in `TestFiles/DD/`. Each was
  converted twice, plain and with `--extract-images --track-changes --render-comments
  --render-footnotes --render-headers-footers`. That makes 504 runs.
- **`docx2oc`**: the same 252 documents.

That is 957 runs per variant. Each variant was tested as it ships: the executable alone, without the
`libSkiaSharp.so` that publishing leaves beside it (see below). Results:

| Variant | Runs that succeeded | Output compared with ReadyToRun |
|---------|---------------------|---------------------------------|
| ReadyToRun (reference) | 912 / 957 | — |
| ReadyToRun + trimmed | 912 / 957, the same runs | byte-identical: all 158 redline `.docx`, every HTML file and extracted image, all 252 JSON files |
| NativeAOT | 912 / 957, the same runs | byte-identical, the same files |

The trimmed and AOT `docx2oc` outputs come from the source-generated serializer. The reference
comes from the old reflection-based one, so the JSON comparison also covers the serializer change.
Two ReadyToRun runs on the same inputs produced identical bytes, so byte comparison is a fair test.

The 45 failures are the same in every variant and have nothing to do with trimming:

- **43 `redline` runs**, failing with *"Index was out of range"* in
  `IrMarkupRenderer.EmitGapArranged`, are a **pre-existing bug that the switch to ReadyToRun (#855)
  exposed**. The same compares succeed with the compressed pre-#855 binary, with a Debug build and
  with `DOTNET_ReadyToRun=0`. They fail whenever that method runs fully optimized: as ReadyToRun or
  NativeAOT code, or under `DOTNET_TieredCompilation=0`. Marking only `EmitGapArranged` with
  `[MethodImpl(MethodImplOptions.NoOptimization)]` makes them pass, and so does turning off only the
  JIT's induction-variable optimization (`DOTNET_JitEnableInductionVariableOpts=0`), which points to a
  runtime miscompile. The unit tests most likely miss it because a test run executes this method as
  unoptimized tier-0 code. It is not fixed here; it is tracked in #925.
- **2 `docx2html` runs**: `RA001-Tracked-Revisions-01/02.docx` with the full flag set fail with
  *"Duplicate attribute"* in every build, Debug included. Also pre-existing.

### `libSkiaSharp.so` is not in the release

`dotnet publish` places SkiaSharp's native library next to the executable, not inside the single
file. `publish.yml` uploads only the executable (`./publish/redline*`). That has been true since
before #855 and is unchanged here. `docx2html` does load the library when it is present (for example
on `DB007-Spec.docx`), yet the HTML from all 504 conversions was byte-identical with and without it.
The NativeAOT `docx2html` loads it the same way when it is present.

## What switching to NativeAOT would take

NativeAOT is runtime-safe on everything exercised on linux-x64, with zero warnings and
byte-identical output. It was still left off the release for these reasons:

- **Linux glibc requirement.** The AOT binary links against the build machine's C library and
  needs glibc 2.34 or newer. The ReadyToRun and trimmed binaries need 2.27. Building on
  `ubuntu-latest` would drop Ubuntu 20.04, Debian 11 and RHEL/Alma/Rocky 8. To avoid that, build
  inside an older-glibc container, such as the cross-build images the .NET team publishes for this
  purpose.
- **Untested platforms.** Only linux-x64 was run. NativeAOT compiles natively per platform, so the
  `win-x64`, `osx-x64` and `osx-arm64` jobs each need their own run against a smoke corpus before
  shipping. `osx-x64` would also be cross-compiled on the arm64 `macos-latest` runner.
- **Diagnostics.** AOT stack traces carry method names but no file or line numbers, so the
  `REDLINE_DEBUG` / `DOCX2HTML_DEBUG` / `DOCX2OC_DEBUG` traces get less useful.
- **No runtime escape hatch.** ReadyToRun code can be bypassed at run time with
  `DOTNET_ReadyToRun=0`, which works around the `EmitGapArranged` failure above. AOT code cannot,
  so that bug should be fixed before an AOT release.

Switching would mean adding `-p:PublishAot=true` (and dropping `PublishSingleFile` /
`PublishReadyToRun`, which do not apply), building on an older-glibc Linux image, and extending the
new CI step to run the AOT binaries on every platform.

## Reproducing

```bash
# Baseline (the #855 release shape) and the trimmed release shape
dotnet publish tools/redline/redline.csproj -c Release -r linux-x64 --self-contained true \
  -p:PublishSingleFile=true -p:PublishReadyToRun=true [-p:PublishTrimmed=true] \
  -p:TrimmerSingleWarn=false -o out/redline

# NativeAOT
dotnet publish tools/redline/redline.csproj -c Release -r linux-x64 -p:PublishAot=true \
  -p:TrimmerSingleWarn=false -o out/redline-aot

# Whole-application warnings only (the library analyzer warnings print as plain "warning ILxxxx")
grep -E "(Trim|AOT) analysis warning" publish.log
```

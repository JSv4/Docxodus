# Browser AOT coverage experiment — issue #783

This experiment compares the checked-in profile with a profile recorded over the
existing workload plus full editor mounts, conversions with headers/footers,
pagination and anchor stamping, and external annotation-set creation. It changes
the compiled method selection, with the same C# source and JavaScript bundles.

The expanded paths are now part of the standard recorder in
`npm/tests/aot-profile-record.spec.ts`. `DOCXODUS_AOT_PROFILE_OUT` lets an experiment
record to a separate file without replacing the shipping profile. The measurements
below were taken before adopting the expansion and raising the budget to 5.75 MiB.

## Results (2026-09-12)

The expanded profile substantially accelerates annotation-set creation. The other
measured paths improve by about 5–8%, except the large asynchronous mount, which
improves by only 2%. These modest changes should not be read as precise guarantees
from three repetitions on one machine. The annotation gain persists after a longer
warm-up, while the profile does not remove most editor-mount or anchor-indexing work.

Warm values below are medians of 15 samples per arm (calls 4–8 from each of three
fresh runtimes). Lower milliseconds are better.

| Document | Operation | Original (ms) | Expanded (ms) | Speedup |
|---|---|---:|---:|---:|
| HC031, 9 pages | Bare conversion | 261.3 | 244.0 | 1.07× |
| HC031, 9 pages | Conversion + headers/footers | 260.8 | 246.0 | 1.06× |
| HC031, 9 pages | Conversion + anchors | 300.3 | 278.1 | 1.08× |
| HC031, 9 pages | Paginated conversion | 304.0 | 287.3 | 1.06× |
| HC031, 9 pages | Create annotation set | 163.0 | 76.4 | 2.13× |
| HC031, 9 pages | Synchronous paginated mount | 423.2 | 401.4 | 1.05× |
| HC031, 9 pages | Asynchronous paginated mount | 606.2 | 561.2 | 1.08× |
| NVCA, 53 pages | Bare conversion | 1344.6 | 1275.1 | 1.05× |
| NVCA, 53 pages | Conversion + headers/footers | 1356.8 | 1295.3 | 1.05× |
| NVCA, 53 pages | Conversion + anchors | 1917.6 | 1802.5 | 1.06× |
| NVCA, 53 pages | Paginated conversion | 1872.1 | 1744.5 | 1.07× |
| NVCA, 53 pages | Create annotation set | 1392.1 | 429.1 | 3.24× |
| NVCA, 53 pages | Synchronous paginated mount | 2988.2 | 2825.4 | 1.06× |
| NVCA, 53 pages | Asynchronous paginated mount | 5654.7 | 5530.6 | 1.02× |

The smaller document was still warming through eight annotation calls. A follow-up
used 30 calls per runtime, three repetitions, and retained calls 16–30 (45 samples
per arm): **133.8 → 70.2 ms, 1.91× faster**.
This confirms a sustained benefit beyond the earlier warm-up curve.

First-call medians, measured after runtime boot:

| Document / operation | Original (ms) | Expanded (ms) |
|---|---:|---:|
| HC031 / Create annotation set | 1500.7 | 1047.1 |
| HC031 / Asynchronous paginated mount | 1893.5 | 1795.1 |
| NVCA / Create annotation set | 2699.4 | 1193.4 |
| NVCA / Asynchronous paginated mount | 7489.3 | 7160.4 |

**Size tradeoff:** the framework grows from 5,411,569 to 5,704,974
Brotli bytes (5.16 → 5.44 MiB), an increase of
286.5 KiB / 5.4%. The expanded build **fails the existing
5.25 MiB size gate** by 195.3 KiB. Native WASM grows from 8,485,229
to 10,345,936 bytes. The accepted shipping budget is now 5.75 MiB.

All **672 main-run outputs and 180 follow-up outputs** have matching hashes across
calls and profiles for each document/operation, after removing annotation timestamps.
Both mount modes retain their original page counts. HC031 also has identical
synchronous/asynchronous DOM hashes; NVCA has a mode-specific DOM difference in
both profiles, so the comparison here is each operation against itself across
the two AOT builds, not a general assertion of mount-mode parity.

Raw samples and generated summaries stay outside version control. The reproduction
commands below write them to `/tmp/docxodus-issue-783/`. Environment: .NET SDK
10.0.301, Mono browser-wasm runtime 10.0.9, Chromium 143.0.7499.4 / Playwright 1.57.0,
Linux on an Intel Core Ultra 7 258V (8 logical CPUs).

## Validation

- Expanded profile recorder: 1 passed. TypeScript source and test type-checks passed.
- Expanded build: **23 browser tests passed**, covering trimming canaries,
  asynchronous mounts, header/footer editing, incremental annotation overlays,
  and the existing steady-state/runtime checks.
- Original build: **3 steady-state/runtime tests passed** after restoring its
  runtime assets. Control medians were comparable: small compare 30.9 → 31.7 ms,
  heavy compare 1363.1 → 1384.9 ms, single-block refresh 16.4 → 15.1 ms. These
  single control runs do not establish small regressions or improvements.
- The expanded build's failure against the original size gate is recorded above.

The original experiment preserved the shipping profile. The follow-up change adopts
the expanded recording workload and profile, raises the size budget, and fixes profile
cache invalidation in `wasm/DocxodusWasm/AotProfile.targets`.

The final recorder was run end to end again on 2026-09-13. Its 23,591 method records
have the same selection, including generic instances, as the measured expanded profile.
The fresh shipping build is **5,707,354 bytes Brotli (5.44 MiB)**, within the new
5.75 MiB budget; the small byte-size difference from the original experiment does not
change the selected methods. **40 browser tests passed**, adding the worker suite to
the checks above. Eight cache-regression scenarios also passed. Real direct publishes
with expanded → original → expanded profiles reproduced the expected native hashes;
an unchanged publish reused AOT outputs and left the native hash unchanged. Both
expanded and unchanged publishes produced native SHA-256
`f4bdf7f5e5fabaad8a32ca1d25b333374488406dc5a3d87ae56d9332bce1553b`.

## Method

- Source: commit `0e01c0cf2a5ab40b482775167edfbe62176c9e64`.
- Added training inputs: `HC031-Complicated-Document.docx` and
  `DB002-Sections-With-Headers.docx`. Each exercises headers-only conversion,
  conversion with all three requested options and notes, annotation-set creation,
  synchronous paginated mount, and asynchronous paginated mount.
- Measurement inputs: HC031 (42,336 bytes, 9 rendered pages) and
  `NVCA-Model-COI.docx` (147,622 bytes, 53 rendered pages). NVCA was already used by
  the original compare/bare-conversion workload; it is not used to train the added
  browser paths.
- Three repetitions per document/operation/build. Each uses a fresh Chromium
  process and .NET runtime, then makes eight consecutive calls. The first call is
  reported separately. Calls 4–8 provide 15 warm samples across the repetitions.
- Build order alternates for each pair and repetition. Builds are served from
  immutable snapshots over loopback, with caching disabled. No builds or other
  test suites run during the measurements.
- First-call time starts after runtime boot. It does not include download or
  runtime initialization. Warm timings measure actual API completion, with editor
  mounts including layout and the asynchronous mount's yields.
- Conversion and annotation measurements call the raw engine bridge. They exclude
  the TypeScript wrapper's normalization and worker messaging. Each operation gets
  a fresh runtime, so its first-call conditions differ from the issue's shared,
  already-warm runtime. The reporter's original file and slower machine are not
  available; these are controlled comparisons on repository fixtures.
- Editor options: `paginated: true`, `editable: true`, `columnWidth: "section"`;
  asynchronous mounts use windows of 24 body units. Viewport: 1280 × 900.
- Every bridge call is timed. Output hashing, DOM inspection, and editor cleanup
  occur outside the measured interval. The annotation hash omits its creation and
  update timestamps. Full HTML/DOM and normalized annotation hashes must agree
  across calls and profiles.

## Build-cache finding

On SDK 10.0.301 / browser-wasm runtime 10.0.9, changing only
`WasmAotProfilePath` reused the cached AOT code: the log said
`Everything is up-to-date, nothing to precompile`, and `dotnet.native.wasm` was
byte-identical to the original. The AOT cache contains assembly file hashes but
does not include the `.aotprofile` input. Clearing the generated `for-publish`
directory forced the intended compilation. The AOT trimming-token manifests then
grew from 7,110 to 10,532 entries; this is distinct from the profile's raw method
record count, which includes generic instantiations.

The fix hashes profile contents before AOT and removes stale bitcode, object and
trimming-token outputs when the hash changes. It records the hash only after a
successful AOT compile and preserves incremental builds for an unchanged profile.
It applies to direct `dotnet publish` as well as the shell scripts. The experiment's
first cached rebuild was discarded and never used as the expanded measurement arm.

## Reproduce

The scripts assume dependencies and the wasm-tools workload are already installed.
Run these commands from the repository root. Snapshots refuse to overwrite an
existing arm; use a fresh directory for each experiment.

```bash
mkdir -p /tmp/docxodus-issue-783
git show 0e01c0cf2a5ab40b482775167edfbe62176c9e64:wasm/DocxodusWasm/docxodus.aotprofile > /tmp/docxodus-issue-783/original.aotprofile
./scripts/build-wasm.sh -p:WasmAotProfilePath=/tmp/docxodus-issue-783/original.aotprofile
npm --prefix npm run build:js
node benchmarks/aot-coverage/snapshot.mjs original /tmp/docxodus-issue-783/original.aotprofile

./scripts/build-wasm.sh -p:RunAOTCompilation=false -p:WasmProfilers=aot
npm --prefix npm run build:js
node npm/scripts/stage-web.mjs --tests
cd npm
DOCXODUS_RECORD_AOT_PROFILE=1 \
DOCXODUS_AOT_PROFILE_OUT=/tmp/docxodus-issue-783/expanded.aotprofile \
  ./node_modules/.bin/playwright test aot-profile-record.spec.ts --project=chromium --reporter=line
cd ..

# The project now invalidates stale AOT outputs when the profile changes.
./scripts/build-wasm.sh -p:WasmAotProfilePath=/tmp/docxodus-issue-783/expanded.aotprofile
node benchmarks/aot-coverage/snapshot.mjs expanded /tmp/docxodus-issue-783/expanded.aotprofile

node benchmarks/aot-coverage/run.mjs --out /tmp/docxodus-issue-783/results.json
node benchmarks/aot-coverage/summarize.mjs /tmp/docxodus-issue-783/results.json /tmp/docxodus-issue-783/summary.json
# Restore the checked-in shipping configuration after measuring variants.
npm --prefix npm run build
```

`run.mjs` also accepts `--root`, `--arms`, `--ops`, `--fixtures`, `--repetitions`,
`--iterations`, and `--warmup`. `DOCXODUS_CHROMIUM_PATH` selects an installed Chromium
when Playwright's default browser is unavailable.

The extra warm-up check is reproducible with:

```bash
node benchmarks/aot-coverage/run.mjs --fixtures HC031-Complicated-Document.docx --ops annotation.create --repetitions 3 --iterations 30 --warmup 15 --out /tmp/docxodus-issue-783/long-warmup.json
```

# Atomic typing latency (issue #788)

`executeBatch` produces a package-hashed receipt and can roll back arbitrary package
mutations. Interactive text plus typing formatting can instead use
`session.replaceMatch(match, replacement, format)`, which returns an ordinary edit
result and commits both changes as one undo/version unit using in-memory snapshots.

This benchmark compares those paths on the issue's exact NVCA fixture: replace
`document` with `new document` in the body paragraph beginning “The Certificate of
Incorporation”, then make the 12 replacement characters bold. It uses the core and
WASM directly, without rendering, React, or a registered page map.

## Measured results (2026-09-14)

Both paths used the same Release build and checked-in AOT profile, on Linux x64,
Intel Core Ultra 7 258V (8 logical CPUs), .NET SDK 10.0.301 / runtime 10.0.9,
and headless Chromium 143.0.7499.4, cross-origin isolated, without CPU throttling.

| Synchronous wall time | `executeBatch` | `replaceMatch(..., format)` |
| --- | ---: | ---: |
| First edit, two fresh contexts per path | 1,131.4–1,164.0 ms | 108.5–108.8 ms |
| Repeats after undo/redo, four samples per path | 958.6–970.2 ms | 47.4–53.8 ms |

All twelve attempts passed text, Bold, version, undo, and redo checks, with no browser
errors. The typing path made only `ReplaceTextAtSpanWithFormat`; the batch path still
returned its package hash. A separate freshly built main baseline (`e14bbedf`) reproduced
the report at 1,154.6–1,183.9 ms initially and 983.3–1,006.0 ms on repeats.

Validation: 4,394 .NET tests passed (three existing skips), 29 Chromium regression tests,
eight Python batch tests, TypeScript checks, and the Release WASM build/size gate passed.

## Reproduction

From the repository root, after installing npm dependencies and Playwright Chromium:

```sh
npm --prefix npm run build
node benchmarks/atomic-typing/run.mjs /tmp/atomic-typing.json
```

Both paths normally use the same build. To compare against a separately built or
extracted baseline package (containing `package.json` and `dist/`), set
`DOCXODUS_BATCH_ROOT=/path/to/baseline`. The typing path always uses this checkout's
`npm/dist`. Run measurements without concurrent builds or test suites.

The script alternates batch / typing / typing / batch in fresh browser contexts.
Each context measures the first edit and two repeats after undo/redo. Timers cover
only the synchronous editing call. Outside each interval it verifies text, bold,
one version increment, and exact text/format restoration through undo and redo;
the batch must also return its receipt hash. JSON includes individual bridge times,
fixture and runtime fingerprints, browser version, CPU, and browser errors.

These are diagnostic samples on one machine, not percentiles or a latency guarantee.
Full-package batch receipt costs remain; callers select the typing operation explicitly.

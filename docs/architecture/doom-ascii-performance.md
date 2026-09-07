The Doom ASCII renderer keeps all 320×200 source pixels, including the original
HUD and sprites. Each pixel retains its tone glyph. Ink runs use the existing
per-cell color-error bound, with a linear longest-prefix allocator instead of
a frame-wide merge heap. Dense document rendering reuses the conversion of
unchanged formatting templates; every frame still replaces the live DOCX text
and goes through the editor's normal refresh. Printable ASCII also omits unused
complex-script font properties. Browser tests verify identical rendered text
and styles with those properties present or absent.

Measured on 2026-09-07, Linux, Intel Core Ultra 7 258V, 1300×1100 viewport at
DPR 2. Each phase runs for 12 seconds after entering Freedoom E1M1. These are
unthrottled document frames observed at browser animation frames; multiple
mutations before presentation count once. Movement phases must also change
the document's content signature, so a frozen picture cannot pass.

| Workload | Previous Chromium build | Optimized Chromium | Chromium, 2× CPU slowdown | Optimized Firefox* |
| --- | ---: | ---: | ---: | ---: |
| Stationary room | 15.17 FPS | 26.90 FPS | 10.46 FPS | 16.83 FPS |
| Turning | 15.61 FPS | 25.69 FPS | 12.12 FPS | 16.65 FPS |
| Moving and firing | 17.69 FPS | 24.27 FPS | 14.64 FPS | 20.28 FPS |
| Turning and firing | 16.86 FPS | 27.00 FPS | 14.58 FPS | 19.83 FPS |

Chromium was 143.0.7499.4; Firefox was 144.0.2. The previous build was preserved
before the source changes. Gameplay depends on engine timing, so these are
matching input workloads rather than identical recorded framebuffer sequences.
CPU throttling is a synthetic headroom check, not a prediction for another
device. The Chromium p95 frame interval was 50.1 ms in all four optimized phases.

The benchmark verifies 200 rows, 64,000 source cells plus 200 row guards, 199
line breaks, no image/canvas/SVG surface, incremental refresh, and identical
text in the live DOM, session XML, source projection, and saved/reopened DOCX.
It saves `report.json`, `gameplay.png`, `source.png`, and `frame.docx` under
`npm/test-results/doom-ascii-benchmark/` by default.

Validation passed 12 native dense-render/plan tests, the demo logic and TypeScript
checks, and 11 Chromium browser checks covering contrast at DPR 1/2, source-cell
preservation, gameplay input, clipboard, history, save/reopen, and phone layout.
Firefox passed the DPR 2 color checks and document round-trip. Its DPR 1 contrast
and adjacent-row color checks fail with exactly the same measurements on both
the original and optimized builds; its synthetic clipboard check also fails on
both. These pre-existing Firefox limitations remain unresolved. The throughput
results do not imply visual parity across browsers at every display density.

To reproduce from the repository root:

```sh
npm --prefix npm run build
npm --prefix npm run pretest
python3 -m http.server 8082 --directory npm/dist/wasm
```

In another terminal:

```sh
npm --prefix npm run benchmark:doom-ascii
DOOM_BENCH_CPU=2 npm --prefix npm run benchmark:doom-ascii
```

Use `ARCADE_URL` for another server, `DOOM_BENCH_WAD` for another same-origin
IWAD, and `DOOM_BENCH_OUTPUT` for a different artifact directory. The benchmark
loads `./embed.bundle.js` so it tests the locally built engine. The demo pages'
default CDN pin continues to select the published release.

*Firefox automation requires care: the Firefox bundled with this checkout's
older Playwright attaches a debugger that disables optimized WebAssembly. That
produced an apparent 2.8–3.2 FPS baseline and 4.8–7.7 FPS after these changes.
Those measurements describe the debug tier, not ordinary Firefox execution.
The Firefox results in the table used a temporary copy of the same browser
with the upstream `allowUnobservedWasm` and `allowUnobservedAsmJS` flags applied
to its automation runtime. See the [upstream implementation](https://raw.githubusercontent.com/microsoft/playwright/main/browser_patches/firefox/juggler/content/Runtime.js).
The installed browser and repository dependencies were not modified.

To measure a matching Playwright Firefox build containing that fix, use
`DOOM_BENCH_BROWSER=firefox` and, when needed, `DOOM_BENCH_EXECUTABLE` to select
its executable. CPU throttling is supported only for Chromium.

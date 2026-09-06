# DOOM rendering continuation — 2026-09-06

This is an unfinished investigation handed from a local session to Codex Cloud at the user's request. The experimental patch is **not applied** to the demo. Remove or consolidate these temporary investigation assets when the follow-up is complete.

## Task and constraints

Continue **JSv4/Docxodus PR #721**, branch **codex/doom-ascii-rendering**. The user wants original DOOM running on the native document surface, a polished Reddit GIF showing title → menu selection → initial gameplay, PNGs, and an updated PR. Fix player-visible rendering, not merely the GIF encoder. Keep optimizations generally applicable; do not detect menus, hide the title transition, replace the game, or change the engine to make the capture look better.

Latest feedback: “new palette works and improves some things (particularly menu) but the gameplay is more washed out. The issue remains the Doom logo remains ghosted when menu opens.” `reported-menu.png` is the user's September 6, 12:14 screenshot. Earlier screenshot paths under `/home/jman/Pictures` will not exist in Cloud.

## Pushed baseline and CI

The production baseline is **381740a1106e6ed0d1f830bfba319caf38ea9aff**. All its GitHub checks passed, including automated review. The CI and Playwright workflows previously stopped because the combined main-branch history APIs + dense-text AOT profile exceeded the 5 MiB WASM wire budget. The fix adds Release-only `WasmOptConfigurationFlags Include="-Oz"` to the SDK post-link pass. It preserves default `-O2` compilation/linking, the full AOT profile, exported APIs, and the unchanged size gate.

Exact local CI toolchain: SDK 10.0.400, wasm-tools workload set 10.0.400.1, runtime 10.0.11. The integrated release measured 5,214,762 bytes Brotli locally. GitHub measured 5091 KiB against 5120 KiB. CI run 34046800799 passed; Playwright run 34046800796 passed with 722 passed, one passed on retry, and 11 skipped. The retry was the new screenshot contrast check: ratio error .06452 vs .06 tolerance. Twelve subsequent local repetitions passed with identical measurements; the assertion was not loosened and no speculative test fix was committed.

Focused verification also passed: all 16 Chromium checks in `demo-arcade-doom-ascii`, `history-client`, `history-live-viewer`, `trim-validation`, and `wasm-steady-state`; demo logic; TypeScript; both npm release-package boundaries. Warm medians: small compare 30.6 ms, heavy compare 1.34 s, editor refresh 15.3 ms.

## Findings so far

- Capture the framebuffer and both renderers while paused. `current-menu-source.png` is the exact raw 320×200 original DOOM frame. It already contains the large title-art DOOM logo behind the smaller menu logo and menu text. The ASCII projection makes their separation much less clear. Do not claim the large background logo is necessarily an uncleared prior frame, and do not dismiss the user's visual complaint.
- On title, menu, and gameplay frames, the authoritative document text matched `asciiFramebuffer()`. DOM text matched the XML after normalizing nonbreaking spaces. Every computed HTML run color matched its XML run color. Switching ASCII → Original → ASCII also reproduced the issue. This is evidence against stale DOM text/colors, not proof that all visual rendering is correct.
- The color merge's weighted RGB error bound of 80 permits a saturated red cell to become orange (`FF6438`). On one paused menu, estimated average per-channel error on red menu pixels was 20.75. Reducing the bound to 60 reduced it to 3.25; menu runs increased about 3025 → 5474, gameplay about 934 → 966. At 45, menu runs were 5921 and gameplay 1685. Measure actual performance and visual quality before choosing a bound.
- The current 23-character tone ramp has sparse low/middle coverage levels. A local experiment measured all 95 printable ASCII glyphs in the actual native editor at the authored font size, tracking, exact line pitch, bold, and DPR 2. The candidate uses that full measured table and a color merge bound of 60. It shows more dark gameplay texture but is **not yet accepted or fully tested**.
- Current calibration is tied to DPR 2. The user may be viewing DPR 1 or another effective scale. This remains an important unresolved hypothesis. `dpr1-*.png` are fresh baseline captures at DPR 1; the candidate has only been captured at DPR 2 so far. Check real display density and zoom rather than judging only enlarged high-density screenshots.

## Files and reproduction

Production: `docs/demo/doom-ascii.js` (projection), `doom-cart.js` (engine framebuffer and PNG), `ascii-scenes.js` (frame XML and font pin), and `ascii-arcade.js` (driver). The generic dense-text converter lives in the C# library. Relevant regression files are under `docs/demo/tools/` and `npm/tests/`.

`compare.mjs` captures title, menu, and initial gameplay from the original shareware IWAD, saves raw source PNGs and both surface renderings, and records text/color comparisons. It defaults to DPR 2; set `DPR=1` for an ordinary-density screen. `PROJECTION_PATH` can point to an experimental JS module without replacing the shipped projection. Output defaults to `/tmp/docxodus-render-compare`; choose `OUTPUT_DIR` for separate runs.

`calibrate.mjs` measures all printable glyphs, one fully visible 200-row native paragraph per glyph. It saves normalized coverage to `/tmp/docxodus-glyph-coverage.json`; set `COVERAGE_PATH` and `DPR` as needed. A first attempt to capture all samples in one tall paragraph was invalid because the editor scroll container clipped the later rows. Only the per-frame calibration is included here.

`candidate.patch` records the unapplied prototype. `glyph-coverage-dpr2.json` is its measured table. Candidate screenshots are comparisons, not release media. Keep the runtime/glyph dimensions unchanged unless measurements justify changing them.

Use `.github/workflows/playwright.yml` for the full setup. The cloud environment needs .NET 10 with wasm-tools, Node 22, Playwright Chromium/Firefox, and ffmpeg for media. Build and stage the browser package with `npm --prefix npm ci`, `npm --prefix npm run build`, and `npm --prefix npm run pretest`. Fetch the original shareware data using `node npm/scripts/fetch-doom-iwad.mjs --shareware`, then serve `npm/dist/wasm` on port 8082. The pinned external engine needs network access; see the production capture script's digest-checked mirror option if offline. No engine or IWAD binaries are committed.

Run `node docs/diagnostics/doom-ascii-handoff/compare.mjs` or `calibrate.mjs` after staging. The comparison's `domText` field does not normalize NBSP and can read false even when the text is identical after whitespace normalization; its `wrongXmlColors` includes invisible spaces whose run ink is coalesced. Compare visible glyph colors and normalize text before interpreting those fields.

## Finish the work

1. Compare source, Original, and ASCII at the same paused frame at DPR 1 and 2; resolve lost menu edge contrast and washed gameplay without game-specific rules.
2. Validate the chosen general change with meaningful dark-tone and contrasting-edge regressions, same-frame projection, copy/paste, save/reopen, undo/redo, phone geometry, and representative performance. Keep the WASM size gate green.
3. Regenerate the original DOOM title → menu → gameplay GIF, MP4, PNGs, DOCX frame, and metadata using `npm/scripts/capture-doom-ascii.mjs`. Preserve real timing and opaque GIF updates. Do not wait for an attract-demo backdrop to conceal the reported transition.
4. Update PR #721 and its preview/validation details; push the final fix and check CI. No merge or public deployment has been requested. The demo's CDN pin is still 12.1.0, so local captures use `?engine=./embed.bundle.js` until a library release and demo pin update happen.

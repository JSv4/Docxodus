import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gzipSync } from 'node:zlib';

// Gzip budgets for the bundles browsers download from the CDN and the hosted demos. Each is the
// measured minified size plus ~15% headroom, so an accidentally unminified build (roughly 1.7x
// larger) or a large new dependency fails here instead of shipping. Raise a budget deliberately,
// in the same change that needs it. export-browser.bundle.js is absent on purpose: it runs inside
// @docxodus/export's pinned Chromium, where download size does not matter.
const budgetsKiB = {
  // Complete margin-comment flow adds 0.65–0.67 KiB over main 784c0335:
  // embed/iife/editor/pagination 144.114/144.177/123.390/40.976 ->
  // 144.779/144.839/124.060/41.630 KiB. Raise only exhausted gates to whole KiB.
  // Continuous-section geometry and page-context float measurement exhaust these three gates.
  // Raise only those budgets to the next whole KiB; the paginator still fits its 41 KiB gate.
  // Page-owned STYLEREF measurement/reflow adds 1.15–1.20 KiB over parent 1f7dc8dd:
  // embed/iife/editor/pagination 141.922/141.928/121.294/38.779 ->
  // 143.116/143.126/122.441/39.978 KiB. Round only these exhausted gates to whole KiB.
  // Authored-clipping proofs add 1.12–1.18 KiB on main a6996d45:
  // embed/iife/editor/pagination 140.630/140.659/119.954/37.468 ->
  // 141.770/141.777/121.128/38.611 KiB. Round only exhausted gates up to whole KiB.
  // Linked-story flow with frame bounds and keep checks adds 1.02–1.07 KiB on main c84cf2ad:
  // embed/iife/editor/pagination 139.258/139.293/118.623/36.097 ->
  // 140.328/140.359/119.646/37.161 KiB. Raise only these exhausted gates by 1 KiB.
  // The balloon appendix (balloonOverflow) adds about 1.5 KiB to every bundle that carries the
  // paginator. Embed/iife/editor/pagination measured 133.3/133.4/112.8/31.8 KiB before it and
  // 135.0/135.1/114.5/33.4 KiB with it; each budget rose by 2 KiB.
  // Body-table fitting on current main measures 137.128/137.164 KiB; retain a 138 KiB gate.
  // Zero-font source coverage brings the ESM bundle just above 138 KiB.
  // Whole inline-layout continuation adds 0.25 KiB (138.947 -> 139.196 on main).
  'embed.bundle.js': 145,
  // Floating-picture height measurement brings this bundle just above 138 KiB.
  // The same whole-box continuation moves IIFE from 138.989 to 139.234 KiB.
  'embed.iife.js': 145,
  // Oversized merged-grid recovery adds 0.39 KiB (117.832 -> 118.220 on current main).
  'editor.bundle.js': 125,
  // Whole inline-image continuation adds about 0.25 KiB to the shared paginator.
  // Oversized floating-picture fitting adds 0.20 KiB here (34.902 -> 35.105).
  // Whole inline-layout continuation adds 0.24 KiB (35.801 -> 36.038 on main).
  // Independent textbox source inventories add 45 compressed bytes (39.978 -> 40.021 KiB).
  'pagination.bundle.js': 42,
  'session.bundle.js': 9,
  'docxodus.worker.js': 6,
  'worker-proxy.bundle.js': 4,
};

const dist = join(dirname(dirname(fileURLToPath(import.meta.url))), 'dist');
const failures = [];
for (const [file, budget] of Object.entries(budgetsKiB)) {
  const kib = gzipSync(readFileSync(join(dist, file)), { level: 9 }).length / 1024;
  const line = `${file.padEnd(24)} ${kib.toFixed(1).padStart(6)} KiB gzip (budget ${budget} KiB)`;
  console.log(line);
  if (kib > budget) failures.push(line);
}
if (failures.length > 0) {
  console.error(`\nBundle size budget exceeded:\n${failures.join('\n')}`);
  process.exit(1);
}

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
  // Authored textbox ink verification adds 410/416/432 bytes to embed/IIFE/pagination
  // (145.345/145.401/42.213 KiB). Raise only these exhausted gates by 0.5 KiB.
  // Shared inline-frame flow/exclusions add about 1.55 KiB on main 97c304f7:
  // embed/IIFE/editor/pagination 146.882/146.966/126.173/43.767 KiB.
  // Raise only the four exhausted gates; the other bundles are unchanged.
  // Recursive endnote-comment splitting adds 0.4–0.5 KiB over 647aee2d:
  // embed/IIFE/editor/pagination 147.3/147.4/126.6/44.2 KiB. Raise these gates by 0.5 KiB.
  // Frame line intervals add 1.72–1.79 KiB over bf76d0da:
  // embed/IIFE/editor/pagination 149.133/149.151/128.362/45.982 KiB.
  // Round only the four exhausted gates up to the next 0.5 KiB.
  // Fixed-anchor frame clearance adds about 0.45 KiB over e3864e55.
  // Raise only the four exhausted gates by 0.5 KiB.
  // Tracked-paragraph fragmentation with leading drawings adds less than 0.4 KiB over 10531183.
  // Round only the three exhausted gates up by 0.5 KiB; pagination remains below 47 KiB.
  // Shared column fragmentation adds 531–565 gzip bytes over 563a4c37.
  // Round the four exhausted paginator-carrying gates up to the next 0.5 KiB.
  // Split-paragraph editor repaint/focus handling adds 194–195 gzip bytes.
  // Round only the two exhausted embed gates up to the next 0.5 KiB.
  // Authored-table clipping diagnostics add 583–615 gzip bytes over dc973c45.
  // Round only the four exhausted paginator-carrying gates to the next 0.5 KiB.
  // Authored-paper coordinates atop oversized keeps add 157 editor / 143 pagination bytes
  // over 3a958cc5: 134179 / 49703 bytes, exceeding the gates by 35 / 39 bytes.
  // Add 0.1 KiB only to those two exhausted gates; retain both embed budgets.
  // Whole-equation fragmentation and its layout guards measure 155718/155750/134411/49939
  // gzip bytes (embed/IIFE/editor/pagination). Raise only exhausted gates to the next 0.1 KiB.
  // Resizing page-offset proofs measure 155976/156009/134666/50174 bytes. Raise these
  // exhausted gates; leave roughly 0.1 KiB above the paginator's 48.998 KiB result.
  // Single-line grid slots measure 156378/156432/135049/50561 gzip bytes.
  // Raise only these four exhausted gates to the next 0.1 KiB.
  // Group-member placement proof measures 156790/156839/135472/51015 gzip bytes.
  // Round only the four exhausted gates up to the next 0.1 KiB.
  'embed.bundle.js': 153.2,
  // Floating-picture height measurement brings this bundle just above 138 KiB.
  // The same whole-box continuation moves IIFE from 138.989 to 139.234 KiB.
  'embed.iife.js': 153.2,
  // Oversized merged-grid recovery adds 0.39 KiB (117.832 -> 118.220 on current main).
  'editor.bundle.js': 132.3,
  // Whole inline-image continuation adds about 0.25 KiB to the shared paginator.
  // Oversized floating-picture fitting adds 0.20 KiB here (34.902 -> 35.105).
  // Whole inline-layout continuation adds 0.24 KiB (35.801 -> 36.038 on main).
  // Independent textbox source inventories add 45 compressed bytes (39.978 -> 40.021 KiB).
  // Export-only table layout stays out of this bundle; its fragment callback brings the shared
  // paginator to 47,629 bytes (46.513 KiB), 13 bytes over the old gate. Round only this gate up.
  // Page-remainder table splitting and its keep guards add 203 gzip bytes over 9dd93836.
  // Round only this exhausted gate up by 0.5 KiB (CI measures 47.1 KiB).
  'pagination.bundle.js': 49.9,
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

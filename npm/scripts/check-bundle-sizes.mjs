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
  // Floating-paragraph admission measures 156920/156958/135611 gzip bytes
  // for embed/IIFE/editor. Round only these exhausted gates up to the next 0.1 KiB.
  // Natural minimum-line recovery adds 255–275 gzip bytes over a376a611.
  // Round only the four exhausted gates to the next 0.1 KiB.
  // Fragment-owned picture exclusions and collision checks measure 157968/157997/136673/52145
  // gzip bytes. Round only these four exhausted gates to the next 0.1 KiB.
  // Uniform cell/list minimum recovery adds 92–105 gzip bytes over d6579170:
  // 158062/158089/136778/52244 bytes. Round these exhausted gates to the next 0.1 KiB.
  // Plain story minimum recovery and shared natural paint add 76–85 gzip bytes over 7c0a32b1:
  // 158154/158178/136878/52335 bytes. Raise the four exhausted gates by 0.1 KiB.
  // Attached footnote preparation measures 158251/158277/136980/52429 gzip bytes.
  // Round only these four exhausted gates to the next 0.1 KiB.
  // Validated manual-width atoms measure 158815/158851/137562/52995 gzip bytes.
  // Raise only these four exhausted gates; the fitter itself stays export-only.
  // Small uniform minima add 79–82 gzip bytes: 158896/158935/137642/53075.
  // Round only the three exhausted gates to the next 0.1 KiB.
  // Shared mixed-line metrics add 2824–2877 gzip bytes over a9e18754:
  // embed/IIFE/editor/pagination 161748/161812/140503/55899 bytes.
  // Round only these four exhausted gates to the next 0.1 KiB.
  // Word-confirmed minimum baseline placement adds 90–108 gzip bytes:
  // embed/IIFE/editor/pagination 161840/161902/140611/55989 bytes.
  // Round only these exhausted gates to the next 0.1 KiB.
  // Shared declared font stacks in balloons/appendices add 61/50 gzip bytes here and
  // in pagination (161934/56039 bytes). Raise only these exhausted gates by 0.1 KiB.
  // Per-line marker, break and alignment metrics (#1565) measure 158.5/158.5/137.7/55.0 KiB
  // for embed/IIFE/editor/pagination on CI. Round only these exhausted gates to the next 0.1 KiB.
  // The last line's text-height fit (#1571) measures 162561/162632/141268/56624 gzip bytes
  // for embed/IIFE/editor/pagination. Round only these exhausted gates to the next 0.1 KiB.
  // The empty section mark's page-break placement (#1572) measures 162610/162688/141324/56676.
  // Embed keeps one byte and editor/pagination are over; round these three to the next 0.1 KiB.
  // Mapped symbol markers' ascent boxes and break-offset spans (#1567, #1565) measure
  // 162690/162780/141404/56766 gzip bytes. Round only these exhausted gates to the next 0.1 KiB.
  // Condensed justified lines (#1550) measure 162717/162811/141440/56806 gzip bytes.
  // Embed and editor are over; round only these two to the next 0.1 KiB.
  // The balloon setting and markup-sheet geometry (#1552, #1559, #1577), minimum-height and aligned
  // tall-row splitting and linked text-box overflow (#1558) and section-grid line heights (#1551),
  // merged together, measure 163388/163455/142035/57417 gzip bytes for embed/IIFE/editor/pagination.
  // Round only these four exhausted gates to the next 0.1 KiB.
  'embed.bundle.js': 159.6,
  // Floating-picture height measurement brings this bundle just above 138 KiB.
  // The same whole-box continuation moves IIFE from 138.989 to 139.234 KiB.
  // The appended bullet-normalization option adds 29 gzip bytes (161996 -> 162025).
  // Only IIFE exhausts its gate; increase it by 0.1 KiB.
  'embed.iife.js': 159.7,
  // Oversized merged-grid recovery adds 0.39 KiB (117.832 -> 118.220 on current main).
  'editor.bundle.js': 138.8,
  // Whole inline-image continuation adds about 0.25 KiB to the shared paginator.
  // Oversized floating-picture fitting adds 0.20 KiB here (34.902 -> 35.105).
  // Whole inline-layout continuation adds 0.24 KiB (35.801 -> 36.038 on main).
  // Independent textbox source inventories add 45 compressed bytes (39.978 -> 40.021 KiB).
  // Export-only table layout stays out of this bundle; its fragment callback brings the shared
  // paginator to 47,629 bytes (46.513 KiB), 13 bytes over the old gate. Round only this gate up.
  // Page-remainder table splitting and its keep guards add 203 gzip bytes over 9dd93836.
  // Round only this exhausted gate up by 0.5 KiB (CI measures 47.1 KiB).
  'pagination.bundle.js': 56.1,
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

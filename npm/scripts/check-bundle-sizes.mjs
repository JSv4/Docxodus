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
  // Whole inline-image atoms add ~0.25 KiB to the shared paginator: embed/iife/editor/
  // pagination measured 136.950/136.991/116.339/33.813 KiB before and
  // 137.195/137.229/116.591/34.065 KiB after. Raise only the exhausted budgets by 1 KiB.
  // Bookmark-aware table wrappers add 0.044 KiB (137.986 -> 138.030); restore 1 KiB headroom.
  // Combined continuation and frame fixes measure 139.140/139.183/118.505 KiB.
  'embed.bundle.js': 140,
  // Empty tab continuation adds 0.037 KiB here (137.979 -> 138.016); restore 1 KiB headroom.
  'embed.iife.js': 140,
  // Including floated-picture geometry adds 0.158 KiB here (116.941 -> 117.099).
  'editor.bundle.js': 119,
  // Oversized merged-grid recovery adds 0.389 KiB (34.870 -> 35.259); retain a tight 36 KiB gate.
  // Shared presentation preservation measures 36.036 KiB gzip in the combined renderer.
  'pagination.bundle.js': 37,
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

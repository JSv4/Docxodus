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
  // The balloon appendix (balloonOverflow) adds about 1.5 KiB to every bundle that carries the
  // paginator. Embed/iife/editor/pagination measured 133.3/133.4/112.8/31.8 KiB before it and
  // 135.0/135.1/114.5/33.4 KiB with it; each budget rose by 2 KiB.
  // Body-table fitting on current main measures 137.128/137.164 KiB; retain a 138 KiB gate.
  // Zero-font source coverage brings the ESM bundle just above 138 KiB.
  'embed.bundle.js': 139,
  // Floating-picture height measurement brings this bundle just above 138 KiB.
  'embed.iife.js': 139,
  // Individually fitted inline pictures bring the current editor to 117.2 KiB gzip.
  'editor.bundle.js': 118,
  // Whole inline-image continuation adds about 0.25 KiB to the shared paginator.
  'pagination.bundle.js': 35,
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

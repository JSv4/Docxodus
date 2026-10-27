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
  // The balloon layout's appendix of cut-short comments (balloonOverflow) adds about 1.5 KiB to
  // every bundle that carries the paginator; before it, these measured 133.4/112.8/31.8 KiB.
  'embed.bundle.js': 137,
  'embed.iife.js': 137,
  // Main 44ebb107 measures 111.6/30.8 KiB for editor/pagination. Word-matched kept-note
  // retries and continuations (#1373, #1374) add about 0.5 KiB to each; allow that reviewed
  // growth and variation between gzip implementations without changing the other budgets.
  'editor.bundle.js': 115,
  'pagination.bundle.js': 34,
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

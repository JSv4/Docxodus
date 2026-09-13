// Issue #783: alternate immutable builds, with a fresh Chromium process/runtime
// for each operation. Output includes every sample, bridge timings, and output hashes.
import { createServer } from 'node:http';
import { createReadStream } from 'node:fs';
import { mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { cpus, platform, release } from 'node:os';
import { dirname, extname, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const require = createRequire(join(repo, 'npm/package.json'));
const { chromium } = require('playwright');
const { transform } = require('esbuild');
const args = new Map();
for (let i = 2; i < process.argv.length; i += 2) args.set(process.argv[i], process.argv[i + 1]);
const bundleRoot = resolve(args.get('--root') ?? '/tmp/docxodus-issue-783/bundles');
const out = resolve(args.get('--out') ?? '/tmp/docxodus-issue-783/results.json');
const arms = (args.get('--arms') ?? 'original,expanded').split(',');
const operations = (args.get('--ops') ?? [
  'html.bare', 'html.headers', 'html.anchors', 'html.paginated',
  'annotation.create', 'editor.open', 'editor.openAsync',
].join(',')).split(',');
const fixtures = (args.get('--fixtures') ?? 'HC031-Complicated-Document.docx,NVCA-Model-COI.docx').split(',');
const repetitions = Number(args.get('--repetitions') ?? 3);
const iterations = Number(args.get('--iterations') ?? 8);
const warmup = Number(args.get('--warmup') ?? 3);
const source = await readFile(join(repo, 'npm/tests/wasm-browser-workload.ts'), 'utf8');
const { code } = await transform(source, { loader: 'ts', format: 'esm', target: 'es2022' });
const { runWasmBrowserWorkload } = await import('data:text/javascript;base64,' + Buffer.from(code).toString('base64'));

const mime = { '.html': 'text/html', '.js': 'text/javascript', '.wasm': 'application/wasm', '.json': 'application/json' };
const server = createServer(async (request, response) => {
  try {
    const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    if (pathname === '/favicon.ico') { response.writeHead(204).end(); return; }
    const path = resolve(bundleRoot, '.' + pathname);
    if (!path.startsWith(bundleRoot + sep) || !(await stat(path)).isFile()) {
      response.writeHead(404).end(); return;
    }
    response.writeHead(200, { 'Content-Type': mime[extname(path)] ?? 'application/octet-stream', 'Cache-Control': 'no-store' });
    createReadStream(path).pipe(response);
  } catch { response.writeHead(404).end(); }
});
await new Promise((resolve, reject) => {
  server.once('error', reject);
  server.listen(0, '127.0.0.1', resolve);
});
const port = server.address().port;
const results = {
  timestamp: new Date().toISOString(),
  environment: { cpu: cpus()[0].model, logicalCpus: cpus().length, platform: platform(), release: release(), node: process.version },
  method: { arms, repetitions, iterations, warmup, freshBrowserPerCase: true, alternatedArmOrder: true, viewport: { width: 1280, height: 900 } },
  bundles: {}, fixtures: {}, runs: [],
};
const median = values => {
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
};
await mkdir(dirname(out), { recursive: true });
try {
  for (const arm of arms) results.bundles[arm] = JSON.parse(await readFile(join(bundleRoot, arm, 'build.json'), 'utf8'));
  if (results.bundles.original && results.bundles.expanded) {
    for (const key of ['sourceCommit', 'sdk', 'editorSha256', 'sessionSha256']) {
      if (results.bundles.original[key] !== results.bundles.expanded[key]) throw new Error(`Builds differ in ${key}`);
    }
    if (results.bundles.original.nativeSha256 === results.bundles.expanded.nativeSha256) {
      throw new Error('Native binaries are identical: clear the AOT build cache before measuring the new profile');
    }
  }
  for (const fixture of fixtures) {
    const bytes = await readFile(join(repo, 'TestFiles', fixture));
    results.fixtures[fixture] = { bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') };
  }
  for (let repetition = 0; repetition < repetitions; repetition++) {
    for (let f = 0; f < fixtures.length; f++) {
      const fixture = fixtures[f];
      const bytes = Array.from(await readFile(join(repo, 'TestFiles', fixture)));
      for (let o = 0; o < operations.length; o++) {
        const operation = operations[o];
        const order = (repetition + f + o) % 2 ? [...arms].reverse() : arms;
        for (const arm of order) {
          const browser = await chromium.launch({
            headless: true,
            ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
          });
          try {
            results.environment.chromium = browser.version();
            const page = await browser.newPage({ viewport: results.method.viewport });
            const errors = [];
            page.on('pageerror', error => errors.push(String(error)));
            page.on('console', message => { if (message.type() === 'error') errors.push(message.text()); });
            await page.goto(`http://127.0.0.1:${port}/${arm}/test-harness.html`);
            await page.waitForFunction(() => window.DocxodusReady === true, { timeout: 60000 });
            let timer;
            let samples;
            try {
              samples = await Promise.race([
                page.evaluate(runWasmBrowserWorkload, { bytes, operation, iterations }),
                new Promise((_, reject) => {
                  timer = setTimeout(() => reject(new Error('operation exceeded 180 seconds')), 180000);
                }),
              ]);
            } finally { clearTimeout(timer); }
            if (errors.length) throw new Error(errors.join('\n'));
            results.runs.push({ arm, repetition, fixture, operation, samples });
            await writeFile(out, JSON.stringify(results, null, 2) + '\n');
            const warm = samples.slice(warmup);
            console.log(`${arm.padEnd(8)} rep=${repetition + 1} ${fixture} ${operation}: first=${samples[0].wallMs.toFixed(1)} ms warm=${warm.length ? median(warm.map(s => s.wallMs)).toFixed(1) : '-'} ms`);
          } finally { await browser.close(); }
        }
      }
    }
  }
} finally { server.close(); }
console.log(`Wrote ${results.runs.length} cases to ${out}`);

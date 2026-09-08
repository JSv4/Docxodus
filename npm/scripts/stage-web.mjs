import { cp, copyFile, mkdir, readdir, rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const npmRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const repoRoot = dirname(npmRoot);
const dist = join(npmRoot, 'dist');
const webroot = join(dist, 'wasm');
const site = join(dist, 'site');
const demo = join(site, 'demo');

// One asset inventory for the local examples, tests and Pages. No source bundling,
// downloads or version substitutions happen here: these are the package's built bytes.
const runtime = [
  ['pagination.bundle.js', 'pagination.bundle.js'],
  ['session.bundle.js', 'session.bundle.js'],
  ['editor.bundle.js', 'editor.bundle.js'],
  ['embed.bundle.js', 'embed.bundle.js'],
  ['docxodus.worker.js', 'docxodus.worker.js'],
  ['worker-proxy.bundle.js', 'worker-proxy.js'],
  ['export-browser.bundle.js', 'export-browser.js'],
  ['export-assets.json', 'export-assets.json'],
];

await mkdir(webroot, { recursive: true });
// Incremental staging must retire files left by older builds, too.
for (const retired of ['freedoom-e1m1.js', 'tools/wad2cart.mjs', 'tools/wad2cart.test.mjs', 'history-example.js']) {
  await rm(join(webroot, retired), { force: true });
}
await rm(site, { recursive: true, force: true });
await cp(join(repoRoot, 'docs'), site, { recursive: true });
await mkdir(join(demo, 'wasm'), { recursive: true });
await cp(join(webroot, '_framework'), join(demo, 'wasm', '_framework'), { recursive: true });
for (const [source, target] of runtime) {
  await copyFile(join(dist, source), join(webroot, target));
  await copyFile(join(dist, source), join(demo, target));
}
for (const entry of await readdir(join(repoRoot, 'docs', 'demo'), { withFileTypes: true })) {
  if (entry.name === 'README.md') continue;
  const target = entry.name.endsWith('.html') && entry.name !== 'player.html'
    ? `demo-${entry.name}` : entry.name;
  await cp(join(repoRoot, 'docs', 'demo', entry.name), join(webroot, target), { recursive: true });
}
for (const entry of await readdir(join(npmRoot, 'examples'))) {
  if (entry.endsWith('.html')) await copyFile(join(npmRoot, 'examples', entry), join(webroot, entry));
}
// Preserve the old local example URL while sharing its entire implementation.
await copyFile(join(npmRoot, 'examples', 'editor.html'), join(webroot, 'history.html'));
await copyFile(join(repoRoot, 'docs', 'demo', 'docxodus-demo-guide.docx'), join(webroot, 'sample.docx'));
if (process.argv.includes('--tests')) {
  for (const name of ['test-harness.html', 'worker-test-harness.html', 'profiling-harness.html', 'standalone-export-harness.html']) {
    await copyFile(join(npmRoot, 'tests', name), join(webroot, name));
  }
}
console.log('Staged local editors in dist/wasm and the deployable site in dist/site (same package build).');

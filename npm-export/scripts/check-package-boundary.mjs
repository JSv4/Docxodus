import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const cache = mkdtempSync(join(tmpdir(), 'docxodus-export-pack-'));

try {
  const packed = spawnSync(
    npmCommand,
    ['pack', '--dry-run', '--json', '--ignore-scripts'],
    {
      cwd: packageRoot,
      encoding: 'utf8',
      env: { ...process.env, npm_config_cache: cache },
    },
  );
  if (packed.error) throw packed.error;
  assert.equal(packed.status, 0, `npm pack audit failed:\n${packed.stderr || packed.stdout}`);
  const [manifest] = JSON.parse(packed.stdout);
  const paths = manifest.files.map(({ path }) => path);
  const packageJson = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8'));

  assert.equal(packageJson.license, 'MIT');
  assert.equal(
    readFileSync(join(packageRoot, 'LICENSE'), 'utf8'),
    readFileSync(join(packageRoot, '..', 'LICENSE'), 'utf8').replace(/^\uFEFF/, ''),
    'the companion license must match the repository license (without its legacy BOM)',
  );
  assert.equal(packageJson.dependencies['playwright-core'], '1.57.0');
  assert.equal(packageJson.dependencies['@playwright/browser-chromium'], '1.57.0');
  assert.equal(packageJson.dependencies.fontkit, '2.0.4');
  assert.equal(packageJson.dependencies['pdf-lib'], '1.17.1');
  assert.equal(packageJson.peerDependencies.docxodus, packageJson.version,
    'the published companion must require the exact matching docxodus version');
  assert.equal(packageJson.bin.docxodus, './dist/cli.js');
  assert.equal(packageJson.bin['docxodus-export-host'], './dist/host.js');

  for (const required of [
    'LICENSE',
    'README.md',
    'dist/index.js',
    'dist/index.d.ts',
    'dist/cli.js',
    'dist/host.js',
    'dist/fonts/index.js',
  ]) {
    assert.ok(paths.includes(required), `package is missing required file: ${required}`);
  }
  const unexpected = paths.filter((path) => {
    if (path === 'LICENSE' || path === 'README.md' || path === 'package.json') return false;
    if (/^dist\/[^/]+\.(?:js|d\.ts)(?:\.map)?$/.test(path)) return false;
    if (/^dist\/fonts\/(?:discovery|index|resolver)\.(?:js|d\.ts)(?:\.map)?$/.test(path)) return false;
    return true;
  });
  assert.deepEqual(unexpected, [], `package contains undeclared files:\n${unexpected.join('\n')}`);
  const forbidden = paths.filter((path) =>
    /(?:browser-chromium|node_modules|example|fixture|spec|test)/i.test(path));
  assert.deepEqual(forbidden, [], `package contains development/browser payloads:\n${forbidden.join('\n')}`);
  console.log(`@docxodus/export package boundary: ${paths.length} runtime/license files`);
} finally {
  rmSync(cache, { recursive: true, force: true });
}

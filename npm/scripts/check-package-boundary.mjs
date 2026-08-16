import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));
const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const cache = mkdtempSync(join(tmpdir(), 'docxodus-pack-'));

try {
  // --ignore-scripts prevents this check from recursively invoking lifecycle
  // hooks. The isolated cache also makes the audit deterministic and leaves no
  // npm state in a developer's home directory.
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
  assert.equal(
    packed.status,
    0,
    `npm pack audit failed:\n${packed.stderr || packed.stdout}`,
  );

  const [manifest] = JSON.parse(packed.stdout);
  const paths = manifest.files.map(({ path }) => path);
  const packageJson = JSON.parse(readFileSync(join(packageRoot, 'package.json'), 'utf8'));
  const packageLicense = readFileSync(join(packageRoot, 'LICENSE'), 'utf8');
  const repositoryLicense = readFileSync(join(packageRoot, '..', 'LICENSE'), 'utf8')
    .replace(/^\uFEFF/, '');

  assert.equal(packageJson.license, 'MIT', 'the npm package must remain MIT');
  assert.equal(packageLicense, repositoryLicense, 'npm LICENSE must match the repository MIT license');
  assert.match(packageLicense, /Copyright \(c\) Microsoft Corporation/,
    'npm LICENSE must preserve the inherited Microsoft notice');
  assert.match(packageLicense, /Copyright \(c\) 2025-2026 John Scrudato IV/,
    'npm LICENSE must credit John Scrudato IV');
  for (const required of [
    'LICENSE',
    'README.md',
    'dist/wasm/_framework/dotnet.js',
    'dist/export-browser.bundle.js',
    'dist/export-browser.d.ts',
    'dist/export-assets.json',
    'dist/render-report-v1.schema.json',
  ]) {
    assert.ok(paths.includes(required), `npm package is missing required runtime file: ${required}`);
  }

  const entrypoints = new Set([packageJson.main, packageJson.module, packageJson.types]);
  const collectEntrypoints = (value) => {
    if (typeof value === 'string') entrypoints.add(value);
    else if (value && typeof value === 'object') Object.values(value).forEach(collectEntrypoints);
  };
  collectEntrypoints(packageJson.exports);
  collectEntrypoints(packageJson.typesVersions);
  for (const entrypoint of entrypoints) {
    const path = entrypoint.replace(/^\.\//, '');
    assert.ok(paths.includes(path), `npm package is missing declared entrypoint: ${path}`);
  }

  const allowed = paths.filter((path) => {
    if (path === 'LICENSE' || path === 'README.md' || path === 'package.json') return false;
    if (/^dist\/[^/]+\.(?:js|d\.ts)(?:\.map)?$/.test(path)) return false;
    if (/^dist\/[^/]+\.json$/.test(path)) return false;
    if (path.startsWith('dist/wasm/_framework/')) return false;
    return true;
  });
  assert.deepEqual(
    allowed,
    [],
    `npm package contains files outside the runtime allowlist:\n${allowed.join('\n')}`,
  );

  const forbidden = paths.filter((path) =>
    /(?:arcade|freedoom|ascii|demo|example|harness|test)/i.test(path));
  assert.deepEqual(
    forbidden,
    [],
    `npm package contains demo or test machinery:\n${forbidden.join('\n')}`,
  );

  console.log(`npm package boundary: ${paths.length} runtime/license files; no demo or test machinery`);
} finally {
  rmSync(cache, { recursive: true, force: true });
}

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { copyFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { build } from 'esbuild';

/** Exercise only files npm will ship, from a consumer with no editor dependencies. */
export async function checkCoreEntry(packageRoot, paths, temporaryRoot) {
  const consumer = join(temporaryRoot, 'consumer');
  const installed = join(consumer, 'node_modules', 'docxodus');
  for (const path of paths) {
    const destination = join(installed, path);
    mkdirSync(dirname(destination), { recursive: true });
    copyFileSync(join(packageRoot, path), destination);
  }
  writeFileSync(join(consumer, 'package.json'), '{"type":"module"}\n');

  // The browser barrel still needs a bundler. Leave core external so the test
  // verifies both entry points share the same live module and WASM instance.
  await build({
    entryPoints: [join(packageRoot, 'dist', 'index.js')],
    bundle: true,
    format: 'esm',
    external: ['./core.js'],
    outfile: join(installed, 'dist', 'browser-test.mjs'),
  });

  const run = (args, label) => {
    const result = spawnSync(process.execPath, args, {
      cwd: consumer,
      encoding: 'utf8',
      timeout: 60_000,
    });
    if (result.error) throw result.error;
    assert.equal(result.status, 0, `${label} failed:\n${result.stdout}\n${result.stderr}`);
  };
  copyFileSync(join(packageRoot, 'tests', 'node-core.mjs'), join(consumer, 'node-core.mjs'));
  run(['node-core.mjs'], 'plain Node ESM core consumer');

  // Check both modern exports-map resolution and the legacy typesVersions path.
  writeFileSync(join(consumer, 'consumer.mts'), `
import * as core from 'docxodus/core';
import type * as browser from 'docxodus';
const engine: Omit<typeof browser, 'DocxEditor' | 'CommentGutter' | 'mountRibbon'> = core;
const browserEngine: typeof core = {} as typeof browser;
const options: core.ConversionOptions = { renderAnnotations: true };
const session: core.DocxSession = core.openDocxSession(new Uint8Array());
const html: Promise<string> = core.convertDocxToHtml(session.save(), options);
// @ts-expect-error Editor APIs belong to the full browser entry.
core.DocxEditor;
`);
  for (const resolution of ['NodeNext', 'node']) {
    run([
      join(packageRoot, 'node_modules', 'typescript', 'bin', 'tsc'),
      '--noEmit', '--strict', '--skipLibCheck', '--target', 'ES2020',
      '--module', resolution === 'NodeNext' ? 'NodeNext' : 'ESNext',
      '--moduleResolution', resolution, 'consumer.mts',
    ], `core declarations (${resolution})`);
  }
  console.log('npm core entry: plain Node ESM, WASM operations, shared browser state, and consumer types passed');
}

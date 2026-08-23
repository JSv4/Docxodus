import { expect, test } from '@playwright/test';
import { mkdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  assertBuildOwningLifecycle,
  javascriptGraphEvidence,
} from './visual-parity/build-provenance.js';

test.describe('generated-PDF build provenance', () => {
  const work = join(tmpdir(), `docxodus-build-evidence-${process.pid}`);
  test.beforeEach(() => {
    rmSync(work, { recursive: true, force: true });
    mkdirSync(join(work, 'nested'), { recursive: true });
    writeFileSync(join(work, 'index.js'), 'export const value = 1;\n');
    writeFileSync(join(work, 'nested/module.js'), 'export const sibling = 2;\n');
    writeFileSync(join(work, 'index.js.map'), 'ignored source map');
  });
  test.afterEach(() => rmSync(work, { recursive: true, force: true }));

  test('is stable but detects tampering in an imported sibling module', () => {
    const first = javascriptGraphEvidence(work);
    expect(first).toMatchObject({ files: 2, sha256: expect.stringMatching(/^[0-9a-f]{64}$/) });
    expect(javascriptGraphEvidence(work)).toEqual(first);
    writeFileSync(join(work, 'nested/module.js'), 'export const sibling = 3;\n');
    expect(javascriptGraphEvidence(work).sha256).not.toBe(first.sha256);
  });

  test('rejects symlinks in the emitted module graph', () => {
    if (process.platform === 'win32') return;
    symlinkSync(join(work, 'index.js'), join(work, 'linked.js'));
    expect(() => javascriptGraphEvidence(work)).toThrow(/symlink/);
  });

  test('rejects an active direct Playwright invocation that skipped the owned builds', () => {
    expect(() => assertBuildOwningLifecycle(false, undefined)).not.toThrow();
    expect(() => assertBuildOwningLifecycle(true, 'test:generated-pdf-parity')).not.toThrow();
    // CI builds both packages as its own steps and then runs the prebuilt runner, so this name
    // must be accepted or the workflow pays for a second trimmed WASM publish.
    expect(() => assertBuildOwningLifecycle(true, 'test:generated-pdf-parity:prebuilt'))
      .not.toThrow();
    expect(() => assertBuildOwningLifecycle(true, undefined)).toThrow(/direct Playwright/);
    expect(() => assertBuildOwningLifecycle(true, 'test')).toThrow(/direct Playwright/);
  });
});

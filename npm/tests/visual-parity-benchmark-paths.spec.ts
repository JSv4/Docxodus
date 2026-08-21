import { expect, test } from '@playwright/test';
import { mkdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  assertSafeCaseId,
  prepareExternalOutputRoot,
  resolveTrackedRegularFile,
} from './visual-parity/benchmark-paths.js';

test.describe('generated-PDF benchmark path confinement', () => {
  const work = join(tmpdir(), `docxodus-benchmark-paths-${process.pid}`);
  const repository = join(work, 'repository');
  const outside = join(work, 'outside');

  test.beforeEach(() => {
    rmSync(work, { recursive: true, force: true });
    mkdirSync(repository, { recursive: true });
    mkdirSync(outside, { recursive: true });
  });
  test.afterEach(() => rmSync(work, { recursive: true, force: true }));

  test('rejects traversal identifiers and repository-relative path escapes', () => {
    expect(() => assertSafeCaseId('pdf-legal-contract')).not.toThrow();
    for (const id of ['../escape', 'nested/case', '', '-leading', 'trailing-']) {
      expect(() => assertSafeCaseId(id)).toThrow(/unsafe/);
    }
    expect(() => resolveTrackedRegularFile(repository, '../outside/file.docx')).toThrow();
    expect(() => resolveTrackedRegularFile(repository, 'nested\\file.docx')).toThrow();
  });

  test('accepts a regular tracked file and rejects a symlinked artifact root', () => {
    writeFileSync(join(repository, 'fixture.docx'), 'fixture');
    expect(resolveTrackedRegularFile(repository, 'fixture.docx')).toBe(join(repository, 'fixture.docx'));
    if (process.platform === 'win32') return;
    const linked = join(outside, 'linked-root');
    symlinkSync(repository, linked, 'dir');
    expect(() => prepareExternalOutputRoot(repository, linked, 0, new Set()))
      .toThrow(/non-symlink|inside the repository/);
  });

  test('separates retries and rejects symlink bootstrap artifacts', () => {
    const root = join(outside, 'artifacts');
    const first = prepareExternalOutputRoot(repository, root, 0, new Set());
    expect(first).toBe(root);
    expect(prepareExternalOutputRoot(repository, root, 1, new Set())).toBe(join(root, 'retry-1'));
    if (process.platform === 'win32') return;
    const bootstrap = join(root, 'ci-context.json');
    symlinkSync(join(repository, 'missing'), bootstrap);
    expect(() => prepareExternalOutputRoot(repository, root, 0, new Set(['ci-context.json'])))
      .toThrow(/unsafe/);
  });
});

import { expect, test } from '@playwright/test';
import { realpathSync } from 'node:fs';
import { basename, dirname } from 'node:path';
import {
  commandVersion,
  pinExecutable,
  resolveExecutable,
} from './visual-parity/toolchain.js';

test.describe('visual-parity executable evidence', () => {
  test('resolves, invokes, and hashes the same absolute executable', () => {
    const resolved = resolveExecutable(basename(process.execPath), dirname(process.execPath));
    expect(resolved).toBe(realpathSync(process.execPath));
    expect(commandVersion(resolved, ['--version'])).toMatch(/^v\d+/);
    const pinned = pinExecutable(
      basename(process.execPath),
      ['--version'],
      dirname(process.execPath),
    );
    expect(pinned.path).toBe(realpathSync(process.execPath));
    expect(pinned.evidence).toMatchObject({
      command: basename(process.execPath),
      executable: basename(process.execPath),
      executableSha256: expect.stringMatching(/^[0-9a-f]{64}$/),
      version: expect.stringMatching(/^v\d+/),
    });
  });

  test('rejects unresolved and relative version probes', () => {
    expect(() => resolveExecutable('definitely-not-a-docxodus-tool', '')).toThrow(/not found/);
    expect(() => commandVersion('node', ['--version'])).toThrow(/absolute/);
  });
});

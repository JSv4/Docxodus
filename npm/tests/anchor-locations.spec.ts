import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import type { DocxSession } from '../src/session.js';
import type { WorkerDocxSession } from '../src/worker-proxy.js';

test('typed and raw bridge queries preserve all/subset/empty semantics without an edit', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const result = await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const session = D.openTypedSession(D.DocxSessionBridge.CreateBlankDocx(), JSON.stringify({ captureInitialProjection: false })) as DocxSession;
    try {
      const all = session.getAnchorLocations();
      const key = Object.keys(all.locations).find(k => k.startsWith('p:'))!;
      const unknown = 'not " an anchor';
      const selected = session.getAnchorLocations([key, unknown, key]);
      return { all, selected, key, unknown, none: session.getAnchorLocations([]),
        explicitNull: session.getAnchorLocations(null), undo: session.undo() };
    } finally { session.close(); }
  });
  expect(result.all.version).toBe(0);
  expect(result.all.locations[result.key]).toEqual({ anchorId: result.key, partUri: '/word/document.xml', elementPath: '/*[1]/*[1]/*[1]' });
  expect(result.selected.locations).toEqual({ [result.key]: result.all.locations[result.key], [result.unknown]: null });
  expect(result.none.locations).toEqual({});
  expect(result.explicitNull).toEqual(result.all);
  expect(result.undo).toBe(false);
});

test('malformed raw arrays are refused and typed queries surface their errors', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  const result = await page.evaluate(() => {
    const D = (window as any).Docxodus, bridge = D.DocxSessionBridge;
    const session = D.openTypedSession(bridge.CreateBlankDocx()) as DocxSession;
    const rawHandle = bridge.OpenSession(bridge.CreateBlankDocx(), '');
    try {
      const invalid = ['{', '42', '[null]', '[true]'].map(json => JSON.parse(bridge.GetAnchorLocations(rawHandle, json)));
      let error = '';
      try { session.getAnchorLocations([null] as any); } catch (e) { error = String(e); }
      return { invalid, error, valid: session.getAnchorLocations([]) };
    } finally { session.close(); bridge.CloseSession(rawHandle); }
  });
  expect(result.invalid.every(value => typeof value.error === 'string')).toBe(true);
  expect(result.error).toContain('malformed anchor id array');
  expect(result.valid.locations).toEqual({});
});

test('worker session exposes the same read-only location envelope', async ({ page }) => {
  const bytes = Array.from(readFileSync(new URL('../../TestFiles/HC006-Test-01.docx', import.meta.url)));
  await page.goto('/worker-test-harness.html');
  const result = await page.evaluate(async bytes => {
    const worker = await (window as any).createDocxodusWorker();
    const session = await worker.openDocxSession(new Uint8Array(bytes), { captureInitialProjection: false }) as WorkerDocxSession;
    try {
      const all = await session.getAnchorLocations();
      const key = Object.keys(all.locations)[0];
      return { all, key, selected: await session.getAnchorLocations([key, 'unknown']), none: await session.getAnchorLocations([]) };
    } finally { await session.close(); worker.terminate(); }
  }, bytes);
  expect(result.all.version).toBe(0);
  expect(result.all.locations[result.key]?.elementPath).toMatch(/^\/\*\[1\]/);
  expect(result.selected.locations).toEqual({ [result.key]: result.all.locations[result.key], unknown: null });
  expect(result.none.locations).toEqual({});
});

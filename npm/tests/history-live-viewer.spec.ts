import { test, expect } from '@playwright/test';
import { build } from 'esbuild';

let bundle: string;
test.beforeAll(async () => {
  bundle = (await build({ entryPoints: ['dist/embed.js'], bundle: true, format: 'esm', write: false })).outputFiles[0].text;
});
test.beforeEach(async ({ page }) => {
  await page.route('**/history-viewer.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><div id="a"></div><div id="b"></div>' }));
  await page.route('**/history-viewer-api.js', route => route.fulfill({ contentType: 'text/javascript', body: bundle }));
  await page.goto('http://localhost:8083/history-viewer.html');
  await page.evaluate(async () => {
    const moduleUrl = '/history-viewer-api.js';
    const api = await import(moduleUrl);
    await api.initialize('http://localhost:8083/wasm/');
    (window as any).historyApi = api;
  });
});

test('two viewers follow accepted logs, pause at time, resume, and install restore boundaries', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const storage = api.createMemoryHistoryStorage();
    const a = api.openDocxHistory(storage); const b = api.openDocxHistory(storage);
    const make = (text: string) => {
      const session = api.openDocxSession(api.createBlankDocx());
      const anchor = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'));
      session.replaceText(anchor, text); const bytes = session.save(); session.close(); return bytes;
    };
    const metadata = (hour: number) => ({ author: 'host', createdAt: `2026-01-01T${hour}:00:00Z` });
    const original = make('Initial document');
    const first = await a.createVersion('doc', null, original, metadata(12));
    const options = { wasmBasePath: 'http://localhost:8083/wasm/' };
    const left = await api.createHistoryViewer('#a', a, 'doc', options);
    const right = await api.createHistoryViewer('#b', b, 'doc', options);
    const second = await a.createVersion('doc', first.head, make('Second document'), metadata(13));
    await Promise.all([left.refresh(), right.refresh()]);
    const equal = left.html === right.html;
    await left.showTime('2026-01-01T12:30:00Z');
    const third = await b.createVersion('doc', second.head, make('Third document'), metadata(14));
    await Promise.all([left.refresh(), right.refresh()]);
    const paused = !left.following && left.sequence === '0' && left.html.includes('Initial document');
    const caughtUpWhilePaused = left.head.revision === third.head.revision;
    await left.resume();
    const resumed = left.following && left.sequence === '2' && left.html.includes('Third document');
    const restored = await a.restoreVersion('doc', third.head, first.version.id, metadata(15));
    const [lr, rr] = await Promise.all([left.refresh(), right.refresh()]);
    const exact = Array.from(left.exportDisplayed()).join(',') === Array.from(original).join(',');
    const duplicate = await left.refresh();
    const copiedHead = left.head; copiedHead.revision = '999';
    const final = { equal, paused, caughtUpWhilePaused, resumed, reset: lr.reset && rr.reset,
      exact, finalSequence: left.sequence, duplicate: duplicate.entries.length, headOwned: left.head.revision === restored.head.revision,
      scoped: document.querySelectorAll('[data-docxodus-embed-root]').length === 2,
      text: document.querySelector('#a')!.textContent };
    left.destroy(); right.destroy(); a.close(); b.close();
    return { ...final, empty: document.querySelector('#a')!.childNodes.length === 0 };
  });
  expect(result.equal).toBe(true);
  expect(result.paused).toBe(true);
  expect(result.caughtUpWhilePaused).toBe(true);
  expect(result.resumed).toBe(true);
  expect(result.reset).toBe(true);
  expect(result.exact).toBe(true);
  expect(result.finalSequence).toBe('3');
  expect(result.duplicate).toBe(0);
  expect(result.headOwned).toBe(true);
  expect(result.scoped).toBe(true);
  expect(result.text).toContain('Initial document');
  expect(result.empty).toBe(true);
});

test('corrupt snapshots retain the prior frame and head; a later retry recovers', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const backing = api.createMemoryHistoryStorage();
    let damaged: string | null = null;
    const storage = { ...backing, async readBlob(ref: any) {
      const bytes = await backing.readBlob(ref);
      if (bytes && ref.digest.value === damaged) bytes[0] ^= 1;
      return bytes;
    } };
    const history = api.openDocxHistory(storage);
    const bytes = api.createBlankDocx();
    const metadata = { author: 'host', createdAt: '2026-01-01T12:00:00Z' };
    const first = await history.createVersion('doc', null, bytes, metadata);
    const viewer = await api.createHistoryViewer('#a', history, 'doc', { wasmBasePath: 'http://localhost:8083/wasm/' });
    const priorHtml = viewer.html;
    const session = api.openDocxSession(bytes);
    session.replaceText(Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:')), 'Recovered document');
    const nextBytes = session.save(); session.close();
    const second = await history.createVersion('doc', first.head, nextBytes, metadata);
    damaged = second.state.snapshot.blob.digest.value;
    let error = '';
    try { await viewer.refresh(); } catch (failure: any) { error = failure.code; }
    const retained = viewer.head.revision === first.head.revision && viewer.html === priorHtml && viewer.sequence === '0';
    damaged = null;
    await viewer.refresh();
    const recovered = viewer.head.revision === second.head.revision && viewer.html.includes('Recovered document');
    viewer.destroy(); history.close();
    return { error, retained, recovered };
  });
  expect(result).toEqual({ error: 'PayloadMismatch', retained: true, recovered: true });
});

test('queued refreshes do not duplicate commits or repaint after destruction', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const backing = api.createMemoryHistoryStorage();
    let block = false;
    let release!: () => void;
    let entered!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    const started = new Promise<void>(resolve => { entered = resolve; });
    const storage = { ...backing, async readHead(id: string) {
      if (block) { entered(); await gate; } return backing.readHead(id);
    } };
    const history = api.openDocxHistory(storage);
    const metadata = { author: 'host', createdAt: '2026-01-01T12:00:00Z' };
    const original = api.createBlankDocx();
    const first = await history.createVersion('doc', null, original, metadata);
    const viewer = await api.createHistoryViewer('#a', history, 'doc', { wasmBasePath: 'http://localhost:8083/wasm/' });
    const session = api.openDocxSession(original);
    session.replaceText(Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:')), 'Intermediate');
    const modified = session.save(); session.close();
    const second = await history.createVersion('doc', first.head, modified, metadata);
    // Returning to identical content via imports still advances displayed sequence, without reset.
    await history.createVersion('doc', second.head, original, metadata);
    const [one, two] = await Promise.all([viewer.refresh(), viewer.refresh()]);
    const sequence = viewer.sequence;
    block = true;
    const pending = viewer.refresh().then(() => 'unexpected', (error: any) => error.code);
    await started; viewer.destroy(); release();
    const closed = await pending;
    history.close();
    return { first: one.entries.length, second: two.entries.length, sequence, closed, empty: document.querySelector('#a')!.childNodes.length === 0 };
  });
  expect(result).toEqual({ first: 2, second: 0, sequence: '2', closed: 'Closed', empty: true });
});

test('renderer failure preserves the prior head and the queue remains retryable', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const history = api.openDocxHistory(api.createMemoryHistoryStorage());
    const metadata = { author: 'host', createdAt: '2026-01-01T12:00:00Z' };
    const original = api.createBlankDocx();
    const first = await history.createVersion('doc', null, original, metadata);
    let failRendering = false;
    let exports = 0;
    // Fault-inject at the renderer input, after the real core update validation boundary.
    const input = {
      readChangesSince: (...args: any[]) => history.readChangesSince(...args),
      exportVersion: async (...args: any[]) => {
        exports++;
        return failRendering ? new Uint8Array([1, 2, 3]) : history.exportVersion(...args);
      },
    };
    const viewer = await api.createHistoryViewer('#a', input, 'doc', { wasmBasePath: 'http://localhost:8083/wasm/' });
    const initialHtml = viewer.html;
    const paragraph = document.querySelector('#a p');
    await viewer.refresh();
    const duplicateSkipped = exports === 1 && paragraph === document.querySelector('#a p');
    const session = api.openDocxSession(original);
    session.replaceText(Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:')), 'After retry');
    const bytes = session.save(); session.close();
    const next = await history.createVersion('doc', first.head, bytes, metadata);
    failRendering = true;
    let failed = false;
    try { await viewer.refresh(); } catch { failed = true; }
    const retained = viewer.head.revision === first.head.revision && viewer.html === initialHtml
      && paragraph === document.querySelector('#a p');
    failRendering = false;
    await viewer.refresh();
    const recovered = viewer.head.revision === next.head.revision && viewer.html.includes('After retry');
    viewer.destroy(); history.close();
    return { duplicateSkipped, failed, retained, recovered };
  });
  expect(result).toEqual({ duplicateSkipped: true, failed: true, retained: true, recovered: true });
});

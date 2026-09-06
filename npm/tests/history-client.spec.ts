import { test, expect } from '@playwright/test';
import { build } from 'esbuild';

let bundle: string;
test.beforeAll(async () => {
  const result = await build({ entryPoints: ['dist/index.js'], bundle: true, format: 'esm', write: false });
  bundle = result.outputFiles[0].text;
});

test.beforeEach(async ({ page }) => {
  await page.route('**/history-test.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><title>History test</title>' }));
  await page.route('**/history-api.js', route => route.fulfill({ contentType: 'text/javascript', body: bundle }));
  await page.goto('http://localhost:8083/history-test.html');
  await page.evaluate(async () => {
    const moduleUrl = '/history-api.js';
    const api = await import(moduleUrl);
    await api.initialize('http://localhost:8083/wasm/');
    (window as any).historyApi = api;
  });
});

test('two host-backed clients reopen, replay, render at time, and restore exact versions', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const storage = api.createMemoryHistoryStorage();
    const a = api.openDocxHistory(storage);
    const b = api.openDocxHistory(storage);
    const original = api.createBlankDocx();
    const metadata = { author: 'alice', createdAt: '2026-01-01T12:00:00Z' };
    const first = await a.createVersion('doc', null, original, metadata);
    const session = api.openDocxSession(original);
    const anchor = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'));
    const edit = session.replaceText(anchor, 'shared historical text');
    if (!edit.success) throw new Error(JSON.stringify(edit));
    const edited = session.save(); session.close();
    const second = await b.createVersion('doc', first.head, edited, { ...metadata, createdAt: '2026-01-01T13:00:00Z' });
    const sequence = await a.resolveSequenceAtTime('doc', '2026-01-01T13:00:00Z');
    const replay = await a.replay('doc', sequence);
    const rendered = api.openDocxSession(replay);
    const markdown = rendered.project().markdown; rendered.close();
    const bytes = await a.exportVersion('doc', first.version.id);
    const page = await a.listVersions('doc', null, 1);
    const rest = await a.listVersions('doc', page.next);
    let stale = '';
    try { await a.createVersion('doc', first.head, original, metadata); }
    catch (error: any) { stale = error.code; }
    const restored = await b.restoreVersion('doc', second.head, first.version.id, metadata);
    const updates = await a.readChangesSince('doc', first.head);
    const duplicate = await b.readChangesSince('doc', restored.head);
    a.close(); b.close();
    const reopened = api.openDocxHistory(storage);
    const latest = await reopened.read('doc');
    const materialized = await reopened.materialize('doc', restored.state.sequence);
    reopened.close();
    let closed = '';
    try { await reopened.read('doc'); } catch (error: any) { closed = error.code; }
    return {
      sequence, markdown, stale, closed, epoch: latest.state.epoch, revisionType: typeof latest.head.revision,
      exact: bytes.every((n: number, i: number) => n === original[i]) && bytes.length === original.length,
      restored: materialized.every((n: number, i: number) => n === original[i]) && materialized.length === original.length,
      firstId: first.version.id.digest.value, listedId: rest.versions[0].id.digest.value,
      updates: updates.entries.map((entry: any) => entry.commit.sequence), reset: updates.reset, duplicates: duplicate.entries.length,
    };
  });
  expect(result.sequence).toBe('1');
  expect(result.markdown).toContain('shared historical text');
  expect(result.stale).toBe('StaleHead');
  expect(result.closed).toBe('Closed');
  expect(result.epoch).toBe('1');
  expect(result.revisionType).toBe('string');
  expect(result.exact).toBe(true);
  expect(result.restored).toBe(true);
  expect(result.listedId).toBe(result.firstId);
  expect(result.updates).toEqual(['1', '2']);
  expect(result.reset).toBe(true);
  expect(result.duplicates).toBe(0);
});

test('async storage retains captured inputs, rejects close while active, and detects corrupted payloads', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const backing = api.createMemoryHistoryStorage();
    let release!: () => void;
    const gate = new Promise<void>(resolve => { release = resolve; });
    let paused = true;
    let corrupt = false;
    const storage = {
      ...backing,
      async readHead(id: string) { if (paused) { paused = false; await gate; } return backing.readHead(id); },
      async readBlob(ref: any) {
        const bytes = await backing.readBlob(ref);
        if (corrupt && bytes) bytes[0] ^= 1;
        return bytes;
      },
    };
    const client = api.openDocxHistory(storage);
    const original = api.createBlankDocx();
    const bytes = original.slice();
    const metadata = { author: 'original', createdAt: '2026-01-01T00:00:00Z' };
    const pending = client.createVersion('doc', null, bytes, metadata);
    bytes.fill(0); metadata.author = 'mutated';
    let activeClose = false;
    try { client.close(); } catch { activeClose = true; }
    release();
    const created = await pending;
    const exported = await client.exportVersion('doc', created.version.id);
    corrupt = true;
    let corruption = '';
    try { await client.read('doc'); } catch (error: any) { corruption = error.code; }
    client.close();
    return { activeClose, author: created.version.record.metadata.author, corruption,
      exact: original.length === exported.length && original.every((n: number, i: number) => n === exported[i]) };
  });
  expect(result).toEqual({ activeClose: true, author: 'original', corruption: 'PayloadMismatch', exact: true });
});

test('memory adapter makes one concurrent CAS winner and owns its blobs and heads', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const storage = api.createMemoryHistoryStorage();
    const bytes = new Uint8Array([1, 2, 3]);
    const digest = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), n => n.toString(16).padStart(2, '0')).join('');
    const reference = { digest: { algorithm: 'SHA-256', value: digest }, length: 3 };
    const put = storage.putBlob(reference, bytes); bytes.fill(9); await put;
    const firstRead = await storage.readBlob(reference); firstRead.fill(7);
    const retained = Array.from(await storage.readBlob(reference));
    const winners = await Promise.all(Array.from({ length: 8 }, () => storage.advanceHead('doc', null, reference)));
    const winner = winners.find(Boolean); winner.revision = '99'; winner.state.digest.value = 'bad';
    const actual = await storage.readHead('doc');
    let mismatch = '';
    try { await storage.putBlob(reference, new Uint8Array([3, 2, 1])); } catch (error: any) { mismatch = error.code; }
    return { retained, count: winners.filter(Boolean).length, revision: actual.revision, digest: actual.state.digest.value, expectedDigest: digest, mismatch };
  });
  expect(result.retained).toEqual([1, 2, 3]);
  expect(result.count).toBe(1);
  expect(result.revision).toBe('1');
  expect(result.digest).toBe(result.expectedDigest);
  expect(result.mismatch).toBe('PayloadMismatch');
});

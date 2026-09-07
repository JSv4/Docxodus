import { test, expect } from '@playwright/test';
import { build } from 'esbuild';
import { readFile, writeFile } from 'node:fs/promises';
import { deflateRawSync } from 'node:zlib';

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

test('durable request IDs recover lost acknowledgements and return original results after later writes', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = (window as any).historyApi;
    const backing = api.createMemoryHistoryStorage();
    let loseAck = true;
    const storage = {
      ...backing,
      async advanceHead(id: string, expected: any, state: any) {
        const head = await backing.advanceHead(id, expected, state);
        if (head && loseAck) { loseAck = false; throw new Error('Lost durable acknowledgement'); }
        return head;
      },
    };
    const a = api.openDocxHistory(storage);
    const bytes = api.createBlankDocx();
    const metadata = { author: 'actor', createdAt: '2026-01-01T12:00:00Z' };
    let lost = false;
    try { await a.createVersion('doc', null, bytes, metadata, 'create-id'); } catch { lost = true; }
    const first = await a.read('doc');
    const recovered = await a.createVersion('doc', null, bytes, metadata, 'create-id');
    const restored = await a.restoreVersion('doc', first.head, first.version.id, metadata, 'restore-id');
    const later = await a.createVersion('doc', restored.head, bytes, metadata);
    a.close();
    const b = api.openDocxHistory(storage);
    const oldCreate = await b.createVersion('doc', null, bytes, metadata, 'create-id');
    const oldRestore = await b.restoreVersion('doc', first.head, first.version.id, metadata, 'restore-id');
    let conflict = '';
    try { await b.createVersion('doc', later.head, bytes, metadata, 'create-id'); }
    catch (error: any) { conflict = error.code; }
    const current = await b.read('doc');
    const versions = await b.listVersions('doc'); b.close();
    return { lost, recovered: recovered.head, oldCreate: oldCreate.head, first: first.head,
      oldRestore: oldRestore.head, restored: restored.head, current: current.head, later: later.head,
      conflict, count: versions.versions.length, requestId: first.state.requests.current.id,
      revisionType: typeof first.state.requests.revision };
  });
  expect(result.lost).toBe(true);
  expect(result.recovered).toEqual(result.first); expect(result.oldCreate).toEqual(result.first);
  expect(result.oldRestore).toEqual(result.restored); expect(result.current).toEqual(result.later);
  expect(result.conflict).toBe('RequestConflict'); expect(result.count).toBe(3);
  expect(result.requestId).toBe('create-id'); expect(result.revisionType).toBe('string');
});

test('real history file opens readonly, drives controls, imports and continues without its original store', async ({ page }, testInfo) => {
  test.setTimeout(120_000);
  const input = await readFile('../TestFiles/HistoryArchive/agreement.docxhistory');
  const expected = await readFile('../TestFiles/HistoryArchive/agreement-v1.docx');
  const result = await page.evaluate(async ({ encoded, expected }) => {
    const api = (window as any).historyApi;
    const decode = (s: string) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
    const encode = (b: Uint8Array) => { let s = ''; for (const value of b) s += String.fromCharCode(value); return btoa(s); };
    const bytes = decode(encoded);
    const pending = api.openDocxHistoryArchive(bytes); bytes.fill(0);
    const reader = await pending;
    const page1 = await reader.listVersions(null, 1);
    const rest = await reader.listVersions(page1.next, 25);
    const first = rest.versions[rest.versions.length - 1];
    const before = await reader.exportDocx(first.id);
    const revision = rest.versions.find((v: any) => v.record.metadata.label === 'Counsel revision');
    const redline = await reader.compareVersions(first.id, revision.id); // First comparison: exercises WASM cold-path warmup.
    let readOnly = '';
    try { await (reader as any).call('create'); } catch (error: any) { readOnly = error.code; }
    const readonlyShape = typeof reader.createVersion === 'undefined' && typeof reader.restoreVersion === 'undefined';
    const originalHead = reader.info.head;
    reader.close(); let closed = '';
    try { await reader.read(); } catch (error: any) { closed = error.code; }
    const backing = api.createMemoryHistoryStorage(); let loseAck = true;
    const store = { ...backing, async initializeHead(id: string, head: any) {
      const value = await backing.initializeHead(id, head);
      if (loseAck) { loseAck = false; throw new Error('Lost initialization acknowledgement'); }
      return value;
    } };
    const history = api.openDocxHistory(store); let lost = false;
    try { await history.importHistoryArchive(decode(encoded)); } catch { lost = true; }
    const imported = await history.importHistoryArchive(decode(encoded));
    const doc = history.document(imported.archive.documentId);
    const retry = await doc.createVersion(null, before, first.record.metadata, 'initial');
    const saved = await doc.createVersion(imported.view.head, before,
      { author: 'frontend', createdAt: '2026-09-01T12:00:00Z', label: 'Browser checkpoint' }, 'browser-checkpoint');
    const restored = await doc.restoreVersion(saved.head, revision.id,
      { author: 'frontend', createdAt: '2026-09-01T13:00:00Z' }, 'browser-restore');
    let conflict = '';
    try { await history.importHistoryArchive(decode(encoded)); } catch (error: any) { conflict = error.code; }
    const portable = await doc.exportHistoryArchive(); history.close();
    const standalone = await api.openDocxHistoryArchive(portable);
    const current = await standalone.read(); const latest = await standalone.exportDocx(); standalone.close();
    const parsed = api.openDocxSession(latest); parsed.close();
    return { readOnly, readonlyShape, closed, originalHead, importedHead: imported.view.head,
      alreadyPresent: imported.alreadyPresent, lost, retryRevision: retry.head.revision, savedRevision: saved.head.revision,
      restoredHead: restored.head, currentHead: current.head, conflict, count: rest.versions.length + 1,
      exact: encode(before) === expected, redline: encode(redline), archive: encode(portable), latest: encode(latest) };
  }, { encoded: input.toString('base64'), expected: expected.toString('base64') });
  expect(result.readOnly).toBe('ReadOnly'); expect(result.readonlyShape).toBe(true); expect(result.closed).toBe('Closed');
  expect(result.importedHead).toEqual(result.originalHead); expect(result.alreadyPresent).toBe(true); expect(result.lost).toBe(true);
  expect(result.retryRevision).toBe('1'); expect(result.savedRevision).toBe('5'); expect(result.currentHead).toEqual(result.restoredHead);
  expect(result.conflict).toBe('ImportConflict'); expect(result.exact).toBe(true); expect(result.count).toBe(4);
  await writeFile(testInfo.outputPath('browser-continued.docxhistory'), Buffer.from(result.archive, 'base64'));
  await writeFile(testInfo.outputPath('browser-latest.docx'), Buffer.from(result.latest, 'base64'));
  await writeFile(testInfo.outputPath('browser-arbitrary-comparison.docx'), Buffer.from(result.redline, 'base64'));
});

test('old storage adapters still work but cannot import; malformed archives fail explicitly', async ({ page }) => {
  const encoded = (await readFile('../TestFiles/HistoryArchive/agreement.docxhistory')).toString('base64');
  const result = await page.evaluate(async encoded => {
    const api = (window as any).historyApi;
    const bytes = Uint8Array.from(atob(encoded), c => c.charCodeAt(0));
    const backing = api.createMemoryHistoryStorage(); let puts = 0;
    const history = api.openDocxHistory({ ...backing, initializeHead: undefined,
      async putBlob(ref: any, bytes: Uint8Array) { puts++; await backing.putBlob(ref, bytes); } });
    let unsupported = ''; try { await history.importHistoryArchive(bytes); } catch (error: any) { unsupported = error.code; }
    const writesAtFailure = puts;
    const legacy = await history.createVersion('legacy', null, api.createBlankDocx(), { author: 'legacy', createdAt: '2026-01-01T00:00:00Z' });
    history.close(); let malformed = '';
    try { await api.openDocxHistoryArchive(bytes.subarray(0, bytes.length - 1)); } catch (error: any) { malformed = error.code; }
    return { unsupported, writesAtFailure, legacyRevision: legacy.head.revision, malformed };
  }, encoded);
  expect(result).toEqual({ unsupported: 'InitializationUnsupported', writesAtFailure: 0, legacyRevision: '1', malformed: 'InvalidManifest' });
});

test('real charter archive exposes conflict decisions and the exact retained proposal', async ({ page }) => {
  test.setTimeout(120_000);
  const encoded = (await readFile('../TestFiles/HistoryArchive/charter-collaboration.docxhistory')).toString('base64');
  const proposal = (await readFile('../TestFiles/HistoryArchive/charter-collaboration-conflicting-proposal.docx')).toString('base64');
  const result = await page.evaluate(async ({ encoded, proposal }) => {
    const api = (window as any).historyApi;
    const decode = (s: string) => Uint8Array.from(atob(s), c => c.charCodeAt(0));
    const reader = await api.openDocxHistoryArchive(decode(encoded));
    const update = await reader.readOperationsSince(null);
    const conflict = update.operations.find((op: any) => op.record.status === 'conflict');
    const fetched = await reader.getOperation(conflict.id);
    const exported = await reader.exportOperationProposal(conflict.id); const expected = decode(proposal);
    const duplicates = await reader.readOperationsSince(update.view.head);
    reader.close();
    return { count: update.operations.length, status: fetched.record.status, revisionType: typeof fetched.record.revision,
      exact: exported.length === expected.length && exported.every((b: number, i: number) => b === expected[i]),
      resolved: update.operations.some((op: any) => op.record.status === 'accepted'
        && op.input.request.resolves?.digest.value === conflict.id.digest.value), duplicates: duplicates.operations.length };
  }, { encoded, proposal });
  expect(result).toEqual({ count: 3, status: 'conflict', revisionType: 'string', exact: true, resolved: true, duplicates: 0 });
});

test('small compressed upload with an oversized entry fails before browser inflation', async ({ page }) => {
  // A valid one-entry ZIP whose 65 MiB payload compresses below 100 KiB. It intentionally
  // has no history manifest: entry-size rejection must precede that later format check.
  const data = Buffer.alloc(65 * 1024 * 1024);
  const packed = deflateRawSync(data, { level: 9 }); const name = Buffer.from('blobs/' + 'a'.repeat(64));
  const table = Uint32Array.from({ length: 256 }, (_, i) => {
    for (let bit = 0; bit < 8; bit++) i = (i & 1) ? (0xedb88320 ^ (i >>> 1)) : (i >>> 1);
    return i >>> 0;
  });
  let crc = 0xffffffff; for (const byte of data) crc = table[(crc ^ byte) & 255] ^ (crc >>> 8); crc = (crc ^ 0xffffffff) >>> 0;
  const local = Buffer.alloc(30); local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4); local.writeUInt16LE(8, 8);
  local.writeUInt16LE(33, 12); local.writeUInt32LE(crc, 14); local.writeUInt32LE(packed.length, 18);
  local.writeUInt32LE(data.length, 22); local.writeUInt16LE(name.length, 26);
  const central = Buffer.alloc(46); central.writeUInt32LE(0x02014b50); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
  central.writeUInt16LE(8, 10); central.writeUInt16LE(33, 14); central.writeUInt32LE(crc, 16);
  central.writeUInt32LE(packed.length, 20); central.writeUInt32LE(data.length, 24); central.writeUInt16LE(name.length, 28);
  const footer = Buffer.alloc(22); footer.writeUInt32LE(0x06054b50); footer.writeUInt16LE(1, 8); footer.writeUInt16LE(1, 10);
  footer.writeUInt32LE(central.length + name.length, 12); footer.writeUInt32LE(local.length + name.length + packed.length, 16);
  const archive = Buffer.concat([local, name, packed, central, name, footer]); expect(archive.length).toBeLessThan(100_000);
  const result = await page.evaluate(async encoded => {
    try { await (window as any).historyApi.openDocxHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0))); }
    catch (error: any) { return { code: error.code, message: error.message }; }
    return null;
  }, archive.toString('base64'));
  expect(result?.code).toBe('ResourceLimit'); expect(result?.message).toContain('Archive entry');
});

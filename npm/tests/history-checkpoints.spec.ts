import { readFile } from 'node:fs/promises';
import { test, expect } from './history-controls-harness.js';
import type { HistoryCheckpointRequest } from '../src/history-checkpoints.js';

test('a lost save acknowledgement survives reload and recovers the original agreement, even after a later save', async ({ page }) => {
  const agreement = (await readFile('../TestFiles/HistoryArchive/agreement-v1.docx')).toString('base64');
  const first = await page.evaluate(async encoded => {
    const api = window.historyApi;
    const store = await api.openIndexedDbHistoryStore('checkpoint-reload');
    const journal = store.journal('agreement');
    const history = api.openDocxHistory({ ...store.storage, async advanceHead(...args) {
      const pending = await journal.read();
      if (!pending) throw new Error('A request must be durable before publication.');
      const head = await store.storage.advanceHead(...args);
      if (head) throw new Error('Connection dropped after commit');
      return head;
    } });
    const doc = history.document('agreement');
    const commands = await api.HistoryCheckpoints.open(doc, journal);
    const empty = await doc.listVersions();
    const bytes = Uint8Array.from(atob(encoded), c => c.charCodeAt(0));
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z', label: 'Signed agreement' };
    const pending = commands.save(bytes, metadata).then(() => '', (error: Error) => error.message);
    bytes.fill(0); metadata.label = 'Mutated';
    const failure = await pending;
    const captured = await journal.read();
    const first = await doc.read();
    history.close();
    const other = api.openDocxHistory(store.storage);
    await other.document('agreement').createVersion(first!.head, api.createBlankDocx(),
      { author: 'Other editor', createdAt: '2026-09-01T13:00:00Z' }, 'later');
    other.close(); store.close();
    return { empty: empty.versions.length, failure, requestId: captured?.id, head: first!.head, pending: commands.hasPending };
  }, agreement);
  expect(first.empty).toBe(0);
  expect(first.failure).toBeTruthy();
  expect(first.pending).toBe(true);
  await page.reload();
  await page.evaluate(async () => {
    const url = '/history-controls-api.js'; window.historyApi = await import(url);
    await window.historyApi.initialize('http://localhost:8083/wasm/');
  });
  const recovered = await page.evaluate(async () => {
    const api = window.historyApi;
    const store = await api.openIndexedDbHistoryStore('checkpoint-reload');
    const history = api.openDocxHistory(store.storage);
    const doc = history.document('agreement');
    const commands = await api.HistoryCheckpoints.open(doc, store.journal('agreement'));
    const result = await commands.retry();
    const bytes = await doc.exportDocx(result.version.id);
    const output = { head: result.head, label: result.version.record.metadata.label,
      current: commands.view!.head, latest: (await doc.read())!.head, requestId: result.state.requests?.current?.id,
      remaining: await store.journal('agreement').read(), count: (await doc.listVersions()).versions.length,
      bytes: Array.from(bytes) };
    history.close(); store.close(); return output;
  });
  expect(recovered.head).toEqual(first.head);
  expect(recovered.current).toEqual(first.head);
  expect(recovered.latest.revision).toBe('2');
  expect(recovered.label).toBe('Signed agreement');
  expect(recovered.requestId).toBe(first.requestId);
  expect(recovered.remaining).toBeNull();
  expect(recovered.count).toBe(2);
  expect(Buffer.from(recovered.bytes)).toEqual(Buffer.from(agreement, 'base64'));
});

test('stale drafts require refresh; restoring appends a checkpoint and keeps every later version', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    const store = await api.openIndexedDbHistoryStore('stale-draft');
    const history = api.openDocxHistory(store.storage);
    const doc = history.document('agreement');
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z' };
    const original = api.createBlankDocx();
    const commands = await api.HistoryCheckpoints.open(doc, store.journal('agreement'));
    const first = await commands.save(original, metadata);
    const session = api.openDocxSession(original);
    const paragraph = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'))!;
    const edit = session.replaceText(paragraph, 'My unsaved agreement draft');
    if (!edit.success) throw new Error(JSON.stringify(edit));
    const draft = session.save(); session.close();
    await doc.createVersion(first.head, original, metadata, 'other-tab');
    let stale = '';
    try { await commands.save(draft, metadata); } catch (error) { stale = (error as Error).message; }
    const needsRefresh = commands.needsRefresh;
    const pending = commands.hasPending;
    await commands.refresh();
    const saved = await commands.save(draft, metadata);
    const restored = await commands.restore(first.version.id, { ...metadata, label: 'Restore original' });
    const retained = await doc.exportDocx(saved.version.id);
    const latest = await doc.exportDocx(restored.version.id);
    const count = (await doc.listVersions()).versions.length;
    history.close(); store.close();
    return { stale, needsRefresh, pending, count, source: restored.version.record.restoredFrom,
      first: first.version.id, retained: Array.from(retained), draft: Array.from(draft), latest: Array.from(latest), original: Array.from(original) };
  });
  expect(result.stale).toBeTruthy();
  expect(result.needsRefresh).toBe(true);
  expect(result.pending).toBe(false);
  expect(result.count).toBe(4);
  expect(result.source).toEqual(result.first);
  expect(result.retained).toEqual(result.draft);
  expect(result.latest).toEqual(result.original);
});

test('independent database connections serialize head initialization, compare-and-swap, and pending requests', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    const a = await api.openIndexedDbHistoryStore('two-tabs');
    const b = await api.openIndexedDbHistoryStore('two-tabs');
    const bytes = new Uint8Array([1, 2, 3]);
    const value = Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256', bytes)), n => n.toString(16).padStart(2, '0')).join('');
    const state = { digest: { algorithm: 'SHA-256' as const, value }, length: bytes.length };
    await a.storage.putBlob(state, bytes); bytes.fill(0);
    const head = { revision: '9007199254740993', state };
    const [initialized, advanced] = await Promise.all([
      a.storage.initializeHead!('agreement', head), b.storage.advanceHead('agreement', null, state),
    ]);
    const initialHead = await b.storage.readHead('agreement');
    const winners = await Promise.all([a.storage.advanceHead('agreement', initialHead, state), b.storage.advanceHead('agreement', initialHead, state)]);
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z' };
    const first = await a.journal('agreement').put({ id: 'one', documentId: 'agreement', kind: 'save', head: null, bytes, metadata });
    const second = await b.journal('agreement').put({ ...first, id: 'two', metadata: { ...metadata, author: 'Someone else' } });
    await b.journal('agreement').remove('two');
    const remaining = await a.journal('agreement').read();
    let mismatch = '';
    try { await b.storage.putBlob(state, bytes); } catch (error) { mismatch = (error as InstanceType<typeof api.DocxHistoryError>).code; }
    const retained = await b.storage.readBlob(state);
    a.close(); b.close();
    return { initialized, advanced, initialHead, winners: winners.filter(Boolean).length,
      first, second, remaining, mismatch, retained: Array.from(retained!) };
  });
  expect(Number(result.initialized.initialized) + Number(result.advanced !== null)).toBe(1);
  expect(result.initialHead).toEqual(result.initialized.initialized ? result.initialized.head : result.advanced);
  expect(result.winners).toBe(1);
  expect(result.first).toEqual(result.second);
  expect(result.remaining).toEqual(result.first);
  expect(result.mismatch).toBe('PayloadMismatch');
  expect(result.retained).toEqual([1, 2, 3]);
});

test('checkpoint guards refuse a foreign document, an empty retry and a re-entrant command', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    // Only the contract HistoryCheckpoints relies on: insert-or-return-existing per document, and a
    // remove that ignores an ID it does not hold. Durability is IndexedDB's job, tested separately.
    const journal = (seed: HistoryCheckpointRequest | null = null) => {
      let stored = seed;
      return {
        read: async () => stored,
        put: async (request: HistoryCheckpointRequest) => (stored ??= request),
        remove: async (requestId: string) => { if (stored?.id === requestId) stored = null; },
      };
    };
    const client = api.openDocxHistory(api.createMemoryHistoryStorage());
    const doc = client.document('agreement');
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z' };
    const code = (error: unknown) => (error as InstanceType<typeof api.DocxHistoryError>).code;

    let foreignDocument = '';
    try {
      await api.HistoryCheckpoints.open(doc, journal({ id: crypto.randomUUID(), documentId: 'a-different-agreement',
        kind: 'save', head: null, bytes: api.createBlankDocx(), metadata }));
    } catch (error) { foreignDocument = code(error); }

    const commands = await api.HistoryCheckpoints.open(doc, journal());
    let emptyRetry = '';
    try { await commands.retry(); } catch (error) { emptyRetry = code(error); }
    const pendingAfterEmptyRetry = commands.hasPending;

    // save() reaches its exclusion guard synchronously, so the second command sees the first running.
    const running = commands.save(api.createBlankDocx(), metadata);
    let reentrant = '';
    try { await commands.save(api.createBlankDocx(), metadata); } catch (error) { reentrant = code(error); }
    await running;
    const afterRunning = await commands.refresh();
    client.close();
    return { foreignDocument, emptyRetry, pendingAfterEmptyRetry, reentrant, published: afterRunning?.head.revision };
  });
  expect(result.foreignDocument).toBe('InvalidRequest');
  expect(result.emptyRetry).toBe('InvalidRequest');
  expect(result.pendingAfterEmptyRetry).toBe(false);
  expect(result.reentrant).toBe('Busy');
  // The rejected re-entrant command neither published a second version nor blocked the first.
  expect(result.published).toBe('1');
});

test('a competing tab pending checkpoint is adopted: the local draft is refused and the original publishes', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    let stored: HistoryCheckpointRequest | null = null;
    const journal = {
      read: async () => stored,
      put: async (request: HistoryCheckpointRequest) => (stored ??= request),
      remove: async (requestId: string) => { if (stored?.id === requestId) stored = null; },
    };
    const client = api.openDocxHistory(api.createMemoryHistoryStorage());
    const doc = client.document('agreement');
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z' };

    const distinct = (text: string) => {
      const session = api.openDocxSession(api.createBlankDocx());
      const paragraph = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'))!;
      const edit = session.replaceText(paragraph, text);
      if (!edit.success) throw new Error(JSON.stringify(edit));
      const bytes = session.save(); session.close(); return bytes;
    };
    const otherBytes = distinct('The other tab agreement');
    const myBytes = distinct('My own agreement draft');

    // This tab opens with an empty journal; the competing tab records its request afterwards.
    const commands = await api.HistoryCheckpoints.open(doc, journal);
    const otherRequest: HistoryCheckpointRequest = { id: crypto.randomUUID(), documentId: 'agreement',
      kind: 'save', head: null, bytes: otherBytes, metadata: { ...metadata, label: 'Other tab draft' } };
    await journal.put(otherRequest);

    let refused = '';
    try { await commands.save(myBytes, { ...metadata, label: 'My draft' }); }
    catch (error) { refused = (error as InstanceType<typeof api.DocxHistoryError>).code; }
    const adopted = commands.hasPending;

    const recovered = await commands.retry();
    const published = await doc.exportDocx(recovered.version.id);
    const listed = await doc.listVersions();
    client.close();
    return { refused, adopted, remaining: stored, label: recovered.version.record.metadata.label,
      labels: listed.versions.map(version => version.record.metadata.label),
      published: Array.from(published), other: Array.from(otherBytes), mine: Array.from(myBytes) };
  });
  expect(result.refused).toBe('PendingRequest');
  expect(result.adopted).toBe(true);
  expect(result.label).toBe('Other tab draft');
  // The local draft never reached storage: exactly one version exists, and it is the other tab's.
  expect(result.labels).toEqual(['Other tab draft']);
  expect(result.published).toEqual(result.other);
  expect(result.published).not.toEqual(result.mine);
  expect(result.remaining).toBeNull();
});

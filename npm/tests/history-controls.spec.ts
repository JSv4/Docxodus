import { readFile } from 'node:fs/promises';
import { test, expect } from './history-controls-harness.js';
import type { HistoryControls } from '../src/history-controls.js';
import type { HistoryCheckpointRequest } from '../src/history-checkpoints.js';

declare global {
  interface Window {
    historyPanel: HistoryControls;
    historyPreview: { title: string; bytes: number[]; revisions: number; markdown: string } | null;
    releaseHistory: () => void;
    disposeHistory: () => Promise<void>;
    historyAcknowledgements: Array<{ action: string; label: string | null | undefined; revision: string }>;
  }
}

test('browse, page, preview, compare and share a real agreement without changing its history', async ({ page }, testInfo) => {
  const archive = (await readFile('../TestFiles/HistoryArchive/agreement.docxhistory')).toString('base64');
  await page.evaluate(async encoded => {
    const api = window.historyApi;
    const reader = await api.openDocxHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0)));
    window.historyPreview = null;
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, {
      reader, pageSize: 2, documentName: 'Agreement.docx',
      preview(bytes, title) {
        const session = api.openDocxSession(bytes);
        window.historyPreview = { title, bytes: Array.from(bytes), revisions: session.listRevisions().length, markdown: session.project().markdown };
        session.close();
      },
    });
    await window.historyPanel.ready;
    window.disposeHistory = async () => { await window.historyPanel.destroy(); reader.close(); };
  }, archive);
  const versions = page.getByLabel('Version', { exact: true });
  await expect(page.getByRole('button', { name: 'Save checkpoint', exact: true })).toBeHidden();
  await expect(versions.locator('option')).toHaveCount(2);
  await page.getByRole('button', { name: 'Load older versions' }).click();
  await expect(versions.locator('option')).toHaveCount(4);
  await expect(page.getByRole('button', { name: 'Load older versions' })).toBeHidden();
  await versions.selectOption('3');
  await page.getByRole('button', { name: 'Preview selected' }).click();
  await expect.poll(() => page.evaluate(() => window.historyPreview?.bytes.length)).toBeGreaterThan(100);
  const preview = await page.evaluate(() => window.historyPreview!);
  expect(Buffer.from(preview.bytes)).toEqual(await readFile('../TestFiles/HistoryArchive/agreement-v1.docx'));
  expect(preview.markdown).toContain('Services');
  await versions.selectOption('2');
  await page.getByLabel('Compare from', { exact: true }).selectOption('3');
  await page.getByRole('button', { name: 'Compare versions' }).click();
  await expect.poll(() => page.evaluate(() => window.historyPreview?.revisions)).toBeGreaterThan(0);
  expect((await page.evaluate(() => window.historyPreview!.title))).toBe('Comparison with tracked changes');
  const downloadEvent = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download selected' }).click();
  const download = await downloadEvent;
  await download.saveAs(testInfo.outputPath(download.suggestedFilename()));
  expect(await readFile(testInfo.outputPath(download.suggestedFilename()))).toEqual(await readFile('../TestFiles/HistoryArchive/agreement-v2.docx'));
  await page.getByText('Download with history', { exact: true }).click();
  await expect(page.getByText(/Includes retained drafts and collaboration proposals/)).toBeVisible();
  const historyDownload = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download .docxhistory', exact: true }).click();
  const portable = await historyDownload;
  expect(portable.suggestedFilename()).toBe('Agreement.docxhistory');
  await portable.saveAs(testInfo.outputPath('shared-agreement.docxhistory'));
  const count = await page.evaluate(async encoded => {
    const reader = await window.historyApi.openDocxHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0)));
    const count = (await reader.listVersions()).versions.length; reader.close(); return count;
  }, (await readFile(testInfo.outputPath('shared-agreement.docxhistory'))).toString('base64'));
  expect(count).toBe(4);
  await page.setViewportSize({ width: 390, height: 844 });
  expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth)).toBe(true);
  await page.screenshot({ path: testInfo.outputPath('history-phone.png'), fullPage: true });
  await page.evaluate(() => window.disposeHistory());
  await expect(page.getByRole('region', { name: 'Document history' })).toHaveCount(0);
});

test('saving disables conflicting controls and a lost restore acknowledgement retries its original target', async ({ page }) => {
  const archive = (await readFile('../TestFiles/HistoryArchive/agreement.docxhistory')).toString('base64');
  await page.evaluate(async encoded => {
    const api = window.historyApi;
    const store = await api.openIndexedDbHistoryStore('panel-restore');
    let pause = false; let loseAck = false;
    const history = api.openDocxHistory({ ...store.storage, async advanceHead(...args) {
      if (pause) { pause = false; await new Promise<void>(resolve => { window.releaseHistory = resolve; }); }
      const head = await store.storage.advanceHead(...args);
      if (loseAck) { loseAck = false; throw new Error('Lost restore acknowledgement'); }
      return head;
    } });
    const imported = await history.importHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0)));
    const doc = history.document(imported.archive.documentId);
    const checkpoints = await api.HistoryCheckpoints.open(doc, store.journal(doc.documentId));
    window.historyAcknowledgements = [];
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, {
      reader: doc, checkpoints, capture: () => api.createBlankDocx(), preview: () => { throw new Error('Unexpected draft replacement'); },
      onCheckpoint: (view, action) => {
        window.historyAcknowledgements.push({ action, label: view.version.record.metadata.label, revision: view.head.revision });
      },
    });
    await window.historyPanel.ready; pause = true;
    window.disposeHistory = async () => {
      const versions = await doc.listVersions();
      if (versions.versions.length !== 6) throw new Error(`Expected six checkpoints, found ${versions.versions.length}`);
      const latest = await doc.read();
      const bytes = await doc.exportDocx(latest!.version.id);
      window.historyPreview = { title: latest!.version.record.metadata.label!, bytes: Array.from(bytes), markdown: '', revisions: 0 };
      await window.historyPanel.destroy(); history.close(); store.close();
    };
    // The next call saves; the following restore loses its acknowledgement.
    const save = checkpoints.save.bind(checkpoints);
    checkpoints.save = async (...args) => { const view = await save(...args); loseAck = true; return view; };
  }, archive);
  await page.getByLabel('Checkpoint name (optional)').fill('Working draft');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(page.getByRole('region', { name: 'Document history' })).toHaveAttribute('aria-busy', 'true');
  await expect(page.getByRole('button', { name: 'Restore selected' })).toBeDisabled();
  await expect(page.getByRole('button', { name: 'Refresh history' })).toBeDisabled();
  await expect.poll(() => page.evaluate(() => typeof window.releaseHistory)).toBe('function');
  await page.evaluate(() => window.releaseHistory());
  await expect(page.getByRole('status')).toContainText('Checkpoint saved');
  await page.getByLabel('Version', { exact: true }).selectOption('4');
  await page.getByLabel('Checkpoint name (optional)').fill('Return to original');
  await page.getByRole('button', { name: 'Restore selected' }).click();
  await expect(page.getByRole('status')).toContainText('could not be confirmed');
  await expect(page.getByRole('button', { name: 'Save checkpoint', exact: true })).toBeDisabled();
  await page.getByLabel('Version', { exact: true }).selectOption('0');
  await page.getByLabel('Checkpoint name (optional)').fill('Changed after failure');
  await page.getByRole('button', { name: 'Retry checkpoint' }).click();
  await expect(page.getByRole('status')).toContainText('Checkpoint recovered');
  await page.evaluate(() => window.disposeHistory());
  const recovered = await page.evaluate(() => window.historyPreview!);
  expect(recovered.title).toBe('Return to original');
  expect(await page.evaluate(() => window.historyAcknowledgements)).toEqual([
    { action: 'save', label: 'Working draft', revision: '5' },
    { action: 'retry', label: 'Return to original', revision: '6' },
  ]);
  expect(Buffer.from(recovered.bytes)).toEqual(await readFile('../TestFiles/HistoryArchive/agreement-v1.docx'));
});

test('resolved collaboration conflicts remain distinguishable and retain the exact proposal download', async ({ page }, testInfo) => {
  const archive = (await readFile('../TestFiles/HistoryArchive/charter-collaboration.docxhistory')).toString('base64');
  await page.evaluate(async encoded => {
    const api = window.historyApi;
    const reader = await api.openDocxHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0)));
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, { reader, pageSize: 1, preview() {} });
    await window.historyPanel.ready;
    window.disposeHistory = async () => { await window.historyPanel.destroy(); reader.close(); };
  }, archive);
  await page.getByText('Recorded collaboration', { exact: true }).click();
  await page.getByRole('button', { name: 'Load activity' }).click();
  await expect(page.getByRole('listitem')).toHaveCount(1);
  await page.getByRole('button', { name: 'Load older activity' }).click();
  await page.getByRole('button', { name: 'Load older activity' }).click();
  await expect(page.getByRole('listitem')).toHaveCount(3);
  await expect(page.getByRole('button', { name: 'Load older activity' })).toBeHidden();
  const conflict = page.getByRole('listitem').filter({ hasText: 'Conflict resolved' });
  await expect(conflict).toHaveCount(1);
  await expect(page.getByText('Conflict needs review', { exact: false })).toHaveCount(0);
  await conflict.getByRole('button', { name: 'View decision' }).click();
  await expect(conflict).toContainText('Document edit');
  const downloadEvent = page.waitForEvent('download');
  await conflict.getByRole('button', { name: 'Download proposal' }).click();
  const download = await downloadEvent;
  await download.saveAs(testInfo.outputPath('retained-proposal.docx'));
  expect(await readFile(testInfo.outputPath('retained-proposal.docx'))).toEqual(await readFile('../TestFiles/HistoryArchive/charter-collaboration-conflicting-proposal.docx'));
  await page.evaluate(() => window.disposeHistory());
});

test('a captured editor view stays pinned when mounting, and closing a slow preview waits without repainting', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    const store = await api.openIndexedDbHistoryStore('captured-editor');
    let entered!: () => void; let release!: () => void;
    const started = new Promise<void>(resolve => { entered = resolve; });
    const gate = new Promise<void>(resolve => { release = resolve; });
    let slow = false;
    const client = api.openDocxHistory({ ...store.storage, async readBlob(reference) {
      if (slow) { entered(); await gate; } return store.storage.readBlob(reference);
    } });
    const doc = client.document('agreement');
    const bytes = api.createBlankDocx();
    const metadata = { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z' };
    const first = await doc.createVersion(null, bytes, metadata, 'original');
    await doc.createVersion(first.head, bytes, { ...metadata, label: 'A newer version' }, 'newer');
    const checkpoints = await api.HistoryCheckpoints.open(doc, store.journal(doc.documentId), first);
    let previews = 0;
    const panel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, {
      reader: doc, checkpoints, capture: () => bytes, preview: () => { previews++; },
    });
    await panel.ready;
    const captured = checkpoints.view!.head;
    let stale = '';
    try { await checkpoints.save(bytes, metadata); } catch (error) { stale = (error as InstanceType<typeof api.DocxHistoryError>).code; }
    slow = true;
    Array.from(panel.element.querySelectorAll('button')).find(button => button.textContent === 'Preview selected')!.click();
    await started;
    let closed = false;
    const disposal = panel.destroy().then(() => { closed = true; });
    await Promise.resolve();
    const waits = !closed; release(); await disposal;
    const retained = (await doc.listVersions()).versions.length;
    client.close(); store.close();
    return { captured, first: first.head, stale, waits, previews, retained, empty: document.querySelector('#controls')!.childElementCount === 0 };
  });
  expect(result.captured).toEqual(result.first);
  expect(result.stale).toBe('StaleHead');
  expect(result.waits).toBe(true);
  expect(result.previews).toBe(0);
  expect(result.retained).toBe(2);
  expect(result.empty).toBe(true);
});

test('keyboard version selection and a local-time lookup export the intended agreement', async ({ page }) => {
  const original = (await readFile('../TestFiles/HistoryArchive/agreement-v1.docx')).toString('base64');
  const time = await page.evaluate(async encoded => {
    const api = window.historyApi;
    const client = api.openDocxHistory(api.createMemoryHistoryStorage());
    const doc = client.document('time-lookup');
    const first = await doc.createVersion(null, Uint8Array.from(atob(encoded), c => c.charCodeAt(0)),
      { author: 'Taylor', createdAt: '2026-09-01T12:00:00Z', label: '<img src=x onerror=alert(1)>' }, 'first');
    await doc.createVersion(first.head, api.createBlankDocx(), { author: 'Taylor', createdAt: '2026-09-01T14:00:00Z', label: 'Later draft' }, 'second');
    window.historyPreview = null;
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, {
      reader: doc, preview(bytes, title) { window.historyPreview = { bytes: Array.from(bytes), title, markdown: '', revisions: 0 }; },
    });
    await window.historyPanel.ready;
    window.disposeHistory = async () => { await window.historyPanel.destroy(); client.close(); };
    const local = new Date('2026-09-01T13:00:00Z');
    return new Date(local.getTime() - local.getTimezoneOffset() * 60_000).toISOString().slice(0, 16);
  }, original);
  const versions = page.getByLabel('Version', { exact: true });
  await versions.focus(); await page.keyboard.press('ArrowDown');
  await expect(versions).toHaveValue('1');
  await expect(page.locator('#controls img')).toHaveCount(0);
  await expect(versions.locator('option').last()).toContainText('<img src=x onerror=alert(1)>');
  await page.getByText('Find a version by time', { exact: true }).click();
  await page.getByLabel('Saved at or before (local time)').fill(time);
  await page.getByRole('button', { name: 'Preview at time' }).click();
  await expect.poll(() => page.evaluate(() => window.historyPreview?.title)).toBe('Version at selected time');
  expect(Buffer.from(await page.evaluate(() => window.historyPreview!.bytes))).toEqual(Buffer.from(original, 'base64'));
  await page.evaluate(() => window.disposeHistory());
});

test('mounting refuses an out-of-range page size and controls bound to another document', async ({ page }) => {
  const result = await page.evaluate(async () => {
    const api = window.historyApi;
    const client = api.openDocxHistory(api.createMemoryHistoryStorage());
    const reader = client.document('agreement');
    const elsewhere = client.document('a-different-agreement');
    let stored: HistoryCheckpointRequest | null = null;
    const checkpoints = await api.HistoryCheckpoints.open(elsewhere, {
      read: async () => stored,
      put: async (request: HistoryCheckpointRequest) => (stored ??= request),
      remove: async (requestId: string) => { if (stored?.id === requestId) stored = null; },
    });
    const container = document.querySelector<HTMLElement>('#controls')!;
    const sizes: string[] = [];
    for (const pageSize of [0, 101, 2.5]) {
      try { api.mountHistoryControls(container, { reader, preview() {}, pageSize }); sizes.push('mounted'); }
      catch (error) { sizes.push((error as Error).name); }
    }
    let mismatch = '';
    try { api.mountHistoryControls(container, { reader, checkpoints, preview() {} }); }
    catch (error) { mismatch = (error as Error).message; }
    client.close();
    return { sizes, mismatch, mounted: container.childElementCount };
  });
  expect(result.sizes).toEqual(['RangeError', 'RangeError', 'RangeError']);
  expect(result.mismatch).toContain('same document');
  // A refused mount leaves the host's container exactly as it found it.
  expect(result.mounted).toBe(0);
});

test('history control errors stay distinct and name the failure a host can act on', async ({ page }) => {
  const messages = await page.evaluate(() => {
    const api = window.historyApi;
    const of = (code: string, pending = false) =>
      api.historyControlError(new api.DocxHistoryError(code, `raw ${code} detail`), pending);
    return {
      stale: of('StaleHead'), conflict: of('ImportConflict'), unsupported: of('InitializationUnsupported'),
      resource: of('ResourceLimit'), unsupportedVersion: of('UnsupportedVersion'), damaged: of('InvalidManifest'),
      // Codes with no dedicated explanation still reach the host: pending draws attention to the
      // unconfirmed checkpoint, everything else surfaces the underlying message.
      pending: of('Timeout', true), unrecognised: of('Timeout'),
      // An unconfirmed checkpoint outranks a file-shaped failure, because the draft is the thing
      // at risk. Codes that already say what to do next keep their own instruction.
      pendingFile: of('ResourceLimit', true), pendingStale: of('StaleHead', true),
      plain: api.historyControlError(new Error('network unavailable')),
      thrownValue: api.historyControlError('not an error object'),
    };
  });
  expect(messages.stale).toContain('Your draft is safe');
  expect(messages.conflict).toContain('read-only');
  expect(messages.unsupported).toContain('import history');
  expect(messages.resource).toContain('64 MiB');
  expect(messages.unsupportedVersion).toContain('unsupported version');
  expect(messages.damaged).toContain('damaged or incomplete');
  expect(messages.pendingFile).toBe(messages.pending);
  expect(messages.pendingStale).toBe(messages.stale);
  expect(messages.pending).toContain('Retry checkpoint');
  expect(messages.unrecognised).toContain('raw Timeout detail');
  expect(messages.plain).toContain('network unavailable');
  expect(messages.thrownValue).toContain('Please try again.');
  // Nothing else collapses into a single unhelpful sentence.
  const distinct = Object.entries(messages).filter(([key]) => !key.startsWith('pending') || key === 'pending');
  expect(new Set(distinct.map(([, text]) => text)).size).toBe(distinct.length);
});

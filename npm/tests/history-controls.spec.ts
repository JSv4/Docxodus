import { readFile } from 'node:fs/promises';
import { test, expect } from './history-controls-harness.js';
import type { HistoryControls } from '../src/history-controls.js';

declare global {
  interface Window {
    historyPanel: HistoryControls;
    historyPreview: { title: string; bytes: number[]; revisions: number; markdown: string } | null;
    releaseHistory: () => void;
    disposeHistory: () => Promise<void>;
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
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, {
      reader: doc, checkpoints, capture: () => api.createBlankDocx(), preview: () => { throw new Error('Unexpected draft replacement'); },
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
  expect(Buffer.from(recovered.bytes)).toEqual(await readFile('../TestFiles/HistoryArchive/agreement-v1.docx'));
});

test('resolved collaboration conflicts remain distinguishable and retain the exact proposal download', async ({ page }, testInfo) => {
  const archive = (await readFile('../TestFiles/HistoryArchive/charter-collaboration.docxhistory')).toString('base64');
  await page.evaluate(async encoded => {
    const api = window.historyApi;
    const reader = await api.openDocxHistoryArchive(Uint8Array.from(atob(encoded), c => c.charCodeAt(0)));
    window.historyPanel = api.mountHistoryControls(document.querySelector<HTMLElement>('#controls')!, { reader, preview() {} });
    await window.historyPanel.ready;
    window.disposeHistory = async () => { await window.historyPanel.destroy(); reader.close(); };
  }, archive);
  await page.getByText('Recorded collaboration', { exact: true }).click();
  await page.getByRole('button', { name: 'Load activity' }).click();
  const conflict = page.getByRole('listitem').filter({ hasText: 'Conflict resolved' });
  await expect(conflict).toHaveCount(1);
  await expect(page.getByText('Conflict needs review', { exact: false })).toHaveCount(0);
  const downloadEvent = page.waitForEvent('download');
  await conflict.getByRole('button', { name: 'Download proposal' }).click();
  const download = await downloadEvent;
  await download.saveAs(testInfo.outputPath('retained-proposal.docx'));
  expect(await readFile(testInfo.outputPath('retained-proposal.docx'))).toEqual(await readFile('../TestFiles/HistoryArchive/charter-collaboration-conflicting-proposal.docx'));
  await page.evaluate(() => window.disposeHistory());
});

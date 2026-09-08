import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { build } from 'esbuild';

let script: string;
let html: string;
test.beforeAll(async () => {
  script = (await build({ entryPoints: ['examples/history.ts'], bundle: true, format: 'esm', write: false })).outputFiles[0].text;
  html = await readFile('examples/history.html', 'utf8');
});
test.beforeEach(async ({ context, page }) => {
  await context.route('**/history.html', route => route.fulfill({ contentType: 'text/html', body: html }));
  await context.route('**/history-example.js', route => route.fulfill({ contentType: 'text/javascript', body: script }));
  page.on('dialog', dialog => void dialog.accept());
  await page.goto('/history.html');
  await expect(page.getByRole('button', { name: 'New document', exact: true })).toBeEnabled();
});

test('resume a portable agreement, save a checkpoint, reopen after reload, and keep conflicting imports read-only', async ({ page }, testInfo) => {
  const file = page.getByLabel('Open a DOCX or history file');
  const versions = page.getByLabel('Version', { exact: true });
  await file.setInputFiles('../TestFiles/HistoryArchive/agreement.docxhistory');
  await expect(versions.locator('option')).toHaveCount(4);
  await expect(page.getByRole('button', { name: 'Save checkpoint', exact: true })).toBeHidden();
  await expect(page.locator('#editor')).not.toContainText('Services');
  await versions.selectOption('3');
  await page.getByRole('button', { name: 'Preview selected' }).click();
  await expect(page.getByRole('dialog')).toBeVisible();
  await expect(page.locator('#preview')).toContainText('Services');
  await expect(page.getByRole('button', { name: 'Use as draft' })).toBeHidden();
  await page.keyboard.press('Escape');
  await page.getByRole('button', { name: 'Resume editing this history' }).click();
  await expect(page.locator('#editor')).toContainText('Services');
  await expect(page.getByRole('button', { name: 'Save checkpoint', exact: true })).toBeEnabled();
  await page.getByLabel('Checkpoint name (optional)').fill('Browser review');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(versions.locator('option')).toHaveCount(5);
  await page.reload();
  await expect(versions.locator('option')).toHaveCount(5);
  await expect(versions).toHaveValue('0');
  await expect(versions.locator('option').first()).toContainText('Browser review');
  await expect(page.locator('#editor')).toContainText('Services');
  await page.getByText('Download with history', { exact: true }).click();
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download .docxhistory', exact: true }).click();
  await (await downloaded).saveAs(testInfo.outputPath('reviewed-agreement.docxhistory'));
  await file.setInputFiles('../TestFiles/HistoryArchive/agreement.docxhistory');
  await page.getByRole('button', { name: 'Resume editing this history' }).click();
  await expect(page.locator('#message')).toContainText('different local history');
  await expect(versions.locator('option')).toHaveCount(4);
  await expect(page.getByRole('button', { name: 'Preview selected' })).toBeEnabled();
  await expect(page.locator('#editor')).toContainText('Services');
  await page.getByRole('button', { name: 'Back to my draft' }).click();
  await expect(versions.locator('option')).toHaveCount(5);
  await page.screenshot({ path: testInfo.outputPath('history-editor.png'), fullPage: true });
});

test('a stale save and malformed or oversized uploads preserve the open draft and its local history', async ({ context, page }) => {
  const file = page.getByLabel('Open a DOCX or history file');
  await file.setInputFiles('../TestFiles/HistoryArchive/agreement-v1.docx');
  await expect(page.locator('#editor')).toContainText('Services');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(1);
  const other = await context.newPage(); other.on('dialog', dialog => void dialog.accept());
  await other.goto('/history.html');
  await expect(other.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(1);
  await page.getByLabel('Checkpoint name (optional)').fill('Another editor saved');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(2);
  const draft = other.locator('#editor [data-anchor][contenteditable="true"]').first();
  await draft.click(); await other.keyboard.press('End');
  await draft.pressSequentially(' — My unsaved agreement changes');
  await other.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(other.locator('#history [role="status"]')).toContainText('A newer checkpoint exists');
  await expect(draft).toContainText('My unsaved agreement changes');
  await expect(other.getByRole('button', { name: 'Save checkpoint', exact: true })).toBeDisabled();
  for (const name of ['broken.docxhistory', 'broken.docx']) {
    await other.getByLabel('Open a DOCX or history file').setInputFiles({ name, mimeType: 'application/octet-stream', buffer: Buffer.from('This is not a document archive') });
    await expect(other.locator('#message')).toContainText('Your document is unchanged');
    await expect(draft).toContainText('My unsaved agreement changes');
  }
  await other.evaluate(() => {
    const file = new File([new Uint8Array(64 * 1024 * 1024 + 1)], 'too-large.docxhistory');
    file.arrayBuffer = async () => { throw new Error('Oversized uploads must be rejected before reading bytes'); };
    const transfer = new DataTransfer(); transfer.items.add(file);
    const input = document.querySelector<HTMLInputElement>('#file')!; input.files = transfer.files;
    input.dispatchEvent(new Event('change'));
  });
  await expect(other.locator('#message')).toContainText('browser processing limits');
  await expect(draft).toContainText('My unsaved agreement changes');
  await other.getByRole('button', { name: 'Refresh history' }).click();
  await expect(other.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(2);
  await other.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(other.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(3);
  await other.getByRole('button', { name: 'Open latest' }).click();
  await expect(other.locator('#preview')).toContainText('My unsaved agreement changes');
  await other.getByRole('button', { name: 'Close preview' }).click();
  await other.close();
});

test('saved agreements need no discard warning, while newer edits remain protected', async ({ page }) => {
  const confirmations: string[] = [];
  page.on('dialog', dialog => { if (dialog.type() === 'confirm') confirmations.push(dialog.message()); });
  await page.getByLabel('Open a DOCX or history file').setInputFiles('../TestFiles/HistoryArchive/agreement-v1.docx');
  await expect(page.locator('#editor')).toContainText('Services');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(page.locator('#history [role="status"]')).toContainText('Checkpoint saved');
  await page.getByRole('button', { name: 'New document', exact: true }).click();
  await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(0);
  expect(confirmations).toEqual([]);

  const draft = page.locator('#editor [data-anchor][contenteditable="true"]').first();
  await draft.click(); await draft.pressSequentially('Terms to retain in a checkpoint');
  await page.getByRole('button', { name: 'Save checkpoint', exact: true }).click();
  await expect(page.locator('#history [role="status"]')).toContainText('Checkpoint saved');
  await draft.click(); await page.keyboard.press('End');
  await draft.pressSequentially(' and newer unsaved changes');
  await page.getByRole('button', { name: 'New document', exact: true }).click();
  await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(0);
  expect(confirmations).toHaveLength(1);
  expect(confirmations[0]).toContain('Replace your open draft?');
});

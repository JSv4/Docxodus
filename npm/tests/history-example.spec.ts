import { test, expect, type Page, type Dialog } from '@playwright/test';

// Exercise the deployable site itself. No source bundling, request interception or engine override.
const APP = 'http://localhost:8084/demo/app.html';
const agreement = '../TestFiles/HistoryArchive/agreement-v1.docx';
async function ready(page: Page) {
  await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'ready', { timeout: 45000 });
  await expect(page.locator('[data-dxr="loader"]')).toBeHidden();
}
async function history(page: Page) {
  await page.getByRole('button', { name: 'Version history', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Save version', exact: true })).toBeEnabled();
}
async function closeHistory(page: Page) { await page.getByRole('button', { name: 'Close version history' }).click(); }
async function openAgreement(page: Page) {
  await page.getByLabel('Open a document or history file').setInputFiles(agreement);
  await expect(page.locator('[data-dxr="editor"]')).toContainText('Services');
}
async function save(page: Page, label: string) {
  await page.getByLabel('Version name (optional)').fill(label);
  await page.getByRole('button', { name: 'Save version', exact: true }).click();
  await expect(page.locator('.dx-history [role="status"]')).toContainText('Version saved');
}

// Publish successfully, but leave the durable request behind as if the IDB acknowledgement was lost.
async function loseNextJournalAcknowledgement(page: Page) {
  await page.evaluate(() => {
    const journal = (window as any).__ribbon.history.draft.checkpoints.journal;
    const remove = journal.remove.bind(journal);
    journal.remove = async (_id: string) => {
      journal.remove = remove;
      throw new Error('Lost journal acknowledgement');
    };
  });
}
async function reopenPending(page: Page) {
  const accept = (dialog: Dialog) => { void dialog.accept(); };
  page.on('dialog', accept);
  try { await page.reload(); await ready(page); }
  finally { page.off('dialog', accept); }
}
async function pendingHistory(page: Page) {
  await page.getByRole('button', { name: 'Version history', exact: true }).click();
  await expect(page.getByRole('button', { name: 'Retry save', exact: true })).toBeEnabled();
}

test.beforeEach(async ({ page }) => { await page.goto(APP); await ready(page); });

test('the built editor saves versions, resumes after reload, compares, restores and downloads portable history', async ({ page }, testInfo) => {
  await openAgreement(page); await history(page); await save(page, 'First draft');
  const versions = page.getByLabel('Version', { exact: true });
  await expect(versions.locator('option')).toHaveCount(1);
  await closeHistory(page);
  const block = page.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
  await block.click(); await page.keyboard.press('End'); await block.pressSequentially(' — Terms reviewed');
  await history(page); await save(page, 'Reviewed');
  await expect(versions.locator('option')).toHaveCount(2);
  await page.reload(); await ready(page);
  await expect(page.locator('[data-dxr="editor"]')).toContainText('Terms reviewed');
  await history(page); await expect(versions.locator('option')).toHaveCount(2);
  await page.locator('summary').filter({ hasText: 'Compare versions' }).click();
  await page.getByRole('button', { name: 'Compare versions', exact: true }).click();
  await expect(page.getByRole('dialog', { name: 'Version preview' })).toBeVisible();
  await page.getByRole('button', { name: 'Back to version history' }).click();
  await versions.selectOption('1');
  page.once('dialog', dialog => void dialog.dismiss());
  await page.getByRole('button', { name: 'Restore selected' }).click();
  await expect(versions.locator('option')).toHaveCount(2);
  page.once('dialog', dialog => void dialog.accept());
  await page.getByRole('button', { name: 'Restore selected' }).click();
  await expect(versions.locator('option')).toHaveCount(3);
  await expect(page.locator('[data-dxr="editor"]')).not.toContainText('Terms reviewed');
  await page.getByText('Download with history', { exact: true }).click();
  const downloaded = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Download with version history', exact: true }).click();
  await (await downloaded).saveAs(testInfo.outputPath('reviewed-agreement.docxhistory'));
  await page.getByRole('dialog', { name: 'Version history', exact: true }).evaluate(el => { el.scrollTop = 0; });
  await page.screenshot({ path: testInfo.outputPath('version-history-desktop.png'), fullPage: true });
});

test('a history file opens separately; importing, conflicts and malformed uploads preserve the document', async ({ page }) => {
  const file = page.getByLabel('Open a document or history file');
  await file.setInputFiles('../TestFiles/HistoryArchive/agreement.docxhistory');
  const versions = page.getByLabel('Version', { exact: true });
  await expect(versions.locator('option')).toHaveCount(4);
  await expect(page.getByRole('button', { name: 'Save version', exact: true })).toBeHidden();
  await expect(page.locator('[data-dxr="editor"]')).not.toContainText('Services');
  await page.getByRole('button', { name: 'Preview selected' }).click();
  await expect(page.getByRole('dialog', { name: 'Version preview' })).toContainText('Services');
  await expect(page.getByRole('button', { name: 'Use as draft' })).toBeHidden();
  await page.getByRole('button', { name: 'Back to version history' }).click();
  await page.getByRole('button', { name: 'Continue editing this document' }).click();
  await expect(page.locator('[data-dxr="editor"]')).toContainText('Services');
  await save(page, 'Browser review');
  await expect(versions.locator('option')).toHaveCount(5);
  await closeHistory(page);
  await file.setInputFiles('../TestFiles/HistoryArchive/agreement.docxhistory');
  await page.getByRole('button', { name: 'Continue editing this document' }).click();
  await expect(page.locator('.dxr-history-status')).toContainText('different local history');
  await page.getByRole('button', { name: 'Back to my document' }).click();
  await expect(versions.locator('option')).toHaveCount(5);
  await closeHistory(page);
  const handle = await page.evaluate(() => (window as any).__ribbon.editor.sessionHandle);
  for (const name of ['broken.docxhistory', 'broken.docx']) {
    await file.setInputFiles({ name, mimeType: 'application/octet-stream', buffer: Buffer.from('not a zip') });
    await expect(page.locator('[data-dxr="status"]')).toContainText('unchanged');
    expect(await page.evaluate(() => (window as any).__ribbon.editor.sessionHandle)).toBe(handle);
    await expect(page.locator('[data-dxr="editor"]')).toContainText('Services');
  }
});

test('two tabs keep stale drafts; canceled replacement and saved-document switching retain versions', async ({ page, context }) => {
  await openAgreement(page); await history(page); await save(page, 'Initial'); await closeHistory(page);
  const other = await context.newPage(); await other.goto(APP); await ready(other);
  await history(page); await save(page, 'Another tab'); await closeHistory(page);
  const block = other.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
  await block.click(); await other.keyboard.press('End'); await block.pressSequentially(' My unsaved draft');
  await history(other);
  await other.getByRole('button', { name: 'Save version', exact: true }).click();
  await expect(other.locator('.dx-history [role="status"]')).toContainText('A newer saved version exists');
  await expect(block).toContainText('My unsaved draft');
  await expect(other.getByRole('button', { name: 'Save version', exact: true })).toBeDisabled();
  await other.getByRole('button', { name: 'Refresh history', exact: true }).click(); await save(other, 'My draft');
  await closeHistory(other);
  await block.click(); await other.keyboard.press('End'); await block.pressSequentially(' keep this');
  other.once('dialog', dialog => void dialog.dismiss());
  await other.getByRole('button', { name: 'New', exact: true }).click();
  await expect(block).toContainText('keep this');
  await history(other); await save(other, 'Retained'); await closeHistory(other);
  await other.getByRole('button', { name: 'New', exact: true }).click();
  await history(other); await save(other, 'Second document');
  const recent = other.getByLabel('Saved documents on this device');
  await expect(recent.locator('option')).toHaveCount(3);
  const currentId = await recent.inputValue();
  await closeHistory(other);
  const newBlock = other.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
  await newBlock.fill('Keep my new draft');
  await history(other);
  other.once('dialog', dialog => void dialog.dismiss());
  await recent.selectOption({ label: 'agreement-v1.docx' });
  await expect(recent).toHaveValue(currentId);
  await expect(newBlock).toContainText('Keep my new draft');
  other.once('dialog', dialog => void dialog.accept());
  await recent.selectOption({ label: 'agreement-v1.docx' });
  await expect(other.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(4);
  await expect(other.locator('[data-dxr="editor"]')).toContainText('keep this');
  await other.close();
});

test('history is optional, storage denial leaves editing available, and the drawer fits a phone', async ({ page }, testInfo) => {
  await page.goto(APP + '?history=0'); await ready(page);
  await expect(page.getByRole('button', { name: 'Version history', exact: true })).toHaveCount(0);
  await page.goto(APP); await ready(page);
  expect(await page.evaluate(async () => (await indexedDB.databases()).length)).toBe(0);
  await page.setViewportSize({ width: 390, height: 844 }); await history(page);
  const drawer = page.getByRole('dialog', { name: 'Version history', exact: true });
  const box = await drawer.boundingBox();
  expect(box!.x).toBeGreaterThanOrEqual(0); expect(box!.x + box!.width).toBeLessThanOrEqual(390);
  expect(await drawer.evaluate(el => el.scrollWidth <= el.clientWidth)).toBe(true);
  await page.screenshot({ path: testInfo.outputPath('version-history-phone.png'), fullPage: true });
  await page.keyboard.press('Escape');
  await expect(page.getByRole('button', { name: 'Version history', exact: true })).toBeFocused();
  await page.reload(); await ready(page);
  await page.evaluate(() => { indexedDB.open = () => { throw new DOMException('Storage unavailable', 'SecurityError'); }; });
  await page.getByRole('button', { name: 'Version history', exact: true }).click();
  await expect(page.locator('.dxr-history-status')).toContainText('Storage unavailable');
  await closeHistory(page);
  await expect(page.locator('[data-dxr="editor"] [contenteditable="true"]').first()).toBeEditable();
});

for (const entry of [
  { name: 'local editor', url: 'http://localhost:8082/editor.html', tap: false },
  { name: 'landing editor', url: 'http://localhost:8084/demo/?demo=editor', tap: false },
  { name: 'compact player', url: 'http://localhost:8084/demo/player.html', tap: true },
]) test(`${entry.name} exposes the shared history drawer from the default build`, async ({ page }) => {
  const engineRequests: string[] = [];
  page.on('request', request => { if (request.url().includes('embed.bundle.js')) engineRequests.push(request.url()); });
  await page.goto(entry.url);
  if (entry.tap) await page.locator('#start').click();
  await ready(page); await history(page);
  await expect(page.getByLabel('Version name (optional)')).toBeVisible();
  expect(engineRequests.length).toBeGreaterThan(0);
  expect(engineRequests.every(url => new URL(url).origin === new URL(entry.url).origin)).toBe(true);
});

test('a document started blank also reopens its saved version after reload', async ({ page }) => {
  await page.goto(APP + '?blank=1'); await ready(page);
  const block = page.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
  await block.fill('A fresh document worth keeping');
  await history(page); await save(page, 'Started here');
  await page.reload(); await ready(page);
  await expect(page.locator('[data-dxr="editor"]')).toContainText('A fresh document worth keeping');
  await history(page);
  await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(1);
});

for (const reload of [false, true]) {
  test(`a recovered save clears unsaved warnings${reload ? ' after reload' : ''}`, async ({ page }) => {
    await openAgreement(page); await history(page);
    await loseNextJournalAcknowledgement(page);
    await page.getByRole('button', { name: 'Save version', exact: true }).click();
    await expect(page.locator('.dx-history [role="status"]')).toContainText('could not be confirmed');
    await closeHistory(page);
    if (reload) await reopenPending(page);
    await pendingHistory(page);
    await page.getByRole('button', { name: 'Retry save', exact: true }).click();
    await expect(page.locator('.dx-history [role="status"]')).toContainText('Saved version recovered');
    await closeHistory(page);
    // An unexpected unsaved-change prompt would cancel New and leave the agreement open.
    await page.getByRole('button', { name: 'New', exact: true }).click();
    await expect(page.locator('[data-dxr="docname"]')).toHaveText('Untitled.docx');
  });

  test(`a recovered restore confirms newer edits and installs its original target${reload ? ' after reload' : ''}`, async ({ page }) => {
    await openAgreement(page); await history(page); await save(page, 'Original'); await closeHistory(page);
    const block = page.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
    await block.click(); await page.keyboard.press('End'); await block.pressSequentially(' Terms reviewed');
    await history(page); await save(page, 'Reviewed');
    await page.getByLabel('Version', { exact: true }).selectOption('1');
    await loseNextJournalAcknowledgement(page);
    page.once('dialog', dialog => void dialog.accept());
    await page.getByRole('button', { name: 'Restore selected' }).click();
    await expect(page.locator('.dx-history [role="status"]')).toContainText('could not be confirmed');
    await expect(page.locator('[data-dxr="editor"]')).toContainText('Terms reviewed');
    await closeHistory(page);
    if (reload) await reopenPending(page);
    await block.click(); await page.keyboard.press('End'); await block.pressSequentially(' After failure');
    await pendingHistory(page);
    // Changing the selection cannot change the durable restore target.
    await page.getByLabel('Version', { exact: true }).selectOption('0');
    page.once('dialog', dialog => { expect(dialog.message()).toContain('unsaved changes'); void dialog.dismiss(); });
    await page.getByRole('button', { name: 'Retry save', exact: true }).click();
    await expect(page.locator('.dx-history [role="status"]')).toContainText('Restore retry canceled');
    await expect(page.locator('[data-dxr="editor"]')).toContainText('After failure');
    await expect(page.getByRole('button', { name: 'Retry save', exact: true })).toBeEnabled();
    page.once('dialog', dialog => void dialog.accept());
    await page.getByRole('button', { name: 'Retry save', exact: true }).click();
    await expect(page.locator('.dx-history [role="status"]')).toContainText('Version restored');
    await expect(page.getByLabel('Version', { exact: true }).locator('option')).toHaveCount(3);
    await expect(page.locator('[data-dxr="editor"]')).not.toContainText(/Terms reviewed|After failure/);
  });
}

test('recovering another tab’s pending save never marks this tab’s different draft as saved', async ({ page, context }) => {
  await openAgreement(page); await history(page); await save(page, 'Initial'); await closeHistory(page);
  const other = await context.newPage(); await other.goto(APP); await ready(other);
  const block = (p: Page) => p.locator('[data-dxr="editor"] [contenteditable="true"][data-anchor]').first();
  await block(page).fill('Saved in the first tab');
  await history(page); await loseNextJournalAcknowledgement(page);
  await page.getByRole('button', { name: 'Save version', exact: true }).click();
  await expect(page.locator('.dx-history [role="status"]')).toContainText('could not be confirmed');
  await block(other).fill('Unsaved in the second tab');
  await history(other);
  await other.getByRole('button', { name: 'Save version', exact: true }).click();
  await expect(other.locator('.dx-history [role="status"]')).toContainText('could not be confirmed');
  await other.getByRole('button', { name: 'Retry save', exact: true }).click();
  await expect(other.locator('.dx-history [role="status"]')).toContainText('Saved version recovered');
  await closeHistory(other);
  const warned = other.waitForEvent('dialog');
  const create = other.getByRole('button', { name: 'New', exact: true }).click();
  const dialog = await warned; expect(dialog.message()).toContain('unsaved changes');
  await dialog.dismiss(); await create;
  await expect(block(other)).toContainText('Unsaved in the second tab');
  await other.close();
});

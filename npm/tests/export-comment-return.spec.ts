import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { commentReturnFixture } from './docx-comment-return-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const { later, frame } of [{ later: false, frame: false }, { later: true, frame: false }, { later: false, frame: true }]) test(`delivered comment returns to rendered text (later=${later}, frame=${frame})`, async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(commentReturnFixture(true, later, frame), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(later ? 2 : 1);
  await writeFile(testInfo.outputPath('comment-return.pdf'), result.pdf!);
  await writeFile(testInfo.outputPath('comment-return.html'), result.html!);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(result.pageCount);
    let text = '';
    for (let number = 1; number <= pdf.numPages; number++) {
      const page = await pdf.getPage(number), viewport = page.getViewport({ scale: 1 });
      expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
      text += (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ') + ' ';
    }
    expect(text.replace(/\s+/g, ' ')).toContain('Maple navigation anchor.');
    expect(text).toContain('MapleBodyReturn');
    expect(text).toContain('MapleNote');
    if (later) expect(text.replace(/\s+/g, ' ')).toContain('Maple return target.');
  } finally { await task.destroy(); }
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    await expect(page.locator('.page-box #comment-ref-6')).toHaveCount(1);
    if (frame) await expect(page.locator('[data-docx-text-frame] #comment-ref-6')).toHaveCount(1);
    const links = page.locator('.page-comment-margin .comment-margin-backref');
    await expect(links).toHaveCount(later ? 2 : 1);
    for (const link of await links.all()) {
      await expect(link).toHaveAttribute('href', '#comment-ref-6');
      await expect(link).not.toHaveAttribute('aria-disabled', 'true');
    }
    await expect(page.locator('.page-comment-margin [id]')).toHaveCount(0);
    await links.first().focus(); await page.keyboard.press('Enter');
    expect(await page.evaluate(() => location.hash)).toBe('#comment-ref-6');
    if (later) expect(await page.evaluate(() => scrollY)).toBeGreaterThan(500);
  } finally { await page.close(); }
});

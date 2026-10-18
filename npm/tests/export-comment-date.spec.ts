import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { commentDateFixture } from './docx-comment-date-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const modern of [false, true]) test(`delivered margin date retains its UTC day and year (modern=${modern})`, async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(commentDateFixture(modern), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(1);
  await writeFile(testInfo.outputPath('comment-date.pdf'), result.pdf!);
  await writeFile(testInfo.outputPath('comment-date.html'), result.html!);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
    expect(text.replace(/\s+/g, ' ')).toContain('Cedar date anchor.');
    expect(text.replace(/\s+/g, ' ')).toContain('DateNote Check the timestamp.');
    expect(text.replace(/\s+/g, ' ')).toContain(modern ? 'Sep 10, 2026' : 'Sep 9, 2026');
  } finally { await task.destroy(); }
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    const note = page.locator('.page-comment-margin .comment-margin-note');
    await expect(note).toHaveCount(1);
    await expect(note).toHaveAttribute('data-date', modern ? '2026-09-10T04:45:00Z' : '2026-09-09T23:45:00Z');
    await expect(note.locator('.comment-margin-date')).toHaveText(modern ? 'Sep 10, 2026' : 'Sep 9, 2026');
  } finally { await page.close(); }
});

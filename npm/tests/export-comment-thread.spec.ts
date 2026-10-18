import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { commentThreadFixture } from './docx-comment-thread-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('delivered thread keeps uniform note text and contained header glyphs', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(commentThreadFixture(true, true, true), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(1);
  await writeFile(testInfo.outputPath('comment-thread.pdf'), result.pdf!);
  await writeFile(testInfo.outputPath('comment-thread.html'), result.html!);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
    expect(text.replace(/\s+/g, ' ')).toContain('Cedar thread anchor.');
    for (let i = 0; i < 3; i++) expect(text.replace(/\s/g, '')).toContain(`CedarThread${i}`);
  } finally { await task.destroy(); }
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const notes = await page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('.page-comment-margin .comment-margin-note')).map(note => {
      const box = note.getBoundingClientRect();
      const column = note.closest('.page-comment-margin')!.getBoundingClientRect();
      const body = note.querySelector('.comment-margin-note-body p')!.getBoundingClientRect();
      const header = note.querySelector('.comment-margin-note-header')!;
      return { fontSize: parseFloat(getComputedStyle(note).fontSize), width: box.width,
        bodyTop: body.top - column.top, bodyBottom: body.bottom - column.top, columnHeight: column.height,
        parts: Array.from(header.children).map(child => {
          const range = document.createRange(); range.selectNodeContents(child);
          const glyph = range.getBoundingClientRect();
          return { left: glyph.left - box.left, right: glyph.right - box.left, width: glyph.width, height: glyph.height };
        }) };
    }));
    expect(notes).toHaveLength(3);
    for (const note of notes) {
      expect(note.fontSize).toBeCloseTo(notes[0].fontSize, 2);
      expect(note.parts).toHaveLength(4);
      expect(note.bodyTop).toBeGreaterThanOrEqual(-1);
      expect(note.bodyBottom).toBeLessThanOrEqual(note.columnHeight + 1);
      for (const part of note.parts) {
        expect(part.width).toBeGreaterThan(0); expect(part.height).toBeGreaterThan(0);
        expect(part.left).toBeGreaterThanOrEqual(-1); expect(part.right).toBeLessThanOrEqual(note.width + 1);
      }
    }
  } finally { await page.close(); }
});

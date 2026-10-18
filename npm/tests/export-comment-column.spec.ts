import { test, expect } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { longCommentFixture, LONG_COMMENT_TOKEN } from './docx-long-comment-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('delivered HTML exposes the complete long note while PDF page dimensions stay unchanged', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(longCommentFixture(240, false, true), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'margin', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(1);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(viewport.width).toBeCloseTo(612, 0);
    expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
    expect(text.replace(/\s+/g, ' ')).toContain('Cedar inventory note anchor.');
  } finally { await task.destroy(); }
  const page = await browser.newPage();
  try {
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const facts = await page.evaluate((last) => {
      const column = document.querySelector<HTMLElement>('.page-comment-margin')!;
      const walker = document.createTreeWalker(column, NodeFilter.SHOW_TEXT);
      let node: Node | null, tail: Text | null = null;
      while ((node = walker.nextNode())) if (node.textContent?.includes(last)) tail = node as Text;
      if (!tail) throw new Error('The delivered HTML lost the last comment token');
      column.scrollTop = column.scrollHeight;
      const range = document.createRange(), start = tail.data.indexOf(last);
      range.setStart(tail, start); range.setEnd(tail, start + last.length);
      const rect = range.getBoundingClientRect(), box = column.getBoundingClientRect();
      return { overflow: getComputedStyle(column).overflowY, width: rect.width, height: rect.height,
        left: rect.left - box.left - column.clientLeft, right: rect.right - box.left - column.clientLeft,
        top: rect.top - box.top - column.clientTop, bottom: rect.bottom - box.top - column.clientTop,
        viewportWidth: column.clientWidth, viewportHeight: column.clientHeight,
        body: column.textContent ?? '' };
    }, LONG_COMMENT_TOKEN);
    expect(['auto', 'scroll']).toContain(facts.overflow);
    for (let i = 0; i < 240; i++) expect(facts.body).toContain(`CedarNote${String(i).padStart(3, '0')}`);
    expect(facts.width).toBeGreaterThan(0); expect(facts.height).toBeGreaterThan(0);
    expect(facts.left).toBeGreaterThanOrEqual(-1); expect(facts.right).toBeLessThanOrEqual(facts.viewportWidth + 1);
    expect(facts.top).toBeGreaterThanOrEqual(-1); expect(facts.bottom).toBeLessThanOrEqual(facts.viewportHeight + 1);
  } finally { await page.close(); }
});

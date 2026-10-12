import { expect, test } from '@playwright/test';
import { pictureDocx } from './docx-picture-fixture.js';

for (const tight of [false, true]) {
  test(`${tight ? 'tight contour' : 'square wrap'} preserves picture placement and all source text`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, false, false,
      ) as string;
      if (html.startsWith('{')) throw new Error(html);
      document.body.innerHTML = '<main id="fixture"></main>';
      const pagination = (window as any).DocxodusPagination.paginateHtml(html, document.getElementById('fixture'),
        { showPageNumbers: false, pageGap: 0 });
      const content = document.querySelector<HTMLElement>('#fixture .page-content')!;
      const image = content.querySelector<HTMLImageElement>('img')!;
      const paragraph = content.querySelector('p')!;
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      const starts = new Map<number, number>();
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = node.textContent ?? '';
        for (let i = 0; i < text.length; i++) {
          if (/\s/.test(text[i])) continue;
          const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
          const rect = range.getBoundingClientRect();
          const row = Math.round(rect.top);
          starts.set(row, Math.min(starts.get(row) ?? Infinity, rect.left - content.getBoundingClientRect().left));
        }
      }
      return { pages: pagination.totalPages, text: paragraph.textContent,
        imageX: image.getBoundingClientRect().left - content.getBoundingClientRect().left,
        imageWidth: image.getBoundingClientRect().width,
        rows: Array.from(starts).sort(([a], [b]) => a - b).map(([, x]) => x),
      };
    }, { bytes: Array.from(pictureDocx(tight)) });
    expect(result.pages).toBe(1);
    expect(result.text?.trim()).toBe(Array.from({ length: 120 }, (_, i) => `token${i}`).join(' '));
    expect(result.imageX).toBeCloseTo(0, 1);
    expect(result.imageWidth * 0.75).toBeCloseTo(72, 1);
    expect(result.rows.length).toBeGreaterThan(3);
    if (tight) expect(result.rows[1]).toBeLessThan(result.rows[0] - 5);
    else expect(result.rows[1]).toBeCloseTo(result.rows[0], 1);
    expect(result.rows[result.rows.length - 1]).toBeCloseTo(0, 1);
  });
}

import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { textboxTabWidthDocx } from './docx-textbox-tab-width-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });
for (const drawing of [true, false]) for (const long of [true, false]) {
  test(`${drawing ? 'DrawingML' : 'VML'} textbox ${long ? 'wrapped' : 'fitting'} tab prefix preserves every word`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(textboxTabWidthDocx(drawing, long), { browser,
      outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const text = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text.match(/Willow\d+/g)).toEqual(Array.from({ length: long ? 45 : 2 }, (_, index) => `Willow${String(index).padStart(2, '0')}`));
      expect(text.match(/RegisterClosed/g)).toHaveLength(1);
    } finally { await task.destroy(); }
    expect(result.pageMap.pages).toHaveLength(1);
    await page.setContent(result.html!);
    const frames = await page.locator('.page-box span').evaluateAll(spans => spans.filter(span =>
      getComputedStyle(span).display === 'inline-block' && span.getBoundingClientRect().height > 400).map(box => ({
      width: box.getBoundingClientRect().width * .75, height: box.getBoundingClientRect().height * .75,
      horizontalOverflow: box.scrollWidth - box.clientWidth, verticalOverflow: box.scrollHeight - box.clientHeight,
      alignedSegments: Array.from(box.querySelectorAll('span')).filter(span => getComputedStyle(span).display === 'inline-flex').length })));
    expect(frames).toHaveLength(1);
    const geometry = frames[0];
    expect(geometry.width).toBeCloseTo(174, 1); expect(geometry.height).toBeCloseTo(360, 1);
    expect(geometry.horizontalOverflow).toBeLessThanOrEqual(1); expect(geometry.verticalOverflow).toBeLessThanOrEqual(1);
    expect(geometry.alignedSegments > 0).toBe(!long);
  });
}

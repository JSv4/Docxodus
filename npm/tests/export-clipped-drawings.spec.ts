import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateClippedDrawingDocx, type ClippedDrawingOptions } from './docx-clipped-drawings-fixture.js';
import { decodePng } from './visual-parity/png.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const cases: Array<[string, ClippedDrawingOptions]> = [
  ['off-paper header line', { story: 'header', top: 850, line: true }],
  ['off-paper body rectangle', { story: 'body', top: 850 }],
  ['visible header rectangle', { story: 'header', top: 28 }],
  ['straddling header rectangle', { story: 'header', top: 790 }],
];

for (const [name, drawing] of cases) {
  test(`${name} preserves the finalized pages and all visible content`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(generateClippedDrawingDocx(drawing), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(2);
    expect(result.renderReport.unsupportedContent).toEqual([]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(2);
      for (let index = 1; index <= 2; index++) {
        const sheet = await pdf.getPage(index);
        const content = await sheet.getTextContent();
        const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
        expect(text).toContain('Merlin survey ledger.');
        expect(text).toContain(index === 1 ? 'Merlin ridge record.' : 'Merlin river record.');
        if (drawing.top < 800) {
          const operators = await sheet.getOperatorList();
          expect(operators.fnArray.some((operation, i) => operation === OPS.setFillRGBColor
            && operators.argsArray[i][0] === '#53407a')).toBe(true);
        }
      }
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const anchors = page.locator('.page-box [data-docx-drawing-anchor]');
    await expect(anchors).toHaveCount(drawing.story === 'header' ? 2 : 1);
    // The paper clip keeps the complete source drawing and its original position in HTML.
    for (const anchor of await anchors.all()) {
      const top = await anchor.evaluate(element => (element.getBoundingClientRect().top
        - element.closest('.page-box')!.getBoundingClientRect().top) * 0.75);
      expect(top).toBeCloseTo(drawing.top, 1);
    }
    for (const paper of await page.locator('.page-box').all()) {
      const image = decodePng(await paper.screenshot());
      const purpleRows = new Set<number>();
      for (let y = 0; y < image.height; y++) {
        let purple = 0;
        for (let x = 0; x < image.width; x++) {
          const offset = (y * image.width + x) * 4;
          if (image.data[offset] === 83 && image.data[offset + 1] === 64 && image.data[offset + 2] === 122) purple++;
        }
        if (purple > 100) purpleRows.add(y);
      }
      if (drawing.top >= 800) expect(purpleRows.size).toBe(0);
      else expect(Math.abs(purpleRows.size - Math.min(24, 800 - drawing.top) * 4 / 3)).toBeLessThanOrEqual(2);
    }
  });
}

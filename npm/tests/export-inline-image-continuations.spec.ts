import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { inlineImageParagraphDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const labels of [false, true]) {
  test(`an inline image paragraph continues whole images${labels ? ' and labels' : ''}`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(inlineImageParagraphDocx(labels), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const pages: string[] = [];
      let images = 0;
      for (let index = 1; index <= pdf.numPages; index++) {
        const pdfPage = await pdf.getPage(index);
        const content = await pdfPage.getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
        images += (await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject).length;
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(images).toBe(4);
      expect(pages.join(' ').match(/Image inventory complete\./g)).toHaveLength(1);
      if (labels) expect(pages.join(' ').match(/Maple\d/g)).toEqual(['Maple1', 'Maple2', 'Maple3', 'Maple4']);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await page.locator('img').evaluateAll(images => Promise.all(images.map(image => (image as HTMLImageElement).decode())));
    const images = await page.locator('.page-content img').evaluateAll(images => images.map(image => {
      const band = image.closest('.page-content')!.getBoundingClientRect();
      const rect = image.getBoundingClientRect();
      return { width: rect.width * 0.75, height: rect.height * 0.75,
        visible: rect.top >= band.top - 1 && rect.bottom <= band.bottom + 1 };
    }));
    expect(images).toHaveLength(4);
    expect(images.every(image => image.visible)).toBe(true);
    for (let i = 0; i < 4; i++) {
      expect(images[i].width).toBeCloseTo([410, 360, 425, 390][i], 0);
      expect(images[i].height).toBeCloseTo([250, 280, 260, 300][i], 0);
    }
  });
}

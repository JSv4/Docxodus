import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { floatingPictureDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [height, tight] of [[300, false], [920, false], [920, true]] as const) {
test(`a ${tight ? 'tight' : 'square'} floating picture ${height}pt tall reserves its visible height and fits a page`, async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(floatingPictureDocx(height, tight), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages: string[] = [];
    let images = 0;
    for (let index = 1; index <= pdf.numPages; index++) {
      const pdfPage = await pdf.getPage(index), content = await pdfPage.getTextContent();
      pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      images += (await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject).length;
    }
    expect(images).toBe(1);
    expect(pages.join(' ').match(/Hazel\d+/g)).toEqual(Array.from({ length: 26 }, (_, index) => `Hazel${String(index + 1).padStart(2, '0')}`));
    expect(pages.join(' ').match(/Hazel anchor\./g)).toHaveLength(1);
    expect(pages.join(' ').match(/Float inventory complete\./g)).toHaveLength(1);
    expect(pages.findIndex(text => text.includes('Hazel anchor.'))).toBe(1);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const picture = await page.locator('.page-content img').evaluate(image => {
    const rect = image.getBoundingClientRect(), body = image.closest('.page-content')!.getBoundingClientRect();
    return { width: rect.width * 0.75, height: rect.height * 0.75, overflow: rect.bottom - body.bottom,
      frameHeight: image.closest('[data-docx-wrap-text]')?.getBoundingClientRect().height };
  });
  if (height === 300) {
    expect(picture.width).toBeCloseTo(240, 1);
    expect(picture.height).toBeCloseTo(300, 1);
  } else {
    expect(picture.height).toBeLessThan(648);
    expect(picture.width / picture.height).toBeCloseTo(240 / height, 3);
  }
  expect(picture.overflow).toBeLessThanOrEqual(1);
  if (tight) expect(picture.frameHeight! * 0.75).toBeCloseTo(picture.height, 1);
});

}

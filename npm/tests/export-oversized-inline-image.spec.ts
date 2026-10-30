import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { oversizedInlineImageDocx, fittingTilePng } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [height, count] of [[640, 1], [870, 1], [640, 2]]) {
test(`a paragraph with ${count} pictures ${height}pt tall preserves its bitmap and fits proportionally when needed`, async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(oversizedInlineImageDocx(height, count), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    let text = '', images = 0;
    for (let index = 1; index <= pdf.numPages; index++) {
      const pdfPage = await pdf.getPage(index);
      const content = await pdfPage.getTextContent();
      text += content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      images += (await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject).length;
    }
    expect(images).toBe(count);
    expect(text.match(/Before the tall tile\./g)).toHaveLength(1);
    expect(text.match(/After the tall tile\./g)).toHaveLength(1);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const images = await page.locator('.page-content img').evaluateAll(images => images.map(image => {
    const rect = image.getBoundingClientRect(), body = image.closest('.page-content')!.getBoundingClientRect();
    return { source: (image as HTMLImageElement).src, width: rect.width, height: rect.height,
      fits: rect.top >= body.top - 1 && rect.bottom <= body.bottom + 1 && rect.left >= body.left - 1 && rect.right <= body.right + 1 };
  }));
  expect(images).toHaveLength(count);
  for (const image of images) {
  expect(image.source).toBe(`data:image/png;base64,${fittingTilePng().toString('base64')}`);
  expect(image.fits).toBe(true);
  expect(image.height * 0.75).toBeLessThan(648);
  expect(image.width / image.height).toBeCloseTo(320 / height, 3);
  if (height === 640) expect(image.height * 0.75).toBeCloseTo(height, 1);
  }
});
}

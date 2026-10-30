import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { wideBodyTableDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

test('an over-wide asymmetric table fits its text and images on every continuation', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(wideBodyTableDocx(), { browser, outputs: ['html', 'pdf'],
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
    expect(pages.join(' ').match(/Ochre\d+[A-E]/g)).toEqual(Array.from({ length: 39 }, (_, row) =>
      Array.from({ length: 5 }, (_, column) => `Ochre${String(row + 1).padStart(2, '0')}${String.fromCharCode(65 + column)}`)).flat());
    expect(pages.join(' ').match(/Batch complete\./g)).toHaveLength(1);
    expect(images).toBe(5);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const geometry = await page.locator('.page-content').evaluateAll(bodies => bodies.flatMap(body =>
    Array.from(body.querySelectorAll('table'), table => ({
      rightOverflow: table.getBoundingClientRect().right - body.getBoundingClientRect().right,
      bottomOverflow: table.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom,
      fonts: Array.from(table.querySelectorAll('p'), p => getComputedStyle(p).fontSize),
    }))));
  expect(geometry.every(table => table.rightOverflow <= 1 && table.bottomOverflow <= 1)).toBe(true);
  expect(geometry.flatMap(table => table.fonts).every(font => parseFloat(font) >= 14)).toBe(true);
});

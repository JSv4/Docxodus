import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { citingBodyTableDocx } from './docx-citing-body-table-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const citingRow of [1, 43]) {
  test(`an oversized table citing a note in row ${citingRow} preserves rows and note placement`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(citingBodyTableDocx(citingRow), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const pages: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(pages.join(' ').match(/Kestrel\d+/g)).toEqual(Array.from({ length: 57 }, (_, index) =>
        `Kestrel${String(index + 1).padStart(3, '0')}`));
      expect(pages.join(' ').match(/Calibration verified\./g)).toHaveLength(1);
      expect(pages.join(' ').match(/Sampling complete\./g)).toHaveLength(1);
      const citationPage = pages.findIndex(text => text.includes(`Kestrel${String(citingRow).padStart(3, '0')}`));
      expect(pages.findIndex(text => text.includes('Calibration verified.'))).toBe(citationPage);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const overflow = await page.locator('.page-content, .page-footnotes').evaluateAll(bands => bands.flatMap(band =>
      Array.from(band.querySelectorAll('p'), p => p.getBoundingClientRect().bottom - band.getBoundingClientRect().bottom)));
    expect(overflow.every(height => height <= 1)).toBe(true);
  });
}

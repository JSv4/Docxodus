import { expect, test } from '@playwright/test';
import { PDFDocument } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generatePageGeometryDocx } from './docx-page-geometry-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const [name, sizes] of Object.entries({
  integral: [[595, 842]], portrait: [[595, 842.5]], landscape: [[842.5, 595]],
  fractionalWidth: [[602.5, 852]], mixed: [[595, 842.5], [842.5, 595], [612, 792]],
})) {
  test(`${name} pages retain requested layout dimensions and bounded physical PDF geometry`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const source = generatePageGeometryDocx(sizes as [number, number][]);
    const result = await renderDocxArtifacts(source, {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(sizes.length);
    expect(result.pageMap.pages.map(p => [p.width, p.height])).toEqual(sizes);
    expect(result.renderReport.pages.map(p => [p.width, p.height])).toEqual(sizes);
    const pdf = await PDFDocument.load(result.pdf!);
    for (const [index, page] of pdf.getPages().entries()) {
      for (const box of [page.getMediaBox(), page.getCropBox()]) {
        expect(box.x).toBe(0);
        expect(box.y).toBe(0);
        expect(Math.abs(box.width - sizes[index][0])).toBeLessThanOrEqual(0.75);
        expect(Math.abs(box.height - sizes[index][1])).toBeLessThanOrEqual(0.75);
      }
    }
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const textPdf = await task.promise;
      for (let i = 1; i <= sizes.length; i++) {
        const content = await (await textPdf.getPage(i)).getTextContent();
        expect(content.items.filter(item => 'str' in item).map(item => item.str).join(''))
          .toBe(`Geometry${i} Willow archive page.`);
      }
    } finally { await task.destroy(); }
    writeFileSync(testInfo.outputPath(`${name}.docx`), source);
    writeFileSync(testInfo.outputPath(`${name}.pdf`), result.pdf!);
  });
}

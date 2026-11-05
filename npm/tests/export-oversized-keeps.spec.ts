import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { oversizedKeepDocx } from './docx-oversized-keeps-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const keep of ['keepLines', 'widowControl'] as const) {
  test(`${keep} preserves every inventory line when a paragraph exceeds a page`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(oversizedKeepDocx(keep), { browser, outputs: ['pdf'],
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
      expect(pages.join(' ').match(/Copper\d+/g)).toEqual(Array.from({ length: 73 }, (_, index) =>
        `Copper${String(index + 1).padStart(3, '0')}`));
      expect(pages.join(' ').match(/Final receipt signed\./g)).toHaveLength(1);
      expect(pages.every(page => (page.match(/Copper\d+/g)?.length ?? 0) >= 2)).toBe(true);
    } finally { await task.destroy(); }
  });
}

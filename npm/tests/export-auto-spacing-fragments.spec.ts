import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { autoSpacingParagraphDocx } from './docx-auto-spacing-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const note of [false, true]) {
  test(`automatic line spacing continues formatted lines${note ? ' with their note' : ''}`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(autoSpacingParagraphDocx(note), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, pages: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(pages.join(' ').match(/Juniper\d+/g)).toEqual(Array.from({ length: 83 }, (_, index) => `Juniper${String(index + 1).padStart(3, '0')}`));
      expect(pages.join(' ').match(/Checkpoint inventory complete\./g)).toHaveLength(1);
      if (note) {
        expect(pages.join(' ').match(/Juniper note checked\./g)).toHaveLength(1);
        expect(pages.findIndex(text => text.includes('Juniper note checked.')))
          .toBe(pages.findIndex(text => text.includes('Juniper036')));
      }
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const geometry = await page.locator('.page-content').evaluateAll(bodies => bodies.flatMap(body =>
      Array.from(body.querySelectorAll('p'), paragraph => paragraph.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom)));
    expect(geometry.every(overflow => overflow <= 1)).toBe(true);
  });
}

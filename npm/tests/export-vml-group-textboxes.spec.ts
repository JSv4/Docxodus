import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { vmlGroupDocx } from './docx-vml-group-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const header of [false, true]) {
  for (const kind of ['group', 'nested', 'separate'] as const) {
    test(`${header ? 'header' : 'body'} ${kind} textboxes retain both PDF stories and independent bounds`, async ({ browser, page }) => {
      test.setTimeout(120_000);
      const result = await renderDocxArtifacts(vmlGroupDocx(header, kind), {
        browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
      });
      expect(result.pageCount).toBe(1);
      await page.setContent(result.html!);
      const boxes = await page.locator('.page-box [data-docx-textbox]').evaluateAll(elements => elements.map(box => {
        const b = box.getBoundingClientRect(), p = box.closest('.page-box')!.getBoundingClientRect();
        return { text: box.textContent?.trim(), x: (b.x-p.x)*.75, y: (b.y-p.y)*.75,
          width: b.width*.75, height: b.height*.75 };
      }));
      expect(boxes.map(box => box.text)).toEqual(['ALDER BOX NORTH', 'BIRCH BOX SOUTH']);
      for (const [index, box] of boxes.entries()) {
        expect(box.x).toBeCloseTo(72 + index*150, 1);
        expect(box.y).toBeCloseTo(100 + index*50, 1);
        expect(box.width).toBeCloseTo(120, 1);
        expect(box.height).toBeCloseTo(28, 1);
      }
      const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
      try {
        const content = await (await (await task.promise).getPage(1)).getTextContent();
        const items = content.items.filter(item => 'str' in item);
        for (const [index, label] of ['ALDER BOX NORTH', 'BIRCH BOX SOUTH'].entries()) {
          const matches = items.filter(item => item.str === label);
          expect(matches).toHaveLength(1);
          expect(matches[0].transform[4]).toBeCloseTo(72 + index*150, 1);
          expect(792-matches[0].transform[5]).toBeGreaterThan(100 + index*50);
          expect(792-matches[0].transform[5]).toBeLessThan(114 + index*50);
        }
      } finally { await task.destroy(); }
    });
  }
}

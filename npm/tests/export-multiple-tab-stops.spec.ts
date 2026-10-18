import { test, expect } from '@playwright/test';
import { writeFileSync } from 'node:fs';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { multipleTabFixture, tabEntries } from './docx-multiple-tab-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('delivered PDF aligns every title and page number to its declared tab stop', async ({ browser }, info) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(multipleTabFixture(), { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const items = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item);
    const at = (text: string) => {
      const item = items.find(item => item.str.trim() === text);
      expect(item, `PDF text ${text}`).toBeDefined(); return item!;
    };
    for (const [, title, number] of tabEntries) {
      const titleItem = at(title), numberItem = at(number);
      expect(titleItem.transform[4]).toBeCloseTo(162, 0);
      expect(numberItem.transform[4] + numberItem.width).toBeCloseTo(522, 0);
      expect(numberItem.transform[5]).toBeCloseTo(titleItem.transform[5], 0);
    }
    writeFileSync(info.outputPath('multiple-tab-stops.pdf'), result.pdf!);
  } finally { await task.destroy(); }
});

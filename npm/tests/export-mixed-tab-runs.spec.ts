import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { mixedTabRows, mixedTabRunDocx } from './docx-mixed-tab-run-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const story of ['body', 'drawing', 'vml'] as const) for (const partition of ['split', 'mixed', 'single'] as const) {
  test(`${story} tab columns preserve every line with ${partition} runs`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const input = mixedTabRunDocx(story, partition);
    writeFileSync(testInfo.outputPath('kestrel-ledger.docx'), input);
    const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const lines = await page.evaluate(tokens => tokens.map(row => row.map(token => {
      const walk = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      let node: Node | null;
      while ((node = walk.nextNode())) if (node.textContent === token) {
        const range = document.createRange(); range.selectNodeContents(node);
        return Array.from(range.getClientRects(), r => ({ x: r.x, y: r.y, width: r.width, height: r.height }));
      }
      return [];
    })), mixedTabRows);
    for (const row of lines) {
      for (const rects of row) expect(rects).toHaveLength(1);
      expect(Math.abs(row[0][0].y - row[1][0].y)).toBeLessThan(.02);
      expect(Math.abs(row[0][0].y - row[2][0].y)).toBeLessThan(.02);
      expect(row[1][0].x - row[0][0].x).toBeCloseTo(140 * 4 / 3, 1);
      expect(row[2][0].x - row[0][0].x).toBeCloseTo(300 * 4 / 3, 1);
    }
    const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const items = (await (await (await pdf.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const text = items.map(item => item.str).join(' ');
      for (const token of [...mixedTabRows.flat(), 'Kestrel ledger complete']) expect(text).toContain(token);
      for (const row of mixedTabRows) {
        const columns = row.map(token => items.find(item => item.str === token)!);
        expect(columns.every(Boolean)).toBe(true);
        expect(Math.abs(columns[0].transform[5] - columns[1].transform[5])).toBeLessThan(.01);
        expect(Math.abs(columns[0].transform[5] - columns[2].transform[5])).toBeLessThan(.01);
        expect(Math.abs(columns[1].transform[4] - columns[0].transform[4] - 140)).toBeLessThan(.8);
        expect(Math.abs(columns[2].transform[4] - columns[0].transform[4] - 300)).toBeLessThan(.8);
      }
    } finally { await pdf.destroy(); }
  });
}

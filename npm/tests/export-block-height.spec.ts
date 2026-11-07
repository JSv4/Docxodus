import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { blockHeightDocx } from './docx-block-height-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function pdfWords(bytes: Uint8Array): Promise<Array<{ text: string; y: number; page: number }>> {
  const loading = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await loading.promise, words = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n);
      for (const item of (await page.getTextContent()).items)
        if ('str' in item && item.str.trim()) words.push({ text: item.str.trim(), y: page.view[3] - item.transform[5], page: n });
    }
    return words;
  } finally { await loading.destroy(); }
}

for (const [multiple, mark, imageHeight] of [[1, 12, 100], [1, 8, 100], [1, 20, 100], [1.15, 12, 100], [2, 12, 100], [1, 12, 8], [2, 12, 8]]) {
  test(`picture-only paragraph: ${multiple} spacing, ${mark}pt mark, ${imageHeight}pt picture`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(blockHeightDocx({ kind: 'image', multiple, mark, imageHeight }), {
      browser, outputs: ['html', 'pdf'], fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
      reviewProfile: 'final', commentProfile: 'hidden',
    });
    const words = await pdfWords(result.pdf!);
    expect(words.map(word => word.text)).toEqual(['PineAbove', 'PineBelow']);
    expect(Math.abs(words[1].y - words[0].y - (Math.max(imageHeight, 13.8) + 13.8 * multiple))).toBeLessThan(.5);
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const bounds = await page.locator('.page-content p').filter({ has: page.locator('img') })
      .evaluate(p => ({ height: p.getBoundingClientRect().height * .75,
        pictureTop: (p.querySelector('img')!.getBoundingClientRect().top - p.getBoundingClientRect().top) * .75 }));
    expect(Math.abs(bounds.height - (Math.max(imageHeight, 13.8) + 13.8 * (multiple - 1)))).toBeLessThan(.05);
    expect(Math.abs(bounds.pictureTop - Math.max(13.8 - imageHeight, 0))).toBeLessThan(.05);
  });
}

test('a picture sharing its line with text retains the text descent', async ({ browser }) => {
  const result = await renderDocxArtifacts(blockHeightDocx({ kind: 'image', mixed: true }), {
    browser, outputs: ['pdf'], fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
    reviewProfile: 'final', commentProfile: 'hidden',
  });
  const words = await pdfWords(result.pdf!);
  expect(words.map(word => word.text)).toEqual(['PineAbove', 'PineCaption', 'PineBelow']);
  expect(Math.abs(words[2].y - words[0].y - 117)).toBeLessThan(.5);
});

for (const [top, bottom, rule, borders, lines, rows, expectedHeight] of [
  [0, 0, 'atLeast', 0, 1, 5, 30],
  [6, 0, 'atLeast', 0, 1, 5, 36],
  [6, 6, 'atLeast', 0, 1, 5, 42],
  [6, 6, 'atLeast', 2, 1, 5, 44],
  [6, 6, 'atLeast', 0, 3, 5, 53.4],
  [6, 6, 'exact', 0, 1, 5, 30],
  [6, 6, 'auto', 0, 1, 5, 25.8],
  [6, 6, 'atLeast', 2, 1, 22, 44],
] as const) {
  test(`${rule} rows with ${top}/${bottom}pt margins, ${borders}pt borders, ${lines} lines, ${rows} rows`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(blockHeightDocx({ kind: 'row', top, bottom, rule, borders, lines, rows }), {
      browser, outputs: ['html', 'pdf'], fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
      reviewProfile: 'final', commentProfile: 'hidden',
    });
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const heights = await page.locator('.page-content tr').evaluateAll(rows => rows.map(row => row.getBoundingClientRect().height * .75));
    expect(heights).toHaveLength(rows);
    for (const height of heights) expect(Math.abs(height - expectedHeight)).toBeLessThan(.1);
    expect(result.pageCount).toBe(rows > 20 ? 2 : 1);
    const words = await pdfWords(result.pdf!);
    const expected = Array.from({ length: rows }, (_, row) => ['A', 'B'].flatMap(column =>
      Array.from({ length: lines }, (_, line) => `Pine${row + 1}${column}${line + 1}`))).flat().concat('PineAfter');
    expect(words.map(word => word.text).sort()).toEqual(expected.sort());
    const starts = words.filter(word => /^Pine\d+A1$/.test(word.text));
    for (let i = 1; i < starts.length; i++) if (starts[i].page === starts[i - 1].page)
      expect(Math.abs(starts[i].y - starts[i - 1].y - expectedHeight)).toBeLessThan(.8);
  });
}

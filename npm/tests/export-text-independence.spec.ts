import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { fitTexts, textFitDocx, type FitNeighbor } from './docx-text-fit-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const neighbor of ['plain', 'bold', 'table', 'list', 'link'] as FitNeighbor[])
  for (const alignment of ['justify', 'left'] as const)
    check(alignment, neighbor, undefined, alignment === 'justify' ? 'pier as' : 'Where');
for (const owner of ['run', 'defaults', 'style'] as const)
  for (const halfPoints of [16, 28]) check('left', 'bold', { owner, halfPoints }, halfPoints === 16 ? 'Where Water' : 'Where');

function check(alignment: keyof typeof fitTexts, neighbor: FitNeighbor,
  kern: Parameters<typeof textFitDocx>[2], ending: string): void {
  test(`${alignment} PDF with ${neighbor}, kerning ${JSON.stringify(kern)}`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(textFitDocx(alignment, neighbor, kern), { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const content = await (await pdf.getPage(1)).getTextContent();
      const rows = new Map<number, Array<{ text: string; x: number }>>();
      for (const item of content.items) if ('str' in item && item.str.trim()) {
        const y = Math.round(item.transform[5] * 100) / 100;
        const row = rows.get(y) ?? []; row.push({ text: item.str.trim(), x: item.transform[4] }); rows.set(y, row);
      }
      // PDF paint order can put list markers first; line ownership is geometric.
      const lines = Array.from(rows).sort(([a], [b]) => b - a)
        .map(([, row]) => row.sort((a, b) => a.x - b.x).map(item => item.text).join(' '));
      expect(lines[0]).toMatch(new RegExp(ending + '$'));
      expect(lines.slice(0, 2).join(' ')).toBe(fitTexts[alignment]);
    } finally { await task.destroy(); }
  });
}

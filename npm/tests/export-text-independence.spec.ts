import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { fitText, fitTexts, textFitDocx, type FitNeighbor, type FitTarget } from './docx-text-fit-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const neighbor of ['plain', 'bold', 'table', 'list', 'link'] as FitNeighbor[])
  for (const alignment of ['justify', 'left'] as const)
    check(alignment, neighbor, undefined, alignment === 'justify' ? 'pier as' : 'Where');
for (const owner of ['run', 'defaults', 'style'] as const)
  for (const halfPoints of [16, 28]) check('left', 'bold', { owner, halfPoints }, halfPoints === 16 ? 'Where Water' : 'Where');
// Word kerns 12pt text from a w:kern minimum of 12pt (24), not 12.5pt or 0, and only on request,
// whether the paragraph holds one run or several (#1562).
for (const [halfPoints, ending] of [[24, 'Where Water'], [25, 'Where'], [0, 'Where']] as const)
  check('left', 'bold', { owner: 'run', halfPoints }, ending);
check('left', 'plain', undefined, 'Where', 'bold-last');
check('left', 'plain', { owner: 'run', halfPoints: 2 }, 'Where Water', 'bold-last');
check('justify', 'plain', undefined, 'pier as', 'bold-last');
// Where Word ends line 1 of the justified paragraph in each placement (#1550), beside the
// left-aligned and compatibility-mode-14 controls that keep natural spaces.
for (const [target, justified, natural] of [
  ['unspaced', 'pier as', 'pier'],
  ['bold-the', 'pier as', 'pier'],
  ['bold', 'pier', 'pier'],
  ['link', 'pier as', 'pier'],
  ['list', 'northern', 'our'],
  ['cell', 'pier', 'northern'],
] as Array<[FitTarget, string, string]>) {
  check('justify', 'plain', undefined, justified, target);
  check('left', 'plain', undefined, natural, target);
  check('justify', 'plain', undefined, natural, target, 14);
}

function check(alignment: keyof typeof fitTexts, neighbor: FitNeighbor,
  kern: Parameters<typeof textFitDocx>[2], ending: string, target: FitTarget = 'paragraph', mode = 15): void {
  const name = target === 'paragraph' ? `with ${neighbor}, kerning ${JSON.stringify(kern)}`
    : `in a ${target} target, mode ${mode}${kern ? `, kerning ${JSON.stringify(kern)}` : ''}`;
  test(`${alignment} PDF ${name}`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(textFitDocx(alignment, neighbor, kern, target, mode), { browser,
      outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const content = await (await pdf.getPage(1)).getTextContent();
      // PDF paint order can put list markers first, and a marker or styled run can sit a little
      // off the text baseline; line ownership is geometric.
      const items = content.items.flatMap(item => 'str' in item && item.str.trim()
        ? [{ text: item.str.trim(), x: item.transform[4], y: item.transform[5] }] : [])
        .sort((a, b) => b.y - a.y);
      const rows: Array<typeof items> = [];
      for (const item of items) {
        if (!rows.length || rows.at(-1)![0].y - item.y > 2) rows.push([]);
        rows.at(-1)!.push(item);
      }
      const lines = rows.map(row => row.sort((a, b) => a.x - b.x).map(item => item.text).join(' ').replace(/^1\. /, ''));
      expect(lines[0]).toMatch(new RegExp(ending + '$'));
      // Runs split a word into separate PDF text items; compare the characters.
      expect(lines.slice(0, 2).join('').replace(/\s/g, '')).toBe(fitText(alignment, target).replace(/\s/g, ''));
    } finally { await task.destroy(); }
  });
}

import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { gridTableDocx, type GridTableRow } from './docx-table-grid-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

/**
 * Issue #1750: grid columns that no row reaches are not part of the table. The reference layout drops
 * them and widens the columns the rows use, in proportion, until the table spans its preferred width;
 * a column any row reaches stays. Each case gives the cell widths in points the reference lays out
 * (docs/ooxml_corner_cases.md), checked as the distances between the starts of adjacent cells' text.
 */
const issueGrid = [1800, 2200, 1500, 1500, 1500, 3000, 3200];
const issueRow: GridTableRow = { cells: [[1800], [2200], [4500, 3]], props: '<w:gridAfter w:val="2"/>' };
const grid = [1000, 2500, 1500, 2000, 2000];
const short: GridTableRow = { cells: [[1000], [2500], [1500]], props: '<w:gridAfter w:val="2"/>' };
const full: GridTableRow = { cells: [[1000], [2500], [1500], [2000], [2000]] };
const four: GridTableRow = { cells: [[1000], [2500], [1500], [2000]], props: '<w:gridAfter w:val="1"/>' };

const cases: Array<[string, number[], GridTableRow[], number, Record<string, number[]>]> = [
  ['the issue: every row leaves two of seven columns', issueGrid, [issueRow, issueRow, issueRow], 14700,
    { A: [155.65, 190.25] }],
  ['every row leaves the last two columns', grid, [short, short, short], 9000, { A: [90, 225] }],
  ['a column only one row reaches stays', grid, [short, four, short], 9000, { A: [64.3, 160.7], B: [64.3, 160.7, 96.45] }],
  ['a row that fills the grid keeps every column (control)', grid, [short, short, full], 9000,
    { A: [50, 125], C: [50, 125, 75, 100] }],
];

for (const [name, columns, rows, width, expected] of cases) {
  test(`unused grid columns: ${name}`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(gridTableDocx(columns, rows, width), { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const page = await (await task.promise).getPage(1);
      const starts = Object.fromEntries((await page.getTextContent()).items
        .filter(item => 'str' in item && /^[A-C]\d$/.test(item.str)).map(item => 'str' in item ? [item.str, item.transform[4]] : []));
      for (const [row, widths] of Object.entries(expected)) {
        widths.forEach((cellWidth, index) =>
          expect(starts[`${row}${index + 2}`] - starts[`${row}${index + 1}`], `${row}${index + 1}`).toBeCloseTo(cellWidth, 0));
      }
    } finally { await task.destroy(); }
  });
}

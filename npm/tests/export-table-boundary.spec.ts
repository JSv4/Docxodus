import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { tableBoundaryDocx, type TableBoundaryOptions } from './docx-table-boundary-fixture.js';

/**
 * Word starts a table's first row where the block above it ends (issue #1244): a paragraph's own
 * after-spacing is the only gap, and an adjacent table adds none. Lines are exact 16pt and rows at
 * least 18pt, so the distance between two baselines is known in advance; the gaps below are Word for
 * the web's (docs/ooxml_corner_cases.md). Before the fix every table here but the control sat 7.5pt lower.
 */

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function baselines(bytes: Uint8Array): Promise<Record<string, number>> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const page = await (await task.promise).getPage(1);
    const items = (await page.getTextContent()).items.filter(item => 'str' in item);
    return Object.fromEntries(items.filter(item => /^Boundary/.test(item.str)).map(item => [item.str, item.transform[5]]));
  } finally { await task.destroy(); }
}

const cases: [string, TableBoundaryOptions, Record<string, number>][] = [
  ['no after-spacing', { afterTwips: null }, { BoundaryCell: 16, BoundaryAfter: 18 }],
  ['zero after-spacing', { afterTwips: '0' }, { BoundaryCell: 16, BoundaryAfter: 18 }],
  ['zero after-spacing and borders', { afterTwips: '0', borders: true }, { BoundaryCell: 17, BoundaryAfter: 19 }],
  ['12pt after-spacing (control)', { afterTwips: '240' }, { BoundaryCell: 28, BoundaryAfter: 18 }],
  ['an adjacent table', { afterTwips: '0', adjacent: true }, { BoundaryCell: 16, BoundarySecond: 18, BoundaryAfter: 18 }],
];

for (const [name, options, gaps] of cases) {
  test(`a table after a paragraph with ${name} starts where Word starts it`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = tableBoundaryDocx(options);
    const actual = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], reviewProfile: 'final',
      commentProfile: 'hidden', timeoutMs: 120_000 });
    writeFileSync(testInfo.outputPath('boundary.docx'), bytes);
    writeFileSync(testInfo.outputPath('boundary.pdf'), actual.pdf!);
    const y = await baselines(actual.pdf!);
    let previous = y.BoundaryBefore;
    for (const [marker, gap] of Object.entries(gaps)) {
      // Within 0.75pt: a collapsed CSS border lays out half its width above the cell, where Word puts
      // all of it, so a bordered table's first line sits 0.5pt higher than Word's.
      expect(Math.abs(previous - y[marker] - gap), `${marker} sits ${gap}pt below the line above it`)
        .toBeLessThanOrEqual(0.75);
      previous = y[marker];
    }
  });
}

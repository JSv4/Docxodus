import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { cellMargins108, tableIndentDocx, type TableIndentOptions } from './docx-table-indent-fixture.js';

/**
 * Where the first cell's text starts, measured from the body text edge, against Word for the web's
 * PDF of the same table (issue #1219; table in docs/ooxml_corner_cases.md). A cell with no margin
 * stated anywhere has none; a table that names no style takes the default table style's; before
 * compatibility mode 15 a styled or indented table puts its first cell's text at the indent.
 */

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function lineStarts(bytes: Uint8Array): Promise<Record<string, number>> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const page = await (await task.promise).getPage(1);
    const items = (await page.getTextContent()).items.filter(item => 'str' in item);
    return Object.fromEntries(items.filter(item => /^(Indent|Outer|Inner)/.test(item.str)).map(item => [item.str, item.transform[4]]));
  } finally { await task.destroy(); }
}

const indent0 = '<w:tblInd w:w="0" w:type="dxa"/>';
const cases: [string, TableIndentOptions, number][] = [
  ['mode 15, no cell margin stated anywhere', { mode: '15' }, 1],
  ['mode 15, 108-twip margins (control)', { mode: '15', tableProps: cellMargins108 }, 6],
  ['mode 15, inheriting the default table style', { mode: '15', defaultTableStyle: true }, 6],
  ['mode 15, a 720-twip indent', { mode: '15', tableProps: '<w:tblInd w:w="720" w:type="dxa"/>' + cellMargins108 }, 42],
  ['mode 14, a stated zero indent', { mode: '14', tableProps: indent0 + cellMargins108 }, 0],
  ['mode 14, inheriting the default table style', { mode: '14', defaultTableStyle: true }, 0],
  ['mode 14, no indent and no style', { mode: '14', tableProps: cellMargins108 }, 5.5],
  ['no stated mode, a stated zero indent', { mode: null, tableProps: indent0 + cellMargins108 }, 0],
];

for (const [name, options, wordOffset] of cases) {
  test(`table text in ${name} starts where Word starts it`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = tableIndentDocx(options);
    const actual = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], reviewProfile: 'final',
      commentProfile: 'hidden', timeoutMs: 120_000 });
    writeFileSync(testInfo.outputPath('indent.docx'), bytes);
    writeFileSync(testInfo.outputPath('indent.pdf'), actual.pdf!);
    const x = await lineStarts(actual.pdf!);
    expect(x.IndentAfter).toBeCloseTo(x.IndentBefore, 1);
    expect(x.IndentCell - x.IndentBefore).toBeCloseTo(wordOffset, 0);
  });
}


const wordFixtureRoot = new URL('../../TestFiles/TableHorizontalPlacement/', import.meta.url);
const wordCases = JSON.parse(readFileSync(new URL('word-observations.json', wordFixtureRoot), 'utf8')) as {
  cases: Array<{ id: string; inputSha256: string; wordMarkers: Record<string, number[]> }>;
};
for (const observed of wordCases.cases) {
  test(`${observed.id}: nested table margins match the captured Word origins`, async ({ browser }) => {
    test.setTimeout(120_000);
    const bytes = new Uint8Array(readFileSync(fileURLToPath(new URL(observed.id + '.docx', wordFixtureRoot))));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(observed.inputSha256);
    const actual = await renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], reviewProfile: 'final',
      commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(actual.pageCount).toBe(1);
    const starts = await lineStarts(actual.pdf!);
    for (const marker of ['OuterLeft', 'InnerLeft'])
      expect(starts[marker]).toBeCloseTo(observed.wordMarkers[marker][0], 0);
  });
}

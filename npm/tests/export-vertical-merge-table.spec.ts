import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateVerticalMergeDocx, type VerticalMergeOptions } from './docx-vertical-merge-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

const cases: Array<{ name: string; options: VerticalMergeOptions; pages: number }> = [
  { name: 'unmerged control', options: { merged: false }, pages: 2 },
  { name: 'merge before the page boundary', options: {}, pages: 2 },
  { name: 'merge across the preferred boundary', options: { mergeStart: 10, mergeSpan: 3 }, pages: 2 },
  { name: 'merge at the next page start', options: { mergeStart: 11, mergeSpan: 2 }, pages: 2 },
  { name: 'combined vertical and horizontal merge', options: { mergeStart: 10, mergeSpan: 2, colSpan: 2 }, pages: 2 },
  { name: 'merge in a later fragment', options: { mergeStart: 20, mergeSpan: 3, rowCount: 28 }, pages: 3 },
  { name: 'overlapping merges in different columns', options: { mergeStart: 9, mergeSpan: 3, rightMerge: { start: 11, span: 3 } }, pages: 2 },
];

for (const { name, options, pages } of cases) {
  test(`${name} exports intact rows and merged cells without clipping`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateVerticalMergeDocx(options);
    const result = await renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000 });
    expect(result.pageCount).toBe(pages);
    const tableFragments = result.pageMap.fragments.filter(f => f.anchorId.startsWith('tbl:'));
    expect(tableFragments).toHaveLength(pages);
    expect(new Set(tableFragments.map(f => f.anchorId)).size).toBe(1);
    expect(new Set(tableFragments.map(f => f.fragmentId)).size).toBe(pages);
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const topology = await page.locator('.page-content table').evaluateAll((tables, overlap) => tables.map(element => {
      const table = element as HTMLTableElement;
      const band = table.closest('.page-content')!.getBoundingClientRect();
      const rect = table.getBoundingClientRect();
      return {
        rows: Array.from(table.rows).map(row => row.textContent!.match(overlap ? /Row\d+/ : /Right\d+/)![0]),
        spans: Array.from(table.rows).flatMap((row, index) => Array.from(row.cells)
          .filter(cell => cell.rowSpan > 1).map(cell => ({ start: index, span: cell.rowSpan, colSpan: cell.colSpan }))),
        count: table.rows.length,
        fits: rect.top >= band.top - 1 && rect.bottom <= band.bottom + 1,
        colgroup: !!table.querySelector('colgroup'),
      };
    }), !!options.rightMerge);
    const rowCount = options.rowCount ?? 18;
    const allRows = Array.from({ length: rowCount }, (_, index) => index + 1);
    const expectedRight = allRows.filter(row => !options.rightMerge || row <= options.rightMerge.start
      || row >= options.rightMerge.start + options.rightMerge.span).map(row => `Right${String(row).padStart(2, '0')}`);
    expect(topology.flatMap(t => t.rows)).toEqual(options.rightMerge
      ? allRows.map(row => `Row${String(row).padStart(2, '0')}`) : expectedRight);
    expect(topology.every(t => t.fits && t.colgroup)).toBe(true);
    const spans = topology.flatMap(t => t.spans);
    expect(spans).toHaveLength(options.merged === false ? 0 : options.rightMerge ? 2 : 1);
    if (options.merged !== false) {
      expect(spans[0]).toMatchObject({ span: options.mergeSpan ?? 2, colSpan: options.colSpan ?? 1 });
      expect(topology.every(t => t.spans.every(s => s.start + s.span <= t.count))).toBe(true);
    }
    expect(await page.locator('.page-content table[data-source-anchor-id]').count()).toBe(pages);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const text: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const physical = await pdf.getPage(index);
        expect(physical.view).toEqual([0, 0, 612, 792]);
        text.push((await physical.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(text.join(' ').match(/Right\d+/g)).toEqual(expectedRight);
      const start = options.mergeStart ?? 1;
      const span = options.mergeSpan ?? 2;
      const expectedLeft = allRows
        .filter(row => row <= start || row >= start + span).map(row => `Left${String(row).padStart(2, '0')}`);
      expect(text.join(' ').match(/Left\d+/g)).toEqual(expectedLeft);
    } finally { await task.destroy(); }
    writeFileSync(testInfo.outputPath(`${name.replace(/ /g, '-')}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`${name.replace(/ /g, '-')}.pdf`), result.pdf!);
  });
}

for (const profile of [
  { reviewProfile: 'markup', commentProfile: 'margin' },
  { reviewProfile: 'final', commentProfile: 'hidden' },
] as const) test(`a merge group taller than a page retains complete text in ${profile.reviewProfile}/${profile.commentProfile}`, async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(generateVerticalMergeDocx({ mergeSpan: 18 }), {
    browser, outputs: ['pdf'], ...profile, limits: { finalPages: 2 }, timeoutMs: 120_000,
  });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise, text: string[] = [];
    expect(pdf.numPages).toBe(2);
    for (let index = 1; index <= pdf.numPages; index++)
      text.push((await (await pdf.getPage(index)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' '));
    const paragraphs = text.join(' ').replace(/\s+/g, ' ').match(/(?:Left|Right)\d{2} (?:Cedar tray|Archive register)\./g);
    expect(paragraphs).toEqual(['Left01 Cedar tray.',
      ...Array.from({ length: 18 }, (_, index) => `Right${String(index + 1).padStart(2, '0')} Archive register.`)]);
  } finally { await task.destroy(); }
});

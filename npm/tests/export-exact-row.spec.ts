import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateExactRowDocx } from './docx-exact-row-fixture.js';

/** Exact-height table rows (`w:hRule="exact"`) in the PDF export: Word draws only what fits the row. */

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

const options = { outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
  commentProfile: 'margin' as const, timeoutMs: 120_000 };

async function pageTexts(bytes: Uint8Array): Promise<string[]> {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages: string[] = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent();
      pages.push(content.items.map(item => ('str' in item ? item.str : '')).join(' '));
    }
    return pages;
  } finally { await task.destroy(); }
}

test('a row that hides a whole paragraph exports, clipped like any other overflow (issue #1168)', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const bytes = generateExactRowDocx({ clippedParagraphs: 3 });
  const actual = await renderDocxArtifacts(bytes, { browser, ...options });
  writeFileSync(testInfo.outputPath('hidden-paragraph.docx'), bytes);
  writeFileSync(testInfo.outputPath('hidden-paragraph.pdf'), actual.pdf!);

  expect(actual.pageCount).toBe(1);
  // The row keeps its exact height: the clipped cell lays out like the cell that only cuts a line.
  const control = await renderDocxArtifacts(generateExactRowDocx({ clippedParagraphs: 2 }), { browser, ...options });
  const rows = (result: typeof actual) => result.pageMap.fragments
    .filter(fragment => fragment.anchorId.startsWith('tr:')).map(fragment => fragment.geometry);
  expect(rows(actual).length).toBeGreaterThan(0);
  expect(rows(actual)).toEqual(rows(control));
  // The text layer keeps the hidden paragraph whole, after the cell's visible text (issue #1140).
  const text = (await pageTexts(actual.pdf!)).join(' ');
  for (const word of ['Alder', 'Birch', 'Cedar', 'Value01', 'Row02', 'Closing']) expect(text).toContain(word);
  expect(text.indexOf('Cedar')).toBeGreaterThan(text.indexOf('Birch'));
});

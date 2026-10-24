import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateExactRowDocx } from './docx-exact-row-fixture.js';
import { decodePng } from './visual-parity/png.js';

/** Positive exact row heights clip; exact zero uses the content height, as Word does. */

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

for (const name of ['band-clip', 'wide-control']) {
  test(`hidden header-row words stay whole under the ${name} band`, async ({ browser }) => {
    test.setTimeout(120_000);
    const path = fileURLToPath(new URL(`../../TestFiles/ExactRowBands/tables-exact-row-header-${name}.docx`, import.meta.url));
    const actual = await renderDocxArtifacts(new Uint8Array(readFileSync(path)), { browser, ...options });
    const pages = await pageTexts(actual.pdf!);
    expect(pages).toHaveLength(1);
    const words = pages[0].trim().split(/\s+/);
    for (const word of ['HdrAlder', 'HdrValue', 'BodyWillow']) expect(words.filter(value => value === word)).toHaveLength(1);
    for (const word of ['HdrBirch', 'HdrCedar']) expect(words.filter(value => value === word).length).toBeLessThanOrEqual(1);
    expect(words.filter(word => !['HdrAlder', 'HdrBirch', 'HdrCedar', 'HdrValue', 'BodyWillow'].includes(word))).toEqual([]);
  });
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

for (const tableBorders of [true, false]) {
  test(`a zero-height exact row ${tableBorders ? 'with' : 'without'} borders exports at its content height (issue #1339)`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    // Native Word treats exact zero as auto. Both paragraphs stay visible and the row
    // contributes their full height, with or without table borders.
    const bytes = generateExactRowDocx({ clippedParagraphs: 2, firstRowTwips: 0, tableBorders });
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    writeFileSync(testInfo.outputPath('zero-height.docx'), bytes);
    writeFileSync(testInfo.outputPath('zero-height.pdf'), actual.pdf!);

    expect(actual.pageCount).toBe(1);
    const text = (await pageTexts(actual.pdf!)).join(' ');
    for (const word of ['Alder', 'Birch', 'Value01', 'Value02', 'Row02', 'Closing']) expect(text).toContain(word);
    const rows = actual.pageMap.fragments.filter(fragment => fragment.anchorId.startsWith('tr:'));
    expect(rows).toHaveLength(3);
    expect(rows[0].geometry.height).toBeGreaterThan(20);
  });
}

test('a table made only of a borderless exact-zero row keeps its content (issue #1339)', async ({ browser }) => {
  test.setTimeout(120_000);
  const actual = await renderDocxArtifacts(
    generateExactRowDocx({ bodyRows: 1, firstRowTwips: 0, tableBorders: false }), { browser, ...options });

  expect(actual.pageCount).toBe(1);
  const text = (await pageTexts(actual.pdf!)).join(' ');
  expect(text).toContain('Closing');
  expect(text).toContain('Alder');
  expect(actual.pageMap.fragments.some(fragment => fragment.anchorId.startsWith('tbl:'))).toBe(true);
});

for (const rule of ['exact', 'atLeast'] as const) {
  test(`${rule}-height header rows repeat at the top of every page the table continues on (issue #1171)`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateExactRowDocx({ headerRule: rule, bodyRows: 70 });
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    writeFileSync(testInfo.outputPath(`header-${rule}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`header-${rule}.pdf`), actual.pdf!);

    const pages = await pageTexts(actual.pdf!);
    expect(pages.length).toBeGreaterThan(1);
    for (const page of pages) expect(page.trimStart()).toMatch(/^Heading Ledger\s+Heading Value\b/);
    // The body rows are not repeated: each is printed once, in order.
    const body = pages.join(' ').match(/Row\d\d ledger/g);
    expect(body).toEqual(Array.from({ length: 69 }, (_, i) => `Row${String(i + 2).padStart(2, '0')} ledger`));
  });
}

test('a paragraph border the exact-height row hides is clipped with the row (issue #1171)', async ({ browser, page }, testInfo) => {
  test.setTimeout(120_000);
  const bytes = generateExactRowDocx({ clippedParagraphs: 2, borders: true });
  const actual = await renderDocxArtifacts(bytes, { browser, ...options, outputs: ['html', 'pdf'] });
  writeFileSync(testInfo.outputPath('borders.docx'), bytes);
  writeFileSync(testInfo.outputPath('borders.pdf'), actual.pdf!);

  // The finalized print document the PDF is printed from, rasterized at one pixel per CSS pixel.
  await page.setViewportSize({ width: 1000, height: 1200 });
  await page.setContent(actual.html!);
  const row = await page.locator('.page-box tr').first().evaluate((tr) => {
    const rect = tr.getBoundingClientRect();
    return { top: rect.top, bottom: rect.bottom };
  });
  const image = decodePng(await page.screenshot({ fullPage: true }));
  // Pixel rows crossed by a red stroke: a long run of red, unlike the colour fringes of anti-aliased text.
  const redRows = new Set<number>();
  for (let y = 0; y < image.height; y++) {
    let run = 0;
    for (let x = 0; x < image.width && run < 100; x++) {
      const at = (y * image.width + x) * 4;
      run = image.data[at] > 180 && image.data[at + 1] < 100 && image.data[at + 2] < 100 ? run + 1 : 0;
    }
    if (run >= 100) redRows.add(y);
  }
  // The group's top border lies inside the row and still shows; its bottom border lies below the
  // row, under the Birch line the row cuts, and must not.
  expect([...redRows].some(y => y >= row.top && y < row.bottom), JSON.stringify({ row, red: [...redRows] })).toBe(true);
  expect([...redRows].filter(y => y < row.top - 1 || y >= row.bottom + 1)).toEqual([]);
});

/** The overlay shapes the export paints for the boxed (green) paragraph, and its exact row's span. */
async function boxShapes(page: import('@playwright/test').Page, html: string) {
  await page.setViewportSize({ width: 1000, height: 1200 });
  await page.setContent(html);
  return page.evaluate(() => {
    const box = document.querySelector('.page-box')!.getBoundingClientRect();
    const row = document.querySelector('.page-box tr')!.getBoundingClientRect();
    const text = document.querySelector('[data-docx-exact-row-cell] p, [data-docx-exact-row-cell]')!.getBoundingClientRect();
    return {
      row: { top: row.top - box.top, bottom: row.bottom - box.top },
      textLeft: text.left - box.left,
      shapes: Array.from(document.querySelectorAll('.page-box > svg > rect'))
        .filter(shape => shape.getAttribute('fill') === 'rgb(0, 170, 0)')
        .map(shape => ({ x: Number(shape.getAttribute('x')), y: Number(shape.getAttribute('y')),
          width: Number(shape.getAttribute('width')), height: Number(shape.getAttribute('height')) })),
    };
  });
}

test('an exact row clips paragraph borders by its height only, not beside the text (issue #1171)', async ({ browser, page }) => {
  test.setTimeout(120_000);
  // The row holds the boxed paragraph whole: all four sides show, the left one in the cell margin.
  const fits = await renderDocxArtifacts(generateExactRowDocx({ box: true, firstRowTwips: 900 }),
    { browser, ...options, outputs: ['html'] });
  const whole = await boxShapes(page, fits.html!);
  expect(whole.shapes).toHaveLength(4);
  expect(Math.min(...whole.shapes.map(shape => shape.x))).toBeLessThan(whole.textLeft);

  // A one-line row cuts the box: its bottom side is hidden, and the sides stop at the row's foot.
  const cut = await renderDocxArtifacts(generateExactRowDocx({ box: true }), { browser, ...options, outputs: ['html'] });
  const trimmed = await boxShapes(page, cut.html!);
  expect(trimmed.shapes).toHaveLength(3);
  for (const shape of trimmed.shapes) {
    expect(shape.y).toBeGreaterThanOrEqual(trimmed.row.top - 1);
    expect(shape.y + shape.height).toBeLessThanOrEqual(trimmed.row.bottom + 1);
  }
  const sides = trimmed.shapes.filter(shape => shape.height > shape.width);
  expect(sides).toHaveLength(2);
  expect(Math.min(...sides.map(shape => shape.x))).toBeLessThan(trimmed.textLeft);
});

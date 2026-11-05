import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateRowHeightUnitDocx } from './docx-row-height-unit-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function facts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const doc = await task.promise;
    const pages = [];
    for (let i = 1; i <= doc.numPages; i++) {
      const page = await doc.getPage(i), content = await page.getTextContent();
      pages.push({ view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
        text: item.str, transform: item.transform, width: item.width, height: item.height, font: content.styles[item.fontName],
      })) });
    }
    return pages;
  } finally { await task.destroy(); }
}

const options = { outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
  commentProfile: 'margin' as const, timeoutMs: 120_000 };
for (const [value, twips] of [['36pt', '720'], ['0.5in', '720'], ['1.27cm', '720'],
  ['12.7mm', '720'], ['3pc', '720'], ['3pi', '720'], ['25.075pt', '502']]) {
  test(`${value} row height prints like its whole-twip control`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateRowHeightUnitDocx(value);
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    const expected = await renderDocxArtifacts(generateRowHeightUnitDocx(twips), { browser, ...options });
    const physical = await facts(actual.pdf!);
    expect(physical).toHaveLength(1);
    expect(physical).toEqual(await facts(expected.pdf!));
    expect(physical[0].items.map(item => item.text).join(' ').match(/Alder\d+/g)).toEqual(['Alder01', 'Alder02', 'Alder03']);
    writeFileSync(testInfo.outputPath(`${value}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`${value}.pdf`), actual.pdf!);
  });
}

for (const [name, rows, rule, pages] of [['exact', 3, 'exact', 1], ['multipage', 30, 'atLeast', 2]] as const) {
  test(`${name} rows preserve complete cell inventory and control geometry`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = generateRowHeightUnitDocx('36pt', rows, rule);
    const actual = await renderDocxArtifacts(bytes, { browser, ...options });
    const expected = await renderDocxArtifacts(generateRowHeightUnitDocx('720', rows, rule), { browser, ...options });
    const physical = await facts(actual.pdf!);
    expect(physical).toHaveLength(pages);
    expect(physical).toEqual(await facts(expected.pdf!));
    const text = physical.flatMap(p => p.items).map(item => item.text).join(' ');
    for (const marker of ['Alder', 'Value']) expect(text.match(new RegExp(`${marker}\\d+`, 'g')))
      .toEqual(Array.from({ length: rows }, (_, i) => `${marker}${String(i + 1).padStart(2, '0')}`));
    writeFileSync(testInfo.outputPath(`${name}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`${name}.pdf`), actual.pdf!);
  });
}

test('a malformed height fails with a property-specific conversion diagnostic', async ({ browser }) => {
  await expect(renderDocxArtifacts(generateRowHeightUnitDocx('687,25'), { browser, ...options }))
    .rejects.toMatchObject({ code: 'conversion_failure', phase: 'docx_conversion',
      message: expect.stringContaining('w:trHeight/@w:val') });
});

test('unit and whole-twip oversized exact rows clip at the paper edge and retain the final blank page', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const control = await renderDocxArtifacts(generateRowHeightUnitDocx('15000', 1, 'exact'), { browser, ...options });
  const expected = await facts(control.pdf!);
  expect(expected.map(page => page.view)).toEqual([[0, 0, 612, 792], [0, 0, 612, 792]]);
  expect(expected[0].items.map(item => item.text).join(' ').trim().split(/\s+/)).toEqual([
    'Alder01', 'Original', 'height', 'record.', 'Value01', 'Complete', 'second', 'cell.',
  ]);
  expect(expected[1].items.map(item => item.text).join('').trim()).toBe('');
  // Unit strings are a converter compatibility input; the whole-twip control is schema valid.
  const bytes = generateRowHeightUnitDocx('750pt', 1, 'exact');
  const actual = await renderDocxArtifacts(bytes, { browser, ...options });
  expect(await facts(actual.pdf!)).toEqual(expected);
  writeFileSync(testInfo.outputPath('oversized-unit.docx'), bytes);
  writeFileSync(testInfo.outputPath('oversized-unit.pdf'), actual.pdf!);
  writeFileSync(testInfo.outputPath('oversized-twips.pdf'), control.pdf!);
});

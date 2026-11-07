import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { tableDocx } from './docx-authored-table-width-fixture.js';


test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [type, width, centered, rows, indent, expectedWidth, expectedLeft, layout] of [
  ['dxa', 500, true, 1, 0, 500, 56],
  ['dxa', 500, true, 24, 0, 500, 56],
  ['dxa', 580, false, 1, 0, 580, 72],
  ['dxa', 500, false, 1, 0, 500, 72],
  ['dxa', 500, true, 1, 0, 500, 56, 'inferred'],
  ['dxa', 500, false, 1, 0, 468, 72, 'autofit'],
  ['dxa', 500, false, 1, -45, 500, 27],
  ['auto', 500, false, 1, 0, 468, 72],
  ['pct', 500, false, 1, 0, 468, 72],
] as const) {
  test(`${type} ${width}pt ${centered ? 'center' : 'left'}, ${rows} rows, ${indent}pt indent, ${layout ?? 'default'} layout preserves authored geometry`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(tableDocx(width, type, centered, rows, indent, { layout }), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden',
      fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
    });
    expect(result.renderReport.warnings.filter(w => w.code === 'source_table_content_clipped')).toEqual([]);
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const tables = await page.locator('.page-content table').evaluateAll(tables => tables.map(table => {
      const rect = table.getBoundingClientRect(), paper = table.closest('.page-box')!.getBoundingClientRect();
      return { width: rect.width * 612 / paper.width, left: (rect.left - paper.left) * 612 / paper.width };
    }));
    expect(tables).toHaveLength(rows === 24 ? 2 : 1);
    for (const table of tables) {
      expect(Math.abs(table.width - expectedWidth)).toBeLessThan(0.6);
      expect(Math.abs(table.left - expectedLeft)).toBeLessThan(0.6);
    }
    const pdf = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const document = await pdf.promise, words = [];
      for (let p = 1; p <= document.numPages; p++)
        words.push(...(await (await document.getPage(p)).getTextContent()).items.flatMap(item => 'str' in item ? [item.str] : []));
      expect(words.join(' ').match(/Quill\d+[AB]|Finish/g)).toEqual([
        ...Array.from({ length: rows }, (_, row) => [`Quill${row + 1}A`, `Quill${row + 1}B`]).flat(), 'Finish',
      ]);
    } finally { await pdf.destroy(); }
  });
}

const clippedText = 'The ferry crews repaint the harbor signs every autumn before the storms arrive at the pier';

for (const kind of ['text', 'inline', 'promoted', 'textbox']) test(`balloon layout retains the paper-clipped table ${kind} diagnostic`, async ({ browser }) => {
  const bytes = tableDocx(580, 'dxa', false, 1, 0, { markup: true,
    ...(kind === 'text' ? { text: clippedText } : { imageWidth: 290, anchored: kind !== 'inline', textbox: kind === 'textbox' }),
  });
  const options = { browser, outputs: ['html', 'pdf'] as ('html' | 'pdf')[], reviewProfile: 'markup' as const,
    commentProfile: 'margin' as const, markupPresentation: 'word' as const, commentColumnPlacement: 'balloons' as const,
    fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true };
  const result = await renderDocxArtifacts(bytes, options);
  expect(result.html).toContain('page-markup-page');
  expect(result.renderReport.warnings.filter(w => w.code === 'source_table_content_clipped')).toHaveLength(1);
  await expect(renderDocxArtifacts(bytes, { ...options, unsupportedContent: 'strict' }))
    .rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});

test('paper clips an authored wide table with a source warning', async ({ browser }) => {
  const result = await renderDocxArtifacts(tableDocx(580, 'dxa', false, 1, 0, { text: clippedText }), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden',
    fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
  });
  const warnings = result.renderReport.warnings.filter(w => w.code === 'source_table_content_clipped');
  expect(warnings).toHaveLength(1);
  expect(warnings[0].anchorId).toMatch(/^p:/);
  expect(result.pageMap.fragments.some(f => f.anchorId === warnings[0].anchorId)).toBe(true);
});

test('strict export rejects table text cut by the paper edge', async ({ browser }) => {
  await expect(renderDocxArtifacts(tableDocx(580, 'dxa', false, 1, 0, { text: clippedText }), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
    fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});

for (const anchored of [false, true]) for (const strict of [false, true]) test(`partly off-paper ${anchored ? 'promoted' : 'inline'} table picture ${strict ? 'fails strict export' : 'reports clipping'}`, async ({ browser, page }) => {
  const pending = renderDocxArtifacts(tableDocx(580, 'dxa', false, 1, 0, { imageWidth: 290, anchored }), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden',
    unsupportedContent: strict ? 'strict' : 'warn',
    fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
  });
  if (strict) {
    await expect(pending).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
  } else {
    const result = await pending;
    const warnings = result.renderReport.warnings.filter(w => w.code === 'source_table_content_clipped');
    expect(warnings).toHaveLength(1);
    expect(warnings[0].anchorId).toMatch(/^p:/);
    await page.setContent(result.html!);
    expect(await page.locator('.page-box img').evaluate(image => image.getBoundingClientRect().right
      > image.closest('.page-box')!.getBoundingClientRect().right)).toBe(true);
  }
});

test('promoted table textbox retains partial-clipping diagnostics', async ({ browser }) => {
  const bytes = tableDocx(580, 'dxa', false, 1, 0, { imageWidth: 290, anchored: true, textbox: true });
  const result = await renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'],
    reviewProfile: 'final', commentProfile: 'hidden', strictFonts: true,
    fontDirectories: ['/usr/share/fonts/truetype/liberation'],
  });
  expect(result.renderReport.warnings.filter(w => w.code === 'source_table_content_clipped')).toHaveLength(1);
  expect(result.html).toContain('page-drawing-layer');
  await expect(renderDocxArtifacts(bytes, { browser, outputs: ['pdf'], reviewProfile: 'final',
    commentProfile: 'hidden', unsupportedContent: 'strict',
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});

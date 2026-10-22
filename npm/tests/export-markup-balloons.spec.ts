import { test, expect } from '@playwright/test';
import { writeFile } from 'node:fs/promises';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { markupBalloonsFixture } from './docx-markup-balloons-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('the PDF prints Word\'s All Markup balloons on the source paper size', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    commentColumnPlacement: 'balloons', timeoutMs: 120_000,
  });
  await writeFile(testInfo.outputPath('markup-balloons.pdf'), result.pdf!);
  expect(result.renderReport.options.commentColumnPlacement).toBe('balloons');
  expect(result.pageMap.pages.map(p => [p.width, p.height])).toEqual([[612, 792]]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(1);
    const page = await pdf.getPage(1), viewport = page.getViewport({ scale: 1 });
    expect(viewport.width).toBeCloseTo(612, 0); expect(viewport.height).toBeCloseTo(792, 0);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ')
      .replace(/\s+/g, ' ');
    for (const expected of ['Formatted: Font: Bold', 'Formatted: Indent: Left: 0.5"', 'Commented [AA1]:',
      'Commented [BB2R1]:', 'The buyer pays on delivery.']) {
      expect(text).toContain(expected);
    }
  } finally { await task.destroy(); }
});

test('the final profile prints comment balloons and no formatting balloons', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture('a4'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'margin', commentColumnPlacement: 'balloons',
    timeoutMs: 120_000,
  });
  expect(result.pageMap.pages.map(p => [Math.round(p.width), Math.round(p.height)])).toEqual([[595, 842]]);
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const page = await (await task.promise).getPage(1);
    const text = (await page.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ')
      .replace(/\s+/g, ' ');
    expect(text).toContain('Commented [AA1]:');
    expect(text).not.toContain('Formatted');
  } finally { await task.destroy(); }
});

test('the default export names no comment column placement', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(markupBalloonsFixture(), {
    browser, outputs: ['html'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
    timeoutMs: 120_000,
  });
  expect('commentColumnPlacement' in result.renderReport.options).toBe(false);
  expect(result.html).not.toContain('page-balloon');
});

for (const [name, options] of [
  ['inline comments', { reviewProfile: 'final', commentProfile: 'inline' }],
  ['the Docxodus presentation', { reviewProfile: 'markup', commentProfile: 'margin' }],
] as const) {
  test(`balloons with ${name} are an invalid argument`, async ({ browser }) => {
    await expect(renderDocxArtifacts(markupBalloonsFixture(), {
      browser, outputs: ['html'], ...options, commentColumnPlacement: 'balloons', timeoutMs: 60_000,
    })).rejects.toHaveProperty('code', 'invalid_argument');
  });
}

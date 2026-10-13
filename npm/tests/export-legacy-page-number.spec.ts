import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateLegacyPageNumberDocx } from './docx-legacy-page-number-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function lines(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    return await Promise.all(Array.from({ length: pdf.numPages }, async (_, index) => {
      const page = await pdf.getPage(index + 1);
      const content = await page.getTextContent();
      return { view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
        text: item.str, transform: item.transform, width: item.width, height: item.height,
        font: content.styles[item.fontName],
      })) };
    }));
  } finally { await task.destroy(); }
}

for (const [name, options] of Object.entries({ ordinary: {}, restartedRoman: { start: 7, roman: true } })) {
  test(`${name} legacy header numbers equal decimal PAGE fields on every PDF page`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const exportOptions = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'markup' as const,
      commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const actual = await renderDocxArtifacts(generateLegacyPageNumberDocx(options), exportOptions);
    const control = await renderDocxArtifacts(generateLegacyPageNumberDocx({ ...options, kind: 'field' }), exportOptions);
    expect(actual.pageCount).toBe(3);
    expect(actual.pageMap.fragments.filter(f => f.story === 'header')).toHaveLength(3);
    const pages = await lines(actual.pdf!);
    expect(pages).toEqual(await lines(control.pdf!));
    for (const [index, page] of pages.entries()) {
      expect(page.items.map(item => item.text).filter(Boolean))
        .toEqual([String((options.start ?? 1) + index), `Body${index + 1} Cedar archive page.`]);
    }
    writeFileSync(testInfo.outputPath(`${name}.docx`), generateLegacyPageNumberDocx(options));
    writeFileSync(testInfo.outputPath(`${name}.pdf`), actual.pdf!);
  });
}

test('non-editable legacy body blocks follow their own page while ordinary body PAGE caches stay unchanged', async ({ browser }) => {
  test.setTimeout(120_000);
  const exportOptions = { browser, outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
    commentProfile: 'margin' as const, timeoutMs: 120_000 };
  const actual = await renderDocxArtifacts(generateLegacyPageNumberDocx({ bodyNumber: true, start: 7, roman: true }), exportOptions);
  const control = await renderDocxArtifacts(generateLegacyPageNumberDocx({ bodyNumber: true, kind: 'field', start: 7, roman: true }), exportOptions);
  const actualPages = await lines(actual.pdf!);
  const controlPages = await lines(control.pdf!);
  for (const [index, page] of actualPages.entries()) {
    expect(page.items.map(item => item.text.trim()).filter(Boolean).join(' '))
      .toBe(`${7 + index} Body${index + 1} Cedar archive page. Current page: ${7 + index}`);
    expect(controlPages[index].items.map(item => item.text.trim()).filter(Boolean).join(' '))
      .toBe(`${7 + index} Body${index + 1} Cedar archive page. Current page: 1`);
  }
});

for (const kind of ['empty', 'tab'] as const) {
  test(`${kind} header still has measurable fragments and complete PDF body text`, async ({ browser }) => {
    test.setTimeout(120_000);
    const actual = await renderDocxArtifacts(generateLegacyPageNumberDocx({ kind }), {
      browser, outputs: ['pdf'], reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
    });
    expect(actual.pageCount).toBe(3);
    expect(actual.pageMap.fragments.filter(f => f.story === 'header')).toHaveLength(3);
    expect((await lines(actual.pdf!)).map(page => page.items.map(item => item.text).join('').trim()))
      .toEqual(['Body1 Cedar archive page.', 'Body2 Cedar archive page.', 'Body3 Cedar archive page.']);
  });
}

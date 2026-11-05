import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateLineBreakDocx, type LineBreakOptions } from './docx-line-break-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function inspect(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages = [];
    for (let index = 1; index <= pdf.numPages; index++) {
      const page = await pdf.getPage(index);
      const content = await page.getTextContent();
      pages.push({ view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
        text: item.str, transform: item.transform, width: item.width, height: item.height,
        font: content.styles[item.fontName],
      })), links: (await page.getAnnotations()).filter(a => a.subtype === 'Link').map(a => ({ url: a.url, rect: a.rect })) });
    }
    return pages;
  } finally { await task.destroy(); }
}

for (const [name, options] of Object.entries<LineBreakOptions>({
  plain: {}, formatted: { formatted: true }, blanks: { blankEvery: 36 }, hyperlinks: { hyperlink: true },
})) {
  test(`${name} hard breaks produce the control's PDF lines without loss or added page-top lines`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const exportOptions = { browser, outputs: ['html', 'pdf'] as const,
      reviewProfile: 'markup' as const, commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const control = await renderDocxArtifacts(generateLineBreakDocx({ ...options, paragraphs: true }), exportOptions);
    const source = generateLineBreakDocx(options);
    const actual = await renderDocxArtifacts(source, exportOptions);
    expect(actual.pageCount).toBe(control.pageCount);
    expect(actual.pageCount).toBe(3);
    const fragments = actual.pageMap.fragments.filter(f => f.story === 'body');
    expect(fragments).toHaveLength(3);
    expect(new Set(fragments.map(f => f.anchorId)).size).toBe(1);
    expect(new Set(fragments.map(f => f.fragmentId)).size).toBe(3);
    const actualPdf = await inspect(actual.pdf!);
    const expectedPdf = await inspect(control.pdf!);
    expect(actualPdf).toEqual(expectedPdf);
    expect(actualPdf.flatMap(p => p.items).map(i => i.text).join(' ').match(/Line\d+/g))
      .toEqual(Array.from({ length: 100 }, (_, index) => `Line${String(index + 1).padStart(3, '0')}`));
    writeFileSync(testInfo.outputPath(`${name}.docx`), source);
    writeFileSync(testInfo.outputPath(`${name}.pdf`), actual.pdf!);
  });
}

for (const constraint of ['keepLines', 'widowControl'] as const) {
  test(`${constraint} continues an oversized hard-break paragraph without losing text`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(generateLineBreakDocx({ [constraint]: true }), {
      browser, outputs: ['pdf'], reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(3);
    expect((await inspect(result.pdf!)).flatMap(p => p.items).map(i => i.text).join(' ').match(/Line\d+/g))
      .toEqual(Array.from({ length: 100 }, (_, index) => `Line${String(index + 1).padStart(3, '0')}`));
  });
}

import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateDrawingAnchorDocx } from './docx-drawing-anchor-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

test('zero relative textbox dimensions export the same verified PDF geometry as the absolute control', async ({ browser }, testInfo) => {
  test.setTimeout(120_000);
  const options = {
    horizontal: { relativeFrom: 'page' as const, offsetEmu: 914400 },
    vertical: { relativeFrom: 'page' as const, offsetEmu: 1143000 },
    prefix: 'PREFIX',
    text: 'Amber specimen label.',
    extentEmu: { cx: 2743200, cy: 914400 },
  };
  const exportOptions = {
    browser, outputs: ['html', 'pdf'] as const,
    reviewProfile: 'markup' as const, commentProfile: 'margin' as const,
    timeoutMs: 120_000,
  };
  const control = await renderDocxArtifacts(generateDrawingAnchorDocx(options), exportOptions);
  for (const relativeSize of [{ width: 0 }, { height: 0 }, { width: 0, height: 0 }]) {
    const actual = await renderDocxArtifacts(generateDrawingAnchorDocx({ ...options, relativeSize }), exportOptions);
    expect(actual.pageCount).toBe(1);
    expect(actual.pageMap).toEqual(control.pageMap);
    const inspect = async (bytes: Uint8Array) => {
      const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const page = await pdf.getPage(1);
        const content = await page.getTextContent();
        return { view: page.view, items: content.items.filter(item => 'str' in item).map(item => ({
          text: item.str, transform: item.transform, width: item.width, height: item.height,
          direction: item.dir, hasEOL: item.hasEOL, font: content.styles[item.fontName],
        })) };
      } finally { await task.destroy(); }
    };
    const printed = await inspect(actual.pdf!);
    const expected = await inspect(control.pdf!);
    // Full text transforms and extents must match, in addition to successful export.
    expect(printed).toEqual(expected);
    expect(printed.items.map(item => item.text).join(' '))
      .toContain('Amber specimen label.');
    writeFileSync(testInfo.outputPath(`textbox-${Object.keys(relativeSize).join('-')}.pdf`), actual.pdf!);
  }
});

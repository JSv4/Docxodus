import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateIndentationUnitDocx } from './docx-indentation-unit-fixture.js';

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

for (const [value, twips] of [['12pt', '240'], ['36pt', '720'], ['0.5in', '720'], ['1.27cm', '720'],
  ['12.7mm', '720'], ['3pc', '720'], ['3pi', '720']]) {
  test(`${value} indentation prints exactly like its twip control`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
      commentProfile: 'margin' as const, timeoutMs: 120_000 };
    const bytes = generateIndentationUnitDocx(value);
    const actual = await renderDocxArtifacts(bytes, options);
    const expected = await renderDocxArtifacts(generateIndentationUnitDocx(twips), options);
    const actualFacts = await facts(actual.pdf!);
    expect(actualFacts).toHaveLength(1);
    expect(actualFacts).toEqual(await facts(expected.pdf!));
    const text = actualFacts.flatMap(p => p.items).map(i => i.text).join(' ');
    for (const marker of ['EdgeMarker', 'LogicalMarker', 'FirstMarker', 'BorderMarker'])
      expect(text.match(new RegExp(marker, 'g'))).toHaveLength(1);
    writeFileSync(testInfo.outputPath(`${value}.docx`), bytes);
    writeFileSync(testInfo.outputPath(`${value}.pdf`), actual.pdf!);
  });
}

test('an unsupported unit remains an explicit conversion failure', async ({ browser }) => {
  await expect(renderDocxArtifacts(generateIndentationUnitDocx('12px'), { browser, outputs: ['pdf'],
    reviewProfile: 'markup', commentProfile: 'margin', timeoutMs: 120_000,
  })).rejects.toMatchObject({ code: 'conversion_failure', phase: 'docx_conversion' });
});

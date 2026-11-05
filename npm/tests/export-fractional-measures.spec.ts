import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { fractionalMeasuresDocx } from './docx-fractional-measures-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

async function pdfFacts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise;
    const pages = [];
    for (let index = 1; index <= pdf.numPages; index++) {
      const page = await pdf.getPage(index);
      const content = await page.getTextContent();
      pages.push({ view: page.view, items: content.items.filter(item => 'str' in item)
        .map(item => ({ text: item.str, transform: item.transform, width: item.width, height: item.height })) });
    }
    return pages;
  } finally { await task.destroy(); }
}

const options = { outputs: ['pdf'] as const, reviewProfile: 'markup' as const,
  markupPresentation: 'word' as const, commentProfile: 'margin' as const,
  commentColumnPlacement: 'balloons' as const, timeoutMs: 120_000 };

for (const [name, tab, rowHeight] of [
  ['style tab', '1827.875', '34.3625pt'],
  ['row height', '91.39375pt', '687.25'],
  ['both', '1827.875', '687.25'],
]) {
  test(`fractional ${name} exports complete text with the unit control's geometry`, async ({ browser }) => {
    const actual = await renderDocxArtifacts(fractionalMeasuresDocx({ tab, rowHeight }), { browser, ...options });
    const control = await renderDocxArtifacts(fractionalMeasuresDocx({
      tab: '91.39375pt', rowHeight: '34.3625pt', pageWidth: '612.025pt', pageHeight: '792.025pt',
    }), { browser, ...options });
    const pages = await pdfFacts(actual.pdf!);
    expect(pages).toHaveLength(1);
    expect(pages).toEqual(await pdfFacts(control.pdf!));
    expect(pages[0].items.map(item => item.text).join(' ').replace(/\s+/g, ' ').trim()).toBe(
      'Wren field inventory. Trail sample 07. All items retained. Field inspection complete.',
    );
  });
}

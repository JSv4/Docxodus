import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateBorderReopenDocx } from './docx-border-reopen-fixture.js';

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const [style, control] of [['single', false], ['dashed', false], ['single', true]] as const) {
  test(`${style} ${control ? 'control' : 'fractional'} border spacing stays consistent through HTML and PDF`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(generateBorderReopenDocx(style, control), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word',
      commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000,
    });
    expect(result.pageCount).toBe(1);
    expect(result.pageMap.fragments.filter(fragment => fragment.story === 'body')).toHaveLength(19);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(result.pageCount);
      const content = await (await pdf.getPage(1)).getTextContent();
      const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (let index = 1; index <= 18; index++) expect(text).toContain(`Kestrel card ${index}: specimen logged.`);
      expect(text).toContain('Inventory complete.');
    } finally { await task.destroy(); }
  });
}

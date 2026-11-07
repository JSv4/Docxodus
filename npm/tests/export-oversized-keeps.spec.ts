import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { oversizedKeepDocx } from './docx-oversized-keeps-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [keep, lines, opening, expected] of [
  ['keepLines', 73, 0, [40, 33]], ['widowControl', 73, 0, [40, 33]],
  ['keepNext', 73, 0, [40, 33]], ['styledKeepNext', 73, 0, [40, 33]], ['bothKeeps', 73, 0, [40, 33]],
  ['keepNext', 73, 8, [0, 40, 33]],
  ['keepNext', 80, 0, [40, 40, 0]],
  ['keepNext', 3, 37, [0, 3]], ['none', 3, 37, [3, 0]],
] as const) {
  test(`${keep}, ${lines} lines after ${opening}: preserves content and page membership`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(oversizedKeepDocx(keep, lines, opening), { browser, outputs: ['pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin',
      commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const pages: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        pages.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(pages.join(' ').match(/Copper\d+/g)).toEqual(Array.from({ length: lines }, (_, index) =>
        `Copper${String(index + 1).padStart(3, '0')}`));
      expect(pages.join(' ').match(/Final receipt signed\./g)).toHaveLength(1);
      expect(pages.map(page => page.match(/Copper\d+/g)?.length ?? 0)).toEqual(expected);
      expect(pages.at(-1)).toContain('Final receipt signed.');
      expect(pages.join(' ').match(/Opening\d+/g) ?? []).toEqual(Array.from({ length: opening }, (_, i) => `Opening${i + 1}`));
    } finally { await task.destroy(); }
  });
}

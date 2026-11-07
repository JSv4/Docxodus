import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { pageBreakDocx, type PageBreakCase } from './docx-page-break-fixture.js';

const cases: Array<[PageBreakCase, boolean?, boolean?]> = [
  ['own'], ['end'], ['mid'], ['none'], ['blank'], ['line'],
  ['leading'], ['repeated'], ['trailing'], ['space'],
  ['own', false], ['own', true], ['end', true], ['mid', true],
  ['end', false, true], ['end', true, true],
];
for (const [kind, splitMark, linked] of cases) {
  test(`page break ${kind}, split mark ${splitMark}, link ${linked}: retains only authored lines`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(pageBreakDocx(kind, splitMark, linked), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden',
      fontDirectories: ['/usr/share/fonts/truetype/liberation'], strictFonts: true,
    });
    if (linked) {
      await page.setContent(result.html!);
      await expect(page.locator('#Target')).toHaveCount(1);
      await expect(page.locator('a[href="#Target"]').filter({ hasText: 'First' })).toHaveCount(1);
    }
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(['none', 'blank', 'line', 'trailing'].includes(kind) ? 1 : 2);
      const words: Array<{ text: string; page: number; y: number }> = [];
      for (let page = 1; page <= pdf.numPages; page++) {
        for (const item of (await (await pdf.getPage(page)).getTextContent()).items)
          if ('str' in item && item.str.trim()) words.push({ text: item.str.trim(), page, y: item.transform[5] });
      }
      expect(words.map(w => w.text)).toEqual(kind === 'mid' ? ['First', 'After', 'Last']
        : kind === 'leading' ? ['Last'] : kind === 'trailing' ? ['First'] : ['First', 'Last']);
      const first = words[0], last = words.at(-1)!;
      if (kind === 'leading') expect(last.page).toBe(2);
      const offset = kind === 'blank' ? 32 : ['mid', 'none', 'line', 'space'].includes(kind) || splitMark ? 16 : 0;
      expect(Math.abs(first.y - last.y - offset)).toBeLessThan(0.4);
      if (kind === 'mid') expect(Math.abs(words[1].y - first.y)).toBeLessThan(0.4);
    } finally { await task.destroy(); }
  });
}

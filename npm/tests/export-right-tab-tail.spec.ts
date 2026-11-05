import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { LONG_TAB_TEXT, rightTabTailDocx } from './docx-right-tab-tail-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true } });

for (const [name, options] of [
  ['long right tail', {}],
  ['mixed tab and tail run', { mixedRun: true }],
  ['left control', { alignment: 'left' }],
  ['center control', { alignment: 'center' }],
  ['long prefix control', { longPrefix: true, longTail: false }],
  ['long prefix and tail', { longPrefix: true }],
  ['short right leader control', { longTail: false }],
  ['larger following font', { longTail: false, tailHalfPoints: 72,
    tailText: 'Sizing words must remain visible through TAIL COMPLETE' }],
  ['expanded letter spacing', { longTail: false, tailSpacing: 120,
    tailText: 'Oak ash elm fir pine yew birch cedar hemlock TAIL COMPLETE' }],
  ['wide glyphs', { longTail: false,
    tailText: 'WWWW '.repeat(9) + 'TAIL COMPLETE' }],
  ['multiple pages', { tailText: Array(4).fill(LONG_TAB_TEXT).join(' ') }],
  ['VML shape', { box: 'shape' }],
  ['VML rectangle', { box: 'rect' }],
  ...[22, 23, 24].map(count => [`wrapped prefix ${count}`, { longTail: false,
    prefixText: Array(count).fill('alpha').join(' '), tailText: 'WWWW '.repeat(9) + 'TAIL COMPLETE' }] as const),
  ['wrapped near-full tail', { longTail: false,
    prefixText: Array(23).fill('alpha').join(' '), tailText: 'WWWW '.repeat(7) + 'TAIL COMPLETE' }],
  ['wrapped short tail control', { longTail: false,
    prefixText: Array(23).fill('alpha').join(' '), tailText: 'END' }],
] as const) {
  test(`${name} keeps all text visible inside the page (#1488)`, async ({ browser, page }) => {
    const settings = Object.assign({ longPrefix: false, longTail: true, tailText: '', prefixText: '' }, options);
    const result = await renderDocxArtifacts(rightTabTailDocx(settings), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, pages: string[] = [];
      for (let n = 1; n <= pdf.numPages; n++) pages.push((await (await pdf.getPage(n)).getTextContent()).items
        .filter(item => 'str' in item).map(item => item.str).join(' '));
      const text = pages.join(' ').replace(/\s+/g, '');
      expect(text).toContain((settings.tailText || (settings.longTail ? LONG_TAB_TEXT : 'FINISH')).replace(/\s+/g, ''));
      const prefix = settings.prefixText || (settings.longPrefix ? LONG_TAB_TEXT : 'Item');
      expect(text).toContain(prefix.replace(/\s+/g, ''));
      const expected = prefix +
        (settings.tailText || (settings.longTail ? LONG_TAB_TEXT : 'FINISH'));
      expect((text.match(/Wren\d\d/g) ?? []).length).toBe((expected.match(/Wren\d\d/g) ?? []).length);
      if (name === 'multiple pages') expect(pdf.numPages).toBeGreaterThan(1);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await page.evaluate(() => document.fonts.ready);
    const clipped = await page.evaluate(() => {
      const errors: string[] = [];
      for (const band of Array.from(document.querySelectorAll('.page-content, [data-docx-textbox]'))) {
        const bounds = band.getBoundingClientRect(), walk = document.createTreeWalker(band, NodeFilter.SHOW_TEXT);
        let node: Node | null;
        while ((node = walk.nextNode())) for (const word of node.textContent!.matchAll(/\S+/g)) {
          const range = document.createRange(); range.setStart(node, word.index); range.setEnd(node, word.index + word[0].length);
          for (const rect of Array.from(range.getClientRects())) if (rect.left < bounds.left - 1 || rect.right > bounds.right + 1)
            errors.push(word[0]);
        }
      }
      return errors;
    });
    expect(clipped).toEqual([]);
  });
}

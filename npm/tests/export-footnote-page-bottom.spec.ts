import { expect, test } from '@playwright/test';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { footnotePageBottomDocx, type FootnotePageBottomOptions } from './docx-footnote-page-bottom-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

/**
 * Issue #1738: a double-spaced page's last line may end the page at its text height, its leading
 * running under the note band. The fit test let a citing paragraph's notes take that leading while
 * the band limit kept it for the body, so the band clipped its last note and the export failed.
 *
 * Each case lists the notes that start on each page, as Word lays the same document out: under
 * widow control a paragraph that brings a note keeps its leading beside it, so the fifth
 * paragraph's citing line moves to page 2 with its note; a paragraph that brings none, or one
 * without widow control, ends page 1 with its leading under the band. In compatibility mode 15
 * the band clears every line's leading (#1751), so those last lines move to page 2 instead.
 */
const cases: Array<[string, FootnotePageBottomOptions, string[][], number[]?]> = [
  ['a citing line at the page foot moves with its note', {}, [['1', '2', '3', '4'], ['5']]],
  ['the next page keeps citing beside its notes', { paragraphs: 6 }, [['1', '2', '3', '4'], ['5', '6']]],
  ['a paragraph citing nothing ends the page under the band', { cites: [2, 1, 1, 1, 0] }, [['1', '2', '3', '4', '5']]],
  ['without widow control the last line ends the page under the band', { widowControl: false }, [['1', '2', '3', '4', '5']]],
  ['without widow control a long note continues from its citing page', { widowControl: false, longNote: 5 },
    [['1', '2', '3', '4', '5'], []]],
  // Word in mode 15, which also gives the body lines on each page.
  ['in mode 15 a paragraph citing nothing keeps its leading above the band',
    { cites: [2, 1, 1, 1, 0], compatibilityMode: 15 }, [['1', '2', '3', '4', '5'], []], [23, 2]],
  ['in mode 15 the last line without widow control keeps its leading above the band',
    { widowControl: false, compatibilityMode: 15 }, [['1', '2', '3', '4', '5'], []], [24, 1]],
  ['in mode 15 a citing line at the page foot still moves with its note', { compatibilityMode: 15 },
    [['1', '2', '3', '4'], ['5']], [23, 2]],
  // Controls that already exported.
  ['four double-spaced paragraphs share one page', { paragraphs: 4 }, [['1', '2', '3', '4']]],
  ['one-and-a-half spacing', { line: 360, paragraphs: 8 }, [['1', '2', '3', '4', '5', '6'], ['7', '8']]],
];

for (const [name, options, expected, lines] of cases) {
  test(`footnotes at a page foot: ${name}`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(footnotePageBottomDocx(options), { browser, outputs: ['html'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    await page.setContent(result.html!);
    const pages = await page.locator('.page-box').evaluateAll(boxes => boxes.map(box => {
      const band = box.querySelector('.page-footnotes');
      return {
        cited: Array.from(box.querySelectorAll('.page-content a.footnote-ref'), ref => ref.getAttribute('data-footnote-id')),
        started: Array.from(box.querySelectorAll('.page-footnotes .footnote-item:not(.footnote-continuation)'),
          note => note.getAttribute('data-footnote-id')),
        overflow: Math.max(0, ...Array.from(band?.querySelectorAll('p') ?? [],
          p => p.getBoundingClientRect().bottom - band!.getBoundingClientRect().bottom)),
        lines: (() => {
          const tops = new Set<number>();
          const walker = document.createTreeWalker(box.querySelector('.page-content')!, NodeFilter.SHOW_TEXT);
          for (let text = walker.nextNode(); text; text = walker.nextNode()) {
            if (!/\S/.test(text.textContent ?? '') || text.parentElement!.closest('a.footnote-ref')) continue;
            const range = document.createRange();
            range.selectNodeContents(text);
            for (const rect of Array.from(range.getClientRects())) if (rect.height > 0) tops.add(Math.round(rect.top));
          }
          return tops.size;
        })(),
      };
    }));
    expect(pages.map(p => p.started)).toEqual(expected);
    if (lines) expect(pages.map(p => p.lines)).toEqual(lines);
    for (const { cited, started, overflow } of pages) {
      expect(started).toEqual(cited);
      expect(overflow).toBeLessThanOrEqual(1);
    }
  });
}

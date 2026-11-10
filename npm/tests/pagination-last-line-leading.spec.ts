import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

/**
 * Issue #1571: Word puts the extra leading of an auto line-spacing multiple below the text
 * (issue #908) and lets the last line on a page need only its text height. Word 2024 printed
 * 24 double-spaced lines where the text area held 23 lines plus 13.75pt, and 23 at 13.70pt;
 * the natural single line is 13.8pt. Spacing after already runs past the margin.
 *
 * The text area is `whole` full line pitches plus `sparePt`. An explicit 13.8pt natural line
 * keeps the cases independent of the installed serif's metrics.
 */
const cases = [
  { name: 'a double-spaced line whose text does not fit stays off the page', multiple: 2, whole: 23, sparePt: 13, lines: 23 },
  { name: 'a double-spaced line needs only its text height', multiple: 2, whole: 23, sparePt: 15, lines: 24 },
  { name: 'a 1.5-spaced line needs only its text height', multiple: 1.5, whole: 31, sparePt: 15, lines: 32 },
  { name: 'a double-spaced paragraph fragment needs only its last text height', multiple: 2, whole: 23, sparePt: 15, lines: 24, long: true },
  { name: 'a long double-spaced paragraph keeps a line whose text does not fit off the page', multiple: 2, whole: 23, sparePt: 13, lines: 23, long: true },
  { name: 'a run CSS centres in its line, without the lift, keeps the full-line test', multiple: 2, whole: 23, sparePt: 15, lines: 23, centred: true },
  { name: 'single lines with spacing after keep their fit', multiple: 1, whole: 26, sparePt: 15, lines: 27, afterPt: 10 },
  { name: 'single lines with spacing after keep a line that does not fit off the page', multiple: 1, whole: 26, sparePt: 13, lines: 26, afterPt: 10 },
];

for (const entry of cases) test(`last line on a page: ${entry.name}`, async ({ page }) => {
  await page.setContent('<style>body{font:12pt/13.8pt "Liberation Serif","Times New Roman",serif}</style>');
  await page.addScriptTag({ path: bundle });
  const lines = await page.evaluate(({ multiple, whole, sparePt, afterPt, long, centred }) => {
    const natural = 13.8;
    const pitch = natural * multiple + (afterPt ?? 0);
    const height = whole * pitch + sparePt;
    // The converter's automatic-spacing markup: each run's line is the multiple of its natural
    // line, lifted so the extra leading sits below the text.
    const paragraph = (text: string) => `<p style="margin:0 0 ${afterPt ?? 0}pt;--docx-auto-line-spacing:${multiple}">`
      + `<span style="position:relative;line-height:calc(1lh * var(--docx-auto-line-spacing));`
      + (centred ? "" : "top:min(0px,calc(1lh * (1 / var(--docx-auto-line-spacing) - 1) / 2))") + `">${text}</span></p>`;
    const body = long
      ? paragraph(Array.from({ length: 1400 }, (_, index) => `word${index + 1}`).join(' '))
      : Array.from({ length: whole * 2 + 5 }, (_, index) => paragraph(`Line ${index + 1} of the harbor log.`)).join('');
    document.body.insertAdjacentHTML('beforeend', `<div id="staging"><section data-section-index="0"
      data-page-width="612" data-page-height="${height + 144}" data-content-width="468" data-content-height="${height}"
      data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72">${body}</section></div>
      <div id="container"></div>`);
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
      document.getElementById('container'), { fragmentParagraphs: true, showPageNumbers: false }).paginate();
    const first = document.querySelector('#container .page-content')!;
    const tops = new Set<number>();
    for (const element of Array.from(first.querySelectorAll('p'))) {
      const range = document.createRange();
      range.selectNodeContents(element);
      for (const rect of Array.from(range.getClientRects())) if (rect.height > 0) tops.add(Math.round(rect.top));
    }
    return tops.size;
  }, entry);
  expect(lines).toBe(entry.lines);
});

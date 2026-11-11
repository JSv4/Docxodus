import { expect, test, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const presentation of ['editor', 'viewer', 'export']) for (const entry of [
  { name: '18pt grid', grid: 18, pitch: 18, count: 36, pages: 4 },
  { name: '12pt grid spans two slots', grid: 12, pitch: 24, count: 27, pages: 5 },
  { name: 'unrelated table', grid: 18, pitch: 18, count: 36, pages: 4, table: true },
  { name: 'no grid control', grid: 0, pitch: 2355 / 2048 * 12, count: 46, pages: 3 },
  // Word 2024 (#1551): the greater of the slots the text needs and the multiple of one pitch.
  { name: '1.15 lines on an 18pt grid', grid: 18, multiple: 1.15, pitch: 20.7, count: 31, pages: 4 },
  { name: 'double lines on a 12pt grid', grid: 12, multiple: 2, pitch: 24, count: 27, pages: 5 },
  { name: '0.8 lines on an 18pt grid', grid: 18, multiple: .8, pitch: 18, count: 36, pages: 4 },
  // A page's last grid line needs only its text, which Word centres: 19 x 33pt + 20.4pt fits 648pt.
  { name: '1.5 lines and 6pt after on an 18pt grid', grid: 18, multiple: 1.5, after: 6, pitch: 33, count: 20, pages: 6 },
  { name: '1.15 lines without a grid', grid: 0, multiple: 1.15, pitch: 2355 / 2048 * 12 * 1.15, count: 40, pages: 3 },
]) test(`${entry.multiple ? 'multiple' : 'single'} line grid ${presentation}: ${entry.name}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  // An automatic multiple carries the converter's run line height and lift.
  await page.setContent(`<style>p {font:12pt/normal 'Liberation Serif'; margin:0 0 ${entry.after ?? 0}pt; padding:0;
    --docx-auto-line-spacing:${entry.multiple ?? 1}; --docx-line-grid:${entry.grid}pt}
    ${entry.multiple ? `p > span {line-height:calc(1lh * var(--docx-auto-line-spacing)); position:relative;
      top:min(0px, calc(1lh * (1 / var(--docx-auto-line-spacing) - 1) / 2))}` : ''}</style>
    <div id="staging"><div data-section-index="0" data-page-width="612" data-page-height="792"
      data-content-width="468" data-content-height="648" data-margin-top="72" data-margin-right="72"
      data-margin-bottom="72" data-margin-left="72">
      ${Array.from({ length: 120 }, (_, i) => `<p><span>Cedar line ${i + 1}</span></p>`).join('')}
      ${entry.table ? '<table><tr><td><b>Separate table</b></td></tr></table>' : ''}
    </div></div><div id="pages"></div>`);
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const result = await page.evaluate(async presentation => {
    const staging = document.getElementById('staging')!;
    const pages = document.getElementById('pages')!;
    const inkOffset = (p: Element) => {
      const range = document.createRange(); range.selectNodeContents(p.firstElementChild!);
      return range.getBoundingClientRect().top - p.getBoundingClientRect().top;
    };
    const before = inkOffset(staging.querySelector('p')!);
    const moduleUrl = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(moduleUrl);
    applyUnroundedNormalLineHeights(staging);
    const resolved = staging.innerHTML;
    applyUnroundedNormalLineHeights(staging);
    const repeated = staging.innerHTML === resolved;
    const textUrl = 'http://localhost:8083/word-text.js';
    const { prepareWordText, paintWordText } = await import(textUrl);
    if (presentation === 'export') prepareWordText(staging);
    const engine = new (window as any).DocxodusPagination.PaginationEngine(staging, pages,
      { scale: 1, showPageNumbers: false, alignBaselinesToWord: presentation !== 'editor' });
    engine.paginate();
    if (presentation === 'export') paintWordText(pages);
    const read = () => {
      const paragraphs = Array.from(pages.querySelector('.page-content')!.querySelectorAll('p'));
      const tops = paragraphs.map(p => p.getBoundingClientRect().top);
      return { count: paragraphs.length, pages: pages.querySelectorAll('.page-box').length,
        pitch: (tops.at(-1)! - tops[0]) / (tops.length - 1) * .75, ink: inkOffset(paragraphs[0]) };
    };
    return { before, first: read(), repeated };
  }, presentation);
  expect(result.first.count).toBe(entry.count);
  expect(result.first.pages).toBe(entry.pages);
  expect(Math.abs(result.first.pitch - entry.pitch)).toBeLessThan(1 / 64);
  // The preexisting natural-height pass removes whole-pixel line rounding. A multiple's ink
  // starts in a CSS-rounded line box, which that pass also moves.
  if (presentation === 'editor' && !entry.multiple) expect(Math.abs(result.first.ink - result.before)).toBeLessThan(.5);
  expect(result.repeated).toBe(true);
});

// Word 2024 puts 36 lines of this paragraph on an 18pt grid in a 646pt body and 35 in a 644pt one:
// the last line needs only its centred text, 15.9pt of its 18pt slot (#1551).
for (const entry of [{ table: false, height: 646, rows: 36 }, { table: false, height: 644, rows: 35 },
  { table: true, height: 646, rows: 36 }, { table: true, height: 646, rows: 36, shifted: true }])
test(`wrapped grid paragraph, table=${entry.table}, body=${entry.height}pt, arbitrary offset=${!!entry.shifted}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.setContent(`<style>p {font:12pt/normal 'Liberation Serif';margin:0;--docx-auto-line-spacing:1;--docx-line-grid:18pt}</style>
    <div id="staging"><div data-section-index="0" data-page-width="612" data-page-height="${entry.height + 144}"
      data-content-width="468" data-content-height="${entry.height}" data-margin-top="72" data-margin-right="72"
      data-margin-bottom="72" data-margin-left="72"><p><span style="white-space:pre-wrap">${Array.from({ length: 900 }, (_, i) => `Cedar${i}`).join(' ')}</span></p>
      ${entry.table ? '<table><tr><td>Separate table</td></tr></table>' : ''}</div></div><div id="pages"></div>`);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const result = await page.evaluate(async shifted => {
    await document.fonts.ready;
    const metricsUrl = 'http://localhost:8083/line-metrics.js', textUrl = 'http://localhost:8083/word-text.js';
    const { applyUnroundedNormalLineHeights } = await import(metricsUrl);
    const { prepareWordText } = await import(textUrl);
    const staging = document.getElementById('staging')!, pages = document.getElementById('pages')!;
    applyUnroundedNormalLineHeights(staging); prepareWordText(staging);
    const fitted = staging.querySelector('p > span')!.getAttribute('data-docx-word-fit');
    if (shifted) Object.assign(staging.querySelector<HTMLElement>('p > span')!.style, { position: 'relative', top: '-1px' });
    new (window as any).DocxodusPagination.PaginationEngine(staging, pages,
      { scale: 1, showPageNumbers: false, fragmentParagraphs: true }).paginate();
    const paragraph = pages.querySelector('.page-content p')!;
    const rows = new Set<number>();
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const match of node.textContent!.matchAll(/\S+/g)) {
      const range = document.createRange(); range.setStart(node, match.index!); range.setEnd(node, match.index! + match[0].length);
      rows.add(Math.round(range.getBoundingClientRect().top * 4));
    }
    return { fitted, rows: rows.size, text: Array.from(pages.querySelectorAll('.page-content p')).map(p => p.textContent).join(' ') };
  }, !!entry.shifted);
  expect(result.fitted).toBe('0');
  if (entry.shifted) expect(result.rows).toBeGreaterThan(entry.rows);
  else expect(result.rows).toBe(entry.rows);
  expect(result.text.match(/Cedar\d+/g)).toEqual(Array.from({ length: 900 }, (_, i) => `Cedar${i}`));
});

test('unmeasured mixed, raised, nested, and exact lines keep their existing geometry', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.setContent(`<style>p {font:12pt/normal 'Liberation Serif';margin:0}</style><main></main>`);
  const controls = await page.evaluate(async () => {
    const moduleUrl = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights } = await import(moduleUrl);
    const root = document.querySelector('main')!;
    const pairs = [
      '<span>Plain</span><span style="font-size:20pt">Large</span>',
      '<span style="vertical-align:super">Raised</span>', '<span><strong>Nested</strong></span>',
      '<span style="line-height:12pt">Exact</span>',
    ].map(content => [0, 18].map(grid => {
      const paragraph = document.createElement('p'); paragraph.innerHTML = content;
      paragraph.style.setProperty('--docx-line-grid', `${grid}pt`); root.append(paragraph); return paragraph;
    }));
    applyUnroundedNormalLineHeights(root);
    return pairs.map(pair => pair.map(p => ({ height: p.getBoundingClientRect().height,
      lines: Array.from(p.querySelectorAll<HTMLElement>('span')).map(s => s.style.lineHeight),
      grid: p.querySelector<HTMLElement>('span')!.style.getPropertyValue('--docx-word-grid-height') })));
  });
  for (const [control, grid] of controls) { expect(grid).toEqual(control); expect(grid.grid).toBe(''); }
});

/** Each paragraph's height in points after the natural-height pass, and its last run's text offset from its top. */
async function gridLines(page: Page, paragraphs: Array<{ html: string; style: string }>) {
  await page.goto('/test-harness.html');
  // Standards mode, as the harness and export pages render: quirks mode ignores the paragraph's own
  // line height on a line that holds only inline elements. Direct children carry a multiple's lift.
  await page.setContent(`<!doctype html><style>p {font:12pt/normal 'Liberation Serif';margin:0}
    p[style*="auto-line-spacing"] > * {line-height:calc(1lh * var(--docx-auto-line-spacing)); position:relative;
      top:min(0px, calc(1lh * (1 / var(--docx-auto-line-spacing) - 1) / 2))}</style><main></main>`);
  return page.evaluate(async paragraphs => {
    await document.fonts.ready;
    const moduleUrl = 'http://localhost:8083/line-metrics.js';
    const { applyUnroundedNormalLineHeights, prepareMixedLineHeights } = await import(moduleUrl);
    const root = document.querySelector('main')!;
    const elements = paragraphs.map(({ html, style }) => {
      const paragraph = document.createElement('p'); paragraph.innerHTML = html; paragraph.style.cssText = style;
      root.append(paragraph); return paragraph;
    });
    // In the paginator's order.
    prepareMixedLineHeights(root);
    applyUnroundedNormalLineHeights(root);
    return elements.map(p => {
      const range = document.createRange(); range.selectNodeContents(Array.from(p.querySelectorAll('span')).at(-1)!);
      const box = p.getBoundingClientRect();
      return [box.height, range.getBoundingClientRect().top - box.top].map(px => Math.round(px * .75 * 10) / 10);
    });
  }, paragraphs);
}

// Word 2024 (#1551): an at-least line is max(whole grid lines, minimum). The converter keeps the minimum's
// line height, which a line the grid pass cannot measure (a bold word) and a taller minimum keep.
test('at-least lines take whole grid lines or their minimum', async ({ page }) => {
  const atLeast = (size: number, minimum: number, grid: number) =>
    `font-size:${size}pt;line-height:${minimum}pt;--docx-min-line-height:${minimum}pt;--docx-line-grid:${grid}pt`;
  const lines = await gridLines(page, [
    { html: '<span>Plain</span>', style: atLeast(12, 15, 18) },
    { html: '<span>Plain</span><b>Bold</b>', style: atLeast(12, 15, 18) },
    { html: '<span>Plain</span>', style: atLeast(24, 30, 18) },
    { html: '<span>Plain</span>', style: atLeast(24, 30, 0) },
    { html: '<span>Plain</span>', style: atLeast(12, 20, 18) },
    // Mixed sizes are measured line by line, as without a grid.
    ...[18, 0].map(grid => ({ html: '<span>Plain</span><span style="font-size:30pt">Large</span>', style: atLeast(12, 20, grid) })),
  ]);
  expect(lines.slice(0, 5).map(([height]) => height)).toEqual([18, 15, 36, 30, 20]);
  expect(lines[5][0]).toBeGreaterThan(30);
  expect(lines[5][0]).toBe(lines[6][0]);
  // A grid line's text sits where a single-spaced grid line's does.
  const [[, single], [, larger]] = await gridLines(page, [{ html: '<span>Plain</span>', style: '--docx-line-grid:18pt' },
    { html: '<span>Plain</span>', style: 'font-size:24pt;--docx-line-grid:18pt' }]);
  expect([lines[0][1], lines[2][1]]).toEqual([single, larger]);
});

// Word 2024 (#1551) lays out a paragraph that shows a tracked insertion or deletion like one without it,
// on a grid or off it. The converter's markup wraps those runs in <ins> or <del>.
test('tracked-change wrappers keep their runs on the line rules', async ({ page }) => {
  const inserted = '<span>Plain </span><ins><span>added</span></ins>';
  const lines = await gridLines(page, [
    { html: inserted, style: '--docx-line-grid:18pt' },
    { html: '<span>Plain</span><del><span> gone</span></del>', style: '--docx-line-grid:18pt' },
    { html: inserted, style: '' },
    { html: inserted, style: '--docx-auto-line-spacing:1.5;--docx-line-grid:18pt' },
    { html: inserted, style: 'font-size:24pt;line-height:30pt;--docx-min-line-height:30pt;--docx-line-grid:18pt' },
    // Off the grid, an at-least line is max(natural, minimum), as without the markup.
    { html: '<span>Plain</span><del><span> gone</span></del>', style: 'font-size:24pt;line-height:15pt;--docx-min-line-height:15pt' },
    { html: inserted, style: 'font-size:10pt;--docx-min-line-height:13pt' },
    // A wrapped run with a line height of its own stays off the grid, as an unwrapped one does.
    ...[18, 0].map(grid => ({ html: '<ins><span style="line-height:12pt">Exact</span></ins>', style: `--docx-line-grid:${grid}pt` })),
  ]);
  expect(lines.map(([height]) => height)).toEqual([18, 18, 13.8, 27, 36, 27.6, 13, 13.8, 13.8]);
  // The inserted text sits where an unrevised grid line's does: its wrapper moves it once.
  const [[, single], [, larger]] = await gridLines(page, [{ html: '<span>Plain</span>', style: '--docx-line-grid:18pt' },
    { html: '<span>Plain</span>', style: 'font-size:24pt;--docx-line-grid:18pt' }]);
  expect([lines[0][1], lines[4][1]]).toEqual([single, larger]);
});

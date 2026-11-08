import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const presentation of ['editor', 'viewer', 'export']) for (const entry of [
  { name: '18pt grid', grid: 18, pitch: 18, count: 36, pages: 4 },
  { name: '12pt grid spans two slots', grid: 12, pitch: 24, count: 27, pages: 5 },
  { name: 'unrelated table', grid: 18, pitch: 18, count: 36, pages: 4, table: true },
  { name: 'no grid control', grid: 0, pitch: 2355 / 2048 * 12, count: 46, pages: 3 },
]) test(`single line grid ${presentation}: ${entry.name}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.setContent(`<style>p {font:12pt/normal 'Liberation Serif'; margin:0; padding:0;
    --docx-auto-line-spacing:1; --docx-line-grid:${entry.grid}pt}</style>
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
  // The preexisting natural-height pass removes whole-pixel line rounding.
  if (presentation === 'editor') expect(Math.abs(result.first.ink - result.before)).toBeLessThan(.5);
  expect(result.repeated).toBe(true);
});

for (const entry of [{ table: false }, { table: true }, { table: true, shifted: true }])
test(`wrapped grid paragraph, table=${entry.table}, arbitrary offset=${!!entry.shifted}`, async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.setContent(`<style>p {font:12pt/normal 'Liberation Serif';margin:0;--docx-auto-line-spacing:1;--docx-line-grid:18pt}</style>
    <div id="staging"><div data-section-index="0" data-page-width="612" data-page-height="790"
      data-content-width="468" data-content-height="646" data-margin-top="72" data-margin-right="72"
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
  if (entry.shifted) expect(result.rows).toBeGreaterThan(35);
  else expect(result.rows).toBe(35);
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

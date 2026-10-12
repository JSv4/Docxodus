import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';

const wordText = readFileSync(new URL('../dist/word-text.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const lineMetrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const geometry = `data-section-index="0" data-page-width="444" data-page-height="700"
  data-content-width="300" data-content-height="556" data-margin-top="72"
  data-margin-bottom="72" data-margin-left="72" data-margin-right="72"`;
const words = Array.from({ length: 150 }, (_, i) => `word${i}`).join('  ');
const styles = `<style>p{font:11pt 'Liberation Serif',serif;line-height:normal;margin:0 0 6pt;--docx-auto-line-spacing:1.15}
  span.run{line-height:calc(1lh * 1.15);position:relative;top:min(0px,calc(1lh * (1/1.15 - 1)/2));--docx-font-kerning:none}</style>`;

test('fractional paragraph paint preserves every character, block box and PageMap anchor', async ({ page }) => {
  await page.setContent(`${styles}<div id="staging"><section ${geometry}>
    <p data-source-anchor-id="p:body:plain" style="text-align:justify"><span class="run">${words}</span></p>
    <p data-source-anchor-id="p:body:tail"><span class="run">Following text</span></p>
    </section></div><div id="container"></div>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  await page.addScriptTag({ path: new URL('../dist/pagination.bundle.js', import.meta.url).pathname });
  const result = await page.evaluate(() => {
    const api = window as any;
    const root = document.getElementById('staging')!;
    api.applyUnroundedNormalLineHeights(root);
    const prepared = api.prepareWordText(root);
    const engine = new api.DocxodusPagination.PaginationEngine(root, document.getElementById('container'), { showPageNumbers: false });
    const pagination = engine.paginate();
    const container = document.getElementById('container')!;
    const before = engine.materializePageMap(0, 'word-text-test');
    const painted = api.paintWordText(container);
    const after = engine.materializePageMap(0, 'word-text-test');
    const paragraphs = Array.from(container.querySelectorAll('p'));
    return { prepared, painted, pages: pagination.totalPages, before, after,
      text: paragraphs.map(p => p.textContent), svgCount: container.querySelectorAll('svg').length,
      lines: Array.from(paragraphs[0].querySelectorAll('svg > text')).map(t => Number(t.getAttribute('y')) / 4 * 0.75),
      repeat: api.paintWordText(container),
    };
  });
  expect(result.prepared).toBe(2);
  expect(result.painted).toBe(2);
  expect(result.pages).toBe(1);
  expect(result.text).toEqual([words, 'Following text']);
  expect(result.before).toEqual(result.after);
  expect(result.svgCount).toBe(2);
  expect(result.repeat).toBe(0);
  expect(result.lines.length).toBeGreaterThan(3);
  for (let i = 1; i < result.lines.length; i++)
    expect(result.lines[i] - result.lines[i - 1]).toBeCloseTo(14.5, 1);
});

for (const [name, content] of [
  ['picture', '<img width="20" height="20">'],
  ['table', '<table><tr><td>Cell</td></tr></table>'],
  ['note', '<a class="footnote-ref">1</a>'],
  ['columns with rich text', '<section data-cols="2"><p><span class="run"><b>Rich column</b></span></p></section>'],
]) {
  test(`complex ${name} flow retains its original HTML layout`, async ({ page }) => {
    await page.setContent(`${styles}<main><p><span class="run">AV A plain paragraph</span></p>${content}</main>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    const result = await page.evaluate(() => {
      const api = window as any;
      const root = document.querySelector('main')!;
      api.applyUnroundedNormalLineHeights(root);
      const before = root.innerHTML;
      return { count: api.prepareWordText(root), before, after: root.innerHTML };
    });
    expect(result.count).toBe(0);
    expect(result.after).toBe(result.before);
  });
}

test('rich runs and unsupported scripts retain their source structure', async ({ page }) => {
  await page.setContent(`${styles}<main><p><span class="run">Plain text beside rich content</span></p>
    <p><span class="run"><b>Rich</b></span></p>
    <p><span class="run">漢字</span></p><p><span class="run" style="text-decoration:underline">Underlined</span></p>
    <p style="padding-top:4pt"><span class="run">Padded paragraph</span></p>
    <p style="border:1px solid"><span class="run">Bordered paragraph</span></p>
    <p><span class="run" data-source-anchor-id="p:body:run">Run identity</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const result = await page.evaluate(() => {
    const api = window as any;
    const root = document.querySelector('main')!;
    api.applyUnroundedNormalLineHeights(root);
    const before = root.innerHTML;
    return { count: api.prepareWordText(root), before, after: root.innerHTML };
  });
  expect(result.count).toBe(0);
  expect(result.after).toBe(result.before);
});

test('the source kerning choice overrides the Word default', async ({ page }) => {
  await page.setContent(`${styles}<main><p><span class="run">AV A</span></p>
    <p><span class="run" style="--docx-font-kerning:normal">AV A</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const result = await page.evaluate(() => {
    const api = window as any;
    const root = document.querySelector('main')!;
    api.applyUnroundedNormalLineHeights(root);
    api.prepareWordText(root);
    return Array.from(root.querySelectorAll('span.run')).map(run => {
      const range = document.createRange(); range.selectNodeContents(run);
      return { kerning: getComputedStyle(run).fontKerning, width: range.getBoundingClientRect().width };
    });
  });
  expect(result.map(r => r.kerning)).toEqual(['none', 'normal']);
  expect(result[1].width).toBeLessThan(result[0].width);
});

test('painting checks cancellation while measuring a long paragraph', async ({ page }) => {
  await page.setContent(`${styles}<main><p><span class="run">${words}</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const result = await page.evaluate(() => {
    const api = window as any;
    const root = document.querySelector('main')!;
    api.applyUnroundedNormalLineHeights(root);
    api.prepareWordText(root);
    const before = root.textContent;
    let checks = 0;
    try {
      api.paintWordText(root, () => { if (++checks === 3) throw new Error('cancelled'); });
      return { cancelled: false, before, after: root.textContent };
    } catch (error) {
      return { cancelled: (error as Error).message === 'cancelled', before, after: root.textContent };
    }
  });
  expect(result.cancelled).toBe(true);
  expect(result.after).toBe(result.before);
});

test('a fragmented column container keeps all of its HTML text and geometry', async ({ page }) => {
  await page.setContent(`${styles}<main><div style="column-count:2;column-gap:24pt;height:100pt;width:300pt">
    <p><span class="run">${words}</span></p><p><span class="run">Following paragraph</span></p>
    </div></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const result = await page.evaluate(() => {
    const api = window as any; const root = document.querySelector('main')!;
    api.applyUnroundedNormalLineHeights(root);
    const prepared = api.prepareWordText(root);
    const paragraphs = Array.from(root.querySelectorAll('p'));
    const before = paragraphs.map(p => ({ text: p.textContent, box: p.getBoundingClientRect().toJSON() }));
    const fragments = paragraphs[0].getClientRects().length;
    const painted = api.paintWordText(root);
    return { prepared, painted, fragments, before,
      after: paragraphs.map(p => ({ text: p.textContent, box: p.getBoundingClientRect().toJSON() })),
      pending: root.querySelectorAll('[data-docx-word-text-prepared]').length };
  });
  expect(result.prepared).toBe(2);
  expect(result.fragments).toBeGreaterThan(1);
  expect(result.painted).toBe(0);
  expect(result.after).toEqual(result.before);
  expect(result.pending).toBe(0);
});

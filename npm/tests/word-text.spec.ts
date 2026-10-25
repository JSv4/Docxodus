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
  ['inline picture', '<p><span><img width="20" height="20"></span><span class="run">Inline picture text</span></p>'],
  ['positioned picture', '<p><span><span style="float:left;position:relative"><img width="20" height="20"></span></span><span class="run">Positioned picture text</span></p>'],
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

test('a convex both-sides contour places short text before the picture and longer text after it', async ({ page }) => {
  const text = 'a b An authored caption retains every word.';
  await page.setContent(`${styles}<main class="page-content" style="position:relative;width:400pt;height:400pt">
    <p data-source-anchor-id="p:body:picture"><span class="run"><span data-docx-edge-picture="true" data-docx-wrap-text="bothSides"
      style="float:left;width:72pt;height:72pt;shape-outside:polygon(50% 0%,100% 50%,50% 100%,0% 50%) border-box;
        shape-margin:5pt;--docx-wrap-distance-left:5pt;--docx-wrap-distance-right:5pt"><img alt="Authored contour" style="width:72pt;height:72pt"></span></span><span class="run">${text}</span></p>
    <p data-source-anchor-id="p:body:tail"><span class="run">Following text</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const result = await page.evaluate(() => {
    const api = window as any; const root = document.querySelector<HTMLElement>('main')!;
    api.applyUnroundedNormalLineHeights(root);
    const prepared = api.prepareWordText(root);
    const paragraphs = Array.from(root.querySelectorAll('p'));
    const before = paragraphs.map(p => p.getBoundingClientRect().toJSON());
    const image = root.querySelector('img')!; const imageBefore = image.getBoundingClientRect().toJSON();
    const painted = api.paintWordText(root);
    return { prepared, painted, before, after: paragraphs.map(p => p.getBoundingClientRect().toJSON()),
      imageSame: image === root.querySelector('img'), imageBefore, imageAfter: image.getBoundingClientRect().toJSON(),
      text: paragraphs[0].textContent,
      x: Array.from(paragraphs[0].querySelectorAll('svg tspan')).map(t => Number(t.getAttribute('x')) / 4 * 0.75) };
  });
  expect(result.prepared).toBe(2);
  expect(result.painted).toBe(2);
  expect(result.text).toBe(text);
  expect(result.after).toEqual(result.before);
  expect(result.imageSame).toBe(true);
  expect(result.imageAfter).toEqual(result.imageBefore);
  expect(result.x[0]).toBeCloseTo(0, 1);
  expect(result.x[1]).toBeLessThan(20);
  expect(result.x[2]).toBeGreaterThan(40);
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

for (const mode of [14, 15]) {
  test(`mode ${mode} fits words with its measured space limit and paints natural final-line spaces`, async ({ page }) => {
    await page.setContent(`<style>
      p { font: 12px monospace; line-height:normal; margin:0; text-align:justify; --docx-auto-line-spacing:1 }
      .run { white-space:pre-wrap; line-height:normal; --docx-font-kerning:none }
      </style><main data-docx-compatibility-mode="${mode}"><p><span class="run">WWWW WWWW WWWW WWWW WWWW</span></p></main>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const api = window as any, p = document.querySelector('p')!;
      const canvas = document.createElement('canvas').getContext('2d')!;
      canvas.font = '12px monospace';
      const word = canvas.measureText('WWWW').width, space = canvas.measureText(' ').width;
      p.style.width = `${3 * word + 2 * space * .8}px`;
      api.applyUnroundedNormalLineHeights(document.querySelector('main'));
      api.prepareWordText(document.querySelector('main'));
      const run = p.firstElementChild!;
      const rows = new Map<number, string[]>();
      const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const match of node.textContent!.matchAll(/\S+/g)) {
        const range = document.createRange(); range.setStart(node, match.index!); range.setEnd(node, match.index! + match[0].length);
        const y = range.getBoundingClientRect().top;
        const words = rows.get(y) ?? []; words.push(match[0]); rows.set(y, words);
      }
      const height = p.getBoundingClientRect().height;
      api.paintWordText(document.querySelector('main'));
      const lines = Array.from(p.querySelectorAll('svg > text'));
      const last = Array.from(lines.at(-1)!.querySelectorAll('tspan')).map(w => Number(w.getAttribute('x')) / 4);
      return { wordsPerLine: Array.from(rows.values()).map(words => words.length),
        lastGap: last.length > 1 ? last[1] - last[0] : null, naturalGap: word + space,
        heightBefore: height, heightAfter: p.getBoundingClientRect().height };
    });
    expect(actual.wordsPerLine).toEqual(mode === 15 ? [3, 2] : [2, 2, 1]);
    if (mode === 15) expect(actual.lastGap).toBeCloseTo(actual.naturalGap, 2);
    expect(actual.heightAfter).toBe(actual.heightBefore);
  });
}

test('a word split across visual lines keeps its complete HTML fallback', async ({ page }) => {
  await page.setContent(`<style>p{font:12px monospace;line-height:normal;width:50px;margin:0;overflow-wrap:anywhere;--docx-auto-line-spacing:1}
    .run{white-space:pre-wrap;line-height:normal}</style><main><p><span class="run">${'LongWord'.repeat(10)}</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const actual = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!, p = root.querySelector('p')!;
    api.applyUnroundedNormalLineHeights(root); api.prepareWordText(root);
    const before = p.innerHTML, box = p.getBoundingClientRect().toJSON();
    return { painted: api.paintWordText(root), before, after: p.innerHTML, box, afterBox: p.getBoundingClientRect().toJSON() };
  });
  expect(actual.painted).toBe(0);
  expect(actual.after).toBe(actual.before);
  expect(actual.afterBox).toEqual(actual.box);
});

for (const sectionTag of ['section', 'div']) {
test(`a ${sectionTag} page continuation keeps justification through the last line of each non-final fragment`, async ({ page }) => {
  await page.setContent(`${styles}<style>span.run{white-space:pre-wrap;position:static}</style><div id="staging"><${sectionTag} ${geometry.replace('700', '260').replace('556', '116')}>
    <p data-source-anchor-id="p:body:plain" style="text-align:justify"><span class="run">${words}</span></p>
    </${sectionTag}></div><div id="container"></div>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  await page.addScriptTag({ path: new URL('../dist/pagination.bundle.js', import.meta.url).pathname });
  const actual = await page.evaluate(() => {
    const api = window as any, root = document.getElementById('staging')!, container = document.getElementById('container')!;
    api.applyUnroundedNormalLineHeights(root); api.prepareWordText(root);
    const engine = new api.DocxodusPagination.PaginationEngine(root, container, { showPageNumbers:false, fragmentParagraphs:true });
    const result = engine.paginate(), before = engine.materializePageMap(0, 'fragment-test');
    const fragments = Array.from(container.querySelectorAll('p'));
    const alignments = fragments.map(p => getComputedStyle(p).textAlignLast);
    const painted = api.paintWordText(container);
    return { pages:result.totalPages, painted, alignments, before, after:engine.materializePageMap(0, 'fragment-test'),
      text: fragments.map(p => p.textContent).join('') };
  });
  expect(actual.pages).toBeGreaterThan(1);
  expect(actual.painted).toBe(actual.pages);
  expect(actual.text).toBe(words);
  expect(actual.after).toEqual(actual.before);
  expect(actual.alignments.slice(0, -1)).toEqual(Array(actual.pages - 1).fill('justify'));
  expect(actual.alignments.at(-1)).toBe('auto');
});
}


test('modern Word can compress a final justified line that exceeds its natural width', async ({ page }) => {
  await page.setContent(`<style>p{font:12px monospace;line-height:normal;margin:0;text-align:justify;--docx-auto-line-spacing:1}
    .run{white-space:pre-wrap;line-height:normal}</style><main data-docx-compatibility-mode="15"><p><span class="run">WWWW WWWW WWWW</span></p></main>`);
  await page.addScriptTag({ content: lineMetrics });
  await page.addScriptTag({ content: wordText });
  const actual = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!, p = root.querySelector('p')!;
    const context = document.createElement('canvas').getContext('2d')!; context.font = '12px monospace';
    const word = context.measureText('WWWW').width, space = context.measureText(' ').width;
    p.style.width = `${word * 3 + space * 2 * .8}px`;
    api.applyUnroundedNormalLineHeights(root); api.prepareWordText(root); api.paintWordText(root);
    const lines = p.querySelectorAll('svg > text'), last = lines[0].lastElementChild!;
    return { lines:lines.length, end:Number(last.getAttribute('x')) / 4 + word, width:p.getBoundingClientRect().width };
  });
  expect(actual.lines).toBe(1);
  expect(actual.end).toBeCloseTo(actual.width, 2);
});

for (const mode of ['', 'data-docx-compatibility-mode="14"']) {
  test(`unchanged legacy line choices preserve their measured word starts (${mode || 'omitted mode'})`, async ({ page }) => {
    await page.setContent(`${styles}<style>.run{white-space:pre-wrap;position:static;font-kerning:none}</style><main ${mode}>
      <p style="width:468pt;font-size:16pt;font-weight:bold"><span class="run">Service agreement layout</span></p>
      <p style="width:468pt;text-align:justify"><span class="run">The parties will inspect the sample, record the agreed date, and send the written schedule to the project office. Each report includes the measured quantity and the name of the person responsible for the next review.</span></p>
      </main>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const api = window as any, root = document.querySelector('main')!;
      api.applyUnroundedNormalLineHeights(root);
      const before = Array.from(root.querySelectorAll('p')).map(p => {
        const node = p.firstElementChild!.firstChild!, box = p.getBoundingClientRect();
        return Array.from(node.textContent!.matchAll(/\S+/g), m => {
          const r = document.createRange(); r.setStart(node, m.index!); r.setEnd(node, m.index! + m[0].length);
          return r.getBoundingClientRect().left - box.left;
        });
      });
      api.prepareWordText(root); api.paintWordText(root);
      const after = Array.from(root.querySelectorAll('p')).map(p =>
        Array.from(p.querySelectorAll('tspan')).map(w => Number(w.getAttribute('x')) / 4));
      return { before, after };
    });
    expect(actual.after).toEqual(actual.before);
  });
}

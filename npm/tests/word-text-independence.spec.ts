import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { wordTextScript } from './word-text-script.js';
import { fitTexts } from './docx-text-fit-fixture.js';

const metrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const justified = fitTexts.justify, kerned = fitTexts.left;
const neighbors = {
  plain: '<p><span>Closing note.</span></p>',
  bold: '<p><span>A short </span><span style="font-weight:bold">closing</span><span> note.</span></p>',
  table: '<table><tr><td><p><span>Cell text.</span></p></td></tr></table>',
  list: '<ol><li><p><span>List text.</span></p></li></ol>',
  link: '<p><a href="https://example.com">Link text.</a></p>',
};

for (const [name, neighbor] of Object.entries(neighbors)) for (const kind of ['justify', 'unkerned', 'kerned'] as const) {
  test(`${kind} line choices do not depend on ${name} content elsewhere`, async ({ page }) => {
    const text = kind === 'justify' ? justified : kerned;
    await page.setContent(`<style>p{font:12pt/normal 'Liberation Serif';margin:0;--docx-auto-line-spacing:1}
      span{white-space:pre-wrap}</style><main data-docx-compatibility-mode="15" style="width:468pt">
      <p id="target" style="text-align:${kind === 'justify' ? 'justify' : 'left'}"><span
        style="font-kerning:${kind === 'kerned' ? 'normal' : 'none'}">${text}</span></p>
      <div id="neighbor">${neighbor}</div></main>`);
    await page.addScriptTag({ content: metrics });
    await page.addScriptTag({ content: wordTextScript });
    const result = await page.evaluate(async () => {
      await document.fonts.ready;
      const root = document.querySelector('main')!, target = document.getElementById('target')!;
      const api = window as any;
      api.applyUnroundedNormalLineHeights(root);
      const neighbor = document.getElementById('neighbor')!, before = neighbor.innerHTML;
      api.prepareWordText(root);
      const rows = new Map<number, string[]>();
      const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const word of node.textContent!.matchAll(/\S+/g)) {
        const range = document.createRange();
        range.setStart(node, word.index!); range.setEnd(node, word.index! + word[0].length);
        const y = Math.round(range.getBoundingClientRect().top * 4);
        const row = rows.get(y) ?? []; row.push(word[0]); rows.set(y, row);
      }
      return { first: Array.from(rows.values())[0].join(' '), text: target.textContent,
        neighborPreserved: neighbor.innerHTML === before };
    });
    expect(result.first).toMatch(kind === 'justify' ? /pier as$/ : kind === 'kerned' ? /Where Water$/ : /Where$/);
    expect(result.text).toBe(text);
    if (name !== 'plain') expect(result.neighborPreserved).toBe(true);
  });
}

test('a plain RTL edge picture preserves its LTR neighbor eligibility', async ({ page }) => {
  await page.setContent(`<style>p{font:12pt/normal 'Liberation Serif';margin:0;--docx-auto-line-spacing:1}
    span{font:inherit;white-space:pre-wrap}</style><main style="width:468pt">
    <p id="ltr"><span>Plain left to right text.</span></p><p id="rtl" dir="rtl"><span><span
      data-docx-edge-picture="true" style="float:left;width:20px;height:20px"><img width="20" height="20"></span></span>
      <span>Logical text.</span></p></main>`);
  await page.addScriptTag({ content: metrics });
  await page.addScriptTag({ content: wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector('main')!, rtl = document.getElementById('rtl')!, api = window as any;
    api.applyUnroundedNormalLineHeights(root);
    const before = rtl.innerHTML;
    return { count: api.prepareWordText(root), rtlPreserved: before === rtl.innerHTML,
      marked: Array.from(root.querySelectorAll('[data-docx-word-text-prepared]'), p => p.id) };
  });
  expect(result).toEqual({ count: 1, rtlPreserved: true, marked: ['ltr'] });
});

for (const [name, owner, eligible] of [
  ['footnote', '<div id="pagination-footnote-registry"><div data-footnote-id="1"><span class="footnote-content">CONTENT</span></div></div>', false],
  ['margin comment', '<div id="pagination-comment-margin-registry"><div data-comment-id="1">CONTENT</div></div>', false],
  ['textbox', '<div data-docx-textbox="true">CONTENT</div>', false],
  ['paragraph frame', '<div data-docx-text-frame="true">CONTENT</div>', false],
  ['drawing', '<div data-docx-drawing-anchor="true">CONTENT</div>', false],
  ['running story', '<div id="pagination-hf-registry">CONTENT</div>', true],
] as const) test(`${name} retains its own text preparation policy`, async ({ page }) => {
  await page.setContent(`<style>p{font:12pt/normal 'Liberation Serif';--docx-auto-line-spacing:1}</style>
    <main style="width:468pt"><p><span>Ordinary body text.</span></p>
    ${owner.replace('CONTENT', '<p id="owned"><span>Separate story continuation.</span></p>')}</main>`);
  await page.addScriptTag({ content: metrics });
  await page.addScriptTag({ content: wordTextScript });
  const result = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!, owned = document.getElementById('owned')!;
    api.applyUnroundedNormalLineHeights(root);
    const before = owned.outerHTML;
    return { count: api.prepareWordText(root), preserved: before === owned.outerHTML };
  });
  expect(result).toEqual({ count: eligible ? 2 : 1, preserved: !eligible });
});

for (const exact of [false, true]) test(`${exact ? 'exact' : 'automatic'} text keeps its line and paint policy in a mixed document`, async ({ page }) => {
  const results = [];
  for (const mixed of [false, true]) {
    const paragraph = (id: string, fixed: boolean) => `<p id="${id}" style="${fixed
      ? 'line-height:20pt;--docx-exact-line-height:20pt;--docx-auto-line-spacing:0'
      : '--docx-auto-line-spacing:1'}"><span ${fixed ? 'data-docx-exact-run="true" style="line-height:20pt;vertical-align:top"' : ''}>${justified}</span></p>`;
    await page.setContent(`<style>p{font:12pt/normal 'Liberation Serif';text-align:justify;margin:0}
      span{white-space:pre-wrap}</style><main data-docx-compatibility-mode="15" style="width:468pt">
      ${paragraph('target', exact)}${mixed ? paragraph('neighbor', !exact) : ''}</main>`);
    await page.addScriptTag({ content: metrics });
    await page.addScriptTag({ content: wordTextScript });
    results.push(await page.evaluate(async () => {
      await document.fonts.ready;
      const api = window as any, root = document.querySelector('main')!, target = document.getElementById('target')!;
      api.applyUnroundedNormalLineHeights(root);
      api.prepareWordText(root);
      const before = target.getBoundingClientRect().height, html = target.innerHTML;
      api.paintWordText(root);
      return { before, after: target.getBoundingClientRect().height, html,
        svg: target.querySelectorAll('svg').length, text: target.textContent };
    }));
  }
  expect(results[1]).toEqual(results[0]);
  expect(results[0].after).toBe(results[0].before);
  expect(results[0].svg).toBe(exact ? 0 : 1);
  expect(results[0].text).toBe(justified);
});

for (const align of ['justify', 'left'] as const) test(`a ${align} mixed-run paragraph keeps its runs and text while taking Word's lines`, async ({ page }) => {
  await page.setContent(`<style>p{font:12pt/normal 'Liberation Serif';margin:0;--docx-auto-line-spacing:1}
    span{white-space:pre-wrap}</style><main data-docx-compatibility-mode="15" style="width:468pt">
    <p id="target" style="text-align:${align}"><span>${justified.slice(0, 13)}</span><span
      style="font-weight:bold">${justified.slice(13, 16)}</span><span>${justified.slice(16)}</span></p></main>`);
  await page.addScriptTag({ content: metrics });
  await page.addScriptTag({ content: wordTextScript });
  const result = await page.evaluate(async () => {
    await document.fonts.ready;
    const root = document.querySelector('main')!, target = document.getElementById('target')!, api = window as any;
    api.applyUnroundedNormalLineHeights(root);
    api.prepareWordText(root);
    const words: Array<{ text: string; top: number }> = [];
    const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const word of node.textContent!.matchAll(/\S+/g)) {
      const range = document.createRange();
      range.setStart(node, word.index!); range.setEnd(node, word.index! + word[0].length);
      words.push({ text: word[0], top: Math.round(range.getBoundingClientRect().top) });
    }
    const condensed = Array.from(target.querySelectorAll<HTMLElement>('[data-docx-condensed-line]'));
    return { first: words.filter(word => word.top === words[0].top).map(word => word.text).join(' '),
      text: target.textContent, runs: Array.from(target.children, run => run.getAttribute('style') ?? ''),
      tops: [...new Set(condensed.map(line => Math.round(line.getBoundingClientRect().top)))], top: words[0].top,
      spacing: condensed.map(line => Number.parseFloat(getComputedStyle(line).wordSpacing)) };
  });
  expect(result.first).toMatch(align === 'justify' ? /pier as$/ : /pier$/);
  expect(result.text).toBe(justified);
  expect(result.runs).toHaveLength(3);
  expect(result.runs[1]).toContain('font-weight: bold');
  // Only the line Word condenses narrows its ordinary spaces.
  expect(result.tops).toEqual(align === 'justify' ? [result.top] : []);
  for (const spacing of result.spacing) expect(spacing).toBeLessThan(0);
});

test('a condensed line keeps the Word baseline of its uniform runs under an enlarged minimum', async ({ page }) => {
  const paragraph = (id: string, align: string) => `<p id="${id}" style="text-align:${align};line-height:24pt;
    --docx-min-line-height:24pt"><span>${justified.slice(0, 60)}</span><span>${justified.slice(60)}</span></p>`;
  await page.setContent(`<style>p{font:12pt 'Liberation Serif';margin:0}span{white-space:pre-wrap}</style>
    <main data-docx-compatibility-mode="15" style="width:468pt">${paragraph('justified', 'justify')}${paragraph('left', 'left')}</main>`);
  await page.addScriptTag({ content: metrics });
  await page.addScriptTag({ content: wordTextScript });
  const result = await page.evaluate(async () => {
    await document.fonts.ready;
    const api = window as any, root = document.querySelector('main')!;
    api.prepareWordText(root);
    api.alignBaselinesToWord(root);
    return ['justified', 'left'].map(id => {
      const p = document.getElementById(id)!;
      return { condensed: p.querySelectorAll('[data-docx-condensed-line]').length,
        tops: Array.from(p.children, run => getComputedStyle(run).top) };
    });
  });
  expect(result[0].condensed).toBeGreaterThan(0);
  expect(result[1].condensed).toBe(0);
  expect(result[0].tops).toEqual(result[1].tops);
  expect(result[0].tops[0]).not.toBe('auto');
});

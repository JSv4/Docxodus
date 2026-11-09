import { expect, test } from '@playwright/test';
import { wordTextScript } from './word-text-script.js';

const prefix = 'Copper lanterns swing above the quiet harbor wall while gulls nap';
const style = `p{font:12pt 'Liberation Mono';line-height:18px;margin:0;--docx-auto-line-spacing:1}
  span{white-space:pre-wrap;line-height:18px;--docx-font-kerning:none}`;

test('fixed cells distinguish an exact fit from a twip-narrower band and an extra character', async ({page}) => {
  await page.setContent(`<style>${style}</style><main data-docx-compatibility-mode="14">
    <p style="width:468pt"><span>${prefix} TAIL</span></p>
    <p style="width:467.95pt"><span>${prefix} TAIL</span></p>
    <p style="width:468pt"><span>${prefix}s TAIL</span></p></main>`);
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({content:wordTextScript});
  const lines = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!;
    api.prepareWordText(root); api.paintWordText(root);
    return Array.from(root.querySelectorAll('p'), p => p.querySelector('svg text')!.textContent!.trim());
  });
  expect(lines).toEqual([prefix, prefix.slice(0, -4), prefix.slice(0, -4)]);
});

test('fixed-cell measurement precedes authored scaling and retains raw glyph dimensions', async ({page}) => {
  await page.setContent(`<style>${style}p{font-size:11pt}</style><main data-docx-compatibility-mode="14">
    ${[.8, 1, 1.5].map(scale => `<p style="width:${429 * scale}pt"><span style="--docx-character-scale:${scale}">${prefix} TAIL</span></p>`).join('')}
    </main>`);
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({content:wordTextScript});
  const actual = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!;
    api.prepareWordText(root); api.paintWordText(root);
    return Array.from(root.querySelectorAll('p'), p => {
      const line = p.querySelector('svg text')!, spans = Array.from(line.querySelectorAll('tspan'));
      return {text:line.textContent!.trim(), secondX:(spans[1].getBoundingClientRect().left - p.getBoundingClientRect().left) * .75,
        glyphs:spans.map(span => {
          const text = span.textContent!.trimEnd(), range = document.createRange();
          range.setStart(span.firstChild!, 0); range.setEnd(span.firstChild!, text.length);
          return {text, width:range.getBoundingClientRect().width * .75};
        })};
    });
  });
  for (const [i, scale] of [.8, 1, 1.5].entries()) {
    expect(actual[i].text).toBe(prefix);
    expect(actual[i].secondX).toBeCloseTo(46.2 * scale, 2);
    for (const glyph of actual[i].glyphs)
      expect(glyph.width).toBeCloseTo(glyph.text.length * 1229 / 2048 * 11 * scale, 1);
  }
});

test('fallback glyphs and an explicitly kerned run retain measured advances', async ({page}) => {
  await page.setContent(`<style>${style}
    @font-face{font-family:PartialMono;src:local('Liberation Mono');unicode-range:U+20-59,U+5B-7E}
    </style><main data-docx-compatibility-mode="14">
    <p style="width:79.2pt;font-family:PartialMono,'Liberation Sans'"><span>Copper ZZZZ TAIL</span></p>
    <p style="width:468pt"><span style="--docx-font-kerning:normal">${prefix} TAIL</span></p></main>`);
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({content:wordTextScript});
  const lines = await page.evaluate(() => {
    const api = window as any, root = document.querySelector('main')!;
    api.prepareWordText(root); api.paintWordText(root);
    return Array.from(root.querySelectorAll('p'), p => p.querySelector('svg text')!.textContent!.trim());
  });
  expect(lines).toEqual(['Copper', prefix.slice(0, -4)]);
});

for (const wrap of ['normal', 'anywhere']) test(`a raw overwide unbreakable word with ${wrap} wrapping cannot gain an unsafe owned line`, async ({page}) => {
  const text = 'X'.repeat(65) + ' tail';
  await page.setContent(`<style>${style}p{width:468pt;overflow-wrap:${wrap}}</style>
    <main data-docx-compatibility-mode="14"><p><span>${text}</span></p></main>`);
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({content:wordTextScript});
  const result = await page.evaluate(() => {
    const root = document.querySelector('main')!;
    (window as any).prepareWordText(root);
    return {text:root.textContent, fitted:root.querySelector('[data-docx-word-fit]') !== null,
      lines:root.querySelectorAll('[data-docx-word-line]').length};
  });
  expect(result).toEqual({text, fitted:false, lines:0});
});

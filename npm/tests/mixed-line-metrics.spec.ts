import { expect, test, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

const metrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const natural = (points: number) => 2355 / 2048 * points * 4 / 3;
const words = Array.from({ length: 36 }, (_, index) => `Iris${index + 1}`).join(' ');
const marker = '<span data-list-marker="true" style="display:inline-flex;align-items:baseline;width:18pt;text-indent:0">'
  + '<span style="flex-shrink:0;white-space:nowrap"><span data-list-marker="true">1.</span></span>'
  + '<span data-list-marker="true" style="flex:1"><span data-docx-tab="left" '
  + 'style="display:inline-block;width:100%;height:0"></span></span></span>';

// These are the converter's unresolved styles, before the shared natural-height pass:
// a largest-run paragraph font, explicit run sizes, and either normal, a fixed
// atLeast height, or the converter's inherited 1lh multiplier and matching lift.
function paragraph(options: {
  width?: number; paragraphStyle?: string; firstStyle?: string; secondStyle?: string;
  secondAttributes?: string; text?: string; multiple?: number; minimum?: number;
} = {}): string {
  const { width = 180, multiple = 1, minimum = 0 } = options;
  const policy = minimum ? `--docx-min-line-height:${minimum}pt;line-height:${minimum >= 14 ? `${minimum}pt` : 'normal'}`
    : `--docx-auto-line-spacing:${multiple}`;
  const automatic = !minimum && multiple !== 1
    ? 'line-height:calc(1lh * var(--docx-auto-line-spacing));position:relative;'
      + 'top:min(0px,calc(1lh * (1 / var(--docx-auto-line-spacing) - 1) / 2));' : '';
  return `<section data-section-index="0"><p id="target" data-source-anchor-id="p:body:mixed" data-anchor="mixed"
    style="margin:0;font:20pt/normal 'Liberation Serif';width:${width}px;${policy};${options.paragraphStyle ?? ''}">`
    + `<span data-docx-font-request="font-0001" style="font-size:20pt;${automatic}${options.firstStyle ?? ''}">Big </span>`
    + `<span data-docx-font-request="font-0001" ${options.secondAttributes ?? ''}
      style="font-size:10pt;${automatic}${options.secondStyle ?? ''}">${options.text ?? words}</span></p></section>`;
}

async function exercise(page: Page, markup: string, widths: number[] = []) {
  await page.setContent(markup);
  await page.addScriptTag({ content: metrics });
  await page.evaluate(() => document.fonts.ready);
  return page.evaluate(widths => {
    const api = window as unknown as {
      prepareMixedLineHeights(root: HTMLElement): number;
      applyUnroundedNormalLineHeights(root: Element): number;
      isMixedLineMetric(element: HTMLElement): boolean;
    };
    const paragraph = document.querySelector<HTMLElement>('#target')!;
    const pristine = paragraph.cloneNode(true) as HTMLElement;
    const snapshot = (p: HTMLElement) => {
      const box = p.getBoundingClientRect();
      const tokens: Array<{ text: string; x: number; y: number; width: number; height: number; rects: number }> = [];
      const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
        for (const match of node.data.matchAll(/\S+/gu)) {
          const range = document.createRange();
          range.setStart(node, match.index); range.setEnd(node, match.index + match[0].length);
          const rect = range.getBoundingClientRect();
          tokens.push({ text: match[0], x: rect.left - box.left, y: rect.top - box.top,
            width: rect.width, height: rect.height, rects: range.getClientRects().length });
        }
      }
      return { html: p.outerHTML, text: p.textContent, height: box.height, tokens,
        anchor: p.getAttribute('data-source-anchor-id'), editorAnchor: p.getAttribute('data-anchor'),
        struts: Array.from(p.querySelectorAll<HTMLElement>('[data-docx-line-metric]')).map(strut => {
          const rect = strut.getBoundingClientRect();
          return { top: rect.top - box.top, height: rect.height, valid: api.isMixedLineMetric(strut) };
        }) };
    };
    const before = snapshot(paragraph);
    const count = api.prepareMixedLineHeights(paragraph);
    const after = snapshot(paragraph);
    const repeat = api.prepareMixedLineHeights(paragraph);
    const repeated = snapshot(paragraph);
    api.applyUnroundedNormalLineHeights(paragraph);
    const afterNaturalPass = snapshot(paragraph);
    const clone = paragraph.cloneNode(true) as HTMLElement;
    clone.removeAttribute('id'); paragraph.parentElement!.append(clone);
    const cloneCount = api.prepareMixedLineHeights(clone);
    const cloned = snapshot(clone); clone.remove();
    const resized = widths.map(width => {
      paragraph.style.width = `${width}px`;
      const count = api.prepareMixedLineHeights(paragraph), actual = snapshot(paragraph);
      const fresh = pristine.cloneNode(true) as HTMLElement;
      fresh.removeAttribute('id'); fresh.style.width = `${width}px`;
      paragraph.parentElement!.append(fresh);
      const freshCount = api.prepareMixedLineHeights(fresh), expected = snapshot(fresh);
      fresh.remove();
      return { width, count, freshCount, actual, expected };
    });
    return { count, repeat, cloneCount, before, after, repeated, afterNaturalPass, cloned, resized };
  }, widths);
}

function horizontal(tokens: Array<{ text: string; x: number; width: number }>) {
  return tokens.map(({ text, x, width }) => ({ text, x, width }));
}

test('equal vertical metrics preserve styled text for ordinary horizontal fitting', async ({ page }) => {
  const result = await exercise(page, paragraph({ paragraphStyle: 'font-size:10pt',
    firstStyle: 'font-size:10pt;font-weight:bold' }));
  expect(result.count).toBe(0);
  expect(result.after.html).toBe(result.before.html);
});

test('one-word visual lines retain every line instead of collapsing equal x coordinates', async ({ page }) => {
  const result = await exercise(page, paragraph({ width: 42, text: Array(12).fill('word').join(' ') }));
  expect(result.count).toBe(1);
  expect(result.after.struts).toHaveLength(13);
  expect(new Set(result.after.tokens.map(token => token.y)).size).toBe(13);
  expect(Math.abs(result.after.height - natural(20) - 12 * natural(10))).toBeLessThan(.25);
  expect(horizontal(result.after.tokens)).toEqual(horizontal(result.before.tokens));
  for (let index = 1; index < result.after.tokens.length; index++)
    expect(result.after.tokens[index].y).toBeGreaterThan(result.after.tokens[index - 1].y);
  expect(result.after.text).toBe(result.before.text);
});

test('a token occupying multiple rectangles is refused without changing its source', async ({ page }) => {
  const result = await exercise(page, paragraph({ width: 42, paragraphStyle: 'overflow-wrap:anywhere',
    text: 'ExtraordinarilyLongUnbrokenWord' }));
  expect(result.before.tokens.some(token => token.rects > 1)).toBe(true);
  expect(result.count).toBe(0);
  expect(result.after.html).toBe(result.before.html);
});

for (const [name, options] of [
  ['paragraph height override', { minimum: 16, paragraphStyle: 'line-height:40px' }],
  ['run height override', { secondStyle: 'line-height:50px' }],
  ['relative left offset', { multiple: 1.15, firstStyle: 'left:17px' }],
  ['relative right offset', { multiple: 1.15, firstStyle: 'right:17px' }],
  ['relative bottom offset', { multiple: 1.15, firstStyle: 'bottom:3px' }],
  ['unowned negative lift', { multiple: 1.15, firstStyle: 'top:-2px' }],
  ['dynamic field', { secondAttributes: 'data-field="PAGE"' }],
  // Off a grid; on one, the multiple is whole grid lines (below).
  ['multiple below one', { multiple: .8 }],
] as const) test(`mixed metrics preserve the ${name}`, async ({ page }) => {
  const result = await exercise(page, paragraph(options));
  expect(result.count).toBe(0);
  expect(result.after.html).toBe(result.before.html);
  expect(result.after.tokens).toEqual(result.before.tokens);
});

for (const [owner, attribute, anchor] of [
  ['footnote', 'data-footnote-id="1"', 'p:fn:mixed'],
  ['comment', 'data-comment-id="1"', 'p:comment:mixed'],
] as const) test(`mixed body metrics leave the actual ${owner} owner unchanged`, async ({ page }) => {
  const source = paragraph().replace('p:body:mixed', anchor)
    .replace('<p id="target"', `<div ${attribute}><p id="target"`).replace('</p></section>', '</p></div></section>');
  const result = await exercise(page, source);
  expect(result.count).toBe(0);
  expect(result.after.html).toBe(result.before.html);
});

for (const minimum of [13, 16]) test(`the ${minimum}pt floor is applied once to each mixed visual line`, async ({ page }) => {
  const result = await exercise(page, paragraph({ minimum }));
  expect(result.count).toBe(1);
  expect(result.after.struts.length).toBeGreaterThan(2);
  const expected = natural(20) + (result.after.struts.length - 1) * minimum * 4 / 3;
  expect(Math.abs(result.after.height - expected)).toBeLessThan(.25);
  expect(horizontal(result.after.tokens)).toEqual(horizontal(result.before.tokens));
  expect(result.after.struts.every(strut => strut.valid)).toBe(true);
});

test('the converter automatic lift is admitted without changing horizontal glyph geometry', async ({ page }) => {
  const result = await exercise(page, paragraph({ multiple: 1.15 }));
  expect(result.count).toBe(1);
  const expected = (natural(20) + (result.after.struts.length - 1) * natural(10)) * 1.15;
  expect(Math.abs(result.after.height - expected)).toBeLessThan(.25);
  expect(horizontal(result.after.tokens)).toEqual(horizontal(result.before.tokens));
});

function numbered(prefix = marker) {
  return '<section data-section-index="0"><p id="target" data-source-anchor-id="li:body:mixed" data-anchor="mixed" '
    + 'style="font:18pt/normal Liberation Serif;--docx-auto-line-spacing:1;width:240px;margin:0 0 0 36pt;text-indent:-18pt">'
    + prefix + `<span data-docx-font-request="font-0001" style="font-size:12pt">${words}</span></p></section>`;
}

test('a converter marker raises only its first line ascent and retains its identity', async ({ page }) => {
  const result = await exercise(page, numbered());
  expect(result.count).toBe(1);
  expect(result.after.struts.length).toBeGreaterThan(2);
  // Word (#1565): the 18pt number's ascent with line gap over the 12pt text's descent.
  const markerLine = (1912 / 2048 * 18 + 443 / 2048 * 12) * 4 / 3;
  expect(Math.abs(result.after.struts[0].height - markerLine)).toBeLessThan(1 / 32);
  for (const strut of result.after.struts.slice(1)) expect(Math.abs(strut.height - natural(12))).toBeLessThan(1 / 32);
  expect(horizontal(result.after.tokens)).toEqual(horizontal(result.before.tokens));
  expect(result.after.tokens.filter(token => token.text === '1.')).toHaveLength(1);
  expect(result.after.anchor).toBe('li:body:mixed');
});

for (const [name, prefix] of [
  ['block marker', marker.replace('display:inline-flex', 'display:block')],
  ['floated marker child', marker.replace('flex-shrink:0', 'float:left;flex-shrink:0')],
  ['raised marker text', marker.replace('<span data-list-marker="true">1.', '<span data-list-marker="true" style="vertical-align:super">1.')],
] as const) test(`unowned ${name} layout remains unchanged`, async ({ page }) => {
  const result = await exercise(page, numbered(prefix));
  expect(result.count).toBe(0);
  expect(result.after.html).toBe(result.before.html);
});

test('repeated and cloned preparation preserve source anchors and resolved line metrics', async ({ page }) => {
  const result = await exercise(page, paragraph({ minimum: 16 }));
  expect(result.count).toBe(1);
  expect(result.repeat).toBe(0); expect(result.cloneCount).toBe(0);
  expect(result.repeated).toEqual(result.after);
  expect(result.afterNaturalPass).toEqual(result.after);
  expect(result.cloned.tokens).toEqual(result.after.tokens);
  expect(result.cloned.struts).toEqual(result.after.struts);
  expect(result.after.anchor).toBe('p:body:mixed');
  expect(result.after.editorAnchor).toBe('mixed');
  expect(result.after.text).toBe(result.before.text);
});

test('changed-width preparation matches fresh source at both narrower and wider widths', async ({ page }) => {
  const result = await exercise(page, paragraph({ minimum: 16 }), [72, 240, 180]);
  expect(result.count).toBe(1);
  for (const { count, freshCount, actual, expected } of result.resized) {
    expect(count).toBe(1); expect(freshCount).toBe(1);
    expect(actual.text).toBe(result.before.text);
    expect(actual.anchor).toBe('p:body:mixed');
    expect(actual.tokens).toEqual(expected.tokens);
    expect(actual.struts).toEqual(expected.struts);
    expect(actual.height).toBe(expected.height);
  }
  expect(result.resized[0].actual.struts.length).toBeGreaterThan(result.resized[1].actual.struts.length);
  expect(result.resized[2].actual.tokens).toEqual(result.after.tokens);
});

// Word 2024 (#1551) gives each visual line of mixed text the whole grid lines its text needs, with its text
// centred in them. On a 36pt grid, that puts a 10pt line's text lower than its own height below the baseline.
// A multiple below one changes nothing on a grid (a 24pt word at 0.8 lines on an 18pt grid: 36pt, as single).
for (const [pitch, multiple] of [[18, 1], [36, 1], [18, .8]])
test(`mixed visual lines take whole ${pitch}pt grid lines${multiple !== 1 ? ` at ${multiple} lines` : ''}`, async ({ page }) => {
  const result = await exercise(page, paragraph({ width: 42, text: Array(12).fill('word').join(' '), multiple,
    paragraphStyle: `--docx-line-grid:${pitch}pt` }));
  expect(result.count).toBe(1);
  expect(result.after.struts.map(strut => strut.height / (pitch * 4 / 3))).toEqual([pitch === 18 ? 2 : 1, ...Array(12).fill(1)]);
  expect(result.after.struts.every(strut => strut.valid)).toBe(true);
  expect(Math.abs(result.after.height - result.after.struts.reduce((sum, strut) => sum + strut.height, 0))).toBeLessThan(.25);
  expect(horizontal(result.after.tokens)).toEqual(horizontal(result.before.tokens));
});

// On an 18pt grid the 20pt line takes two grid lines and each 10pt line one, so 4, 5 and 4 lines fill 100pt pages.
for (const [viewer, pitch, lines] of [[false, 0, [7, 6]], [true, 0, [7, 6]], [false, 18, [4, 5, 4]], [true, 18, [4, 5, 4]]] as const)
test(`the ${viewer ? 'viewer' : 'editable engine'} paginates mixed lines${pitch ? ` on a ${pitch}pt grid` : ''} with complete source identities`, async ({ page }) => {
  const source = paragraph({ width: 42, text: Array(12).fill('word').join(' '), paragraphStyle: pitch ? `--docx-line-grid:${pitch}pt` : '' });
  await page.setContent('<div id="staging" style="display:none">' + source.replace('<section data-section-index="0">',
    '<section data-section-index="0" data-page-width="175.5" data-page-height="244" '
      + 'data-content-width="31.5" data-content-height="100" data-margin-top="72" '
      + 'data-margin-right="72" data-margin-bottom="72" data-margin-left="72">')
    + '</div><div id="pages"></div>');
  await page.evaluate(() => document.fonts.ready);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const result = await page.evaluate(viewer => {
    const pages = document.getElementById('pages')!;
    const engine = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), pages, {
      scale: 1, showPageNumbers: false, fragmentParagraphs: true, alignBaselinesToWord: viewer,
    });
    engine.paginate();
    const map = engine.materializePageMap(0, 'mixed-line-metrics') as {
      pages: Array<{ pageNumber: number }>;
      fragments: Array<{ fragmentId: string; anchorId: string; pageNumber: number; fragmentIndex: number }>;
    };
    const fragments = Array.from(pages.querySelectorAll<HTMLElement>('.page-content p'));
    return { pageNumbers: map.pages.map(p => p.pageNumber), map: map.fragments,
      text: fragments.map(p => p.textContent).join(' ').replace(/\s+/g, ' ').trim(),
      lines: fragments.map(p => p.querySelectorAll('[data-docx-line-metric]').length),
      sourceAnchors: fragments.map(p => p.dataset.sourceAnchorId),
      editorAnchors: pages.querySelectorAll('[data-anchor="mixed"]').length,
      htmlIds: pages.querySelectorAll('#target').length };
  }, viewer);
  const pageNumbers = lines.map((_, index) => index + 1);
  expect(result.pageNumbers).toEqual(pageNumbers);
  expect(result.lines).toEqual(lines);
  expect(result.text).toBe('Big ' + Array(12).fill('word').join(' '));
  expect(result.sourceAnchors).toEqual(lines.map(() => 'p:body:mixed'));
  expect(result.editorAnchors).toBe(1); expect(result.htmlIds).toBe(1);
  expect(result.map.map(fragment => fragment.anchorId)).toEqual(result.sourceAnchors);
  expect(result.map.map(fragment => fragment.pageNumber)).toEqual(pageNumbers);
  expect(result.map.map(fragment => fragment.fragmentIndex)).toEqual(pageNumbers.map(number => number - 1));
  expect(new Set(result.map.map(fragment => fragment.fragmentId)).size).toBe(lines.length);
});

import { test, expect } from '@playwright/test';
import * as path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.dirname(fileURLToPath(import.meta.url));
const geometry = `data-section-index="0" data-page-width="612" data-page-height="792"
  data-content-width="468" data-content-height="648" data-margin-top="72"
  data-margin-bottom="72" data-margin-left="72" data-margin-right="72"`;

test('column breaks retain both sides of the canonical paragraph in PageMap', async ({ page }) => {
  await page.setContent(`<style>p{font:11pt/14pt serif;margin:0}</style>
    <div id="staging"><section ${geometry} data-section-type="continuous" data-cols="2" data-col-gap="24"
      style="column-count:2;column-gap:24pt">
      <p data-source-anchor-id="p:body:lead">Left column</p>
      <p data-source-anchor-id="p:body:split">Before break</p>
      <p data-source-anchor-id="p:body:split" data-column-break-before="true"
        style="break-before:column">After break</p>
    </section></div><div id="container"></div>`);
  await page.addScriptTag({ path: path.join(root, '../dist/pagination.bundle.js') });
  const result = await page.evaluate(() => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine(
      document.getElementById('staging'), document.getElementById('container'), { showPageNumbers: false });
    engine.paginate();
    return { map: engine.materializePageMap(0, 'column-break-test'),
      text: document.getElementById('container')!.textContent };
  });
  expect(result.map.pages).toHaveLength(1);
  const split = result.map.fragments.filter((f: any) => f.anchorId === 'p:body:split');
  expect(split).toHaveLength(2);
  expect(split[0].geometry.x).toBeCloseTo(72, 0);
  expect(split[1].geometry.x).toBeCloseTo(318, 0);
  expect(split.map((f: any) => f.fragmentIndex)).toEqual([0, 1]);
  expect(result.text).toContain('Before break');
  expect(result.text).toContain('After break');
});

test('flattened print endnotes retain their marker without list punctuation', async ({ page }) => {
  await page.setContent(`<style>p{font:11pt/14pt serif;margin:0}ol{padding:0;margin:0}
    section.endnotes hr{width:144pt;height:1px;border:0;background:black;margin:0 0 4pt 0}</style>
    <div id="staging"><section ${geometry}>
      <p data-source-anchor-id="p:body:lead">Body</p>
      <section class="endnotes"><hr><ol><li id="en-1" value="1" data-source-anchor-id="en:en:note">
        <p data-source-anchor-id="p:en:paragraph"><span data-source-anchor-id="en:en:note"><span data-endnote-number="true">i</span> Note text</span></p>
      </li></ol></section>
    </section></div><div id="container"></div>`);
  await page.addScriptTag({ path: path.join(root, '../dist/pagination.bundle.js') });
  const result = await page.evaluate(() => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine(
      document.getElementById('staging'), document.getElementById('container'), { showPageNumbers: false });
    engine.paginate();
    const rule = document.querySelector('#container hr') as HTMLElement;
    return { text: document.getElementById('container')!.textContent,
      ruleWidth: rule.getBoundingClientRect().width * 0.75,
      ruleBorder: getComputedStyle(rule).borderTopWidth,
      map: engine.materializePageMap(0, 'endnote-test') };
  });
  expect(result.text).toContain('i Note text');
  expect(result.text).not.toContain('1.');
  expect(result.text).not.toContain('i.');
  expect(result.ruleWidth).toBeCloseTo(144, 1);
  expect(result.ruleBorder).toBe('0px');
  expect(result.map.fragments.some((f: any) => f.anchorId === 'en:en:note')).toBe(true);
});

for (const source of ['class', 'list', 'inline']) {
  test(`flattened endnotes preserve ${source} typography and glyph baselines`, async ({ page }) => {
    const size = source === 'list' ? '9pt' : '11pt';
    const declaration = `font-size:${size};font-family:'Liberation Serif',serif`;
    const content = `<span data-source-anchor-id="en:en:note"><span data-endnote-number="true">i</span><span class="ink"> Note text</span></span>`;
    await page.setContent(`<style>
      p{line-height:normal;margin:0;--docx-auto-line-spacing:1.15}
      ol{padding:0;margin:0}.ink{font-size:${size};line-height:calc(1lh * 1.15);
        position:relative;top:min(0px,calc(1lh * (1 / 1.15 - 1) / 2))}
      ${source === 'class' ? `.note{${declaration}}` : ''}
      </style><div id="staging"><section ${geometry}>
      <p id="control" style="${declaration}">${content.replace('en:en:note', 'p:body:control')}</p>
      <section class="endnotes"><ol><li id="en-1" ${source === 'list' ? `style="${declaration}"` : ''}>
        <p class="note" data-source-anchor-id="p:en:paragraph" ${source === 'inline' ? `style="${declaration}"` : ''}>${content}</p>
      </li></ol></section></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: path.join(root, '../dist/pagination.bundle.js') });
    const result = await page.evaluate(() => {
      const engine = new (window as any).DocxodusPagination.PaginationEngine(
        document.getElementById('staging'), document.getElementById('container'),
        { showPageNumbers: false, fragmentParagraphs: true, alignBaselinesToWord: true });
      engine.paginate();
      const measure = (p: HTMLElement) => {
        const top = p.getBoundingClientRect().top;
        const style = getComputedStyle(p);
        return { font: style.fontSize, family: style.fontFamily,
          text: p.querySelector('.ink')!.getBoundingClientRect().top - top,
          marker: p.querySelector('[data-endnote-number]')!.getBoundingClientRect().top - top };
      };
      return { control: measure(document.querySelector('#container #control')!),
        note: measure(document.querySelector('#container [data-pagination-safe-endnote]')!),
        nested: document.querySelectorAll('#container section.endnotes').length };
    });
    expect(result.nested).toBe(0);
    expect(result.note.font).toBe(result.control.font);
    expect(result.note.family).toBe(result.control.family);
    expect(result.note.text).toBeCloseTo(result.control.text, 2);
    expect(result.note.marker).toBeCloseTo(result.control.marker, 2);
  });
}

test('automatically spaced endnotes still split with their authored font', async ({ page }) => {
  const text = Array.from({ length: 180 }, (_, i) => `Word${i}`).join(' ');
  await page.setContent(`<style>
    p.note{font:11pt/normal 'Liberation Serif',serif;margin:0;--docx-auto-line-spacing:1.15}
    ol{padding:0;margin:0}.ink{line-height:calc(1lh * 1.15);position:relative;
      top:min(0px,calc(1lh * (1 / 1.15 - 1) / 2))}
    </style><div id="staging"><section ${geometry.replace('data-content-height="648"', 'data-content-height="80"').replace('data-page-height="792"', 'data-page-height="224"')}>
    <section class="endnotes"><ol><li id="en-1" data-source-anchor-id="en:en:note">
      <p class="note" data-source-anchor-id="p:en:paragraph"><span data-source-anchor-id="en:en:note"><span data-endnote-number="true">i</span><span class="ink"> ${text}</span></span></p>
    </li></ol></section></section></div><div id="container"></div>`);
  await page.addScriptTag({ path: path.join(root, '../dist/pagination.bundle.js') });
  const result = await page.evaluate(() => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine(
      document.getElementById('staging'), document.getElementById('container'),
      { showPageNumbers: false, fragmentParagraphs: true, alignBaselinesToWord: true });
    engine.paginate();
    const notes = Array.from(document.querySelectorAll<HTMLElement>('#container [data-pagination-safe-endnote]'));
    return { fonts: notes.map(p => getComputedStyle(p).fontSize),
      words: notes.map(p => p.textContent).join(' ').match(/Word\d+/g),
      map: engine.materializePageMap(0, 'endnote-font-test') };
  });
  expect(result.fonts.length).toBeGreaterThan(1);
  expect(result.fonts.every(font => Math.abs(parseFloat(font) - 44 / 3) < .001)).toBe(true);
  expect(result.words).toEqual(text.split(' '));
  expect(result.map.fragments.filter((f: any) => f.anchorId === 'p:en:paragraph').length).toBeGreaterThan(1);
});

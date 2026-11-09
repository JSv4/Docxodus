import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

import { wordTextScript as wordText } from './word-text-script.js';
const lineMetrics = readFileSync(new URL('../dist/line-metrics.js', import.meta.url), 'utf8').replace(/^export /gm, '');
const observations = JSON.parse(readFileSync(new URL('../../TestFiles/JustificationFlow/word-observations.json', import.meta.url), 'utf8'));
const escape = (text: string) => text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

for (const source of observations.cases) {
  test(`${source.id} retains every independently captured native line`, async ({ page }) => {
    const bytes = readFileSync(new URL(`../../TestFiles/JustificationFlow/${source.file}`, import.meta.url));
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(source.inputSha256);
    // Liberation Serif's Latin advances match these Times New Roman probes. The
    // separate public-export checks use the exact declared Times New Roman font.
    await page.setContent(`<style>p{font:${source.fontSizePt}pt 'Liberation Serif',serif;line-height:normal;
      margin:0 0 6pt;--docx-auto-line-spacing:1.15} .run{white-space:pre-wrap;position:static;
      line-height:calc(1lh * 1.15);--docx-font-kerning:none}</style><main data-docx-compatibility-mode="15">
      ${source.paragraphs.map((p: any, i: number) => `<p data-source-anchor-id="p:body:${i}"
        style="width:${p.widthPt}pt;text-indent:${p.firstLineIndentPt}pt;text-align:justify"><span class="run">${escape(p.text)}</span></p>`).join('')}</main>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const api = window as any, root = document.querySelector('main')!;
      const before = Array.from(root.querySelectorAll('p')).map(p => ({ text:p.textContent, id:p.dataset.sourceAnchorId }));
      api.applyUnroundedNormalLineHeights(root);
      const prepared = api.prepareWordText(root);
      const boxes = Array.from(root.querySelectorAll('p')).map(p => ({
        height:p.getBoundingClientRect().height,
        pitch:Number.parseFloat(getComputedStyle(p.querySelector('.run')!).lineHeight) }));
      const painted = api.paintWordText(root);
      return { prepared, painted, before, boxes,
        after:Array.from(root.querySelectorAll('p')).map(p => ({ text:p.textContent, id:p.dataset.sourceAnchorId })),
        lines:Array.from(root.querySelectorAll('p')).map(p => Array.from(p.querySelectorAll('svg > text')).map(t => t.textContent!.trim().split(/\s+/))) };
    });
    expect(actual.prepared).toBe(source.paragraphs.length);
    expect(actual.painted).toBe(source.paragraphs.length);
    expect(actual.after).toEqual(actual.before);
    expect(actual.lines).toEqual(source.paragraphs.map((p: any) => p.nativeLines));
    for (let index = 0; index < actual.boxes.length; index++)
      expect(Math.abs(actual.boxes[index].height - actual.boxes[index].pitch * source.paragraphs[index].nativeLines.length)).toBeLessThan(1 / 32);
  });
}

for (const position of ['static', 'relative']) {
for (const source of observations.cases) {
  test(`${source.id} retains native page membership through paragraph continuation (${position})`, async ({ page }) => {
    await page.setContent(`<style>p{font:${source.fontSizePt}pt 'Liberation Serif',serif;line-height:normal;
      margin:0 0 6pt;--docx-auto-line-spacing:1.15} .run{white-space:pre-wrap;position:${position};
      top:min(0px,calc(1lh * (1/1.15 - 1)/2));
      line-height:calc(1lh * 1.15);--docx-font-kerning:none}</style><div id="staging" data-docx-compatibility-mode="15"><div data-section-index="0"
      data-page-width="612" data-page-height="792" data-content-width="468" data-content-height="648"
      data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72">
      ${source.paragraphs.map((p: any, i: number) => `<p data-source-anchor-id="p:body:${i.toString(16).padStart(32, '0')}"
        data-widow-control="false" style="width:${p.widthPt}pt;text-indent:${p.firstLineIndentPt}pt;text-align:justify"><span class="run">${escape(p.text)}</span></p>`).join('')}
      </div></div><div id="container"></div>`);
    await page.addScriptTag({ content: lineMetrics });
    await page.addScriptTag({ content: wordText });
    await page.addScriptTag({ path: new URL('../dist/pagination.bundle.js', import.meta.url).pathname });
    const actual = await page.evaluate(count => {
      const api = window as any, root = document.getElementById('staging')!, container = document.getElementById('container')!;
      const text = Array.from(root.querySelectorAll('p')).map(p => p.textContent).join('');
      api.applyUnroundedNormalLineHeights(root); api.prepareWordText(root);
      const engine = new api.DocxodusPagination.PaginationEngine(root, container, { showPageNumbers:false, fragmentParagraphs:true });
      const result = engine.paginate(), before = engine.materializePageMap(0, 'native-page-fit');
      const painted = api.paintWordText(container);
      const paragraphs = Array.from(container.querySelectorAll('p'));
      const lines: Array<Array<{page:number;words:string[]}>> = Array.from({length:count}, () => []);
      for (const p of paragraphs) {
        const index = Number.parseInt(p.dataset.sourceAnchorId!.split(':').at(-1)!, 16);
        const page = Number(p.closest<HTMLElement>('.page-box')!.dataset.pageNumber);
        for (const line of Array.from(p.querySelectorAll('svg > text')))
          lines[index].push({page, words:line.textContent!.trim().split(/\s+/)});
      }
      return { pages:result.totalPages, painted, fragments:paragraphs.length, lines, before,
        after:engine.materializePageMap(0, 'native-page-fit'), textSame:paragraphs.map(p => p.textContent).join('') === text };
    }, source.paragraphs.length);
    expect(actual.pages).toBe(source.nativePageCount);
    expect(actual.painted).toBe(actual.fragments);
    expect(actual.textSame).toBe(true);
    expect(actual.before).toEqual(actual.after);
    expect(actual.lines).toEqual(source.paragraphs.map((p: any) => p.nativeLines.map((words: string[], i: number) => ({page:p.nativeLinePages[i], words}))));
  });
}
}

for (const mode of [14, 15]) {
  for (const indent of [0, 36]) {
    test(`mode ${mode} distinguishes raw font and fixed-cell boundaries with ${indent}pt indent`, async ({ page }) => {
      // Liberation Mono advances are 1229/2048 em: 128 glyphs at 12pt = 921.75pt exactly.
      // Word fits the issue's 65-glyph prefix into 468pt despite its 468.076171875pt
      // raw advance. Preserve the browser control and separately test the inferred
      // 144-twip fixed-cell policy: 128 cells occupy 921.6pt.
      const prefix = 'Copper '.repeat(18) + 'ox';
      await page.setContent(`<style>p{font:12pt 'Liberation Mono',monospace;line-height:18px;margin:0;
        --docx-auto-line-spacing:1;text-indent:${indent}pt}span{white-space:pre-wrap;font-kerning:none;line-height:18px}</style>
        <main data-docx-compatibility-mode="${mode}">${[921.55, 921.6, 921.65, 921.70, 921.75, 921.80].map(width =>
          `<p style="width:${width + indent}pt"><span>${prefix} TAIL</span></p>`).join('')}</main>`);
      await page.addScriptTag({ content: wordText });
      await page.evaluate(() => document.fonts.ready);
      const facts = await page.evaluate(prefix => {
        const root = document.querySelector('main')!, context = document.createElement('canvas').getContext('2d')!;
        context.font = '16px "Liberation Mono"'; context.fontKerning = 'none';
        const lines = () => Array.from(root.querySelectorAll('p')).map(p => {
          const rows = new Map<number, string>(), walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            for (let i = 0; i < node.textContent!.length; i++) {
              const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
              const y = range.getBoundingClientRect().top;
              rows.set(y, (rows.get(y) ?? '') + node.textContent![i]);
            }
          }
          return Array.from(rows.values()).map(text => text.trimEnd());
        });
        const before = lines(), prepared = (window as any).prepareWordText(root);
        return { width: context.measureText(prefix).width, before, after: lines(), prepared };
      }, prefix);
      expect(facts.width).toBe(1229);
      expect(facts.prepared).toBe(6);
      expect(facts.before).toEqual([...Array.from({length:4}, () => [prefix.slice(0, -3), 'ox TAIL']),
        [prefix, 'TAIL'], [prefix, 'TAIL']]);
      expect(facts.after).toEqual([[prefix.slice(0, -3), 'ox TAIL'], ...Array.from({length:5}, () => [prefix, 'TAIL'])]);
    });
  }
}

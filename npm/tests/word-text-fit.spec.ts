import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

const wordText = readFileSync(new URL('../dist/word-text.js', import.meta.url), 'utf8').replace(/^export /gm, '');
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

for (const source of observations.cases) {
  test(`${source.id} retains native page membership through paragraph continuation`, async ({ page }) => {
    await page.setContent(`<style>p{font:${source.fontSizePt}pt 'Liberation Serif',serif;line-height:normal;
      margin:0 0 6pt;--docx-auto-line-spacing:1.15} .run{white-space:pre-wrap;position:static;
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

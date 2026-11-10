import { expect, test } from '@playwright/test';
import { wordTextScript as wordText } from './word-text-script.js';

const sentence = 'Soft amber lamps guide seven quiet boats toward shore.';
const tail = ' The dock stays open until dawn.';
const styles = `<style>p{font:12pt 'Liberation Sans';line-height:15pt;--docx-exact-line-height:15pt;
  margin:0;text-align:justify} .run{white-space:pre-wrap;vertical-align:top;line-height:15pt;font-kerning:none}</style>`;

for (const indent of [0, 12]) {
  test(`repeated spaces fit after ${indent}pt indent beside unchanged automatic text`, async ({ page }) => {
    await page.setContent(`${styles}<main data-docx-compatibility-mode="15">
      <p id="auto" style="line-height:normal;--docx-exact-line-height:0;--docx-auto-line-spacing:1"><span class="run" style="line-height:normal;vertical-align:baseline">Unchanged automatic paragraph.</span></p>
      <p id="exact" style="font:12px monospace;line-height:15pt;text-indent:${indent}pt"><span class="run" data-docx-exact-run="true">WWWW  WWWW  WWWW  WWWW  WWWW</span></p></main>`);
    await page.addScriptTag({ content: wordText });
    const result = await page.evaluate(indent => {
      const root = document.querySelector('main')!, p = document.getElementById('exact')!, auto = document.getElementById('auto')!;
      const context = document.createElement('canvas').getContext('2d')!; context.font='12px monospace';
      p.style.width = `${context.measureText('WWWW').width * 3 + context.measureText(' ').width * 4 * .8 + indent * 4 / 3}px`;
      const before = auto.outerHTML, text = root.textContent;
      (window as any).prepareWordText(root);
      const lines = Array.from(p.querySelectorAll<HTMLElement>('[data-docx-word-line]'));
      return {before,after:auto.outerHTML,text,afterText:root.textContent,
        lines:lines.map(l => l.textContent!.trim()),firstIndent:lines[0]?.style.marginLeft,
        finalSpacing:lines.at(-1)?.style.wordSpacing};
    }, indent);
    expect(result.after).toBe(result.before);
    expect(result.afterText).toBe(result.text);
    expect(result.lines).toEqual(['WWWW  WWWW  WWWW','WWWW  WWWW']);
    expect(result.firstIndent).toBe(indent ? '16px' : '');
    expect(result.finalSpacing).toBe('0px');
  });
}

for (const pitch of [15]) {
  test(`fitted exact ${pitch}pt lines retain full pitch and visible ink across pages`, async ({ page }) => {
    const text = Array(12).fill(sentence + tail).join(' ');
    await page.setContent(`${styles}<style>p,.run{line-height:${pitch}pt;--docx-exact-line-height:${pitch}pt}</style>
      <div id="staging" data-docx-compatibility-mode="15"><div data-section-index="0"
      data-page-width="388" data-page-height="150" data-content-width="298" data-content-height="60"
      data-margin-top="45" data-margin-bottom="45" data-margin-left="45" data-margin-right="45">
      <p data-source-anchor-id="p:body:exact" data-widow-control="false"><span class="run" data-docx-exact-run="true">${text}</span></p>
      </div></div><div id="container"></div>`);
    await page.addScriptTag({ content: wordText });
    await page.addScriptTag({ path: new URL('../dist/pagination.bundle.js', import.meta.url).pathname });
    const actual = await page.evaluate(() => {
      const api = window as any, root = document.getElementById('staging')!, container = document.getElementById('container')!;
      api.prepareWordText(root);
      const count = root.querySelectorAll('[data-docx-word-line]').length;
      const engine = new api.DocxodusPagination.PaginationEngine(root, container, {showPageNumbers:false,fragmentParagraphs:true});
      const result = engine.paginate();
      const paragraphs = Array.from(container.querySelectorAll('p'));
      return {pages:result.totalPages,count,text:paragraphs.map(p => p.textContent).join(''),
        fragments:paragraphs.map(p => ({height:p.getBoundingClientRect().height,
          lines:p.querySelectorAll('[data-docx-word-line]').length,
          preserved:p.hasAttribute('data-docx-exact-ink')}))};
    });
    expect(actual.count).toBeGreaterThan(12);
    expect(actual.pages).toBeGreaterThan(1);
    expect(actual.text).toBe(text);
    expect(actual.fragments.reduce((n,p) => n+p.lines,0)).toBe(actual.count);
    for (const fragment of actual.fragments)
      expect(Math.abs(fragment.height - fragment.lines * pitch * 4 / 3)).toBeLessThan(.25);
  });
}

for (const mode of [14, 15]) for (const width of [293, 298, 310]) for (const final of [false, true]) {
  test(`exact line fit: mode ${mode}, ${width}pt, final=${final}`, async ({ page }) => {
    await page.setContent(`${styles}<main data-docx-compatibility-mode="${mode}"><p style="width:${width}pt">
      <span class="run" data-docx-exact-run="true">${sentence + (final ? '' : tail)}</span></p></main>`);
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const root = document.querySelector<HTMLElement>('main')!, p = root.querySelector('p')!, run = p.querySelector('.run')!;
      const read = () => {
        const rows = new Map<number, Array<{text:string;left:number;width:number}>>();
        const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const word of node.textContent!.matchAll(/\S+/g)) {
          const range = document.createRange(); range.setStart(node, word.index!); range.setEnd(node, word.index! + word[0].length);
          const rect = range.getBoundingClientRect(), row = rows.get(rect.top) ?? [];
          row.push({text:word[0],left:rect.left,width:rect.width}); rows.set(rect.top,row);
        }
        return Array.from(rows.values());
      };
      const before = read(), original = root.innerHTML;
      (window as any).prepareWordText(root);
      const after = read(), box = p.getBoundingClientRect();
      const canvas = document.createElement('canvas').getContext('2d')!;
      canvas.font = "16px 'Liberation Sans'";
      return {before,after,unchanged:root.innerHTML === original,text:p.textContent!.trim(),height:box.height,
        naturalSpace:canvas.measureText(' ').width,painted:(window as any).paintWordText(root),
        gaps:after.flatMap(row => row.slice(1).map((word,i) => word.left-row[i].left-row[i].width))};
    });
    const fits = width === 310 || (mode === 15 && width === 298);
    expect(actual.after[0].map(w => w.text).join(' ')).toBe(fits ? sentence : sentence.replace(' shore.', ''));
    expect(actual.text).toBe(sentence + (final ? '' : tail));
    expect(actual.painted).toBe(0);
    expect(actual.height).toBeCloseTo(actual.after.length * 20, 2);
    // Range endpoints snap to layout units; moving a word can change their difference by one unit.
    actual.after.flat().forEach((word,i) =>
      expect(Math.abs(word.width - actual.before.flat()[i].width)).toBeLessThanOrEqual(1 / 64));
    for (const gap of actual.gaps) expect(gap + 1 / 32).toBeGreaterThanOrEqual(actual.naturalSpace * .75);
    if (mode === 14) expect(actual.unchanged).toBe(true);
  });
}

for (const [name, content, attributes] of [
  ['hard break', `${sentence}<br>${tail}`, ''],
  ['tab', `${sentence}\t${tail}`, ''],
  ['mixed font', `${sentence}<span style="font-family:serif">${tail}</span>`, ''],
  ['bidi', sentence + tail, 'dir="rtl"'],
  ['left aligned', sentence + tail, 'style="text-align:left"'],
  ['overlapping ink', sentence + tail, 'style="line-height:4pt;--docx-exact-line-height:4pt"'],
  ['shifted run', sentence + tail, ''],
  ['aligned final line', sentence + tail, 'style="text-align-last:right"'],
]) {
  test(`exact ${name} retains its existing flow`, async ({ page }) => {
    await page.setContent(`${styles}<main data-docx-compatibility-mode="15" style="width:298pt"><p ${attributes}>
      <span class="run" data-docx-exact-run="true" style="${name === 'overlapping ink' ? 'line-height:4pt' :
        name === 'shifted run' ? 'position:relative;left:0;right:0;top:-3pt' : ''}">${content}</span></p></main>`);
    await page.addScriptTag({ content: wordText });
    const actual = await page.evaluate(() => {
      const root = document.querySelector('main')!, before = root.innerHTML;
      (window as any).prepareWordText(root);
      return {before,after:root.innerHTML};
    });
    expect(actual.after).toBe(actual.before);
  });
}

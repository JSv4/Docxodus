import { expect, test } from '@playwright/test';
import { wordTextScript } from './word-text-script.js';

const styles = `<style>p{font:12pt/18px 'Liberation Mono';margin:0;width:144pt;--docx-auto-line-spacing:1}
  .run{white-space:pre-wrap;font-kerning:none;line-height:18px}</style>`;
const words = Array.from({ length: 72 }, (_, i) => `W${String(i).padStart(4, '0')}`);

for (const [scale, perLine] of [[.8, 4], [1, 3], [1.5, 2]]) {
  test(`${scale * 100}% width controls wrapping and survives page continuations`, async ({ page }) => {
    await page.setContent(`${styles}<div id="staging"><div data-section-index="0"
      data-page-width="288" data-page-height="279" data-content-width="144" data-content-height="135"
      data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72">
      <p data-source-anchor-id="p:body:00000000000000000000000000000001" data-widow-control="false">
        <span class="run" style="--docx-character-scale:${scale}">${words.join(' ')}</span></p>
      </div></div><div id="container"></div>`);
    await page.addScriptTag({ content: wordTextScript });
    await page.addScriptTag({ path: new URL('../dist/pagination.bundle.js', import.meta.url).pathname });
    const result = await page.evaluate(() => {
      const api = window as any, root = document.getElementById('staging')!, container = document.getElementById('container')!;
      const prepared = api.prepareWordText(root), first = root.innerHTML;
      api.prepareWordText(root);
      const repeated = root.innerHTML;
      const engine = new api.DocxodusPagination.PaginationEngine(root, container, { showPageNumbers:false, fragmentParagraphs:true });
      const pages = engine.paginate().totalPages, before = engine.materializePageMap(0, 'width');
      const painted = api.paintWordText(container);
      return { prepared, repeated: first === repeated, pages, before, after: engine.materializePageMap(0, 'width'), painted,
        paragraphs: container.querySelectorAll('p').length,
        lines: Array.from(container.querySelectorAll('svg > text')).map(t => t.textContent!.trim()),
        heights: Array.from(container.querySelectorAll('svg > text')).map(t => t.getBoundingClientRect().height) };
    });
    expect(result.prepared).toBe(1);
    expect(result.repeated).toBe(true);
    expect(result.painted).toBe(result.paragraphs);
    expect(result.pages).toBe(Math.ceil(Math.ceil(words.length / perLine) / 10));
    expect(result.after).toEqual(result.before);
    expect(result.lines).toEqual(Array.from({ length: Math.ceil(words.length / perLine) }, (_, i) =>
      words.slice(i * perLine, (i + 1) * perLine).join(' ')));
    expect(Math.max(...result.heights) - Math.min(...result.heights)).toBeLessThan(1 / 64);
  });
}

for (const align of ['left', 'center', 'right', 'justify']) {
  test(`character width preserves the ${align} origin, height and baseline`, async ({ page }) => {
    await page.setContent(`${styles}<main>${[.8, 1, 1.5].map(scale =>
      `<p style="text-align:${align};text-indent:18pt"><span class="run" style="--docx-character-scale:${scale}">WidthMarker</span></p>`).join('')}</main>`);
    await page.addScriptTag({ content: wordTextScript });
    const result = await page.evaluate(() => {
      const api = window as any, root = document.querySelector('main')!;
      api.prepareWordText(root); api.paintWordText(root);
      return Array.from(root.querySelectorAll('p')).map(p => {
        const text = p.querySelector('text')!, span = text.querySelector('tspan')!;
        return { width:span.getBoundingClientRect().width, x:span.getBoundingClientRect().left - p.getBoundingClientRect().left,
          height:p.getBoundingClientRect().height, baseline:Number(text.getAttribute('y')) };
      });
    });
    for (const [i, scale] of [.8, 1, 1.5].entries()) {
      expect(result[i].width / result[1].width).toBeCloseTo(scale, 3);
      const remaining = 192 - 24 - result[i].width;
      expect(result[i].x).toBeCloseTo(24 + (align === 'right' ? remaining : align === 'center' ? remaining / 2 : 0), 0);
      expect(result[i].height).toBe(result[1].height);
      expect(result[i].baseline).toBe(result[1].baseline);
    }
  });
}

for (const [name, wrapper, content] of [
  ['header', '<div data-hf-type="header-default">CONTENT</div>', 'WidthMarker'],
  ['cell', '<table><tr><td>CONTENT</td></tr></table>', 'WidthMarker'],
  ['columns', '<div style="column-count:2">CONTENT</div>', 'WidthMarker'],
  ['picture neighbor', '<div><img width="1" height="1">CONTENT</div>', 'WidthMarker'],
  ['explicit break', 'CONTENT', 'Width<br>Marker'],
  ['wide word', 'CONTENT', 'W'.repeat(100)],
  ['tab', 'CONTENT', 'Width\tMarker'],
  ['nonbreaking space', 'CONTENT', 'Width\u00a0Marker'],
]) test(`scaled ${name} retains its existing owner atomically`, async ({ page }) => {
  await page.setContent(`${styles}<main>${wrapper.replace('CONTENT',
    `<p><span class="run" style="--docx-character-scale:1.5">${content}</span></p>`)}</main>`);
  await page.addScriptTag({ content: wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector('main')!, before = root.innerHTML, api = window as any;
    const prepared = api.prepareWordText(root), painted = api.paintWordText(root);
    return { prepared, painted, unchanged:root.innerHTML === before };
  });
  expect(result).toEqual({ prepared:0, painted:0, unchanged:true });
});

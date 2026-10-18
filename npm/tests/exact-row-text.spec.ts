import { expect, test } from '@playwright/test';
import { build } from 'esbuild';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { decodePng } from './visual-parity/png.js';

let module = '';
test.beforeAll(async () => {
  const result = await build({ stdin: { resolveDir: join(dirname(fileURLToPath(import.meta.url)), '../src'),
    contents: `export { preserveClippedExactRowText } from './exact-row-text.ts';
      export { renderedSourceMeasurerExpression } from './pagination.ts';` },
    bundle: true, format: 'esm', write: false });
  module = result.outputFiles[0].text;
});

for (const band of ['page-header', 'page-footer', 'page-content', 'inset-body', 'unclipped-control']) {
  for (const offset of [36, 36.2]) {
    test(`exact-row text stays whole under ${band} at ${offset}pt`, async ({ page }) => {
      await page.setContent(`<style>
        @page { size: 612pt 792pt; margin: 0; }
        body { margin: 0; font: 11pt/15pt 'Liberation Serif'; }
        p { margin: 0; }
        .page-box { position: relative; width: 612pt; height: 792pt; overflow: hidden; }
        .band { position: absolute; top: ${offset}pt; left: 72pt; width: 468pt; height: 36pt; }
      </style><div class="page-box"><div class="band ${band === 'inset-body' ? 'page-content' : band}" style="${band === 'inset-body'
        ? 'overflow:visible;clip-path:inset(-3px -96px -3px -96px)' : band === 'page-content'
        ? 'overflow-x:visible;overflow-y:clip' : `overflow:${band === 'unclipped-control' ? 'visible' : 'hidden'}`}">
        <div style="display:flex">${['Hdr', 'Cell'].map(prefix =>
          `<div data-docx-exact-row-cell="true" style="width:50%;height:15pt;overflow:hidden">`
          + ['Alder', 'Birch', 'Cedar'].map(word => `<p><span style="position:${offset === 36 ? 'static' : 'relative'}">${prefix}${word}</span></p>`).join('')
          + '</div>').join('')}</div>
        <p>HdrValue</p><span id="side" style="position:absolute;left:-60pt;top:0">SideAlder</span>
      </div></div>`);
      await page.addScriptTag({ type: 'module', content: `${module}\nwindow.rewrite = preserveClippedExactRowText;
        const observedMeasureSource = Function('return ' + renderedSourceMeasurerExpression())();
        window.sideGeometry = () => observedMeasureSource(window, document.getElementById('side'), document.querySelector('.page-box')).visible;` });
      await page.waitForFunction(() => (window as any).rewrite);
      const before = decodePng(await page.screenshot());
      const geometry = await page.evaluate(() => (window as any).sideGeometry());
      expect(await page.evaluate(() => (window as any).rewrite(document.body))).toBe(2);
      const revised = await page.evaluate(() => (window as any).sideGeometry());
      for (const edge of ['left', 'right', 'top', 'bottom'])
        expect(Math.abs(geometry[edge] - revised[edge])).toBeLessThanOrEqual(0.5);
      const after = decodePng(await page.screenshot());
      let changed = 0;
      for (let i = 0; i < before.data.length; i += 4)
        if (before.data.slice(i, i + 3).some((value, channel) => Math.abs(value - after.data[i + channel]) > 8)) changed++;
      expect(changed, 'the rewrite preserves visible clipping').toBeLessThan(10);
      const task = getDocument({ data: new Uint8Array(await page.pdf({ preferCSSPageSize: true })), verbosity: 0 });
      try {
        const pdf = await task.promise;
        const content = await (await pdf.getPage(1)).getTextContent();
        const words = content.items.flatMap(item => 'str' in item ? item.str.split(/\s+/).filter(Boolean) : []);
        expect(words).toContain('HdrAlder');
        expect(words).toContain('HdrValue');
        const rowWords = ['Hdr', 'Cell'].flatMap(prefix => ['Alder', 'Birch', 'Cedar'].map(word => prefix + word));
        expect(words.filter(word => ![...rowWords, 'HdrValue', 'SideAlder'].includes(word))).toEqual([]);
        const observed = words.filter(word => rowWords.includes(word));
        expect(observed).toEqual(rowWords.filter(word => words.includes(word)));
        if (['page-content', 'inset-body', 'unclipped-control'].includes(band)) expect(words).toContain('SideAlder');
        for (const word of ['HdrAlder', 'HdrBirch', 'HdrCedar', 'HdrValue'])
          expect(words.filter(value => value === word).length).toBeLessThanOrEqual(1);
      } finally { await task.destroy(); }
    });
  }
}

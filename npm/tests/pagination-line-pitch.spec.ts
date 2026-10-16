import { test, expect } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = dirname(fileURLToPath(import.meta.url));

/**
 * Issue #1151: the paginated view lays single-spaced lines out at the font's own line height.
 *
 * Chromium rounds a `line-height: normal` line to a whole CSS pixel. The export has resolved that
 * since #850; the paginated view and editor did not, so 11pt Times New Roman advanced 12.75pt a
 * line where Word and LibreOffice advance 12.65pt, and the error moved page breaks.
 *
 * The case uses the repository's own test font, so it does not depend on installed fonts. Its
 * `hhea` table gives ascender 1901 + |descender| 483 + lineGap 0 = 2384 per 2048 units/em, so at
 * 10pt its natural line is 11.64pt; Chromium's rounding makes it 15px, 11.25pt.
 */
const NATURAL_EM = 2384 / 2048;
const FONT_SIZE_PT = 10;
const fontBase64 = readFileSync(join(__dirname, '..', '..', 'docs', 'demo', 'fonts', 'docxodus-canvas-mono.woff2'))
  .toString('base64');

test('paginated single-spaced lines advance by the font\'s unrounded natural height', async ({ page }) => {
  const paragraphs = Array.from({ length: 60 }, (_, index) => `<p class="pitch">line ${index + 1}</p>`).join('');
  await page.setContent(`
    <style>
      @font-face { font-family: 'Pitch Test'; src: url(data:font/woff2;base64,${fontBase64}) format('woff2'); }
      .pitch { font-family: 'Pitch Test'; font-size: ${FONT_SIZE_PT}pt; margin: 0; padding: 0; }
    </style>
    <div id="staging">
      <div data-section-index="0"
           data-page-width="400" data-page-height="320"
           data-content-width="380" data-content-height="300"
           data-margin-top="10" data-margin-right="10"
           data-margin-bottom="10" data-margin-left="10">${paragraphs}</div>
    </div>
    <div id="container"></div>`);
  await page.evaluate(async () => {
    await document.fonts.load(`10pt 'Pitch Test'`);
    await document.fonts.ready;
  });
  await page.addScriptTag({ path: join(__dirname, '../dist/pagination.bundle.js') });

  const firstPage = await page.evaluate(() => {
    const staging = document.getElementById('staging') as HTMLElement;
    const container = document.getElementById('container') as HTMLElement;
    const { PaginationEngine } = (window as any).DocxodusPagination;
    new PaginationEngine(staging, container, { scale: 1, showPageNumbers: false }).paginate();
    const lines = Array.from(container.querySelector('.page-content')!.querySelectorAll('p'));
    const tops = lines.map((line) => line.getBoundingClientRect().top);
    return { lines: lines.length, pitchPt: (tops[tops.length - 1] - tops[0]) / (tops.length - 1) * 0.75 };
  });

  const naturalPt = NATURAL_EM * FONT_SIZE_PT;
  expect(firstPage.pitchPt).toBeCloseTo(naturalPt, 1);
  // 300pt of body holds 25 lines of 11.64pt; at the rounded 11.25pt it held 26.
  expect(firstPage.lines).toBe(Math.floor(300 / naturalPt));
});

test('the read-only paginated view puts glyphs on Word\'s baseline without moving any line', async ({ page }) => {
  // Word sets a single-spaced line's baseline its natural height less the font's descent below the
  // line's top; Chromium rounds ascent and descent and lands up to a pixel off (issue #942). The
  // export has corrected glyph placement since #942; the read-only paginated view now does too.
  // An editable surface (the paginated editor, direct engine callers by default) keeps it off.
  const paragraphs = Array.from({ length: 3 }, (_, index) => `<p class="pitch">line ${index + 1}</p>`).join('');
  const html = `
    <style>
      @font-face { font-family: 'Pitch Test'; src: url(data:font/woff2;base64,${fontBase64}) format('woff2'); }
      .pitch { font-family: 'Pitch Test'; font-size: ${FONT_SIZE_PT}pt; margin: 0; padding: 0; }
    </style>
    <div id="root"><div id="pagination-staging">
      <div data-section-index="0"
           data-page-width="400" data-page-height="320"
           data-content-width="380" data-content-height="300"
           data-margin-top="10" data-margin-right="10"
           data-margin-bottom="10" data-margin-left="10">${paragraphs.replace(/<p class="pitch">([^<]*)<\/p>/g, '<p class="pitch"><span>$1</span></p>')}</div>
    </div><div id="pagination-container"></div></div>`;
  await page.setContent('<div id="host"></div>');
  await page.evaluate(async (fontCss) => {
    const style = document.createElement('style'); style.textContent = fontCss; document.head.appendChild(style);
    await document.fonts.load(`10pt 'Pitch Test'`);
  }, `@font-face { font-family: 'Pitch Test'; src: url(data:font/woff2;base64,${fontBase64}) format('woff2'); }`);
  await page.addScriptTag({ path: join(__dirname, '../dist/pagination.bundle.js') });

  const layout = await page.evaluate((markup) => {
    const { paginateHtml, PaginationEngine } = (window as any).DocxodusPagination;
    const measure = (container: HTMLElement) => {
      const paragraph = container.querySelector<HTMLElement>('.page-content p')!;
      const span = paragraph.querySelector<HTMLElement>('span')!;
      return { paragraphTop: paragraph.getBoundingClientRect().top - container.getBoundingClientRect().top,
        position: getComputedStyle(span).position, top: getComputedStyle(span).top };
    };
    const host = document.getElementById('host')!;
    paginateHtml(markup, host, { scale: 1, showPageNumbers: false });
    const viewer = measure(host);
    host.innerHTML = markup;
    new PaginationEngine(host.querySelector('#pagination-staging'), host.querySelector('#pagination-container'),
      { scale: 1, showPageNumbers: false }).paginate();
    const engine = measure(host);
    return { viewer, engine };
  }, html);

  expect(layout.viewer.position).toBe('relative');
  expect(Number.parseFloat(layout.viewer.top)).not.toBe(0);
  expect(layout.engine.position).toBe('static');
  expect(layout.viewer.paragraphTop).toBeCloseTo(layout.engine.paragraphTop, 2);
});

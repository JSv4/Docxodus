import { expect, test, type Page } from '@playwright/test';
import { build } from 'esbuild';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
// Bundled, so the module's own imports (exact-row-text.js) resolve inside the inline script.
let module = '';
test.beforeAll(async () => {
  const result = await build({
    entryPoints: [join(here, '../src/paragraph-borders.ts')], bundle: true, format: 'esm', write: false,
  });
  module = result.outputFiles[0].text;
});

// Chromium lays boxes out in 1/64 CSS pixel units.
const LAYOUT_UNIT = 1 / 64;
const close = (actual: number, expected: number) => expect(Math.abs(actual - expected)).toBeLessThanOrEqual(LAYOUT_UNIT);

// Issue #1116: a declared 0.25pt paragraph border is laid out and painted at 0.25pt in the export,
// not at the one CSS pixel (0.75pt) the browser draws for any border.
async function render(page: Page, style: 'solid' | 'dashed') {
  await page.setContent(`
    <style>body { margin: 0; font: 11pt serif; } .page-box { position: relative; width: 612pt; height: 200pt; }</style>
    <div class="page-box">
      <div style="padding: 72pt">
        <div id="card" data-docx-exact-border="true" style="
          --docx-border-top-width: 0.25pt; --docx-border-left-width: 0.25pt;
          border-top: ${style} rgb(36, 107, 125) 0.75pt; border-left: ${style} rgb(36, 107, 125) 0.75pt;
          border-bottom: solid rgb(36, 107, 125) 0.75pt;
          padding: 4pt; margin-left: -4.75pt"><p id="text" style="margin: 0">BorderCard Juniper.</p></div>
      </div>
    </div>`);
  await page.addScriptTag({ type: 'module', content: `${module}\nwindow.borders = { prepareParagraphBorders, paintParagraphBorders };` });
  await page.waitForFunction(() => (window as any).borders);
  return page.evaluate(() => {
    const { prepareParagraphBorders, paintParagraphBorders } = (window as any).borders;
    const text = () => document.getElementById('text')!.getBoundingClientRect();
    const before = text();
    const prepared = prepareParagraphBorders(document.body);
    const after = text();
    const painted = paintParagraphBorders(document.body);
    const shapes = Array.from(document.querySelectorAll('.page-box > svg > *')).map((shape) =>
      Object.fromEntries(Array.from(shape.attributes).map((a) => [a.name, a.value])));
    const card = getComputedStyle(document.getElementById('card')!);
    return {
      prepared, painted, shapes,
      before: { left: before.left, top: before.top }, after: { left: after.left, top: after.top },
      borders: [card.borderTopStyle, card.borderLeftStyle, card.borderBottomStyle],
    };
  });
}

test('a declared quarter-point border is laid out and painted at its physical width', async ({ page }) => {
  const result = await render(page, 'solid');
  expect(result.prepared).toBe(1);
  expect(result.painted).toBe(1);
  // The recorded sides leave CSS layout; the omitted-width bottom side keeps its CSS border.
  expect(result.borders).toEqual(['none', 'none', 'solid']);
  // The text stays at the left text edge and moves up by the 0.5pt the pixel border over-spaced.
  close(result.after.left, result.before.left);
  close(result.after.top, result.before.top - 0.5 * 4 / 3);
  const quarterPoint = 0.25 * 4 / 3;
  const [top, left] = result.shapes;
  expect(Number(top.height)).toBeCloseTo(quarterPoint, 6);
  expect(Number(left.width)).toBeCloseTo(quarterPoint, 6);
  // The left border ends 4pt (+ its own width) outside the text edge.
  close(Number(left.x) + quarterPoint + 4 * 4 / 3, result.after.left);
  expect(top.fill).toBe('rgb(36, 107, 125)');
});

test('a dashed quarter-point border keeps the dash rhythm of its one-pixel CSS border', async ({ page }) => {
  const result = await render(page, 'dashed');
  const quarterPoint = 0.25 * 4 / 3;
  const [top] = result.shapes;
  expect(Number(top['stroke-width'])).toBeCloseTo(quarterPoint, 6);
  // Only the width changes: the dash and gap stay those of the 1px border CSS drew (3px / 2px).
  expect(top['stroke-dasharray'].split(' ').map(Number)).toEqual([3, 2]);
});

for (const border of ['solid', 'dashed']) {
  test(`${border} prepared border spacing survives HTML serialization`, async ({ page }) => {
    await page.setContent(`<style>body { margin:0; font:12px serif; }</style>${
      Array.from({ length: 12 }, (_, index) => `<div data-docx-exact-border="true"
        style="--docx-border-top-width:1.25pt; --docx-border-bottom-width:1.25pt;
          border-top:${border} 1.25pt teal; border-bottom:${border} 1.25pt teal; padding:1pt">
        <p style="margin:0">Kestrel card ${index + 1}.</p></div>`).join('')}`);
    await page.addScriptTag({ type: 'module', content: `${module}\nwindow.borders = { prepareParagraphBorders };` });
    await page.waitForFunction(() => (window as any).borders);
    const before = await page.evaluate(() => {
      (window as any).borders.prepareParagraphBorders(document.body);
      document.querySelectorAll('script').forEach(script => script.remove());
      return { html: document.documentElement.outerHTML,
        geometry: Array.from(document.querySelectorAll('p'), p => p.getBoundingClientRect().toJSON()) };
    });
    await page.setContent(before.html);
    expect(await page.locator('p').evaluateAll(paragraphs => paragraphs.map(p =>
      p.getBoundingClientRect().toJSON()))).toEqual(before.geometry);
  });
}

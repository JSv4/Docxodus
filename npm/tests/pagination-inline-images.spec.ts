import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
import { fittingTilePng } from './docx-page-fitting-fixture.js';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const source = `data:image/png;base64,${fittingTilePng().toString('base64')}`;

for (const segmenter of [true, false]) {
  test(`inline picture boundaries preserve image identities with grapheme segmenter ${segmenter}`, async ({ page }) => {
    await page.setContent(`<style>body{font:11pt/15pt Arial}p{margin:0}</style><div id="staging">
      <section data-section-index="0" data-page-width="220" data-page-height="200"
      data-content-width="180" data-content-height="160" data-margin-top="20" data-margin-bottom="20"
      data-margin-left="20" data-margin-right="20"><p id="gallery" data-anchor="p:body:gallery">
      ${[1,2,3,4].map(index => `<span id="carrier${index}" style="position:relative;top:-0.5px"><img id="image${index}" data-anchor="img:body:tile${index}"
        src="${source}" style="width:120pt;height:95pt"></span>`).join('')}</p><p>Orchid Å🙂 complete.</p>
      </section></div><div id="container"></div>`);
    await page.locator('img').evaluateAll(images => Promise.all(images.map(image => (image as HTMLImageElement).decode())));
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(segmenter => {
      if (!segmenter) Object.defineProperty(Intl, 'Segmenter', { configurable: true, value: undefined });
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      const container = document.getElementById('container')!;
      const ids = Array.from(container.querySelectorAll('[id]'), element => element.id);
      return { text: container.textContent, ids, images: Array.from(container.querySelectorAll('img'), image => ({
        id: image.id, source: image.src, width: image.getBoundingClientRect().width * 0.75,
        height: image.getBoundingClientRect().height * 0.75,
        overflow: image.getBoundingClientRect().bottom - image.closest('.page-content')!.getBoundingClientRect().bottom,
      })) };
    }, segmenter);
    expect(new Set(result.ids).size).toBe(result.ids.length);
    expect(result.images.map(image => image.id)).toEqual(['image1', 'image2', 'image3', 'image4']);
    expect(result.images.every(image => image.source === source && image.overflow <= 1)).toBe(true);
    for (const image of result.images) {
      expect(image.width).toBeCloseTo(120, 1);
      expect(image.height).toBeCloseTo(95, 1);
    }
    expect(result.text?.match(/Orchid Å🙂 complete\./g)).toHaveLength(1);
  });
}

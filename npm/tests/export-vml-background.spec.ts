import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { generateVmlBackgroundDocx } from './docx-vml-background-fixture.js';
import { decodePng } from './visual-parity/png.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

async function pdfFacts(bytes: Uint8Array) {
  const task = getDocument({ data: new Uint8Array(bytes), verbosity: 0 });
  try {
    const pdf = await task.promise, page = await pdf.getPage(1);
    const content = await page.getTextContent();
    const operators = await page.getOperatorList();
    return { count: pdf.numPages, view: page.view,
      images: operators.fnArray.filter(operation => operation === OPS.paintImageXObject || operation === OPS.paintInlineImageXObject).length,
      text: content.items.filter(item => 'str' in item).map(item => ({ text: item.str, transform: item.transform })) };
  } finally { await task.destroy(); }
}

test('a centered VML background uses the same page placement as its DrawingML control', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'markup' as const,
    markupPresentation: 'word' as const, commentProfile: 'margin' as const,
    commentColumnPlacement: 'balloons' as const, timeoutMs: 120_000 };
  const actual = await renderDocxArtifacts(generateVmlBackgroundDocx('vml'), options);
  const control = await renderDocxArtifacts(generateVmlBackgroundDocx('drawing'), options);
  const facts = await pdfFacts(actual.pdf!);
  expect(facts).toEqual(await pdfFacts(control.pdf!));
  expect(facts.count).toBe(1);
  expect(facts.images).toBe(1);
  expect(facts.text.map(item => item.text).join(' ')).toContain('Oriole background ledger.');
  expect(facts.text.map(item => item.text).join(' ')).toContain('Oriole body record.');
  await page.setContent(actual.html!);
  await page.evaluate(() => document.fonts.ready);
  expect(await page.locator('.page-box [data-docx-drawing-anchor]').count()).toBe(1);
  const placement = await page.locator('.page-box [data-docx-drawing-anchor]').evaluate(frame => {
    const paper = frame.closest('.page-box')!.getBoundingClientRect(), box = frame.getBoundingClientRect();
    const bitmap = frame.querySelector('img')!.getBoundingClientRect();
    return { left: (box.left - paper.left) * .75, top: (box.top - paper.top) * .75,
      width: box.width * .75, height: box.height * .75, imageWidth: bitmap.width * .75, imageHeight: bitmap.height * .75,
      behind: frame.getAttribute('data-docx-anchor-behind'), zIndex: Number(getComputedStyle(frame).zIndex) };
  });
  expect(placement).toMatchObject({ behind: 'true', width: 648, height: 828, imageWidth: 648, imageHeight: 828 });
  expect(placement.left).toBeCloseTo(0, 1); expect(placement.top).toBeCloseTo(0, 1);
  expect(placement.zIndex).toBeLessThan(0);
  const raster = decodePng(await page.locator('.page-box').screenshot());
  const pixel = (x: number, y: number) => Array.from(raster.data.slice((y * raster.width + x) * 4, (y * raster.width + x) * 4 + 3));
  expect(pixel(5, 5)).toEqual([41, 117, 137]);
  expect(pixel(raster.width - 6, raster.height - 6)).toEqual([242, 232, 212]);
});

for (const kind of ['inline', 'wrapped'] as const) test(`a ${kind} VML image retains its established flow fallback`, async ({ browser }) => {
  const result = await renderDocxArtifacts(generateVmlBackgroundDocx(kind), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  expect(result.pageCount).toBe(1);
  expect(result.html).not.toContain('data-docx-drawing-anchor');
  const facts = await pdfFacts(result.pdf!);
  expect(facts.text.map(item => item.text).join(' ')).toContain('Oriole body record.');
});

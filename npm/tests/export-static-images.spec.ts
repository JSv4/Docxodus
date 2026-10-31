import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { staticImageDocx } from './docx-static-image-fixture.js';
import { decodePng } from './visual-parity/png.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const format of ['png', 'tiff', 'wmf', 'anisotropic-wmf', 'tiny-wmf', 'wide-wmf', 'tall-wmf', 'group-wmf'] as const)
for (const header of ['wide-wmf', 'tall-wmf', 'group-wmf'].includes(format) ? [false] : [false, true]) {
  test(`${format} paints in the ${header ? 'header' : 'body'} with its authored extent`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const wmf = format.endsWith('wmf');
    const widthScale = format === 'wide-wmf' || format === 'group-wmf' ? 2 : 1, heightScale = format === 'tall-wmf' ? 2 : 1;
    const result = await renderDocxArtifacts(staticImageDocx(format, header), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000 });
    expect(result.renderReport.warnings.filter(w => w.code === 'image_decode_failed' || w.code === 'unsupported_content_placeholder')).toEqual([]);
    await page.setContent(result.html!);
    const image = page.locator(header ? '.page-header img' : '.page-content img');
    await expect(image).toHaveCount(1);
    const rendered = await image.evaluate(async (element: HTMLImageElement, { format, widthScale, heightScale }) => {
      await element.decode(); const rect = element.getBoundingClientRect();
      const wmf = format.endsWith('wmf');
      const canvas = document.createElement('canvas'); canvas.width = wmf ? 96 * widthScale : 2; canvas.height = wmf ? 64 * heightScale : 3;
      const context = canvas.getContext('2d')!; context.drawImage(element, 0, 0, canvas.width, canvas.height);
      return { width: rect.width * 0.75, height: rect.height * 0.75, natural: [element.naturalWidth, element.naturalHeight],
        pixel: [...context.getImageData(format === 'anisotropic-wmf' ? 16 : wmf ? 8 * widthScale : 0, wmf ? 16 * heightScale : 0, 1, 1).data] };
    }, { format, widthScale, heightScale });
    expect(rendered.width).toBeCloseTo(72 * widthScale, 1); expect(rendered.height).toBeCloseTo(48 * heightScale, 1);
    expect(rendered.natural).toEqual(format === 'tiny-wmf' ? [1, 1] : wmf ? [96, 64] : [2, 3]);
    expect(rendered.pixel).toEqual(wmf ? [48, 80, 128, 255] : [255, 0, 0, 255]);
    if (wmf) {
      // drawImage first paints at intrinsic size and then stretches that raster. Inspect the
      // delivered page instead: viewport-relative SVG pens are resolved at the CSS image size.
      const painted = decodePng(await image.screenshot());
      const stroke = format === 'anisotropic-wmf' ? 8 : 2 * widthScale;
      const coverage = (x: number, y: number) => (255 - painted.data[(y * painted.width + x) * 4]) / (255 - 48);
      const horizontal = Array.from({ length: 24 * heightScale }, (_, y) => coverage(32 * widthScale, y)).reduce((a, b) => a + b, 0);
      const vertical = Array.from({ length: 24 * widthScale }, (_, x) => coverage(x, 32 * heightScale)).reduce((a, b) => a + b, 0);
      expect(Math.abs(horizontal - stroke)).toBeLessThan(0.15);
      expect(Math.abs(vertical - stroke)).toBeLessThan(0.15);
    }
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise; expect(pdf.numPages).toBe(1);
      const page = await pdf.getPage(1), operations = await page.getOperatorList();
      // PDF.js folds a path's paint operation into constructPath's first argument.
      const painted = operations.fnArray.filter((op, i) => wmf
        ? op === OPS.constructPath && operations.argsArray[i][0] === OPS.stroke
        : op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject);
      expect(painted.length)
        .toBeGreaterThanOrEqual(wmf ? format === 'anisotropic-wmf' ? 1 : 2 : 1);
      expect((await page.getTextContent()).items.filter(i => 'str' in i).map(i => i.str).join('')).toBe('CypressAfter');
    } finally { await task.destroy(); }
  });
}

for (const format of ['unsupported-tiff', 'unsupported-wmf'] as const) test(`${format} remains explicit under warn and strict policies`, async ({ browser }) => {
  test.setTimeout(120_000);
  const input = staticImageDocx(format, false);
  const options = { browser, outputs: ['html', 'pdf'] as Array<'html' | 'pdf'>, reviewProfile: 'final' as const, commentProfile: 'hidden' as const, timeoutMs: 120_000 };
  const result = await renderDocxArtifacts(input, { ...options, unsupportedContent: 'warn' });
  expect(result.renderReport.warnings.some(w => w.code === 'unsupported_content_placeholder')).toBe(true);
  expect(result.html).toContain('CypressAfter');
  await expect(renderDocxArtifacts(input, { ...options, unsupportedContent: 'strict' })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'docx_conversion' });
});

import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { staticImageDocx } from './docx-static-image-fixture.js';
import { decodePng } from './visual-parity/png.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const kind of ['emf', 'wide-emf', 'tall-emf', 'anchored-emf', 'vml-emf', 'group-emf', 'geometric-emf'] as const)
for (const header of kind === 'emf' ? [false, true] : [false]) {
  test(`${kind} retains its frame, overflowing outline and text in ${header ? 'header' : 'body'} HTML/PDF`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(staticImageDocx(kind, header), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict', timeoutMs: 120_000 });
    expect(result.renderReport.warnings.filter(w => w.code === 'unsupported_content_placeholder')).toEqual([]);
    await page.setContent(result.html!);
    const svg = page.locator(`${kind === 'anchored-emf' ? '.page-drawing-layer' : header ? '.page-header' : '.page-content'} svg[data-docx-vector-frame="emf"]`);
    await expect(svg).toHaveCount(1);
    const geometry = await svg.evaluate((element: SVGSVGElement) => {
      const shape = element.querySelector('rect')!, matrix = shape.getScreenCTM()!, bounds = element.getBoundingClientRect();
      const paper = element.closest('.page-box')!.getBoundingClientRect();
      const left = new DOMPoint(10, 10).matrixTransform(matrix), right = new DOMPoint(90, 90).matrixTransform(matrix);
      return { frame: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
        left: { x: left.x, y: left.y }, right: { x: right.x, y: right.y },
        paper: { x: paper.x, y: paper.y },
        effect: shape.getAttribute('vector-effect') };
    });
    const width = kind === 'wide-emf' || kind === 'group-emf' ? 192 : 96, height = kind === 'tall-emf' ? 128 : 64;
    expect(geometry.frame.width).toBeCloseTo(width, 1); expect(geometry.frame.height).toBeCloseTo(height, 1);
    if (kind === 'anchored-emf') {
      expect(geometry.frame.x - geometry.paper.x).toBeCloseTo(96, 1);
      expect(geometry.frame.y - geometry.paper.y).toBeCloseTo(192, 1);
    }
    expect(geometry.right.x - geometry.left.x).toBeCloseTo(width * 2500 / 2646, 1);
    expect(geometry.right.x).toBeGreaterThan(geometry.frame.x + geometry.frame.width);
    expect(geometry.effect).toBe(kind === 'geometric-emf' ? null : 'non-scaling-stroke');
    const screenshot = decodePng(await page.screenshot({ fullPage: true }));
    const row = Math.round((geometry.left.y + geometry.right.y) / 2);
    const coverage = Array.from({ length: 10 }, (_, i) => Math.round(geometry.right.x) - 5 + i)
      .map(x => (255 - screenshot.data[(row * screenshot.width + x) * 4]) / 255).reduce((a, b) => a + b, 0);
    const expectedStroke = kind === 'geometric-emf' ? width * 62.5 / 2646 : 1;
    expect(Math.abs(coverage - expectedStroke)).toBeLessThan(0.2);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise; expect(pdf.numPages).toBe(1);
      const pdfPage = await pdf.getPage(1), ops = await pdfPage.getOperatorList();
      expect(ops.fnArray.filter(op => op === OPS.constructPath).length).toBeGreaterThan(0);
      const viewport = pdfPage.getViewport({ scale: 4 / 3 });
      const factory = pdf.canvasFactory as { create(width: number, height: number): { canvas: HTMLCanvasElement; context: CanvasRenderingContext2D } };
      const target = factory.create(Math.ceil(viewport.width), Math.ceil(viewport.height));
      await pdfPage.render({ ...target, canvasContext: target.context, viewport }).promise;
      const pdfRow = Math.round((geometry.left.y + geometry.right.y) / 2 - geometry.paper.y);
      const pdfX = Math.round(geometry.right.x - geometry.paper.x) - 5;
      const pixels = target.context.getImageData(pdfX, pdfRow, 10, 1).data;
      const pdfCoverage = Array.from({ length: 10 }, (_, i) => (255 - pixels[i * 4]) / 255).reduce((a, b) => a + b, 0);
      expect(Math.abs(pdfCoverage - expectedStroke)).toBeLessThan(0.25);
      expect((await pdfPage.getTextContent()).items.filter(i => 'str' in i).map(i => i.str).join('')).toBe('CypressAfter');
    } finally { await task.destroy(); }
  });
}

test('an unknown record after the rectangle rejects the whole EMF under warn and strict policies', async ({ browser }) => {
  test.setTimeout(120_000);
  const input = staticImageDocx('unsupported-emf', false);
  const options = { browser, outputs: ['html', 'pdf'] as Array<'html' | 'pdf'>, reviewProfile: 'final' as const, commentProfile: 'hidden' as const, timeoutMs: 120_000 };
  const result = await renderDocxArtifacts(input, { ...options, unsupportedContent: 'warn' });
  expect(result.html).toContain('[EMF IMAGE]'); expect(result.html).not.toContain('data-docx-vector-frame');
  await expect(renderDocxArtifacts(input, { ...options, unsupportedContent: 'strict' })).rejects.toMatchObject({ code: 'resource_policy_failure' });
});

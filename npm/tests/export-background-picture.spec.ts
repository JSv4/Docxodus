import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { backgroundPictureDocx } from './docx-background-picture-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const variant of ['page', 'paragraph', 'relative'] as const) {
  test(`${variant}-anchored background preserves body space and repeats on every page`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const options = { browser, outputs: ['html', 'pdf'] as const, reviewProfile: 'markup' as const,
      markupPresentation: 'word' as const, commentProfile: 'margin' as const, commentColumnPlacement: 'balloons' as const };
    const origin = variant === 'relative' ? 'page' : variant;
    const result = await renderDocxArtifacts(backgroundPictureDocx(origin, false, variant === 'relative'), options);
    const control = await renderDocxArtifacts(backgroundPictureDocx(origin, true), options);
    expect(result.pageCount).toBe(2);
    expect(result.pageMap.fragments.filter(f => f.story === 'body')).toEqual(control.pageMap.fragments.filter(f => f.story === 'body'));
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      const text: string[] = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const pdfPage = await pdf.getPage(i);
        const content = await pdfPage.getTextContent();
        text.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
        expect((await pdfPage.getOperatorList()).fnArray.filter(op => op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject)).toHaveLength(1);
      }
      for (const label of ['first delivery.', 'second delivery.', 'Six crates accounted for.'])
        expect(text.join(' ').split(label)).toHaveLength(2);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const positions = await page.locator('.page-box').evaluateAll(boxes => boxes.map(box => {
      const anchor = box.querySelector<HTMLElement>(':scope > [data-docx-drawing-anchor]')!;
      const image = anchor.querySelector('img')!;
      const bounds = box.getBoundingClientRect(), rect = image.getBoundingClientRect();
      return { x: rect.left - bounds.left, y: rect.top - bounds.top, width: rect.width / bounds.width,
        height: rect.height / bounds.height, z: Number(getComputedStyle(anchor).zIndex) };
    }));
    for (const position of positions) {
      expect(Math.abs(position.x)).toBeLessThan(2); expect(Math.abs(position.y)).toBeLessThan(2);
      expect(position.width).toBeCloseTo(1, 2); expect(position.height).toBeCloseTo(1, 2);
      expect(position.z).toBeLessThan(0);
    }
  });
}

import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { positionedVmlShapeDocx } from './docx-positioned-vml-shape-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const kind of ['oval', 'rect'] as const) for (const positioned of [true, false]) {
  test(`${positioned ? 'positioned' : 'inline'} legacy ${kind} preserves the body and its placeholder`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(positionedVmlShapeDocx(kind, positioned), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden',
    });
    expect(result.pageCount).toBe(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdfPage = await (await task.promise).getPage(1);
      const text = (await pdfPage.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (const token of ['Maple inventory begins.', 'Maple inventory complete.', '[SHAPE]']) expect(text).toContain(token);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const shape = page.locator('[data-docx-shape="unrendered"]').first();
    expect(await shape.getAttribute('data-docx-drawing-anchor')).toBe(positioned ? 'true' : null);
    expect(await shape.evaluate(element => element.getBoundingClientRect().height * .75)).toBeCloseTo(positioned ? 550 : 72, 1);
    if (positioned) {
      expect(await shape.getAttribute('data-docx-anchor-behind')).toBe('true');
      expect(await shape.evaluate(element => getComputedStyle(element).position)).toBe('absolute');
    }
  });
}

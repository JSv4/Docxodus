import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { floatingShapeDocx } from './docx-floating-shape-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const height of [72, 760]) for (const behind of [false, true]) {
  test(`${height}pt drawing preserves ink and text with behindDoc=${behind}`, async ({ browser, page }) => {
    const bytes = floatingShapeDocx({ height, behind, wrap: 'topAndBottom', mixed: true, after: 'Harbor inventory complete.' });
    const result = await renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden' });
    expect(result.pageCount).toBe(1);
    expect(result.warnings.filter(warning => warning.code === 'drawing_clearance_relaxed')).toHaveLength(height === 760 ? 1 : 0);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdfPage = await (await task.promise).getPage(1);
      const text = (await pdfPage.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text).toContain('AnchorWord'); expect(text).toContain('Harbor inventory complete.');
      expect((await pdfPage.getOperatorList()).fnArray).toContain(OPS.constructPath);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const shapeHeight = await page.locator('[data-docx-shape-anchor]').evaluate(element => element.getBoundingClientRect().height * .75);
    expect(shapeHeight).toBeCloseTo(height, 1);
    const clearance = await page.locator('[data-docx-shape-clearance]').evaluate(element => element.getBoundingClientRect().height * .75);
    expect(clearance).toBeCloseTo(height === 760 ? 0 : height, 1);
  });
}

test('strict export rejects clearance that cannot fit any page', async ({ browser }) => {
  await expect(renderDocxArtifacts(floatingShapeDocx({ height: 760, behind: true }), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', unsupportedContent: 'strict',
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'output_verification' });
});

test('relaxing an impossible exclusion retains a fitting neighbor in the same paragraph', async ({ browser, page }) => {
  const result = await renderDocxArtifacts(floatingShapeDocx({ layers: [
    { order: 1, fill: 'E2EDE7', behind: true, height: 760 },
    { order: 2, fill: '19647E', behind: true, height: 72 },
  ] }), { browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden' });
  expect(result.warnings.filter(warning => warning.code === 'drawing_clearance_relaxed')).toHaveLength(1);
  await page.setContent(result.html!);
  expect(await page.locator('[data-docx-shape-clearance]').evaluate(element => element.getBoundingClientRect().height * .75)).toBeCloseTo(72, 1);
});

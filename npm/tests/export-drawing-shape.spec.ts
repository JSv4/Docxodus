import { expect, test } from '@playwright/test';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { SHAPE_VARIANTS, shapeDocx } from './docx-drawing-shape-fixture.js';

/**
 * PDF export of a paragraph holding only a DrawingML rectangle (alone, or in a group, optionally a
 * tracked deletion) used to fail with output_verification_failure: the rectangle produced no HTML,
 * so its paragraph had no measurable fragment (issue #1270). It now completes, with the rectangle.
 */

test.use({ launchOptions: {
  chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const variant of SHAPE_VARIANTS) {
  test(`${variant.name}: exports in markup mode with the rectangle and its line`, async ({ browser }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(shapeDocx(variant), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', commentProfile: 'hidden', timeoutMs: 120_000,
    });

    expect(result.pageCount).toBe(1);
    expect(result.pdf!.byteLength).toBeGreaterThan(0);
    expect(result.html).toContain('fill="#19647E"');
    // The rectangle's paragraph is at least the rectangle's 72pt tall.
    const heights = result.pageMap.fragments.map((fragment) => fragment.geometry.height);
    expect(Math.max(...heights)).toBeGreaterThanOrEqual(72);
  });
}

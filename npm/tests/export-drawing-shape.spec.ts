import { expect, test } from '@playwright/test';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { inlineDrawingDocx, SHAPE_VARIANTS, shapeDocx, UNDRAWABLE_DRAWINGS } from './docx-drawing-shape-fixture.js';

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
    expect(result.warnings.filter((warning) => warning.phase === 'chart_svg_materialization')).toEqual([]);
    expect(result.renderReport.unsupportedContent).toEqual([]);
    expect(result.pdf!.byteLength).toBeGreaterThan(0);
    expect(result.html).toContain('fill="#19647E"');
    // The rectangle's paragraph is at least the rectangle's 72pt tall.
    const heights = result.pageMap.fragments.map((fragment) => fragment.geometry.height);
    expect(Math.max(...heights)).toBeGreaterThanOrEqual(72);
  });
}

for (const [kind, drawing] of Object.entries(UNDRAWABLE_DRAWINGS)) {
  test(`${kind}: an inline drawing that cannot be drawn is reported, and rejected under strict`, async ({ browser }) => {
    test.setTimeout(120_000);
    const source = inlineDrawingDocx(drawing.uri, drawing.data, false, drawing.parts);
    const options = { browser, outputs: ['pdf'] as const, reviewProfile: 'markup' as const, commentProfile: 'hidden' as const,
      timeoutMs: 120_000 };

    // Its extent is kept (the export completes), and the loss is reported rather than silent.
    const result = await renderDocxArtifacts(source, options);
    expect(result.pageCount).toBe(1);
    expect(result.warnings.map((warning) => warning.code)).toContain('unsupported_content_placeholder');
    expect(result.renderReport.unsupportedContent).toEqual([
      expect.objectContaining({ contentType: 'Other', action: 'placeholder' }),
    ]);

    await expect(renderDocxArtifacts(source, { ...options, unsupportedContent: 'strict' }))
      .rejects.toMatchObject({ code: 'resource_policy_failure' });
  });
}

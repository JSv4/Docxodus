import { createHash } from 'node:crypto';
import { expect, test } from '@playwright/test';
import { canonicalJson } from './visual-parity/canonical-json.js';
import { assertSupportedPdfResult } from './visual-parity/pdf-result.js';

const sha256 = (value: Uint8Array | string): string =>
  createHash('sha256').update(value).digest('hex');

function validResult(): Record<string, any> {
  const pdf = new Uint8Array(32).fill(1);
  const fingerprint = 'a'.repeat(64);
  const pageMap = {
    schemaVersion: 1,
    mode: 'paginated',
    availability: 'available',
    documentVersion: 0,
    rendererFingerprint: fingerprint,
    pages: [{ pageNumber: 1, pageInSection: 1, width: 612, height: 792, sectionIndex: 0, pageName: 'page-1' }],
    fragments: [{
      fragmentId: 'p1-f0-body:document:p1',
      anchorId: 'body:document:p1',
      fragmentIndex: 0,
      pageNumber: 1,
      geometry: { x: 1, y: 2, width: 3, height: 4 },
      story: 'body',
      inTableCell: false,
    }],
  };
  return {
    pdf,
    pageCount: 1,
    pageMap,
    rendererFingerprint: fingerprint,
    warnings: [],
    renderReport: {
      status: 'complete',
      source: { rawPackageBytesDigest: 'b'.repeat(64) },
      options: { reviewProfile: 'final', commentProfile: 'hidden' },
      environment: { rendererFingerprint: fingerprint, verification: 'nodeVerified', fidelityTier: 'releaseBaselined' },
      // Deliberately a separate array: sharing the reference with pageMap.pages makes every
      // mutation below change both sides at once, so the report-vs-PageMap divergence check
      // could be deleted with both tests still green.
      pages: structuredClone(pageMap.pages),
      // Mirrors what CI actually reports: the contract substitutes Calibri with metric-compatible
      // Carlito, so `status` is "substituted", never "resolved".
      fonts: [{
        requestedFamily: 'Calibri',
        resolvedFamily: 'Carlito',
        status: 'substituted',
        source: 'configured',
        metricCompatible: true,
        glyphCoverage: 'complete',
      }],
      warnings: [],
      bindings: { pdfDigest: sha256(pdf), pageMapDigest: sha256(canonicalJson(pageMap)) },
    },
  };
}

const expectation = {
  sourceSha256: 'b'.repeat(64),
  reviewProfile: 'final',
  commentProfile: 'hidden',
};

test.describe('supported generated-PDF result verification', () => {
  test('reports a degraded font environment without failing the contract', () => {
    // WC036 uses a glyph Carlito lacks; 001-DeletedRun asks for "Calibri Light", which no
    // contract face matches metrically. Both are host font-set shortfalls, so they are surfaced
    // rather than blamed on the renderer.
    for (const degrade of [
      (v: Record<string, any>) => { v.renderReport.fonts[0].metricCompatible = false; },
      (v: Record<string, any>) => { v.renderReport.fonts[0].glyphCoverage = 'partial'; },
    ]) {
      const value = validResult();
      degrade(value);
      const verified = assertSupportedPdfResult(value, expectation);
      expect(verified.fontEnvironment.healthy).toBe(false);
      expect(verified.fontEnvironment.degraded).toHaveLength(1);
      expect(verified.fontEnvironment.degraded[0]).toMatchObject({
        requested: 'Calibri', resolved: 'Carlito',
      });
    }
    // A healthy substitution reports healthy.
    expect(assertSupportedPdfResult(validResult(), expectation).fontEnvironment)
      .toMatchObject({ healthy: true, faces: 1, degraded: [] });
  });

  test('accepts metric-compatible substitution, which the contract guarantees', () => {
    // `verification` stays "browserObserved" here on purpose: that label requires every face to
    // be "resolved", and a substitution contract can never produce it. Asserting it would gate on
    // an unreachable value — which is exactly what this harness used to do.
    const value = validResult();
    value.renderReport.environment.verification = 'browserObserved';
    expect(() => assertSupportedPdfResult(value, expectation)).not.toThrow();
  });

  test('accepts a fully bound result', () => {
    expect(assertSupportedPdfResult(validResult(), expectation)).toMatchObject({
      pageCount: 1,
      rendererFingerprint: 'a'.repeat(64),
    });
  });

  test('rejects independent count, fingerprint, PageMap, and report drift', () => {
    const mutations = [
      (value: Record<string, any>) => { value.pageCount = 2; },
      (value: Record<string, any>) => { value.pageMap.rendererFingerprint = 'c'.repeat(64); },
      (value: Record<string, any>) => { value.pageMap.pages[0].width = Number.POSITIVE_INFINITY; },
      (value: Record<string, any>) => { value.pageMap.fragments[0].pageNumber = 2; },
      // `verification` is no longer gated — see the acceptance test above. The tier still is,
      // and it is what rejects an injected browser.
      (value: Record<string, any>) => { value.renderReport.environment.fidelityTier = 'unbaselined'; },
      (value: Record<string, any>) => { value.renderReport.bindings.pdfDigest = 'd'.repeat(64); },
      (value: Record<string, any>) => { value.warnings.push({ code: 'unexpected' }); },
      // Only the report's inventory moves — the clause that exists to catch exactly this.
      (value: Record<string, any>) => { value.renderReport.pages[0].width = 611; },
      (value: Record<string, any>) => { value.renderReport.pages.push(value.renderReport.pages[0]); },
      // The faces must come from the configured directories, not whatever the browser had.
      (value: Record<string, any>) => { value.renderReport.fonts[0].source = 'browser'; },
      // A substitution that changes metrics is a layout change wearing a font's name.
      (value: Record<string, any>) => { value.renderReport.fonts[0].status = 'missing'; },
      (value: Record<string, any>) => { value.renderReport.fonts = []; },
    ];
    for (const mutate of mutations) {
      const value = validResult();
      mutate(value);
      expect(() => assertSupportedPdfResult(value, expectation)).toThrow();
    }
  });
});

import assert from "node:assert/strict";
import { test } from "node:test";
import {
  CURRENT_RENDER_REPORT_SCHEMA,
  CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  hasCurrentRenderReportDiscriminator,
  isCurrentCompleteRenderReport,
} from "../dist/contracts.js";

test("accepts only the current v2 browser-materializer report discriminator", () => {
  const hash = "0".repeat(64);
  const limits = Object.fromEntries([
    "compressedDocxBytes", "opcEntries", "expandedOpcBytes", "xmlPartBytes",
    "opcUriCharacters", "opcCompressionRatio", "htmlOutputBytes", "pdfOutputBytes",
    "pageMapOutputBytes", "renderReportOutputBytes", "pdfParserExpandedBytes", "finalPages",
    "domNodes", "automaticResources", "automaticResourceBytes", "renderDiagnostics",
    "fontDirectoryEntries", "fontFiles", "fontFileBytes", "fontTotalBytes", "fontRequests",
    "fontSampleCodePoints",
  ].map((key) => [key, 1]));
  const valid = {
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: CURRENT_RENDER_REPORT_SCHEMA_VERSION,
    source: { rawPackageBytesDigest: hash, byteLength: 1, documentVersion: 0 },
    options: {
      reviewProfile: "markup",
      reviewProfileAlreadyApplied: false,
      commentProfile: "hidden",
      title: "test",
      outputs: [],
      layoutDigest: hash,
      runtimePolicyDigest: hash,
      policy: { unsupportedContent: "warn", strictFonts: false, timeoutMs: 1, limits },
    },
    readiness: [{
      phase: "pagination",
      status: "complete",
      elapsedMs: 0,
      pending: [],
      diagnostics: [
        "sections_processed", "page_runs_processed", "source_anchors_inventoried",
        "note_references_inventoried",
      ].map((code) => ({ code, severity: "info", message: code, count: 0 })),
    }],
    fonts: [],
    resources: [],
    unsupportedContent: [],
    fontIdentity: { schemaVersion: 1, digest: hash, verification: "browserObserved" },
    warnings: [],
    status: "complete",
    environment: {
      rendererFingerprint: hash,
      verification: "browserObserved",
      fidelityTier: "experimental",
      observed: {
        runtimeKind: "browser",
        locale: "en-US",
        timezone: "UTC",
        viewport: [1, 1],
        deviceScaleFactor: 1,
        media: {
          colorScheme: "light",
          reducedMotion: "reduce",
          forcedColors: "none",
          printMedia: true,
        },
        networkIsolation: "contextRestricted",
      },
    },
    pages: [{ pageNumber: 1, pageInSection: 1, pageName: "test", width: 1, height: 1 }],
    bindings: { pageMapDigest: hash, artifactRequestIds: [] },
  };
  assert.equal(CURRENT_RENDER_REPORT_SCHEMA,
    "https://docxodus.dev/schemas/render/render-report/v2");
  assert.equal(CURRENT_RENDER_REPORT_SCHEMA_VERSION, 2);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  }), true);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: "https://docxodus.dev/schemas/render/render-report/v1",
    schemaVersion: 1,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: 1,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: "https://docxodus.dev/schemas/render/render-report/v1",
    schemaVersion: 2,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator(null), false);
  assert.equal(isCurrentCompleteRenderReport({
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  }), false);
  assert.equal(isCurrentCompleteRenderReport(valid), true);
  const incomplete = structuredClone(valid);
  delete incomplete.source;
  assert.equal(isCurrentCompleteRenderReport(incomplete), false);
  const badFont = structuredClone(valid);
  badFont.fonts = [{
    requestKey: "not-a-digest",
    requestedFamily: "serif",
    status: "unverified",
    source: "browser",
  }];
  assert.equal(isCurrentCompleteRenderReport(badFont), false);
  const badResource = structuredClone(valid);
  badResource.resources = [{
    kind: "chart",
    status: "inline",
    readiness: "failed",
    contentKey: hash,
  }];
  assert.equal(isCurrentCompleteRenderReport(badResource), false);
});

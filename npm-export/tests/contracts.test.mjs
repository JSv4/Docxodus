import assert from "node:assert/strict";
import { test } from "node:test";
import {
  CURRENT_RENDER_REPORT_SCHEMA,
  CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  fromBrowserFailure,
  hasCurrentRenderReportDiscriminator,
  isCurrentCompleteRenderReport,
  isCurrentFailedRenderReport,
} from "../dist/contracts.js";

test("accepts only a closed current v3 browser-materializer report", () => {
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
    fontReadiness: [],
    resources: [],
    unsupportedContent: [],
    fontIdentity: {
      resolverContract: "https://docxodus.dev/contracts/font-resolver/v1",
      substitutionContractVersion: 1,
      substitutionContractDigest: hash,
      resolutionDigest: hash,
    },
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
    "https://docxodus.dev/schemas/render/render-report/v3");
  assert.equal(CURRENT_RENDER_REPORT_SCHEMA_VERSION, 3);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  }), true);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: "https://docxodus.dev/schemas/render/render-report/v1",
    schemaVersion: 1,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: "https://docxodus.dev/schemas/render/render-report/v2",
    schemaVersion: 2,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: CURRENT_RENDER_REPORT_SCHEMA,
    schemaVersion: 1,
  }), false);
  assert.equal(hasCurrentRenderReportDiscriminator({
    schema: "https://docxodus.dev/schemas/render/render-report/v2",
    schemaVersion: 3,
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
    requestId: "font-0001",
    requestedFamily: "serif",
    requestedFamilies: ["serif"],
    requestedStyle: "normal",
    requestedWeight: 400,
    requestedStretch: 100,
    sampleCodePointCount: 1,
    sampleDigest: "not-a-digest",
    status: "unverified",
    source: "browser",
    glyphCoverage: "unverified",
  }];
  assert.equal(isCurrentCompleteRenderReport(badFont), false);
  const badReadiness = structuredClone(valid);
  badReadiness.fontReadiness = [{
    requestKey: hash,
    requestedFamily: "serif",
    available: true,
    bytesBase64: "forbidden",
  }];
  assert.equal(isCurrentCompleteRenderReport(badReadiness), false);
  const duplicateReadiness = structuredClone(valid);
  duplicateReadiness.fontReadiness = [
    { requestKey: hash, requestedFamily: "serif", available: true },
    { requestKey: hash, requestedFamily: "sans-serif", available: true },
  ];
  assert.equal(isCurrentCompleteRenderReport(duplicateReadiness), false);
  const oldIdentity = structuredClone(valid);
  oldIdentity.fontIdentity = { schemaVersion: 1, digest: hash, verification: "browserObserved" };
  assert.equal(isCurrentCompleteRenderReport(oldIdentity), false);
  const badResource = structuredClone(valid);
  badResource.resources = [{
    kind: "chart",
    status: "inline",
    readiness: "failed",
    contentKey: hash,
  }];
  assert.equal(isCurrentCompleteRenderReport(badResource), false);

  const {
    status: _status,
    environment,
    pages,
    bindings,
    readiness: _readiness,
    ...failedBase
  } = structuredClone(valid);
  const failed = {
    ...failedBase,
    readiness: [{
      phase: "font_loading",
      status: "failed",
      elapsedMs: 1,
      pending: ["font:configured"],
    }],
    status: "failed",
    failure: {
      code: "resource_policy_failure",
      severity: "error",
      phase: "font_loading",
      message: "Font policy failed.",
      remediation: "Supply a permitted font.",
      pending: ["font:configured"],
    },
    environment: { verification: environment.verification },
    partial: { pages, bindings },
    unavailable: [],
  };
  assert.equal(isCurrentFailedRenderReport(failed), true);
  assert.equal(isCurrentCompleteRenderReport(failed), false);

  const twoTerminalPhases = structuredClone(failed);
  twoTerminalPhases.readiness.push({
    phase: "output_verification",
    status: "cancelled",
    elapsedMs: 0,
    pending: [],
  });
  assert.equal(isCurrentFailedRenderReport(twoTerminalPhases), false);
  const malformedFailure = structuredClone(failed);
  malformedFailure.failure.bytesBase64 = "forbidden";
  assert.equal(isCurrentFailedRenderReport(malformedFailure), false);

  const acceptedError = fromBrowserFailure({
    code: "resource_policy_failure",
    phase: "font_loading",
    message: "Font policy failed.",
    remediation: "Supply a permitted font.",
    report: failed,
  });
  assert.notEqual(acceptedError.report, failed);
  assert.equal(acceptedError.report?.status, "failed");
  failed.failure.message = "mutated after validation";
  assert.equal(acceptedError.report?.failure.message, "Font policy failed.");
  const rejectedError = fromBrowserFailure({
    code: "resource_policy_failure",
    phase: "font_loading",
    report: malformedFailure,
  });
  assert.equal(rejectedError.report, undefined);
});

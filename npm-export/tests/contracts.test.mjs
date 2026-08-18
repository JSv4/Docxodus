import assert from "node:assert/strict";
import { test } from "node:test";
import {
  CURRENT_RENDER_REPORT_SCHEMA,
  CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  fromBrowserFailure,
  hasCurrentRenderReportDiscriminator,
  isCurrentCompleteRenderReport,
  isCurrentFailedRenderReport,
  isCurrentPageMap,
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
  const validPageMap = {
    schemaVersion: 1,
    mode: "paginated",
    availability: "available",
    documentVersion: 0,
    rendererFingerprint: hash,
    pages: [{ ...structuredClone(valid.pages[0]), sectionIndex: 0 }],
    fragments: [{
      fragmentId: "p1-f0-a",
      anchorId: "a",
      fragmentIndex: 0,
      pageNumber: 1,
      geometry: { x: 0, y: 0, width: 1, height: 1 },
      story: "body",
      inTableCell: false,
    }],
  };
  assert.equal(isCurrentPageMap(validPageMap, limits), true);
  for (const mutate of [
    (map) => { map.path = "/private/source"; },
    (map) => { map.schemaVersion = 2; },
    (map) => { map.mode = "continuous"; },
    (map) => { map.fragments.push(structuredClone(map.fragments[0])); },
    (map) => { map.fragments[0].fragmentIndex = 1; },
    (map) => { map.fragments[0].geometry.x = 2; },
  ]) {
    const malformed = structuredClone(validPageMap);
    mutate(malformed);
    assert.equal(isCurrentPageMap(malformed, limits), false);
  }
  const raisedPolicyLimit = structuredClone(valid);
  raisedPolicyLimit.options.policy.limits.fontRequests = Number.MAX_SAFE_INTEGER;
  assert.equal(isCurrentCompleteRenderReport(raisedPolicyLimit), false);
  const incomplete = structuredClone(valid);
  delete incomplete.source;
  assert.equal(isCurrentCompleteRenderReport(incomplete), false);
  const badFont = structuredClone(valid);
  badFont.fonts = [{
    requestId: "font-0001",
    requestedFamily: "serif",
    requestedFamilies: ["serif"],
    requestedFamilyKinds: ["generic"],
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
  const overlongFontRequestId = structuredClone(valid);
  overlongFontRequestId.fonts = [{
    ...badFont.fonts[0],
    requestId: `font-${"1".repeat(124)}`,
    sampleDigest: hash,
  }];
  assert.equal(overlongFontRequestId.fonts[0].requestId.length, 129);
  assert.equal(isCurrentCompleteRenderReport(overlongFontRequestId), false);
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
  const earlyFinalFailure = structuredClone(failed);
  earlyFinalFailure.options.reviewProfile = "final";
  earlyFinalFailure.options.reviewProfileAlreadyApplied = false;
  delete earlyFinalFailure.derivedProfileSource;
  assert.equal(isCurrentFailedRenderReport(earlyFinalFailure), true);
  const forbiddenMarkupDerived = structuredClone(failed);
  forbiddenMarkupDerived.derivedProfileSource = {
    rawPackageBytesDigest: hash,
    byteLength: 1,
  };
  assert.equal(isCurrentFailedRenderReport(forbiddenMarkupDerived), false);

  const twoTerminalPhases = structuredClone(failed);
  twoTerminalPhases.readiness.push({
    phase: "output_verification",
    status: "cancelled",
    elapsedMs: 0,
    pending: [],
  });
  assert.equal(isCurrentFailedRenderReport(twoTerminalPhases), false);
  const cancelledNonCancellation = structuredClone(failed);
  cancelledNonCancellation.readiness[0].status = "cancelled";
  assert.equal(isCurrentFailedRenderReport(cancelledNonCancellation), false);
  const failedCancellation = structuredClone(failed);
  failedCancellation.failure.code = "operation_cancelled";
  assert.equal(isCurrentFailedRenderReport(failedCancellation), false);
  failedCancellation.readiness[0].status = "cancelled";
  assert.equal(isCurrentFailedRenderReport(failedCancellation), true);
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
  const contradictoryError = fromBrowserFailure({
    code: "operation_cancelled",
    phase: "cleanup",
    message: "Contradictory bridge message.",
    remediation: "Contradictory bridge remediation.",
    report: acceptedError.report,
  });
  assert.equal(contradictoryError.code, "resource_policy_failure");
  assert.equal(contradictoryError.phase, "font_loading");
  assert.equal(contradictoryError.message, "Font policy failed.");
  assert.equal(contradictoryError.remediation, "Supply a permitted font.");
  const reboundError = fromBrowserFailure({
    code: "resource_policy_failure",
    phase: "font_loading",
    report: acceptedError.report,
  }, ["pdf"]);
  assert.deepEqual(reboundError.report?.options.outputs, ["pdf"]);
  assert.equal(reboundError.report?.unavailable.find(
    ({ field }) => field === "bindings.htmlDigest",
  )?.reasonCode, "notRequested");
  assert.equal(reboundError.report?.unavailable.find(
    ({ field }) => field === "bindings.pdfDigest",
  )?.reasonCode, "notReached");
  const verificationFailure = structuredClone(acceptedError.report);
  verificationFailure.failure.phase = "output_verification";
  verificationFailure.readiness[0].phase = "output_verification";
  verificationFailure.partial.bindings.htmlDigest = hash;
  const verificationError = fromBrowserFailure({ report: verificationFailure }, ["pdf"]);
  assert.equal(verificationError.report?.unavailable.find(
    ({ field }) => field === "bindings.pdfDigest",
  )?.reasonCode, "failedVerification");
  const dualPartial = fromBrowserFailure({ report: verificationFailure }, ["html", "pdf"]);
  assert.equal(dualPartial.report?.partial?.bindings?.htmlDigest, hash);
  assert.equal(dualPartial.report?.unavailable.some(
    ({ field }) => field === "bindings.htmlDigest",
  ), false);
  const forgedBrowserPdf = structuredClone(verificationFailure);
  forgedBrowserPdf.partial.bindings.pdfDigest = "3".repeat(64);
  forgedBrowserPdf.partial.bindings.pdfByteDeterministic = false;
  forgedBrowserPdf.partial.bindings.volatilePdfMetadata = {};
  const strippedBrowserPdf = fromBrowserFailure({ report: forgedBrowserPdf }, ["pdf"]);
  assert.equal(strippedBrowserPdf.report?.partial?.bindings?.pdfDigest, undefined);
  assert.equal(strippedBrowserPdf.report?.partial?.bindings?.pdfByteDeterministic, undefined);
  assert.equal(strippedBrowserPdf.report?.partial?.bindings?.volatilePdfMetadata, undefined);
  const rejectedError = fromBrowserFailure({
    code: "resource_policy_failure",
    phase: "font_loading",
    report: malformedFailure,
  });
  assert.equal(rejectedError.report, undefined);

  const expectedFailureContract = {
    source: structuredClone(acceptedError.report.source),
    options: structuredClone(acceptedError.report.options),
    retainedPolicy: { strictFonts: true, runtimePolicyDigest: "1".repeat(64) },
  };
  const reboundStrict = fromBrowserFailure({ report: acceptedError.report }, ["pdf"],
    expectedFailureContract);
  assert.equal(reboundStrict.report?.options.policy.strictFonts, true);
  assert.equal(reboundStrict.report?.options.runtimePolicyDigest, "1".repeat(64));
  for (const mutate of [
    (report) => { report.source.rawPackageBytesDigest = "2".repeat(64); },
    (report) => { report.options.title = "forged title"; },
    (report) => { report.options.policy.limits.fontRequests += 1; },
  ]) {
    const forged = structuredClone(acceptedError.report);
    mutate(forged);
    assert.equal(fromBrowserFailure({ report: forged }, ["pdf"], expectedFailureContract).report,
      undefined);
  }
});

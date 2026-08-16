import type { Browser } from "playwright-core";
import type {
  CommentProfile,
  CompleteRenderReport,
  DocxodusExportErrorCode,
  EnvironmentVerification,
  ExportPhase,
  ExportResourceLimits,
  FailedRenderReport,
  FontFaceStyle,
  FontResolver,
  PaginatedHtmlOptions,
  PaginatedHtmlResult,
  PaginatedRenderMetadata,
  RenderReport,
  RenderWarning,
  ReviewProfile,
} from "docxodus/export-browser";

export const CURRENT_RENDER_REPORT_SCHEMA =
  "https://docxodus.dev/schemas/render/render-report/v3" as const;
export const CURRENT_RENDER_REPORT_SCHEMA_VERSION = 3 as const;

export function hasCurrentRenderReportDiscriminator(value: unknown): value is {
  schema: typeof CURRENT_RENDER_REPORT_SCHEMA;
  schemaVersion: typeof CURRENT_RENDER_REPORT_SCHEMA_VERSION;
} {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    && (value as { schema?: unknown }).schema === CURRENT_RENDER_REPORT_SCHEMA
    && (value as { schemaVersion?: unknown }).schemaVersion === CURRENT_RENDER_REPORT_SCHEMA_VERSION;
}

function reportRecord(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function reportKeys(value: Record<string, unknown>, allowed: readonly string[]): boolean {
  return Object.keys(value).every((key) => allowed.includes(key));
}

function reportStrings(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === "string");
}

export function isCurrentCompleteRenderReport(value: unknown): value is CompleteRenderReport {
  if (!hasCurrentRenderReportDiscriminator(value)) return false;
  const report = value as unknown as Record<string, unknown>;
  const digest = (candidate: unknown): boolean =>
    typeof candidate === "string" && /^[0-9a-f]{64}$/.test(candidate);
  const integer = (candidate: unknown, minimum = 0): candidate is number =>
    Number.isSafeInteger(candidate) && (candidate as number) >= minimum;
  const finite = (candidate: unknown, minimum = 0): candidate is number =>
    typeof candidate === "number" && Number.isFinite(candidate) && candidate >= minimum;
  if (!reportKeys(report, [
    "schema", "schemaVersion", "source", "derivedProfileSource", "options", "readiness",
    "fonts", "fontReadiness", "resources", "unsupportedContent", "fontIdentity", "warnings",
    "status", "environment", "pages", "bindings",
  ])) return false;

  const source = reportRecord(report.source);
  if (!source || !reportKeys(source, ["rawPackageBytesDigest", "byteLength", "documentVersion"])
    || !digest(source.rawPackageBytesDigest) || !integer(source.byteLength, 1)
    || !integer(source.documentVersion)) return false;
  const derived = report.derivedProfileSource === undefined
    ? undefined : reportRecord(report.derivedProfileSource);
  if (report.derivedProfileSource !== undefined && (!derived
    || !reportKeys(derived, ["rawPackageBytesDigest", "byteLength"])
    || !digest(derived.rawPackageBytesDigest) || !integer(derived.byteLength, 1))) return false;

  const options = reportRecord(report.options);
  const policy = options ? reportRecord(options.policy) : undefined;
  const limits = policy ? reportRecord(policy.limits) : undefined;
  const limitKeys = [
    "compressedDocxBytes", "opcEntries", "expandedOpcBytes", "xmlPartBytes",
    "opcUriCharacters", "opcCompressionRatio", "htmlOutputBytes", "pdfOutputBytes",
    "pageMapOutputBytes", "renderReportOutputBytes", "pdfParserExpandedBytes", "finalPages",
    "domNodes", "automaticResources", "automaticResourceBytes", "renderDiagnostics",
    "fontDirectoryEntries", "fontFiles", "fontFileBytes", "fontTotalBytes", "fontRequests",
    "fontSampleCodePoints",
  ] as const;
  const outputKey = JSON.stringify(options?.outputs);
  if (!options || !reportKeys(options, [
    "reviewProfile", "reviewProfileAlreadyApplied", "commentProfile", "title", "outputs",
    "layoutDigest", "runtimePolicyDigest", "policy",
  ]) || !["final", "original", "markup"].includes(String(options.reviewProfile))
    || typeof options.reviewProfileAlreadyApplied !== "boolean"
    || !["hidden", "inline", "endnotes", "margin"].includes(String(options.commentProfile))
    || typeof options.title !== "string"
    || !["[]", '["html"]', '["pdf"]', '["html","pdf"]'].includes(outputKey)
    || !digest(options.layoutDigest) || !digest(options.runtimePolicyDigest)
    || !policy || !reportKeys(policy, ["unsupportedContent", "strictFonts", "timeoutMs", "limits"])
    || !["warn", "strict"].includes(String(policy.unsupportedContent))
    || typeof policy.strictFonts !== "boolean" || !integer(policy.timeoutMs, 1)
    || !limits || !reportKeys(limits, limitKeys)
    || !limitKeys.every((key) => integer(limits[key], 1))) return false;
  const profileRequiresDerived = ["final", "original"].includes(String(options.reviewProfile))
    && options.reviewProfileAlreadyApplied === false;
  if (profileRequiresDerived !== (derived !== undefined)
    || (options.reviewProfile === "markup" && derived !== undefined)
    || (options.reviewProfileAlreadyApplied === true && derived !== undefined)) return false;

  const phases = new Set([
    "input_validation", "package_preflight", "browser_launch", "wasm_initialization",
    "docx_conversion", "font_loading", "image_decoding", "chart_svg_materialization",
    "pagination", "running_story_placement", "page_tree_stability", "pdf_print",
    "output_verification", "output_write", "filesystem_commit", "cleanup",
  ]);
  const diagnosticCodes = new Set([
    "sections_processed", "page_runs_processed", "source_anchors_inventoried",
    "note_references_inventoried",
  ]);
  if (!Array.isArray(report.readiness) || report.readiness.some((entry) => {
    const item = reportRecord(entry);
    if (!item || !reportKeys(item, ["phase", "status", "elapsedMs", "pending", "diagnostics"])
      || !phases.has(String(item.phase)) || item.status !== "complete"
      || !finite(item.elapsedMs) || !reportStrings(item.pending)) return true;
    if (item.phase !== "pagination") return item.diagnostics !== undefined;
    if (!Array.isArray(item.diagnostics) || item.diagnostics.length !== 4) return true;
    const seen = new Set<string>();
    return item.diagnostics.some((diagnostic) => {
      const row = reportRecord(diagnostic);
      if (!row || !reportKeys(row, ["code", "severity", "message", "count"])
        || !diagnosticCodes.has(String(row.code)) || seen.has(String(row.code))
        || row.severity !== "info" || typeof row.message !== "string" || row.message.length === 0
        || !integer(row.count)) return true;
      seen.add(String(row.code));
      return false;
    });
  }) || report.readiness.filter((entry) => reportRecord(entry)?.phase === "pagination").length !== 1) {
    return false;
  }

  const fontKeys: string[] = [];
  const fontsValid = Array.isArray(report.fonts) && report.fonts.every((font) => {
    const item = reportRecord(font);
    if (!item || !reportKeys(item, [
      "requestId", "requestedFamily", "requestedFamilies", "requestedStyle", "requestedWeight",
      "requestedStretch", "sampleCodePointCount", "sampleDigest", "resolvedFamily", "resolvedFace",
      "status", "source", "format", "fileSha256", "version", "faceMatch", "metricCompatible",
      "glyphCoverage", "missingCodePointCount", "browserFallbackAvailable", "licenseEvidence",
    ]) || typeof item.requestId !== "string" || !/^font-[0-9]{4,}$/.test(item.requestId)
      || typeof item.requestedFamily !== "string" || item.requestedFamily.length === 0
      || item.requestedFamily.length > 512
      || !reportStrings(item.requestedFamilies) || item.requestedFamilies.length === 0
      || item.requestedFamilies.length > 256
      || item.requestedFamilies.some((family) => family.length === 0 || family.length > 512)
      || !item.requestedFamilies.includes(item.requestedFamily)
      || !["normal", "italic", "oblique"].includes(String(item.requestedStyle))
      || !integer(item.requestedWeight, 1) || (item.requestedWeight as number) > 1000
      || !finite(item.requestedStretch, 50) || (item.requestedStretch as number) > 200
      || !integer(item.sampleCodePointCount)
      || (item.sampleCodePointCount as number) > (limits?.fontSampleCodePoints as number)
      || !digest(item.sampleDigest)
      || (item.resolvedFamily !== undefined
        && (typeof item.resolvedFamily !== "string" || item.resolvedFamily.length === 0
          || item.resolvedFamily.length > 512))
      || (item.resolvedFace !== undefined
        && (typeof item.resolvedFace !== "string" || item.resolvedFace.length === 0
          || item.resolvedFace.length > 512))
      || !["resolved", "substituted", "missing", "load_failed", "unverified"].includes(String(item.status))
      || !["browser", "configured", "attested"].includes(String(item.source))
      || (item.format !== undefined && !["ttf", "otf", "woff", "woff2"].includes(String(item.format)))
      || (item.fileSha256 !== undefined && !digest(item.fileSha256))
      || (item.version !== undefined
        && (typeof item.version !== "string" || item.version.length === 0 || item.version.length > 512))
      || (item.faceMatch !== undefined && !["exact", "synthesized"].includes(String(item.faceMatch)))
      || (item.metricCompatible !== undefined && typeof item.metricCompatible !== "boolean")
      || (item.glyphCoverage !== undefined
        && !["complete", "partial", "unverified"].includes(String(item.glyphCoverage)))
      || (item.missingCodePointCount !== undefined
        && (!integer(item.missingCodePointCount)
          || (item.missingCodePointCount as number) > (item.sampleCodePointCount as number)))
      || (item.browserFallbackAvailable !== undefined
        && typeof item.browserFallbackAvailable !== "boolean")) return false;
    const license = item.licenseEvidence === undefined
      ? undefined : reportRecord(item.licenseEvidence);
    if (item.licenseEvidence !== undefined && (!license
      || !reportKeys(license, ["kind", "identity", "noSubsetting"])
      || !["installable", "previewPrint", "editable", "attested"].includes(String(license.kind))
      || !digest(license.identity) || typeof license.noSubsetting !== "boolean")) return false;
    const selected = item.status === "resolved" || item.status === "substituted"
      || item.status === "load_failed";
    if (item.source === "browser") {
      if ((item.status !== "missing" && item.status !== "unverified")
        || item.glyphCoverage !== "unverified"
        || ["resolvedFamily", "resolvedFace", "format", "fileSha256", "version", "faceMatch",
          "metricCompatible", "missingCodePointCount", "licenseEvidence"].some((key) =>
          item[key] !== undefined)) return false;
    } else if (!selected || typeof item.resolvedFamily !== "string"
      || typeof item.resolvedFace !== "string" || !digest(item.fileSha256)
      || typeof item.version !== "string") return false;
    if (item.source === "configured" && (!license || item.format === undefined)) return false;
    if ((item.status === "missing" || item.status === "unverified") && item.source !== "browser") return false;
    if (item.glyphCoverage === "partial"
      && (!integer(item.missingCodePointCount, 1))) return false;
    if (item.glyphCoverage === "complete" && item.missingCodePointCount !== undefined) return false;
    fontKeys.push(item.requestId);
    return true;
  }) && report.fonts.length <= (limits?.fontRequests as number)
    && new Set(fontKeys).size === fontKeys.length;
  const readinessKeys: string[] = [];
  const fontReadinessValid = Array.isArray(report.fontReadiness)
    && report.fontReadiness.length <= (limits?.fontRequests as number)
    && report.fontReadiness.every((probe) => {
      const item = reportRecord(probe);
      if (!item || !reportKeys(item, ["requestKey", "requestedFamily", "available"])
        || !digest(item.requestKey) || typeof item.requestedFamily !== "string"
        || item.requestedFamily.length === 0 || item.requestedFamily.length > 512
        || typeof item.available !== "boolean") return false;
      readinessKeys.push(item.requestKey as string);
      return true;
    }) && new Set(readinessKeys).size === readinessKeys.length;
  const resourcesValid = Array.isArray(report.resources) && report.resources.every((resource) => {
    const item = reportRecord(resource);
    if (!item || !reportKeys(item, [
      "kind", "status", "readiness", "contentKey", "resource", "anchorId", "message",
      "mediaType", "byteLength",
    ]) || (item.resource !== undefined && typeof item.resource !== "string")
      || (item.anchorId !== undefined && (typeof item.anchorId !== "string" || item.anchorId.length === 0))
      || (item.message !== undefined && (typeof item.message !== "string" || item.message.length === 0))
      || (item.mediaType !== undefined
        && (typeof item.mediaType !== "string" || item.mediaType.length === 0))
      || (item.byteLength !== undefined && !integer(item.byteLength))) return false;
    return item.kind === "external_link"
      ? ["allowed_user_link", "omitted"].includes(String(item.status))
        && item.readiness === undefined && item.contentKey === undefined
      : ["image", "svg", "chart"].includes(String(item.kind))
        && digest(item.contentKey)
        && (item.readiness === "complete"
          ? (item.kind === "image" ? item.status === "embedded" : item.status === "inline")
          : item.readiness === "failed" && item.status === "omitted");
  });
  const unsupportedValid = Array.isArray(report.unsupportedContent)
    && report.unsupportedContent.every((entry) => {
      const item = reportRecord(entry);
      return !!item && reportKeys(item, ["contentType", "elementName", "anchorId", "action"])
        && typeof item.contentType === "string" && item.contentType.length > 0
        && (item.elementName === undefined
          || (typeof item.elementName === "string" && item.elementName.length > 0))
        && (item.anchorId === undefined || (typeof item.anchorId === "string" && item.anchorId.length > 0))
        && item.action === "placeholder";
    });
  const warningsValid = Array.isArray(report.warnings) && report.warnings.every((entry) => {
    const item = reportRecord(entry);
    return !!item && reportKeys(item, [
      "code", "severity", "phase", "message", "remediation", "detail", "partUri",
      "anchorId", "resource",
    ]) && typeof item.code === "string" && item.code.length > 0 && item.severity === "warning"
      && phases.has(String(item.phase)) && typeof item.message === "string" && item.message.length > 0
      && typeof item.remediation === "string" && item.remediation.length > 0
      && ["detail", "partUri", "anchorId", "resource"].every((key) =>
        item[key] === undefined || typeof item[key] === "string");
  });
  const fontIdentity = reportRecord(report.fontIdentity);
  const environment = reportRecord(report.environment);
  const observed = environment ? reportRecord(environment.observed) : undefined;
  const media = observed ? reportRecord(observed.media) : undefined;
  const attested = environment?.attested === undefined
    ? undefined : reportRecord(environment.attested);
  const pagesValid = Array.isArray(report.pages) && report.pages.length > 0
    && report.pages.every((entry) => {
      const item = reportRecord(entry);
      return !!item && reportKeys(item, [
        "pageNumber", "pageInSection", "pageName", "width", "height", "sectionIndex",
      ]) && integer(item.pageNumber, 1) && integer(item.pageInSection, 1)
        && typeof item.pageName === "string" && item.pageName.length > 0
        && finite(item.width, Number.MIN_VALUE) && finite(item.height, Number.MIN_VALUE)
        && (item.sectionIndex === undefined || integer(item.sectionIndex));
    });
  const bindings = reportRecord(report.bindings);
  const volatileMetadata = bindings?.volatilePdfMetadata === undefined
    ? undefined : reportRecord(bindings.volatilePdfMetadata);
  const bindingsValid = !!bindings && reportKeys(bindings, [
    "pageMapDigest", "htmlDigest", "pdfDigest", "artifactRequestIds", "pdfByteDeterministic",
    "volatilePdfMetadata",
  ]) && digest(bindings.pageMapDigest) && reportStrings(bindings.artifactRequestIds)
    && (bindings.htmlDigest === undefined || digest(bindings.htmlDigest))
    && (bindings.pdfDigest === undefined || digest(bindings.pdfDigest))
    && (bindings.pdfByteDeterministic === undefined || bindings.pdfByteDeterministic === false)
    && (bindings.volatilePdfMetadata === undefined || (!!volatileMetadata
      && Object.values(volatileMetadata).every((entry) => typeof entry === "string")))
    && (outputKey.includes("html") ? bindings.htmlDigest !== undefined : bindings.htmlDigest === undefined)
    && (outputKey.includes("pdf")
      ? bindings.pdfDigest !== undefined && bindings.pdfByteDeterministic === false
        && volatileMetadata !== undefined
      : bindings.pdfDigest === undefined && bindings.pdfByteDeterministic === undefined
        && bindings.volatilePdfMetadata === undefined);
  return report.status === "complete"
    && !!fontIdentity && reportKeys(fontIdentity, [
      "resolverContract", "substitutionContractVersion", "substitutionContractDigest",
      "resolutionDigest",
    ])
    && fontIdentity.resolverContract === "https://docxodus.dev/contracts/font-resolver/v1"
    && fontIdentity.substitutionContractVersion === 1
    && digest(fontIdentity.substitutionContractDigest) && digest(fontIdentity.resolutionDigest)
    && !!environment && reportKeys(environment, [
      "rendererFingerprint", "verification", "fidelityTier", "observed", "attested",
      "attestationDigest",
    ]) && digest(environment.rendererFingerprint)
    && ["nodeVerified", "browserObserved", "callerAttested"].includes(String(environment.verification))
    && ["releaseBaselined", "experimental", "unbaselined"].includes(String(environment.fidelityTier))
    && !!observed && reportKeys(observed, [
      "runtimeKind", "playwrightVersion", "browserProduct", "browserBuild", "executableSha256",
      "launchFlags", "operatingSystem", "architecture", "locale", "timezone", "viewport",
      "deviceScaleFactor", "media", "networkIsolation",
    ]) && ["browser", "nodeChromium"].includes(String(observed.runtimeKind))
    && (observed.playwrightVersion === undefined
      || (typeof observed.playwrightVersion === "string" && observed.playwrightVersion.length > 0))
    && (observed.browserProduct === undefined
      || (typeof observed.browserProduct === "string" && observed.browserProduct.length > 0))
    && (observed.browserBuild === undefined
      || (typeof observed.browserBuild === "string" && observed.browserBuild.length > 0))
    && (observed.executableSha256 === undefined || digest(observed.executableSha256))
    && (observed.launchFlags === undefined || (reportStrings(observed.launchFlags)
      && observed.launchFlags.every((flag) => flag.length > 0)
      && new Set(observed.launchFlags).size === observed.launchFlags.length))
    && (observed.operatingSystem === undefined
      || (typeof observed.operatingSystem === "string" && observed.operatingSystem.length > 0))
    && (observed.architecture === undefined
      || (typeof observed.architecture === "string" && observed.architecture.length > 0))
    && typeof observed.locale === "string" && observed.locale.length > 0
    && typeof observed.timezone === "string" && observed.timezone.length > 0
    && Array.isArray(observed.viewport) && observed.viewport.length === 2
    && observed.viewport.every((entry) => finite(entry, Number.MIN_VALUE))
    && finite(observed.deviceScaleFactor, Number.MIN_VALUE)
    && !!media && reportKeys(media, ["colorScheme", "reducedMotion", "forcedColors", "printMedia"])
    && ["light", "dark", "no-preference"].includes(String(media.colorScheme))
    && ["reduce", "no-preference"].includes(String(media.reducedMotion))
    && ["active", "none"].includes(String(media.forcedColors)) && media.printMedia === true
    && ["ownedProcessRestricted", "contextRestricted"].includes(String(observed.networkIsolation))
    && (environment.attestationDigest === undefined || digest(environment.attestationDigest))
    && (attested === undefined || (reportKeys(attested, [
      "chromiumProduct", "chromiumBuild", "executableSha256", "launchFlags", "hostFontsDigest",
      "basis",
    ]) && typeof attested.chromiumProduct === "string" && attested.chromiumProduct.length > 0
      && typeof attested.chromiumBuild === "string" && attested.chromiumBuild.length > 0
      && (attested.executableSha256 === undefined || digest(attested.executableSha256))
      && reportStrings(attested.launchFlags)
      && attested.launchFlags.every((flag) => flag.length > 0)
      && new Set(attested.launchFlags).size === attested.launchFlags.length
      && digest(attested.hostFontsDigest) && typeof attested.basis === "string"
      && attested.basis.length > 0))
    && (environment.verification !== "callerAttested"
      || (attested !== undefined && digest(environment.attestationDigest)))
    && (environment.verification !== "nodeVerified"
      || (observed.runtimeKind === "nodeChromium" && observed.playwrightVersion !== undefined
        && observed.browserProduct !== undefined && observed.browserBuild !== undefined
        && digest(observed.executableSha256) && Array.isArray(observed.launchFlags)
        && observed.operatingSystem !== undefined && observed.architecture !== undefined))
    && fontsValid
    && fontReadinessValid
    && resourcesValid
    && unsupportedValid
    && warningsValid
    && pagesValid
    && bindingsValid;
}

/** Validate a browser failure artifact before retaining it on a Node error. */
export function isCurrentFailedRenderReport(value: unknown): value is FailedRenderReport {
  if (!hasCurrentRenderReportDiscriminator(value)) return false;
  const report = value as unknown as Record<string, unknown>;
  if (!reportKeys(report, [
    "schema", "schemaVersion", "source", "derivedProfileSource", "options", "readiness",
    "fonts", "fontReadiness", "resources", "unsupportedContent", "fontIdentity", "warnings",
    "status", "failure", "environment", "partial", "unavailable",
  ]) || report.status !== "failed") return false;

  const zeroDigest = "0".repeat(64);
  const options = reportRecord(report.options);
  const outputs = Array.isArray(options?.outputs) ? options.outputs : [];
  const { failure: _failure, unavailable: _unavailable, partial: _partial,
    status: _status, environment: _environment, ...base } = report;
  // Reuse the complete-report guard for the entire shared closed vocabulary.
  // Only branch-specific readiness, environment, partial, failure, and
  // unavailable evidence are replaced here and validated independently below.
  const commonCandidate = {
    ...base,
    fontIdentity: report.fontIdentity ?? {
      resolverContract: "https://docxodus.dev/contracts/font-resolver/v1",
      substitutionContractVersion: 1,
      substitutionContractDigest: zeroDigest,
      resolutionDigest: zeroDigest,
    },
    readiness: [{
      phase: "pagination",
      status: "complete",
      elapsedMs: 0,
      pending: [],
      diagnostics: [
        { code: "sections_processed", severity: "info", message: "validated", count: 0 },
        { code: "page_runs_processed", severity: "info", message: "validated", count: 0 },
        { code: "source_anchors_inventoried", severity: "info", message: "validated", count: 0 },
        { code: "note_references_inventoried", severity: "info", message: "validated", count: 0 },
      ],
    }],
    status: "complete",
    environment: {
      rendererFingerprint: zeroDigest,
      verification: "browserObserved",
      fidelityTier: "unbaselined",
      observed: {
        runtimeKind: "browser",
        locale: "und",
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
    pages: [{
      pageNumber: 1,
      pageInSection: 1,
      pageName: "validated",
      width: 1,
      height: 1,
      sectionIndex: 0,
    }],
    bindings: {
      pageMapDigest: zeroDigest,
      artifactRequestIds: [],
      ...(outputs.includes("html") ? { htmlDigest: zeroDigest } : {}),
      ...(outputs.includes("pdf") ? {
        pdfDigest: zeroDigest,
        pdfByteDeterministic: false,
        volatilePdfMetadata: {},
      } : {}),
    },
  };
  if (!isCurrentCompleteRenderReport(commonCandidate)) return false;

  const finite = (candidate: unknown, minimum = 0): candidate is number =>
    typeof candidate === "number" && Number.isFinite(candidate) && candidate >= minimum;
  const integer = (candidate: unknown, minimum = 0): candidate is number =>
    Number.isSafeInteger(candidate) && (candidate as number) >= minimum;
  const digest = (candidate: unknown): candidate is string =>
    typeof candidate === "string" && /^[0-9a-f]{64}$/.test(candidate);
  const diagnosticCodes = new Set([
    "sections_processed", "page_runs_processed", "source_anchors_inventoried",
    "note_references_inventoried",
  ]);
  if (!Array.isArray(report.readiness)) return false;
  let terminalCount = 0;
  for (const entry of report.readiness) {
    const item = reportRecord(entry);
    if (!item || !reportKeys(item, ["phase", "status", "elapsedMs", "pending", "diagnostics"])
      || !PHASES.has(item.phase as ExportPhase)
      || !["complete", "failed", "cancelled"].includes(String(item.status))
      || !finite(item.elapsedMs) || !reportStrings(item.pending)) return false;
    if (item.status === "failed" || item.status === "cancelled") terminalCount++;
    if (item.phase === "pagination" && item.status === "complete") {
      if (!Array.isArray(item.diagnostics) || item.diagnostics.length !== 4) return false;
      const seen = new Set<string>();
      for (const diagnostic of item.diagnostics) {
        const row = reportRecord(diagnostic);
        if (!row || !reportKeys(row, ["code", "severity", "message", "count"])
          || !diagnosticCodes.has(String(row.code)) || seen.has(String(row.code))
          || row.severity !== "info" || typeof row.message !== "string" || row.message.length === 0
          || !integer(row.count)) return false;
        seen.add(String(row.code));
      }
    } else if (item.diagnostics !== undefined) return false;
  }
  if (terminalCount !== 1) return false;

  const failure = reportRecord(report.failure);
  if (!failure || !reportKeys(failure, [
    "code", "severity", "phase", "message", "remediation", "detail", "pending",
    "partUri", "anchorId", "resource",
  ]) || !ERROR_CODES.has(failure.code as DocxodusExportErrorCode)
    || failure.severity !== "error" || !PHASES.has(failure.phase as ExportPhase)
    || typeof failure.message !== "string" || failure.message.length === 0
    || typeof failure.remediation !== "string" || failure.remediation.length === 0
    || ["detail", "partUri", "anchorId", "resource"].some((key) =>
      failure[key] !== undefined && typeof failure[key] !== "string")
    || (failure.pending !== undefined && !reportStrings(failure.pending))) return false;

  const environment = report.environment === undefined
    ? undefined : reportRecord(report.environment);
  if (report.environment !== undefined && (!environment || !reportKeys(environment, [
    "rendererFingerprint", "verification", "fidelityTier", "observed", "attested",
    "attestationDigest",
  ]) || !["nodeVerified", "browserObserved", "callerAttested"].includes(String(environment.verification))
    || (environment.rendererFingerprint !== undefined && !digest(environment.rendererFingerprint))
    || (environment.fidelityTier !== undefined
      && !["releaseBaselined", "experimental", "unbaselined"].includes(String(environment.fidelityTier)))
    || (environment.attestationDigest !== undefined && !digest(environment.attestationDigest)))) return false;
  const observed = environment?.observed === undefined
    ? undefined : reportRecord(environment.observed);
  const media = observed?.media === undefined ? undefined : reportRecord(observed.media);
  if (environment?.observed !== undefined && (!observed || !reportKeys(observed, [
    "runtimeKind", "playwrightVersion", "browserProduct", "browserBuild", "executableSha256",
    "launchFlags", "operatingSystem", "architecture", "locale", "timezone", "viewport",
    "deviceScaleFactor", "media", "networkIsolation",
  ]) || !["browser", "nodeChromium"].includes(String(observed.runtimeKind))
    || ["playwrightVersion", "browserProduct", "browserBuild", "operatingSystem", "architecture"]
      .some((key) => observed[key] !== undefined
        && (typeof observed[key] !== "string" || (observed[key] as string).length === 0))
    || (observed.executableSha256 !== undefined && !digest(observed.executableSha256))
    || (observed.launchFlags !== undefined && !reportStrings(observed.launchFlags))
    || typeof observed.locale !== "string" || observed.locale.length === 0
    || typeof observed.timezone !== "string" || observed.timezone.length === 0
    || !Array.isArray(observed.viewport) || observed.viewport.length !== 2
    || !observed.viewport.every((entry) => finite(entry, Number.MIN_VALUE))
    || !finite(observed.deviceScaleFactor, Number.MIN_VALUE)
    || !media || !reportKeys(media, ["colorScheme", "reducedMotion", "forcedColors", "printMedia"])
    || !["light", "dark", "no-preference"].includes(String(media.colorScheme))
    || !["reduce", "no-preference"].includes(String(media.reducedMotion))
    || !["active", "none"].includes(String(media.forcedColors)) || media.printMedia !== true
    || !["ownedProcessRestricted", "contextRestricted"].includes(String(observed.networkIsolation)))) {
    return false;
  }
  const attested = environment?.attested === undefined
    ? undefined : reportRecord(environment.attested);
  if (environment?.attested !== undefined && (!attested || !reportKeys(attested, [
    "chromiumProduct", "chromiumBuild", "executableSha256", "launchFlags", "hostFontsDigest",
    "basis",
  ]) || typeof attested.chromiumProduct !== "string" || attested.chromiumProduct.length === 0
    || typeof attested.chromiumBuild !== "string" || attested.chromiumBuild.length === 0
    || (attested.executableSha256 !== undefined && !digest(attested.executableSha256))
    || !reportStrings(attested.launchFlags) || !digest(attested.hostFontsDigest)
    || typeof attested.basis !== "string" || attested.basis.length === 0)) return false;
  if (environment?.verification === "callerAttested"
    && (!attested || !digest(environment.attestationDigest))) return false;

  const partial = report.partial === undefined ? undefined : reportRecord(report.partial);
  if (report.partial !== undefined && (!partial || !reportKeys(partial, ["pages", "bindings"]))) {
    return false;
  }
  if (partial?.pages !== undefined && (!Array.isArray(partial.pages)
    || partial.pages.some((entry) => {
      const page = reportRecord(entry);
      return !page || !reportKeys(page, [
        "pageNumber", "pageInSection", "pageName", "width", "height", "sectionIndex",
      ]) || !integer(page.pageNumber, 1) || !integer(page.pageInSection, 1)
        || typeof page.pageName !== "string" || page.pageName.length === 0
        || !finite(page.width, Number.MIN_VALUE) || !finite(page.height, Number.MIN_VALUE)
        || (page.sectionIndex !== undefined && !integer(page.sectionIndex));
    }))) return false;
  const bindings = partial?.bindings === undefined ? undefined : reportRecord(partial.bindings);
  if (partial?.bindings !== undefined && (!bindings || !reportKeys(bindings, [
    "pageMapDigest", "htmlDigest", "pdfDigest", "artifactRequestIds", "pdfByteDeterministic",
    "volatilePdfMetadata",
  ]) || ["pageMapDigest", "htmlDigest", "pdfDigest"].some((key) =>
    bindings[key] !== undefined && !digest(bindings[key]))
    || (bindings.artifactRequestIds !== undefined && !reportStrings(bindings.artifactRequestIds))
    || (bindings.pdfByteDeterministic !== undefined && bindings.pdfByteDeterministic !== false)
    || (bindings.volatilePdfMetadata !== undefined && (() => {
      const metadata = reportRecord(bindings.volatilePdfMetadata);
      return !metadata || Object.values(metadata).some((entry) => typeof entry !== "string");
    })()))) return false;

  if (!Array.isArray(report.unavailable) || report.unavailable.length > 4) return false;
  const unavailableFields = new Set<string>();
  for (const entry of report.unavailable) {
    const item = reportRecord(entry);
    if (!item || !reportKeys(item, ["field", "reasonCode", "detail"])
      || !["environment.rendererFingerprint", "bindings.pageMapDigest", "bindings.htmlDigest",
        "bindings.pdfDigest"].includes(String(item.field))
      || unavailableFields.has(String(item.field))
      || !["notReached", "notRequested", "failedVerification", "discardedOnFailure"]
        .includes(String(item.reasonCode))
      || typeof item.detail !== "string" || item.detail.length === 0) return false;
    unavailableFields.add(String(item.field));
  }
  return true;
}

export type {
  CommentProfile,
  CompleteRenderReport,
  DocxodusExportErrorCode,
  EnvironmentVerification,
  ExportPhase,
  ExportResourceLimits,
  FailedRenderReport,
  PaginatedHtmlOptions,
  PaginatedHtmlResult,
  PaginatedRenderMetadata,
  RenderReport,
  RenderWarning,
  ReviewProfile,
} from "docxodus/export-browser";

export interface FontLicenseAttestation {
  schemaVersion: 1;
  usage: "standalone-document-font-embedding";
  fileSha256: string;
  embeddingPermitted: true;
  permittedOutputs: readonly RenderOutput[];
  subsettingPermitted: boolean;
  basis: string;
  attester?: string;
}

export interface RenderEnvironmentAttestation {
  schemaVersion: 1;
  usage: "docxodus-render-environment";
  chromiumProduct: string;
  chromiumBuild: string;
  executableSha256?: string;
  launchFlags: readonly string[];
  hostFonts: ReadonlyArray<{
    family: string;
    postscriptName: string;
    style: FontFaceStyle;
    weight: number;
    stretch: number;
    fileSha256: string;
    version: string;
  }>;
  basis: string;
}

export interface NodeExportRuntime {
  /** Caller-owned browser. Docxodus creates a fresh context but never closes the browser. */
  browser?: Browser;
  /** Explicit Chromium executable used when `browser` is omitted. */
  browserExecutablePath?: string;
  /** Reserved for the verified font runtime delivered by issue #442. */
  fontDirectories?: readonly string[];
  fontLicenseAttestations?: readonly FontLicenseAttestation[];
  environmentAttestation?: RenderEnvironmentAttestation;
}

/** Validated internal runtime. Resolver functions never enter the public Node option surface. */
export interface ValidatedNodeExportRuntime extends NodeExportRuntime {
  fontDirectories: readonly string[];
  fontLicenseAttestations: readonly FontLicenseAttestation[];
  fontResolver?: FontResolver;
  /** Internal eager catalog validation used before Chromium launch. */
  prepareFonts?: (signal?: AbortSignal) => Promise<void>;
}

export type NodeExportOptions = Omit<PaginatedHtmlOptions, "wasmBasePath" | "fontResolver">
  & NodeExportRuntime;
export type RenderOutput = "html" | "pdf";

export interface PdfExportResult extends PaginatedRenderMetadata {
  pdf: Uint8Array;
}

export interface RenderBatchResult extends PaginatedRenderMetadata {
  html?: string;
  pdf?: Uint8Array;
}

export interface RenderBatchOptions extends NodeExportOptions {
  outputs: readonly RenderOutput[];
}

export interface RenderFileDestinations {
  htmlPath?: string;
  pdfPath?: string;
  pageMapPath?: string;
  reportPath?: string;
}

export interface RenderFileResult extends PaginatedRenderMetadata {
  written: RenderFileDestinations;
}

export class DocxodusExportError extends Error {
  readonly code: DocxodusExportErrorCode;
  readonly phase: ExportPhase;
  readonly remediation: string;
  readonly detail?: string;
  readonly pending?: readonly string[];
  readonly partUri?: string;
  readonly anchorId?: string;
  readonly resource?: string;
  readonly cause?: unknown;
  readonly report?: FailedRenderReport;
  readonly committedDestinations: readonly string[];

  constructor(
    code: DocxodusExportErrorCode,
    phase: ExportPhase,
    message: string,
    remediation: string,
    options: {
      detail?: string;
      pending?: readonly string[];
      partUri?: string;
      anchorId?: string;
      resource?: string;
      cause?: unknown;
      report?: FailedRenderReport;
      committedDestinations?: readonly string[];
    } = {},
  ) {
    super(message);
    this.name = "DocxodusExportError";
    this.code = code;
    this.phase = phase;
    this.remediation = remediation;
    this.detail = options.detail;
    this.pending = options.pending === undefined ? undefined : Object.freeze([...options.pending]);
    this.partUri = options.partUri;
    this.anchorId = options.anchorId;
    this.resource = options.resource;
    this.cause = options.cause;
    this.report = options.report;
    this.committedDestinations = Object.freeze([...(options.committedDestinations ?? [])]);
  }

  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      code: this.code,
      severity: "error",
      phase: this.phase,
      message: this.message,
      remediation: this.remediation,
      ...(this.detail === undefined ? {} : { detail: this.detail }),
      ...(this.pending === undefined ? {} : { pending: this.pending }),
      ...(this.partUri === undefined ? {} : { partUri: this.partUri }),
      ...(this.anchorId === undefined ? {} : { anchorId: this.anchorId }),
      ...(this.resource === undefined ? {} : { resource: this.resource }),
      ...(this.report === undefined ? {} : { report: this.report }),
      ...(this.committedDestinations.length === 0
        ? {}
        : { committedDestinations: this.committedDestinations }),
    };
  }
}

export interface BrowserMaterializationSuccess {
  html?: string;
  pageCount: number;
  pageMap: PaginatedHtmlResult["pageMap"];
  renderReport: CompleteRenderReport;
  warnings: RenderWarning[];
  rendererFingerprint: string;
}

export interface BrowserMaterializationFailure {
  name?: string;
  code?: string;
  phase?: string;
  message?: string;
  remediation?: string;
  detail?: string;
  pending?: string[];
  partUri?: string;
  anchorId?: string;
  resource?: string;
  report?: FailedRenderReport;
}

export function exportError(
  code: DocxodusExportErrorCode,
  phase: ExportPhase,
  message: string,
  remediation: string,
  options: ConstructorParameters<typeof DocxodusExportError>[4] = {},
): never {
  throw new DocxodusExportError(code, phase, message, remediation, options);
}

const ERROR_CODES = new Set<DocxodusExportErrorCode>([
  "invalid_argument",
  "invalid_document",
  "source_digest_mismatch",
  "document_version_unrepresentable",
  "conversion_failure",
  "browser_launch_failure",
  "resource_policy_failure",
  "readiness_timeout",
  "operation_cancelled",
  "pagination_failure",
  "pdf_write_failure",
  "output_write_failure",
  "output_verification_failure",
  "resource_limit",
  "unsupported_runtime",
  "filesystem_failure",
]);

const PHASES = new Set<ExportPhase>([
  "input_validation",
  "package_preflight",
  "browser_launch",
  "wasm_initialization",
  "docx_conversion",
  "font_loading",
  "image_decoding",
  "chart_svg_materialization",
  "pagination",
  "running_story_placement",
  "page_tree_stability",
  "pdf_print",
  "output_verification",
  "output_write",
  "filesystem_commit",
  "cleanup",
]);

export function fromBrowserFailure(value: BrowserMaterializationFailure): DocxodusExportError {
  const report = isCurrentFailedRenderReport(value.report)
    ? structuredClone(value.report)
    : undefined;
  const code = ERROR_CODES.has(value.code as DocxodusExportErrorCode)
    ? value.code as DocxodusExportErrorCode
    : "conversion_failure";
  const phase = PHASES.has(value.phase as ExportPhase)
    ? value.phase as ExportPhase
    : "docx_conversion";
  return new DocxodusExportError(
    code,
    phase,
    value.message ?? "The browser materializer failed without a message.",
    value.remediation ?? "Inspect the render report and source document.",
    {
      detail: value.detail,
      pending: value.pending ?? report?.failure.pending,
      partUri: value.partUri ?? report?.failure.partUri,
      anchorId: value.anchorId ?? report?.failure.anchorId,
      resource: value.resource ?? report?.failure.resource,
      report,
    },
  );
}

export function failedReportFromComplete(
  report: CompleteRenderReport,
  error: DocxodusExportError,
): FailedRenderReport {
  const { status: _status, environment, pages, bindings, readiness, ...base } = report;
  const unavailable: FailedRenderReport["unavailable"] = [];
  if (report.options.outputs.includes("html")) {
    if (!bindings.htmlDigest) {
      unavailable.push({
        field: "bindings.htmlDigest",
        reasonCode: "failedVerification",
        detail: "HTML output did not verify before the operation failed.",
      });
    }
  } else {
    unavailable.push({
      field: "bindings.htmlDigest",
      reasonCode: "notRequested",
      detail: "HTML output was not selected for this operation.",
    });
  }
  if (report.options.outputs.includes("pdf")) {
    if (!bindings.pdfDigest) {
      unavailable.push({
        field: "bindings.pdfDigest",
        reasonCode: "failedVerification",
        detail: "PDF output did not verify before the operation failed.",
      });
    }
  } else {
    unavailable.push({
      field: "bindings.pdfDigest",
      reasonCode: "notRequested",
      detail: "PDF output was not selected for this operation.",
    });
  }
  return {
    ...base,
    readiness: readiness.some((entry) => entry.status !== "complete")
      ? readiness
      : [
          ...readiness,
          {
            phase: error.phase,
            status: error.code === "operation_cancelled" ? "cancelled" : "failed",
            elapsedMs: 0,
            pending: [...(error.pending ?? [])],
          },
        ],
    status: "failed",
    failure: {
      code: error.code,
      severity: "error",
      phase: error.phase,
      message: error.message,
      remediation: error.remediation,
      ...(error.detail === undefined ? {} : { detail: error.detail }),
      ...(error.pending === undefined ? {} : { pending: [...error.pending] }),
      ...(error.partUri === undefined ? {} : { partUri: error.partUri }),
      ...(error.anchorId === undefined ? {} : { anchorId: error.anchorId }),
      ...(error.resource === undefined ? {} : { resource: error.resource }),
    },
    environment,
    partial: { pages, bindings },
    unavailable,
  };
}

export function attachFailedReport(
  error: unknown,
  report: CompleteRenderReport | undefined,
): unknown {
  if (!(error instanceof DocxodusExportError) || error.report || !report) return error;
  return new DocxodusExportError(error.code, error.phase, error.message, error.remediation, {
    detail: error.detail,
    pending: error.pending,
    partUri: error.partUri,
    anchorId: error.anchorId,
    resource: error.resource,
    cause: error.cause,
    report: failedReportFromComplete(report, error),
    committedDestinations: error.committedDestinations,
  });
}

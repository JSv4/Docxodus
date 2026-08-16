import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_TIMEOUT_MS,
  COMMENT_PROFILES,
  REVIEW_PROFILES,
  normalizeFontFamilyName,
  type CompleteRenderReport,
  type ExportResourceLimits,
  type FontFaceStyle,
  type FontResolution,
  type PaginatedHtmlOptions,
  type PaginatedHtmlResult,
} from "docxodus/export-browser";
import { resolve } from "node:path";
import { canonicalJson, canonicalJsonBytes, sha256 } from "./canonical.js";
import { renderInBrowser } from "./browser-session.js";
import type {
  FontLicenseAttestation,
  NodeExportOptions,
  NodeExportRuntime,
  PdfExportResult,
  RenderBatchOptions,
  RenderBatchResult,
  RenderEnvironmentAttestation,
  RenderFileDestinations,
  RenderFileResult,
  RenderOutput,
  ValidatedNodeExportRuntime,
} from "./contracts.js";
import {
  attachFailedReport,
  DocxodusExportError,
  exportError,
} from "./contracts.js";
import {
  prepareDestinations,
  readStableInputFile,
  writeNoReplace,
} from "./files.js";
import { verifyPdf } from "./pdf.js";
import { createNodeFontRuntime } from "./fonts/index.js";

export * from "./contracts.js";
export {
  COMMENT_PROFILES,
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_RESOURCE_LIMITS,
  HARD_EXPORT_TIMEOUT_MS,
  REVIEW_PROFILES,
} from "docxodus/export-browser";

function ownedInput(document: Uint8Array): Uint8Array {
  if (!(document instanceof Uint8Array)) {
    exportError(
      "invalid_document",
      "input_validation",
      "The Node export API requires a Uint8Array DOCX snapshot.",
      "Read the file into a Uint8Array or use renderDocxFile().",
    );
  }
  return new Uint8Array(document);
}

function exactKeys(record: Record<string, unknown>, allowed: readonly string[], label: string): void {
  const extra = Object.keys(record).filter((key) => !allowed.includes(key));
  if (extra.length > 0) {
    exportError(
      "invalid_document",
      "input_validation",
      `${label} contains unknown fields: ${extra.join(", ")}.`,
      "Use the documented attestation schema without extension fields.",
    );
  }
}

function nonEmptyString(value: unknown, label: string): string {
  if (typeof value !== "string" || value.trim() === "") {
    exportError(
      "invalid_document",
      "input_validation",
      `${label} must be a non-empty string.`,
      "Correct the runtime attestation and retry.",
    );
  }
  if (value.length > 1024 || /[\u0000-\u001f\u007f]/u.test(value)) {
    exportError("invalid_document", "input_validation", `${label} is not a bounded plain string.`,
      "Use at most 1024 printable characters.");
  }
  return value;
}

function digestString(value: unknown, label: string): string {
  const digest = nonEmptyString(value, label);
  if (!/^[0-9a-f]{64}$/.test(digest)) {
    exportError(
      "invalid_document",
      "input_validation",
      `${label} must be a lower-case SHA-256 digest.`,
      "Provide exactly 64 hexadecimal characters.",
    );
  }
  return digest;
}

function validateEnvironmentAttestation(
  value: RenderEnvironmentAttestation | undefined,
  limits: Readonly<ExportResourceLimits>,
): RenderEnvironmentAttestation | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    exportError("invalid_document", "input_validation", "environmentAttestation must be an object.",
      "Provide the documented canonical JSON object.");
  }
  const record = value as unknown as Record<string, unknown>;
  exactKeys(record, [
    "chromiumProduct",
    "chromiumBuild",
    "executableSha256",
    "launchFlags",
    "hostFonts",
    "basis",
  ], "environmentAttestation");
  const launchFlags = record.launchFlags;
  const hostFonts = record.hostFonts;
  if (!Array.isArray(launchFlags)) {
    exportError("invalid_document", "input_validation", "launchFlags must be a string array.",
      "Record every attested Chromium launch flag.");
  }
  if (launchFlags.length > 128) {
    exportError("resource_limit", "input_validation", "launchFlags exceeds 128 entries.",
      "Use the bounded canonical Chromium launch configuration.");
  }
  if (!Array.isArray(hostFonts)) {
    exportError("invalid_document", "input_validation", "hostFonts must be an array.",
      "Provide an empty array when no host fonts are attested.");
  }
  if (hostFonts.length > limits.fontFiles) {
    exportError("resource_limit", "input_validation",
      `fontFiles limit exceeded by hostFonts (${hostFonts.length} > ${limits.fontFiles}).`,
      "Attest no more host font faces than the configured fontFiles limit.");
  }
  const resolvedLaunchFlags = launchFlags.map((flag, index) =>
    nonEmptyString(flag, `launchFlags[${index}]`));
  const seen = new Set<string>();
  const seenFaces = new Set<string>();
  const resolvedFonts = hostFonts.map((font, index) => {
    if (!font || typeof font !== "object" || Array.isArray(font)) {
      exportError("invalid_document", "input_validation", `hostFonts[${index}] must be an object.`,
        "Correct the environment attestation.");
    }
    const item = font as Record<string, unknown>;
    exactKeys(item, ["family", "postscriptName", "style", "weight", "stretch", "fileSha256", "version"],
      `hostFonts[${index}]`);
    if (item.style !== "normal" && item.style !== "italic" && item.style !== "oblique") {
      exportError("invalid_document", "input_validation", `hostFonts[${index}].style is invalid.`,
        "Use normal, italic, or oblique.");
    }
    const style = item.style as FontFaceStyle;
    if (!Number.isSafeInteger(item.weight) || (item.weight as number) < 1
      || (item.weight as number) > 1000) {
      exportError("invalid_document", "input_validation", `hostFonts[${index}].weight is invalid.`,
        "Use an integer font weight from 1 through 1000.");
    }
    if (typeof item.stretch !== "number" || !Number.isFinite(item.stretch)
      || item.stretch < 50 || item.stretch > 200) {
      exportError("invalid_document", "input_validation", `hostFonts[${index}].stretch is invalid.`,
        "Use a CSS stretch percentage from 50 through 200.");
    }
    const family = normalizeFontFamilyName(nonEmptyString(item.family, `hostFonts[${index}].family`));
    const postscriptName = normalizeFontFamilyName(
      nonEmptyString(item.postscriptName, `hostFonts[${index}].postscriptName`),
    );
    const faceKey = canonicalJson([
      family.toLowerCase(),
      postscriptName.toLowerCase(),
      style,
      item.weight,
      item.stretch,
    ]);
    if (seenFaces.has(faceKey)) {
      exportError("invalid_document", "input_validation", "hostFonts contains a duplicate face identity.",
        "List each family/PostScript-name/style/weight/stretch face exactly once.");
    }
    seenFaces.add(faceKey);
    const fileSha256 = digestString(item.fileSha256, `hostFonts[${index}].fileSha256`);
    if (seen.has(fileSha256)) {
      exportError("invalid_document", "input_validation", "hostFonts contains a duplicate file digest.",
        "List each attested font file exactly once.");
    }
    seen.add(fileSha256);
    return {
      family,
      postscriptName,
      style,
      weight: item.weight as number,
      stretch: item.stretch,
      fileSha256,
      version: nonEmptyString(item.version, `hostFonts[${index}].version`),
    };
  });
  return Object.freeze({
    chromiumProduct: nonEmptyString(record.chromiumProduct, "chromiumProduct"),
    chromiumBuild: nonEmptyString(record.chromiumBuild, "chromiumBuild"),
    ...(record.executableSha256 === undefined
      ? {}
      : { executableSha256: digestString(record.executableSha256, "executableSha256") }),
    launchFlags: Object.freeze(resolvedLaunchFlags),
    hostFonts: Object.freeze(resolvedFonts),
    basis: nonEmptyString(record.basis, "basis"),
  });
}

function validateRuntime(
  runtime: NodeExportRuntime,
  limits: Readonly<ExportResourceLimits>,
): ValidatedNodeExportRuntime {
  if (runtime.browser && runtime.browserExecutablePath) {
    exportError(
      "invalid_document",
      "input_validation",
      "browser and browserExecutablePath cannot be supplied together.",
      "Inject one caller-owned Chromium browser or provide one executable path.",
    );
  }
  if (runtime.browserExecutablePath !== undefined
    && (typeof runtime.browserExecutablePath !== "string"
      || runtime.browserExecutablePath.trim() === "")) {
    exportError(
      "invalid_document",
      "input_validation",
      "browserExecutablePath must be a non-empty path.",
      "Provide a Chromium executable path or omit the option to use the pinned browser.",
    );
  }
  const fontDirectories = runtime.fontDirectories ?? [];
  if (!Array.isArray(fontDirectories)) {
    exportError("invalid_document", "input_validation", "fontDirectories must be a string array.",
      "Provide each explicit font directory as a non-empty path.");
  }
  if (fontDirectories.length > limits.fontDirectoryEntries) {
    exportError("resource_limit", "input_validation",
      `fontDirectoryEntries limit exceeded by fontDirectories (${fontDirectories.length} > ${limits.fontDirectoryEntries}).`,
      "Configure fewer explicit font directory roots.");
  }
  if (!fontDirectories.every((entry) => typeof entry === "string" && entry.trim() !== "")) {
    exportError("invalid_document", "input_validation", "fontDirectories must be a string array.",
      "Provide each explicit font directory as a non-empty path.");
  }
  const attestations = runtime.fontLicenseAttestations ?? [];
  if (!Array.isArray(attestations)) {
    exportError("invalid_document", "input_validation", "fontLicenseAttestations must be an array.",
      "Provide the documented attestation objects.");
  }
  if (attestations.length > limits.fontFiles) {
    exportError("resource_limit", "input_validation",
      `fontFiles limit exceeded by fontLicenseAttestations (${attestations.length} > ${limits.fontFiles}).`,
      "Attest no more font identities than the configured fontFiles limit.");
  }
  const normalizedAttestations: FontLicenseAttestation[] = [];
  const attestationDigests = new Set<string>();
  for (const [index, attestation] of attestations.entries()) {
    if (!attestation || typeof attestation !== "object" || Array.isArray(attestation)) {
      exportError(
        "invalid_document",
        "input_validation",
        `fontLicenseAttestations[${index}] must be an object.`,
        "Provide the documented font-license attestation schema.",
      );
    }
    const record = attestation as unknown as Record<string, unknown>;
    exactKeys(
      record,
      ["schemaVersion", "usage", "fileSha256", "embeddingPermitted", "basis", "attester"],
      `fontLicenseAttestations[${index}]`,
    );
    if (record.schemaVersion !== 1
      || record.usage !== "standalone-document-font-embedding") {
      exportError(
        "invalid_document",
        "input_validation",
        `fontLicenseAttestations[${index}] has an unsupported schema or usage.`,
        "Use schemaVersion 1 for standalone-document-font-embedding.",
      );
    }
    const fileSha256 = digestString(record.fileSha256,
      `fontLicenseAttestations[${index}].fileSha256`);
    if (attestationDigests.has(fileSha256)) {
      exportError("invalid_document", "input_validation",
        "fontLicenseAttestations contains a duplicate file digest.",
        "List each attested font byte identity once.");
    }
    attestationDigests.add(fileSha256);
    if (record.embeddingPermitted !== true) {
      exportError(
        "invalid_document",
        "input_validation",
        `fontLicenseAttestations[${index}].embeddingPermitted must be true.`,
        "Do not load a font unless embedding permission is affirmatively attested.",
      );
    }
    const basis = nonEmptyString(record.basis, `fontLicenseAttestations[${index}].basis`);
    let attester: string | undefined;
    if (record.attester !== undefined) {
      attester = nonEmptyString(record.attester, `fontLicenseAttestations[${index}].attester`);
    }
    normalizedAttestations.push(Object.freeze({
      schemaVersion: 1 as const,
      usage: "standalone-document-font-embedding" as const,
      fileSha256,
      embeddingPermitted: true as const,
      basis,
      ...(attester ? { attester } : {}),
    }));
  }
  if (attestations.length > 0 && fontDirectories.length === 0) {
    exportError(
      "invalid_document",
      "input_validation",
      "Font-license attestations require at least one font directory.",
      "Remove the unattached attestations or provide the matching font directory.",
    );
  }
  // Capture relative roots against the call-time working directory before the
  // first asynchronous browser operation can observe a changed process cwd.
  const normalizedDirectories = Object.freeze(fontDirectories.map((directory) => resolve(directory)));
  const frozenAttestations = Object.freeze(normalizedAttestations);
  const fontRuntime = normalizedDirectories.length === 0
    ? undefined
    : createNodeFontRuntime(normalizedDirectories, frozenAttestations, limits);
  return {
    browser: runtime.browser,
    browserExecutablePath: runtime.browserExecutablePath,
    fontDirectories: normalizedDirectories,
    fontLicenseAttestations: frozenAttestations,
    ...(fontRuntime ? {
      fontResolver: fontRuntime.resolver,
      prepareFonts: fontRuntime.prepare,
    } : {}),
    environmentAttestation: validateEnvironmentAttestation(runtime.environmentAttestation, limits),
  };
}

function nodeOptionsPreflight(options: NodeExportOptions): void {
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    exportError(
      "invalid_document",
      "input_validation",
      "Export options are required.",
      "Supply explicit reviewProfile and commentProfile values.",
    );
  }
  if (!(REVIEW_PROFILES as readonly string[]).includes(options.reviewProfile)) {
    exportError("invalid_document", "input_validation", "reviewProfile is invalid.",
      "Use final, original, or markup.");
  }
  if (options.reviewProfileAlreadyApplied !== undefined
    && typeof options.reviewProfileAlreadyApplied !== "boolean") {
    exportError("invalid_document", "input_validation",
      "reviewProfileAlreadyApplied must be a boolean.",
      "Use true only for an exact final/original source that contains no tracked revisions.");
  }
  if (options.reviewProfileAlreadyApplied === true && options.reviewProfile === "markup") {
    exportError("invalid_document", "input_validation",
      "reviewProfileAlreadyApplied cannot be used with the markup profile.",
      "Omit reviewProfileAlreadyApplied for markup, which renders the unchanged revision-bearing source.");
  }
  if (!(COMMENT_PROFILES as readonly string[]).includes(options.commentProfile)) {
    exportError("invalid_document", "input_validation", "commentProfile is invalid.",
      "Use hidden, inline, endnotes, or margin.");
  }
  if (options.documentVersion !== undefined
    && (!Number.isSafeInteger(options.documentVersion) || options.documentVersion < 0)) {
    exportError(
      "invalid_document",
      "input_validation",
      "documentVersion must be a non-negative JavaScript safe integer.",
      "Use a value between 0 and Number.MAX_SAFE_INTEGER.",
    );
  }
  if (options.expectedSourceDigest !== undefined
    && !/^[0-9a-f]{64}$/i.test(options.expectedSourceDigest)) {
    exportError(
      "invalid_document",
      "input_validation",
      "expectedSourceDigest must be a SHA-256 hex digest.",
      "Supply exactly 64 hexadecimal characters.",
    );
  }
  if (options.unsupportedContent !== undefined
    && options.unsupportedContent !== "warn"
    && options.unsupportedContent !== "strict") {
    exportError("invalid_document", "input_validation", "unsupportedContent is invalid.",
      "Use warn or strict.");
  }
  if (options.strictFonts !== undefined && typeof options.strictFonts !== "boolean") {
    exportError("invalid_document", "input_validation", "strictFonts must be a boolean.",
      "Use true or false.");
  }
  if (options.title !== undefined && typeof options.title !== "string") {
    exportError("invalid_document", "input_validation", "title must be a string.",
      "Provide plain document-title text or omit the option.");
  }

  const suppliedLimits = options.limits ?? {};
  if (!suppliedLimits || typeof suppliedLimits !== "object" || Array.isArray(suppliedLimits)) {
    exportError("invalid_document", "input_validation", "limits must be an object.",
      "Use keys from ExportResourceLimits with positive integer values.");
  }
  for (const [name, value] of Object.entries(suppliedLimits)) {
    if (!(name in DEFAULT_EXPORT_RESOURCE_LIMITS)) {
      exportError("invalid_document", "input_validation", `Unknown export limit: ${name}.`,
        "Use a key from ExportResourceLimits.");
    }
    const key = name as keyof ExportResourceLimits;
    if (!Number.isSafeInteger(value) || (value as number) <= 0) {
      exportError(
        "invalid_document",
        "input_validation",
        `Export limit ${name} must be a positive safe integer.`,
        "Supply a positive integer no greater than the published default.",
      );
    }
    if ((value as number) > DEFAULT_EXPORT_RESOURCE_LIMITS[key]) {
      exportError(
        "invalid_document",
        "input_validation",
        `Export limit ${name} may only lower the default.`,
        `Use ${DEFAULT_EXPORT_RESOURCE_LIMITS[key]} or less.`,
      );
    }
  }
}

function sourcePreflight(sourceBytes: Uint8Array, options: NodeExportOptions): void {
  const compressedLimit = options.limits?.compressedDocxBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.compressedDocxBytes;
  if (sourceBytes.byteLength > compressedLimit) {
    exportError(
      "resource_limit",
      "package_preflight",
      `compressedDocxBytes limit exceeded (${sourceBytes.byteLength} > ${compressedLimit}).`,
      "Use a smaller document or a deployment with a reviewed limits contract.",
    );
  }
  const actualDigest = options.expectedSourceDigest ? sha256(sourceBytes) : undefined;
  if (options.expectedSourceDigest
    && actualDigest !== options.expectedSourceDigest.toLowerCase()) {
    exportError(
      "invalid_document",
      "package_preflight",
      "The source digest does not match expectedSourceDigest.",
      "Render the exact verified source bytes or update the expected digest.",
      {
        detail: `expected=${options.expectedSourceDigest.toLowerCase()}; actual=${actualDigest}`,
      },
    );
  }
}

function snapshotNodeOptions<T extends NodeExportOptions>(options: T): T {
  if (!options || typeof options !== "object" || Array.isArray(options)) return options;
  const environment = options.environmentAttestation;
  const snapshot = {
    ...options,
    ...(options.limits && typeof options.limits === "object" && !Array.isArray(options.limits)
      ? { limits: { ...options.limits } }
      : {}),
    ...(Array.isArray(options.fontDirectories)
      ? { fontDirectories: [...options.fontDirectories] }
      : {}),
    ...(Array.isArray(options.fontLicenseAttestations)
      ? {
        fontLicenseAttestations: options.fontLicenseAttestations.map((entry) =>
          entry && typeof entry === "object" ? { ...entry } : entry),
      }
      : {}),
    ...(environment && typeof environment === "object" && !Array.isArray(environment)
      ? {
        environmentAttestation: {
          ...environment,
          ...(Array.isArray(environment.launchFlags)
            ? { launchFlags: [...environment.launchFlags] }
            : {}),
          ...(Array.isArray(environment.hostFonts)
            ? {
              hostFonts: environment.hostFonts.map((font) =>
                font && typeof font === "object" ? { ...font } : font),
            }
            : {}),
        },
      }
      : {}),
    ...("outputs" in options && Array.isArray((options as RenderBatchOptions).outputs)
      ? { outputs: [...(options as RenderBatchOptions).outputs] }
      : {}),
  };
  return snapshot as T;
}

function snapshotDestinations(destinations: RenderFileDestinations): RenderFileDestinations {
  if (!destinations || typeof destinations !== "object" || Array.isArray(destinations)) {
    exportError("invalid_document", "input_validation", "destinations must be an object.",
      "Provide at least one documented artifact path.");
  }
  const allowed = ["htmlPath", "pdfPath", "pageMapPath", "reportPath"];
  const unknown = Object.keys(destinations).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) {
    exportError(
      "invalid_document",
      "input_validation",
      `destinations contains unknown fields: ${unknown.join(", ")}.`,
      "Use only htmlPath, pdfPath, pageMapPath, and reportPath.",
    );
  }
  return { ...destinations };
}

function validateOutputs(outputs: readonly RenderOutput[]): RenderOutput[] {
  if (!Array.isArray(outputs)) {
    exportError("invalid_document", "input_validation", "outputs must be an array.",
      "Request html, pdf, or both.");
  }
  const resolved: RenderOutput[] = [];
  for (const output of outputs) {
    if (output !== "html" && output !== "pdf") {
      exportError("invalid_document", "input_validation", `Unknown output kind: ${String(output)}.`,
        "Use html or pdf.");
    }
    if (resolved.includes(output)) {
      exportError("invalid_document", "input_validation", `Duplicate output kind: ${output}.`,
        "List each requested output once.");
    }
    resolved.push(output);
  }
  return resolved;
}

function normalizedTimeout(options: NodeExportOptions): number {
  const timeout = options.timeoutMs ?? DEFAULT_EXPORT_TIMEOUT_MS;
  if (!Number.isSafeInteger(timeout) || timeout <= 0 || timeout > HARD_EXPORT_TIMEOUT_MS) {
    exportError(
      "invalid_document",
      "input_validation",
      `timeoutMs must be an integer from 1 through ${HARD_EXPORT_TIMEOUT_MS}.`,
      "Use the shared export timeout contract.",
    );
  }
  return timeout;
}

function browserOptions(
  options: NodeExportOptions,
): Omit<PaginatedHtmlOptions, "wasmBasePath" | "fontResolver"> {
  return {
    documentVersion: options.documentVersion,
    expectedSourceDigest: options.expectedSourceDigest,
    reviewProfile: options.reviewProfile,
    reviewProfileAlreadyApplied: options.reviewProfileAlreadyApplied,
    commentProfile: options.commentProfile,
    title: options.title,
    unsupportedContent: options.unsupportedContent,
    strictFonts: options.strictFonts,
    timeoutMs: options.timeoutMs,
    limits: options.limits,
  };
}

type BrowserRuntimeIdentity = Awaited<ReturnType<typeof renderInBrowser>>["runtime"];

function applyHostFontAttestation(
  report: CompleteRenderReport,
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation | undefined,
): void {
  if (!attestation || !runtimeAttestationMatches(runtime, attestation)) return;
  let changed = false;
  report.fonts = report.fonts.map((font): FontResolution => {
    if (font.source !== "browser" || font.status !== "unverified") return font;
    const family = font.resolvedFamily ?? font.requestedFamily;
    const match = attestation.hostFonts.find((candidate) =>
      normalizeFontFamilyName(candidate.family).toLowerCase()
        === normalizeFontFamilyName(family).toLowerCase()
      && candidate.style === font.requestedStyle
      && candidate.weight === font.requestedWeight
      && candidate.stretch === font.requestedStretch);
    if (!match) return font;
    changed = true;
    return {
      ...font,
      status: "resolved",
      source: "attested",
      resolvedFamily: match.family,
      resolvedFace: match.postscriptName,
      fileSha256: match.fileSha256,
      version: match.version,
      faceMatch: "exact",
      // document.fonts already proved the sampled request loadable. Exact host
      // attestation binds that observation to this immutable face identity.
      glyphCoverage: "complete",
    };
  });
  if (changed && report.fontIdentity) {
    const previousResolutionDigest = report.fontIdentity.resolutionDigest;
    report.fontIdentity = {
      ...report.fontIdentity,
      resolutionDigest: sha256(canonicalJson({
        schemaVersion: 1,
        previousResolutionDigest,
        resolutions: report.fonts,
      })),
    };
  }
}

function runtimeAttestationMatches(
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation,
): boolean {
  if (attestation.chromiumProduct !== runtime.chromiumProduct) return false;
  if (attestation.chromiumBuild !== runtime.browserVersion) return false;
  if (attestation.executableSha256 !== undefined
    && attestation.executableSha256 !== runtime.executableDigest) return false;
  if (runtime.launchMode !== "injected"
    && canonicalJson(attestation.launchFlags) !== canonicalJson(runtime.launchFlags)) return false;
  return true;
}

function strictFontResolution(
  font: FontResolution,
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation | undefined,
): boolean {
  const exact = font.status === "resolved"
    && font.faceMatch === "exact"
    && font.glyphCoverage === "complete"
    && typeof font.fileSha256 === "string"
    && typeof font.version === "string";
  if (!exact) return false;
  if ((font.source === "configured" || font.source === "attested")
    && font.licenseEvidence !== undefined) return true;
  return font.source === "attested"
    && font.licenseEvidence === undefined
    && attestation !== undefined
    && runtimeAttestationMatches(runtime, attestation)
    && attestation.hostFonts.some((face) =>
      face.fileSha256 === font.fileSha256
      && face.version === font.version
      && face.postscriptName === font.resolvedFace);
}

function reconcileHostFontWarnings(report: CompleteRenderReport): void {
  const stillUnverified = report.fonts.some((font) =>
    font.status === "unverified" || font.source === "browser");
  if (!stillUnverified) {
    report.warnings = report.warnings.filter(({ code }) => code !== "font_environment_unverified");
  }
}

function enforceStrictFontPolicy(
  report: CompleteRenderReport,
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation | undefined,
): void {
  const failures = report.fonts.filter((font) => !strictFontResolution(font, runtime, attestation));
  if (failures.length === 0) return;
  report.warnings = report.warnings.map((warning) => warning.phase === "font_loading"
    ? { ...warning, severity: "error" as const }
    : warning);
  exportError(
    "resource_policy_failure",
    "font_loading",
    "Strict font policy rejected a non-exact, unverified, or incompletely covered font outcome.",
    "Supply exact verified or host-attested faces with complete glyph coverage, or disable strictFonts.",
    {
      detail: failures.map(({ requestId, status }) => `${requestId}:${status}`).join(", "),
    },
  );
}

function finalRendererFingerprint(
  browserMaterializerFingerprint: string,
  runtime: BrowserRuntimeIdentity,
  report: CompleteRenderReport,
  verification: CompleteRenderReport["environment"]["verification"],
  attestation: RenderEnvironmentAttestation | undefined,
): string {
  const fonts = report.fonts.map((font) => ({ ...font }))
    .sort((left, right) => {
      const leftKey = canonicalJson(left);
      const rightKey = canonicalJson(right);
      return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
    });
  return sha256(canonicalJson({
    schemaVersion: 1,
    browserMaterializerFingerprint,
    runtime,
    environment: {
      verification,
      ...(verification === "callerAttested" && attestation
        ? {
          attestedRuntime: {
            chromiumProduct: attestation.chromiumProduct,
            chromiumBuild: attestation.chromiumBuild,
            launchFlags: attestation.launchFlags,
            ...(attestation.executableSha256
              ? { executableSha256: attestation.executableSha256 }
              : {}),
          },
        }
        : {}),
    },
    fontIdentity: report.fontIdentity,
    fonts,
  }));
}

function environmentVerification(
  report: CompleteRenderReport,
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation | undefined,
): "nodeVerified" | "browserObserved" | "callerAttested" {
  const configuredFont = (font: FontResolution): boolean => strictFontResolution(
    font,
    runtime,
    undefined,
  )
    && (font.source === "configured" || font.source === "attested")
    && font.licenseEvidence !== undefined;
  const hostAttestedFont = (font: FontResolution): boolean => strictFontResolution(
    font,
    runtime,
    attestation,
  ) && font.licenseEvidence === undefined;
  const everyFontVerified = report.fonts.every((font) =>
    configuredFont(font) || hostAttestedFont(font));
  if (!everyFontVerified) return "browserObserved";
  if (runtime.launchMode !== "injected") {
    return report.fonts.every(configuredFont)
      ? "nodeVerified"
      : attestation && runtimeAttestationMatches(runtime, attestation)
        ? "callerAttested"
        : "browserObserved";
  }
  return attestation && runtimeAttestationMatches(runtime, attestation)
    ? "callerAttested"
    : "browserObserved";
}

async function prepareFontsBeforeBrowser(
  runtime: ValidatedNodeExportRuntime,
  deadline: number,
): Promise<void> {
  if (!runtime.prepareFonts) return;
  const timeoutMs = deadline - Date.now();
  if (timeoutMs <= 0) {
    exportError(
      "readiness_timeout",
      "font_loading",
      "Export timed out while validating the configured font catalog.",
      "Increase timeoutMs or reduce the configured font catalog.",
    );
  }
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    await runtime.prepareFonts(controller.signal);
  } catch (error) {
    if (controller.signal.aborted) {
      exportError(
        "readiness_timeout",
        "font_loading",
        "Export timed out while validating the configured font catalog.",
        "Increase timeoutMs or reduce the configured font catalog.",
      );
    }
    throw error;
  } finally {
    clearTimeout(timer);
  }
}

async function renderOwned(
  sourceBytes: Uint8Array,
  options: NodeExportOptions,
  outputs: readonly RenderOutput[],
): Promise<RenderBatchResult> {
  const requested = validateOutputs(outputs);
  nodeOptionsPreflight(options);
  sourcePreflight(sourceBytes, options);
  const timeoutMs = normalizedTimeout(options);
  const effectiveLimits = Object.freeze({
    ...DEFAULT_EXPORT_RESOURCE_LIMITS,
    ...(options.limits ?? {}),
  });
  const runtime = validateRuntime(options, effectiveLimits);
  const deadline = Date.now() + timeoutMs;
  let report: CompleteRenderReport | undefined;
  try {
    await prepareFontsBeforeBrowser(runtime, deadline);
    const attestation = runtime.environmentAttestation;
    const deferStrictFontPolicy = options.strictFonts === true
      && (attestation?.hostFonts.length ?? 0) > 0;
    const browser = await renderInBrowser(
      sourceBytes,
      browserOptions({
        ...options,
        timeoutMs,
        ...(deferStrictFontPolicy ? { strictFonts: false } : {}),
      }),
      runtime,
      requested.includes("html"),
      requested.includes("pdf"),
      deadline,
    );
    report = structuredClone(browser.materialization.renderReport);
    applyHostFontAttestation(report, browser.runtime, attestation);
    reconcileHostFontWarnings(report);
    const verification = environmentVerification(report, browser.runtime, attestation);
    const rendererFingerprint = finalRendererFingerprint(
      browser.materialization.rendererFingerprint,
      browser.runtime,
      report,
      verification,
      attestation,
    );
    const pageMap = {
      ...browser.materialization.pageMap,
      rendererFingerprint,
    };
    report.environment = {
      rendererFingerprint,
      verification,
    };
    report.bindings.pageMapDigest = sha256(canonicalJson(pageMap));
    report.bindings.artifactRequestIds = [];
    if (options.strictFonts) enforceStrictFontPolicy(report, browser.runtime, attestation);

    if (browser.pdf) {
      const pdfLimit = options.limits?.pdfOutputBytes
        ?? DEFAULT_EXPORT_RESOURCE_LIMITS.pdfOutputBytes;
      if (browser.pdf.byteLength > pdfLimit) {
        exportError(
          "resource_limit",
          "output_verification",
          `PDF output exceeds pdfOutputBytes (${browser.pdf.byteLength} > ${pdfLimit}).`,
          "Lower document complexity or select a larger permitted limit.",
        );
      }
      const verified = await verifyPdf(browser.pdf, report.pages);
      report.bindings.pdfDigest = verified.digest;
      report.bindings.pdfByteDeterministic = false;
      report.bindings.volatilePdfMetadata = verified.volatileMetadata;
    }

    return {
      ...(browser.materialization.html === undefined ? {} : { html: browser.materialization.html }),
      ...(browser.pdf === undefined ? {} : { pdf: new Uint8Array(browser.pdf) }),
      pageCount: browser.materialization.pageCount,
      pageMap,
      renderReport: report,
      warnings: report.warnings,
      rendererFingerprint,
    };
  } catch (error) {
    const normalized = error instanceof DocxodusExportError
      ? error
      : new DocxodusExportError(
        "output_verification_failure",
        "output_verification",
        "The Node export boundary failed while verifying the materialized artifacts.",
        "Inspect the retained cause and retry with the supported runtime.",
        { cause: error },
      );
    throw attachFailedReport(normalized, report);
  }
}

export function renderDocxArtifacts(
  document: Uint8Array,
  options: RenderBatchOptions,
): Promise<RenderBatchResult> {
  const sourceBytes = ownedInput(document);
  const ownedOptions = snapshotNodeOptions(options);
  return renderOwned(
    sourceBytes,
    ownedOptions,
    (ownedOptions as RenderBatchOptions | undefined)?.outputs as readonly RenderOutput[],
  );
}

export function convertDocxToPdf(
  document: Uint8Array,
  options: NodeExportOptions,
): Promise<PdfExportResult> {
  const sourceBytes = ownedInput(document);
  const ownedOptions = snapshotNodeOptions(options);
  return renderOwned(sourceBytes, ownedOptions, ["pdf"]).then((result) => {
    if (!result.pdf) {
      exportError("pdf_write_failure", "pdf_print", "PDF output was not returned.",
        "Report this invariant failure.");
    }
    return { ...result, pdf: result.pdf };
  });
}

export function convertDocxToStandaloneHtml(
  document: Uint8Array,
  options: NodeExportOptions,
): Promise<PaginatedHtmlResult> {
  const sourceBytes = ownedInput(document);
  const ownedOptions = snapshotNodeOptions(options);
  return renderOwned(sourceBytes, ownedOptions, ["html"]).then((result) => {
    if (result.html === undefined) {
      exportError("conversion_failure", "output_verification", "HTML output was not returned.",
        "Report this invariant failure.");
    }
    return { ...result, html: result.html };
  });
}

export async function renderDocxFile(
  inputPath: string,
  destinations: RenderFileDestinations,
  options: NodeExportOptions,
): Promise<RenderFileResult> {
  const ownedOptions = snapshotNodeOptions(options);
  const ownedDestinations = snapshotDestinations(destinations);
  nodeOptionsPreflight(ownedOptions);
  const maximumBytes = ownedOptions.limits?.compressedDocxBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.compressedDocxBytes;
  const input = await readStableInputFile(inputPath, maximumBytes);
  const prepared = await prepareDestinations(input, ownedDestinations);
  const outputs: RenderOutput[] = [];
  if (ownedDestinations.htmlPath) outputs.push("html");
  if (ownedDestinations.pdfPath) outputs.push("pdf");
  const result = await renderOwned(input.bytes, ownedOptions, outputs);
  const payloads: Record<keyof RenderFileDestinations, Uint8Array | undefined> = {
    htmlPath: result.html === undefined ? undefined : Buffer.from(result.html, "utf8"),
    pdfPath: result.pdf,
    pageMapPath: ownedDestinations.pageMapPath ? canonicalJsonBytes(result.pageMap) : undefined,
    reportPath: ownedDestinations.reportPath ? canonicalJsonBytes(result.renderReport) : undefined,
  };
  const order: Array<keyof RenderFileDestinations> = [
    "htmlPath",
    "pdfPath",
    "pageMapPath",
    "reportPath",
  ];
  const committed: string[] = [];
  try {
    for (const kind of order) {
      const destination = prepared.find((entry) => entry.kind === kind);
      if (!destination) continue;
      const bytes = payloads[kind];
      if (!bytes) {
        exportError("output_write_failure", "output_write", `${kind} bytes are unavailable.`,
          "Report this invariant failure.");
      }
      await writeNoReplace(destination, bytes);
      committed.push(destination.resolvedPath);
    }
  } catch (error) {
    if (error instanceof DocxodusExportError) {
      const reported = attachFailedReport(error, result.renderReport) as DocxodusExportError;
      const committedDestinations = [...new Set([
        ...committed,
        ...reported.committedDestinations,
      ])];
      throw new DocxodusExportError(reported.code, reported.phase, reported.message, reported.remediation, {
        detail: reported.detail,
        cause: reported.cause,
        report: reported.report,
        committedDestinations,
      });
    }
    const normalized = new DocxodusExportError(
      "filesystem_failure",
      "filesystem_commit",
      "Artifact publication failed unexpectedly.",
      "Inspect the retained cause and destination filesystem.",
      { cause: error, committedDestinations: committed },
    );
    throw attachFailedReport(normalized, result.renderReport);
  }
  return {
    ...result,
    written: Object.fromEntries(prepared.map((entry) => [entry.kind, entry.resolvedPath])),
  };
}

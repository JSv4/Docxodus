import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_TIMEOUT_MS,
  type CompleteRenderReport,
  type ExportRuntimeAttestationEvidence,
  type ExportRuntimeObservedFacts,
  type ExportResourceLimits,
  type PaginatedHtmlOptions,
  type PaginatedHtmlResult,
} from "docxodus/export-browser";
import { canonicalJson, canonicalJsonBytes, sha256 } from "./canonical.js";
import {
  renderInBrowser,
  type BrowserRenderOutcome,
  type BrowserRuntimeIdentity,
} from "./browser-session.js";
import type {
  NodeExportOptions,
  NodeExportRuntime,
  PdfExportResult,
  RenderBatchOptions,
  RenderBatchResult,
  RenderEnvironmentAttestation,
  RenderFileDestinations,
  RenderFileResult,
  RenderOutput,
} from "./contracts.js";
import { isAbsolute, resolve } from "node:path";
import {
  CURRENT_RENDER_REPORT_SCHEMA,
  CURRENT_RENDER_REPORT_SCHEMA_VERSION,
  attachFailedReport,
  DocxodusExportError,
  exportError,
  hasCurrentRenderReportDiscriminator,
  isCurrentCompleteRenderReport,
} from "./contracts.js";
import {
  prepareDestinations,
  publishNoReplace,
  readStableInputFile,
} from "./files.js";
import { verifyPdf } from "./pdf.js";

export * from "./contracts.js";
export {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_RESOURCE_LIMITS,
  HARD_EXPORT_TIMEOUT_MS,
} from "docxodus/export-browser";

function ownedInput(document: Uint8Array, options: NodeExportOptions): Uint8Array {
  if (!(document instanceof Uint8Array)) {
    exportError(
      "invalid_argument",
      "input_validation",
      "The Node export API requires a Uint8Array DOCX snapshot.",
      "Read the file into a Uint8Array or use renderDocxFile().",
    );
  }
  const maximum = options.limits?.compressedDocxBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.compressedDocxBytes;
  if (document.byteLength > maximum) {
    exportError(
      "resource_limit",
      "input_validation",
      `compressedDocxBytes limit exceeded before snapshotting (${document.byteLength} > ${maximum}).`,
      "Use a smaller document or a deployment with a reviewed limits contract.",
    );
  }
  if (options.signal?.aborted) {
    exportError(
      "operation_cancelled",
      "input_validation",
      "Export was cancelled before input snapshotting.",
      "Retry with a non-aborted signal.",
      { pending: ["input snapshot"] },
    );
  }
  return new Uint8Array(document);
}

function exactKeys(record: Record<string, unknown>, allowed: readonly string[], label: string): void {
  const extra = Object.keys(record).filter((key) => !allowed.includes(key));
  if (extra.length > 0) {
    exportError(
      "invalid_argument",
      "input_validation",
      `${label} contains unknown fields: ${extra.join(", ")}.`,
      "Use the documented attestation schema without extension fields.",
    );
  }
}

function wellFormed(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return false;
  }
  return true;
}

function materialDigest(domain: string, value: unknown): string {
  if (!/^[\x20-\x7e]+$/.test(domain)) throw new TypeError("Digest domains must be printable ASCII.");
  return sha256(`${domain}\0${canonicalJson(value)}`);
}

function nonEmptyString(value: unknown, label: string): string {
  if (typeof value !== "string" || value.trim() === "" || !wellFormed(value)) {
    exportError(
      "invalid_argument",
      "input_validation",
      `${label} must be a non-empty, well-formed Unicode string.`,
      "Correct the runtime attestation and retry.",
    );
  }
  return value;
}

function digestString(value: unknown, label: string): string {
  const digest = nonEmptyString(value, label);
  if (!/^[0-9a-f]{64}$/.test(digest)) {
    exportError(
      "invalid_argument",
      "input_validation",
      `${label} must be a lower-case SHA-256 digest.`,
      "Provide exactly 64 hexadecimal characters.",
    );
  }
  return digest;
}

function validateEnvironmentAttestation(
  value: RenderEnvironmentAttestation | undefined,
): RenderEnvironmentAttestation | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    exportError("invalid_argument", "input_validation", "environmentAttestation must be an object.",
      "Provide the documented canonical JSON object.");
  }
  const record = value as unknown as Record<string, unknown>;
  exactKeys(record, [
    "schemaVersion",
    "usage",
    "chromiumProduct",
    "chromiumBuild",
    "executableSha256",
    "launchFlags",
    "hostFonts",
    "basis",
  ], "environmentAttestation");
  if (record.schemaVersion !== 1 || record.usage !== "docxodus-render-environment") {
    exportError(
      "invalid_argument",
      "input_validation",
      "environmentAttestation must use schemaVersion 1 and the docxodus-render-environment usage.",
      "Provide the documented versioned environment attestation.",
    );
  }
  const launchFlags = record.launchFlags;
  const hostFonts = record.hostFonts;
  if (!Array.isArray(launchFlags) || !launchFlags.every((item) =>
    typeof item === "string" && item.trim() !== "")) {
    exportError("invalid_argument", "input_validation", "launchFlags must be a string array.",
      "Record every attested Chromium launch flag.");
  }
  if (new Set(launchFlags).size !== launchFlags.length) {
    exportError(
      "invalid_argument",
      "input_validation",
      "launchFlags contains a duplicate value.",
      "Record every effective launch flag exactly once and in launch order.",
    );
  }
  if (!Array.isArray(hostFonts)) {
    exportError("invalid_argument", "input_validation", "hostFonts must be an array.",
      "Provide an empty array when no host fonts are attested.");
  }
  const seen = new Set<string>();
  const resolvedFonts = hostFonts.map((font, index) => {
    if (!font || typeof font !== "object" || Array.isArray(font)) {
      exportError("invalid_argument", "input_validation", `hostFonts[${index}] must be an object.`,
        "Correct the environment attestation.");
    }
    const item = font as Record<string, unknown>;
    exactKeys(
      item,
      ["family", "postscriptName", "style", "weight", "stretch", "fileSha256", "version"],
      `hostFonts[${index}]`,
    );
    if (!Number.isSafeInteger(item.weight) || (item.weight as number) < 1
      || (item.weight as number) > 1000) {
      exportError("invalid_argument", "input_validation", `hostFonts[${index}].weight is invalid.`,
        "Use an integer font weight from 1 through 1000.");
    }
    if (item.style !== "normal" && item.style !== "italic" && item.style !== "oblique") {
      exportError("invalid_argument", "input_validation", `hostFonts[${index}].style is invalid.`,
        "Use normal, italic, or oblique.");
    }
    if (typeof item.stretch !== "number" || !Number.isFinite(item.stretch)
      || item.stretch <= 0 || item.stretch > 1000) {
      exportError("invalid_argument", "input_validation", `hostFonts[${index}].stretch is invalid.`,
        "Use a finite positive font-stretch percentage no greater than 1000.");
    }
    const fileSha256 = digestString(item.fileSha256, `hostFonts[${index}].fileSha256`);
    const resolved = {
      family: nonEmptyString(item.family, `hostFonts[${index}].family`),
      postscriptName: nonEmptyString(item.postscriptName, `hostFonts[${index}].postscriptName`),
      style: item.style as "normal" | "italic" | "oblique",
      weight: item.weight as number,
      stretch: item.stretch,
      fileSha256,
      version: nonEmptyString(item.version, `hostFonts[${index}].version`),
    };
    const faceKey = canonicalJson(resolved);
    if (seen.has(faceKey)) {
      exportError("invalid_argument", "input_validation", "hostFonts contains a duplicate face identity.",
        "List each attested host font face exactly once.");
    }
    seen.add(faceKey);
    return resolved;
  });
  return Object.freeze({
    schemaVersion: 1 as const,
    usage: "docxodus-render-environment" as const,
    chromiumProduct: nonEmptyString(record.chromiumProduct, "chromiumProduct"),
    chromiumBuild: nonEmptyString(record.chromiumBuild, "chromiumBuild"),
    ...(record.executableSha256 === undefined
      ? {}
      : { executableSha256: digestString(record.executableSha256, "executableSha256") }),
    launchFlags: Object.freeze([...(launchFlags as string[])]),
    hostFonts: Object.freeze(resolvedFonts.sort((left, right) => {
      const leftKey = canonicalJson(left);
      const rightKey = canonicalJson(right);
      return leftKey < rightKey ? -1 : leftKey > rightKey ? 1 : 0;
    })),
    basis: nonEmptyString(record.basis, "basis"),
  });
}

function validateRuntime(runtime: NodeExportRuntime): NodeExportRuntime {
  if (runtime.browser && runtime.browserExecutablePath) {
    exportError(
      "invalid_argument",
      "input_validation",
      "browser and browserExecutablePath cannot be supplied together.",
      "Inject one caller-owned Chromium browser or provide one executable path.",
    );
  }
  if (runtime.browser !== undefined
    && (!runtime.browser || typeof runtime.browser !== "object"
      || typeof runtime.browser.browserType !== "function"
      || typeof runtime.browser.isConnected !== "function"
      || typeof runtime.browser.newContext !== "function")) {
    exportError(
      "invalid_argument",
      "input_validation",
      "browser must be a Playwright Browser object.",
      "Inject a connected Playwright Chromium Browser instance.",
    );
  }
  if (runtime.browserExecutablePath !== undefined
    && (typeof runtime.browserExecutablePath !== "string"
      || runtime.browserExecutablePath.trim() === ""
      || !isAbsolute(runtime.browserExecutablePath))) {
    exportError(
      "invalid_argument",
      "input_validation",
      "browserExecutablePath must be a non-empty absolute path.",
      "Provide an absolute Chromium executable path or omit it to use the pinned browser.",
    );
  }
  const fontDirectories = runtime.fontDirectories ?? [];
  if (!Array.isArray(fontDirectories) || !fontDirectories.every((entry) =>
    typeof entry === "string" && entry.trim() !== "")) {
    exportError("invalid_argument", "input_validation", "fontDirectories must be a string array.",
      "Provide each explicit font directory as a non-empty path.");
  }
  const attestations = runtime.fontLicenseAttestations ?? [];
  if (!Array.isArray(attestations)) {
    exportError("invalid_argument", "input_validation", "fontLicenseAttestations must be an array.",
      "Provide the documented attestation objects.");
  }
  const attestedFiles = new Set<string>();
  for (const [index, attestation] of attestations.entries()) {
    if (!attestation || typeof attestation !== "object" || Array.isArray(attestation)) {
      exportError(
        "invalid_argument",
        "input_validation",
        `fontLicenseAttestations[${index}] must be an object.`,
        "Provide the documented font-license attestation schema.",
      );
    }
    const record = attestation as unknown as Record<string, unknown>;
    exactKeys(
      record,
      [
        "schemaVersion",
        "usage",
        "fileSha256",
        "embeddingPermitted",
        "permittedOutputs",
        "subsettingPermitted",
        "basis",
        "attester",
      ],
      `fontLicenseAttestations[${index}]`,
    );
    if (record.schemaVersion !== 1
      || record.usage !== "standalone-document-font-embedding") {
      exportError(
        "invalid_argument",
        "input_validation",
        `fontLicenseAttestations[${index}] has an unsupported schema or usage.`,
        "Use the versioned standalone-document-font-embedding attestation.",
      );
    }
    const fileSha256 = digestString(
      record.fileSha256,
      `fontLicenseAttestations[${index}].fileSha256`,
    );
    if (attestedFiles.has(fileSha256)) {
      exportError(
        "invalid_argument",
        "input_validation",
        "fontLicenseAttestations contains a duplicate file digest.",
        "Provide exactly one embedding decision per font file.",
      );
    }
    attestedFiles.add(fileSha256);
    if (record.embeddingPermitted !== true) {
      exportError(
        "invalid_argument",
        "input_validation",
        `fontLicenseAttestations[${index}].embeddingPermitted must be true.`,
        "Do not load a font unless embedding permission is affirmatively attested.",
      );
    }
    if (!Array.isArray(record.permittedOutputs) || record.permittedOutputs.length === 0
      || record.permittedOutputs.length > 2
      || record.permittedOutputs.some((output) => output !== "html" && output !== "pdf")
      || new Set(record.permittedOutputs).size !== record.permittedOutputs.length) {
      exportError(
        "invalid_argument",
        "input_validation",
        `fontLicenseAttestations[${index}].permittedOutputs is invalid.`,
        "List html, pdf, or both exactly once.",
      );
    }
    if (record.subsettingPermitted !== true && record.subsettingPermitted !== false) {
      exportError(
        "invalid_argument",
        "input_validation",
        `fontLicenseAttestations[${index}].subsettingPermitted must be boolean.`,
        "Record the font license's exact subsetting permission.",
      );
    }
    nonEmptyString(record.basis, `fontLicenseAttestations[${index}].basis`);
    if (record.attester !== undefined) {
      nonEmptyString(record.attester, `fontLicenseAttestations[${index}].attester`);
    }
  }
  if (fontDirectories.length > 0) {
    exportError(
      "unsupported_runtime",
      "font_loading",
      "Explicit font-directory injection is not available in this renderer version.",
      "Use the browser-observed font mode until issue #442 lands the verified font runtime.",
    );
  }
  if (attestations.length > 0 && fontDirectories.length === 0) {
    exportError(
      "invalid_argument",
      "input_validation",
      "Font-license attestations require at least one font directory.",
      "Remove the unattached attestations or provide the matching directory after #442.",
    );
  }
  return {
    browser: runtime.browser,
    browserExecutablePath: runtime.browserExecutablePath,
    fontDirectories: [],
    fontLicenseAttestations: [],
    environmentAttestation: validateEnvironmentAttestation(runtime.environmentAttestation),
  };
}

function nodeOptionsPreflight(options: NodeExportOptions, allowOutputs: boolean): void {
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    exportError(
      "invalid_argument",
      "input_validation",
      "Export options are required.",
      "Supply explicit reviewProfile and commentProfile values.",
    );
  }
  exactKeys(options as unknown as Record<string, unknown>, [
    "documentVersion",
    "expectedSourceDigest",
    "reviewProfile",
    "reviewProfileAlreadyApplied",
    "commentProfile",
    "title",
    "unsupportedContent",
    "strictFonts",
    "timeoutMs",
    "limits",
    "signal",
    "browser",
    "browserExecutablePath",
    "fontDirectories",
    "fontLicenseAttestations",
    "environmentAttestation",
    ...(allowOutputs ? ["outputs"] : []),
  ], "Export options");
  if (options.reviewProfile !== "final"
    && options.reviewProfile !== "original"
    && options.reviewProfile !== "markup") {
    exportError("invalid_argument", "input_validation", "reviewProfile is invalid.",
      "Use final, original, or markup.");
  }
  if (options.commentProfile !== "hidden"
    && options.commentProfile !== "inline"
    && options.commentProfile !== "endnotes"
    && options.commentProfile !== "margin") {
    exportError("invalid_argument", "input_validation", "commentProfile is invalid.",
      "Use hidden, inline, endnotes, or margin.");
  }
  if (options.reviewProfileAlreadyApplied !== undefined
    && typeof options.reviewProfileAlreadyApplied !== "boolean") {
    exportError(
      "invalid_argument",
      "input_validation",
      "reviewProfileAlreadyApplied must be boolean.",
      "Omit it or pass true only for exact policy-derived final/original bytes.",
    );
  }
  if (options.reviewProfileAlreadyApplied === true && options.reviewProfile === "markup") {
    exportError(
      "invalid_argument",
      "input_validation",
      "reviewProfileAlreadyApplied is invalid with the markup profile.",
      "Use unchanged source bytes for markup, or choose final/original.",
    );
  }
  if (options.documentVersion !== undefined
    && (!Number.isSafeInteger(options.documentVersion) || options.documentVersion < 0)) {
    exportError(
      "document_version_unrepresentable",
      "input_validation",
      "documentVersion must be a non-negative JavaScript safe integer.",
      "Use a value between 0 and Number.MAX_SAFE_INTEGER.",
    );
  }
  if (options.expectedSourceDigest !== undefined
    && !/^[0-9a-f]{64}$/.test(options.expectedSourceDigest)) {
    exportError(
      "invalid_argument",
      "input_validation",
      "expectedSourceDigest must be a lower-case SHA-256 hex digest.",
      "Supply exactly 64 lower-case hexadecimal characters.",
    );
  }
  if (options.unsupportedContent !== undefined
    && options.unsupportedContent !== "warn"
    && options.unsupportedContent !== "strict") {
    exportError("invalid_argument", "input_validation", "unsupportedContent is invalid.",
      "Use warn or strict.");
  }
  if (options.strictFonts !== undefined && typeof options.strictFonts !== "boolean") {
    exportError("invalid_argument", "input_validation", "strictFonts must be a boolean.",
      "Use true or false.");
  }
  if (options.title !== undefined
    && (typeof options.title !== "string" || !wellFormed(options.title))) {
    exportError("invalid_argument", "input_validation", "title must be a well-formed Unicode string.",
      "Provide plain document-title text without unpaired UTF-16 surrogates, or omit it.");
  }
  if (options.signal !== undefined
    && (typeof options.signal !== "object"
      || typeof options.signal.addEventListener !== "function"
      || typeof options.signal.removeEventListener !== "function"
      || typeof options.signal.aborted !== "boolean")) {
    exportError(
      "invalid_argument",
      "input_validation",
      "signal must be an AbortSignal.",
      "Pass a standards-compliant AbortSignal or omit it.",
    );
  }

  const suppliedLimits = options.limits ?? {};
  if (!suppliedLimits || typeof suppliedLimits !== "object" || Array.isArray(suppliedLimits)) {
    exportError("invalid_argument", "input_validation", "limits must be an object.",
      "Use keys from ExportResourceLimits with positive integer values.");
  }
  for (const [name, value] of Object.entries(suppliedLimits)) {
    if (!(name in DEFAULT_EXPORT_RESOURCE_LIMITS)) {
      exportError("invalid_argument", "input_validation", `Unknown export limit: ${name}.`,
        "Use a key from ExportResourceLimits.");
    }
    const key = name as keyof ExportResourceLimits;
    if (!Number.isSafeInteger(value) || (value as number) <= 0) {
      exportError(
        "invalid_argument",
        "input_validation",
        `Export limit ${name} must be a positive safe integer.`,
        "Supply a positive integer no greater than the published default.",
      );
    }
    if ((value as number) > DEFAULT_EXPORT_RESOURCE_LIMITS[key]) {
      exportError(
        "invalid_argument",
        "input_validation",
        `Export limit ${name} may only lower the default.`,
        `Use ${DEFAULT_EXPORT_RESOURCE_LIMITS[key]} or less.`,
      );
    }
  }
  const resolvedLimits = { ...DEFAULT_EXPORT_RESOURCE_LIMITS, ...suppliedLimits };
  if (Array.isArray(options.fontDirectories)
    && options.fontDirectories.length > resolvedLimits.fontDirectoryEntries) {
    exportError(
      "resource_limit",
      "input_validation",
      `fontDirectoryEntries limit exceeded (${options.fontDirectories.length} > ${resolvedLimits.fontDirectoryEntries}).`,
      "Use fewer configured font directories.",
    );
  }
  if (Array.isArray(options.fontLicenseAttestations)
    && options.fontLicenseAttestations.length > resolvedLimits.fontFiles) {
    exportError(
      "resource_limit",
      "input_validation",
      `fontFiles limit exceeded by license attestations (${options.fontLicenseAttestations.length} > ${resolvedLimits.fontFiles}).`,
      "Use fewer font attestations.",
    );
  }
  const environment = options.environmentAttestation;
  if (environment && typeof environment === "object" && !Array.isArray(environment)) {
    if (Array.isArray(environment.hostFonts)
      && environment.hostFonts.length > resolvedLimits.fontFiles) {
      exportError(
        "resource_limit",
        "input_validation",
        `fontFiles limit exceeded by environment attestation (${environment.hostFonts.length} > ${resolvedLimits.fontFiles}).`,
        "Attest fewer host font files.",
      );
    }
    if (Array.isArray(environment.launchFlags)
      && environment.launchFlags.length > resolvedLimits.renderDiagnostics) {
      exportError(
        "resource_limit",
        "input_validation",
        `renderDiagnostics limit exceeded by launch flags (${environment.launchFlags.length} > ${resolvedLimits.renderDiagnostics}).`,
        "Attest a bounded Chromium launch configuration.",
      );
    }
  }
  if (allowOutputs) validateOutputs((options as RenderBatchOptions).outputs);
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
    && actualDigest !== options.expectedSourceDigest) {
    exportError(
      "source_digest_mismatch",
      "package_preflight",
      "The source digest does not match expectedSourceDigest.",
      "Render the exact verified source bytes or update the expected digest.",
      {
        detail: `expected=${options.expectedSourceDigest}; actual=${actualDigest}`,
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
    exportError("invalid_argument", "input_validation", "destinations must be an object.",
      "Provide at least one documented artifact path.");
  }
  const allowed = ["htmlPath", "pdfPath", "pageMapPath", "reportPath"];
  const unknown = Object.keys(destinations).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) {
    exportError(
      "invalid_argument",
      "input_validation",
      `destinations contains unknown fields: ${unknown.join(", ")}.`,
      "Use only htmlPath, pdfPath, pageMapPath, and reportPath.",
    );
  }
  for (const [key, value] of Object.entries(destinations)) {
    if (value !== undefined && (typeof value !== "string" || value.trim() === "")) {
      exportError(
        "invalid_argument",
        "input_validation",
        `destinations.${key} must be a non-empty string when provided.`,
        "Remove the field or provide a valid artifact path.",
      );
    }
  }
  return Object.fromEntries(
    Object.entries(destinations)
      .filter((entry): entry is [string, string] => entry[1] !== undefined)
      .map(([key, value]) => [key, resolve(value)]),
  ) as RenderFileDestinations;
}

function validateOutputs(outputs: readonly RenderOutput[]): RenderOutput[] {
  if (!Array.isArray(outputs)) {
    exportError("invalid_argument", "input_validation", "outputs must be an array.",
      "Request html, pdf, or both.");
  }
  if (outputs.length > 2) {
    exportError(
      "invalid_argument",
      "input_validation",
      "outputs may contain at most html and pdf.",
      "Request each supported output at most once.",
    );
  }
  const resolved: RenderOutput[] = [];
  for (const output of outputs) {
    if (output !== "html" && output !== "pdf") {
      exportError("invalid_argument", "input_validation", `Unknown output kind: ${String(output)}.`,
        "Use html or pdf.");
    }
    if (resolved.includes(output)) {
      exportError("invalid_argument", "input_validation", `Duplicate output kind: ${output}.`,
        "List each requested output once.");
    }
    resolved.push(output);
  }
  return [
    ...(resolved.includes("html") ? ["html" as const] : []),
    ...(resolved.includes("pdf") ? ["pdf" as const] : []),
  ];
}

function normalizedTimeout(options: NodeExportOptions): number {
  const timeout = options.timeoutMs ?? DEFAULT_EXPORT_TIMEOUT_MS;
  if (!Number.isSafeInteger(timeout) || timeout <= 0 || timeout > HARD_EXPORT_TIMEOUT_MS) {
    exportError(
      "invalid_argument",
      "input_validation",
      `timeoutMs must be an integer from 1 through ${HARD_EXPORT_TIMEOUT_MS}.`,
      "Use the shared export timeout contract.",
    );
  }
  return timeout;
}

function browserOptions(options: NodeExportOptions): Omit<PaginatedHtmlOptions, "wasmBasePath"> {
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

function runtimeEnvironment(
  report: CompleteRenderReport,
  runtime: BrowserRuntimeIdentity,
  attestation: RenderEnvironmentAttestation | undefined,
): CompleteRenderReport["environment"] {
  const slash = runtime.browserVersion.indexOf("/");
  const browserProduct = slash > 0 ? runtime.browserVersion.slice(0, slash) : "Chromium";
  const browserBuild = slash > 0 ? runtime.browserVersion.slice(slash + 1) : runtime.browserVersion;
  const observed: ExportRuntimeObservedFacts = {
    ...report.environment.observed,
    runtimeKind: "nodeChromium",
    playwrightVersion: runtime.playwrightVersion,
    browserProduct,
    browserBuild,
    ...(runtime.executableDigest === undefined
      ? {}
      : { executableSha256: runtime.executableDigest }),
    launchFlags: [...runtime.launchFlags],
    operatingSystem: runtime.platform,
    architecture: runtime.architecture,
    networkIsolation: runtime.launchMode === "injected"
      ? "contextRestricted"
      : "ownedProcessRestricted",
  };
  const nodeFontsVerified = report.fonts.every((font) =>
    font.status === "resolved" && (font.source === "embedded" || font.source === "configured"));
  const baselineVerification = runtime.launchMode !== "injected" && nodeFontsVerified
    ? "nodeVerified"
    : "browserObserved";
  const fidelityTier = runtime.launchMode === "pinned"
    ? runtime.platform === "linux" && runtime.architecture === "x64"
      ? "releaseBaselined"
      : "experimental"
    : "unbaselined";
  if (!attestation || attestation.executableSha256 === undefined) {
    return {
      ...report.environment,
      verification: baselineVerification,
      fidelityTier,
      observed,
    };
  }

  const attestedFamilies = new Set(attestation.hostFonts.map((font) =>
    font.family.normalize("NFC").toLowerCase()));
  const uncoveredFont = report.fonts.find((font) => font.source === "browser"
    && !attestedFamilies.has((font.resolvedFamily ?? font.requestedFamily).normalize("NFC").toLowerCase()));
  const flagsMatch = runtime.launchMode === "injected"
    || (runtime.launchFlags.length === attestation.launchFlags.length
      && runtime.launchFlags.every((flag, index) => flag === attestation.launchFlags[index]));
  const digestMatches = runtime.executableDigest === undefined
    || runtime.executableDigest === attestation.executableSha256;
  if (attestation.chromiumProduct !== browserProduct
    || attestation.chromiumBuild !== browserBuild
    || !flagsMatch || !digestMatches || uncoveredFont) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The caller environment attestation does not match the observed Chromium render runtime.",
      "Regenerate the attestation from this exact executable, launch policy, and host-font set.",
      {
        detail: uncoveredFont
          ? `unattestedFont=${uncoveredFont.resolvedFamily ?? uncoveredFont.requestedFamily}`
          : "Chromium product, build, executable digest, or launch flags differ.",
      },
    );
  }
  const attested: ExportRuntimeAttestationEvidence = {
    chromiumProduct: attestation.chromiumProduct,
    chromiumBuild: attestation.chromiumBuild,
    executableSha256: attestation.executableSha256,
    launchFlags: [...attestation.launchFlags],
    hostFontsDigest: materialDigest("docxodus:font-configuration:v1", attestation.hostFonts),
    basis: attestation.basis,
  };
  return {
    ...report.environment,
    verification: "callerAttested",
    fidelityTier,
    observed,
    attested,
    attestationDigest: materialDigest("docxodus:environment-attestation:v1", attestation),
  };
}

function verifyBrowserOutcome(
  browser: BrowserRenderOutcome,
  requested: readonly RenderOutput[],
): void {
  const materialization = browser.materialization;
  const expectedOutputs = [
    ...(requested.includes("html") ? ["html" as const] : []),
    ...(requested.includes("pdf") ? ["pdf" as const] : []),
  ];
  const report = materialization.renderReport;
  if (!isCurrentCompleteRenderReport(report)) {
    const version = hasCurrentRenderReportDiscriminator(report)
      ? "a malformed v2 complete report"
      : "an unsupported report discriminator";
    exportError(
      "output_verification_failure",
      "output_verification",
      `The browser materializer returned ${version}; expected ${CURRENT_RENDER_REPORT_SCHEMA} version ${CURRENT_RENDER_REPORT_SCHEMA_VERSION}.`,
      "Use matching hardened docxodus and @docxodus/export package versions; legacy v1 is validation-only.",
    );
  }
  const pageMapDigest = sha256(canonicalJson(materialization.pageMap));
  const outputsMatch = canonicalJson(report.options.outputs) === canonicalJson(expectedOutputs);
  const pageInventoriesMatch = canonicalJson(report.pages) === canonicalJson(materialization.pageMap.pages);
  if (!Number.isSafeInteger(materialization.pageCount) || materialization.pageCount < 1
    || materialization.pageCount !== materialization.pageMap.pages.length
    || materialization.pageCount !== report.pages.length
    || materialization.rendererFingerprint !== materialization.pageMap.rendererFingerprint
    || materialization.rendererFingerprint !== report.environment.rendererFingerprint
    || report.bindings.pageMapDigest !== pageMapDigest
    || !outputsMatch || !pageInventoriesMatch
    || canonicalJson(materialization.warnings) !== canonicalJson(report.warnings)) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The browser materializer returned inconsistent report, PageMap, or renderer identity bindings.",
      "Use matching hardened docxodus and @docxodus/export package versions.",
    );
  }
  if (expectedOutputs.includes("html")) {
    if (materialization.html === undefined
      || report.bindings.htmlDigest !== sha256(materialization.html)) {
      exportError(
        "output_verification_failure",
        "output_verification",
        "The browser materializer HTML bytes do not match the reported digest.",
        "Use the exact finalized standalone HTML returned by the verified materializer.",
      );
    }
  } else if (materialization.html !== undefined || report.bindings.htmlDigest !== undefined) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The browser materializer returned an unrequested HTML artifact or digest.",
      "Bind reports only to selected output artifacts.",
    );
  }
  if (expectedOutputs.includes("pdf") !== (browser.pdf !== undefined)) {
    exportError(
      "output_verification_failure",
      "output_verification",
      "The browser materializer PDF presence does not match the selected outputs.",
      "Return exactly one PDF byte stream when PDF output is selected.",
    );
  }
}

function enforceSerializedLimit(
  value: string,
  maximum: number,
  name: "pageMapOutputBytes" | "renderReportOutputBytes",
): void {
  const actual = Buffer.byteLength(value, "utf8");
  if (actual > maximum) {
    exportError(
      "resource_limit",
      "output_verification",
      `${name} limit exceeded (${actual} > ${maximum}).`,
      "Use a smaller document or lower-complexity diagnostics.",
    );
  }
}

async function renderOwned(
  sourceBytes: Uint8Array,
  options: NodeExportOptions,
  outputs: readonly RenderOutput[],
  allowOutputs: boolean,
): Promise<RenderBatchResult> {
  const requested = validateOutputs(outputs);
  nodeOptionsPreflight(options, allowOutputs);
  sourcePreflight(sourceBytes, options);
  const runtime = validateRuntime(options);
  const timeoutMs = normalizedTimeout(options);
  const deadline = performance.now() + timeoutMs;
  const pdfLimit = options.limits?.pdfOutputBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.pdfOutputBytes;
  const parserLimit = options.limits?.pdfParserExpandedBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.pdfParserExpandedBytes;
  let report: CompleteRenderReport | undefined;
  try {
    const browser = await renderInBrowser(
      sourceBytes,
      browserOptions({ ...options, timeoutMs }),
      runtime,
      requested.includes("html"),
      requested.includes("pdf"),
      deadline,
      pdfLimit,
      options.signal,
    );
    verifyBrowserOutcome(browser, requested);
    const attestation = runtime.environmentAttestation;
    report = structuredClone(browser.materialization.renderReport);
    const environment = runtimeEnvironment(report, browser.runtime, attestation);
    const rendererFingerprint = materialDigest("docxodus:renderer-fingerprint:v1", {
      schemaVersion: 1,
      browserMaterializerFingerprint: browser.materialization.rendererFingerprint,
      runtime: browser.runtime,
      environment: {
        verification: environment.verification,
        fidelityTier: environment.fidelityTier,
        observed: environment.observed,
        attested: environment.attested,
        attestationDigest: environment.attestationDigest,
      },
    });
    const pageMap = structuredClone(browser.materialization.pageMap);
    pageMap.rendererFingerprint = rendererFingerprint;
    environment.rendererFingerprint = rendererFingerprint;
    report.environment = {
      ...environment,
    };
    report.options.outputs = [...requested];
    const pageMapJson = canonicalJson(pageMap);
    enforceSerializedLimit(
      pageMapJson,
      options.limits?.pageMapOutputBytes ?? DEFAULT_EXPORT_RESOURCE_LIMITS.pageMapOutputBytes,
      "pageMapOutputBytes",
    );
    report.bindings.pageMapDigest = sha256(pageMapJson);
    report.bindings.artifactRequestIds = [];
    if (requested.includes("html")) {
      report.bindings.htmlDigest = sha256(browser.materialization.html!);
    } else {
      delete report.bindings.htmlDigest;
    }

    if (browser.pdf) {
      if (browser.pdf.byteLength > pdfLimit) {
        exportError(
          "resource_limit",
          "output_verification",
          `PDF output exceeds pdfOutputBytes (${browser.pdf.byteLength} > ${pdfLimit}).`,
          "Lower document complexity or select a larger permitted limit.",
        );
      }
      const verified = await verifyPdf(browser.pdf, report.pages, parserLimit);
      report.bindings.pdfDigest = verified.digest;
      report.bindings.pdfByteDeterministic = false;
      report.bindings.volatilePdfMetadata = verified.volatileMetadata;
    } else {
      delete report.bindings.pdfDigest;
      delete report.bindings.pdfByteDeterministic;
      delete report.bindings.volatilePdfMetadata;
    }

    const reportJson = canonicalJson(report);
    enforceSerializedLimit(
      reportJson,
      options.limits?.renderReportOutputBytes ?? DEFAULT_EXPORT_RESOURCE_LIMITS.renderReportOutputBytes,
      "renderReportOutputBytes",
    );

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
  nodeOptionsPreflight(options, true);
  const ownedOptions = snapshotNodeOptions(options);
  const sourceBytes = ownedInput(document, ownedOptions);
  return renderOwned(
    sourceBytes,
    ownedOptions,
    (ownedOptions as RenderBatchOptions | undefined)?.outputs as readonly RenderOutput[],
    true,
  );
}

export function convertDocxToPdf(
  document: Uint8Array,
  options: NodeExportOptions,
): Promise<PdfExportResult> {
  nodeOptionsPreflight(options, false);
  const ownedOptions = snapshotNodeOptions(options);
  const sourceBytes = ownedInput(document, ownedOptions);
  return renderOwned(sourceBytes, ownedOptions, ["pdf"], false).then((result) => {
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
  nodeOptionsPreflight(options, false);
  const ownedOptions = snapshotNodeOptions(options);
  const sourceBytes = ownedInput(document, ownedOptions);
  return renderOwned(sourceBytes, ownedOptions, ["html"], false).then((result) => {
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
  nodeOptionsPreflight(options, false);
  const ownedOptions = snapshotNodeOptions(options);
  const ownedDestinations = snapshotDestinations(destinations);
  const maximumBytes = ownedOptions.limits?.compressedDocxBytes
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.compressedDocxBytes;
  const input = await readStableInputFile(inputPath, maximumBytes);
  const prepared = await prepareDestinations(input, ownedDestinations);
  const outputs: RenderOutput[] = [];
  if (ownedDestinations.htmlPath) outputs.push("html");
  if (ownedDestinations.pdfPath) outputs.push("pdf");
  const result = await renderOwned(input.bytes, ownedOptions, outputs, false);
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
  try {
    const publications = [];
    for (const kind of order) {
      const destination = prepared.find((entry) => entry.kind === kind);
      if (!destination) continue;
      const bytes = payloads[kind];
      if (!bytes) {
        exportError("output_write_failure", "output_write", `${kind} bytes are unavailable.`,
          "Report this invariant failure.");
      }
      publications.push({ destination, bytes });
    }
    await publishNoReplace(publications, ownedOptions.signal);
  } catch (error) {
    if (error instanceof DocxodusExportError) {
      const reported = attachFailedReport(error, result.renderReport) as DocxodusExportError;
      throw new DocxodusExportError(reported.code, reported.phase, reported.message, reported.remediation, {
        detail: reported.detail,
        pending: reported.pending,
        partUri: reported.partUri,
        anchorId: reported.anchorId,
        resource: reported.resource,
        cause: reported.cause,
        report: reported.report,
        committedDestinations: reported.committedDestinations,
      });
    }
    const normalized = new DocxodusExportError(
      "filesystem_failure",
      "filesystem_commit",
      "Artifact publication failed unexpectedly.",
      "Inspect the retained cause and destination filesystem.",
      { cause: error },
    );
    throw attachFailedReport(normalized, result.renderReport);
  }
  return {
    ...result,
    written: Object.fromEntries(prepared.map((entry) => [entry.kind, entry.resolvedPath])),
  };
}

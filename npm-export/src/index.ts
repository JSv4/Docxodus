import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_TIMEOUT_MS,
  type CompleteRenderReport,
  type ExportResourceLimits,
  type PaginatedHtmlOptions,
  type PaginatedHtmlResult,
} from "docxodus/export-browser";
import { canonicalJson, canonicalJsonBytes, sha256 } from "./canonical.js";
import { renderInBrowser } from "./browser-session.js";
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

export * from "./contracts.js";
export {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_RESOURCE_LIMITS,
  HARD_EXPORT_TIMEOUT_MS,
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
  return value;
}

function digestString(value: unknown, label: string): string {
  const digest = nonEmptyString(value, label).toLowerCase();
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
  if (!Array.isArray(launchFlags) || !launchFlags.every((item) =>
    typeof item === "string" && item.trim() !== "")) {
    exportError("invalid_document", "input_validation", "launchFlags must be a string array.",
      "Record every attested Chromium launch flag.");
  }
  if (!Array.isArray(hostFonts)) {
    exportError("invalid_document", "input_validation", "hostFonts must be an array.",
      "Provide an empty array when no host fonts are attested.");
  }
  const seen = new Set<string>();
  const resolvedFonts = hostFonts.map((font, index) => {
    if (!font || typeof font !== "object" || Array.isArray(font)) {
      exportError("invalid_document", "input_validation", `hostFonts[${index}] must be an object.`,
        "Correct the environment attestation.");
    }
    const item = font as Record<string, unknown>;
    exactKeys(item, ["family", "style", "weight", "fileSha256", "version"], `hostFonts[${index}]`);
    if (!Number.isSafeInteger(item.weight) || (item.weight as number) < 1
      || (item.weight as number) > 1000) {
      exportError("invalid_document", "input_validation", `hostFonts[${index}].weight is invalid.`,
        "Use an integer font weight from 1 through 1000.");
    }
    const fileSha256 = digestString(item.fileSha256, `hostFonts[${index}].fileSha256`);
    if (seen.has(fileSha256)) {
      exportError("invalid_document", "input_validation", "hostFonts contains a duplicate file digest.",
        "List each attested font file exactly once.");
    }
    seen.add(fileSha256);
    return {
      family: nonEmptyString(item.family, `hostFonts[${index}].family`),
      style: nonEmptyString(item.style, `hostFonts[${index}].style`),
      weight: item.weight as number,
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
    launchFlags: Object.freeze([...(launchFlags as string[])]),
    hostFonts: Object.freeze(resolvedFonts),
    basis: nonEmptyString(record.basis, "basis"),
  });
}

function validateRuntime(runtime: NodeExportRuntime): NodeExportRuntime {
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
  if (!Array.isArray(fontDirectories) || !fontDirectories.every((entry) =>
    typeof entry === "string" && entry.trim() !== "")) {
    exportError("invalid_document", "input_validation", "fontDirectories must be a string array.",
      "Provide each explicit font directory as a non-empty path.");
  }
  const attestations = runtime.fontLicenseAttestations ?? [];
  if (!Array.isArray(attestations)) {
    exportError("invalid_document", "input_validation", "fontLicenseAttestations must be an array.",
      "Provide the documented attestation objects.");
  }
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
      ["fileSha256", "embeddingPermitted", "basis", "attester"],
      `fontLicenseAttestations[${index}]`,
    );
    digestString(record.fileSha256, `fontLicenseAttestations[${index}].fileSha256`);
    if (record.embeddingPermitted !== true) {
      exportError(
        "invalid_document",
        "input_validation",
        `fontLicenseAttestations[${index}].embeddingPermitted must be true.`,
        "Do not load a font unless embedding permission is affirmatively attested.",
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
      "invalid_document",
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

function nodeOptionsPreflight(options: NodeExportOptions): void {
  if (!options || typeof options !== "object" || Array.isArray(options)) {
    exportError(
      "invalid_document",
      "input_validation",
      "Export options are required.",
      "Supply explicit reviewProfile and commentProfile values.",
    );
  }
  if (options.reviewProfile !== "final"
    && options.reviewProfile !== "original"
    && options.reviewProfile !== "markup") {
    exportError("invalid_document", "input_validation", "reviewProfile is invalid.",
      "Use final, original, or markup.");
  }
  if (options.commentProfile !== "hidden"
    && options.commentProfile !== "inline"
    && options.commentProfile !== "endnotes"
    && options.commentProfile !== "margin") {
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

function browserOptions(options: NodeExportOptions): Omit<PaginatedHtmlOptions, "wasmBasePath"> {
  return {
    documentVersion: options.documentVersion,
    expectedSourceDigest: options.expectedSourceDigest,
    reviewProfile: options.reviewProfile,
    commentProfile: options.commentProfile,
    title: options.title,
    unsupportedContent: options.unsupportedContent,
    strictFonts: options.strictFonts,
    timeoutMs: options.timeoutMs,
    limits: options.limits,
  };
}

function environmentVerification(
  report: CompleteRenderReport,
  attestation: RenderEnvironmentAttestation | undefined,
): "browserObserved" | "callerAttested" {
  if (!attestation) return "browserObserved";
  const attestedFamilies = new Set(attestation.hostFonts.map((font) => font.family.toLowerCase()));
  const everyFontCovered = report.fonts.every((font) =>
    attestedFamilies.has((font.resolvedFamily ?? font.requestedFamily).toLowerCase()));
  return everyFontCovered ? "callerAttested" : "browserObserved";
}

async function renderOwned(
  sourceBytes: Uint8Array,
  options: NodeExportOptions,
  outputs: readonly RenderOutput[],
): Promise<RenderBatchResult> {
  const requested = validateOutputs(outputs);
  nodeOptionsPreflight(options);
  sourcePreflight(sourceBytes, options);
  const runtime = validateRuntime(options);
  const timeoutMs = normalizedTimeout(options);
  const deadline = Date.now() + timeoutMs;
  let report: CompleteRenderReport | undefined;
  try {
    const browser = await renderInBrowser(
      sourceBytes,
      browserOptions({ ...options, timeoutMs }),
      runtime,
      requested.includes("html"),
      requested.includes("pdf"),
      deadline,
    );
    const attestation = runtime.environmentAttestation;
    const rendererFingerprint = sha256(canonicalJson({
      schemaVersion: 1,
      browserMaterializerFingerprint: browser.materialization.rendererFingerprint,
      runtime: browser.runtime,
      environmentAttestation: attestation,
    }));
    const pageMap = {
      ...browser.materialization.pageMap,
      rendererFingerprint,
    };
    report = structuredClone(browser.materialization.renderReport);
    report.environment = {
      rendererFingerprint,
      verification: environmentVerification(report, attestation),
    };
    report.bindings.pageMapDigest = sha256(canonicalJson(pageMap));
    report.bindings.artifactRequestIds = [];

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

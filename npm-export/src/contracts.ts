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

export { COMMENT_PROFILES, REVIEW_PROFILES } from "docxodus/export-browser";

export interface FontLicenseAttestation {
  schemaVersion: 1;
  usage: "standalone-document-font-embedding";
  fileSha256: string;
  embeddingPermitted: true;
  basis: string;
  attester?: string;
}

export interface RenderEnvironmentAttestation {
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
  /** Ordered caller-owned roots searched by the verified font resolver. */
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
    this.cause = options.cause;
    this.report = options.report;
    this.committedDestinations = Object.freeze([...(options.committedDestinations ?? [])]);
  }

  toJSON(): Record<string, unknown> {
    return {
      name: this.name,
      code: this.code,
      phase: this.phase,
      message: this.message,
      remediation: this.remediation,
      ...(this.detail === undefined ? {} : { detail: this.detail }),
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
  "invalid_document",
  "conversion_failure",
  "browser_launch_failure",
  "resource_policy_failure",
  "readiness_timeout",
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
      report: value.report,
    },
  );
}

export function failedReportFromComplete(
  report: CompleteRenderReport,
  error: DocxodusExportError,
): FailedRenderReport {
  const { status: _status, environment, pages, bindings, ...base } = report;
  return {
    ...base,
    status: "failed",
    failure: {
      code: error.code,
      phase: error.phase,
      message: error.message,
      remediation: error.remediation,
    },
    environment,
    partial: { pages, bindings },
    unavailable: [
      ...(!bindings.pdfDigest
        ? [{ field: "bindings.pdfDigest" as const, reason: "PDF output did not verify." }]
        : []),
    ],
  };
}

export function attachFailedReport(
  error: unknown,
  report: CompleteRenderReport | undefined,
): unknown {
  if (!(error instanceof DocxodusExportError) || error.report || !report) return error;
  return new DocxodusExportError(error.code, error.phase, error.message, error.remediation, {
    detail: error.detail,
    cause: error.cause,
    report: failedReportFromComplete(report, error),
    committedDestinations: error.committedDestinations,
  });
}

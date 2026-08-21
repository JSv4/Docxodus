import type { Browser } from "playwright-core";
import type {
  CommentProfile,
  CompleteRenderReport,
  DocxodusExportErrorCode,
  EnvironmentVerification,
  ExportPhase,
  ExportResourceLimits,
  FailedRenderReport,
  FontResolver,
  FontResolverRequest,
  FontResolverResponse,
  PaginatedHtmlOptions,
  PaginatedHtmlResult,
  PaginatedRenderMetadata,
  ReadinessOutcome,
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
  FontResolver,
  FontResolverRequest,
  FontResolverResponse,
  PaginatedHtmlOptions,
  PaginatedHtmlResult,
  PaginatedRenderMetadata,
  ReadinessOutcome,
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
    style: "normal" | "italic" | "oblique";
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
  /** Ordered font directories searched for configured faces; order is policy. */
  fontDirectories?: readonly string[];
  fontLicenseAttestations?: readonly FontLicenseAttestation[];
  environmentAttestation?: RenderEnvironmentAttestation;
  /** Built by the host from `fontDirectories`; exposed to the page as a binding. */
  fontResolver?: FontResolver;
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
      pending: value.pending ?? value.report?.failure.pending,
      partUri: value.partUri ?? value.report?.failure.partUri,
      anchorId: value.anchorId ?? value.report?.failure.anchorId,
      resource: value.resource ?? value.report?.failure.resource,
      report: value.report,
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

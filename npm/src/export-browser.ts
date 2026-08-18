/**
 * UI-free browser materializer for deterministic, standalone paginated HTML.
 *
 * Conversion happens in the existing WASM worker. Layout happens in a private,
 * attached browsing context so no caller styles, globals, or editor state can
 * influence the artifact.
 */

import limitsContractJson from "./export-resource-limits-v1.json";
import {
  PaginationEngine,
  type PageMap,
  type PaginationDiagnostic,
} from "./pagination.js";
import {
  createWorkerDocxodus,
  type WorkerDocxodus,
} from "./worker-proxy.js";
import {
  CommentRenderMode,
  PaginationMode,
  type PackageManifest,
  type PackageManifestInspectionLimits,
  type VersionInfo,
} from "./types.js";
import {
  admitPrintVisualResources,
  awaitFinalPrintReadiness,
  documentFontReadiness,
  documentGraphicReadiness,
  documentImageReadiness,
  pageTreeReadiness,
  PrintReadinessError,
  type FontReadinessProbe,
  type PageTreeStabilityProbe,
  type VisualResourceProbe,
} from "./print-readiness.js";
import {
  automaticUrlAllowed,
  cssSecurityTokens,
  dataUrlInfo,
  standaloneSrcsetAllowed,
} from "./standalone-resource-policy.js";
import {
  BrowserFontError,
  createBrowserFontTask,
  inventoryDocumentFontRequests,
  parseCssFontFamily,
  type BrowserFontResult,
  type BrowserFontTask,
} from "./font-runtime.js";
import type {
  FontConfigurationIdentity,
  FontResolution,
  FontResolver,
} from "./font-contract.js";

export { awaitFinalPrintReadiness, PrintReadinessError } from "./print-readiness.js";
export {
  automaticUrlAllowed,
  cssSecurityTokens,
  dataUrlInfo,
} from "./standalone-resource-policy.js";
export type {
  FinalPrintReadinessResult,
  FontReadinessProbe,
  PageTreeStabilityProbe,
  PrintReadinessPhase,
  VisualResourceProbe,
} from "./print-readiness.js";
export {
  fontFamilyKey,
  FONT_RESOLVER_CONTRACT_ID,
  FONT_RESOLVER_SCHEMA_VERSION,
  FONT_SUBSTITUTION_CONTRACT,
  FONT_SUBSTITUTION_CONTRACT_MATERIAL,
  FONT_SUBSTITUTION_CONTRACT_VERSION,
  normalizeFontFamilyName,
} from "./font-contract.js";
export { inventoryDocumentFontRequests, parseCssFontFamily } from "./font-runtime.js";
export type {
  FontConfigurationIdentity,
  FontEmbeddingKind,
  FontFamilyKind,
  FontFaceMatch,
  FontFaceStyle,
  FontFileFormat,
  FontGlyphCoverage,
  FontLicenseEvidence,
  FontMediaType,
  FontRequest,
  FontResolution,
  FontResolutionSource,
  FontResolutionStatus,
  FontResolver,
  FontResolverFace,
  FontResolverOutcome,
  FontResolverRequest,
  FontResolverResponse,
  FontSubstitutionEntry,
} from "./font-contract.js";

export type ReviewProfile = "final" | "original" | "markup";
export type CommentProfile = "hidden" | "inline" | "endnotes" | "margin";
export type UnsupportedContentPolicy = "warn" | "strict";

export type ExportPhase =
  | "input_validation"
  | "package_preflight"
  | "browser_launch"
  | "wasm_initialization"
  | "docx_conversion"
  | "font_loading"
  | "image_decoding"
  | "chart_svg_materialization"
  | "pagination"
  | "running_story_placement"
  | "page_tree_stability"
  | "pdf_print"
  | "output_verification"
  | "output_write"
  | "filesystem_commit"
  | "cleanup";

export type DocxodusExportErrorCode =
  | "invalid_argument"
  | "invalid_document"
  | "source_digest_mismatch"
  | "document_version_unrepresentable"
  | "conversion_failure"
  | "browser_launch_failure"
  | "resource_policy_failure"
  | "readiness_timeout"
  | "operation_cancelled"
  | "pagination_failure"
  | "pdf_write_failure"
  | "output_write_failure"
  | "output_verification_failure"
  | "resource_limit"
  | "unsupported_runtime"
  | "filesystem_failure";

export interface ExportResourceLimits {
  compressedDocxBytes: number;
  opcEntries: number;
  expandedOpcBytes: number;
  xmlPartBytes: number;
  opcUriCharacters: number;
  opcCompressionRatio: number;
  htmlOutputBytes: number;
  pdfOutputBytes: number;
  pageMapOutputBytes: number;
  renderReportOutputBytes: number;
  pdfParserExpandedBytes: number;
  finalPages: number;
  domNodes: number;
  automaticResources: number;
  automaticResourceBytes: number;
  renderDiagnostics: number;
  fontDirectoryEntries: number;
  fontFiles: number;
  fontFileBytes: number;
  fontTotalBytes: number;
  fontRequests: number;
  fontSampleCodePoints: number;
}

interface ExportLimitsContract {
  schemaVersion: 1;
  defaults: ExportResourceLimits;
  hardCeilings: ExportResourceLimits;
  timeoutMs: { default: number; hardCeiling: number };
}

const LIMITS_CONTRACT = limitsContractJson as ExportLimitsContract;
export const DEFAULT_EXPORT_RESOURCE_LIMITS: Readonly<ExportResourceLimits> =
  Object.freeze({ ...LIMITS_CONTRACT.defaults });
export const HARD_EXPORT_RESOURCE_LIMITS: Readonly<ExportResourceLimits> =
  Object.freeze({ ...LIMITS_CONTRACT.hardCeilings });
export const DEFAULT_EXPORT_TIMEOUT_MS = LIMITS_CONTRACT.timeoutMs.default;
export const HARD_EXPORT_TIMEOUT_MS = LIMITS_CONTRACT.timeoutMs.hardCeiling;

export interface PaginatedHtmlOptions {
  documentVersion?: number;
  expectedSourceDigest?: string;
  reviewProfile: ReviewProfile;
  reviewProfileAlreadyApplied?: boolean;
  commentProfile: CommentProfile;
  title?: string;
  unsupportedContent?: UnsupportedContentPolicy;
  strictFonts?: boolean;
  /** Ephemeral browser-side resolver; never serialized into standalone output. */
  fontResolver?: FontResolver;
  timeoutMs?: number;
  limits?: Partial<ExportResourceLimits>;
  signal?: AbortSignal;
  /** Trusted runtime assets only; never a document-resource base URL. */
  wasmBasePath?: string;
}

export interface RenderWarning {
  code: string;
  severity: "warning";
  phase: ExportPhase;
  message: string;
  remediation: string;
  detail?: string;
  partUri?: string;
  anchorId?: string;
  resource?: string;
}

export interface ReadinessOutcome {
  phase: ExportPhase;
  status: "complete" | "failed" | "cancelled";
  elapsedMs: number;
  pending: string[];
  diagnostics?: PaginationDiagnostic[];
}

export interface ResourceOutcome {
  kind: "image" | "svg" | "chart" | "external_link";
  status: "embedded" | "inline" | "allowed_user_link" | "omitted";
  readiness?: "complete" | "failed";
  contentKey?: string;
  resource?: string;
  anchorId?: string;
  message?: string;
  mediaType?: string;
  byteLength?: number;
}

export interface UnsupportedContentOutcome {
  contentType: string;
  elementName?: string;
  anchorId?: string;
  action: "placeholder";
}

export interface RenderReportBase {
  schema: "https://docxodus.dev/schemas/render/render-report/v3";
  schemaVersion: 3;
  source: {
    rawPackageBytesDigest: string;
    byteLength: number;
    documentVersion: number;
  };
  derivedProfileSource?: {
    rawPackageBytesDigest: string;
    byteLength: number;
  };
  options: {
    reviewProfile: ReviewProfile;
    reviewProfileAlreadyApplied: boolean;
    commentProfile: CommentProfile;
    title: string;
    outputs: Array<"html" | "pdf">;
    layoutDigest: string;
    runtimePolicyDigest: string;
    policy: {
      unsupportedContent: UnsupportedContentPolicy;
      strictFonts: boolean;
      timeoutMs: number;
      limits: ExportResourceLimits;
    };
  };
  readiness: ReadinessOutcome[];
  fontIdentity?: FontConfigurationIdentity;
  fonts: FontResolution[];
  fontReadiness: FontReadinessProbe[];
  resources: ResourceOutcome[];
  unsupportedContent: UnsupportedContentOutcome[];
  warnings: RenderWarning[];
}

export type EnvironmentVerification = "nodeVerified" | "browserObserved" | "callerAttested";
export type FidelityTier = "releaseBaselined" | "experimental" | "unbaselined";

export interface ExportRuntimeObservedFacts {
  runtimeKind: "browser" | "nodeChromium";
  playwrightVersion?: string;
  browserProduct?: string;
  browserBuild?: string;
  executableSha256?: string;
  launchFlags?: string[];
  operatingSystem?: string;
  architecture?: string;
  locale: string;
  timezone: string;
  viewport: [number, number];
  deviceScaleFactor: number;
  media: {
    colorScheme: "light" | "dark" | "no-preference";
    reducedMotion: "reduce" | "no-preference";
    forcedColors: "active" | "none";
    printMedia: true;
  };
  networkIsolation: "ownedProcessRestricted" | "contextRestricted";
}

export interface ExportRuntimeAttestationEvidence {
  chromiumProduct: string;
  chromiumBuild: string;
  executableSha256?: string;
  launchFlags: string[];
  hostFontsDigest: string;
  basis: string;
}

export interface CompleteRenderReport extends RenderReportBase {
  status: "complete";
  fontIdentity: FontConfigurationIdentity;
  environment: {
    rendererFingerprint: string;
    verification: EnvironmentVerification;
    fidelityTier: FidelityTier;
    observed: ExportRuntimeObservedFacts;
    attested?: ExportRuntimeAttestationEvidence;
    attestationDigest?: string;
  };
  pages: Array<{
    pageNumber: number;
    pageInSection: number;
    pageName: string;
    width: number;
    height: number;
    sectionIndex?: number;
  }>;
  bindings: {
    pageMapDigest: string;
    htmlDigest?: string;
    pdfDigest?: string;
    artifactRequestIds: string[];
    pdfByteDeterministic?: false;
    volatilePdfMetadata?: Record<string, string>;
  };
}

export interface FailedRenderReport extends RenderReportBase {
  status: "failed";
  failure: {
    code: DocxodusExportErrorCode;
    severity: "error";
    phase: ExportPhase;
    message: string;
    remediation: string;
    detail?: string;
    pending?: string[];
    partUri?: string;
    anchorId?: string;
    resource?: string;
  };
  environment?: {
    rendererFingerprint?: string;
    verification: EnvironmentVerification;
    fidelityTier?: FidelityTier;
    observed?: ExportRuntimeObservedFacts;
    attested?: ExportRuntimeAttestationEvidence;
    attestationDigest?: string;
  };
  partial?: {
    pages?: CompleteRenderReport["pages"];
    bindings?: Partial<CompleteRenderReport["bindings"]>;
  };
  unavailable: Array<{
    field:
      | "environment.rendererFingerprint"
      | "bindings.pageMapDigest"
      | "bindings.htmlDigest"
      | "bindings.pdfDigest";
    reasonCode: "notReached" | "notRequested" | "failedVerification" | "discardedOnFailure";
    detail: string;
  }>;
}

export type RenderReport = CompleteRenderReport | FailedRenderReport;

export interface PaginatedRenderMetadata {
  pageCount: number;
  pageMap: PageMap;
  renderReport: CompleteRenderReport;
  warnings: RenderWarning[];
  rendererFingerprint: string;
}

export interface PaginatedHtmlResult extends PaginatedRenderMetadata {
  html: string;
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
  report?: FailedRenderReport;

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
    } = {},
  ) {
    super(message);
    this.name = "DocxodusExportError";
    this.code = code;
    this.phase = phase;
    this.remediation = remediation;
    this.detail = options.detail;
    this.pending = options.pending;
    this.partUri = options.partUri;
    this.anchorId = options.anchorId;
    this.resource = options.resource;
    this.cause = options.cause;
    this.report = options.report;
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
      ...(this.pending === undefined ? {} : { pending: [...this.pending] }),
      ...(this.partUri === undefined ? {} : { partUri: this.partUri }),
      ...(this.anchorId === undefined ? {} : { anchorId: this.anchorId }),
      ...(this.resource === undefined ? {} : { resource: this.resource }),
      ...(this.report === undefined ? {} : { report: this.report }),
    };
  }
}

interface NormalizedOptions {
  documentVersion: number;
  expectedSourceDigest?: string;
  reviewProfile: ReviewProfile;
  reviewProfileAlreadyApplied: boolean;
  commentProfile: CommentProfile;
  title: string;
  unsupportedContent: UnsupportedContentPolicy;
  strictFonts: boolean;
  fontResolver?: FontResolver;
  timeoutMs: number;
  limits: ExportResourceLimits;
  wasmBasePath: string;
  signal?: AbortSignal;
}

interface ExecutionState {
  startedAt: number;
  deadline: number;
  phase: ExportPhase;
  readiness: ReadinessOutcome[];
  warnings: RenderWarning[];
  fontIdentity?: FontConfigurationIdentity;
  fonts: FontResolution[];
  fontReadiness: FontReadinessProbe[];
  resources: ResourceOutcome[];
  unsupportedContent: UnsupportedContentOutcome[];
  limits: ExportResourceLimits;
  signal?: AbortSignal;
}

interface FinalizedTree {
  frame: HTMLIFrameElement;
  document: Document;
  engine: PaginationEngine;
  pages: HTMLElement[];
}

interface AttemptStateCheckpoint {
  readiness: number;
  warnings: number;
  fonts: number;
  fontReadiness: number;
  fontIdentity?: FontConfigurationIdentity;
  resources: number;
  unsupportedContent: number;
}

interface RuntimeAssetIdentity {
  packageVersion: string;
  graphDigest: string;
  materializerDigest: string;
  assetCount: number;
  verifiedRuntimeAssetCount: number;
}

class PageTreeInstabilityError extends Error {
  readonly pending: readonly string[];

  constructor(message: string, pending: readonly string[] = []) {
    super(message);
    this.name = "PageTreeInstabilityError";
    this.pending = Object.freeze(boundedPendingResources(pending));
  }
}

const REPORT_SCHEMA = "https://docxodus.dev/schemas/render/render-report/v3" as const;
const TEXT_ENCODER = new TextEncoder();
const ALLOWED_REVIEW_PROFILES = new Set<ReviewProfile>(["final", "original", "markup"]);
const ALLOWED_COMMENT_PROFILES = new Set<CommentProfile>(["hidden", "inline", "endnotes", "margin"]);
const ALLOWED_UNSUPPORTED_POLICIES = new Set<UnsupportedContentPolicy>(["warn", "strict"]);
const PACKAGE_LIMIT_FINDINGS = new Set([
  "entry_count_limit_exceeded",
  "entry_expansion_limit_exceeded",
  "entry_uri_limit_exceeded",
  "compression_ratio_limit_exceeded",
  "total_expansion_limit_exceeded",
  "xml_size_limit_exceeded",
]);
const RUNTIME_ASSET_GRAPH_MAX_BYTES = 1024 * 1024;
const RUNTIME_ASSET_COUNT_MAX = 10_000;
const RUNTIME_ASSET_BYTES_MAX = 64 * 1024 * 1024;
const RUNTIME_ASSET_TOTAL_BYTES_MAX = 1024 * 1024 * 1024;

function compareCodeUnits(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}
const STANDALONE_CSP = [
  "default-src 'none'",
  "base-uri 'none'",
  "connect-src 'none'",
  "font-src data:",
  "form-action 'none'",
  "frame-src 'none'",
  "img-src data:",
  "media-src data:",
  "object-src 'none'",
  "script-src 'none'",
  "style-src 'unsafe-inline'",
].join("; ");
const RENDER_CSP = `${STANDALONE_CSP}; navigate-to 'none'`;

function fail(
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
  } = {},
): never {
  throw new DocxodusExportError(code, phase, message, remediation, options);
}

function normalizeOptions(options: PaginatedHtmlOptions): NormalizedOptions {
  if (!options || typeof options !== "object") {
    fail("invalid_argument", "input_validation", "Export options are required.",
      "Supply explicit reviewProfile and commentProfile values.");
  }
  if (!ALLOWED_REVIEW_PROFILES.has(options.reviewProfile)) {
    fail("invalid_argument", "input_validation", "reviewProfile is invalid.",
      "Use final, original, or markup.");
  }
  if (!ALLOWED_COMMENT_PROFILES.has(options.commentProfile)) {
    fail("invalid_argument", "input_validation", "commentProfile is invalid.",
      "Use hidden, inline, endnotes, or margin.");
  }
  const reviewProfileAlreadyApplied = options.reviewProfileAlreadyApplied ?? false;
  if (typeof reviewProfileAlreadyApplied !== "boolean") {
    fail("invalid_argument", "input_validation", "reviewProfileAlreadyApplied must be boolean.",
      "Omit it or pass true only for exact policy-derived final/original bytes.");
  }
  if (reviewProfileAlreadyApplied && options.reviewProfile === "markup") {
    fail("invalid_argument", "input_validation",
      "reviewProfileAlreadyApplied is invalid with the markup profile.",
      "Use unchanged source bytes for markup, or choose final/original.");
  }
  const unsupportedContent = options.unsupportedContent ?? "warn";
  if (!ALLOWED_UNSUPPORTED_POLICIES.has(unsupportedContent)) {
    fail("invalid_argument", "input_validation", "unsupportedContent is invalid.",
      "Use warn or strict.");
  }
  const documentVersion = options.documentVersion ?? 0;
  if (!Number.isSafeInteger(documentVersion) || documentVersion < 0) {
    fail("document_version_unrepresentable", "input_validation",
      "documentVersion must be a non-negative JavaScript safe integer.",
      "Use a value between 0 and Number.MAX_SAFE_INTEGER.");
  }
  if (options.expectedSourceDigest !== undefined
    && !/^[0-9a-f]{64}$/.test(options.expectedSourceDigest)) {
    fail("invalid_argument", "input_validation", "expectedSourceDigest must be a lower-case SHA-256 hex digest.",
      "Supply exactly 64 lower-case hexadecimal characters.");
  }
  const title = options.title ?? "";
  if (typeof title !== "string") {
    fail("invalid_argument", "input_validation", "title must be a string.",
      "Supply a plain document title or omit it for the normative empty title.");
  }
  try {
    assertWellFormedUnicode(title);
  } catch {
    fail("invalid_argument", "input_validation", "title contains an unpaired UTF-16 surrogate.",
      "Supply well-formed Unicode text that can be encoded as strict UTF-8.");
  }
  if (options.strictFonts !== undefined && typeof options.strictFonts !== "boolean") {
    fail("invalid_argument", "input_validation", "strictFonts must be boolean.",
      "Pass true or false.");
  }
  if (options.signal !== undefined
    && (typeof options.signal !== "object"
      || typeof options.signal.addEventListener !== "function"
      || typeof options.signal.removeEventListener !== "function"
      || typeof options.signal.aborted !== "boolean")) {
    fail("invalid_argument", "input_validation", "signal must be an AbortSignal.",
      "Pass a standards-compliant AbortSignal or omit it.");
  }
  let wasmBasePath: string;
  try {
    if (options.wasmBasePath !== undefined
      && (typeof options.wasmBasePath !== "string" || options.wasmBasePath.length === 0)) {
      throw new TypeError("empty or non-string path");
    }
    const resolved = new URL(options.wasmBasePath ?? "./wasm/", import.meta.url);
    if (!new Set(["http:", "https:", "file:"]).has(resolved.protocol)) {
      throw new TypeError("unsupported runtime URL scheme");
    }
    wasmBasePath = resolved.href;
  } catch {
    fail("invalid_argument", "input_validation", "wasmBasePath must be a valid non-empty URL string.",
      "Point it only at the closed, hash-verified Docxodus runtime asset directory.");
  }
  if (options.fontResolver !== undefined && typeof options.fontResolver !== "function") {
    fail("invalid_argument", "input_validation", "fontResolver must be a function.",
      "Provide an asynchronous FontResolver callback or omit the option.");
  }

  const limits = { ...DEFAULT_EXPORT_RESOURCE_LIMITS };
  if (options.limits !== undefined
    && (!options.limits || typeof options.limits !== "object" || Array.isArray(options.limits))) {
    fail("invalid_argument", "input_validation", "limits must be an object.",
      "Supply only lower integer values from ExportResourceLimits.");
  }
  for (const [name, value] of Object.entries(options.limits ?? {})) {
    if (!(name in limits)) {
      fail("invalid_argument", "input_validation", `Unknown export limit: ${name}.`,
        "Use a key from ExportResourceLimits.");
    }
    const key = name as keyof ExportResourceLimits;
    if (!Number.isSafeInteger(value) || value <= 0) {
      fail("invalid_argument", "input_validation", `Export limit ${name} must be a positive safe integer.`,
        "Supply a positive integer no greater than the published default.");
    }
    if (value > DEFAULT_EXPORT_RESOURCE_LIMITS[key]) {
      fail("invalid_argument", "input_validation", `Export limit ${name} may only lower the default.`,
        `Use ${DEFAULT_EXPORT_RESOURCE_LIMITS[key]} or less.`);
    }
    limits[key] = value;
  }

  const timeoutMs = options.timeoutMs ?? LIMITS_CONTRACT.timeoutMs.default;
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0
    || timeoutMs > LIMITS_CONTRACT.timeoutMs.hardCeiling) {
    fail("invalid_argument", "input_validation", "timeoutMs is outside the supported range.",
      `Use an integer from 1 through ${LIMITS_CONTRACT.timeoutMs.hardCeiling}.`);
  }

  return Object.freeze({
    documentVersion,
    expectedSourceDigest: options.expectedSourceDigest,
    reviewProfile: options.reviewProfile,
    reviewProfileAlreadyApplied,
    commentProfile: options.commentProfile,
    title,
    unsupportedContent,
    strictFonts: options.strictFonts ?? false,
    ...(options.fontResolver ? { fontResolver: options.fontResolver } : {}),
    timeoutMs,
    limits: Object.freeze(limits),
    wasmBasePath,
    signal: options.signal,
  });
}

async function ownedBytes(
  document: File | Uint8Array,
  maximum: number,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  if (signal?.aborted) {
    fail("operation_cancelled", "input_validation", "Export was cancelled before input snapshotting.",
      "Retry with a non-aborted signal.");
  }
  if (document instanceof Uint8Array) {
    enforceLimit(document.byteLength, maximum, "compressedDocxBytes", "input_validation");
    return new Uint8Array(document);
  }
  if (typeof File !== "undefined" && document instanceof File) {
    enforceLimit(document.size, maximum, "compressedDocxBytes", "input_validation");
    const bytes = new Uint8Array(await document.arrayBuffer());
    if (signal?.aborted) {
      fail("operation_cancelled", "input_validation", "Export was cancelled while reading the input File.",
        "Retry with a non-aborted signal.");
    }
    enforceLimit(bytes.byteLength, maximum, "compressedDocxBytes", "input_validation");
    return bytes;
  }
  fail("invalid_argument", "input_validation", "document must be a File or Uint8Array.",
    "Pass immutable DOCX bytes or a browser File.");
}

function monotonicNow(): number {
  return globalThis.performance?.now() ?? Date.now();
}

const PENDING_RESOURCE_DETAILS_MAX = 64;
const PENDING_RESOURCE_LABEL_MAX = 512;

function boundedPendingResources(resources: readonly string[]): string[] {
  const bounded = resources.slice(0, PENDING_RESOURCE_DETAILS_MAX).map((resource, index) => {
    const normalized = resource.replace(/[\u0000-\u001f\u007f]+/g, " ").trim()
      || `resource-${index + 1}`;
    return normalized.length <= PENDING_RESOURCE_LABEL_MAX
      ? normalized
      : `${normalized.slice(0, PENDING_RESOURCE_LABEL_MAX - 3)}...`;
  });
  if (resources.length > PENDING_RESOURCE_DETAILS_MAX) {
    bounded.push(`... ${resources.length - PENDING_RESOURCE_DETAILS_MAX} more`);
  }
  return bounded;
}

async function runPhase<T>(
  state: ExecutionState,
  phase: ExportPhase,
  pendingResources: string[] | (() => string[]),
  operation: (signal: AbortSignal) => T | Promise<T>,
): Promise<T> {
  state.phase = phase;
  const started = monotonicNow();
  const pending = (): string[] => {
    try {
      return boundedPendingResources(
        typeof pendingResources === "function" ? pendingResources() : pendingResources,
      );
    } catch {
      return ["pending-resource-inventory"];
    }
  };
  const reportProgress = (
    status: "pending" | "complete" | "failed",
    resources: string[],
    reportedPhase: ExportPhase = phase,
  ): void => {
    const reporter = (globalThis as typeof globalThis & {
      __docxodusReadinessProgress?: (snapshot: {
        phase: ExportPhase;
        status: "pending" | "complete" | "failed";
        pending: string[];
      }) => unknown;
    }).__docxodusReadinessProgress;
    if (typeof reporter === "function") {
      try {
        void reporter({ phase: reportedPhase, status, pending: [...resources] });
      } catch {
        // Progress is diagnostic-only; the readiness result remains authoritative.
      }
    }
  };
  const initialPending = pending();
  if (state.signal?.aborted) {
    fail("operation_cancelled", phase, `Export was cancelled during ${phase}.`,
      "Retry with a non-aborted signal.", { pending: initialPending });
  }
  const remaining = state.deadline - monotonicNow();
  if (remaining <= 0) {
    fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
      "Increase timeoutMs or remove the pending resource.", {
        detail: initialPending.join(", "),
        pending: initialPending,
      });
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abortListener: (() => void) | undefined;
  let progressTimer: ReturnType<typeof setInterval> | undefined;
  let timedOutPending: string[] | undefined;
  const controller = new AbortController();
  let pendingSignature = JSON.stringify(initialPending);
  reportProgress("pending", initialPending);
  if (typeof pendingResources === "function") {
    progressTimer = setInterval(() => {
      const resources = pending();
      const signature = JSON.stringify(resources);
      if (signature === pendingSignature) return;
      pendingSignature = signature;
      reportProgress("pending", resources);
    }, 25);
  }
  try {
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => {
        const resources = pending();
        timedOutPending = resources;
        controller.abort();
        reject(new DocxodusExportError(
          "readiness_timeout",
          phase,
          `Export timed out during ${phase}.`,
          "Increase timeoutMs or remove the pending resource.",
          { detail: resources.join(", "), pending: resources },
        ));
      }, remaining);
    });
    const cancellation = new Promise<never>((_, reject) => {
      if (!state.signal) return;
      abortListener = () => {
        const resources = pending();
        controller.abort();
        reject(new DocxodusExportError(
          "operation_cancelled",
          phase,
          `Export was cancelled during ${phase}.`,
          "Retry with a non-aborted signal.",
          { pending: resources },
        ));
      };
      state.signal.addEventListener("abort", abortListener, { once: true });
    });
    const result = await Promise.race([
      Promise.resolve().then(() => operation(controller.signal)),
      timeout,
      cancellation,
    ]);
    // A synchronous DOM operation cannot be pre-empted by the timer because it
    // blocks the event loop. Reject it immediately after control returns; hot
    // pagination loops also invoke the cooperative checkpoint below.
    if (state.signal?.aborted) {
      fail("operation_cancelled", phase, `Export was cancelled during ${phase}.`,
        "Retry with a non-aborted signal.", { pending: pending() });
    }
    if (monotonicNow() >= state.deadline) {
      const resources = pending();
      fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
        "Increase timeoutMs or remove the pending resource.", {
          detail: resources.join(", "),
          pending: resources,
        });
    }
    state.readiness.push({
      phase,
      status: "complete",
      elapsedMs: Math.max(0, monotonicNow() - started),
      pending: [],
    });
    reportProgress("complete", []);
    return result;
  } catch (error) {
    const resources = timedOutPending
      ?? (error instanceof PrintReadinessError ? [...error.pending] : pending());
    const resolvedError = error instanceof PrintReadinessError
      ? new DocxodusExportError(
        error.reason === "resource_limit"
          ? "resource_limit"
          : error.message.includes("timed out")
            ? "readiness_timeout"
            : "output_verification_failure",
        error.phase,
        error.message,
        error.reason === "resource_limit"
          ? "Use a smaller document or raise the corresponding versioned readiness ceiling."
          : "Remove the unstable resource or retry in the same verified browser environment.",
        { pending: [...error.pending], cause: error },
      )
      : error;
    const failurePhase = resolvedError instanceof DocxodusExportError
      ? resolvedError.phase
      : phase;
    state.readiness.push({
      phase: failurePhase,
      status: resolvedError instanceof DocxodusExportError && resolvedError.code === "operation_cancelled"
        ? "cancelled"
        : "failed",
      elapsedMs: Math.max(0, monotonicNow() - started),
      pending: resources,
    });
    reportProgress("failed", resources, failurePhase);
    throw resolvedError;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (abortListener && state.signal) state.signal.removeEventListener("abort", abortListener);
    if (progressTimer !== undefined) clearInterval(progressTimer);
    controller.abort();
  }
}

function utf8Bytes(value: string): Uint8Array {
  return TEXT_ENCODER.encode(value);
}

function utf8ByteLength(value: string): number {
  let length = 0;
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit < 0x80) length++;
    else if (unit < 0x800) length += 2;
    else if (unit >= 0xd800 && unit <= 0xdbff
      && value.charCodeAt(index + 1) >= 0xdc00 && value.charCodeAt(index + 1) <= 0xdfff) {
      length += 4;
      index++;
    } else length += 3;
  }
  return length;
}

function preflightConvertedHtml(source: string, options: NormalizedOptions): void {
  enforceLimit(utf8ByteLength(source), options.limits.htmlOutputBytes,
    "htmlOutputBytes", "docx_conversion");
  let prospectiveNodes = 2;
  let inText = false;
  for (let index = 0; index < source.length; index++) {
    if (source[index] !== "<") {
      if (!inText && !/\s/.test(source[index])) {
        prospectiveNodes++;
        inText = true;
      }
      continue;
    }
    inText = false;
    const next = source[index + 1];
    if (next && next !== "/" && next !== "!" && next !== "?") prospectiveNodes++;
    if (prospectiveNodes > options.limits.domNodes) {
      fail("resource_limit", "docx_conversion",
        `domNodes limit exceeded before HTML attachment (${prospectiveNodes} > ${options.limits.domNodes}).`,
        "Use a smaller document or a lower-complexity conversion profile.");
    }
  }
}

function countDomNodes(document: Document, maximum: number, phase: ExportPhase): number {
  const walker = document.createTreeWalker(document, 0xffffffff);
  let count = 0;
  while (walker.nextNode()) {
    count++;
    if (count > maximum) {
      fail("resource_limit", phase, `domNodes limit exceeded (${count} > ${maximum}).`,
        "Use a smaller document or a lower-complexity conversion profile.");
    }
  }
  return count;
}

async function sha256(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) {
    fail("unsupported_runtime", "output_verification", "Web Crypto SHA-256 is unavailable.",
      "Run the exporter in a secure, standards-compliant browser context.");
  }
  const owned = new Uint8Array(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", owned.buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("");
}

async function boundedResponseBytes(
  response: Response,
  maximum: number,
  label: string,
): Promise<Uint8Array> {
  const declared = response.headers.get("content-length");
  if (declared !== null && /^(?:0|[1-9]\d*)$/.test(declared)
    && BigInt(declared) > BigInt(maximum)) {
    fail("unsupported_runtime", "wasm_initialization",
      `${label} exceeds its admitted byte length.`,
      "Deploy the closed, bounded runtime asset graph generated with this package.");
  }
  if (!response.body) {
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (bytes.byteLength > maximum) {
      fail("unsupported_runtime", "wasm_initialization",
        `${label} exceeds its admitted byte length.`,
        "Deploy the closed, bounded runtime asset graph generated with this package.");
    }
    return bytes;
  }
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    while (true) {
      const next = await reader.read();
      if (next.done) break;
      total += next.value.byteLength;
      if (total > maximum) {
        await reader.cancel();
        fail("unsupported_runtime", "wasm_initialization",
          `${label} exceeds its admitted byte length.`,
          "Deploy the closed, bounded runtime asset graph generated with this package.");
      }
      chunks.push(next.value);
    }
  } finally {
    reader.releaseLock();
  }
  const bytes = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return bytes;
}

async function loadRuntimeAssetIdentity(
  wasmBasePath: string,
  signal: AbortSignal,
): Promise<RuntimeAssetIdentity> {
  const manifestUrl = new URL("./export-assets.json", import.meta.url);
  const response = await globalThis.fetch(manifestUrl, {
    cache: "no-store",
    credentials: "same-origin",
    signal,
  });
  if (!response.ok) {
    fail("unsupported_runtime", "wasm_initialization",
      `The runtime asset graph could not be loaded (${response.status}).`,
      "Deploy export-assets.json beside the browser export bundle.");
  }
  const graphBytes = await boundedResponseBytes(
    response,
    RUNTIME_ASSET_GRAPH_MAX_BYTES,
    "Runtime asset graph",
  );
  let graphText: string;
  try {
    graphText = new TextDecoder("utf-8", { fatal: true }).decode(graphBytes);
  } catch {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph is not strict UTF-8.",
      "Regenerate export-assets.json without malformed byte sequences.");
  }
  const manifest = strictJsonParse(graphText, (detail) => fail(
    "unsupported_runtime",
    "wasm_initialization",
    "The runtime asset graph is malformed.",
    "Deploy the versioned export-assets.json generated with this bundle.",
    { detail },
  )) as Record<string, unknown>;
  if (!manifest || typeof manifest !== "object" || Array.isArray(manifest)
    || Object.keys(manifest).some((key) => !["schema", "schemaVersion", "packageVersion", "assets"].includes(key))
    || manifest.schema !== "https://docxodus.dev/schemas/export/export-assets/v1"
    || manifest.schemaVersion !== 1 || typeof manifest.packageVersion !== "string"
    || manifest.packageVersion.length === 0
    || !Array.isArray(manifest.assets) || manifest.assets.length === 0
    || manifest.assets.length > RUNTIME_ASSET_COUNT_MAX) {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph is malformed.",
      "Deploy the versioned export-assets.json generated with this bundle.");
  }
  try {
    assertWellFormedUnicode(manifest.packageVersion);
  } catch {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph packageVersion is not well-formed Unicode.",
      "Regenerate export-assets.json from valid package metadata.");
  }
  let aggregateAssetBytes = 0;
  const paths = new Set<string>();
  const assets = manifest.assets.map((entry, index) => {
    if (!entry || typeof entry !== "object") {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset entry ${index} is malformed.`,
        "Regenerate the package runtime asset graph.");
    }
    const record = entry as Record<string, unknown>;
    if (Object.keys(record).some((key) => !["path", "mediaType", "byteLength", "sha256"].includes(key))
      || typeof record.path !== "string" || typeof record.mediaType !== "string"
      || !Number.isSafeInteger(record.byteLength) || (record.byteLength as number) < 0
      || (record.byteLength as number) > RUNTIME_ASSET_BYTES_MAX
      || typeof record.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(record.sha256)) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset entry ${index} has invalid identity fields.`,
        "Regenerate the package runtime asset graph.");
    }
    const path = record.path;
    const segments = path.split("/");
    const extension = path.slice(path.lastIndexOf("."));
    const expectedMediaType: Record<string, string> = {
      ".css": "text/css",
      ".dat": "application/octet-stream",
      ".js": "text/javascript",
      ".json": "application/json",
      ".wasm": "application/wasm",
    };
    if (!path.startsWith("./") || path.includes("\\") || path.includes("?") || path.includes("#")
      || segments.some((segment, segmentIndex) =>
        (segmentIndex > 0 && (segment === "" || segment === "." || segment === "..")))
      || paths.has(path) || expectedMediaType[extension] !== record.mediaType) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset entry ${index} has a non-canonical or duplicate path.`,
        "Regenerate the package runtime asset graph.");
    }
    paths.add(path);
    aggregateAssetBytes += record.byteLength as number;
    if (aggregateAssetBytes > RUNTIME_ASSET_TOTAL_BYTES_MAX) {
      fail("unsupported_runtime", "wasm_initialization",
        "The runtime asset graph exceeds its aggregate byte ceiling.",
        "Deploy a bounded Docxodus runtime package.");
    }
    return {
      path,
      mediaType: record.mediaType as string,
      byteLength: record.byteLength as number,
      sha256: record.sha256 as string,
    };
  });
  for (let index = 1; index < assets.length; index++) {
    if (compareCodeUnits(assets[index - 1].path, assets[index].path) >= 0) {
      fail("unsupported_runtime", "wasm_initialization",
        "Runtime asset entries are not in canonical path order.",
        "Regenerate export-assets.json with the package asset generator.");
    }
  }
  const materializer = assets.find((entry) => entry.path === "./export-browser.bundle.js");
  if (!materializer || !paths.has("./docxodus.worker.js")
    || !paths.has("./wasm/_framework/dotnet.js")) {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph does not identify the browser materializer.",
      "Regenerate export-assets.json from the complete runtime package.");
  }

  const materializerResponse = await globalThis.fetch(import.meta.url, {
    cache: "no-store",
    credentials: "same-origin",
    signal,
  });
  if (!materializerResponse.ok) {
    fail("unsupported_runtime", "wasm_initialization",
      `The loaded browser materializer could not be verified (${materializerResponse.status}).`,
      "Serve the browser export bundle from a readable same-origin URL.");
  }
  const materializerBytes = await boundedResponseBytes(
    materializerResponse,
    materializer.byteLength,
    "Browser materializer",
  );
  const materializerDigest = await sha256(materializerBytes);
  if (materializerDigest !== materializer.sha256) {
    fail("unsupported_runtime", "wasm_initialization",
      "The loaded browser materializer does not match export-assets.json.",
      "Deploy the bundle and asset graph from the same Docxodus build.");
  }

  const verifiedRuntimeAssets = assets.filter((entry) =>
    entry.path === "./docxodus.worker.js" || entry.path.startsWith("./wasm/_framework/"));
  const resolvedWasmBasePath = new URL(
    wasmBasePath.endsWith("/") ? wasmBasePath : `${wasmBasePath}/`,
    import.meta.url,
  );
  for (const asset of verifiedRuntimeAssets) {
    const url = asset.path === "./docxodus.worker.js"
      ? new URL("./docxodus.worker.js", import.meta.url)
      : new URL(asset.path.slice("./wasm/".length), resolvedWasmBasePath);
    const assetResponse = await globalThis.fetch(url, {
      cache: "force-cache",
      credentials: "same-origin",
      signal,
    });
    if (!assetResponse.ok) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset ${asset.path} could not be verified (${assetResponse.status}).`,
        "Deploy every runtime asset named by export-assets.json.");
    }
    const bytes = await boundedResponseBytes(assetResponse, asset.byteLength, `Runtime asset ${asset.path}`);
    if (bytes.byteLength !== asset.byteLength || await sha256(bytes) !== asset.sha256) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset ${asset.path} does not match export-assets.json.`,
        "Deploy the worker, WASM directory, browser bundle, and asset graph from one build.");
    }
  }

  return {
    packageVersion: manifest.packageVersion,
    graphDigest: await sha256(utf8Bytes(canonicalJson({
      schemaVersion: manifest.schemaVersion,
      packageVersion: manifest.packageVersion,
      assets,
    }))),
    materializerDigest,
    assetCount: assets.length,
    verifiedRuntimeAssetCount: verifiedRuntimeAssets.length,
  };
}

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "string") {
    assertWellFormedUnicode(value);
    return value;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("Canonical JSON does not support non-finite numbers");
    return Object.is(value, -0) ? 0 : value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (typeof value === "object") {
    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new TypeError("Canonical JSON supports only plain objects");
    }
    const result: Record<string, unknown> = {};
    for (const key of Object.keys(value as Record<string, unknown>).sort()) {
      assertWellFormedUnicode(key);
      const member = (value as Record<string, unknown>)[key];
      if (member !== undefined) result[key] = canonicalValue(member);
    }
    return result;
  }
  throw new TypeError(`Canonical JSON does not support ${typeof value}`);
}

function assertWellFormedUnicode(value: string): void {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        throw new TypeError("Canonical JSON does not support unpaired UTF-16 surrogates");
      }
      index++;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      throw new TypeError("Canonical JSON does not support unpaired UTF-16 surrogates");
    }
  }
}

/** Serialize report/schema values with recursively sorted object keys and no insignificant space. */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value));
}

async function canonicalMaterialDigest(domain: string, value: unknown): Promise<string> {
  if (!/^[\x20-\x7e]+$/.test(domain)) {
    throw new TypeError("Canonical digest domain tags must be printable ASCII");
  }
  const domainBytes = utf8Bytes(domain);
  const materialBytes = utf8Bytes(canonicalJson(value));
  const input = new Uint8Array(domainBytes.byteLength + 1 + materialBytes.byteLength);
  input.set(domainBytes, 0);
  input[domainBytes.byteLength] = 0;
  input.set(materialBytes, domainBytes.byteLength + 1);
  return sha256(input);
}

function constantTimeDigestEqual(left: string, right: string): boolean {
  let difference = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index++) {
    difference |= (left.charCodeAt(index) || 0) ^ (right.charCodeAt(index) || 0);
  }
  return difference === 0;
}

type JsonRecord = Record<string, unknown>;

function manifestFailure(path: string, detail: string): never {
  fail("invalid_document", "package_preflight",
    `DOCX preflight returned a manifest that violates schema v1 at ${path}.`,
    "Use the matching hardened #493 package-manifest producer and consumer.",
    { detail });
}

function strictJsonParse(
  source: string,
  onFailure: (detail: string) => never = (detail) => manifestFailure("$", detail),
): unknown {
  let cursor = 0;
  const whitespace = () => {
    while (cursor < source.length && /[\u0009\u000a\u000d\u0020]/.test(source[cursor])) cursor++;
  };
  const parseStringToken = (): string => {
    const start = cursor;
    if (source[cursor++] !== '"') throw new SyntaxError(`Expected string at ${start}`);
    while (cursor < source.length) {
      const character = source[cursor++];
      if (character === '"') return JSON.parse(source.slice(start, cursor)) as string;
      if (character.charCodeAt(0) < 0x20) throw new SyntaxError(`Control character at ${cursor - 1}`);
      if (character !== "\\") continue;
      if (cursor >= source.length) throw new SyntaxError("Unterminated JSON escape");
      const escape = source[cursor++];
      if (escape === "u") {
        if (!/^[0-9a-fA-F]{4}$/.test(source.slice(cursor, cursor + 4))) {
          throw new SyntaxError(`Invalid Unicode escape at ${cursor - 2}`);
        }
        cursor += 4;
      } else if (!'"\\/bfnrt'.includes(escape)) {
        throw new SyntaxError(`Invalid JSON escape at ${cursor - 2}`);
      }
    }
    throw new SyntaxError("Unterminated JSON string");
  };
  const parseValue = (depth: number): void => {
    if (depth > 128) throw new SyntaxError("JSON nesting is too deep");
    whitespace();
    const character = source[cursor];
    if (character === '"') {
      parseStringToken();
      return;
    }
    if (character === "{") {
      cursor++;
      whitespace();
      const keys = new Set<string>();
      if (source[cursor] === "}") { cursor++; return; }
      while (cursor < source.length) {
        whitespace();
        const key = parseStringToken();
        if (keys.has(key)) throw new SyntaxError(`Duplicate JSON property ${JSON.stringify(key)}`);
        keys.add(key);
        whitespace();
        if (source[cursor++] !== ":") throw new SyntaxError(`Expected colon at ${cursor - 1}`);
        parseValue(depth + 1);
        whitespace();
        const separator = source[cursor++];
        if (separator === "}") return;
        if (separator !== ",") throw new SyntaxError(`Expected object separator at ${cursor - 1}`);
      }
      throw new SyntaxError("Unterminated JSON object");
    }
    if (character === "[") {
      cursor++;
      whitespace();
      if (source[cursor] === "]") { cursor++; return; }
      while (cursor < source.length) {
        parseValue(depth + 1);
        whitespace();
        const separator = source[cursor++];
        if (separator === "]") return;
        if (separator !== ",") throw new SyntaxError(`Expected array separator at ${cursor - 1}`);
      }
      throw new SyntaxError("Unterminated JSON array");
    }
    const rest = source.slice(cursor);
    const literal = /^(?:true|false|null)/.exec(rest)?.[0]
      ?? /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(rest)?.[0];
    if (!literal) throw new SyntaxError(`Invalid JSON value at ${cursor}`);
    cursor += literal.length;
  };
  try {
    parseValue(0);
    whitespace();
    if (cursor !== source.length) throw new SyntaxError(`Trailing JSON data at ${cursor}`);
    return JSON.parse(source) as unknown;
  } catch (error) {
    return onFailure(error instanceof Error ? error.message : String(error));
  }
}

function recordAt(
  value: unknown,
  path: string,
  required: readonly string[],
  optional: readonly string[] = [],
): JsonRecord {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    manifestFailure(path, "expected an object");
  }
  const record = value as JsonRecord;
  const allowed = new Set([...required, ...optional]);
  for (const key of Object.keys(record)) {
    if (!allowed.has(key)) manifestFailure(`${path}.${key}`, "unknown property");
  }
  for (const key of required) {
    if (!Object.prototype.hasOwnProperty.call(record, key)) {
      manifestFailure(`${path}.${key}`, "missing required property");
    }
  }
  return record;
}

function arrayAt(value: unknown, path: string): unknown[] {
  if (!Array.isArray(value)) manifestFailure(path, "expected an array");
  return value;
}

function stringAt(value: unknown, path: string, allowEmpty = true): string {
  if (typeof value !== "string" || (!allowEmpty && value.length === 0)) {
    manifestFailure(path, "expected a string");
  }
  try {
    assertWellFormedUnicode(value);
  } catch (error) {
    manifestFailure(path, error instanceof Error ? error.message : String(error));
  }
  return value;
}

function integerAt(value: unknown, path: string, minimum = 0): number {
  if (!Number.isSafeInteger(value) || (value as number) < minimum) {
    manifestFailure(path, `expected a safe integer >= ${minimum}`);
  }
  return value as number;
}

function booleanAt(value: unknown, path: string): boolean {
  if (typeof value !== "boolean") manifestFailure(path, "expected a boolean");
  return value;
}

function enumAt<T extends string>(value: unknown, path: string, allowed: readonly T[]): T {
  if (typeof value !== "string" || !allowed.includes(value as T)) {
    manifestFailure(path, `expected one of ${allowed.join(", ")}`);
  }
  return value as T;
}

function digestAt(value: unknown, path: string, nullable: boolean): void {
  if (value === null && nullable) return;
  const digest = recordAt(value, path, ["algorithm", "value"]);
  if (digest.algorithm !== "SHA-256") manifestFailure(`${path}.algorithm`, "expected SHA-256");
  if (typeof digest.value !== "string" || !/^[0-9a-f]{64}$/.test(digest.value)) {
    manifestFailure(`${path}.value`, "expected 64 lower-case hexadecimal digits");
  }
}

function decimalAt(value: unknown, path: string): bigint {
  if (typeof value !== "string" || !/^(?:0|[1-9]\d*)$/.test(value)) {
    manifestFailure(path, "expected a canonical non-negative base-10 integer string");
  }
  return BigInt(value);
}

function validatePackageManifestJson(source: string): PackageManifest {
  const value = strictJsonParse(source);
  const manifest = recordAt(value, "$", [
    "schema", "schemaVersion", "packageKind", "isValid", "rawPackageBytesDigest",
    "orderedOpcContentDigest", "normalizedSemanticDigest", "entries", "contentTypes",
    "relationships", "facts", "findings",
  ]);
  if (manifest.schema !== "https://docxodus.dev/schemas/verification/package-manifest/v1") {
    manifestFailure("$.schema", "unexpected package-manifest discriminator");
  }
  if (manifest.schemaVersion !== 1) manifestFailure("$.schemaVersion", "expected 1");
  enumAt(manifest.packageKind, "$.packageKind", [
    "opc", "zip", "zip-encrypted", "ole-encrypted", "ole", "malformed",
  ] as const);
  booleanAt(manifest.isValid, "$.isValid");
  digestAt(manifest.rawPackageBytesDigest, "$.rawPackageBytesDigest", false);
  digestAt(manifest.orderedOpcContentDigest, "$.orderedOpcContentDigest", true);
  digestAt(manifest.normalizedSemanticDigest, "$.normalizedSemanticDigest", true);

  for (const [index, item] of arrayAt(manifest.entries, "$.entries").entries()) {
    const path = `$.entries[${index}]`;
    const entry = recordAt(item, path, [
      "uri", "occurrence", "contentType", "contentTypeSource", "size", "compressedSize",
      "rawBytesDigest", "normalizedXmlDigest", "isXml", "isEncrypted",
    ]);
    stringAt(entry.uri, `${path}.uri`, false);
    integerAt(entry.occurrence, `${path}.occurrence`);
    if (entry.contentType !== null) stringAt(entry.contentType, `${path}.contentType`, false);
    enumAt(entry.contentTypeSource, `${path}.contentTypeSource`, [
      "override", "default", "implicit", "unresolved",
    ] as const);
    decimalAt(entry.size, `${path}.size`);
    decimalAt(entry.compressedSize, `${path}.compressedSize`);
    digestAt(entry.rawBytesDigest, `${path}.rawBytesDigest`, true);
    digestAt(entry.normalizedXmlDigest, `${path}.normalizedXmlDigest`, true);
    booleanAt(entry.isXml, `${path}.isXml`);
    if (entry.isEncrypted !== null) booleanAt(entry.isEncrypted, `${path}.isEncrypted`);
  }
  for (const [index, item] of arrayAt(manifest.contentTypes, "$.contentTypes").entries()) {
    const path = `$.contentTypes[${index}]`;
    const declaration = recordAt(item, path, ["kind", "key", "contentType", "occurrence"]);
    enumAt(declaration.kind, `${path}.kind`, ["default", "override"] as const);
    stringAt(declaration.key, `${path}.key`, false);
    stringAt(declaration.contentType, `${path}.contentType`, false);
    integerAt(declaration.occurrence, `${path}.occurrence`);
  }
  for (const [index, item] of arrayAt(manifest.relationships, "$.relationships").entries()) {
    const path = `$.relationships[${index}]`;
    const relationship = recordAt(item, path, [
      "ownerUri", "id", "type", "target", "targetMode", "resolvedTargetUri", "isTargetPresent",
    ]);
    stringAt(relationship.ownerUri, `${path}.ownerUri`, false);
    stringAt(relationship.id, `${path}.id`, false);
    stringAt(relationship.type, `${path}.type`, false);
    stringAt(relationship.target, `${path}.target`, false);
    enumAt(relationship.targetMode, `${path}.targetMode`, ["Internal", "External"] as const);
    if (relationship.resolvedTargetUri !== null) {
      stringAt(relationship.resolvedTargetUri, `${path}.resolvedTargetUri`, false);
    }
    if (relationship.isTargetPresent !== null) {
      booleanAt(relationship.isTargetPresent, `${path}.isTargetPresent`);
    }
  }

  const facts = recordAt(manifest.facts, "$.facts", [
    "mainDocumentUri", "isStrictOoxml", "isMacroEnabled", "hasCoreProperties",
    "hasExtendedProperties", "hasCustomProperties", "sectionCount", "paragraphCount", "tableCount",
    "headerPartCount", "footerPartCount", "footnoteCount", "endnoteCount", "styleCount",
    "numberingDefinitionCount", "themePartCount", "mediaPartCount", "customXmlPartCount",
    "drawingCount", "altChunkCount", "fieldCount", "revisions", "annotations",
  ]);
  if (facts.mainDocumentUri !== null) stringAt(facts.mainDocumentUri, "$.facts.mainDocumentUri", false);
  for (const name of [
    "isStrictOoxml", "isMacroEnabled", "hasCoreProperties", "hasExtendedProperties",
    "hasCustomProperties",
  ]) booleanAt(facts[name], `$.facts.${name}`);
  for (const name of [
    "sectionCount", "paragraphCount", "tableCount", "headerPartCount", "footerPartCount",
    "footnoteCount", "endnoteCount", "styleCount", "numberingDefinitionCount", "themePartCount",
    "mediaPartCount", "customXmlPartCount", "drawingCount", "altChunkCount", "fieldCount",
  ]) integerAt(facts[name], `$.facts.${name}`);
  const revisions = recordAt(facts.revisions, "$.facts.revisions", [
    "insertions", "deletions", "moveFrom", "moveTo", "propertyChanges", "structuralChanges",
    "otherChanges", "total",
  ]);
  for (const name of Object.keys(revisions)) integerAt(revisions[name], `$.facts.revisions.${name}`);
  const annotations = recordAt(facts.annotations, "$.facts.annotations", [
    "comments", "commentReplies", "threadedCommentMetadata", "resolvedComments", "people",
    "docxodusAnnotations",
  ]);
  for (const name of Object.keys(annotations)) integerAt(annotations[name], `$.facts.annotations.${name}`);

  for (const [index, item] of arrayAt(manifest.findings, "$.findings").entries()) {
    const path = `$.findings[${index}]`;
    const finding = recordAt(item, path, ["code", "severity", "message", "location"]);
    stringAt(finding.code, `${path}.code`, false);
    enumAt(finding.severity, `${path}.severity`, ["info", "warning", "error"] as const);
    stringAt(finding.message, `${path}.message`, false);
    if (finding.location !== null) {
      const location = recordAt(finding.location, `${path}.location`, [
        "entryUri", "ownerUri", "relationshipId", "targetUri", "propertyPath",
      ]);
      for (const name of Object.keys(location)) {
        if (location[name] !== null) stringAt(location[name], `${path}.location.${name}`);
      }
    }
  }
  return manifest as unknown as PackageManifest;
}

function addWarning(state: ExecutionState, warning: RenderWarning): void {
  enforceDiagnosticAdmission(state, 1);
  state.warnings.push(warning);
}

function addResource(state: ExecutionState, resource: ResourceOutcome): void {
  enforceDiagnosticAdmission(state, 1);
  state.resources.push(resource);
}

function addUnsupportedContent(state: ExecutionState, outcome: UnsupportedContentOutcome): void {
  enforceDiagnosticAdmission(state, 1);
  state.unsupportedContent.push(outcome);
}

function enforceDiagnosticAdmission(state: ExecutionState, additional: number): void {
  const current = state.warnings.length + state.resources.length
    + state.unsupportedContent.length + state.fonts.length + state.fontReadiness.length;
  if (current + additional > state.limits.renderDiagnostics) {
    fail("resource_limit", state.phase,
      `renderDiagnostics limit exceeded (${current + additional} > ${state.limits.renderDiagnostics}).`,
      "Use a smaller document or a versioned deployment policy with a higher diagnostic ceiling.");
  }
}

function addPageTreeRetryWarning(state: ExecutionState): void {
  addWarning(state, {
    code: "page_tree_retry",
    severity: "warning",
    phase: "page_tree_stability",
    message: "The finalized page tree changed during its quiet interval and was rebuilt from pristine converted HTML.",
    remediation: "No action is required unless repeated exports fail page-tree stability.",
  });
}

function enforceLimit(actual: number, maximum: number, name: keyof ExportResourceLimits, phase: ExportPhase): void {
  if (actual > maximum) {
    fail("resource_limit", phase, `${name} limit exceeded (${actual} > ${maximum}).`,
      `Use a smaller document or raise the deployment ceiling in a versioned limits contract.`);
  }
}

function inspectionLimits(options: NormalizedOptions): PackageManifestInspectionLimits {
  return {
    opcEntries: options.limits.opcEntries,
    expandedOpcBytes: options.limits.expandedOpcBytes,
    xmlPartBytes: options.limits.xmlPartBytes,
    opcUriCharacters: options.limits.opcUriCharacters,
    opcCompressionRatio: options.limits.opcCompressionRatio,
  };
}

async function preflightManifest(
  manifest: PackageManifest,
  bytes: Uint8Array,
  options: NormalizedOptions,
  state: ExecutionState,
  compareExpectedDigest: boolean,
): Promise<void> {
  const sourceDigest = manifest.rawPackageBytesDigest.value;
  const recomputedDigest = await sha256(bytes);
  if (!constantTimeDigestEqual(sourceDigest, recomputedDigest)) {
    fail("invalid_document", "package_preflight",
      "The #493 source digest does not match the exact bytes supplied to export.",
      "Use a matching package-manifest producer and immutable source snapshot.", {
        detail: `manifest=${sourceDigest}; recomputed=${recomputedDigest}`,
      });
  }
  if (compareExpectedDigest && options.expectedSourceDigest
    && !constantTimeDigestEqual(options.expectedSourceDigest, sourceDigest)) {
    fail("source_digest_mismatch", "package_preflight", "The source digest does not match expectedSourceDigest.",
      "Render the exact verified source bytes or update the expected digest.", {
        detail: `expected=${options.expectedSourceDigest}; actual=${sourceDigest}`,
      });
  }
  for (const finding of manifest.findings) {
    if (finding.severity === "info") continue;
    addWarning(state, {
      code: finding.code,
      severity: "warning",
      phase: "package_preflight",
      message: finding.message,
      remediation: "Inspect the named package part before relying on export fidelity.",
      ...(finding.location?.entryUri ? { partUri: finding.location.entryUri } : {}),
      ...(finding.location?.targetUri ? { resource: finding.location.targetUri } : {}),
    });
  }
  const packageLimitFinding = manifest.findings.find((finding) =>
    PACKAGE_LIMIT_FINDINGS.has(finding.code));
  if (packageLimitFinding) {
    fail("resource_limit", "package_preflight", packageLimitFinding.message,
      "Use a smaller package or raise the corresponding versioned package-inspection ceiling.",
      { detail: packageLimitFinding.code });
  }

  enforceLimit(bytes.byteLength, options.limits.compressedDocxBytes, "compressedDocxBytes", "package_preflight");
  enforceLimit(manifest.entries.length, options.limits.opcEntries, "opcEntries", "package_preflight");
  let expandedBytes = 0n;
  const expandedLimit = BigInt(options.limits.expandedOpcBytes);
  const xmlLimit = BigInt(options.limits.xmlPartBytes);
  const ratioLimit = BigInt(options.limits.opcCompressionRatio);
  for (const [index, entry] of manifest.entries.entries()) {
    if (entry.uri.length > options.limits.opcUriCharacters) {
      fail("resource_limit", "package_preflight",
        `opcUriCharacters limit exceeded for entry ${index} (${entry.uri.length} > ${options.limits.opcUriCharacters}).`,
        "Use a package with shorter canonical OPC part names.", { partUri: entry.uri });
    }
    const size = decimalAt(entry.size, `$.entries[${index}].size`);
    const compressedSize = decimalAt(entry.compressedSize, `$.entries[${index}].compressedSize`);
    expandedBytes += size;
    if (expandedBytes > expandedLimit) {
      fail("resource_limit", "package_preflight",
        `expandedOpcBytes limit exceeded (${expandedBytes.toString()} > ${expandedLimit.toString()}).`,
        "Use a smaller package or lower-expansion resources.");
    }
    if (entry.isXml && size > xmlLimit) {
      fail("resource_limit", "package_preflight",
        `xmlPartBytes limit exceeded (${size.toString()} > ${xmlLimit.toString()}).`,
        "Split or reduce the XML part before export.", { partUri: entry.uri });
    }
    if ((compressedSize === 0n && size > 0n)
      || (compressedSize > 0n && size > compressedSize * ratioLimit)) {
      fail("resource_limit", "package_preflight",
        `opcCompressionRatio limit exceeded for ${entry.uri}.`,
        "Repackage the document without a suspiciously high expansion ratio.", { partUri: entry.uri });
    }
  }

  if (manifest.packageKind !== "opc" || !manifest.isValid || !manifest.facts.mainDocumentUri) {
    fail("invalid_document", "package_preflight",
      `The input is not a valid DOCX OPC package (${manifest.packageKind}).`,
      "Repair or decrypt the document before export.");
  }
  const mainRelationship = manifest.relationships.find((relationship) =>
    relationship.ownerUri === "/"
    && relationship.targetMode === "Internal"
    && relationship.type.endsWith("/officeDocument"));
  if (!mainRelationship
    || mainRelationship.resolvedTargetUri !== manifest.facts.mainDocumentUri
    || mainRelationship.isTargetPresent !== true
    || !manifest.entries.some((entry) => entry.uri === manifest.facts.mainDocumentUri)) {
    fail("invalid_document", "package_preflight",
      "The manifest's main-document identity is incomplete or inconsistent.",
      "Repair the package-level officeDocument relationship and target part.");
  }
  if (options.reviewProfileAlreadyApplied && manifest.facts.revisions.total !== 0) {
    fail("invalid_argument", "package_preflight",
      "reviewProfileAlreadyApplied input still contains native tracked revisions.",
      "Supply the exact already-accepted/rejected package or set reviewProfileAlreadyApplied to false.");
  }

  const externalAutomatic = manifest.relationships.filter((relationship) =>
    relationship.targetMode === "External"
    && !relationship.type.toLowerCase().endsWith("/hyperlink"));
  for (const relationship of externalAutomatic) {
    const warning: RenderWarning = {
      code: "external_automatic_resource_omitted",
      severity: "warning",
      phase: "package_preflight",
      message: "An external automatic resource is not fetched by standalone export.",
      remediation: "Embed the resource in the DOCX package before export.",
      partUri: relationship.ownerUri,
      resource: relationship.target,
    };
    if (options.unsupportedContent !== "strict") addWarning(state, warning);
    addResource(state, { kind: "external_link", status: "omitted", resource: relationship.target });
  }
  if (externalAutomatic.length > 0 && options.unsupportedContent === "strict") {
    fail("resource_policy_failure", "package_preflight",
      "Strict export rejected an external automatic resource.",
      "Embed all automatic resources or use unsupportedContent: warn.");
  }

  if (manifest.facts.isMacroEnabled) {
    const warning: RenderWarning = {
      code: "macro_content_not_exported",
      severity: "warning",
      phase: "package_preflight",
      message: "Macro content is not active or embedded in standalone HTML.",
      remediation: "Remove macros or use warn policy for a static visual export.",
    };
    if (options.unsupportedContent !== "strict") addWarning(state, warning);
    if (options.unsupportedContent === "strict") {
      fail("resource_policy_failure", "package_preflight", warning.message, warning.remediation);
    }
  }
  if (manifest.facts.altChunkCount > 0) {
    const warning: RenderWarning = {
      code: "altchunk_not_supported",
      severity: "warning",
      phase: "package_preflight",
      message: "Arbitrary altChunk content is not a supported standalone export input.",
      remediation: "Materialize altChunk content into ordinary WordprocessingML before export.",
    };
    if (options.unsupportedContent !== "strict") addWarning(state, warning);
    if (options.unsupportedContent === "strict") {
      fail("resource_policy_failure", "package_preflight", warning.message, warning.remediation);
    }
  }
}

function conversionOptions(options: NormalizedOptions) {
  const commentRenderMode = {
    hidden: CommentRenderMode.Disabled,
    inline: CommentRenderMode.Inline,
    endnotes: CommentRenderMode.EndnoteStyle,
    margin: CommentRenderMode.Margin,
  }[options.commentProfile];
  return {
    pageTitle: options.title,
    cssPrefix: "docx-",
    fabricateClasses: true,
    additionalCss: "",
    commentRenderMode,
    commentCssClassPrefix: "comment-",
    paginationMode: PaginationMode.Paginated,
    paginationScale: 1,
    paginationCssClassPrefix: "page-",
    renderAnnotations: true,
    renderFootnotesAndEndnotes: true,
    renderHeadersAndFooters: true,
    renderTrackedChanges: options.reviewProfile === "markup",
    showDeletedContent: true,
    renderMoveOperations: true,
    renderUnsupportedContentPlaceholders: true,
    stampAnchors: true,
  };
}

function bootstrapHtml(title = ""): string {
  const safeTitle = title.replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;",
  })[character]!);
  return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${RENDER_CSP}"><title>${safeTitle}</title></head><body></body></html>`;
}

async function createIsolatedFrame(
  hostDocument: Document,
  state: ExecutionState,
  html = bootstrapHtml(),
  phase: ExportPhase = "browser_launch",
): Promise<HTMLIFrameElement> {
  if (!hostDocument.defaultView || !hostDocument.documentElement) {
    fail("unsupported_runtime", "browser_launch", "An attached browser document is required.",
      "Call the browser exporter from a live Window, not a worker or detached document.");
  }
  const frame = hostDocument.createElement("iframe");
  frame.dataset.docxodusExportRealm = "v1";
  frame.setAttribute("aria-hidden", "true");
  frame.setAttribute("tabindex", "-1");
  frame.sandbox.add("allow-same-origin");
  frame.style.position = "fixed";
  frame.style.left = "-100000px";
  frame.style.top = "0";
  frame.style.width = "1600px";
  frame.style.height = "1200px";
  frame.style.border = "0";
  frame.style.pointerEvents = "none";
  const loaded = new Promise<void>((resolve, reject) => {
    frame.addEventListener("load", () => resolve(), { once: true });
    frame.addEventListener("error", () => reject(new Error("isolated frame failed to load")), { once: true });
  });
  frame.srcdoc = html;
  (hostDocument.body ?? hostDocument.documentElement).appendChild(frame);
  try {
    await runPhase(state, phase, ["isolated browsing context"], () => loaded);
  } catch (error) {
    frame.remove();
    throw error;
  }
  if (!frame.contentDocument?.defaultView) {
    frame.remove();
    fail("browser_launch_failure", "browser_launch", "The isolated export frame is inaccessible.",
      "Allow same-origin srcdoc frames while retaining the script-free sandbox.");
  }
  return frame;
}

const SVG_URL_PRESENTATION_ATTRIBUTES = new Set([
  "clip-path", "cursor", "fill", "filter", "marker", "marker-end", "marker-mid",
  "marker-start", "mask", "stroke",
]);

function estimateDataUrlBytes(value: string): number | undefined {
  return dataUrlInfo(value.trim())?.byteLength;
}

function policyWarning(
  state: ExecutionState,
  options: NormalizedOptions,
  warning: Omit<RenderWarning, "severity">,
): void {
  if (options.unsupportedContent === "strict") {
    fail("resource_policy_failure", warning.phase, warning.message, warning.remediation, {
      detail: warning.detail,
      partUri: warning.partUri,
      anchorId: warning.anchorId,
      resource: warning.resource,
    });
  }
  addWarning(state, { ...warning, severity: "warning" });
}

function sanitizeCss(
  css: string,
  state: ExecutionState,
  options: NormalizedOptions,
  resourceLabel: string,
): string {
  let candidate = css;
  // Retokenize after every rewrite. Security tokens can otherwise be synthesized
  // when hostile identifier fragments meet at a removed construct boundary.
  for (let pass = 0; pass < 8; pass++) {
    const tokens = cssSecurityTokens(candidate);
    const actionable = tokens.filter((token) =>
      token.kind !== "url" || !automaticUrlAllowed(token.value, true));
    if (actionable.length === 0) return candidate;

    let cursor = 0;
    let sanitized = "";
    for (const token of tokens) {
      sanitized += candidate.slice(cursor, token.start);
      if (token.kind === "import") {
        policyWarning(state, options, {
          code: "css_import_omitted",
          phase: "docx_conversion",
          message: "A CSS import was removed from standalone output.",
          remediation: "Inline the stylesheet and all of its resources.",
          resource: `${resourceLabel}: ${token.value.slice(0, 160)}`,
        });
        // A semicolon cannot be consumed as the optional trailing whitespace of a
        // hex escape, so identifier fragments cannot join across this boundary.
        sanitized += ";";
      } else if (token.kind === "url" && automaticUrlAllowed(token.value, true)) {
        sanitized += candidate.slice(token.start, token.end);
      } else {
        policyWarning(state, options, {
          code: "external_css_resource_omitted",
          phase: "docx_conversion",
          message: "An automatic CSS resource was removed from standalone output.",
          remediation: "Embed the resource as a data URL in the DOCX conversion output.",
          resource: token.value,
        });
        sanitized += "url(\"data:,\")";
      }
      cursor = token.end;
    }
    candidate = sanitized + candidate.slice(cursor);
  }

  policyWarning(state, options, {
    code: "external_css_resource_omitted",
    phase: "docx_conversion",
    message: "CSS could not be proven free of automatic external resources.",
    remediation: "Remove dynamic CSS resource substitutions and inline all resources.",
    resource: resourceLabel,
  });
  return "";
}

function freezeCssMotion(document: Document): void {
  const removeMotionDeclarations = (style: CSSStyleDeclaration): void => {
    for (const property of Array.from(style)) {
      const normalized = property.toLowerCase();
      if (normalized.startsWith("animation") || normalized.startsWith("transition")) {
        style.removeProperty(property);
      }
    }
  };
  const visitRules = (rules: CSSRuleList): void => {
    for (const rule of Array.from(rules)) {
      if ("style" in rule && (rule as CSSStyleRule).style) {
        removeMotionDeclarations((rule as CSSStyleRule).style);
      }
      if ("cssRules" in rule) {
        visitRules((rule as CSSGroupingRule).cssRules);
      }
    }
  };
  for (const styleElement of Array.from(document.querySelectorAll<HTMLStyleElement>("style"))) {
    if (styleElement.dataset.docxodusStaticReadiness === "v1") continue;
    const sheet = styleElement.sheet;
    if (!sheet) continue;
    visitRules(sheet.cssRules);
    styleElement.textContent = Array.from(sheet.cssRules, (rule) => rule.cssText).join("\n");
  }
  for (const element of Array.from(document.querySelectorAll<HTMLElement>("[style]"))) {
    removeMotionDeclarations(element.style);
  }
}

function sanitizeConvertedDocument(
  target: Document,
  convertedHtml: string,
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  const view = target.defaultView as (Window & typeof globalThis) | null;
  if (!view) {
    fail("browser_launch_failure", "browser_launch", "The export realm has no defaultView.",
      "Use an attached same-origin browser frame.");
  }
  const parsed = new view.DOMParser().parseFromString(convertedHtml, "text/html");
  const parserError = parsed.querySelector("parsererror");
  if (parserError || !parsed.body) {
    fail("conversion_failure", "docx_conversion", "Converted HTML could not be parsed.",
      "Inspect converter diagnostics and the source package.");
  }

  for (const element of Array.from(parsed.querySelectorAll<HTMLElement>("script, iframe, object, embed, video, audio, source, track, link, base"))) {
    policyWarning(state, options, {
      code: "active_or_external_content_omitted",
      phase: "docx_conversion",
      message: `A ${element.localName} element was removed from standalone output.`,
      remediation: "Represent the content as supported static WordprocessingML.",
    });
    element.remove();
  }
  for (const meta of Array.from(parsed.querySelectorAll<HTMLMetaElement>("meta[http-equiv]"))) {
    meta.remove();
  }
  const svgAnimationElements = new Set([
    "animate", "animatemotion", "animatetransform", "set", "discard",
  ]);
  for (const animation of Array.from(parsed.querySelectorAll<Element>("svg *"))
    .filter((element) => svgAnimationElements.has(element.localName.toLowerCase()))) {
    animation.remove();
    policyWarning(state, options, {
      code: "svg_animation_omitted",
      phase: "docx_conversion",
      message: "An SVG animation element was removed from deterministic standalone output.",
      remediation: "Materialize the required SVG state as static drawable content before export.",
    });
  }
  for (const style of Array.from(parsed.querySelectorAll<HTMLStyleElement>("style"))) {
    style.textContent = sanitizeCss(style.textContent ?? "", state, options, "stylesheet");
  }

  const urlAttributes = new Set([
    "src", "poster", "data", "action", "formaction", "background", "xlink:href",
  ]);
  for (const element of Array.from(parsed.querySelectorAll<HTMLElement>("*"))) {
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      if (name.startsWith("on")) {
        element.removeAttribute(attribute.name);
        continue;
      }
      if (name === "style") {
        element.setAttribute("style", sanitizeCss(attribute.value, state, options, "inline style"));
        continue;
      }
      if (name === "srcset") {
        if (!standaloneSrcsetAllowed(attribute.value)) {
          element.removeAttribute(attribute.name);
          policyWarning(state, options, {
            code: "srcset_omitted",
            phase: "docx_conversion",
            message: "A responsive image source set was removed because every candidate could not be proven standalone.",
            remediation: "Embed one image as a data URL in src.",
            resource: attribute.value,
          });
        }
        continue;
      }
      if (SVG_URL_PRESENTATION_ATTRIBUTES.has(name)
        && cssSecurityTokens(attribute.value).length > 0) {
        element.setAttribute(
          attribute.name,
          sanitizeCss(attribute.value, state, options, `SVG ${attribute.name}`),
        );
        continue;
      }
      if (name === "href" && element.localName === "a") {
        const href = attribute.value.trim();
        if (href.startsWith("#") || /^(?:https?|mailto|tel):/i.test(href)) {
          if (!href.startsWith("#")) {
            element.setAttribute("rel", "noopener noreferrer");
            addResource(state, { kind: "external_link", status: "allowed_user_link", resource: href });
          }
          continue;
        }
        element.removeAttribute(attribute.name);
        policyWarning(state, options, {
          code: "unsafe_hyperlink_omitted",
          phase: "docx_conversion",
          message: "A hyperlink with an unsupported scheme was removed.",
          remediation: "Use an HTTPS, HTTP, mailto, tel, or document-fragment target.",
          resource: href,
        });
        continue;
      }
      if (name === "ping" && element.localName === "a") {
        element.removeAttribute(attribute.name);
        policyWarning(state, options, {
          code: "hyperlink_ping_omitted",
          phase: "docx_conversion",
          message: "A hyperlink ping target was removed from standalone output.",
          remediation: "Use an ordinary user-activated hyperlink without background tracking requests.",
          resource: attribute.value,
        });
        continue;
      }
      if ((name === "href" && element.localName !== "a") || urlAttributes.has(name)) {
        if (element.namespaceURI === "http://www.w3.org/2000/svg"
          && element.localName === "use"
          && (name === "href" || name === "xlink:href")
          && (!attribute.value.trim().startsWith("#") || attribute.value.trim().length === 1)) {
          element.removeAttribute(attribute.name);
          policyWarning(state, options, {
            code: "external_svg_use_omitted",
            phase: "docx_conversion",
            message: "An external or empty SVG use reference was removed.",
            remediation: "Reference an existing same-document SVG fragment.",
            resource: attribute.value,
          });
          continue;
        }
        const allowFragment = element.namespaceURI === "http://www.w3.org/2000/svg";
        if (!automaticUrlAllowed(attribute.value, allowFragment)) {
          element.removeAttribute(attribute.name);
          policyWarning(state, options, {
            code: "external_automatic_resource_omitted",
            phase: "docx_conversion",
            message: "An automatic external resource was removed from standalone output.",
            remediation: "Embed the resource as a data URL before export.",
            resource: attribute.value,
          });
        }
      }
    }
    if (element.localName === "form") {
      element.removeAttribute("action");
      element.addEventListener("submit", (event) => event.preventDefault());
    }
  }

  // Admit the sanitized detached tree before importing it into the live layout realm or allowing
  // the browser to decode any data URL. This bounds the second DOM copy and decoded resources.
  countDomNodes(parsed, options.limits.domNodes, "docx_conversion");
  const admittedResources = automaticResourceCount(parsed);
  enforceLimit(admittedResources.count, options.limits.automaticResources,
    "automaticResources", "docx_conversion");
  enforceLimit(admittedResources.bytes, options.limits.automaticResourceBytes,
    "automaticResourceBytes", "docx_conversion");

  target.documentElement.lang = parsed.documentElement.lang || "en-US";
  target.title = options.title;
  for (const style of Array.from(parsed.head.querySelectorAll("style"))) {
    target.head.appendChild(target.importNode(style, true));
  }
  const staticReadinessStyle = target.createElement("style");
  staticReadinessStyle.dataset.docxodusStaticReadiness = "v1";
  staticReadinessStyle.textContent = "*,*::before,*::after{animation:none!important;transition:none!important}";
  target.head.appendChild(staticReadinessStyle);
  target.body.replaceChildren(...Array.from(parsed.body.childNodes, (node) => target.importNode(node, true)));
  freezeCssMotion(target);
}

function inventoryConvertedContent(
  document: Document,
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  for (const placeholder of Array.from(document.querySelectorAll<HTMLElement>("[data-content-type]"))) {
    const outcome: UnsupportedContentOutcome = {
      contentType: placeholder.dataset.contentType ?? "Other",
      ...(placeholder.dataset.elementName ? { elementName: placeholder.dataset.elementName } : {}),
      ...(placeholder.closest<HTMLElement>("[data-source-anchor-id]")?.dataset.sourceAnchorId
        ? { anchorId: placeholder.closest<HTMLElement>("[data-source-anchor-id]")!.dataset.sourceAnchorId }
        : {}),
      action: "placeholder",
    };
    addUnsupportedContent(state, outcome);
    policyWarning(state, options, {
      code: "unsupported_content_placeholder",
      phase: "docx_conversion",
      message: `${outcome.contentType} is represented by a visible placeholder.`,
      remediation: "Replace it with a supported static representation for faithful export.",
      ...(outcome.anchorId ? { anchorId: outcome.anchorId } : {}),
    });
  }
}

function recordFontResolution(
  result: BrowserFontResult,
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  enforceLimit(result.resolutions.length, state.limits.fontRequests, "fontRequests", "font_loading");
  enforceDiagnosticAdmission(state, result.resolutions.length);
  state.fontIdentity = { ...result.identity };
  state.fonts.push(...result.resolutions.map((resolution) => ({
    ...resolution,
    requestedFamilies: [...resolution.requestedFamilies],
    requestedFamilyKinds: [...resolution.requestedFamilyKinds],
    ...(resolution.licenseEvidence
      ? { licenseEvidence: { ...resolution.licenseEvidence } }
      : {}),
  })));
  if (result.renderedTextNodeCount > 0 && result.resolutions.length === 0) {
    fail("resource_policy_failure", "font_loading",
      "Rendered text was present, but the canonical font inventory was unexpectedly empty.",
      "Report this font inventory invariant failure before relying on the output.");
  }
  for (const resolution of result.resolutions) {
    const severity = "warning" as const;
    if (resolution.status === "missing") {
      addWarning(state, {
        code: "font_unavailable",
        severity,
        phase: "font_loading",
        message: `No configured face resolved the required font family: ${resolution.requestedFamily}.`,
        remediation: "Install or explicitly supply the required font before export.",
        resource: resolution.requestedFamily,
      });
    }
    if (resolution.status === "load_failed") {
      addWarning(state, {
        code: "font_load_failed",
        severity,
        phase: "font_loading",
        message: `The configured font face for ${resolution.requestedFamily} could not be decoded or loaded.`,
        remediation: "Replace the configured font with a valid browser-supported face.",
        resource: resolution.requestedFamily,
      });
    }
    if (resolution.status === "substituted") {
      addWarning(state, {
        code: "font_substituted",
        severity,
        phase: "font_loading",
        message: `${resolution.requestedFamily} was resolved to ${resolution.resolvedFamily ?? "a configured substitute"}.`,
        remediation: "Supply an exact configured family when substitution is not acceptable.",
        resource: resolution.requestedFamily,
      });
    }
    if (resolution.faceMatch === "synthesized") {
      addWarning(state, {
        code: "font_face_synthesized",
        severity,
        phase: "font_loading",
        message: `The requested style, weight, or stretch for ${resolution.requestedFamily} requires synthesis.`,
        remediation: "Supply an exact configured face for this style, weight, and stretch.",
        resource: resolution.requestedFamily,
      });
    }
    if (resolution.metricCompatible === false) {
      addWarning(state, {
        code: "font_metric_mismatch",
        severity,
        phase: "font_loading",
        message: `The substitute selected for ${resolution.requestedFamily} is not metrically compatible.`,
        remediation: "Supply the exact family or review PageMap and line wrapping before publication.",
        resource: resolution.requestedFamily,
      });
    }
    if (resolution.glyphCoverage === "partial") {
      addWarning(state, {
        code: "font_glyph_coverage_partial",
        severity,
        phase: "font_loading",
        message: `The configured face for ${resolution.requestedFamily} covers only part of the rendered text.`,
        remediation: "Supply a face with complete glyph coverage for the reported sample.",
        resource: resolution.requestedFamily,
      });
    }
  }
  if (result.resolutions.some((resolution) =>
    resolution.status === "unverified" || resolution.source === "browser")) {
    addWarning(state, {
      code: "font_environment_unverified",
      severity: "warning",
      phase: "font_loading",
      message: "The browser loaded the requested CSS font families, but their exact files and substitutions are not attestable.",
      remediation: "Use an explicit verified font resolver when exact font identity is required.",
    });
  }
  const strictFailures = result.resolutions.filter((resolution) =>
    resolution.status !== "resolved"
    || (resolution.source !== "configured" && resolution.source !== "attested")
    || resolution.faceMatch !== "exact"
    || resolution.glyphCoverage !== "complete"
    || !resolution.fileSha256
    || !resolution.licenseEvidence);
  if (options.strictFonts && strictFailures.length > 0) {
    fail("resource_policy_failure", "font_loading",
      "Strict font policy rejected a non-exact, unverified, or incompletely covered font outcome.",
      "Supply exact verified faces with complete glyph coverage or disable strictFonts.", {
        detail: strictFailures.map(({ requestId, status }) => `${requestId}:${status}`).join(", "),
      });
  }
}

function recordExactFontReadiness(
  probes: readonly FontReadinessProbe[],
  state: ExecutionState,
): void {
  enforceLimit(probes.length, state.limits.fontRequests, "fontRequests", "font_loading");
  enforceDiagnosticAdmission(state, probes.length);
  state.fontReadiness.push(...probes.map((probe) => ({ ...probe })));
}

function rethrowBrowserFontError(error: unknown): never {
  if (error instanceof BrowserFontError) {
    fail(
      error.kind === "resource_limit" ? "resource_limit" : "resource_policy_failure",
      "font_loading",
      error.message,
      error.kind === "resource_limit"
        ? "Lower the number or size of configured fonts or raise the versioned font limit."
        : "Return one canonical, digest-verified resolver outcome for every font request.",
      { detail: error.detail, cause: error },
    );
  }
  throw error;
}

function replaceFailedVisual(element: Element, label: string): void {
  const placeholder = element.ownerDocument.createElement("span");
  placeholder.className = "docxodus-export-resource-placeholder";
  placeholder.setAttribute("role", "img");
  placeholder.setAttribute("aria-label", `${label} unavailable`);
  placeholder.textContent = `[${label} unavailable]`;
  const style = element.getAttribute("style");
  if (style) placeholder.setAttribute("style", style);
  element.replaceWith(placeholder);
}

function recordImageReadiness(
  images: HTMLImageElement[],
  probes: VisualResourceProbe[],
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  let htmlImageIndex = 0;
  probes.forEach((probe) => {
    const image = probe.source === "html-image" ? images[htmlImageIndex++] : undefined;
    const source = image?.getAttribute("src") ?? "";
    const metadata = /^data:([^;,]+)/i.exec(source);
    addResource(state, {
      kind: "image",
      status: probe.status === "complete" ? "embedded" : "omitted",
      readiness: probe.status,
      contentKey: probe.contentKey,
      resource: probe.resource,
      ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
      ...(probe.message ? { message: probe.message } : {}),
      mediaType: probe.mediaType ?? metadata?.[1],
      byteLength: probe.byteLength ?? estimateDataUrlBytes(source),
    });
    if (probe.status === "failed") {
      if (image) replaceFailedVisual(image, probe.resource);
      policyWarning(state, options, {
        code: "image_decode_failed",
        phase: "image_decoding",
        message: `An embedded image could not be decoded: ${probe.resource}.`,
        remediation: "Replace the image with a supported, non-corrupt embedded image.",
        ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
        resource: probe.resource,
      });
      if (probe.source !== "html-image") {
        fail("resource_policy_failure", "image_decoding",
          `A ${probe.source} dependency did not produce decoded pixels: ${probe.resource}.`,
          "Embed a supported image and ensure every CSS/SVG reference resolves before export.", {
            detail: probe.message,
            anchorId: probe.anchorId,
            resource: probe.resource,
            pending: [`image:${probe.resource}`],
          });
      }
    }
  });
}

function graphicElements(document: Document): Element[] {
  return Array.from(new Set(Array.from(document.querySelectorAll<Element>(
    "svg, [data-docxodus-materialization]",
  )))).filter((element) =>
    element.closest("[data-docxodus-materialization]") === element
    || !element.parentElement?.closest("[data-docxodus-materialization]"));
}

function recordGraphicReadiness(
  elements: Element[],
  probes: VisualResourceProbe[],
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  let graphicIndex = 0;
  probes.forEach((probe) => {
    addResource(state, {
      kind: probe.kind,
      status: probe.status === "complete" ? "inline" : "omitted",
      readiness: probe.status,
      contentKey: probe.contentKey,
      resource: probe.resource,
      ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
      ...(probe.message ? { message: probe.message } : {}),
    });
    if (probe.status === "failed") {
      const element = probe.source === "graphic" ? elements[graphicIndex] : undefined;
      if (element) replaceFailedVisual(element, probe.resource);
      policyWarning(state, options, {
        code: "graphic_materialization_failed",
        phase: "chart_svg_materialization",
        message: `${probe.kind} content did not finish materializing: ${probe.resource}.`,
        remediation: "Replace the graphic or use a supported static SVG representation.",
        ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
        resource: probe.resource,
      });
      if (probe.source === "svg-use") {
        fail("resource_policy_failure", "chart_svg_materialization",
          `An SVG use reference did not resolve to drawable local content: ${probe.resource}.`,
          "Use a same-document fragment that names an existing drawable SVG target.", {
            detail: probe.message,
            anchorId: probe.anchorId,
            resource: probe.resource,
            pending: [`materialization:${probe.resource}`],
          });
      }
    }
    if (probe.source === "graphic") graphicIndex++;
  });
}

function rewriteCssFragmentUrls(value: string, targets: ReadonlyMap<string, string>): string {
  const tokens = cssSecurityTokens(value);
  let rewritten = value;
  for (const token of [...tokens].reverse()) {
    if (token.kind !== "url" || !token.value.startsWith("#")) continue;
    const target = targets.get(token.value.slice(1));
    if (!target) continue;
    rewritten = `${rewritten.slice(0, token.start)}url("#${target}")${rewritten.slice(token.end)}`;
  }
  return rewritten;
}

function normalizeFragmentTargets(
  document: Document,
  pages: HTMLElement[],
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  const firstFootnote = new Set<string>();
  for (const item of Array.from(document.querySelectorAll<HTMLElement>(
    ".page-footnotes [data-footnote-id]",
  ))) {
    const id = item.dataset.footnoteId;
    if (id && !firstFootnote.has(id)) {
      item.id = `fn-${id}`;
      firstFootnote.add(id);
    }
  }

  const globalTargets = new Map<string, string>();
  const pageTargets = new Map<HTMLElement, Map<string, string>>();
  const occurrences = new Map<string, number>();
  for (const page of pages) {
    const localTargets = new Map<string, string>();
    pageTargets.set(page, localTargets);
    const allocate = (element: Element): void => {
      const original = element.id;
      const occurrence = occurrences.get(original) ?? 0;
      occurrences.set(original, occurrence + 1);
      const resolved = occurrence === 0
        ? original
        : `${original}--page-${page.dataset.pageNumber ?? "0"}-${occurrence}`;
      element.id = resolved;
      if (!localTargets.has(original)) localTargets.set(original, resolved);
      if (!globalTargets.has(original)) globalTargets.set(original, resolved);
    };
    const svgRoots = Array.from(page.querySelectorAll<SVGSVGElement>("svg"))
      .filter((svg) => svg.parentElement?.closest("svg") === null);
    for (const svg of svgRoots) {
      const svgTargets = new Map<string, string>();
      const identified = [
        ...(svg.hasAttribute("id") ? [svg] : []),
        ...Array.from(svg.querySelectorAll<SVGElement>("[id]")),
      ];
      for (const element of identified) {
        const original = element.id;
        allocate(element);
        if (!svgTargets.has(original)) svgTargets.set(original, element.id);
      }
      for (const element of [svg, ...Array.from(svg.querySelectorAll<SVGElement>("*"))]) {
        if (element.localName === "use") {
          for (const name of ["href", "xlink:href"]) {
            const value = element.getAttribute(name);
            if (!value?.startsWith("#")) continue;
            const target = svgTargets.get(value.slice(1));
            if (target) element.setAttribute(name, `#${target}`);
          }
        }
        for (const name of [...SVG_URL_PRESENTATION_ATTRIBUTES, "style"]) {
          const value = element.getAttribute(name);
          if (value) element.setAttribute(name, rewriteCssFragmentUrls(value, svgTargets));
        }
      }
    }
    for (const element of Array.from(page.querySelectorAll<HTMLElement>("[id]"))) {
      if (element.closest("svg")) continue;
      allocate(element);
    }
  }

  for (const link of Array.from(document.querySelectorAll<HTMLAnchorElement>("a[href^='#']"))) {
    const original = link.getAttribute("href")!.slice(1);
    let page: HTMLElement | null = link.closest<HTMLElement>(".page-box");
    const target = (page ? pageTargets.get(page)?.get(original) : undefined)
      ?? globalTargets.get(original);
    if (target) {
      link.setAttribute("href", `#${target}`);
    } else {
      link.removeAttribute("href");
      link.setAttribute("aria-disabled", "true");
      link.tabIndex = -1;
      policyWarning(state, options, {
        code: "fragment_target_unavailable",
        phase: "running_story_placement",
        message: `Fragment target #${original} is unavailable in the final page tree.`,
        remediation: "Repair the bookmark/note target or use warn policy to retain an inert label.",
        resource: `#${original}`,
      });
    }
  }

  const ids = Array.from(document.querySelectorAll<HTMLElement>("[id]"), (element) => element.id);
  if (new Set(ids).size !== ids.length) {
    fail("output_verification_failure", "running_story_placement",
      "The final standalone document contains duplicate fragment IDs.",
      "Report the source document and duplicate target to Docxodus.");
  }
}

function standaloneStyle(pages: HTMLElement[]): string {
  const namedPages = new Map<string, { width: number; height: number }>();
  for (const page of pages) {
    const sectionIndex = Number.parseInt(page.dataset.sectionIndex ?? "0", 10);
    const width = Number.parseFloat(page.style.width);
    const height = Number.parseFloat(page.style.height);
    const name = `docxodus-section-${sectionIndex}`;
    if (!namedPages.has(name)) namedPages.set(name, { width, height });
    page.style.setProperty("page", name);
    page.dataset.pageWidthPt = String(width);
    page.dataset.pageHeightPt = String(height);
  }
  const rules = Array.from(namedPages, ([name, dimensions]) =>
    `@page ${name} { size: ${dimensions.width}pt ${dimensions.height}pt; margin: 0; }`).join("\n");
  return `
@page { margin: 0; }
${rules}
html, body { margin: 0; padding: 0; }
#pagination-container.page-container { min-height: 0 !important; }
@media screen {
  html, body { background: #e5e7eb; }
  #pagination-container.page-container {
    display: flex !important;
    flex-direction: column;
    align-items: center;
    gap: 20px !important;
    padding: 20px !important;
    background: transparent !important;
  }
  .page-box { box-shadow: 0 2px 8px rgba(0, 0, 0, .18); }
}
@media print {
  html, body { margin: 0 !important; padding: 0 !important; background: transparent !important; }
  #pagination-container.page-container {
    display: block !important;
    gap: 0 !important;
    padding: 0 !important;
    background: transparent !important;
  }
  .page-box {
    zoom: 1 !important;
    transform: none !important;
    margin: 0 !important;
    box-shadow: none !important;
    break-after: page;
    page-break-after: always;
  }
  .page-box:last-child { break-after: auto; page-break-after: auto; }
}`;
}

function finalizePageTree(
  document: Document,
  pages: HTMLElement[],
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  document.querySelector("#pagination-staging, .page-staging")?.remove();
  for (const element of Array.from(document.querySelectorAll<HTMLElement>("*"))) {
    element.removeAttribute("contenteditable");
    element.removeAttribute("data-anchor");
    element.removeAttribute("data-committed-text");
    element.removeAttribute("data-docxodus-materialization");
    element.removeAttribute("data-docxodus-materialization-state");
    element.removeAttribute("data-docxodus-materialization-id");
    element.removeAttribute("data-docxodus-materialization-error");
    element.removeAttribute("draggable");
    element.style.removeProperty("will-change");
    element.style.removeProperty("contain");
  }
  for (const page of pages) {
    for (const property of [
      "zoom", "transform", "transform-origin", "margin-right", "margin-bottom", "box-shadow",
    ]) page.style.removeProperty(property);
  }
  normalizeFragmentTargets(document, pages, state, options);
  document.documentElement.dataset.docxodusStandalone = "v1";
  const policy = document.querySelector<HTMLMetaElement>(
    "meta[http-equiv='Content-Security-Policy' i]",
  );
  if (!policy) {
    fail("output_verification_failure", "running_story_placement",
      "The standalone document lost its Content Security Policy.",
      "Report the materializer defect; finalization must retain the closed policy.");
  }
  policy.content = STANDALONE_CSP;
  const style = document.createElement("style");
  style.id = "docxodus-standalone-style";
  style.textContent = standaloneStyle(pages);
  document.head.appendChild(style);
}

function assertNoClippedContent(document: Document): void {
  for (const footnotes of Array.from(document.querySelectorAll<HTMLElement>(".page-footnotes"))) {
    const boundary = footnotes.getBoundingClientRect().bottom;
    const hasVisibleOverflow = Array.from(footnotes.querySelectorAll<HTMLElement>("*"))
      .some((element) => element.getBoundingClientRect().bottom > boundary + 1);
    if (footnotes.scrollHeight > footnotes.clientHeight + 1 && hasVisibleOverflow) {
      fail("pagination_failure", "running_story_placement",
        "A footnote continuation is clipped in the final page tree.",
        "Report the unsupported note structure; eligible text paragraphs are continued losslessly.");
    }
  }
  for (const content of Array.from(document.querySelectorAll<HTMLElement>(".page-content"))) {
    if (content.scrollHeight > content.clientHeight + 1) {
      const pageNumber = content.closest<HTMLElement>(".page-box")?.dataset.pageNumber ?? "unknown";
      const boundary = content.getBoundingClientRect().bottom;
      const overflow = Array.from(content.querySelectorAll<HTMLElement>("*"))
        .map((element) => ({
          element,
          bottom: element.getBoundingClientRect().bottom,
        }))
        .filter(({ bottom }) => bottom > boundary + 1)
        .sort((left, right) => right.bottom - left.bottom)
        .slice(0, 3)
        .map(({ element, bottom }) =>
          `${element.localName}${element.className ? `.${String(element.className).trim().replace(/\s+/g, ".")}` : ""} (+${(bottom - boundary).toFixed(1)}px)`);
      // Collapsed trailing margins contribute to scrollHeight even though no
      // descendant pixels are clipped (a common final blank Word paragraph).
      if (overflow.length === 0) continue;
      fail("pagination_failure", "running_story_placement",
        `Page ${pageNumber} body content is clipped (${content.scrollHeight}px scroll height in a ${content.clientHeight}px band; ${overflow.join(", ")}).`,
        "Split the oversized block or reduce its dimensions before export.");
    }
  }
}

function checkpointAttemptState(state: ExecutionState): AttemptStateCheckpoint {
  return {
    readiness: state.readiness.length,
    warnings: state.warnings.length,
    fonts: state.fonts.length,
    fontReadiness: state.fontReadiness.length,
    fontIdentity: state.fontIdentity ? { ...state.fontIdentity } : undefined,
    resources: state.resources.length,
    unsupportedContent: state.unsupportedContent.length,
  };
}

function restoreAttemptState(state: ExecutionState, checkpoint: AttemptStateCheckpoint): void {
  state.readiness.length = checkpoint.readiness;
  state.warnings.length = checkpoint.warnings;
  state.fonts.length = checkpoint.fonts;
  state.fontReadiness.length = checkpoint.fontReadiness;
  state.fontIdentity = checkpoint.fontIdentity ? { ...checkpoint.fontIdentity } : undefined;
  state.resources.length = checkpoint.resources;
  state.unsupportedContent.length = checkpoint.unsupportedContent;
}

async function pristineAttemptAgreementSignature(
  pageTreeSignature: string,
  state: ExecutionState,
): Promise<string> {
  const fonts = state.fontReadiness.map((font) => ({
    requestKey: font.requestKey,
    available: font.available,
  })).sort((left, right) => compareCodeUnits(left.requestKey, right.requestKey));
  const resources = state.resources
    .filter((resource) => resource.kind !== "external_link")
    .map((resource) => ({
      kind: resource.kind,
      status: resource.status,
      readiness: resource.readiness,
      contentKey: resource.contentKey,
      resource: resource.resource,
      anchorId: resource.anchorId,
      mediaType: resource.mediaType,
      byteLength: resource.byteLength,
    }))
    .sort((left, right) => compareCodeUnits(canonicalJson(left), canonicalJson(right)));
  return canonicalMaterialDigest(
    "docxodus:pristine-attempt-agreement:v1",
    { pageTreeSignature, fonts, resources },
  );
}

function serializeDocument(document: Document): string {
  return `<!doctype html>\n${document.documentElement.outerHTML}`;
}

function automaticResourceCount(document: Document): { count: number; bytes: number } {
  let count = 0;
  let bytes = 0;
  const add = (value: string): void => {
    count++;
    bytes += estimateDataUrlBytes(value) ?? 0;
  };
  const addCssUrls = (value: string): void => {
    for (const token of cssSecurityTokens(value)) {
      if (token.kind === "url" || token.kind === "substitution") add(token.value);
    }
  };
  for (const element of Array.from(document.querySelectorAll<HTMLElement>("*"))) {
    for (const name of ["src", "poster", "background", "data"]) {
      const value = element.getAttribute(name);
      if (value) add(value);
    }
    const srcset = element.getAttribute("srcset");
    if (srcset) add(srcset.replace(/\s+(?:\d+(?:\.\d+)?x|\d+w)\s*$/i, ""));
    const isSvg = element.namespaceURI === "http://www.w3.org/2000/svg";
    if (isSvg) {
      for (const name of ["href", "xlink:href"]) {
        const value = element.getAttribute(name);
        if (value) add(value);
      }
      for (const name of SVG_URL_PRESENTATION_ATTRIBUTES) {
        const value = element.getAttribute(name);
        if (value) addCssUrls(value);
      }
    }
    const inlineStyle = element.getAttribute("style");
    if (inlineStyle) addCssUrls(inlineStyle);
  }
  for (const style of Array.from(document.querySelectorAll<HTMLStyleElement>("style"))) {
    addCssUrls(style.textContent ?? "");
  }
  return { count, bytes };
}

async function layoutDigestForOptions(options: NormalizedOptions): Promise<string> {
  const layoutContract = {
    title: options.title,
    reviewProfile: options.reviewProfile,
    reviewProfileAlreadyApplied: options.reviewProfileAlreadyApplied,
    commentProfile: options.commentProfile,
    pagination: {
      mode: "paginated",
      scale: 1,
      pageGap: 0,
      showPageNumbers: false,
      fragmentParagraphs: true,
      cssPrefix: "page-",
    },
  };
  return canonicalMaterialDigest("docxodus:layout-options:v1", layoutContract);
}

async function runtimePolicyDigestForOptions(
  options: NormalizedOptions,
  runtimeAssets: RuntimeAssetIdentity,
): Promise<string> {
  return canonicalMaterialDigest("docxodus:runtime-policy:v1", {
    assetGraphDigest: runtimeAssets.graphDigest,
    assetPackageVersion: runtimeAssets.packageVersion,
    isolation: {
      attachedSameOriginFrame: true,
      scripts: "denied",
      automaticNetwork: "denied",
      sandbox: ["allow-same-origin"],
    },
    limits: options.limits,
    strictFonts: options.strictFonts,
    timeoutMs: options.timeoutMs,
    unsupportedContent: options.unsupportedContent,
  });
}

function observedRuntimeFacts(document: Document): ExportRuntimeObservedFacts {
  const view = document.defaultView;
  if (!view) {
    fail("unsupported_runtime", "browser_launch", "The render document has no browser identity.",
      "Use an attached browser document.");
  }
  return {
    runtimeKind: "browser",
    locale: view.navigator.language || "und",
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC",
    viewport: [view.innerWidth, view.innerHeight],
    deviceScaleFactor: view.devicePixelRatio,
    media: {
      colorScheme: view.matchMedia("(prefers-color-scheme: dark)").matches
        ? "dark"
        : view.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "no-preference",
      reducedMotion: view.matchMedia("(prefers-reduced-motion: reduce)").matches
        ? "reduce" : "no-preference",
      forcedColors: view.matchMedia("(forced-colors: active)").matches ? "active" : "none",
      printMedia: true,
    },
    networkIsolation: "contextRestricted",
  };
}

async function rendererIdentity(
  document: Document,
  layoutDigest: string,
  runtimePolicyDigest: string,
  runtimeAssets: RuntimeAssetIdentity,
  runtimeVersion: VersionInfo,
  fontIdentity: FontConfigurationIdentity,
  fonts: FontResolution[],
): Promise<{
  rendererFingerprint: string;
  observed: ExportRuntimeObservedFacts;
  fontIdentity: FontConfigurationIdentity;
}> {
  const view = document.defaultView;
  if (!view) {
    fail("unsupported_runtime", "browser_launch", "The render document has no browser identity.",
      "Use an attached browser document.");
  }
  const fontRecords = fonts.map((font) => ({ ...font }))
    .sort((left, right) => compareCodeUnits(canonicalJson(left), canonicalJson(right)));
  const observed = observedRuntimeFacts(document);
  const fingerprint = {
    contract: "docxodus-standalone-browser-v1",
    verification: "browserObserved",
    runtimeAssets,
    runtimeVersion,
    paginatorContractVersion: 1,
    pageMapSchemaVersion: 1,
    renderReportSchemaVersion: 3,
    layoutDigest,
    runtimePolicyDigest,
    fontIdentity,
    fonts: fontRecords,
    observed,
  };
  return {
    rendererFingerprint: await canonicalMaterialDigest(
      "docxodus:renderer-fingerprint:v1",
      fingerprint,
    ),
    observed,
    fontIdentity,
  };
}

function reportPages(pageMap: PageMap): CompleteRenderReport["pages"] {
  return pageMap.pages.map((page) => ({ ...page }));
}

function storyForAnchor(anchorId: string): PageMap["fragments"][number]["story"] {
  const first = anchorId.indexOf(":");
  const second = first < 0 ? -1 : anchorId.indexOf(":", first + 1);
  const scope = first >= 0 && second > first ? anchorId.slice(first + 1, second) : "body";
  if (scope.startsWith("hdr")) return "header";
  if (scope.startsWith("ftr")) return "footer";
  if (scope === "fn") return "footnote";
  if (scope === "en") return "endnote";
  if (scope === "cmt") return "comment";
  return "body";
}

function visibleRectWithinPage(
  document: Document,
  element: HTMLElement,
  page: HTMLElement,
): { left: number; top: number; right: number; bottom: number } {
  const pageRect = page.getBoundingClientRect();
  const rect = element.getBoundingClientRect();
  let left = Math.max(rect.left, pageRect.left);
  let top = Math.max(rect.top, pageRect.top);
  let right = Math.min(rect.right, pageRect.right);
  let bottom = Math.min(rect.bottom, pageRect.bottom);
  const clips = (value: string) =>
    value === "hidden" || value === "clip" || value === "scroll" || value === "auto";
  for (let ancestor = element.parentElement;
    ancestor && ancestor !== page;
    ancestor = ancestor.parentElement) {
    const style = document.defaultView!.getComputedStyle(ancestor);
    const ancestorRect = ancestor.getBoundingClientRect();
    if (clips(style.overflowX)) {
      left = Math.max(left, ancestorRect.left);
      right = Math.min(right, ancestorRect.right);
    }
    if (clips(style.overflowY)) {
      top = Math.max(top, ancestorRect.top);
      bottom = Math.min(bottom, ancestorRect.bottom);
    }
  }
  return { left, top, right, bottom };
}

function assertStandaloneResourceAudit(document: Document): void {
  if (document.querySelector(
    "script, iframe, object, embed, video, audio, source, track, link, base, meta[http-equiv='refresh' i]",
  )) {
    throw new Error("offline output contains active or external-loading content");
  }
  const meta = document.querySelector<HTMLMetaElement>(
    "meta[http-equiv='Content-Security-Policy' i]",
  );
  if (!meta || meta.content !== STANDALONE_CSP) {
    throw new Error("offline output CSP is missing or changed");
  }
  for (const element of Array.from(document.querySelectorAll<HTMLElement>("*"))) {
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase();
      if (name.startsWith("on")) throw new Error(`offline output retained ${attribute.name}`);
      if (name === "style") {
        const unsafe = cssSecurityTokens(attribute.value).find((token) =>
          token.kind !== "url" || !automaticUrlAllowed(token.value, true));
        if (unsafe) throw new Error(`offline inline style retained ${unsafe.kind}`);
      }
      if (name === "srcset" && !standaloneSrcsetAllowed(attribute.value)) {
        throw new Error("offline output retained an unsafe srcset");
      }
      if (name === "href" && element.localName === "a") {
        const href = attribute.value.trim();
        if (!href.startsWith("#") && !/^(?:https?|mailto|tel):/i.test(href)) {
          throw new Error(`offline output retained unsafe hyperlink ${href}`);
        }
      } else if (name === "ping" && element.localName === "a") {
        throw new Error("offline output retained a hyperlink ping target");
      } else if (["src", "poster", "data", "action", "formaction", "background", "xlink:href", "href"].includes(name)) {
        const allowFragment = element.namespaceURI === "http://www.w3.org/2000/svg";
        if (!automaticUrlAllowed(attribute.value, allowFragment)) {
          throw new Error(`offline output retained automatic URL ${attribute.value}`);
        }
      }
    }
    if (element.localName === "style") {
      const unsafe = cssSecurityTokens(element.textContent ?? "").find((token) =>
        token.kind !== "url" || !automaticUrlAllowed(token.value, true));
      if (unsafe) throw new Error(`offline stylesheet retained ${unsafe.kind}`);
    }
  }
}

async function verifyOfflineReopen(
  hostDocument: Document,
  html: string,
  expectedPageMap: PageMap,
  state: ExecutionState,
  signal: AbortSignal,
): Promise<void> {
  const frame = await createIsolatedFrame(hostDocument, state, html, "output_verification");
  try {
    const reopened = frame.contentDocument!;
    const pages = Array.from(reopened.querySelectorAll<HTMLElement>(".page-box"));
    const reopenedReadiness = await awaitFinalPrintReadiness(reopened, {
      timeoutMs: Math.max(1, state.deadline - monotonicNow()),
      signal,
      limits: {
        fontRequests: state.limits.fontRequests,
        fontSampleCodePoints: state.limits.fontSampleCodePoints,
        visualResources: state.limits.automaticResources,
        domNodes: state.limits.domNodes,
        automaticResourceBytes: state.limits.automaticResourceBytes,
      },
    });
    const expectedFonts = state.fontReadiness.map((font) => ({
      requestKey: font.requestKey,
      available: font.available,
    })).sort((left, right) => compareCodeUnits(left.requestKey, right.requestKey));
    const reopenedFonts = reopenedReadiness.fonts.map((font) => ({
      requestKey: font.requestKey,
      available: font.available,
    })).sort((left, right) => compareCodeUnits(left.requestKey, right.requestKey));
    if (canonicalJson(expectedFonts) !== canonicalJson(reopenedFonts)) {
      fail("output_verification_failure", "font_loading",
        "The reopened standalone document changed its exact font-request inventory or availability.",
        "Use the same embedded/configured fonts through materialization and offline reopen.", {
          pending: boundedPendingResources(reopenedReadiness.fonts
            .map((font) => `font:${font.requestedFamily}:${font.requestKey.slice(0, 12)}`)),
        });
    }
    const expectedVisualResources = state.resources
      .filter((resource) => resource.kind !== "external_link" && resource.readiness === "complete")
      .map((resource) => ({
        kind: resource.kind,
        resource: resource.resource ?? "",
        readiness: resource.readiness,
        contentKey: resource.contentKey,
      }))
      .sort((left, right) => compareCodeUnits(canonicalJson(left), canonicalJson(right)));
    const reopenedVisualResources = [
      ...reopenedReadiness.images,
      ...reopenedReadiness.graphics,
    ].map((resource) => ({
      kind: resource.kind,
      resource: resource.resource,
      readiness: resource.status,
      contentKey: resource.contentKey,
    })).sort((left, right) => compareCodeUnits(canonicalJson(left), canonicalJson(right)));
    if (canonicalJson(expectedVisualResources) !== canonicalJson(reopenedVisualResources)) {
      fail("output_verification_failure", "output_verification",
        "The reopened standalone document changed its ready visual-resource inventory.",
        "Materialize every image and graphic before recording the report and final page tree.", {
          pending: boundedPendingResources(reopenedVisualResources.map((resource) =>
            `${resource.kind}:${resource.resource}`)),
        });
    }
    countDomNodes(reopened, state.limits.domNodes, "output_verification");
    const resources = automaticResourceCount(reopened);
    enforceLimit(resources.count, state.limits.automaticResources,
      "automaticResources", "output_verification");
    enforceLimit(resources.bytes, state.limits.automaticResourceBytes,
      "automaticResourceBytes", "output_verification");
    assertStandaloneResourceAudit(reopened);

    if (pages.length !== expectedPageMap.pages.length) {
      throw new Error(`offline page count changed (${pages.length} != ${expectedPageMap.pages.length})`);
    }
    for (let index = 0; index < pages.length; index++) {
      const page = pages[index];
      const expected = expectedPageMap.pages[index];
      const rect = page.getBoundingClientRect();
      const widthPt = rect.width * 72 / 96;
      const heightPt = rect.height * 72 / 96;
      if (Math.abs(widthPt - expected.width) > 0.1
        || Math.abs(heightPt - expected.height) > 0.1
        || Number.parseInt(page.dataset.pageNumber ?? "0", 10) !== expected.pageNumber
        || Number.parseInt(page.dataset.pageInSection ?? "0", 10) !== expected.pageInSection
        || Number.parseInt(page.dataset.sectionIndex ?? "0", 10) !== expected.sectionIndex
        || reopened.defaultView!.getComputedStyle(page).page !== expected.pageName) {
        throw new Error(`offline geometry changed on page ${expected.pageNumber}`);
      }
    }

    const actualFragments: PageMap["fragments"] = [];
    for (const page of pages) {
      const pageNumber = Number.parseInt(page.dataset.pageNumber ?? "0", 10);
      const expectedPage = expectedPageMap.pages.find((candidate) => candidate.pageNumber === pageNumber);
      if (!expectedPage) throw new Error(`offline output added page ${pageNumber}`);
      const pageRect = page.getBoundingClientRect();
      const pointPerRenderedX = expectedPage.width / pageRect.width;
      const pointPerRenderedY = expectedPage.height / pageRect.height;
      for (const element of Array.from(
        page.querySelectorAll<HTMLElement>("[data-source-anchor-id][data-page-fragment-id]"),
      )) {
        const style = reopened.defaultView!.getComputedStyle(element);
        if (style.display === "none" || style.visibility === "hidden") continue;
        const rect = element.getBoundingClientRect();
        const visible = visibleRectWithinPage(reopened, element, page);
        if (rect.width <= 0 || rect.height <= 0
          || visible.right <= visible.left || visible.bottom <= visible.top) continue;
        const anchorId = element.dataset.sourceAnchorId!;
        actualFragments.push({
          fragmentId: element.dataset.pageFragmentId!,
          anchorId,
          fragmentIndex: Number.parseInt(element.dataset.fragmentIndex ?? "-1", 10),
          pageNumber,
          geometry: {
            x: (visible.left - pageRect.left) * pointPerRenderedX,
            y: (visible.top - pageRect.top) * pointPerRenderedY,
            width: (visible.right - visible.left) * pointPerRenderedX,
            height: (visible.bottom - visible.top) * pointPerRenderedY,
          },
          story: storyForAnchor(anchorId),
          inTableCell: element.matches("td,th") || element.closest("td,th") !== null,
        });
      }
    }
    if (actualFragments.length !== expectedPageMap.fragments.length) {
      throw new Error(
        `offline fragment inventory changed (${actualFragments.length} != ${expectedPageMap.fragments.length})`,
      );
    }
    for (let index = 0; index < actualFragments.length; index++) {
      const actual = actualFragments[index];
      const expected = expectedPageMap.fragments[index];
      if (actual.fragmentId !== expected.fragmentId
        || actual.anchorId !== expected.anchorId
        || actual.fragmentIndex !== expected.fragmentIndex
        || actual.pageNumber !== expected.pageNumber
        || actual.story !== expected.story
        || actual.inTableCell !== expected.inTableCell
        || Math.abs(actual.geometry.x - expected.geometry.x) > 0.1
        || Math.abs(actual.geometry.y - expected.geometry.y) > 0.1
        || Math.abs(actual.geometry.width - expected.geometry.width) > 0.1
        || Math.abs(actual.geometry.height - expected.geometry.height) > 0.1) {
        throw new Error(`offline PageMap fragment changed at index ${index}`);
      }
    }
  } finally {
    frame.remove();
  }
}

function reportBase(
  manifest: PackageManifest,
  sourceBytes: Uint8Array,
  derivedManifest: PackageManifest | undefined,
  derivedBytes: Uint8Array | undefined,
  options: NormalizedOptions,
  layoutDigest: string,
  runtimePolicyDigest: string,
  state: ExecutionState,
  fontIdentity?: FontConfigurationIdentity,
): RenderReportBase {
  return {
    schema: REPORT_SCHEMA,
    schemaVersion: 3,
    source: {
      rawPackageBytesDigest: manifest.rawPackageBytesDigest.value.toLowerCase(),
      byteLength: sourceBytes.byteLength,
      documentVersion: options.documentVersion,
    },
    ...(derivedManifest && derivedBytes ? {
      derivedProfileSource: {
        rawPackageBytesDigest: derivedManifest.rawPackageBytesDigest.value,
        byteLength: derivedBytes.byteLength,
      },
    } : {}),
    options: {
      reviewProfile: options.reviewProfile,
      reviewProfileAlreadyApplied: options.reviewProfileAlreadyApplied,
      commentProfile: options.commentProfile,
      title: options.title,
      outputs: ["html"],
      layoutDigest,
      runtimePolicyDigest,
      policy: {
        unsupportedContent: options.unsupportedContent,
        strictFonts: options.strictFonts,
        timeoutMs: options.timeoutMs,
        limits: { ...options.limits },
      },
    },
    readiness: state.readiness.map((outcome) => ({
      ...outcome,
      pending: [...outcome.pending],
      ...(outcome.diagnostics
        ? { diagnostics: outcome.diagnostics.map((diagnostic) => ({ ...diagnostic })) }
        : {}),
    })),
    ...(state.fontIdentity ? { fontIdentity: { ...state.fontIdentity } } : {}),
    fonts: state.fonts.map((font) => ({
      ...font,
      requestedFamilies: [...font.requestedFamilies],
      requestedFamilyKinds: [...font.requestedFamilyKinds],
      ...(font.licenseEvidence ? { licenseEvidence: { ...font.licenseEvidence } } : {}),
    })),
    fontReadiness: state.fontReadiness.map((font) => ({ ...font })),
    resources: state.resources.map((resource) => ({ ...resource })),
    unsupportedContent: state.unsupportedContent.map((outcome) => ({ ...outcome })),
    warnings: state.warnings.map((warning) => ({ ...warning })),
    ...(fontIdentity ? { fontIdentity: { ...fontIdentity } } : {}),
  };
}

function failureReport(
  manifest: PackageManifest,
  sourceBytes: Uint8Array,
  derivedManifest: PackageManifest | undefined,
  derivedBytes: Uint8Array | undefined,
  options: NormalizedOptions,
  layoutDigest: string,
  runtimePolicyDigest: string,
  rendererFingerprint: string | undefined,
  observed: ExportRuntimeObservedFacts | undefined,
  fontIdentity: FontConfigurationIdentity | undefined,
  state: ExecutionState,
  error: DocxodusExportError,
  pageMapWasMaterialized: boolean,
  htmlWasMaterialized: boolean,
  pages?: CompleteRenderReport["pages"],
): FailedRenderReport {
  const unavailable: FailedRenderReport["unavailable"] = [];
  if (!rendererFingerprint) unavailable.push({
    field: "environment.rendererFingerprint",
    reasonCode: "notReached",
    detail: `Failure occurred during ${error.phase} before renderer identity completed.`,
  });
  unavailable.push(
    {
      field: "bindings.pageMapDigest",
      reasonCode: pageMapWasMaterialized ? "discardedOnFailure" : "notReached",
      detail: pageMapWasMaterialized
        ? "The materialized PageMap is discarded because the render did not complete."
        : "PageMap materialization was not reached.",
    },
    {
      field: "bindings.htmlDigest",
      reasonCode: htmlWasMaterialized ? "discardedOnFailure" : "notReached",
      detail: htmlWasMaterialized
        ? "The selected HTML payload is discarded because the render did not complete."
        : "Standalone HTML materialization was not reached.",
    },
    {
      field: "bindings.pdfDigest",
      reasonCode: "notRequested",
      detail: "PDF output was not selected by the browser materializer.",
    },
  );
  return {
    ...reportBase(
      manifest,
      sourceBytes,
      derivedManifest,
      derivedBytes,
      options,
      layoutDigest,
      runtimePolicyDigest,
      state,
      fontIdentity,
    ),
    status: "failed",
    failure: {
      code: error.code,
      severity: "error",
      phase: error.phase,
      message: error.message,
      remediation: error.remediation,
      ...(error.detail ? { detail: error.detail } : {}),
      ...(error.pending ? { pending: [...error.pending] } : {}),
      ...(error.partUri ? { partUri: error.partUri } : {}),
      ...(error.anchorId ? { anchorId: error.anchorId } : {}),
      ...(error.resource ? { resource: error.resource } : {}),
    },
    ...(rendererFingerprint && observed
      ? {
        environment: {
          rendererFingerprint,
          verification: "browserObserved" as const,
          fidelityTier: "unbaselined" as const,
          observed,
        },
      }
      : {}),
    ...(pages ? { partial: { pages } } : {}),
    unavailable,
  };
}

function ensureTerminalReadiness(state: ExecutionState, error: DocxodusExportError): void {
  const last = state.readiness.at(-1);
  if (last && last.status !== "complete") return;
  state.readiness.push({
    phase: error.phase,
    status: error.code === "operation_cancelled" ? "cancelled" : "failed",
    elapsedMs: 0,
    pending: error.pending ? [...error.pending] : [],
  });
}

function asExportError(error: unknown, phase: ExportPhase): DocxodusExportError {
  if (error instanceof DocxodusExportError) return error;
  const message = error instanceof Error ? error.message : String(error);
  const code: DocxodusExportErrorCode = phase === "pagination"
    || phase === "running_story_placement"
    || phase === "page_tree_stability"
    ? "pagination_failure"
    : phase === "output_verification"
      ? "output_verification_failure"
      : "conversion_failure";
  return new DocxodusExportError(code, phase, message,
    "Inspect the render report and source document, then retry with supported content.", { cause: error });
}

/**
 * Convert DOCX bytes into a complete offline HTML document containing only the
 * finalized fixed page tree and its inline resources.
 */
export async function convertDocxToPaginatedHtml(
  document: File | Uint8Array,
  requestedOptions: PaginatedHtmlOptions,
): Promise<PaginatedHtmlResult> {
  // Both snapshots happen before the first await. Later caller mutations cannot
  // split the verified package identity from the bytes transferred to WASM.
  const options = normalizeOptions(requestedOptions);
  const startedAt = monotonicNow();
  const ownedOperationAbort = new AbortController();
  const sourcePromise = ownedBytes(
    document,
    options.limits.compressedDocxBytes,
    options.signal,
  );
  const state: ExecutionState = {
    startedAt,
    deadline: startedAt + options.timeoutMs,
    phase: "input_validation",
    readiness: [],
    warnings: [],
    fonts: [],
    fontReadiness: [],
    resources: [],
    unsupportedContent: [],
    limits: options.limits,
    signal: options.signal,
  };
  let sourceBytes: Uint8Array | undefined;
  let renderBytes: Uint8Array | undefined;
  let derivedBytes: Uint8Array | undefined;
  let worker: WorkerDocxodus | undefined;
  let frame: HTMLIFrameElement | undefined;
  let manifest: PackageManifest | undefined;
  let derivedManifest: PackageManifest | undefined;
  let runtimeAssets: RuntimeAssetIdentity | undefined;
  let runtimeVersion: VersionInfo | undefined;
  let layoutDigest = "";
  let runtimePolicyDigest = "";
  let rendererFingerprint: string | undefined;
  let observed: ExportRuntimeObservedFacts | undefined;
  let fontIdentity: FontConfigurationIdentity | undefined;
  let pageMapWasMaterialized = false;
  let htmlWasMaterialized = false;
  let pagesForFailure: CompleteRenderReport["pages"] | undefined;

  try {
    sourceBytes = await runPhase(state, "input_validation", ["document bytes"], () => sourcePromise);
    if (sourceBytes.byteLength === 0) {
      fail("invalid_document", "input_validation", "The DOCX input is empty.",
        "Pass a non-empty OPC package.");
    }
    layoutDigest = await layoutDigestForOptions(options);

    runtimeAssets = await runPhase(state, "wasm_initialization", ["runtime asset graph"], () =>
      loadRuntimeAssetIdentity(options.wasmBasePath, ownedOperationAbort.signal));
    runtimePolicyDigest = await runtimePolicyDigestForOptions(options, runtimeAssets);
    worker = await runPhase(state, "wasm_initialization", ["WASM worker"], async () => {
      const created = await createWorkerDocxodus({
        wasmBasePath: options.wasmBasePath,
        signal: ownedOperationAbort.signal,
      });
      // runPhase may already have returned a timeout/cancellation while asynchronous
      // worker initialization was still in flight. Never orphan that late worker.
      if (ownedOperationAbort.signal.aborted
        || options.signal?.aborted
        || monotonicNow() >= state.deadline) {
        created.terminate();
      }
      return created;
    });
    runtimeVersion = await runPhase(state, "wasm_initialization", ["WASM runtime identity"], () =>
      worker!.getVersion());
    const manifestJson = await runPhase(state, "package_preflight", ["source package manifest"], () =>
      worker!.generatePackageManifestJson(sourceBytes!, inspectionLimits(options)));
    manifest = validatePackageManifestJson(manifestJson);
    await preflightManifest(manifest, sourceBytes, options, state, true);

    renderBytes = sourceBytes;
    if (options.reviewProfile !== "markup" && !options.reviewProfileAlreadyApplied) {
      derivedBytes = await runPhase(
        state,
        "package_preflight",
        [`${options.reviewProfile} review-profile projection`],
        async () => {
          try {
            return await worker!.projectReviewProfile(
              sourceBytes!,
              options.reviewProfile as "final" | "original",
              options.limits.compressedDocxBytes,
            );
          } catch (error) {
            if (String(error).includes("exceeds compressedDocxBytes")) {
              fail("resource_limit", "package_preflight",
                "The derived review-profile package exceeds compressedDocxBytes.",
                "Use a smaller document or remove revision history before export.");
            }
            throw error;
          }
        },
      );
      enforceLimit(derivedBytes.byteLength, options.limits.compressedDocxBytes,
        "compressedDocxBytes", "package_preflight");
      const derivedManifestJson = await runPhase(
        state,
        "package_preflight",
        ["derived package manifest"],
        () => worker!.generatePackageManifestJson(derivedBytes!, inspectionLimits(options)),
      );
      derivedManifest = validatePackageManifestJson(derivedManifestJson);
      await preflightManifest(derivedManifest, derivedBytes, options, state, false);
      if (derivedManifest.facts.revisions.total !== 0) {
        fail("conversion_failure", "package_preflight",
          `The ${options.reviewProfile} projection retained native tracked revisions.`,
          "Report the projection defect; derived bytes must be fully accepted or rejected exactly once.");
      }
      if (manifest.facts.revisions.total > 0
        && constantTimeDigestEqual(
          manifest.rawPackageBytesDigest.value,
          derivedManifest.rawPackageBytesDigest.value,
        )) {
        fail("conversion_failure", "package_preflight",
          `The ${options.reviewProfile} projection did not change revision-bearing package bytes.`,
          "Report the projection defect; changed review state requires a distinct derived identity.");
      }
      renderBytes = derivedBytes;
    }

    const convertedHtml = await runPhase(state, "docx_conversion", ["WASM conversion"], async () => {
      try {
        return await worker!.convertDocxToHtml(
          renderBytes!,
          conversionOptions(options),
          options.limits.htmlOutputBytes,
        );
      } catch (error) {
        if (String(error).includes("exceeds htmlOutputBytes")) {
          fail("resource_limit", "docx_conversion",
            "Converted HTML exceeds htmlOutputBytes before main-thread materialization.",
            "Use a smaller document or lower-complexity conversion profile.");
        }
        throw error;
      }
    });
    preflightConvertedHtml(convertedHtml, options);
    const attemptCheckpoint = checkpointAttemptState(state);
    let finalized: FinalizedTree | undefined;
    let stableReferenceSignature: string | undefined;
    let fontResolutionReferenceDigest: string | undefined;
    let pageTreeRetries = 0;
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        frame = await createIsolatedFrame(globalThis.document, state, bootstrapHtml(options.title));
        const renderDocument = frame.contentDocument!;
        sanitizeConvertedDocument(renderDocument, convertedHtml, state, options);
        // The converter hides its source staging tree from the viewer, but this
        // is the tree the paginator measures and clones. It is already offscreen;
        // expose only that renderer-owned root to computed-style inventory.
        // Author-hidden descendants remain hidden and are still excluded.
        const measurementStaging = renderDocument.getElementById("pagination-staging") as HTMLElement | null;
        if (measurementStaging) {
          measurementStaging.style.setProperty("visibility", "visible", "important");
          measurementStaging.style.setProperty("pointer-events", "none", "important");
        }
        inventoryConvertedContent(renderDocument, state, options);
        const readinessEvidenceCheckpoint = checkpointAttemptState(state);

        const readinessLimits = {
          fontRequests: options.limits.fontRequests,
          fontSampleCodePoints: options.limits.fontSampleCodePoints,
          visualResources: options.limits.automaticResources,
          domNodes: options.limits.domNodes,
          automaticResourceBytes: options.limits.automaticResourceBytes,
        };
        let fontTask: BrowserFontTask | undefined;
        await runPhase(state, "font_loading", () =>
          fontTask?.pending() ?? ["font:inventory"], async (signal) => {
          try {
            fontTask = createBrowserFontTask(renderDocument, options.fontResolver, options.limits);
            const result = await fontTask.wait(signal);
            if (fontResolutionReferenceDigest === undefined) {
              fontResolutionReferenceDigest = result.identity.resolutionDigest;
            } else if (fontResolutionReferenceDigest !== result.identity.resolutionDigest) {
              fail("resource_policy_failure", "font_loading",
                "The font resolver produced a different configuration across pristine layout attempts.",
                "Use an immutable resolver response for one export operation.", {
                  detail: `first=${fontResolutionReferenceDigest}; current=${result.identity.resolutionDigest}`,
                });
            }
            recordFontResolution(result, state, options);
          } catch (error) {
            rethrowBrowserFontError(error);
          }
        });
        const imageTask = documentImageReadiness(renderDocument, readinessLimits);
        const graphicTask = documentGraphicReadiness(renderDocument, readinessLimits);
        await runPhase(state, "image_decoding", imageTask.pending, async (signal) => {
          await admitPrintVisualResources(renderDocument, readinessLimits, signal);
          const probes = await imageTask.wait(signal);
          const images = Array.from(renderDocument.images);
          if (images.length !== probes.filter(({ source }) => source === "html-image").length) {
            throw new PrintReadinessError(
              "image_decoding",
              "The image inventory changed after its final readiness probe.",
              ["image-inventory"],
            );
          }
          recordImageReadiness(images, probes, state, options);
        });
        await runPhase(
          state,
          "chart_svg_materialization",
          graphicTask.pending,
          async (signal) => {
            const probes = await graphicTask.wait(signal);
            const graphics = graphicElements(renderDocument);
            if (graphics.length !== probes.filter(({ source }) => source === "graphic").length) {
              throw new PrintReadinessError(
                "chart_svg_materialization",
                "The graphic inventory changed after its final readiness probe.",
                ["graphic-inventory"],
              );
            }
            recordGraphicReadiness(graphics, probes, state, options);
          },
        );

        const staging = renderDocument.getElementById("pagination-staging") as HTMLElement | null;
        const container = renderDocument.getElementById("pagination-container") as HTMLElement | null;
        if (!staging || !container) {
          fail("conversion_failure", "docx_conversion",
            "Paginated conversion did not produce staging and page containers.",
            "Use the paginated converter contract and report the malformed conversion output.");
        }
        const engine = new PaginationEngine(staging, container, {
          scale: 1,
          cssPrefix: "page-",
          showPageNumbers: false,
          pageGap: 0,
          fragmentParagraphs: true,
          checkCancellation: () => {
            if (state.signal?.aborted) {
              fail("operation_cancelled", state.phase,
                `Export was cancelled during ${state.phase}.`,
                "Retry with a non-aborted signal.",
                { pending: ["page layout"] });
            }
            if (monotonicNow() >= state.deadline) {
              fail("readiness_timeout", state.phase,
                `Export timed out during ${state.phase}.`,
                "Increase timeoutMs or reduce document layout complexity.",
                { detail: "cooperative pagination checkpoint", pending: ["page layout"] });
            }
          },
          checkPageCount: (prospectivePageCount) => {
            enforceLimit(
              prospectivePageCount,
              options.limits.finalPages,
              "finalPages",
              "pagination",
            );
          },
        });
        const pagination = await runPhase(state, "pagination", ["page layout"], () => engine.paginate());
        const paginationOutcome = state.readiness[state.readiness.length - 1];
        if (paginationOutcome?.phase === "pagination" && paginationOutcome.status === "complete") {
          paginationOutcome.diagnostics = pagination.readiness.diagnostics.map((diagnostic) => ({
            ...diagnostic,
          }));
        }
        enforceLimit(pagination.totalPages, options.limits.finalPages, "finalPages", "pagination");
        const pages = pagination.pages.map((page) => page.element);
        if (pages.length === 0) {
          fail("pagination_failure", "pagination", "Pagination produced no pages.",
            "Verify that the DOCX has a renderable main document body.");
        }

        await runPhase(state, "running_story_placement", ["headers, footers, and notes"], () => {
          finalizePageTree(renderDocument, pages, state, options);
          freezeCssMotion(renderDocument);
          // Running-story placement changes the visible fragment set. This is
          // the sole identity-writing normalization pass; PageMap measurement
          // below remains read-only.
          engine.normalizePageMapFragmentIdentities();
          assertNoClippedContent(renderDocument);
        });
        countDomNodes(renderDocument, options.limits.domNodes, "running_story_placement");
        const automaticResources = automaticResourceCount(renderDocument);
        enforceLimit(automaticResources.count, options.limits.automaticResources,
          "automaticResources", "running_story_placement");
        enforceLimit(automaticResources.bytes, options.limits.automaticResourceBytes,
          "automaticResourceBytes", "running_story_placement");

        // Bind the report and attempt agreement to the exact finalized page DOM,
        // not merely the pre-pagination staging tree. Preserve deterministic
        // omitted-resource evidence for visuals already replaced by placeholders.
        const failedVisualEvidence = state.resources
          .slice(readinessEvidenceCheckpoint.resources)
          .filter((resource) => resource.readiness === "failed");
        state.fontReadiness.length = readinessEvidenceCheckpoint.fontReadiness;
        state.resources.length = readinessEvidenceCheckpoint.resources;
        state.resources.push(...failedVisualEvidence);
        const finalFontTask = documentFontReadiness(renderDocument, readinessLimits);
        const finalImageTask = documentImageReadiness(renderDocument, readinessLimits);
        const finalGraphicTask = documentGraphicReadiness(renderDocument, readinessLimits);
        await runPhase(state, "font_loading", finalFontTask.pending, async (signal) => {
          recordExactFontReadiness(await finalFontTask.wait(signal), state);
        });
        await runPhase(state, "image_decoding", finalImageTask.pending, async (signal) => {
          await admitPrintVisualResources(renderDocument, readinessLimits, signal);
          const probes = await finalImageTask.wait(signal);
          const images = Array.from(renderDocument.images);
          if (images.length !== probes.filter(({ source }) => source === "html-image").length) {
            throw new PrintReadinessError(
              "image_decoding",
              "The finalized image inventory changed after its readiness probe.",
              ["image-inventory"],
            );
          }
          recordImageReadiness(images, probes, state, options);
        });
        await runPhase(
          state,
          "chart_svg_materialization",
          finalGraphicTask.pending,
          async (signal) => {
            const probes = await finalGraphicTask.wait(signal);
            const graphics = graphicElements(renderDocument);
            if (graphics.length !== probes.filter(({ source }) => source === "graphic").length) {
              throw new PrintReadinessError(
                "chart_svg_materialization",
                "The finalized graphic inventory changed after its readiness probe.",
                ["graphic-inventory"],
              );
            }
            recordGraphicReadiness(graphics, probes, state, options);
          },
        );

        const stabilityTask = pageTreeReadiness(renderDocument, pages);
        const stableAttempt = await runPhase(
          state,
          "page_tree_stability",
          stabilityTask.pending,
          async (signal): Promise<{
            stability: PageTreeStabilityProbe;
            agreementSignature: string;
          }> => {
            try {
              const stability = await stabilityTask.wait(signal);
              if (signal.aborted) throw new DOMException("Print readiness was aborted", "AbortError");
              return {
                stability,
                agreementSignature: await pristineAttemptAgreementSignature(
                  stability.signature,
                  state,
                ),
              };
            } catch (error) {
              if (error instanceof PrintReadinessError) {
                throw new PageTreeInstabilityError(error.message, error.pending);
              }
              throw error;
            }
          },
        );
        const { agreementSignature } = stableAttempt;
        if (stableReferenceSignature === undefined) {
          stableReferenceSignature = agreementSignature;
          frame.remove();
          frame = undefined;
          restoreAttemptState(state, attemptCheckpoint);
          if (pageTreeRetries > 0) addPageTreeRetryWarning(state);
          continue;
        }
        if (stableReferenceSignature !== agreementSignature) {
          // A mismatch resets the reference: publication requires two
          // consecutive pristine-tree attempts with the same signature.
          stableReferenceSignature = agreementSignature;
          pageTreeRetries++;
          frame.remove();
          frame = undefined;
          restoreAttemptState(state, attemptCheckpoint);
          addPageTreeRetryWarning(state);
          continue;
        }
        finalized = { frame, document: renderDocument, engine, pages };
        break;
      } catch (error) {
        frame?.remove();
        frame = undefined;
        if (!(error instanceof PageTreeInstabilityError)) throw error;
        // An attempt that was internally unstable cannot participate in the
        // two-consecutive-pristine-layout agreement.
        stableReferenceSignature = undefined;
        pageTreeRetries++;
        restoreAttemptState(state, attemptCheckpoint);
        addPageTreeRetryWarning(state);
        if (attempt === 3) {
          fail("pagination_failure", "page_tree_stability", error.message,
            "Remove asynchronous layout inputs or report the source document to Docxodus.", {
              pending: [...error.pending],
            });
        }
      }
    }
    if (!finalized) {
      fail("pagination_failure", "page_tree_stability",
        "The final page tree did not produce two matching stable signatures within three attempts.",
        "Remove asynchronous layout inputs or report the source document to Docxodus.");
    }
    frame = finalized.frame;
    const { document: renderDocument, engine, pages } = finalized;
    if (!state.fontIdentity) {
      fail("resource_policy_failure", "output_verification",
        "The successful render is missing its font configuration identity.",
        "Report this exporter invariant failure.");
    }
    const renderer = await runPhase(state, "output_verification", ["renderer identity"], () =>
      rendererIdentity(
        renderDocument,
        layoutDigest,
        runtimePolicyDigest,
        runtimeAssets!,
        runtimeVersion!,
        state.fontIdentity!,
        state.fonts,
      ));
    rendererFingerprint = renderer.rendererFingerprint;
    observed = renderer.observed;
    fontIdentity = renderer.fontIdentity;
    const pageMap = await runPhase(state, "output_verification", ["PageMap geometry"], () =>
      engine.materializePageMap(options.documentVersion, rendererFingerprint!));
    pageMapWasMaterialized = true;
    if (pageMap.documentVersion !== options.documentVersion
      || pageMap.rendererFingerprint !== rendererFingerprint
      || pageMap.pages.length !== pages.length) {
      fail("output_verification_failure", "output_verification",
        "PageMap identity does not match the finalized render.",
        "Report the materializer defect; artifact identity fields must agree exactly.");
    }
    const pageNumbers = new Set<number>();
    for (const [index, page] of pageMap.pages.entries()) {
      if (page.pageNumber !== index + 1 || pageNumbers.has(page.pageNumber)
        || !Number.isInteger(page.pageInSection) || page.pageInSection < 1
        || page.sectionIndex === undefined
        || !Number.isInteger(page.sectionIndex) || page.sectionIndex < 0
        || typeof page.pageName !== "string" || page.pageName.length === 0
        || !Number.isFinite(page.width) || page.width <= 0
        || !Number.isFinite(page.height) || page.height <= 0) {
        fail("output_verification_failure", "output_verification",
          `PageMap page ${index} violates the finalized page invariants.`,
          "Report the materializer defect; page identities and geometry must be finite and contiguous.");
      }
      pageNumbers.add(page.pageNumber);
    }
    const fragmentIds = new Set<string>();
    const nextFragmentIndex = new Map<string, number>();
    for (const [index, fragment] of pageMap.fragments.entries()) {
      const page = pageMap.pages[fragment.pageNumber - 1];
      const expectedIndex = nextFragmentIndex.get(fragment.anchorId) ?? 0;
      const geometry = fragment.geometry;
      if (!page || fragmentIds.has(fragment.fragmentId)
        || fragment.fragmentId !==
          `p${fragment.pageNumber}-f${fragment.fragmentIndex}-${fragment.anchorId}`
        || fragment.fragmentIndex !== expectedIndex
        || !Object.values(geometry).every((value) => Number.isFinite(value) && value >= 0)
        || geometry.x + geometry.width > page.width + 0.1
        || geometry.y + geometry.height > page.height + 0.1) {
        fail("output_verification_failure", "output_verification",
          `PageMap fragment ${index} violates the visible-fragment invariants.`,
          "Report the materializer defect; fragments must be unique, ordered, visible, and page-bounded.");
      }
      fragmentIds.add(fragment.fragmentId);
      nextFragmentIndex.set(fragment.anchorId, expectedIndex + 1);
    }
    pagesForFailure = reportPages(pageMap);
    const pageMapJson = canonicalJson(pageMap);
    enforceLimit(utf8ByteLength(pageMapJson), options.limits.pageMapOutputBytes,
      "pageMapOutputBytes", "output_verification");
    const pageMapDigest = await runPhase(state, "output_verification", ["PageMap digest"], () =>
      sha256(utf8Bytes(pageMapJson)));
    const html = serializeDocument(renderDocument);
    htmlWasMaterialized = true;
    enforceLimit(utf8ByteLength(html), options.limits.htmlOutputBytes,
      "htmlOutputBytes", "output_verification");
    await runPhase(state, "output_verification", ["offline reopen"], (signal) =>
      verifyOfflineReopen(globalThis.document, html, pageMap, state, signal));
    const htmlDigest = await runPhase(state, "output_verification", ["HTML digest"], () =>
      sha256(utf8Bytes(html)));

    const report: CompleteRenderReport = {
      ...reportBase(
        manifest,
        sourceBytes,
        derivedManifest,
        derivedBytes,
        options,
        layoutDigest,
        runtimePolicyDigest,
        state,
        fontIdentity,
      ),
      status: "complete",
      fontIdentity,
      environment: {
        rendererFingerprint,
        verification: "browserObserved",
        fidelityTier: "unbaselined",
        observed,
      },
      pages: pagesForFailure,
      bindings: {
        pageMapDigest,
        htmlDigest,
        artifactRequestIds: [],
      },
    };
    enforceLimit(utf8ByteLength(canonicalJson(report)), options.limits.renderReportOutputBytes,
      "renderReportOutputBytes", "output_verification");
    return {
      html,
      pageCount: pages.length,
      pageMap,
      renderReport: report,
      warnings: report.warnings,
      rendererFingerprint,
    };
  } catch (error) {
    const resolved = asExportError(error, state.phase);
    ensureTerminalReadiness(state, resolved);
    if (manifest && sourceBytes) {
      resolved.report = failureReport(
        manifest,
        sourceBytes,
        derivedManifest,
        derivedBytes,
        options,
        layoutDigest,
        runtimePolicyDigest,
        rendererFingerprint,
        observed,
        fontIdentity,
        state,
        resolved,
        pageMapWasMaterialized,
        htmlWasMaterialized,
        pagesForFailure,
      );
      if (utf8ByteLength(canonicalJson(resolved.report)) > options.limits.renderReportOutputBytes) {
        // Preserve the primary structured error while withholding an artifact
        // that would itself exceed the versioned report ceiling.
        resolved.report = undefined;
      }
    }
    throw resolved;
  } finally {
    state.phase = "cleanup";
    ownedOperationAbort.abort();
    frame?.remove();
    worker?.terminate();
  }
}

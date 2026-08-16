/**
 * UI-free browser materializer for deterministic, standalone paginated HTML.
 *
 * Conversion happens in the existing WASM worker. Layout happens in a private,
 * attached browsing context so no caller styles, globals, or editor state can
 * influence the artifact.
 */

import limitsContractJson from "./export-resource-limits-v1.json";
import { PaginationEngine, type PageMap } from "./pagination.js";
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
  timeoutMs?: number;
  limits?: Partial<ExportResourceLimits>;
  signal?: AbortSignal;
  /** Reserved browser-only resolver boundary implemented by issue #442. */
  fontResolver?: unknown;
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
}

export interface FontResolution {
  requestedFamily: string;
  requestedFamilyStack?: string[];
  resolvedFamily?: string;
  status: "resolved" | "substituted" | "missing" | "unverified";
  source: "browser" | "embedded" | "configured";
}

export interface FontConfigurationIdentity {
  schemaVersion: 1;
  digest: string;
  verification: "browserObserved" | "configured";
}

export interface ResourceOutcome {
  kind: "image" | "svg" | "chart" | "external_link";
  status: "embedded" | "inline" | "allowed_user_link" | "omitted";
  resource?: string;
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
  schema: "https://docxodus.dev/schemas/render/render-report/v1";
  schemaVersion: 1;
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
  fonts: FontResolution[];
  resources: ResourceOutcome[];
  unsupportedContent: UnsupportedContentOutcome[];
  warnings: RenderWarning[];
  fontIdentity?: FontConfigurationIdentity;
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
  executableSha256: string;
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
  fonts: FontResolution[];
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
  constructor(message: string) {
    super(message);
    this.name = "PageTreeInstabilityError";
  }
}

const REPORT_SCHEMA = "https://docxodus.dev/schemas/render/render-report/v1" as const;
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
  if (options.fontResolver !== undefined) {
    fail("unsupported_runtime", "font_loading",
      "The browser fontResolver contract is reserved but not implemented on this branch.",
      "Omit fontResolver until issue #442 supplies the versioned resolver.");
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

async function runPhase<T>(
  state: ExecutionState,
  phase: ExportPhase,
  pending: string[],
  operation: () => T | Promise<T>,
): Promise<T> {
  state.phase = phase;
  const started = monotonicNow();
  if (state.signal?.aborted) {
    fail("operation_cancelled", phase, `Export was cancelled during ${phase}.`,
      "Retry with a non-aborted signal.", { pending });
  }
  const remaining = state.deadline - monotonicNow();
  if (remaining <= 0) {
    fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
      "Increase timeoutMs or remove the pending resource.", { pending });
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abortListener: (() => void) | undefined;
  try {
    const timeout = new Promise<never>((_, reject) => {
      timer = setTimeout(() => reject(new DocxodusExportError(
        "readiness_timeout",
        phase,
        `Export timed out during ${phase}.`,
        "Increase timeoutMs or remove the pending resource.",
        { pending },
      )), remaining);
    });
    const cancellation = new Promise<never>((_, reject) => {
      if (!state.signal) return;
      abortListener = () => reject(new DocxodusExportError(
        "operation_cancelled",
        phase,
        `Export was cancelled during ${phase}.`,
        "Retry with a non-aborted signal.",
        { pending },
      ));
      state.signal.addEventListener("abort", abortListener, { once: true });
    });
    const result = await Promise.race([Promise.resolve().then(operation), timeout, cancellation]);
    // A synchronous DOM operation cannot be pre-empted by the timer because it
    // blocks the event loop. Reject it immediately after control returns; hot
    // pagination loops also invoke the cooperative checkpoint below.
    if (state.signal?.aborted) {
      fail("operation_cancelled", phase, `Export was cancelled during ${phase}.`,
        "Retry with a non-aborted signal.", { pending });
    }
    if (monotonicNow() >= state.deadline) {
      fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
        "Increase timeoutMs or remove the pending resource.", { pending });
    }
    state.readiness.push({
      phase,
      status: "complete",
      elapsedMs: Math.max(0, monotonicNow() - started),
      pending: [],
    });
    return result;
  } catch (error) {
    state.readiness.push({
      phase,
      status: error instanceof DocxodusExportError && error.code === "operation_cancelled"
        ? "cancelled"
        : "failed",
      elapsedMs: Math.max(0, monotonicNow() - started),
      pending: [...pending],
    });
    throw error;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (abortListener && state.signal) state.signal.removeEventListener("abort", abortListener);
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

function addFont(state: ExecutionState, font: FontResolution): void {
  enforceLimit(state.fonts.length + 1, state.limits.fontRequests, "fontRequests", "font_loading");
  enforceDiagnosticAdmission(state, 1);
  state.fonts.push(font);
}

function enforceDiagnosticAdmission(state: ExecutionState, additional: number): void {
  const current = state.warnings.length + state.resources.length
    + state.unsupportedContent.length + state.fonts.length;
  if (current + additional > state.limits.renderDiagnostics) {
    fail("resource_limit", state.phase,
      `renderDiagnostics limit exceeded (${current + additional} > ${state.limits.renderDiagnostics}).`,
      "Use a smaller document or a versioned deployment policy with a higher diagnostic ceiling.");
  }
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

const STANDALONE_DATA_MEDIA_TYPES = new Set([
  "image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp", "image/tiff",
  "image/x-icon", "font/woff", "font/woff2", "font/ttf", "font/otf",
  "application/font-woff", "application/vnd.ms-fontobject",
]);

function dataUrlInfo(value: string): { mediaType: string; byteLength: number } | undefined {
  if (!value.startsWith("data:")) return undefined;
  const comma = value.indexOf(",");
  if (comma < 0) return undefined;
  const metadata = value.slice(5, comma);
  const segments = metadata.split(";");
  const mediaType = (segments.shift() ?? "").toLowerCase();
  if (mediaType === "" && value === "data:,") return { mediaType: "", byteLength: 0 };
  if (!STANDALONE_DATA_MEDIA_TYPES.has(mediaType)) return undefined;
  if (segments.some((segment) => segment.toLowerCase() !== "base64")) return undefined;
  if (!segments.some((segment) => segment.toLowerCase() === "base64")) return undefined;
  const payload = value.slice(comma + 1);
  if (!/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(payload)) {
    return undefined;
  }
  const padding = payload.endsWith("==") ? 2 : payload.endsWith("=") ? 1 : 0;
  return { mediaType, byteLength: payload.length / 4 * 3 - padding };
}

function automaticUrlAllowed(value: string, allowFragment = false): boolean {
  const trimmed = value.trim();
  return trimmed === "" || dataUrlInfo(trimmed) !== undefined
    || (allowFragment && trimmed.startsWith("#"));
}

function standaloneSrcsetAllowed(value: string): boolean {
  // Fail closed to one self-contained candidate. A loose comma split would
  // misparse the comma that is part of every data URL and could retain a later
  // network candidate.
  const match = /^\s*(data:\S+?)(?:\s+(?:\d+(?:\.\d+)?x|\d+w))?\s*$/i.exec(value);
  return !!match && dataUrlInfo(match[1]) !== undefined;
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

interface CssSecurityToken {
  kind: "import" | "substitution" | "url";
  start: number;
  end: number;
  value: string;
}

function consumeCssEscape(source: string, start: number): { value: string; end: number } {
  let cursor = start + 1;
  if (cursor >= source.length) return { value: "\ufffd", end: cursor };
  if (source[cursor] === "\r" && source[cursor + 1] === "\n") {
    return { value: "", end: cursor + 2 };
  }
  if (source[cursor] === "\n" || source[cursor] === "\r" || source[cursor] === "\f") {
    return { value: "", end: cursor + 1 };
  }
  const hexStart = cursor;
  while (cursor < source.length && cursor - hexStart < 6 && /[0-9a-f]/i.test(source[cursor])) {
    cursor++;
  }
  if (cursor > hexStart) {
    const point = Number.parseInt(source.slice(hexStart, cursor), 16);
    if (/\s/.test(source[cursor] ?? "")) {
      if (source[cursor] === "\r" && source[cursor + 1] === "\n") cursor += 2;
      else cursor++;
    }
    return {
      value: point === 0 || point > 0x10ffff || (point >= 0xd800 && point <= 0xdfff)
        ? "\ufffd"
        : String.fromCodePoint(point),
      end: cursor,
    };
  }
  return { value: source[cursor], end: cursor + 1 };
}

function consumeCssName(source: string, start: number): { value: string; end: number } {
  let value = "";
  let cursor = start;
  while (cursor < source.length) {
    const character = source[cursor];
    if (character === "\\") {
      const escape = consumeCssEscape(source, cursor);
      value += escape.value;
      cursor = escape.end;
    } else if (/[a-z0-9_-]/i.test(character) || character.charCodeAt(0) >= 0x80) {
      value += character;
      cursor++;
    } else {
      break;
    }
  }
  return { value, end: cursor };
}

function consumeCssComment(source: string, start: number): number {
  const end = source.indexOf("*/", start + 2);
  return end < 0 ? source.length : end + 2;
}

function consumeCssString(source: string, start: number): number {
  const quote = source[start];
  let cursor = start + 1;
  while (cursor < source.length) {
    if (source[cursor] === quote) return cursor + 1;
    // CSS bad-string recovery ends at an unescaped newline and resumes tokenization
    // afterward. Swallowing the remainder would hide later declarations from the audit.
    if (source[cursor] === "\n" || source[cursor] === "\r" || source[cursor] === "\f") {
      return cursor;
    }
    if (source[cursor] === "\\") cursor = consumeCssEscape(source, cursor).end;
    else cursor++;
  }
  return cursor;
}

function consumeCssFunction(source: string, openParenthesis: number): number {
  let depth = 0;
  let cursor = openParenthesis;
  while (cursor < source.length) {
    if (source.startsWith("/*", cursor)) {
      cursor = consumeCssComment(source, cursor);
    } else if (source[cursor] === "\"" || source[cursor] === "'") {
      cursor = consumeCssString(source, cursor);
    } else if (source[cursor] === "\\") {
      cursor = consumeCssEscape(source, cursor).end;
    } else if (source[cursor] === "(") {
      depth++;
      cursor++;
    } else if (source[cursor] === ")") {
      depth--;
      cursor++;
      if (depth === 0) return cursor;
    } else {
      cursor++;
    }
  }
  return cursor;
}

function decodeCssEscapedText(source: string, stripComments: boolean): string {
  let decoded = "";
  for (let cursor = 0; cursor < source.length;) {
    if (stripComments && source.startsWith("/*", cursor)) {
      cursor = consumeCssComment(source, cursor);
    } else if (source[cursor] === "\\") {
      const escape = consumeCssEscape(source, cursor);
      decoded += escape.value;
      cursor = escape.end;
    } else {
      decoded += source[cursor++];
    }
  }
  return decoded;
}

function decodeCssUrlComponent(source: string): string {
  const trimmed = source.trim();
  if (trimmed.length >= 2 && (trimmed[0] === "\"" || trimmed[0] === "'")
    && trimmed.at(-1) === trimmed[0]) {
    // Comment delimiters are ordinary payload bytes inside a CSS string.
    return decodeCssEscapedText(trimmed.slice(1, -1), false);
  }
  return decodeCssEscapedText(trimmed, true).trim();
}

/** Tokenize the security-relevant CSS grammar while honoring escapes, strings, and comments. */
function cssSecurityTokens(css: string): CssSecurityToken[] {
  const tokens: CssSecurityToken[] = [];
  const functionStack: string[] = [];
  for (let cursor = 0; cursor < css.length;) {
    if (css.startsWith("/*", cursor)) {
      cursor = consumeCssComment(css, cursor);
      continue;
    }
    if (css[cursor] === "\"" || css[cursor] === "'") {
      const end = consumeCssString(css, cursor);
      const context = functionStack.at(-1);
      const isImageSource = context === "image-set" || context === "-webkit-image-set";
      if (isImageSource
        && end > cursor && css[end - 1] === css[cursor]) {
        tokens.push({
          kind: "url",
          start: cursor,
          end,
          value: decodeCssEscapedText(css.slice(cursor + 1, end - 1), false),
        });
      }
      cursor = Math.max(end, cursor + 1);
      continue;
    }
    if (css[cursor] === "@") {
      const name = consumeCssName(css, cursor + 1);
      if (name.value.toLowerCase() === "import") {
        let end = name.end;
        let depth = 0;
        while (end < css.length) {
          if (css.startsWith("/*", end)) end = consumeCssComment(css, end);
          else if (css[end] === "\"" || css[end] === "'") end = consumeCssString(css, end);
          else if (css[end] === "(") { depth++; end++; }
          else if (css[end] === ")") { depth = Math.max(0, depth - 1); end++; }
          else if (css[end] === ";" && depth === 0) { end++; break; }
          else end++;
        }
        tokens.push({ kind: "import", start: cursor, end, value: css.slice(cursor, end) });
        cursor = end;
        continue;
      }
    }
    if (css[cursor] === "\\" || /[a-z_-]/i.test(css[cursor]) || css.charCodeAt(cursor) >= 0x80) {
      const name = consumeCssName(css, cursor);
      if (name.value.toLowerCase() === "url" && css[name.end] === "(") {
        let end = name.end + 1;
        while (end < css.length) {
          if (css.startsWith("/*", end)) end = consumeCssComment(css, end);
          else if (css[end] === "\"" || css[end] === "'") end = consumeCssString(css, end);
          else if (css[end] === "\\") end = consumeCssEscape(css, end).end;
          else if (css[end++] === ")") break;
        }
        const innerEnd = css[end - 1] === ")" ? end - 1 : end;
        tokens.push({
          kind: "url",
          start: cursor,
          end,
          value: decodeCssUrlComponent(css.slice(name.end + 1, innerEnd)),
        });
        cursor = end;
        continue;
      }
      if (css[name.end] === "(") {
        const functionName = name.value.toLowerCase();
        const isImageSource = functionStack.some(
          (context) => context === "image-set" || context === "-webkit-image-set",
        );
        if (isImageSource
          && (functionName === "var" || functionName === "env" || functionName === "if")) {
          // A substitution can resolve to a quoted image-set source, including a
          // custom property declared elsewhere. It cannot be proven standalone
          // from the authored token stream, so remove the entire substitution.
          const end = consumeCssFunction(css, name.end);
          tokens.push({
            kind: "substitution",
            start: cursor,
            end,
            value: css.slice(cursor, end),
          });
          cursor = end;
          continue;
        }
        functionStack.push(functionName);
        cursor = name.end + 1;
        continue;
      }
      cursor = Math.max(name.end, cursor + 1);
      continue;
    }
    if (css[cursor] === "(") {
      functionStack.push("");
      cursor++;
      continue;
    }
    if (css[cursor] === ")") {
      functionStack.pop();
      cursor++;
      continue;
    }
    cursor++;
  }
  return tokens;
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
  target.body.replaceChildren(...Array.from(parsed.body.childNodes, (node) => target.importNode(node, true)));
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

  for (const image of Array.from(document.querySelectorAll<HTMLImageElement>("img"))) {
    const source = image.getAttribute("src") ?? "";
    const metadata = /^data:([^;,]+)/i.exec(source);
    addResource(state, {
      kind: "image",
      status: "embedded",
      resource: image.alt || undefined,
      mediaType: metadata?.[1],
      byteLength: estimateDataUrlBytes(source),
    });
  }
  for (const svg of Array.from(document.querySelectorAll<SVGSVGElement>("svg"))) {
    addResource(state, {
      kind: svg.classList.contains("chart") || svg.closest("[class*='chart']") ? "chart" : "svg",
      status: "inline",
    });
  }
}

async function awaitFonts(document: Document): Promise<void> {
  if (document.fonts) await document.fonts.ready;
}

function inventoryBrowserObservedFonts(document: Document, state: ExecutionState): void {
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");

  const candidates = new Set<Element>([
    document.body,
    ...Array.from(document.querySelectorAll<HTMLElement>("[data-source-anchor-id]")),
  ]);
  const families = new Set<string>();
  for (const element of candidates) {
    if (!element.textContent?.trim()) continue;
    const family = view.getComputedStyle(element).fontFamily.trim();
    if (family) families.add(family);
  }
  for (const requestedFamily of Array.from(families).sort(compareCodeUnits)) {
    addFont(state, {
      requestedFamily,
      status: "unverified",
      source: "browser",
    });
  }
  if (families.size > 0) {
    addWarning(state, {
      code: "font_environment_unverified",
      severity: "warning",
      phase: "font_loading",
      message: "The browser loaded the requested CSS font families, but their exact files and substitutions are not attestable.",
      remediation: "Use the verified font resolver from issue #442 when exact font identity is required.",
    });
  }
}

async function decodeImages(document: Document): Promise<void> {
  for (const image of Array.from(document.images)) {
    if (typeof image.decode === "function") await image.decode();
    if (!image.complete || image.naturalWidth <= 0 || image.naturalHeight <= 0) {
      throw new Error(`Image failed to decode${image.alt ? `: ${image.alt}` : ""}`);
    }
  }
}

function validateInlineSvg(document: Document): void {
  for (const svg of Array.from(document.querySelectorAll<SVGSVGElement>("svg"))) {
    if (!svg.hasAttribute("viewBox")
      && !(Number.parseFloat(svg.getAttribute("width") ?? "") > 0
        && Number.parseFloat(svg.getAttribute("height") ?? "") > 0)) {
      throw new Error("Inline SVG has neither a viewBox nor explicit dimensions");
    }
  }
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
    for (const element of Array.from(page.querySelectorAll<HTMLElement>("[id]"))) {
      const original = element.id;
      const occurrence = occurrences.get(original) ?? 0;
      occurrences.set(original, occurrence + 1);
      const resolved = occurrence === 0
        ? original
        : `${original}--page-${page.dataset.pageNumber ?? "0"}-${occurrence}`;
      element.id = resolved;
      if (!localTargets.has(original)) localTargets.set(original, resolved);
      if (!globalTargets.has(original)) globalTargets.set(original, resolved);
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
        "Split the oversized footnote paragraph; full continuation support is tracked by issue #489.");
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

function treeSignature(document: Document, pages: HTMLElement[]): string {
  const geometry = pages.map((page) => {
    const rect = page.getBoundingClientRect();
    return [
      page.dataset.pageNumber,
      page.dataset.sectionIndex,
      rect.width.toFixed(3),
      rect.height.toFixed(3),
      page.scrollWidth,
      page.scrollHeight,
    ];
  });
  const fragments = Array.from(
    document.querySelectorAll<HTMLElement>(".page-box [data-source-anchor-id]"),
    (element) => {
      const rect = element.getBoundingClientRect();
      const style = document.defaultView!.getComputedStyle(element);
      return [
        element.dataset.sourceAnchorId,
        element.dataset.pageNumber,
        element.dataset.fragmentIndex,
        rect.left.toFixed(3),
        rect.top.toFixed(3),
        rect.width.toFixed(3),
        rect.height.toFixed(3),
        style.display,
        style.visibility,
      ];
    },
  );
  return canonicalJson({
    fragments,
    geometry,
    nodes: document.querySelectorAll("*").length,
    textLength: document.body.textContent?.length ?? 0,
  });
}

async function animationFrame(document: Document): Promise<void> {
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  await new Promise<void>((resolve) => view.requestAnimationFrame(() => resolve()));
}

async function awaitStableTree(document: Document, pages: HTMLElement[]): Promise<string> {
  const view = document.defaultView as (Window & typeof globalThis) | null;
  if (!view) throw new Error("render document has no defaultView");

  let mutations = 0;
  let resizes = 0;
  const mutationObserver = new view.MutationObserver((records) => {
    mutations += records.length;
  });
  const resizeObserver = typeof view.ResizeObserver === "function"
    ? new view.ResizeObserver((records) => {
      resizes += records.length;
    })
    : undefined;
  mutationObserver.observe(document.documentElement, {
    attributes: true,
    characterData: true,
    childList: true,
    subtree: true,
  });
  for (const page of pages) resizeObserver?.observe(page);

  try {
    // Let initial ResizeObserver delivery and style/layout work settle before
    // beginning the contractual quiet interval.
    await animationFrame(document);
    await animationFrame(document);
    const first = treeSignature(document, pages);
    mutations = 0;
    resizes = 0;
    // The render frame is intentionally script-disabled; use the caller realm's
    // timer while continuing to observe and measure only the render realm.
    await new Promise<void>((resolve) => globalThis.setTimeout(resolve, 100));
    await animationFrame(document);
    await animationFrame(document);
    const second = treeSignature(document, pages);
    if (first !== second || mutations !== 0 || resizes !== 0) {
      throw new PageTreeInstabilityError(
        `Final page tree changed during the quiet interval (mutations=${mutations}, resizes=${resizes})`,
      );
    }
    return second;
  } finally {
    mutationObserver.disconnect();
    resizeObserver?.disconnect();
  }
}

function checkpointAttemptState(state: ExecutionState): AttemptStateCheckpoint {
  return {
    readiness: state.readiness.length,
    warnings: state.warnings.length,
    fonts: state.fonts.length,
    resources: state.resources.length,
    unsupportedContent: state.unsupportedContent.length,
  };
}

function restoreAttemptState(state: ExecutionState, checkpoint: AttemptStateCheckpoint): void {
  state.readiness.length = checkpoint.readiness;
  state.warnings.length = checkpoint.warnings;
  state.fonts.length = checkpoint.fonts;
  state.resources.length = checkpoint.resources;
  state.unsupportedContent.length = checkpoint.unsupportedContent;
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
  const fontDigest = await canonicalMaterialDigest(
    "docxodus:font-configuration:v1",
    { verification: "browserObserved", fonts: fontRecords },
  );
  const fontIdentity: FontConfigurationIdentity = {
    schemaVersion: 1,
    digest: fontDigest,
    verification: "browserObserved",
  };
  const observed = observedRuntimeFacts(document);
  const fingerprint = {
    contract: "docxodus-standalone-browser-v1",
    verification: "browserObserved",
    runtimeAssets,
    runtimeVersion,
    paginatorContractVersion: 1,
    pageMapSchemaVersion: 1,
    renderReportSchemaVersion: 1,
    layoutDigest,
    runtimePolicyDigest,
    fontConfigurationDigest: fontDigest,
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
): Promise<void> {
  const frame = await createIsolatedFrame(hostDocument, state, html, "output_verification");
  try {
    const reopened = frame.contentDocument!;
    const pages = Array.from(reopened.querySelectorAll<HTMLElement>(".page-box"));
    await awaitFonts(reopened);
    await decodeImages(reopened);
    validateInlineSvg(reopened);
    await awaitStableTree(reopened, pages);
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
    schemaVersion: 1,
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
    readiness: state.readiness.map((outcome) => ({ ...outcome, pending: [...outcome.pending] })),
    fonts: state.fonts.map((font) => ({ ...font })),
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
    if (options.strictFonts) {
      fail("unsupported_runtime", "font_loading",
        "Strict font verification is not yet available in the browser materializer.",
        "Use the browser-observed font mode until issue #442 lands.");
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
    let firstAttemptSignature: string | undefined;
    for (let attempt = 1; attempt <= 2; attempt++) {
      try {
        frame = await createIsolatedFrame(globalThis.document, state, bootstrapHtml(options.title));
        const renderDocument = frame.contentDocument!;
        sanitizeConvertedDocument(renderDocument, convertedHtml, state, options);
        inventoryConvertedContent(renderDocument, state, options);

        await runPhase(state, "font_loading", ["document.fonts.ready"], async () => {
          await awaitFonts(renderDocument);
          inventoryBrowserObservedFonts(renderDocument, state);
        });
        await runPhase(state, "image_decoding", ["embedded images"], () =>
          decodeImages(renderDocument));
        await runPhase(state, "chart_svg_materialization", ["inline SVG"], () =>
          validateInlineSvg(renderDocument));

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
        });
        const pagination = await runPhase(state, "pagination", ["page layout"], () => engine.paginate());
        enforceLimit(pagination.totalPages, options.limits.finalPages, "finalPages", "pagination");
        const pages = pagination.pages.map((page) => page.element);
        if (pages.length === 0) {
          fail("pagination_failure", "pagination", "Pagination produced no pages.",
            "Verify that the DOCX has a renderable main document body.");
        }

        await runPhase(state, "running_story_placement", ["headers, footers, and notes"], () => {
          finalizePageTree(renderDocument, pages, state, options);
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

        const stableSignature = await runPhase(state, "page_tree_stability", ["fixed page tree"], () =>
          awaitStableTree(renderDocument, pages));
        if (attempt === 1) {
          firstAttemptSignature = stableSignature;
          frame.remove();
          frame = undefined;
          restoreAttemptState(state, attemptCheckpoint);
          continue;
        }
        if (!firstAttemptSignature || firstAttemptSignature !== stableSignature) {
          throw new PageTreeInstabilityError(
            "Two layouts created from the same pristine converted HTML produced different final page trees",
          );
        }
        finalized = { frame, document: renderDocument, engine, pages };
        break;
      } catch (error) {
        frame?.remove();
        frame = undefined;
        if (!(error instanceof PageTreeInstabilityError) || attempt === 2) throw error;
        restoreAttemptState(state, attemptCheckpoint);
        addWarning(state, {
          code: "page_tree_retry",
          severity: "warning",
          phase: "page_tree_stability",
          message: "The finalized page tree changed during its quiet interval and was rebuilt from pristine converted HTML.",
          remediation: "No action is required unless repeated exports fail page-tree stability.",
        });
      }
    }
    if (!finalized) {
      fail("pagination_failure", "page_tree_stability",
        "The final page tree did not stabilize within two attempts.",
        "Remove asynchronous layout inputs or report the source document to Docxodus.");
    }
    frame = finalized.frame;
    const { document: renderDocument, engine, pages } = finalized;
    const renderer = await runPhase(state, "output_verification", ["renderer identity"], () =>
      rendererIdentity(
        renderDocument,
        layoutDigest,
        runtimePolicyDigest,
        runtimeAssets!,
        runtimeVersion!,
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
    await runPhase(state, "output_verification", ["offline reopen"], () =>
      verifyOfflineReopen(globalThis.document, html, pageMap, state));
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

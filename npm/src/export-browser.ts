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
  type VersionInfo,
} from "./types.js";
import {
  documentFontReadiness,
  documentGraphicReadiness,
  documentImageReadiness,
  pageTreeReadiness,
  PrintReadinessError,
  type FontReadinessProbe,
  type PageTreeStabilityProbe,
  type VisualResourceProbe,
} from "./print-readiness.js";

export { awaitFinalPrintReadiness, PrintReadinessError } from "./print-readiness.js";
export type {
  FinalPrintReadinessResult,
  FontReadinessProbe,
  PageTreeStabilityProbe,
  PrintReadinessPhase,
  VisualResourceProbe,
} from "./print-readiness.js";

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
  | "invalid_document"
  | "conversion_failure"
  | "browser_launch_failure"
  | "resource_policy_failure"
  | "readiness_timeout"
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
  htmlOutputBytes: number;
  pdfOutputBytes: number;
  finalPages: number;
  domNodes: number;
  automaticResources: number;
  automaticResourceBytes: number;
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
  commentProfile: CommentProfile;
  title?: string;
  unsupportedContent?: UnsupportedContentPolicy;
  strictFonts?: boolean;
  timeoutMs?: number;
  limits?: Partial<ExportResourceLimits>;
  /** Override only when the package's `dist/wasm` directory is hosted elsewhere. */
  wasmBasePath?: string;
}

export interface RenderWarning {
  code: string;
  severity: "warning" | "error";
  phase: ExportPhase;
  message: string;
  remediation: string;
  partUri?: string;
  anchorId?: string;
  resource?: string;
}

export interface ReadinessOutcome {
  phase: ExportPhase;
  status: "complete" | "failed";
  elapsedMs: number;
  pending: string[];
  diagnostics?: PaginationDiagnostic[];
}

export interface FontResolution {
  requestedFamily: string;
  resolvedFamily?: string;
  status: "resolved" | "substituted" | "missing" | "unverified";
  source: "browser" | "embedded" | "configured";
}

export interface ResourceOutcome {
  kind: "image" | "svg" | "chart" | "external_link";
  status: "embedded" | "inline" | "allowed_user_link" | "omitted";
  readiness?: "complete" | "failed";
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
    commentProfile: CommentProfile;
    layoutDigest: string;
  };
  readiness: ReadinessOutcome[];
  fonts: FontResolution[];
  resources: ResourceOutcome[];
  unsupportedContent: UnsupportedContentOutcome[];
  warnings: RenderWarning[];
}

export type EnvironmentVerification = "nodeVerified" | "browserObserved" | "callerAttested";

export interface CompleteRenderReport extends RenderReportBase {
  status: "complete";
  environment: {
    rendererFingerprint: string;
    verification: EnvironmentVerification;
  };
  pages: Array<{
    pageNumber: number;
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
    phase: ExportPhase;
    message: string;
    remediation: string;
  };
  environment?: {
    rendererFingerprint?: string;
    verification: EnvironmentVerification;
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
    reason: string;
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
  readonly cause?: unknown;
  report?: FailedRenderReport;

  constructor(
    code: DocxodusExportErrorCode,
    phase: ExportPhase,
    message: string,
    remediation: string,
    options: { detail?: string; cause?: unknown; report?: FailedRenderReport } = {},
  ) {
    super(message);
    this.name = "DocxodusExportError";
    this.code = code;
    this.phase = phase;
    this.remediation = remediation;
    this.detail = options.detail;
    this.cause = options.cause;
    this.report = options.report;
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
    };
  }
}

interface NormalizedOptions {
  documentVersion: number;
  expectedSourceDigest?: string;
  reviewProfile: ReviewProfile;
  commentProfile: CommentProfile;
  title: string;
  unsupportedContent: UnsupportedContentPolicy;
  strictFonts: boolean;
  timeoutMs: number;
  limits: ExportResourceLimits;
  wasmBasePath: string;
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
const CSP = [
  "default-src 'none'",
  "base-uri 'none'",
  "connect-src 'none'",
  "font-src data:",
  "form-action 'none'",
  "frame-src 'none'",
  "img-src data:",
  "media-src data:",
  "navigate-to 'none'",
  "object-src 'none'",
  "script-src 'none'",
  "style-src 'unsafe-inline'",
].join("; ");

function fail(
  code: DocxodusExportErrorCode,
  phase: ExportPhase,
  message: string,
  remediation: string,
  options: { detail?: string; cause?: unknown } = {},
): never {
  throw new DocxodusExportError(code, phase, message, remediation, options);
}

function normalizeOptions(options: PaginatedHtmlOptions): NormalizedOptions {
  if (!options || typeof options !== "object") {
    fail("invalid_document", "input_validation", "Export options are required.",
      "Supply explicit reviewProfile and commentProfile values.");
  }
  if (!ALLOWED_REVIEW_PROFILES.has(options.reviewProfile)) {
    fail("invalid_document", "input_validation", "reviewProfile is invalid.",
      "Use final, original, or markup.");
  }
  if (!ALLOWED_COMMENT_PROFILES.has(options.commentProfile)) {
    fail("invalid_document", "input_validation", "commentProfile is invalid.",
      "Use hidden, inline, endnotes, or margin.");
  }
  const unsupportedContent = options.unsupportedContent ?? "warn";
  if (!ALLOWED_UNSUPPORTED_POLICIES.has(unsupportedContent)) {
    fail("invalid_document", "input_validation", "unsupportedContent is invalid.",
      "Use warn or strict.");
  }
  const documentVersion = options.documentVersion ?? 0;
  if (!Number.isSafeInteger(documentVersion) || documentVersion < 0) {
    fail("invalid_document", "input_validation",
      "documentVersion must be a non-negative JavaScript safe integer.",
      "Use a value between 0 and Number.MAX_SAFE_INTEGER.");
  }
  if (options.expectedSourceDigest !== undefined
    && !/^[0-9a-f]{64}$/i.test(options.expectedSourceDigest)) {
    fail("invalid_document", "input_validation", "expectedSourceDigest must be a SHA-256 hex digest.",
      "Supply exactly 64 hexadecimal characters.");
  }

  const limits = { ...DEFAULT_EXPORT_RESOURCE_LIMITS };
  for (const [name, value] of Object.entries(options.limits ?? {})) {
    if (!(name in limits)) {
      fail("invalid_document", "input_validation", `Unknown export limit: ${name}.`,
        "Use a key from ExportResourceLimits.");
    }
    const key = name as keyof ExportResourceLimits;
    if (!Number.isSafeInteger(value) || value <= 0) {
      fail("invalid_document", "input_validation", `Export limit ${name} must be a positive safe integer.`,
        "Supply a positive integer no greater than the published default.");
    }
    if (value > DEFAULT_EXPORT_RESOURCE_LIMITS[key]) {
      fail("invalid_document", "input_validation", `Export limit ${name} may only lower the default.`,
        `Use ${DEFAULT_EXPORT_RESOURCE_LIMITS[key]} or less.`);
    }
    limits[key] = value;
  }

  const timeoutMs = options.timeoutMs ?? LIMITS_CONTRACT.timeoutMs.default;
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs <= 0
    || timeoutMs > LIMITS_CONTRACT.timeoutMs.hardCeiling) {
    fail("invalid_document", "input_validation", "timeoutMs is outside the supported range.",
      `Use an integer from 1 through ${LIMITS_CONTRACT.timeoutMs.hardCeiling}.`);
  }

  return Object.freeze({
    documentVersion,
    expectedSourceDigest: options.expectedSourceDigest?.toLowerCase(),
    reviewProfile: options.reviewProfile,
    commentProfile: options.commentProfile,
    title: options.title ?? "Document",
    unsupportedContent,
    strictFonts: options.strictFonts ?? false,
    timeoutMs,
    limits: Object.freeze(limits),
    wasmBasePath: options.wasmBasePath ?? new URL("./wasm/", import.meta.url).href,
  });
}

async function ownedBytes(document: File | Uint8Array): Promise<Uint8Array> {
  if (document instanceof Uint8Array) return new Uint8Array(document);
  if (typeof File !== "undefined" && document instanceof File) {
    return new Uint8Array(await document.arrayBuffer());
  }
  fail("invalid_document", "input_validation", "document must be a File or Uint8Array.",
    "Pass immutable DOCX bytes or a browser File.");
}

function monotonicNow(): number {
  return globalThis.performance?.now() ?? Date.now();
}

async function runPhase<T>(
  state: ExecutionState,
  phase: ExportPhase,
  pendingResources: string[] | (() => string[]),
  operation: (signal: AbortSignal) => T | Promise<T>,
): Promise<T> {
  state.phase = phase;
  const started = monotonicNow();
  const pending = (): string[] => [
    ...(typeof pendingResources === "function" ? pendingResources() : pendingResources),
  ];
  const reportProgress = (status: "pending" | "complete" | "failed", resources: string[]): void => {
    const reporter = (globalThis as typeof globalThis & {
      __docxodusReadinessProgress?: (snapshot: {
        phase: ExportPhase;
        status: "pending" | "complete" | "failed";
        pending: string[];
      }) => unknown;
    }).__docxodusReadinessProgress;
    if (typeof reporter === "function") {
      try {
        void reporter({ phase, status, pending: [...resources] });
      } catch {
        // Progress is diagnostic-only; the readiness result remains authoritative.
      }
    }
  };
  const remaining = state.deadline - Date.now();
  if (remaining <= 0) {
    fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
      "Increase timeoutMs or remove the pending resource.", { detail: pending().join(", ") });
  }
  let timer: ReturnType<typeof setTimeout> | undefined;
  let timedOutPending: string[] | undefined;
  const controller = new AbortController();
  reportProgress("pending", pending());
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
          { detail: resources.join(", ") },
        ));
      }, remaining);
    });
    const result = await Promise.race([
      Promise.resolve().then(() => operation(controller.signal)),
      timeout,
    ]);
    // A synchronous DOM operation cannot be pre-empted by the timer because it
    // blocks the event loop. Reject it immediately after control returns; hot
    // pagination loops also invoke the cooperative checkpoint below.
    if (Date.now() >= state.deadline) {
      fail("readiness_timeout", phase, `Export timed out during ${phase}.`,
        "Increase timeoutMs or remove the pending resource.", { detail: pending().join(", ") });
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
    const resources = timedOutPending ?? pending();
    state.readiness.push({
      phase,
      status: "failed",
      elapsedMs: Math.max(0, monotonicNow() - started),
      pending: resources,
    });
    reportProgress("failed", resources);
    throw error;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    controller.abort();
  }
}

function utf8Bytes(value: string): Uint8Array {
  return TEXT_ENCODER.encode(value);
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

async function loadRuntimeAssetIdentity(wasmBasePath: string): Promise<RuntimeAssetIdentity> {
  const manifestUrl = new URL("./export-assets.json", import.meta.url);
  const response = await globalThis.fetch(manifestUrl, { cache: "no-store", credentials: "same-origin" });
  if (!response.ok) {
    fail("unsupported_runtime", "wasm_initialization",
      `The runtime asset graph could not be loaded (${response.status}).`,
      "Deploy export-assets.json beside the browser export bundle.");
  }
  const manifest = await response.json() as {
    schemaVersion?: unknown;
    packageVersion?: unknown;
    assets?: unknown;
  };
  if (manifest.schemaVersion !== 1 || typeof manifest.packageVersion !== "string"
    || !Array.isArray(manifest.assets) || manifest.assets.length === 0) {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph is malformed.",
      "Deploy the versioned export-assets.json generated with this bundle.");
  }
  const assets = manifest.assets.map((entry, index) => {
    if (!entry || typeof entry !== "object") {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset entry ${index} is malformed.`,
        "Regenerate the package runtime asset graph.");
    }
    const record = entry as Record<string, unknown>;
    if (typeof record.path !== "string" || typeof record.mediaType !== "string"
      || !Number.isSafeInteger(record.byteLength) || (record.byteLength as number) < 0
      || typeof record.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(record.sha256)) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset entry ${index} has invalid identity fields.`,
        "Regenerate the package runtime asset graph.");
    }
    return {
      path: record.path,
      mediaType: record.mediaType,
      byteLength: record.byteLength,
      sha256: record.sha256,
    };
  }).sort((left, right) => String(left.path).localeCompare(String(right.path)));
  const materializer = assets.find((entry) => entry.path === "./export-browser.bundle.js");
  if (!materializer) {
    fail("unsupported_runtime", "wasm_initialization",
      "The runtime asset graph does not identify the browser materializer.",
      "Regenerate export-assets.json from the complete runtime package.");
  }

  const materializerResponse = await globalThis.fetch(import.meta.url, {
    cache: "no-store",
    credentials: "same-origin",
  });
  if (!materializerResponse.ok) {
    fail("unsupported_runtime", "wasm_initialization",
      `The loaded browser materializer could not be verified (${materializerResponse.status}).`,
      "Serve the browser export bundle from a readable same-origin URL.");
  }
  const materializerDigest = await sha256(new Uint8Array(await materializerResponse.arrayBuffer()));
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
    });
    if (!assetResponse.ok) {
      fail("unsupported_runtime", "wasm_initialization",
        `Runtime asset ${asset.path} could not be verified (${assetResponse.status}).`,
        "Deploy every runtime asset named by export-assets.json.");
    }
    const bytes = new Uint8Array(await assetResponse.arrayBuffer());
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
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("Canonical JSON does not support non-finite numbers");
    return Object.is(value, -0) ? 0 : value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (typeof value === "object") {
    const result: Record<string, unknown> = {};
    for (const key of Object.keys(value as Record<string, unknown>).sort()) {
      const member = (value as Record<string, unknown>)[key];
      if (member !== undefined) result[key] = canonicalValue(member);
    }
    return result;
  }
  throw new TypeError(`Canonical JSON does not support ${typeof value}`);
}

/** Serialize report/schema values with recursively sorted object keys and no insignificant space. */
export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value));
}

function addWarning(state: ExecutionState, warning: RenderWarning): void {
  state.warnings.push(warning);
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

function preflightManifest(
  manifest: PackageManifest,
  bytes: Uint8Array,
  options: NormalizedOptions,
  state: ExecutionState,
): void {
  if (!manifest || manifest.schemaVersion !== 1 || !manifest.rawPackageBytesDigest?.value) {
    fail("invalid_document", "package_preflight", "DOCX preflight returned an invalid manifest.",
      "Validate the package with the Docxodus package verifier.");
  }
  const sourceDigest = manifest.rawPackageBytesDigest.value.toLowerCase();
  if (options.expectedSourceDigest && options.expectedSourceDigest !== sourceDigest) {
    fail("invalid_document", "package_preflight", "The source digest does not match expectedSourceDigest.",
      "Render the exact verified source bytes or update the expected digest.", {
        detail: `expected=${options.expectedSourceDigest}; actual=${sourceDigest}`,
      });
  }
  for (const finding of manifest.findings) {
    if (finding.severity === "info") continue;
    addWarning(state, {
      code: finding.code,
      severity: finding.severity,
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
  enforceLimit(
    manifest.entries.reduce((sum, entry) => sum + entry.size, 0),
    options.limits.expandedOpcBytes,
    "expandedOpcBytes",
    "package_preflight",
  );
  const largestXml = manifest.entries.reduce(
    (largest, entry) => entry.isXml ? Math.max(largest, entry.size) : largest,
    0,
  );
  enforceLimit(largestXml, options.limits.xmlPartBytes, "xmlPartBytes", "package_preflight");

  if (manifest.packageKind !== "opc" || !manifest.isValid || !manifest.facts.mainDocumentUri) {
    fail("invalid_document", "package_preflight",
      `The input is not a valid DOCX OPC package (${manifest.packageKind}).`,
      "Repair or decrypt the document before export.");
  }

  const externalAutomatic = manifest.relationships.filter((relationship) =>
    relationship.targetMode === "External"
    && !relationship.type.toLowerCase().endsWith("/hyperlink"));
  for (const relationship of externalAutomatic) {
    const warning: RenderWarning = {
      code: "external_automatic_resource_omitted",
      severity: options.unsupportedContent === "strict" ? "error" : "warning",
      phase: "package_preflight",
      message: "An external automatic resource is not fetched by standalone export.",
      remediation: "Embed the resource in the DOCX package before export.",
      partUri: relationship.ownerUri,
      resource: relationship.target,
    };
    addWarning(state, warning);
    state.resources.push({ kind: "image", status: "omitted", resource: relationship.target });
  }
  if (externalAutomatic.length > 0 && options.unsupportedContent === "strict") {
    fail("resource_policy_failure", "package_preflight",
      "Strict export rejected an external automatic resource.",
      "Embed all automatic resources or use unsupportedContent: warn.");
  }

  if (manifest.facts.isMacroEnabled) {
    const warning: RenderWarning = {
      code: "macro_content_not_exported",
      severity: options.unsupportedContent === "strict" ? "error" : "warning",
      phase: "package_preflight",
      message: "Macro content is not active or embedded in standalone HTML.",
      remediation: "Remove macros or use warn policy for a static visual export.",
    };
    addWarning(state, warning);
    if (options.unsupportedContent === "strict") {
      fail("resource_policy_failure", "package_preflight", warning.message, warning.remediation);
    }
  }
  if (manifest.facts.altChunkCount > 0) {
    const warning: RenderWarning = {
      code: "altchunk_not_supported",
      severity: options.unsupportedContent === "strict" ? "error" : "warning",
      phase: "package_preflight",
      message: "Arbitrary altChunk content is not a supported standalone export input.",
      remediation: "Materialize altChunk content into ordinary WordprocessingML before export.",
    };
    addWarning(state, warning);
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

function bootstrapHtml(title = "Document"): string {
  const safeTitle = title.replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;",
  })[character]!);
  return `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${CSP}"><title>${safeTitle}</title></head><body></body></html>`;
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

function automaticUrlAllowed(value: string): boolean {
  const trimmed = value.trim();
  return trimmed === "" || trimmed.startsWith("data:") || trimmed.startsWith("#");
}

function standaloneSrcsetAllowed(value: string): boolean {
  // Fail closed to one self-contained candidate. A loose comma split would
  // misparse the comma that is part of every data URL and could retain a later
  // network candidate.
  return /^\s*data:[^,\s]+,[^,\s]+(?:\s+(?:\d+(?:\.\d+)?x|\d+w))?\s*$/i.test(value);
}

const SVG_URL_PRESENTATION_ATTRIBUTES = new Set([
  "clip-path", "cursor", "fill", "filter", "marker", "marker-end", "marker-mid",
  "marker-start", "mask", "stroke",
]);

function estimateDataUrlBytes(value: string): number | undefined {
  const comma = value.indexOf(",");
  if (comma < 0 || !value.startsWith("data:")) return undefined;
  const metadata = value.slice(5, comma);
  const payload = value.slice(comma + 1);
  return /;base64(?:;|$)/i.test(metadata)
    ? Math.floor(payload.replace(/\s/g, "").length * 3 / 4)
    : utf8Bytes(decodeURIComponent(payload)).byteLength;
}

function policyWarning(
  state: ExecutionState,
  options: NormalizedOptions,
  warning: Omit<RenderWarning, "severity">,
): void {
  const resolved: RenderWarning = {
    ...warning,
    severity: options.unsupportedContent === "strict" ? "error" : "warning",
  };
  addWarning(state, resolved);
  if (options.unsupportedContent === "strict") {
    fail("resource_policy_failure", resolved.phase, resolved.message, resolved.remediation, {
      detail: resolved.resource,
    });
  }
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
      token.kind !== "url" || !automaticUrlAllowed(token.value));
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
      } else if (token.kind === "url" && automaticUrlAllowed(token.value)) {
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
            state.resources.push({ kind: "external_link", status: "allowed_user_link", resource: href });
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
      if ((name === "href" && element.localName !== "a") || urlAttributes.has(name)) {
        if (!automaticUrlAllowed(attribute.value)) {
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
    state.unsupportedContent.push(outcome);
    policyWarning(state, options, {
      code: "unsupported_content_placeholder",
      phase: "docx_conversion",
      message: `${outcome.contentType} is represented by a visible placeholder.`,
      remediation: "Replace it with a supported static representation for faithful export.",
      ...(outcome.anchorId ? { anchorId: outcome.anchorId } : {}),
    });
  }
}

function recordFontReadiness(
  probes: FontReadinessProbe[],
  state: ExecutionState,
  options: NormalizedOptions,
): void {
  for (const probe of probes) {
    state.fonts.push({
      requestedFamily: probe.requestedFamily,
      status: probe.available ? "unverified" : "missing",
      source: "browser",
    });
    if (!probe.available) {
      addWarning(state, {
        code: "font_unavailable",
        severity: options.strictFonts ? "error" : "warning",
        phase: "font_loading",
        message: `The browser could not load the required font family: ${probe.requestedFamily}.`,
        remediation: "Install or explicitly supply the required font before export.",
        resource: probe.requestedFamily,
      });
    }
  }
  if (probes.some(({ available }) => available)) {
    addWarning(state, {
      code: "font_environment_unverified",
      severity: options.strictFonts ? "error" : "warning",
      phase: "font_loading",
      message: "The browser loaded the requested CSS font families, but their exact files and substitutions are not attestable.",
      remediation: "Use the verified font resolver from issue #442 when exact font identity is required.",
    });
  }
  if (options.strictFonts && probes.length > 0) {
    fail("resource_policy_failure", "font_loading",
      "Strict font policy requires verified font files, but this runtime can only observe browser availability.",
      "Use the verified font resolver delivered by issue #442 or disable strictFonts.", {
        detail: probes.map(({ requestedFamily }) => requestedFamily).join(", "),
      });
  }
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
  probes.forEach((probe, index) => {
    const image = images[index];
    const source = image?.getAttribute("src") ?? "";
    const metadata = /^data:([^;,]+)/i.exec(source);
    state.resources.push({
      kind: "image",
      status: probe.status === "complete" ? "embedded" : "omitted",
      readiness: probe.status,
      resource: probe.resource,
      ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
      ...(probe.message ? { message: probe.message } : {}),
      mediaType: metadata?.[1],
      byteLength: estimateDataUrlBytes(source),
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
  probes.forEach((probe, index) => {
    state.resources.push({
      kind: probe.kind,
      status: probe.status === "complete" ? "inline" : "omitted",
      readiness: probe.status,
      resource: probe.resource,
      ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
      ...(probe.message ? { message: probe.message } : {}),
    });
    if (probe.status === "failed") {
      const element = elements[index];
      if (element) replaceFailedVisual(element, probe.resource);
      policyWarning(state, options, {
        code: "graphic_materialization_failed",
        phase: "chart_svg_materialization",
        message: `${probe.kind} content did not finish materializing: ${probe.resource}.`,
        remediation: "Replace the graphic or use a supported static SVG representation.",
        ...(probe.anchorId ? { anchorId: probe.anchorId } : {}),
        resource: probe.resource,
      });
    }
  });
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
    reviewProfile: options.reviewProfile,
    commentProfile: options.commentProfile,
    unsupportedContent: options.unsupportedContent,
    pagination: {
      mode: "paginated",
      scale: 1,
      pageGap: 0,
      showPageNumbers: false,
      fragmentParagraphs: true,
      cssPrefix: "page-",
    },
  };
  return sha256(utf8Bytes(canonicalJson(layoutContract)));
}

async function rendererIdentity(
  document: Document,
  layoutDigest: string,
  runtimeAssets: RuntimeAssetIdentity,
  runtimeVersion: VersionInfo,
  fonts: FontResolution[],
): Promise<string> {
  const view = document.defaultView;
  if (!view) {
    fail("unsupported_runtime", "browser_launch", "The render document has no browser identity.",
      "Use an attached browser document.");
  }
  const navigatorWithUaData = view.navigator as Navigator & {
    userAgentData?: {
      brands?: Array<{ brand: string; version: string }>;
      mobile?: boolean;
      platform?: string;
      getHighEntropyValues?: (hints: string[]) => Promise<Record<string, unknown>>;
    };
  };
  let highEntropyUserAgent: Record<string, unknown> | undefined;
  try {
    highEntropyUserAgent = await navigatorWithUaData.userAgentData?.getHighEntropyValues?.([
      "architecture", "bitness", "fullVersionList", "model", "platformVersion", "wow64",
    ]);
  } catch {
    // High-entropy UA values are optional browser observations. The ordinary
    // user agent and platform below remain part of the fingerprint.
  }
  const fontRecords = fonts.map((font) => ({ ...font }))
    .sort((left, right) => canonicalJson(left).localeCompare(canonicalJson(right)));
  const fingerprint = {
    contract: "docxodus-standalone-browser-v1",
    verification: "browserObserved",
    runtimeAssets,
    runtimeVersion,
    paginatorContractVersion: 1,
    pageMapSchemaVersion: 1,
    renderReportSchemaVersion: 1,
    layoutDigest,
    fontConfigurationDigest: await sha256(utf8Bytes(canonicalJson(fontRecords))),
    fonts: fontRecords,
    browser: {
      userAgent: view.navigator.userAgent,
      platform: view.navigator.platform,
      language: view.navigator.language,
      languages: Array.from(view.navigator.languages),
      userAgentData: navigatorWithUaData.userAgentData
        ? {
          brands: navigatorWithUaData.userAgentData.brands,
          mobile: navigatorWithUaData.userAgentData.mobile,
          platform: navigatorWithUaData.userAgentData.platform,
          highEntropy: highEntropyUserAgent,
        }
        : undefined,
      viewport: [view.innerWidth, view.innerHeight],
      devicePixelRatio: view.devicePixelRatio,
      hardwareConcurrency: view.navigator.hardwareConcurrency,
      maxTouchPoints: view.navigator.maxTouchPoints,
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      timezoneOffsetMinutes: new Date().getTimezoneOffset(),
      screen: {
        width: view.screen.width,
        height: view.screen.height,
        colorDepth: view.screen.colorDepth,
        pixelDepth: view.screen.pixelDepth,
      },
      media: {
        print: view.matchMedia("print").matches,
        darkColorScheme: view.matchMedia("(prefers-color-scheme: dark)").matches,
        reducedMotion: view.matchMedia("(prefers-reduced-motion: reduce)").matches,
        forcedColors: view.matchMedia("(forced-colors: active)").matches,
      },
    },
  };
  return sha256(utf8Bytes(canonicalJson(fingerprint)));
}

function finalPages(pages: HTMLElement[]): CompleteRenderReport["pages"] {
  return pages.map((page, index) => ({
    pageNumber: Number.parseInt(page.dataset.pageNumber ?? String(index + 1), 10),
    width: Number.parseFloat(page.dataset.pageWidthPt ?? page.style.width),
    height: Number.parseFloat(page.dataset.pageHeightPt ?? page.style.height),
    sectionIndex: Number.parseInt(page.dataset.sectionIndex ?? "0", 10),
  }));
}

async function verifyOfflineReopen(
  hostDocument: Document,
  html: string,
  expectedPages: CompleteRenderReport["pages"],
  state: ExecutionState,
): Promise<void> {
  const frame = await createIsolatedFrame(hostDocument, state, html, "output_verification");
  try {
    const reopened = frame.contentDocument!;
    const pages = Array.from(reopened.querySelectorAll<HTMLElement>(".page-box"));
    if (pages.length !== expectedPages.length) {
      throw new Error(`offline page count changed (${pages.length} != ${expectedPages.length})`);
    }
    for (let index = 0; index < pages.length; index++) {
      const page = pages[index];
      const expected = expectedPages[index];
      const rect = page.getBoundingClientRect();
      const widthPt = rect.width * 72 / 96;
      const heightPt = rect.height * 72 / 96;
      if (Math.abs(widthPt - expected.width) > 0.1
        || Math.abs(heightPt - expected.height) > 0.1
        || Number.parseInt(page.dataset.sectionIndex ?? "0", 10) !== expected.sectionIndex) {
        throw new Error(`offline geometry changed on page ${expected.pageNumber}`);
      }
    }
    if (reopened.querySelector("script, link[rel='stylesheet'], iframe, object, embed")) {
      throw new Error("offline output contains an active or external-loading element");
    }
  } finally {
    frame.remove();
  }
}

function reportBase(
  manifest: PackageManifest,
  sourceBytes: Uint8Array,
  options: NormalizedOptions,
  layoutDigest: string,
  state: ExecutionState,
): RenderReportBase {
  return {
    schema: REPORT_SCHEMA,
    schemaVersion: 1,
    source: {
      rawPackageBytesDigest: manifest.rawPackageBytesDigest.value.toLowerCase(),
      byteLength: sourceBytes.byteLength,
      documentVersion: options.documentVersion,
    },
    options: {
      reviewProfile: options.reviewProfile,
      commentProfile: options.commentProfile,
      layoutDigest,
    },
    readiness: state.readiness.map((outcome) => ({
      ...outcome,
      pending: [...outcome.pending],
      ...(outcome.diagnostics
        ? { diagnostics: outcome.diagnostics.map((diagnostic) => ({ ...diagnostic })) }
        : {}),
    })),
    fonts: state.fonts.map((font) => ({ ...font })),
    resources: state.resources.map((resource) => ({ ...resource })),
    unsupportedContent: state.unsupportedContent.map((outcome) => ({ ...outcome })),
    warnings: state.warnings.map((warning) => ({ ...warning })),
  };
}

function failureReport(
  manifest: PackageManifest,
  sourceBytes: Uint8Array,
  options: NormalizedOptions,
  layoutDigest: string,
  rendererFingerprint: string | undefined,
  state: ExecutionState,
  error: DocxodusExportError,
  pages?: CompleteRenderReport["pages"],
): FailedRenderReport {
  const unavailable: FailedRenderReport["unavailable"] = [];
  if (!rendererFingerprint) unavailable.push({
    field: "environment.rendererFingerprint",
    reason: `Failure occurred during ${error.phase} before renderer identity completed.`,
  });
  unavailable.push(
    { field: "bindings.pageMapDigest", reason: "No verified PageMap is published for a failed render." },
    { field: "bindings.htmlDigest", reason: "No HTML artifact is published for a failed render." },
    { field: "bindings.pdfDigest", reason: "PDF output was not requested by the browser materializer." },
  );
  return {
    ...reportBase(manifest, sourceBytes, options, layoutDigest, state),
    status: "failed",
    failure: {
      code: error.code,
      phase: error.phase,
      message: error.message,
      remediation: error.remediation,
    },
    ...(rendererFingerprint
      ? { environment: { rendererFingerprint, verification: "browserObserved" as const } }
      : {}),
    ...(pages ? { partial: { pages } } : {}),
    unavailable,
  };
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
  const sourcePromise = ownedBytes(document);
  const state: ExecutionState = {
    startedAt: Date.now(),
    deadline: Date.now() + options.timeoutMs,
    phase: "input_validation",
    readiness: [],
    warnings: [],
    fonts: [],
    resources: [],
    unsupportedContent: [],
  };
  let sourceBytes: Uint8Array | undefined;
  let worker: WorkerDocxodus | undefined;
  let frame: HTMLIFrameElement | undefined;
  let manifest: PackageManifest | undefined;
  let runtimeAssets: RuntimeAssetIdentity | undefined;
  let runtimeVersion: VersionInfo | undefined;
  let layoutDigest = "";
  let rendererFingerprint: string | undefined;
  let pagesForFailure: CompleteRenderReport["pages"] | undefined;

  try {
    sourceBytes = await runPhase(state, "input_validation", ["document bytes"], () => sourcePromise);
    enforceLimit(sourceBytes.byteLength, options.limits.compressedDocxBytes,
      "compressedDocxBytes", "input_validation");
    if (sourceBytes.byteLength === 0) {
      fail("invalid_document", "input_validation", "The DOCX input is empty.",
        "Pass a non-empty OPC package.");
    }
    if (options.reviewProfile === "original") {
      fail("unsupported_runtime", "input_validation",
        "The original revision profile is not yet available in the browser materializer.",
        "Use final or markup until issue #444 supplies shared reject-revision projection.");
    }
    layoutDigest = await layoutDigestForOptions(options);

    runtimeAssets = await runPhase(state, "wasm_initialization", ["runtime asset graph"], () =>
      loadRuntimeAssetIdentity(options.wasmBasePath));
    worker = await runPhase(state, "wasm_initialization", ["WASM worker"], () =>
      createWorkerDocxodus({ wasmBasePath: options.wasmBasePath }));
    runtimeVersion = await runPhase(state, "wasm_initialization", ["WASM runtime identity"], () =>
      worker!.getVersion());
    manifest = await runPhase(state, "package_preflight", ["package manifest"], () =>
      worker!.generatePackageManifest(sourceBytes!));
    preflightManifest(manifest, sourceBytes, options, state);

    const convertedHtml = await runPhase(state, "docx_conversion", ["WASM conversion"], () =>
      worker!.convertDocxToHtml(sourceBytes!, conversionOptions(options)));
    const attemptCheckpoint = checkpointAttemptState(state);
    let finalized: FinalizedTree | undefined;
    let stableReferenceSignature: string | undefined;
    let pageTreeRetries = 0;
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        frame = await createIsolatedFrame(globalThis.document, state, bootstrapHtml(options.title));
        const renderDocument = frame.contentDocument!;
        sanitizeConvertedDocument(renderDocument, convertedHtml, state, options);
        inventoryConvertedContent(renderDocument, state, options);

        const fontTask = documentFontReadiness(renderDocument);
        await runPhase(state, "font_loading", fontTask.pending, async (signal) => {
          recordFontReadiness(await fontTask.wait(signal), state, options);
        });
        const images = Array.from(renderDocument.images);
        const imageTask = documentImageReadiness(renderDocument);
        await runPhase(state, "image_decoding", imageTask.pending, async (signal) => {
          recordImageReadiness(images, await imageTask.wait(signal), state, options);
        });
        const graphics = graphicElements(renderDocument);
        const graphicTask = documentGraphicReadiness(renderDocument);
        await runPhase(
          state,
          "chart_svg_materialization",
          graphicTask.pending,
          async (signal) => {
            recordGraphicReadiness(graphics, await graphicTask.wait(signal), state, options);
          },
        );

        rendererFingerprint = await rendererIdentity(
          renderDocument,
          layoutDigest,
          runtimeAssets,
          runtimeVersion,
          state.fonts,
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
            if (Date.now() >= state.deadline) {
              fail("readiness_timeout", state.phase,
                `Export timed out during ${state.phase}.`,
                "Increase timeoutMs or reduce document layout complexity.",
                { detail: "cooperative pagination checkpoint" });
            }
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
          assertNoClippedContent(renderDocument);
        });
        enforceLimit(renderDocument.querySelectorAll("*").length,
          options.limits.domNodes, "domNodes", "running_story_placement");
        const automaticResources = automaticResourceCount(renderDocument);
        enforceLimit(automaticResources.count, options.limits.automaticResources,
          "automaticResources", "running_story_placement");
        enforceLimit(automaticResources.bytes, options.limits.automaticResourceBytes,
          "automaticResourceBytes", "running_story_placement");

        const stabilityTask = pageTreeReadiness(renderDocument, pages);
        const stability = await runPhase(
          state,
          "page_tree_stability",
          stabilityTask.pending,
          async (signal): Promise<PageTreeStabilityProbe> => {
            try {
              return await stabilityTask.wait(signal);
            } catch (error) {
              if (error instanceof PrintReadinessError) {
                throw new PageTreeInstabilityError(error.message);
              }
              throw error;
            }
          },
        );
        if (stableReferenceSignature === undefined) {
          stableReferenceSignature = stability.signature;
          frame.remove();
          frame = undefined;
          restoreAttemptState(state, attemptCheckpoint);
          if (pageTreeRetries > 0) addPageTreeRetryWarning(state);
          continue;
        }
        if (stableReferenceSignature !== stability.signature) {
          // A mismatch resets the reference: publication requires two
          // consecutive pristine-tree attempts with the same signature.
          stableReferenceSignature = stability.signature;
          throw new PageTreeInstabilityError(
            "Consecutive layouts created from the same pristine converted HTML produced different final page trees",
          );
        }
        finalized = { frame, document: renderDocument, engine, pages };
        break;
      } catch (error) {
        frame?.remove();
        frame = undefined;
        if (!(error instanceof PageTreeInstabilityError)) throw error;
        pageTreeRetries++;
        restoreAttemptState(state, attemptCheckpoint);
        addPageTreeRetryWarning(state);
        if (attempt === 3) throw error;
      }
    }
    if (!finalized) {
      fail("pagination_failure", "page_tree_stability",
        "The final page tree did not produce two matching stable signatures within three attempts.",
        "Remove asynchronous layout inputs or report the source document to Docxodus.");
    }
    frame = finalized.frame;
    const { document: renderDocument, engine, pages } = finalized;
    const finalRendererFingerprint = rendererFingerprint;
    if (!finalRendererFingerprint) {
      fail("output_verification_failure", "output_verification",
        "Renderer identity is unavailable after successful pagination.",
        "Retry in an attached standards-compliant browser.");
    }
    pagesForFailure = finalPages(pages);
    const pageMap = await runPhase(state, "output_verification", ["PageMap geometry"], () =>
      engine.materializePageMap(options.documentVersion, finalRendererFingerprint));
    const pageMapDigest = await runPhase(state, "output_verification", ["PageMap digest"], () =>
      sha256(utf8Bytes(canonicalJson(pageMap))));
    const html = serializeDocument(renderDocument);
    enforceLimit(utf8Bytes(html).byteLength, options.limits.htmlOutputBytes,
      "htmlOutputBytes", "output_verification");
    await runPhase(state, "output_verification", ["offline reopen"], () =>
      verifyOfflineReopen(globalThis.document, html, pagesForFailure!, state));
    const htmlDigest = await runPhase(state, "output_verification", ["HTML digest"], () =>
      sha256(utf8Bytes(html)));

    return await runPhase(state, "output_verification", ["final artifact assembly"], () => {
      const report: CompleteRenderReport = {
        ...reportBase(manifest!, sourceBytes!, options, layoutDigest, state),
        status: "complete",
        environment: { rendererFingerprint: finalRendererFingerprint, verification: "browserObserved" },
        pages: pagesForFailure!,
        bindings: {
          pageMapDigest,
          htmlDigest,
          artifactRequestIds: [],
        },
      };
      return {
        html,
        pageCount: pages.length,
        pageMap,
        renderReport: report,
        warnings: report.warnings,
        rendererFingerprint: finalRendererFingerprint,
      };
    });
  } catch (error) {
    const resolved = asExportError(error, state.phase);
    if (manifest && sourceBytes) {
      resolved.report = failureReport(
        manifest,
        sourceBytes,
        options,
        layoutDigest,
        rendererFingerprint,
        state,
        resolved,
        pagesForFailure,
      );
    }
    throw resolved;
  } finally {
    state.phase = "cleanup";
    frame?.remove();
    worker?.terminate();
  }
}

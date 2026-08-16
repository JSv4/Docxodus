import { createHash, randomUUID } from "node:crypto";
import type { BigIntStats } from "node:fs";
import { lstat, mkdtemp, open, realpath, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { isAbsolute, join, resolve } from "node:path";
import { chromium, type Browser, type BrowserContext, type Page } from "playwright-core";
import {
  automaticUrlAllowed,
  cssSecurityTokens,
  dataUrlInfo,
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  type FontResolverRequest,
  type PaginatedHtmlOptions,
} from "docxodus/export-browser";
import { loadVerifiedAssetGraph } from "./assets.js";
import type {
  BrowserMaterializationFailure,
  BrowserMaterializationSuccess,
  ExportPhase,
  NodeExportRuntime,
  ValidatedNodeExportRuntime,
} from "./contracts.js";
import {
  attachFailedReport,
  DocxodusExportError,
  exportError,
  fromBrowserFailure,
} from "./contracts.js";

const PLAYWRIGHT_VERSION = "1.57.0";
const EXECUTABLE_BYTES_MAX = 1024 * 1024 * 1024;
const FILE_READ_CHUNK_BYTES = 64 * 1024;
const PDF_STREAM_CHUNK_BYTES = 64 * 1024;
const CLEANUP_TIMEOUT_MS = 10_000;
const DENIED_REQUEST_DETAILS_MAX = 64;
const READINESS_PENDING_DETAILS_MAX = 64;
const READINESS_PENDING_LABEL_MAX = 512;
const LAUNCH_FLAGS = Object.freeze([
  "--disable-background-networking",
  "--disable-component-update",
  "--disable-default-apps",
  "--disable-domain-reliability",
  "--disable-sync",
  "--host-resolver-rules=MAP * ~NOTFOUND",
  "--metrics-recording-only",
  "--no-default-browser-check",
  "--no-first-run",
]);
const EXPORT_PHASES = new Set<ExportPhase>([
  "input_validation", "package_preflight", "browser_launch", "wasm_initialization",
  "docx_conversion", "font_loading", "image_decoding", "chart_svg_materialization",
  "pagination", "running_story_placement", "page_tree_stability", "pdf_print",
  "output_verification", "output_write", "filesystem_commit", "cleanup",
]);

export interface RequestLogEntry {
  url: string;
  method: string;
  resourceType: string;
  disposition: "allowed" | "denied";
}

export interface BrowserRuntimeIdentity {
  chromiumProduct: "chromium";
  browserVersion: string;
  executableDigest?: string;
  launchMode: "injected" | "explicit" | "pinned";
  launchFlags: readonly string[];
  playwrightVersion: string;
  assetManifestDigest: string;
  packageVersion: string;
  platform: NodeJS.Platform;
  architecture: string;
  chromiumSandbox: boolean | "caller-owned";
}

export interface BrowserRenderOutcome {
  materialization: BrowserMaterializationSuccess;
  pdf?: Uint8Array;
  runtime: BrowserRuntimeIdentity;
  requestLog: RequestLogEntry[];
}

interface BridgeResponse {
  ok: boolean;
  result?: BrowserMaterializationSuccess & { pdfHtml?: string };
  error?: BrowserMaterializationFailure;
}

interface FontBindingResponse {
  ok: boolean;
  result?: unknown;
  error?: Record<string, unknown>;
}

interface ReadinessProgress {
  phase: ExportPhase;
  status: "pending" | "complete" | "failed";
  pending: string[];
}

export interface OwnedExportBrowserSession {
  browser: Browser;
  close(): Promise<void>;
}

interface VerifiedExecutable {
  path: string;
  digest: string;
  identity: BigIntStats;
}

interface HostOwnedBrowserIdentity {
  executableDigest: string;
  launchMode: "explicit" | "pinned";
}

const HOST_OWNED_BROWSERS = new WeakMap<Browser, HostOwnedBrowserIdentity>();

function boundedPending(pending: string | readonly string[]): string[] {
  const values = typeof pending === "string" ? [pending] : pending;
  const bounded = values.slice(0, READINESS_PENDING_DETAILS_MAX).map((value, index) => {
    const normalized = value.replace(/[\u0000-\u001f\u007f]+/g, " ").trim()
      || `resource-${index + 1}`;
    return normalized.length <= READINESS_PENDING_LABEL_MAX
      ? normalized
      : `${normalized.slice(0, READINESS_PENDING_LABEL_MAX - 3)}...`;
  });
  if (values.length > READINESS_PENDING_DETAILS_MAX) {
    bounded.push(`... ${values.length - READINESS_PENDING_DETAILS_MAX} more`);
  }
  return bounded;
}

function parseReadinessProgress(value: unknown): ReadinessProgress | undefined {
  if (!value || typeof value !== "object") return undefined;
  const progress = value as Partial<ReadinessProgress>;
  if (typeof progress.phase !== "string" || !EXPORT_PHASES.has(progress.phase as ExportPhase)
    || (progress.status !== "pending" && progress.status !== "complete" && progress.status !== "failed")
    || !Array.isArray(progress.pending)
    || !progress.pending.every((entry) => typeof entry === "string")) {
    return undefined;
  }
  return {
    phase: progress.phase as ExportPhase,
    status: progress.status,
    pending: boundedPending(progress.pending),
  };
}

function timeoutError(phase: ExportPhase, pending: string | readonly string[]) {
  const resources = boundedPending(pending);
  return new DocxodusExportError(
    "readiness_timeout",
    phase,
    `Export timed out during ${phase}.`,
    "Increase timeoutMs or reduce document/runtime complexity.",
    { detail: resources.join(", "), pending: resources },
  );
}

function cancellationError(
  phase: ExportPhase,
  pending: string | readonly string[],
): DocxodusExportError {
  const resources = boundedPending(pending);
  return new DocxodusExportError(
    "operation_cancelled",
    phase,
    `Export was cancelled during ${phase}.`,
    "Retry with a non-aborted signal.",
    { pending: resources },
  );
}

function remaining(deadline: number, phase: "browser_launch" | "wasm_initialization" | "pdf_print"): number {
  const value = deadline - performance.now();
  if (value <= 0) throw timeoutError(phase, phase);
  return value;
}

async function bounded<T>(
  context: BrowserContext | undefined,
  deadline: number,
  phase: "browser_launch" | "wasm_initialization" | "pdf_print",
  pending: string,
  signal: AbortSignal | undefined,
  operation: () => Promise<T>,
  readinessProgress?: () => ReadinessProgress | undefined,
): Promise<T> {
  if (signal?.aborted) throw cancellationError(phase, pending);
  const timeoutMs = remaining(deadline, phase);
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abortListener: (() => void) | undefined;
  try {
    const contenders: Array<Promise<T>> = [];
    if (signal) {
      contenders.push(new Promise<never>((_, reject) => {
        abortListener = () => {
          void context?.close().catch(() => undefined);
          const reported = readinessProgress?.();
          const progress = reported?.status === "pending" ? reported : undefined;
          reject(cancellationError(
            progress?.phase ?? phase,
            progress && progress.pending.length > 0 ? progress.pending : pending,
          ));
        };
        signal.addEventListener("abort", abortListener, { once: true });
        if (signal.aborted) abortListener();
      }));
    }
    if (!signal?.aborted) contenders.push(operation());
    contenders.push(new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          void context?.close().catch(() => undefined);
          const reported = readinessProgress?.();
          const progress = reported?.status === "pending" ? reported : undefined;
          reject(timeoutError(
            progress?.phase ?? phase,
            progress && progress.pending.length > 0 ? progress.pending : pending,
          ));
        }, timeoutMs);
      }));
    return await Promise.race(contenders);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (signal && abortListener) signal.removeEventListener("abort", abortListener);
  }
}

function sameIdentity(left: BigIntStats, right: BigIntStats): boolean {
  return left.dev === right.dev
    && left.ino === right.ino
    && left.size === right.size
    && left.mtimeNs === right.mtimeNs
    && left.ctimeNs === right.ctimeNs;
}

async function removeOwnedTemporaryDirectory(path: string, identity: BigIntStats): Promise<void> {
  let current: BigIntStats;
  try {
    current = await lstat(path, { bigint: true });
  } catch (cause) {
    if ((cause as NodeJS.ErrnoException).code === "ENOENT") return;
    throw cause;
  }
  if (!current.isDirectory() || current.isSymbolicLink()
    || current.dev !== identity.dev || current.ino !== identity.ino) {
    throw new Error("Refusing to remove a replaced Chromium temporary directory.");
  }
  await rm(path, { recursive: true, force: false });
}

async function digestExecutable(
  path: string,
  deadline: number,
  signal: AbortSignal | undefined,
): Promise<VerifiedExecutable> {
  let handle: Awaited<ReturnType<typeof open>> | undefined;
  try {
    if (!isAbsolute(path)) throw new Error("executable path is not absolute");
    const requestedPath = resolve(path);
    const requestedIdentity = await lstat(requestedPath, { bigint: true });
    if (!requestedIdentity.isFile() || requestedIdentity.isSymbolicLink()) {
      throw new Error("executable path is not an ordinary non-symlink file");
    }
    const resolvedPath = await realpath(path);
    if (resolvedPath !== requestedPath) {
      throw new Error("executable path contains a symlink or non-canonical component");
    }
    handle = await open(resolvedPath, "r");
    const before = await handle.stat({ bigint: true });
    if (!before.isFile()) throw new Error("not a regular file");
    if (before.size > BigInt(EXECUTABLE_BYTES_MAX)) {
      throw new Error(`executable exceeds ${EXECUTABLE_BYTES_MAX} bytes`);
    }
    const hash = createHash("sha256");
    const chunk = Buffer.allocUnsafe(FILE_READ_CHUNK_BYTES);
    let position = 0;
    while (BigInt(position) < before.size) {
      if (signal?.aborted) throw cancellationError("browser_launch", "Chromium verification");
      remaining(deadline, "browser_launch");
      const requested = Number(
        before.size - BigInt(position) > BigInt(chunk.byteLength)
          ? BigInt(chunk.byteLength)
          : before.size - BigInt(position),
      );
      const { bytesRead } = await handle.read(chunk, 0, requested, position);
      if (bytesRead === 0) throw new Error("executable ended while it was being hashed");
      hash.update(chunk.subarray(0, bytesRead));
      position += bytesRead;
    }
    const probe = await handle.read(chunk, 0, 1, position);
    const after = await handle.stat({ bigint: true });
    const pathAfter = await stat(resolvedPath, { bigint: true });
    if (probe.bytesRead !== 0 || !sameIdentity(requestedIdentity, before)
      || !sameIdentity(before, after) || !sameIdentity(after, pathAfter)) {
      throw new Error("executable changed while it was being verified");
    }
    return { path: resolvedPath, digest: hash.digest("hex"), identity: after };
  } catch (cause) {
    if (cause instanceof DocxodusExportError) throw cause;
    return exportError(
      "browser_launch_failure",
      "browser_launch",
      `The configured Chromium executable cannot be verified: ${path}`,
      "Provide a readable Chromium executable built for this host.",
      { cause },
    );
  } finally {
    await handle?.close().catch(() => undefined);
  }
}

function browserEnvironment(temporaryDirectory: string): NodeJS.ProcessEnv {
  const inherited = ["SystemRoot", "WINDIR", "COMSPEC", "PATHEXT"] as const;
  const environment: NodeJS.ProcessEnv = {};
  for (const name of inherited) {
    if (process.env[name] !== undefined) environment[name] = process.env[name];
  }
  return {
    ...environment,
    HOME: temporaryDirectory,
    USERPROFILE: temporaryDirectory,
    XDG_CACHE_HOME: temporaryDirectory,
    XDG_CONFIG_HOME: temporaryDirectory,
    XDG_RUNTIME_DIR: temporaryDirectory,
    TMPDIR: temporaryDirectory,
    TMP: temporaryDirectory,
    TEMP: temporaryDirectory,
    TZ: "UTC",
  };
}

async function launchBrowser(
  runtime: NodeExportRuntime,
  deadline: number,
  signal?: AbortSignal,
): Promise<{
  browser: Browser;
  owned: boolean;
  temporaryDirectory?: string;
  temporaryIdentity?: BigIntStats;
  launchMode: BrowserRuntimeIdentity["launchMode"];
  executableDigest?: string;
}> {
  if (runtime.browser && runtime.browserExecutablePath) {
    exportError(
      "invalid_argument",
      "input_validation",
      "browser and browserExecutablePath cannot be supplied together.",
      "Inject one caller-owned Chromium browser or provide one executable path.",
    );
  }
  if (runtime.browser) {
    if (runtime.browser.browserType().name() !== "chromium" || !runtime.browser.isConnected()) {
      exportError(
        "browser_launch_failure",
        "browser_launch",
        "The injected browser is not a connected Chromium browser.",
        "Inject a connected Playwright Chromium Browser instance.",
      );
    }
    const hostIdentity = HOST_OWNED_BROWSERS.get(runtime.browser);
    return {
      browser: runtime.browser,
      owned: false,
      launchMode: hostIdentity?.launchMode ?? "injected",
      executableDigest: hostIdentity?.executableDigest,
    };
  }

  let temporaryDirectory: string;
  let temporaryIdentity: BigIntStats;
  try {
    temporaryDirectory = await mkdtemp(join(tmpdir(), "docxodus-export-"));
    temporaryIdentity = await lstat(temporaryDirectory, { bigint: true });
  } catch (cause) {
    exportError(
      "filesystem_failure",
      "browser_launch",
      "A private temporary directory could not be created for Chromium.",
      "Verify temporary-directory permissions and available space.",
      { cause },
    );
  }
  const explicit = runtime.browserExecutablePath;
  let launchPromise: Promise<Browser> | undefined;
  try {
    const executable = await digestExecutable(explicit ?? chromium.executablePath(), deadline, signal);
    launchPromise = chromium.launch({
      executablePath: executable.path,
      headless: true,
      chromiumSandbox: true,
      args: [...LAUNCH_FLAGS],
      timeout: remaining(deadline, "browser_launch"),
      env: browserEnvironment(temporaryDirectory),
    });
    const browser = await bounded(
      undefined,
      deadline,
      "browser_launch",
      "Chromium process launch",
      signal,
      () => launchPromise!,
    );
    const executableAfter = await stat(executable.path, { bigint: true });
    if (!sameIdentity(executable.identity, executableAfter)) {
      await browser.close().catch(() => undefined);
      exportError(
        "browser_launch_failure",
        "browser_launch",
        "The Chromium executable changed while the browser was launching.",
        "Use an immutable, deployment-owned Chromium installation.",
      );
    }
    return {
      browser,
      owned: true,
      temporaryDirectory,
      temporaryIdentity,
      launchMode: explicit ? "explicit" : "pinned",
      executableDigest: executable.digest,
    };
  } catch (cause) {
    if (launchPromise) {
      void launchPromise.then((browser) => browser.close()).catch(() => undefined);
    }
    let cleanupFailure: unknown;
    await removeOwnedTemporaryDirectory(temporaryDirectory, temporaryIdentity)
      .catch((error) => { cleanupFailure = error; });
    if (cause instanceof DocxodusExportError) {
      if (cleanupFailure === undefined) throw cause;
      throw new DocxodusExportError(cause.code, cause.phase, cause.message, cause.remediation, {
        detail: cause.detail,
        pending: cause.pending,
        partUri: cause.partUri,
        anchorId: cause.anchorId,
        resource: cause.resource,
        cause: new AggregateError([
          ...(cause.cause === undefined ? [] : [cause.cause]),
          cleanupFailure,
        ]),
        report: cause.report,
      });
    }
    exportError(
      "browser_launch_failure",
      "browser_launch",
      "Chromium could not be launched for export.",
      explicit
        ? "Verify browserExecutablePath and its shared-library dependencies."
        : "Install @playwright/browser-chromium during deployment or provide browserExecutablePath.",
      { cause: cleanupFailure === undefined ? cause : new AggregateError([cause, cleanupFailure]) },
    );
  }
}

async function cleanupStep(label: string, operation: () => Promise<void>): Promise<unknown | undefined> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    await Promise.race([
      operation(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(
          () => reject(new Error(`${label} exceeded the ${CLEANUP_TIMEOUT_MS}ms cleanup limit.`)),
          CLEANUP_TIMEOUT_MS,
        );
      }),
    ]);
    return undefined;
  } catch (error) {
    return error;
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

/** Internal host lifecycle: one owned browser, one fresh context per render batch. */
export async function openOwnedExportBrowserSession(
  browserExecutablePath: string | undefined,
  timeoutMs: number,
  signal?: AbortSignal,
): Promise<OwnedExportBrowserSession> {
  const launch = await launchBrowser({ browserExecutablePath }, performance.now() + timeoutMs, signal);
  if (!launch.owned || !launch.temporaryDirectory || !launch.temporaryIdentity) {
    exportError(
      "browser_launch_failure",
      "browser_launch",
      "The export host did not acquire an owned Chromium browser.",
      "Report this host lifecycle invariant failure.",
    );
  }
  HOST_OWNED_BROWSERS.set(launch.browser, {
    executableDigest: launch.executableDigest!,
    launchMode: launch.launchMode as "explicit" | "pinned",
  });
  let closed = false;
  return {
    browser: launch.browser,
    async close(): Promise<void> {
      if (closed) return;
      closed = true;
      HOST_OWNED_BROWSERS.delete(launch.browser);
      const failures: unknown[] = [];
      const browserFailure = await cleanupStep("Owned Chromium shutdown", () => launch.browser.close());
      if (browserFailure !== undefined) failures.push(browserFailure);
      const directoryFailure = await cleanupStep("Chromium temporary-directory removal", () =>
        removeOwnedTemporaryDirectory(launch.temporaryDirectory!, launch.temporaryIdentity!));
      if (directoryFailure !== undefined) failures.push(directoryFailure);
      if (failures.length > 0) {
        exportError(
          "filesystem_failure",
          "cleanup",
          "The export host could not close its owned browser runtime cleanly.",
          "Verify temporary-directory permissions and process cleanup.",
          { cause: new AggregateError(failures) },
        );
      }
    },
  };
}

function safeDiagnosticUrl(value: string): string {
  try {
    const url = new URL(value);
    const search = url.search ? "?<redacted>" : "";
    const safe = `${url.protocol}//${url.host}${url.pathname}${search}`;
    return safe.length <= 2048 ? safe : `${safe.slice(0, 2045)}...`;
  } catch {
    return value.length <= 2048 ? value : `${value.slice(0, 2045)}...`;
  }
}

/**
 * Install the host-owned bindings used by the script-disabled print verifier.
 * This stays in the same module as the verifier so tests can exercise the
 * exact production bridge without widening the package's public entry point.
 */
export async function exposePrintReadinessBindings(
  page: Page,
  onProgress?: (value: unknown) => void,
): Promise<void> {
  await page.exposeBinding("__docxodusReadinessProgress", (_source, value: unknown) => {
    onProgress?.(value);
  });
  await page.exposeBinding("__docxodusCssSecurityTokens", (_source, value: unknown) => {
    if (typeof value !== "string") return [];
    return cssSecurityTokens(value).map((token) => {
      const info = token.kind === "url" ? dataUrlInfo(token.value.trim()) : undefined;
      return {
        kind: token.kind,
        value: token.value,
        allowed: token.kind === "url" && automaticUrlAllowed(token.value),
        ...(info ? { mediaType: info.mediaType, byteLength: info.byteLength } : {}),
      };
    });
  });
  await page.exposeBinding("__docxodusAutomaticImageUrl", (_source, value: unknown) => {
    if (typeof value !== "string" || !automaticUrlAllowed(value)) return undefined;
    const info = dataUrlInfo(value.trim());
    return info?.mediaType.startsWith("image/") ? info : undefined;
  });
}

export async function validatePrintDocument(
  page: Page,
  expectedPageMap: BrowserMaterializationSuccess["pageMap"],
  timeoutMs: number,
  limits: {
    fontRequests: number;
    fontSampleCodePoints: number;
    visualResources: number;
    domNodes: number;
    automaticResourceBytes: number;
  },
  expectedFonts: BrowserMaterializationSuccess["renderReport"]["fontReadiness"],
  expectedResources: BrowserMaterializationSuccess["renderReport"]["resources"],
): Promise<void> {
  const readiness = await page.evaluate(async ({
    timeoutMs: readinessTimeoutMs,
    limits: readinessLimits,
    expectedFonts: expectedFontOutcomes,
    expectedResources: expectedVisualOutcomes,
  }) => {
    type Phase = "font_loading" | "image_decoding" | "chart_svg_materialization" | "page_tree_stability";
    class ReadinessFailure extends Error {
      constructor(
        readonly code: "readiness_timeout" | "resource_limit" | "output_verification_failure",
        readonly phase: Phase,
        message: string,
        readonly pending: string[],
      ) {
        super(message);
      }
    }
    const deadline = performance.now() + readinessTimeoutMs;
    const pendingLimit = 64;
    const labelLimit = 512;
    const bound = (values: readonly string[]): string[] => {
      const result = values.slice(0, pendingLimit).map((value, index) => {
        const normalized = value.replace(/[\u0000-\u001f\u007f]+/g, " ").trim()
          || `resource-${index + 1}`;
        return normalized.length <= labelLimit
          ? normalized
          : `${normalized.slice(0, labelLimit - 3)}...`;
      });
      if (values.length > pendingLimit) result.push(`... ${values.length - pendingLimit} more`);
      return result;
    };
    const publish = async (phase: Phase, status: "pending" | "complete" | "failed", pending: string[]) => {
      const reporter = (globalThis as typeof globalThis & {
        __docxodusReadinessProgress?: (value: {
          phase: Phase;
          status: "pending" | "complete" | "failed";
          pending: string[];
        }) => unknown;
      }).__docxodusReadinessProgress;
      if (typeof reporter !== "function") return;
      try {
        await reporter({ phase, status, pending: bound(pending) });
      } catch {
        // Progress is diagnostic-only.
      }
    };
    const abortError = () => new DOMException("Print readiness was aborted", "AbortError");
    const wait = async (milliseconds: number, signal: AbortSignal): Promise<void> => {
      if (signal.aborted) throw abortError();
      await new Promise<void>((resolve, reject) => {
        let settled = false;
        const timer = setTimeout(() => {
          if (settled) return;
          settled = true;
          signal.removeEventListener("abort", onAbort);
          resolve();
        }, milliseconds);
        const onAbort = () => {
          if (settled) return;
          settled = true;
          clearTimeout(timer);
          reject(abortError());
        };
        signal.addEventListener("abort", onAbort, { once: true });
        if (signal.aborted) onAbort();
      });
    };
    const frame = (signal: AbortSignal) => new Promise<void>((resolve, reject) => {
      let settled = false;
      const request = requestAnimationFrame(() => {
        if (settled) return;
        settled = true;
        signal.removeEventListener("abort", onAbort);
        resolve();
      });
      const onAbort = () => {
        if (settled) return;
        settled = true;
        cancelAnimationFrame(request);
        reject(abortError());
      };
      signal.addEventListener("abort", onAbort, { once: true });
      if (signal.aborted) onAbort();
    });
    const run = async <T>(
      phase: Phase,
      pending: () => string[],
      operation: (signal: AbortSignal) => Promise<T>,
    ): Promise<T> => {
      const remainingMs = deadline - performance.now();
      if (remainingMs <= 0) {
        throw new ReadinessFailure(
          "readiness_timeout", phase, `Print readiness timed out during ${phase}.`, bound(pending()),
        );
      }
      const controller = new AbortController();
      let timer: ReturnType<typeof setTimeout> | undefined;
      await publish(phase, "pending", pending());
      try {
        const result = await Promise.race([
          operation(controller.signal),
          new Promise<never>((_, reject) => {
            timer = setTimeout(() => {
              const resources = bound(pending());
              controller.abort();
              reject(new ReadinessFailure(
                "readiness_timeout", phase, `Print readiness timed out during ${phase}.`, resources,
              ));
            }, remainingMs);
          }),
        ]);
        await publish(phase, "complete", []);
        return result;
      } catch (error) {
        await publish(phase, "failed", error instanceof ReadinessFailure ? error.pending : pending());
        throw error;
      } finally {
        if (timer !== undefined) clearTimeout(timer);
        controller.abort();
      }
    };
    const label = (value: string, fallback: string): string => {
      const normalized = value.replace(/[\u0000-\u001f\u007f]+/g, " ").trim() || fallback;
      return normalized.length <= 256 ? normalized : `${normalized.slice(0, 253)}...`;
    };
    const sha256 = async (domain: string, value: string): Promise<string> => {
      const material = new TextEncoder().encode(`${domain}\u0000${value}`);
      const digest = await crypto.subtle.digest("SHA-256", material);
      return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    };
    type CssImageToken = {
      kind: "import" | "substitution" | "url";
      value: string;
      allowed: boolean;
      mediaType?: string;
      byteLength?: number;
    };
    const cssTokens = async (value: string): Promise<CssImageToken[]> => {
      const binding = (globalThis as typeof globalThis & {
        __docxodusCssSecurityTokens?: (css: string) => Promise<CssImageToken[]>;
      }).__docxodusCssSecurityTokens;
      if (typeof binding !== "function") {
        throw new ReadinessFailure(
          "output_verification_failure", "image_decoding",
          "The canonical CSS resource tokenizer is unavailable in the print context.",
          ["css-background-tokenizer"],
        );
      }
      return binding(value);
    };
    const imageUrlInfo = async (value: string): Promise<{
      mediaType: string;
      byteLength: number;
    } | undefined> => {
      const binding = (globalThis as typeof globalThis & {
        __docxodusAutomaticImageUrl?: (url: string) => Promise<{
          mediaType: string;
          byteLength: number;
        } | undefined>;
      }).__docxodusAutomaticImageUrl;
      return typeof binding === "function" ? binding(value) : undefined;
    };
    const resourceAnchor = (element: Element): string | undefined =>
      element.closest<HTMLElement>("[data-source-anchor-id]")?.dataset.sourceAnchorId;
    const uniqueLabels = (labels: string[]): string[] => {
      const totals = new Map<string, number>();
      const seen = new Map<string, number>();
      labels.forEach((entry) => totals.set(entry, (totals.get(entry) ?? 0) + 1));
      return labels.map((entry) => {
        if ((totals.get(entry) ?? 0) === 1) return entry;
        const occurrence = (seen.get(entry) ?? 0) + 1;
        seen.set(entry, occurrence);
        return `${entry}#${occurrence}`;
      });
    };
    const mapPool = async <T, U>(
      values: readonly T[],
      operation: (value: T, index: number) => Promise<U>,
    ): Promise<U[]> => {
      const results = new Array<U>(values.length);
      let cursor = 0;
      await Promise.all(Array.from({ length: Math.min(16, values.length) }, async () => {
        while (cursor < values.length) {
          const index = cursor++;
          results[index] = await operation(values[index], index);
        }
      }));
      return results;
    };
    let resourceMutationVersion = 0;
    const resourceObserver = new MutationObserver(() => { resourceMutationVersion++; });
    resourceObserver.observe(document.documentElement, {
      attributes: true,
      attributeFilter: [
        "src", "srcset", "href", "xlink:href", "style", "class", "data-docxodus-materialization",
        "data-docxodus-materialization-state", "data-docxodus-materialization-id",
      ],
      childList: true,
      characterData: true,
      subtree: true,
    });
    try {
      let fontPending = ["font-inventory", "document.fonts.ready"];
      const fonts = await run("font_loading", () => fontPending, async (signal) => {
        const candidates = new Map<string, { family: string; sample: string }>();
        let sampled = 0;
        let visited = 0;
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (++visited % 1_024 === 0) await wait(0, signal);
          const element = node.parentElement;
          const text = node.textContent?.trim();
          if (!element || !text || element.closest("script,style,template,noscript")) continue;
          const style = getComputedStyle(element);
          if (style.display === "none" || style.visibility === "hidden"
            || style.contentVisibility === "hidden") continue;
          const familySpec = style.fontFamily.trim();
          if (!familySpec) continue;
          const specification = [
            style.fontStyle || "normal",
            style.fontVariant || "normal",
            style.fontWeight || "400",
            style.fontStretch || "normal",
            style.fontSize || "12px",
            familySpec,
          ].join(" ");
          const remaining = Math.max(0, readinessLimits.fontSampleCodePoints - sampled);
          const addition = Array.from(text).slice(0, Math.min(256, remaining)).join("");
          const existing = candidates.get(specification);
          if (existing && addition) {
            const candidateRemaining = Math.max(0, 256 - Array.from(existing.sample).length);
            const appended = Array.from(addition).slice(0, candidateRemaining).join("");
            existing.sample += appended;
            sampled += Array.from(appended).length;
          } else if (!existing) {
            if (candidates.size >= readinessLimits.fontRequests) {
              throw new ReadinessFailure(
                "resource_limit", "font_loading",
                `Font readiness exceeded its ${readinessLimits.fontRequests}-request limit.`,
                [`font-request-limit:${readinessLimits.fontRequests}`],
              );
            }
            const sample = addition || " ";
            sampled += Array.from(addition).length;
            candidates.set(specification, { family: label(familySpec, "sans-serif"), sample });
          }
        }
        const entries = Array.from(candidates, ([specification, value]) => ({ specification, ...value }))
          .sort((left, right) => left.specification < right.specification ? -1 : 1);
        fontPending = entries.map(({ family }) => `font:${family}`);
        const probes = await Promise.all(entries.map(async ({ specification, family, sample }) => {
          const requestKey = await sha256(
            "docxodus:font-request:v1",
            JSON.stringify({ specification, sample }),
          );
          try {
            await document.fonts.load(specification, sample);
            return {
              requestKey,
              requestedFamily: family,
              available: document.fonts.check(specification, sample),
            };
          } catch {
            return { requestKey, requestedFamily: family, available: false };
          }
        }));
        await document.fonts.ready;
        fontPending = [];
        return probes;
      });
      const expected = expectedFontOutcomes.map((font) => ({
        requestKey: font.requestKey,
        available: font.available,
      })).sort((left, right) => left.requestKey < right.requestKey ? -1 : 1);
      const actual = fonts.map((font) => ({
        requestKey: font.requestKey,
        available: font.available,
      })).sort((left, right) => left.requestKey < right.requestKey ? -1 : 1);
      if (JSON.stringify(expected) !== JSON.stringify(actual)) {
        throw new ReadinessFailure(
          "output_verification_failure", "font_loading",
          "The print document changed its exact font-request inventory or availability.",
          bound(fonts.map((font) =>
            `font:${font.requestedFamily}:${font.requestKey.slice(0, 12)}`)),
        );
      }

      let imagePending = ["image-inventory"];
      const imageProbes = await run("image_decoding", () => imagePending, async (signal) => {
        type Dependency = {
          source: "html-image" | "css-background" | "svg-image";
          element: Element;
          url: string;
          fallbackLabel: string;
          contentMaterial: string;
          image?: HTMLImageElement;
        };
        const dependencies: Dependency[] = [];
        const admit = (dependency: Dependency): void => {
          if (dependencies.length >= readinessLimits.visualResources) {
            throw new ReadinessFailure(
              "resource_limit", "image_decoding",
              `Image readiness exceeded its ${readinessLimits.visualResources}-resource limit.`,
              [`image-resource-limit:${readinessLimits.visualResources}`],
            );
          }
          dependencies.push(dependency);
        };
        Array.from(document.images).forEach((image, index) => {
          const base = image.alt.trim() || resourceAnchor(image) || `image-${index + 1}`;
          admit({
            source: "html-image",
            element: image,
            image,
            url: image.currentSrc || image.getAttribute("src") || "",
            fallbackLabel: label(`html-image:${label(base, `image-${index + 1}`)}`, `html-image-${index + 1}`),
            contentMaterial: JSON.stringify({
              src: image.getAttribute("src") ?? "",
              srcset: image.getAttribute("srcset") ?? "",
            }),
          });
        });
        const elements = Array.from(document.querySelectorAll<Element>("body, body *"));
        if (elements.length > readinessLimits.domNodes) {
          throw new ReadinessFailure(
            "resource_limit", "image_decoding",
            `CSS background readiness exceeded its ${readinessLimits.domNodes}-element work limit.`,
            [`css-background-element-limit:${readinessLimits.domNodes}`],
          );
        }
        const backgrounds = new Map<string, { tokens: CssImageToken[]; digest: string }>();
        let backgroundCodeUnits = 0;
        for (const [elementIndex, element] of elements.entries()) {
          if ((elementIndex + 1) % 256 === 0) await wait(0, signal);
          for (const pseudo of [null, "::before", "::after"] as const) {
            const style = getComputedStyle(element, pseudo);
            if (style.display === "none" || style.visibility === "hidden"
              || style.contentVisibility === "hidden"
              || (pseudo !== null && (style.content === "none" || style.content === "normal"))) continue;
            const background = style.backgroundImage.trim();
            if (!background || background === "none") continue;
            backgroundCodeUnits += background.length;
            if (backgroundCodeUnits > readinessLimits.automaticResourceBytes) {
              throw new ReadinessFailure(
                "resource_limit", "image_decoding",
                `CSS background readiness exceeded its ${readinessLimits.automaticResourceBytes}-code-unit work limit.`,
                [`css-background-code-unit-limit:${readinessLimits.automaticResourceBytes}`],
              );
            }
            let evidence = backgrounds.get(background);
            if (!evidence) {
              evidence = {
                tokens: await cssTokens(background),
                digest: await sha256("docxodus:computed-background:v1", background),
              };
              backgrounds.set(background, evidence);
            }
            for (const [urlIndex, token] of evidence.tokens.entries()) {
              if (token.kind !== "url" || token.value.trim() === "data:,") continue;
              const pseudoLabel = pseudo === null ? "element" : pseudo.slice(2);
              admit({
                source: "css-background",
                element,
                url: token.value.trim(),
                fallbackLabel: label(
                  `css-background:${resourceAnchor(element) || `${elementIndex + 1}-${pseudoLabel}-${urlIndex + 1}`}`,
                  `css-background-${elementIndex + 1}-${pseudoLabel}-${urlIndex + 1}`,
                ),
                contentMaterial: JSON.stringify({
                  backgroundDigest: evidence.digest,
                  pseudo: pseudoLabel,
                  urlIndex,
                }),
              });
            }
          }
        }
        Array.from(document.querySelectorAll<SVGImageElement>("svg image"))
          .forEach((element, index) => {
            const url = (element.getAttribute("href")
              ?? element.getAttribute("xlink:href") ?? "").trim();
            admit({
              source: "svg-image",
              element,
              url,
              fallbackLabel: label(
                `svg-image:${resourceAnchor(element) || index + 1}`,
                `svg-image-${index + 1}`,
              ),
              contentMaterial: JSON.stringify({ url, markup: element.outerHTML }),
            });
          });
        const labels = uniqueLabels(dependencies.map(({ fallbackLabel }) => fallbackLabel));
        const combinedVisualCount = dependencies.length
          + Array.from(new Set(Array.from(document.querySelectorAll<Element>(
            "svg,[data-docxodus-materialization]",
          )))).filter((element) => element.closest("[data-docxodus-materialization]") === element
            || !element.parentElement?.closest("[data-docxodus-materialization]")).length
          + document.querySelectorAll("svg use").length;
        if (combinedVisualCount > readinessLimits.visualResources) {
          throw new ReadinessFailure(
            "resource_limit", "image_decoding",
            `Combined visual readiness exceeded its ${readinessLimits.visualResources}-resource limit.`,
            [`visual-resource-limit:${readinessLimits.visualResources}`],
          );
        }
        imagePending = labels.map((entry) => `image:${entry}`);
        const cssDecodeCache = new Map<string, Promise<void>>();
        const probes = await mapPool(dependencies, async (dependency, index) => {
          const pending = imagePending[index];
          try {
            if (dependency.image) {
              if (typeof dependency.image.decode === "function") await dependency.image.decode();
              if (signal.aborted) throw abortError();
              if (!dependency.image.complete || dependency.image.naturalWidth <= 0
                || dependency.image.naturalHeight <= 0) {
                throw new Error("the browser reported no decoded pixels");
              }
            } else {
              const info = await imageUrlInfo(dependency.url);
              if (!info) throw new Error("the image reference is not an allowed embedded image data URL");
              if (dependency.source === "svg-image") {
                const svgImage = dependency.element as SVGImageElement & { decode?: () => Promise<void> };
                if (typeof svgImage.decode !== "function") {
                  throw new Error("SVG image decoding is unavailable");
                }
                await svgImage.decode();
                const bounds = svgImage.getBBox();
                if (!(bounds.width > 0 && bounds.height > 0)) {
                  throw new Error("the SVG image produced no drawable bounds");
                }
              } else {
                let decoding = cssDecodeCache.get(dependency.url);
                if (!decoding) {
                  decoding = (async () => {
                    const image = new Image();
                    try {
                      image.src = dependency.url;
                      await image.decode();
                      if (!image.complete || image.naturalWidth <= 0 || image.naturalHeight <= 0) {
                        throw new Error("the browser reported no decoded pixels");
                      }
                    } finally {
                      image.removeAttribute("src");
                    }
                  })();
                  cssDecodeCache.set(dependency.url, decoding);
                }
                await decoding;
              }
            }
            return {
              kind: "image" as const,
              resource: labels[index],
              readiness: "complete" as const,
              contentKey: await sha256(
                "docxodus:visual-resource:v1",
                JSON.stringify({ source: dependency.source, material: dependency.contentMaterial }),
              ),
            };
          } catch (error) {
            if (signal.aborted) throw error;
            throw new ReadinessFailure(
              "output_verification_failure", "image_decoding",
              `Image dependency failed in the print document: ${labels[index]}.`, [pending],
            );
          }
        });
        imagePending = [];
        return probes;
      });

      let graphicPending = ["graphic-inventory"];
      const graphicProbes = await run("chart_svg_materialization", () => graphicPending, async (signal) => {
        while (true) {
          const graphics = Array.from(new Set(Array.from(document.querySelectorAll<Element>(
            "svg,[data-docxodus-materialization]",
          )))).filter((element) => element.closest("[data-docxodus-materialization]") === element
            || !element.parentElement?.closest("[data-docxodus-materialization]"));
          const uses = Array.from(document.querySelectorAll<SVGUseElement>("svg use"));
          if (imageProbes.length + graphics.length + uses.length > readinessLimits.visualResources) {
            throw new ReadinessFailure(
              "resource_limit", "chart_svg_materialization",
              `Combined visual readiness exceeded its ${readinessLimits.visualResources}-resource limit.`,
              [`visual-resource-limit:${readinessLimits.visualResources}`],
            );
          }
          graphicPending = graphics.flatMap((element, index) =>
            element.getAttribute("data-docxodus-materialization-state") === "pending"
              ? [`materialization:${label(
                element.getAttribute("data-docxodus-materialization-id") || `graphic-${index + 1}`,
                `graphic-${index + 1}`,
              )}`]
              : []);
          if (graphicPending.length > 0) {
            await frame(signal);
            continue;
          }
          const graphicLabels = uniqueLabels(graphics.map((element, index) => {
            const kind = element.getAttribute("data-docxodus-materialization") === "chart"
              || element.classList.contains("chart")
              || element.closest("[class*='chart']") !== null ? "chart" : "svg";
            const base = element.getAttribute("data-docxodus-materialization-id")?.trim()
              || resourceAnchor(element) || `${kind}-${index + 1}`;
            return label(`graphic:${kind}:${label(base, `${kind}-${index + 1}`)}`, `graphic-${index + 1}`);
          }));
          const probes = await mapPool(graphics, async (element, index) => {
            const state = element.getAttribute("data-docxodus-materialization-state");
            const id = graphicLabels[index];
            if (element.hasAttribute("data-docxodus-materialization") && state !== "complete") {
              throw new ReadinessFailure(
                "output_verification_failure", "chart_svg_materialization",
                `Graphic materialization is not complete: ${id}.`, [`materialization:${id}`],
              );
            }
            const svg = element.localName === "svg"
              ? element as SVGSVGElement
              : element.querySelector<SVGSVGElement>("svg");
            const width = Number.parseFloat(svg?.getAttribute("width") ?? "");
            const height = Number.parseFloat(svg?.getAttribute("height") ?? "");
            if (!svg || (!svg.hasAttribute("viewBox") && !(width > 0 && height > 0))
              || !svg.querySelector(
                "path,rect,circle,ellipse,line,polyline,polygon,text,image,use,foreignObject",
              )) {
              throw new ReadinessFailure(
                "output_verification_failure", "chart_svg_materialization",
                `Graphic is incomplete in the print document: ${id}.`, [`materialization:${id}`],
              );
            }
            const kind = element.getAttribute("data-docxodus-materialization") === "chart"
              || element.classList.contains("chart")
              || element.closest("[class*='chart']") !== null ? "chart" as const : "svg" as const;
            return {
              kind,
              resource: id,
              readiness: "complete" as const,
              contentKey: await sha256(
                "docxodus:visual-resource:v1",
                JSON.stringify({ source: "graphic", markup: element.outerHTML }),
              ),
            };
          });
          const useLabels = uniqueLabels(uses.map((element, index) => label(
            `svg-use:${resourceAnchor(element) || index + 1}`,
            `svg-use-${index + 1}`,
          )));
          const targetDigests = new Map<Element, string>();
          let targetMarkupCodeUnits = 0;
          for (const [index, element] of uses.entries()) {
            const href = (element.getAttribute("href")
              ?? element.getAttribute("xlink:href") ?? "").trim();
            let target: Element | null = null;
            if (href.startsWith("#") && href.length > 1) {
              let id = href.slice(1);
              try { id = decodeURIComponent(id); } catch { id = ""; }
              if (id) target = document.getElementById(id);
            }
            const drawable = target && (target.matches(
              "path,rect,circle,ellipse,line,polyline,polygon,text,image,use,foreignObject",
            ) || target.querySelector(
              "path,rect,circle,ellipse,line,polyline,polygon,text,image,use,foreignObject",
            ));
            let bounds: DOMRect | SVGRect | undefined;
            try { bounds = element.getBBox(); } catch { bounds = undefined; }
            if (!target || target === element || target.contains(element) || !drawable
              || !bounds || !(bounds.width > 0 || bounds.height > 0)) {
              throw new ReadinessFailure(
                "output_verification_failure", "chart_svg_materialization",
                `SVG use did not resolve to drawable local content: ${useLabels[index]}.`,
                [`materialization:${useLabels[index]}`],
              );
            }
            let targetDigest = targetDigests.get(target);
            if (!targetDigest) {
              const markup = target.outerHTML;
              targetMarkupCodeUnits += markup.length;
              if (targetMarkupCodeUnits > readinessLimits.automaticResourceBytes) {
                throw new ReadinessFailure(
                  "resource_limit", "chart_svg_materialization",
                  `SVG use readiness exceeded its ${readinessLimits.automaticResourceBytes}-code-unit work limit.`,
                  [`svg-use-target-code-unit-limit:${readinessLimits.automaticResourceBytes}`],
                );
              }
              targetDigest = await sha256("docxodus:svg-use-target:v1", markup);
              targetDigests.set(target, targetDigest);
            }
            probes.push({
              kind: "svg" as const,
              resource: useLabels[index],
              readiness: "complete" as const,
              contentKey: await sha256(
                "docxodus:visual-resource:v1",
                JSON.stringify({ source: "svg-use", href, targetDigest }),
              ),
            });
          }
          graphicPending = [];
          return probes;
        }
      });

      const expectedVisuals = expectedVisualOutcomes
        .filter((resource) => resource.kind !== "external_link"
          && resource.readiness === "complete"
          && !resource.resource?.startsWith("css-background:"))
        .map((resource) => ({
          kind: resource.kind,
          resource: resource.resource ?? "",
          readiness: resource.readiness,
          contentKey: resource.contentKey,
        })).sort((left, right) => JSON.stringify(left) < JSON.stringify(right) ? -1 : 1);
      const actualVisuals = [...imageProbes, ...graphicProbes]
        .filter((resource) => !resource.resource.startsWith("css-background:"))
        .sort((left, right) => JSON.stringify(left) < JSON.stringify(right) ? -1 : 1);
      if (JSON.stringify(expectedVisuals) !== JSON.stringify(actualVisuals)) {
        throw new ReadinessFailure(
          "output_verification_failure", "chart_svg_materialization",
          "The print document changed its exact visual-resource inventory or outcomes.",
          bound(actualVisuals.map((resource) => `${resource.kind}:${resource.resource}`)),
        );
      }

      const pages = Array.from(document.querySelectorAll<HTMLElement>(".page-box"));
      if (pages.length === 0) {
        throw new ReadinessFailure(
          "output_verification_failure", "page_tree_stability",
          "The final print document contains no page boxes.", ["page-tree:missing"],
        );
      }
      resourceMutationVersion += resourceObserver.takeRecords().length;
      const resourceVersion = resourceMutationVersion;
      let treePending = [`page-tree:${pages.length}-pages`, "quiet-interval:100ms"];
      await run("page_tree_stability", () => treePending, async (signal) => {
        let mutations = 0;
        let resizes = 0;
        const mutationsObserver = new MutationObserver((records) => { mutations += records.length; });
        const resizeObserver = typeof ResizeObserver === "function"
          ? new ResizeObserver((records) => { resizes += records.length; })
          : undefined;
        const signature = async () => {
          const geometry = pages.map((pageElement) => {
            const rect = pageElement.getBoundingClientRect();
            return [rect.left, rect.top, rect.width, rect.height, pageElement.scrollWidth, pageElement.scrollHeight];
          });
          const serialized = JSON.stringify({ geometry, pages: pages.map((entry) => entry.outerHTML) });
          const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(serialized));
          return Array.from(new Uint8Array(digest), (value) =>
            value.toString(16).padStart(2, "0")).join("");
        };
        mutationsObserver.observe(document.documentElement, {
          attributes: true, characterData: true, childList: true, subtree: true,
        });
        pages.forEach((entry) => resizeObserver?.observe(entry));
        try {
          await frame(signal);
          await frame(signal);
          const first = await signature();
          mutations = 0;
          resizes = 0;
          await wait(100, signal);
          await frame(signal);
          await frame(signal);
          const second = await signature();
          if (first !== second || mutations !== 0 || resizes !== 0) {
            throw new ReadinessFailure(
              "output_verification_failure", "page_tree_stability",
              `The print page tree changed during its quiet interval (mutations=${mutations}, resizes=${resizes}).`,
              [`page-tree:${pages.length}-pages`],
            );
          }
          treePending = [];
        } finally {
          mutationsObserver.disconnect();
          resizeObserver?.disconnect();
        }
      });
      resourceMutationVersion += resourceObserver.takeRecords().length;
      if (resourceMutationVersion !== resourceVersion) {
        throw new ReadinessFailure(
          "output_verification_failure", "page_tree_stability",
          "The print resource inventory changed after readiness completed.",
          [`page-tree:${pages.length}-pages`],
        );
      }
      return { ok: true as const };
    } catch (error) {
      const failure = error instanceof ReadinessFailure
        ? error
        : new ReadinessFailure(
          "output_verification_failure", "page_tree_stability",
          error instanceof Error ? error.message : String(error),
          ["print-readiness"],
        );
      return {
        ok: false as const,
        failure: {
          code: failure.code,
          phase: failure.phase,
          message: failure.message,
          pending: bound(failure.pending),
        },
      };
    } finally {
      resourceObserver.disconnect();
    }
  }, { timeoutMs, limits, expectedFonts, expectedResources });
  if (!readiness.ok) {
    exportError(
      readiness.failure.code,
      readiness.failure.phase,
      readiness.failure.message,
      "Inspect the exact reopened print document and its pending resources.",
      { pending: readiness.failure.pending, detail: readiness.failure.pending.join(", ") },
    );
  }

  const audit = await page.evaluate(async () => {
    const pages = Array.from(document.querySelectorAll<HTMLElement>(".page-box"));
    const fragments: Array<{
      fragmentId: string;
      anchorId: string;
      fragmentIndex: number;
      pageNumber: number;
      geometry: { x: number; y: number; width: number; height: number };
      story: "body" | "header" | "footer" | "footnote" | "endnote" | "comment";
      inTableCell: boolean;
    }> = [];
    const story = (anchorId: string): typeof fragments[number]["story"] => {
      const first = anchorId.indexOf(":");
      const second = first < 0 ? -1 : anchorId.indexOf(":", first + 1);
      const scope = first >= 0 && second > first ? anchorId.slice(first + 1, second) : "body";
      if (scope.startsWith("hdr")) return "header";
      if (scope.startsWith("ftr")) return "footer";
      if (scope === "fn") return "footnote";
      if (scope === "en") return "endnote";
      if (scope === "cmt") return "comment";
      return "body";
    };
    for (const pageElement of pages) {
      const pageNumber = Number.parseInt(pageElement.dataset.pageNumber ?? "0", 10);
      const pageRect = pageElement.getBoundingClientRect();
      const pageWidth = Number.parseFloat(pageElement.dataset.pageWidthPt ?? "0");
      const pageHeight = Number.parseFloat(pageElement.dataset.pageHeightPt ?? "0");
      for (const element of Array.from(pageElement.querySelectorAll<HTMLElement>(
        "[data-source-anchor-id][data-page-fragment-id]",
      ))) {
        const style = getComputedStyle(element);
        if (style.display === "none" || style.visibility === "hidden") continue;
        const rect = element.getBoundingClientRect();
        let left = Math.max(rect.left, pageRect.left);
        let top = Math.max(rect.top, pageRect.top);
        let right = Math.min(rect.right, pageRect.right);
        let bottom = Math.min(rect.bottom, pageRect.bottom);
        const clips = (value: string) =>
          value === "hidden" || value === "clip" || value === "scroll" || value === "auto";
        for (let ancestor = element.parentElement;
          ancestor && ancestor !== pageElement;
          ancestor = ancestor.parentElement) {
          const ancestorStyle = getComputedStyle(ancestor);
          const ancestorRect = ancestor.getBoundingClientRect();
          if (clips(ancestorStyle.overflowX)) {
            left = Math.max(left, ancestorRect.left);
            right = Math.min(right, ancestorRect.right);
          }
          if (clips(ancestorStyle.overflowY)) {
            top = Math.max(top, ancestorRect.top);
            bottom = Math.min(bottom, ancestorRect.bottom);
          }
        }
        if (rect.width <= 0 || rect.height <= 0 || right <= left || bottom <= top) continue;
        const anchorId = element.dataset.sourceAnchorId!;
        fragments.push({
          fragmentId: element.dataset.pageFragmentId!,
          anchorId,
          fragmentIndex: Number.parseInt(element.dataset.fragmentIndex ?? "-1", 10),
          pageNumber,
          geometry: {
            x: (left - pageRect.left) * pageWidth / pageRect.width,
            y: (top - pageRect.top) * pageHeight / pageRect.height,
            width: (right - left) * pageWidth / pageRect.width,
            height: (bottom - top) * pageHeight / pageRect.height,
          },
          story: story(anchorId),
          inTableCell: element.matches("td,th") || element.closest("td,th") !== null,
        });
      }
    }
    return {
      fontsLoaded: document.fonts.status === "loaded",
      activeElements: document.querySelectorAll(
        "script, iframe, object, embed, link[rel=stylesheet], base, form, meta[http-equiv=refresh]",
      ).length,
      pages: pages.map((element) => {
        const rect = element.getBoundingClientRect();
        return {
          pageNumber: Number.parseInt(element.dataset.pageNumber ?? "0", 10),
          pageInSection: Number.parseInt(element.dataset.pageInSection ?? "0", 10),
          sectionIndex: Number.parseInt(element.dataset.sectionIndex ?? "-1", 10),
          pageName: getComputedStyle(element).page,
          width: rect.width * 72 / 96,
          height: rect.height * 72 / 96,
        };
      }),
      fragments,
    };
  });
  if (!audit.fontsLoaded || audit.activeElements !== 0
    || audit.pages.length !== expectedPageMap.pages.length
    || audit.fragments.length !== expectedPageMap.fragments.length) {
    exportError(
      "output_verification_failure",
      "pdf_print",
      "The finalized print document changed before PDF capture.",
      "Retry with the verified standalone materializer and pinned Chromium runtime.",
    );
  }
  for (const [index, actual] of audit.pages.entries()) {
    const expected = expectedPageMap.pages[index];
    if (actual.pageNumber !== expected.pageNumber
      || actual.pageInSection !== expected.pageInSection
      || actual.sectionIndex !== expected.sectionIndex
      || actual.pageName !== expected.pageName
      || Math.abs(actual.width - expected.width) > 0.1
      || Math.abs(actual.height - expected.height) > 0.1) {
      exportError(
        "output_verification_failure",
        "pdf_print",
        `The finalized print geometry changed on page ${index + 1}.`,
        "Preserve the finalized page boxes unchanged across print media activation.",
      );
    }
  }
  for (const [index, actual] of audit.fragments.entries()) {
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
      exportError(
        "output_verification_failure",
        "pdf_print",
        `The finalized print PageMap changed at fragment ${index}.`,
        "Preserve the exact serialized page tree across print media activation.",
      );
    }
  }
}

function decodedBase64Length(value: string): number {
  if (value.length % 4 !== 0
    || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) {
    throw new Error("Chromium returned non-canonical base64 stream data.");
  }
  return value.length / 4 * 3
    - (value.endsWith("==") ? 2 : value.endsWith("=") ? 1 : 0);
}

async function printPdfStream(
  context: BrowserContext,
  page: Page,
  maximumBytes: number,
): Promise<Uint8Array> {
  const session = await context.newCDPSession(page);
  let stream: string | undefined;
  try {
    const result = await session.send("Page.printToPDF", {
      printBackground: true,
      preferCSSPageSize: true,
      displayHeaderFooter: false,
      scale: 1,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      transferMode: "ReturnAsStream",
      generateTaggedPDF: true,
      generateDocumentOutline: false,
    });
    stream = result.stream;
    if (!stream) throw new Error("Chromium did not return a PDF stream handle.");
    const chunks: Buffer[] = [];
    let total = 0;
    while (true) {
      const chunk = await session.send("IO.read", {
        handle: stream,
        size: PDF_STREAM_CHUNK_BYTES,
      });
      if (chunk.base64Encoded !== true) {
        throw new Error("Chromium returned a non-base64 PDF stream chunk.");
      }
      const byteLength = decodedBase64Length(chunk.data);
      if (total + byteLength > maximumBytes) {
        exportError(
          "resource_limit",
          "pdf_print",
          `pdfOutputBytes limit exceeded while streaming (${total + byteLength} > ${maximumBytes}).`,
          "Lower document complexity or select a larger permitted limit.",
        );
      }
      const bytes = Buffer.from(chunk.data, "base64");
      if (bytes.byteLength !== byteLength) throw new Error("Chromium PDF stream length changed while decoding.");
      chunks.push(bytes);
      total += bytes.byteLength;
      if (chunk.eof) break;
    }
    return Buffer.concat(chunks, total);
  } finally {
    if (stream) await session.send("IO.close", { handle: stream }).catch(() => undefined);
    await session.detach().catch(() => undefined);
  }
}

export async function renderInBrowser(
  sourceBytes: Uint8Array,
  browserOptions: Omit<PaginatedHtmlOptions, "wasmBasePath" | "fontResolver">,
  runtime: ValidatedNodeExportRuntime,
  includeHtml: boolean,
  includePdf: boolean,
  deadline: number,
  pdfMaximumBytes: number,
  signal?: AbortSignal,
): Promise<BrowserRenderOutcome> {
  let graph: Awaited<ReturnType<typeof loadVerifiedAssetGraph>>;
  try {
    graph = await bounded(
      undefined,
      deadline,
      "wasm_initialization",
      "verified runtime asset loading",
      signal,
      loadVerifiedAssetGraph,
    );
  } catch (cause) {
    if (cause instanceof DocxodusExportError) throw cause;
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The verified browser runtime asset graph could not be loaded.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  }
  // A routed HTTPS origin provides the browser materializer the secure context
  // required by Web Crypto without opening a listening socket or consulting DNS.
  const origin = `https://${randomUUID().replaceAll("-", "")}.docxodus.invalid`;
  const inputPath = `/input-${randomUUID()}.docx`;
  const requestLog: RequestLogEntry[] = [];
  const denied: string[] = [];
  const maximumDiagnostics = browserOptions.limits?.renderDiagnostics
    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.renderDiagnostics;
  let requestLogOverflow = false;
  let deniedCount = 0;
  const recordRequest = (entry: RequestLogEntry): void => {
    if (requestLog.length >= maximumDiagnostics) {
      requestLogOverflow = true;
      return;
    }
    requestLog.push(entry);
  };
  const recordDenied = (url: string, method: string, resourceType: string): void => {
    deniedCount++;
    const safeUrl = safeDiagnosticUrl(url);
    if (denied.length < DENIED_REQUEST_DETAILS_MAX) denied.push(safeUrl);
    recordRequest({ url: safeUrl, method, resourceType, disposition: "denied" });
  };
  let inputReads = 0;
  let context: BrowserContext | undefined;
  let printContext: BrowserContext | undefined;
  let launch: Awaited<ReturnType<typeof launchBrowser>> | undefined;
  let outcome: BrowserRenderOutcome | undefined;
  let materialization: BrowserMaterializationSuccess | undefined;
  let primaryError: unknown;
  let cleanupError: unknown;
  let currentPhase: ExportPhase = "browser_launch";
  let lastReadinessProgress: ReadinessProgress | undefined;
  const fontAbortController = new AbortController();

  try {
    launch = await launchBrowser(runtime, deadline, signal);
    try {
      const contextPromise = launch.browser.newContext({
        viewport: { width: 1280, height: 960 },
        deviceScaleFactor: 1,
        locale: "en-US",
        timezoneId: "UTC",
        colorScheme: "light",
        reducedMotion: "reduce",
        serviceWorkers: "block",
        acceptDownloads: false,
        javaScriptEnabled: true,
      });
      context = await bounded(
        undefined,
        deadline,
        "browser_launch",
        "isolated browser context creation",
        signal,
        () => contextPromise,
      ).catch((error) => {
        void contextPromise.then((lateContext) => lateContext.close()).catch(() => undefined);
        throw error;
      });
    } catch (cause) {
      if (cause instanceof DocxodusExportError) throw cause;
      exportError(
        "browser_launch_failure",
        "browser_launch",
        "A fresh isolated Chromium context could not be created.",
        "Use a connected supported Chromium browser and inspect its process limits.",
        { cause },
      );
    }
    currentPhase = "wasm_initialization";
    await context.routeWebSocket("**/*", async (webSocket) => {
      recordDenied(webSocket.url(), "GET", "websocket");
      await webSocket.close({ code: 1008, reason: "Docxodus closed runtime graph" });
    });
    await context.route("**/*", async (route) => {
      const request = route.request();
      let url: URL;
      try {
        url = new URL(request.url());
      } catch {
        recordDenied(request.url(), request.method(), request.resourceType());
        await route.abort("blockedbyclient");
        return;
      }
      const entry: Omit<RequestLogEntry, "disposition"> = {
        url: safeDiagnosticUrl(request.url()),
        method: request.method(),
        resourceType: request.resourceType(),
      };
      if (url.origin !== origin || request.method() !== "GET" || url.search !== ""
        || url.username !== "" || url.password !== "") {
        recordDenied(request.url(), request.method(), request.resourceType());
        await route.abort("blockedbyclient");
        return;
      }
      if (url.pathname === inputPath) {
        inputReads++;
        if (inputReads !== 1) {
          recordDenied(request.url(), request.method(), request.resourceType());
          await route.abort("blockedbyclient");
          return;
        }
        recordRequest({ ...entry, disposition: "allowed" });
        await route.fulfill({
          status: 200,
          body: Buffer.from(sourceBytes.buffer, sourceBytes.byteOffset, sourceBytes.byteLength),
          contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          headers: { "Cache-Control": "no-store" },
        });
        return;
      }
      const asset = graph.assets.get(url.pathname);
      if (!asset) {
        recordDenied(request.url(), request.method(), request.resourceType());
        await route.abort("blockedbyclient");
        return;
      }
      recordRequest({ ...entry, disposition: "allowed" });
      await route.fulfill({
        status: 200,
        body: asset.body,
        contentType: asset.contentType,
        headers: {
          "Cache-Control": "no-store",
          "X-Content-Type-Options": "nosniff",
          ...asset.headers,
        },
      });
    });

    const page = await bounded(
      context,
      deadline,
      "wasm_initialization",
      "isolated page creation",
      signal,
      () => context!.newPage(),
    );
    page.on("popup", (popup) => {
      recordDenied(popup.url(), "GET", "popup");
      void popup.close();
    });
    page.on("download", (download) => {
      recordDenied(download.url(), "GET", "download");
      void download.cancel();
    });
    if (runtime.fontResolver) {
      await page.exposeBinding("__docxodusResolveFonts", async (
        _source,
        request: FontResolverRequest,
      ): Promise<FontBindingResponse> => {
        try {
          return {
            ok: true,
            result: await runtime.fontResolver!(request, fontAbortController.signal),
          };
        } catch (error) {
          const normalized = error instanceof DocxodusExportError
            ? error
            : new DocxodusExportError(
              "resource_policy_failure",
              "font_loading",
              "The configured font resolver failed.",
              "Inspect the configured font directories and attestations.",
              { cause: error },
            );
          return { ok: false, error: normalized.toJSON() };
        }
      });
    }
    await page.exposeBinding("__docxodusReadinessProgress", (_source, value: unknown) => {
      const progress = parseReadinessProgress(value);
      if (progress) {
        lastReadinessProgress = progress;
        currentPhase = lastReadinessProgress.phase;
      }
    });
    page.setDefaultTimeout(remaining(deadline, "wasm_initialization"));
    await bounded(context, deadline, "wasm_initialization", "browser materializer bootstrap", signal, async () => {
      await page.goto(`${origin}/index.html`, { waitUntil: "load" });
      await page.waitForFunction(() =>
        (globalThis as unknown as { __docxodusExportReady?: boolean }).__docxodusExportReady === true);
    });

    const bridgeResponse = await bounded(
      context,
      deadline,
      "wasm_initialization",
      "DOCX materialization",
      signal,
      () => page.evaluate(async ({ inputUrl, options, wantsHtml, wantsPdf }) => {
        const bridge = (globalThis as unknown as {
          __docxodusExportBridge: {
            render(
              input: string,
              browserOptions: unknown,
              includeHtml: boolean,
              includePdf: boolean,
            ): Promise<BridgeResponse>;
          };
        }).__docxodusExportBridge;
        return bridge.render(inputUrl, options, wantsHtml, wantsPdf);
      }, {
        inputUrl: `${origin}${inputPath}`,
        // Preserve the caller's requested timeout in the report, policy digest,
        // and renderer fingerprint. The Node watchdog still enforces the one
        // absolute operation deadline around this browser work.
        options: browserOptions,
        wantsHtml: includeHtml,
        wantsPdf: includePdf,
      }),
      () => lastReadinessProgress,
    );
    if (!bridgeResponse.ok || !bridgeResponse.result) {
      throw fromBrowserFailure(bridgeResponse.error ?? {});
    }
    if (inputReads !== 1) {
      exportError(
        "resource_policy_failure",
        "output_verification",
        `The browser materializer consumed the source snapshot ${inputReads} times.`,
        "Use the single-read closed runtime bridge.",
      );
    }
    const pdfHtml = bridgeResponse.result.pdfHtml;
    delete bridgeResponse.result.pdfHtml;
    if (includePdf !== (typeof pdfHtml === "string")) {
      exportError(
        "output_verification_failure",
        "output_verification",
        "The browser materializer did not return exactly one finalized PDF snapshot when requested.",
        "Use the matching hardened browser materializer.",
      );
    }
    materialization = bridgeResponse.result;
    materialization.renderReport.options.outputs = [
      ...(includeHtml ? ["html" as const] : []),
      ...(includePdf ? ["pdf" as const] : []),
    ];
    if (!includeHtml) delete materialization.renderReport.bindings.htmlDigest;

    let pdf: Uint8Array | undefined;
    if (includePdf) {
      currentPhase = "pdf_print";
      const printStarted = performance.now();
      try {
        const printOrigin = `https://${randomUUID().replaceAll("-", "")}.docxodus.invalid`;
        const printPath = `/final-${randomUUID()}.html`;
        let printReads = 0;
        const printContextPromise = launch.browser.newContext({
          viewport: { width: 1280, height: 960 },
          deviceScaleFactor: 1,
          locale: "en-US",
          timezoneId: "UTC",
          colorScheme: "light",
          reducedMotion: "reduce",
          serviceWorkers: "block",
          acceptDownloads: false,
          javaScriptEnabled: false,
        });
        printContext = await bounded(
          undefined,
          deadline,
          "pdf_print",
          "isolated PDF context creation",
          signal,
          () => printContextPromise,
        ).catch((error) => {
          void printContextPromise.then((lateContext) => lateContext.close()).catch(() => undefined);
          throw error;
        });
        await printContext.routeWebSocket("**/*", async (webSocket) => {
          recordDenied(webSocket.url(), "GET", "websocket");
          await webSocket.close({ code: 1008, reason: "Docxodus closed print snapshot" });
        });
        await printContext.route("**/*", async (route) => {
          const request = route.request();
          let url: URL;
          try {
            url = new URL(request.url());
          } catch {
            recordDenied(request.url(), request.method(), request.resourceType());
            await route.abort("blockedbyclient");
            return;
          }
          if (url.origin !== printOrigin || url.pathname !== printPath || url.search !== ""
            || url.username !== "" || url.password !== "" || request.method() !== "GET") {
            recordDenied(request.url(), request.method(), request.resourceType());
            await route.abort("blockedbyclient");
            return;
          }
          printReads++;
          if (printReads !== 1) {
            recordDenied(request.url(), request.method(), request.resourceType());
            await route.abort("blockedbyclient");
            return;
          }
          recordRequest({
            url: safeDiagnosticUrl(request.url()),
            method: request.method(),
            resourceType: request.resourceType(),
            disposition: "allowed",
          });
          await route.fulfill({
            status: 200,
            body: pdfHtml!,
            contentType: "text/html; charset=utf-8",
            headers: {
              "Cache-Control": "no-store",
              "Content-Security-Policy": "default-src 'none'; img-src data:; media-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'; navigate-to 'none'",
              "X-Content-Type-Options": "nosniff",
            },
          });
        });
        const printPage = await bounded(
          printContext,
          deadline,
          "pdf_print",
          "isolated PDF page creation",
          signal,
          () => printContext!.newPage(),
        );
        printPage.on("popup", (popup) => {
          recordDenied(popup.url(), "GET", "popup");
          void popup.close();
        });
        printPage.on("download", (download) => {
          recordDenied(download.url(), "GET", "download");
          void download.cancel();
        });
        await exposePrintReadinessBindings(printPage, (value: unknown) => {
          const progress = parseReadinessProgress(value);
          if (progress) {
            lastReadinessProgress = progress;
            currentPhase = progress.phase;
          }
        });
        printPage.setDefaultTimeout(remaining(deadline, "pdf_print"));
        lastReadinessProgress = {
          phase: "pdf_print",
          status: "pending",
          pending: ["final print-document readiness", "Chromium PDF stream"],
        };
        pdf = await bounded(
          printContext,
          deadline,
          "pdf_print",
          "Chromium PDF printing",
          signal,
          async () => {
            try {
              await printPage.goto(`${printOrigin}${printPath}`, { waitUntil: "load" });
              await printPage.waitForFunction(() =>
                document.readyState === "complete"
                && document.documentElement.dataset.docxodusStandalone === "v1"
                && document.querySelectorAll(".page-box").length > 0);
              await printPage.emulateMedia({
                media: "print",
                colorScheme: "light",
                reducedMotion: "reduce",
              });
              await validatePrintDocument(
                printPage,
                materialization!.pageMap,
                Math.max(1, remaining(deadline, "pdf_print") - 100),
                {
                  fontRequests: browserOptions.limits?.fontRequests
                    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.fontRequests,
                  fontSampleCodePoints: browserOptions.limits?.fontSampleCodePoints
                    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.fontSampleCodePoints,
                  visualResources: browserOptions.limits?.automaticResources
                    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.automaticResources,
                  domNodes: browserOptions.limits?.domNodes
                    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.domNodes,
                  automaticResourceBytes: browserOptions.limits?.automaticResourceBytes
                    ?? DEFAULT_EXPORT_RESOURCE_LIMITS.automaticResourceBytes,
                },
                materialization!.renderReport.fontReadiness,
                materialization!.renderReport.resources,
              );
              if (printReads !== 1) {
                exportError(
                  "resource_policy_failure",
                  "pdf_print",
                  `The finalized PDF snapshot was loaded ${printReads} times.`,
                  "Use the single-read closed print context.",
                );
              }
              lastReadinessProgress = {
                phase: "pdf_print",
                status: "pending",
                pending: ["Chromium PDF stream"],
              };
              currentPhase = "pdf_print";
              return await printPdfStream(printContext!, printPage, pdfMaximumBytes);
            } catch (cause) {
              if (cause instanceof DocxodusExportError) throw cause;
              throw new DocxodusExportError(
                "pdf_write_failure",
                "pdf_print",
                "Chromium failed to produce PDF bytes.",
                "Retry with the pinned Chromium runtime and inspect browser diagnostics.",
                { cause },
              );
            }
          },
          () => lastReadinessProgress,
        );
        materialization.renderReport.readiness.push({
          phase: "pdf_print",
          status: "complete",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: [],
        });
      } catch (error) {
        const failurePhase = error instanceof DocxodusExportError ? error.phase : "pdf_print";
        const failurePending = error instanceof DocxodusExportError && error.pending
          ? boundedPending(error.pending)
          : boundedPending(error instanceof DocxodusExportError && error.detail
            ? error.detail
            : "offline finalized HTML and Chromium PDF printing");
        materialization.renderReport.readiness.push({
          phase: failurePhase,
          status: error instanceof DocxodusExportError && error.code === "operation_cancelled"
            ? "cancelled"
            : "failed",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: failurePending,
        });
        throw error;
      }
    }
    currentPhase = "output_verification";
    if (requestLogOverflow) {
      exportError(
        "resource_limit",
        "output_verification",
        `renderDiagnostics limit exceeded by browser request evidence (${maximumDiagnostics}).`,
        "Use the closed runtime graph without repeated request attempts.",
      );
    }
    if (deniedCount > 0) {
      exportError(
        "resource_policy_failure",
        "output_verification",
        "The export attempted a request outside the closed runtime asset graph.",
        "Embed automatic resources and remove active or external content.",
        {
          detail: `denied=${deniedCount}\n${denied.join("\n")}`,
          resource: denied[0],
        },
      );
    }
    outcome = {
      materialization: bridgeResponse.result,
      pdf,
      requestLog,
      runtime: {
        chromiumProduct: launch.browser.browserType().name() as "chromium",
        browserVersion: launch.browser.version(),
        executableDigest: launch.executableDigest,
        launchMode: launch.launchMode,
        launchFlags: launch.launchMode === "injected" ? [] : LAUNCH_FLAGS,
        playwrightVersion: PLAYWRIGHT_VERSION,
        assetManifestDigest: graph.manifestDigest,
        packageVersion: graph.packageVersion,
        platform: process.platform,
        architecture: process.arch,
        chromiumSandbox: launch.launchMode === "injected" ? "caller-owned" : true,
      },
    };
  } catch (error) {
    const normalized = error instanceof DocxodusExportError
      ? error
      : new DocxodusExportError(
        currentPhase === "browser_launch"
          ? "browser_launch_failure"
          : currentPhase === "pdf_print"
            ? "pdf_write_failure"
            : currentPhase === "output_verification"
              ? "output_verification_failure"
              : "conversion_failure",
        currentPhase,
        `The export runtime failed unexpectedly during ${currentPhase}.`,
        "Inspect the retained cause and retry with the pinned supported runtime.",
        { cause: error },
      );
    primaryError = attachFailedReport(normalized, materialization?.renderReport);
  } finally {
    fontAbortController.abort(new Error("The browser render context has closed."));
    const cleanupFailures: unknown[] = [];
    if (printContext) {
      const failure = await cleanupStep("PDF browser context shutdown", () => printContext!.close());
      if (failure !== undefined) cleanupFailures.push(failure);
    }
    if (context) {
      const failure = await cleanupStep("Browser context shutdown", () => context!.close());
      if (failure !== undefined) cleanupFailures.push(failure);
    }
    if (launch?.owned) {
      const failure = await cleanupStep("Owned Chromium shutdown", () => launch!.browser.close());
      if (failure !== undefined) cleanupFailures.push(failure);
    }
    if (launch?.temporaryDirectory && launch.temporaryIdentity) {
      const failure = await cleanupStep("Chromium temporary-directory removal", () =>
        removeOwnedTemporaryDirectory(launch!.temporaryDirectory!, launch!.temporaryIdentity!));
      if (failure !== undefined) cleanupFailures.push(failure);
    }
    if (cleanupFailures.length > 0) cleanupError = new AggregateError(cleanupFailures);
  }

  if (primaryError !== undefined) {
    if (cleanupError !== undefined && primaryError instanceof DocxodusExportError) {
      throw new DocxodusExportError(
        primaryError.code,
        primaryError.phase,
        primaryError.message,
        primaryError.remediation,
        {
          detail: primaryError.detail,
          pending: primaryError.pending,
          partUri: primaryError.partUri,
          anchorId: primaryError.anchorId,
          resource: primaryError.resource,
          cause: new AggregateError([
            ...(primaryError.cause === undefined ? [] : [primaryError.cause]),
            cleanupError,
          ]),
          report: primaryError.report,
          committedDestinations: primaryError.committedDestinations,
        },
      );
    }
    throw primaryError;
  }
  if (cleanupError !== undefined) {
    const error = new DocxodusExportError(
      "filesystem_failure",
      "cleanup",
      "The owned browser context or temporary runtime directory could not be removed.",
      "Verify temporary-directory permissions and retry.",
      { cause: cleanupError },
    );
    throw attachFailedReport(error, materialization?.renderReport);
  }
  if (!outcome) {
    exportError(
      "conversion_failure",
      "output_verification",
      "The browser session ended without an artifact or an error.",
      "Report this invariant failure.",
    );
  }
  return outcome;
}

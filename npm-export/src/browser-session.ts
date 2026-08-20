import { createHash, randomUUID } from "node:crypto";
import type { BigIntStats } from "node:fs";
import { lstat, mkdtemp, open, realpath, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { isAbsolute, join, resolve } from "node:path";
import { chromium, type Browser, type BrowserContext, type Page } from "playwright-core";
import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  type PaginatedHtmlOptions,
} from "docxodus/export-browser";
import { loadVerifiedAssetGraph } from "./assets.js";
import type {
  BrowserMaterializationFailure,
  BrowserMaterializationSuccess,
  ExportPhase,
  NodeExportRuntime,
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

export interface RequestLogEntry {
  url: string;
  method: string;
  resourceType: string;
  disposition: "allowed" | "denied";
}

export interface BrowserRuntimeIdentity {
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

function timeoutError(phase: "browser_launch" | "wasm_initialization" | "pdf_print", pending: string) {
  return new DocxodusExportError(
    "readiness_timeout",
    phase,
    `Export timed out during ${phase}.`,
    "Increase timeoutMs or reduce document/runtime complexity.",
    { detail: pending, pending: [pending] },
  );
}

function cancellationError(
  phase: "browser_launch" | "wasm_initialization" | "pdf_print",
  pending: string,
): DocxodusExportError {
  return new DocxodusExportError(
    "operation_cancelled",
    phase,
    `Export was cancelled during ${phase}.`,
    "Retry with a non-aborted signal.",
    { pending: [pending] },
  );
}

function remaining(deadline: number, phase: "browser_launch" | "wasm_initialization" | "pdf_print"): number {
  const value = deadline - Date.now();
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
): Promise<T> {
  if (signal?.aborted) throw cancellationError(phase, pending);
  const timeoutMs = remaining(deadline, phase);
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abortListener: (() => void) | undefined;
  try {
    const contenders: Array<Promise<T>> = [operation()];
    contenders.push(new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          void context?.close().catch(() => undefined);
          reject(timeoutError(phase, pending));
        }, timeoutMs);
      }));
    if (signal) {
      contenders.push(new Promise<never>((_, reject) => {
        abortListener = () => {
          void context?.close().catch(() => undefined);
          reject(cancellationError(phase, pending));
        };
        signal.addEventListener("abort", abortListener, { once: true });
      }));
    }
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
  const launch = await launchBrowser({ browserExecutablePath }, Date.now() + timeoutMs, signal);
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

async function validatePrintDocument(
  page: Page,
  expectedPageMap: BrowserMaterializationSuccess["pageMap"],
): Promise<void> {
  const audit = await page.evaluate(async () => {
    await document.fonts.ready;
    await Promise.all(Array.from(document.images, async (image) => {
      if (!image.complete) {
        await new Promise<void>((resolve) => {
          image.addEventListener("load", () => resolve(), { once: true });
          image.addEventListener("error", () => resolve(), { once: true });
        });
      }
      if (typeof image.decode === "function") await image.decode().catch(() => undefined);
    }));
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
  browserOptions: Omit<PaginatedHtmlOptions, "wasmBasePath">,
  runtime: NodeExportRuntime,
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
        options: browserOptions,
        wantsHtml: includeHtml,
        wantsPdf: includePdf,
      }),
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
        printPage.setDefaultTimeout(remaining(deadline, "pdf_print"));
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
              await validatePrintDocument(printPage, materialization!.pageMap);
              if (printReads !== 1) {
                exportError(
                  "resource_policy_failure",
                  "pdf_print",
                  `The finalized PDF snapshot was loaded ${printReads} times.`,
                  "Use the single-read closed print context.",
                );
              }
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
        );
        materialization.renderReport.readiness.push({
          phase: "pdf_print",
          status: "complete",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: [],
        });
      } catch (error) {
        materialization.renderReport.readiness.push({
          phase: "pdf_print",
          status: error instanceof DocxodusExportError && error.code === "operation_cancelled"
            ? "cancelled"
            : "failed",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: [error instanceof DocxodusExportError && error.detail
            ? error.detail
            : "offline finalized HTML and Chromium PDF printing"],
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

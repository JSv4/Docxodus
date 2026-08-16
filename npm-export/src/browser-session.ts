import { randomUUID } from "node:crypto";
import { readFile, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { mkdtemp } from "node:fs/promises";
import { chromium, type Browser, type BrowserContext, type Page } from "playwright-core";
import type { PaginatedHtmlOptions } from "docxodus/export-browser";
import { loadVerifiedAssetGraph } from "./assets.js";
import { sha256 } from "./canonical.js";
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
const LAUNCH_FLAGS = Object.freeze([
  "--disable-background-networking",
  "--disable-component-update",
  "--disable-default-apps",
  "--disable-domain-reliability",
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
}

export interface BrowserRenderOutcome {
  materialization: BrowserMaterializationSuccess;
  pdf?: Uint8Array;
  runtime: BrowserRuntimeIdentity;
  requestLog: RequestLogEntry[];
}

interface BridgeResponse {
  ok: boolean;
  result?: BrowserMaterializationSuccess;
  error?: BrowserMaterializationFailure;
}

export interface OwnedExportBrowserSession {
  browser: Browser;
  close(): Promise<void>;
}

function timeoutError(phase: "browser_launch" | "wasm_initialization" | "pdf_print", pending: string) {
  return new DocxodusExportError(
    "readiness_timeout",
    phase,
    `Export timed out during ${phase}.`,
    "Increase timeoutMs or reduce document/runtime complexity.",
    { detail: pending },
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
  phase: "wasm_initialization" | "pdf_print",
  pending: string,
  operation: () => Promise<T>,
): Promise<T> {
  const timeoutMs = remaining(deadline, phase);
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      operation(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          // Closing the owned context cancels page work, workers, and PDF printing.
          void context?.close().catch(() => undefined);
          reject(timeoutError(phase, pending));
        }, timeoutMs);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
  }
}

async function digestExecutable(path: string): Promise<string> {
  try {
    const info = await stat(path);
    if (!info.isFile()) throw new Error("not a regular file");
    return sha256(await readFile(path));
  } catch (cause) {
    exportError(
      "browser_launch_failure",
      "browser_launch",
      `The configured Chromium executable cannot be verified: ${path}`,
      "Provide a readable Chromium executable built for this host.",
      { cause },
    );
  }
}

async function launchBrowser(
  runtime: NodeExportRuntime,
  deadline: number,
): Promise<{
  browser: Browser;
  owned: boolean;
  temporaryDirectory?: string;
  launchMode: BrowserRuntimeIdentity["launchMode"];
  executableDigest?: string;
}> {
  if (runtime.browser && runtime.browserExecutablePath) {
    exportError(
      "invalid_document",
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
    return { browser: runtime.browser, owned: false, launchMode: "injected" };
  }

  let temporaryDirectory: string;
  try {
    temporaryDirectory = await mkdtemp(join(tmpdir(), "docxodus-export-"));
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
  try {
    const executablePath = explicit ?? chromium.executablePath();
    const executableDigest = await digestExecutable(executablePath);
    const browser = await chromium.launch({
      ...(explicit ? { executablePath: explicit } : {}),
      headless: true,
      args: [...LAUNCH_FLAGS],
      timeout: remaining(deadline, "browser_launch"),
      env: {
        ...process.env,
        TMPDIR: temporaryDirectory,
        TMP: temporaryDirectory,
        TEMP: temporaryDirectory,
      },
    });
    return {
      browser,
      owned: true,
      temporaryDirectory,
      launchMode: explicit ? "explicit" : "pinned",
      executableDigest,
    };
  } catch (cause) {
    await rm(temporaryDirectory, { recursive: true, force: true }).catch(() => undefined);
    if (cause instanceof DocxodusExportError) throw cause;
    exportError(
      "browser_launch_failure",
      "browser_launch",
      "Chromium could not be launched for export.",
      explicit
        ? "Verify browserExecutablePath and its shared-library dependencies."
        : "Install @playwright/browser-chromium during deployment or provide browserExecutablePath.",
      { cause },
    );
  }
}

/** Internal host lifecycle: one owned browser, one fresh context per render batch. */
export async function openOwnedExportBrowserSession(
  browserExecutablePath: string | undefined,
  timeoutMs: number,
): Promise<OwnedExportBrowserSession> {
  const launch = await launchBrowser({ browserExecutablePath }, Date.now() + timeoutMs);
  if (!launch.owned || !launch.temporaryDirectory) {
    exportError(
      "browser_launch_failure",
      "browser_launch",
      "The export host did not acquire an owned Chromium browser.",
      "Report this host lifecycle invariant failure.",
    );
  }
  let closed = false;
  return {
    browser: launch.browser,
    async close(): Promise<void> {
      if (closed) return;
      closed = true;
      const failures: unknown[] = [];
      await launch.browser.close().catch((error) => failures.push(error));
      await rm(launch.temporaryDirectory!, { recursive: true, force: true })
        .catch((error) => failures.push(error));
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

async function activateFinalDocument(
  context: BrowserContext,
  page: Page,
  deadline: number,
): Promise<void> {
  await bounded(context, deadline, "pdf_print", "offline finalized HTML reopen", async () => {
    await page.evaluate(() => {
      const bridge = (globalThis as unknown as {
        __docxodusExportBridge: { activatePdfDocument(): void };
      }).__docxodusExportBridge;
      bridge.activatePdfDocument();
    });
    await page.waitForFunction(() =>
      document.readyState === "complete"
      && document.documentElement.dataset.docxodusStandalone === "v1"
      && document.querySelectorAll(".page-box").length > 0);
    await page.evaluate(async () => { await document.fonts.ready; });
  });
}

export async function renderInBrowser(
  sourceBytes: Uint8Array,
  browserOptions: Omit<PaginatedHtmlOptions, "wasmBasePath">,
  runtime: NodeExportRuntime,
  includeHtml: boolean,
  includePdf: boolean,
  deadline: number,
): Promise<BrowserRenderOutcome> {
  let graph: Awaited<ReturnType<typeof loadVerifiedAssetGraph>>;
  try {
    graph = await loadVerifiedAssetGraph();
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
  let inputReads = 0;
  let context: BrowserContext | undefined;
  let launch: Awaited<ReturnType<typeof launchBrowser>> | undefined;
  let outcome: BrowserRenderOutcome | undefined;
  let materialization: BrowserMaterializationSuccess | undefined;
  let primaryError: unknown;
  let cleanupError: unknown;
  let currentPhase: ExportPhase = "browser_launch";

  try {
    launch = await launchBrowser(runtime, deadline);
    try {
      context = await launch.browser.newContext({
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
    } catch (cause) {
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
      denied.push(webSocket.url());
      requestLog.push({
        url: webSocket.url(),
        method: "GET",
        resourceType: "websocket",
        disposition: "denied",
      });
      await webSocket.close({ code: 1008, reason: "Docxodus closed runtime graph" });
    });
    await context.route("**/*", async (route) => {
      const request = route.request();
      const url = new URL(request.url());
      const entry: Omit<RequestLogEntry, "disposition"> = {
        url: request.url(),
        method: request.method(),
        resourceType: request.resourceType(),
      };
      if (url.origin !== origin || request.method() !== "GET" || url.search !== "") {
        denied.push(request.url());
        requestLog.push({ ...entry, disposition: "denied" });
        await route.abort("blockedbyclient");
        return;
      }
      if (url.pathname === inputPath) {
        inputReads++;
        if (inputReads !== 1) {
          denied.push(request.url());
          requestLog.push({ ...entry, disposition: "denied" });
          await route.abort("blockedbyclient");
          return;
        }
        requestLog.push({ ...entry, disposition: "allowed" });
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
        denied.push(request.url());
        requestLog.push({ ...entry, disposition: "denied" });
        await route.abort("blockedbyclient");
        return;
      }
      requestLog.push({ ...entry, disposition: "allowed" });
      await route.fulfill({
        status: 200,
        body: asset.body,
        contentType: asset.contentType,
        headers: { "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" },
      });
    });

    const page = await context.newPage();
    page.on("popup", (popup) => { denied.push(popup.url()); void popup.close(); });
    page.on("download", (download) => { denied.push(download.url()); void download.cancel(); });
    page.setDefaultTimeout(remaining(deadline, "wasm_initialization"));
    await bounded(context, deadline, "wasm_initialization", "browser materializer bootstrap", async () => {
      await page.goto(`${origin}/index.html`, { waitUntil: "load" });
      await page.waitForFunction(() =>
        (globalThis as unknown as { __docxodusExportReady?: boolean }).__docxodusExportReady === true);
    });

    const bridgeResponse = await bounded(
      context,
      deadline,
      "wasm_initialization",
      "DOCX materialization",
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
        options: { ...browserOptions, timeoutMs: remaining(deadline, "wasm_initialization") },
        wantsHtml: includeHtml,
        wantsPdf: includePdf,
      }),
    );
    if (!bridgeResponse.ok || !bridgeResponse.result) {
      throw fromBrowserFailure(bridgeResponse.error ?? {});
    }
    materialization = bridgeResponse.result;

    let pdf: Uint8Array | undefined;
    if (includePdf) {
      currentPhase = "pdf_print";
      const printStarted = performance.now();
      try {
        await activateFinalDocument(context, page, deadline);
        await page.emulateMedia({ media: "print", colorScheme: "light", reducedMotion: "reduce" });
        pdf = new Uint8Array(await bounded(context, deadline, "pdf_print", "Chromium PDF printing", () =>
          page.pdf({
            printBackground: true,
            preferCSSPageSize: true,
            tagged: true,
            outline: false,
            displayHeaderFooter: false,
            scale: 1,
            margin: { top: "0", right: "0", bottom: "0", left: "0" },
          }).catch((cause) => {
            throw new DocxodusExportError(
              "pdf_write_failure",
              "pdf_print",
              "Chromium failed to produce PDF bytes.",
              "Retry with the pinned Chromium runtime and inspect browser diagnostics.",
              { cause },
            );
          })));
        materialization.renderReport.readiness.push({
          phase: "pdf_print",
          status: "complete",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: [],
        });
      } catch (error) {
        materialization.renderReport.readiness.push({
          phase: "pdf_print",
          status: "failed",
          elapsedMs: Math.max(0, performance.now() - printStarted),
          pending: [error instanceof DocxodusExportError && error.detail
            ? error.detail
            : "offline finalized HTML and Chromium PDF printing"],
        });
        throw error;
      }
    }
    currentPhase = "output_verification";
    if (denied.length > 0) {
      exportError(
        "resource_policy_failure",
        "output_verification",
        "The export attempted a request outside the closed runtime asset graph.",
        "Embed automatic resources and remove active or external content.",
        { detail: denied.join("\n") },
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
        launchFlags: launch.owned ? LAUNCH_FLAGS : [],
        playwrightVersion: PLAYWRIGHT_VERSION,
        assetManifestDigest: graph.manifestDigest,
        packageVersion: graph.packageVersion,
        platform: process.platform,
        architecture: process.arch,
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
    await context?.close().catch((error) => cleanupFailures.push(error));
    if (launch?.owned) {
      await launch.browser.close().catch((error) => cleanupFailures.push(error));
    }
    if (launch?.temporaryDirectory) {
      await rm(launch.temporaryDirectory, { recursive: true, force: true })
        .catch((error) => cleanupFailures.push(error));
    }
    if (cleanupFailures.length > 0) cleanupError = new AggregateError(cleanupFailures);
  }

  if (primaryError !== undefined) throw primaryError;
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

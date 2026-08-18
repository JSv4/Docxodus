import { createHash } from "node:crypto";
import type { BigIntStats } from "node:fs";
import { open, realpath, stat } from "node:fs/promises";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { DocxodusExportError, exportError } from "./contracts.js";
import { canonicalJson } from "./canonical.js";
import { decodeStrictUtf8, strictJsonParse } from "./strict-json.js";

const RUNTIME_ASSET_GRAPH_MAX_BYTES = 1024 * 1024;
const RUNTIME_ASSET_COUNT_MAX = 10_000;
const RUNTIME_ASSET_BYTES_MAX = 64 * 1024 * 1024;
const RUNTIME_ASSET_TOTAL_BYTES_MAX = 1024 * 1024 * 1024;
const PACKAGE_MANIFEST_MAX_BYTES = 1024 * 1024;
const RESERVED_ROUTES = new Set(["./bootstrap.js", "./export-assets.json", "./index.html"]);
const ASSET_MEDIA_TYPES: Readonly<Record<string, string>> = Object.freeze({
  ".css": "text/css",
  ".dat": "application/octet-stream",
  ".js": "text/javascript",
  ".json": "application/json",
  ".wasm": "application/wasm",
});

class AssetLoadCancelled extends Error {}

function throwIfAssetLoadCancelled(signal: AbortSignal): void {
  if (signal.aborted) throw new AssetLoadCancelled("Runtime asset loading was cancelled.");
}

export function createSharedAbortableLoader<T>(
  operation: (signal: AbortSignal) => Promise<T>,
): (signal?: AbortSignal) => Promise<T> {
  interface ActiveLoad {
    controller: AbortController;
    promise: Promise<T>;
    waiters: number;
  }
  let cached: T | undefined;
  let hasCached = false;
  let active: ActiveLoad | undefined;
  return (signal?: AbortSignal): Promise<T> => {
    if (signal?.aborted) {
      return Promise.reject(new DocxodusExportError(
        "operation_cancelled",
        "wasm_initialization",
        "Export was cancelled while loading the verified runtime asset graph.",
        "Retry with a non-aborted signal.",
        { pending: ["verified runtime asset graph"] },
      ));
    }
    if (hasCached) return Promise.resolve(cached as T);
    if (!active) {
      const controller = new AbortController();
      const state: ActiveLoad = {
        controller,
        promise: undefined as unknown as Promise<T>,
        waiters: 0,
      };
      state.promise = operation(controller.signal).then((value) => {
        if (!controller.signal.aborted) {
          cached = value;
          hasCached = true;
        }
        return value;
      }).finally(() => {
        if (active === state) active = undefined;
      });
      // The final waiter may cancel before the operation observes the shared
      // abort; keep that late rejection handled without caching its result.
      void state.promise.catch(() => undefined);
      active = state;
    }
    const state = active;
    state.waiters++;
    return new Promise<T>((resolve, reject) => {
      let settled = false;
      const finish = (action: () => void): void => {
        if (settled) return;
        settled = true;
        signal?.removeEventListener("abort", onAbort);
        state.waiters--;
        if (state.waiters === 0 && active === state) state.controller.abort();
        action();
      };
      const onAbort = (): void => finish(() => reject(new DocxodusExportError(
        "operation_cancelled",
        "wasm_initialization",
        "Export was cancelled while loading the verified runtime asset graph.",
        "Retry with a non-aborted signal.",
        { pending: ["verified runtime asset graph"] },
      )));
      signal?.addEventListener("abort", onAbort, { once: true });
      if (signal?.aborted) onAbort();
      state.promise.then(
        (value) => finish(() => resolve(value)),
        (error) => finish(() => reject(error)),
      );
    });
  };
}

interface ExportAssetEntry {
  path: string;
  mediaType: string;
  byteLength: number;
  sha256: string;
}

interface ExportAssetManifest {
  schema: "https://docxodus.dev/schemas/export/export-assets/v1";
  schemaVersion: 1;
  packageVersion: string;
  assets: ExportAssetEntry[];
}

interface CompanionPackageManifest {
  name: "@docxodus/export";
  version: string;
}

export interface ServedAsset {
  body: Buffer;
  contentType: string;
  headers?: Readonly<Record<string, string>>;
}

export interface VerifiedAssetGraph {
  assets: ReadonlyMap<string, ServedAsset>;
  packageVersion: string;
  manifestDigest: string;
  coordinatorDigest: string;
}

const BOOTSTRAP_HTML = `<!doctype html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self'; worker-src 'self'; connect-src 'self'; frame-src 'self'; img-src data:; media-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'; navigate-to 'none'">
<title>Docxodus export runtime</title></head><body>
<script type="module" src="/bootstrap.js"></script></body></html>`;

const BOOTSTRAP_JS = `import {
  convertDocxToPaginatedHtml,
  DocxodusExportError,
} from "/export-browser.bundle.js";

globalThis.__docxodusExportBridge = {
  async render(inputUrl, options, includeHtml, stagePdf) {
    try {
      const response = await fetch(inputUrl, { cache: "no-store", credentials: "same-origin" });
      if (!response.ok) throw new Error(\`Input snapshot could not be loaded (\${response.status}).\`);
      const fontResolver = typeof globalThis.__docxodusResolveFonts === "function"
        ? async (request, signal) => {
          if (signal.aborted) throw signal.reason;
          const resolved = await globalThis.__docxodusResolveFonts(request);
          if (!resolved || resolved.ok !== true || !resolved.result) {
            const failure = resolved?.error ?? {};
            throw new DocxodusExportError(
              failure.code ?? "resource_policy_failure",
              failure.phase ?? "font_loading",
              failure.message ?? "The Node font resolver failed without a message.",
              failure.remediation ?? "Inspect the configured font directories and attestations.",
              failure.detail ? { detail: failure.detail } : {},
            );
          }
          if (signal.aborted) throw signal.reason;
          return resolved.result;
        }
        : undefined;
      const result = await convertDocxToPaginatedHtml(
        new Uint8Array(await response.arrayBuffer()),
        { ...options, fontResolver, wasmBasePath: "/wasm/" },
      );
      return {
        ok: true,
        result: {
          ...(includeHtml ? { html: result.html } : {}),
          ...(stagePdf ? { pdfHtml: result.html } : {}),
          pageCount: result.pageCount,
          pageMap: result.pageMap,
          renderReport: result.renderReport,
          warnings: result.warnings,
          rendererFingerprint: result.rendererFingerprint,
        },
      };
    } catch (error) {
      if (error instanceof DocxodusExportError) return { ok: false, error: error.toJSON() };
      return {
        ok: false,
        error: {
          code: "conversion_failure",
          phase: "docx_conversion",
          message: error instanceof Error ? error.message : String(error),
          remediation: "Inspect the source document and browser runtime.",
        },
      };
    }
  },
};
globalThis.__docxodusExportReady = true;
`;

function sha256(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

export function runtimeAssetGraphDigest(
  manifest: Pick<ExportAssetManifest, "schemaVersion" | "packageVersion" | "assets">,
): string {
  return sha256(Buffer.from(canonicalJson({
    schemaVersion: manifest.schemaVersion,
    packageVersion: manifest.packageVersion,
    assets: manifest.assets,
  }), "utf8"));
}

function sameIdentity(left: BigIntStats, right: BigIntStats): boolean {
  return left.dev === right.dev
    && left.ino === right.ino
    && left.size === right.size
    && left.mtimeNs === right.mtimeNs
    && left.ctimeNs === right.ctimeNs;
}

async function readBoundedStableFile(
  path: string,
  maximum: number,
  signal: AbortSignal,
): Promise<Buffer> {
  throwIfAssetLoadCancelled(signal);
  const handle = await open(path, "r");
  try {
    throwIfAssetLoadCancelled(signal);
    const before = await handle.stat({ bigint: true });
    if (!before.isFile()) throw new Error("not a regular file");
    if (before.size > BigInt(maximum)) throw new Error(`file exceeds ${maximum} bytes`);
    const length = Number(before.size);
    const bytes = Buffer.allocUnsafe(length);
    let offset = 0;
    while (offset < length) {
      throwIfAssetLoadCancelled(signal);
      const { bytesRead } = await handle.read(bytes, offset, length - offset, offset);
      if (bytesRead === 0) break;
      offset += bytesRead;
    }
    throwIfAssetLoadCancelled(signal);
    const probe = Buffer.allocUnsafe(1);
    const extra = await handle.read(probe, 0, 1, offset);
    const after = await handle.stat({ bigint: true });
    const pathAfter = await stat(path, { bigint: true });
    throwIfAssetLoadCancelled(signal);
    if (offset !== length || extra.bytesRead !== 0
      || !sameIdentity(before, after) || !sameIdentity(after, pathAfter)) {
      throw new Error("file changed while it was being read");
    }
    return bytes;
  } finally {
    await handle.close();
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

function parseManifest(bytes: Buffer): ExportAssetManifest {
  let value: unknown;
  try {
    value = strictJsonParse(decodeStrictUtf8(bytes, "The export asset manifest"));
  } catch (cause) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest is not valid JSON.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  }
  const record = value as Partial<ExportAssetManifest>;
  if (!record || typeof record !== "object" || Array.isArray(record)
    || Object.keys(record).some((key) =>
      !["schema", "schemaVersion", "packageVersion", "assets"].includes(key))
    || record.schema !== "https://docxodus.dev/schemas/export/export-assets/v1"
    || record.schemaVersion !== 1
    || typeof record.packageVersion !== "string" || record.packageVersion.length === 0
    || !wellFormed(record.packageVersion)
    || !Array.isArray(record.assets)
    || record.assets.length === 0
    || record.assets.length > RUNTIME_ASSET_COUNT_MAX) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest has an unsupported shape.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
    );
  }
  let total = 0;
  const seen = new Set<string>();
  const portablePaths = new Set<string>();
  let previous = "";
  for (const [index, entry] of record.assets.entries()) {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)
      || Object.keys(entry).some((key) =>
        !["path", "mediaType", "byteLength", "sha256"].includes(key))
      || typeof entry.path !== "string" || typeof entry.mediaType !== "string"
      || !Number.isSafeInteger(entry.byteLength) || entry.byteLength < 0
      || entry.byteLength > RUNTIME_ASSET_BYTES_MAX
      || typeof entry.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(entry.sha256)) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        `Runtime asset entry ${index} has invalid identity fields.`,
        "Reinstall matching published packages.",
      );
    }
    const segments = entry.path.split("/");
    const extension = entry.path.slice(entry.path.lastIndexOf("."));
    const portablePath = entry.path.normalize("NFC").toLowerCase();
    if (!entry.path.startsWith("./") || entry.path.includes("\\")
      || entry.path.includes("?") || entry.path.includes("#")
      || !wellFormed(entry.path)
      || segments.some((segment, segmentIndex) =>
        segmentIndex > 0 && (segment === "" || segment === "." || segment === ".."))
      || seen.has(entry.path) || portablePaths.has(portablePath)
      || RESERVED_ROUTES.has(entry.path)
      || ASSET_MEDIA_TYPES[extension] !== entry.mediaType
      || (index > 0 && previous >= entry.path)) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        `Runtime asset entry ${index} has a non-canonical, colliding, or reserved path.`,
        "Regenerate the canonical runtime asset graph.",
      );
    }
    previous = entry.path;
    seen.add(entry.path);
    portablePaths.add(portablePath);
    total += entry.byteLength;
    if (total > RUNTIME_ASSET_TOTAL_BYTES_MAX) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        "The runtime asset graph exceeds its aggregate byte ceiling.",
        "Deploy a bounded Docxodus runtime package.",
      );
    }
  }
  return record as ExportAssetManifest;
}

function safeAssetFile(packageRoot: string, assetPath: string): string {
  if (!/^\.\/[a-z0-9_./-]+$/i.test(assetPath) || assetPath.includes("//")) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      `The runtime asset graph contains an unsafe path: ${assetPath}`,
      "Reinstall the published Docxodus package.",
    );
  }
  const candidate = resolve(packageRoot, assetPath.slice(2));
  const fromRoot = relative(packageRoot, candidate);
  if (fromRoot === "" || fromRoot === ".." || fromRoot.startsWith(`..${sep}`)
    || isAbsolute(fromRoot)) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      `The runtime asset graph escapes its package root: ${assetPath}`,
      "Reinstall the published Docxodus package.",
    );
  }
  return candidate;
}

function requireContainedRealPath(
  packageRoot: string,
  candidate: string,
  label: string,
  signal: AbortSignal,
): Promise<string> {
  throwIfAssetLoadCancelled(signal);
  return realpath(candidate).then((resolvedPath) => {
    throwIfAssetLoadCancelled(signal);
    const fromRoot = relative(packageRoot, resolvedPath);
    if (fromRoot === "" || fromRoot === ".." || fromRoot.startsWith(`..${sep}`)
      || isAbsolute(fromRoot)) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        `${label} resolves outside the installed Docxodus package.`,
        "Reinstall the published package without symlinked runtime assets.",
      );
    }
    return resolvedPath;
  }).catch((cause) => {
    if (cause instanceof AssetLoadCancelled) throw cause;
    if (cause instanceof DocxodusExportError) throw cause;
    return exportError(
      "unsupported_runtime",
      "wasm_initialization",
      `${label} could not be resolved as a local package file.`,
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  });
}

export const loadVerifiedAssetGraph = createSharedAbortableLoader(loadVerifiedAssetGraphCore);

async function loadVerifiedAssetGraphCore(signal: AbortSignal): Promise<VerifiedAssetGraph> {
  throwIfAssetLoadCancelled(signal);
  let manifestUrl: string;
  let materializerUrl: string;
  try {
    manifestUrl = import.meta.resolve("docxodus/export-assets.json");
    materializerUrl = import.meta.resolve("docxodus/export-browser");
  } catch (cause) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The matching Docxodus browser export package could not be resolved.",
      "Install the exact matching docxodus version beside @docxodus/export.",
      { cause },
    );
  }
  if (!manifestUrl.startsWith("file:") || !materializerUrl.startsWith("file:")) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "Docxodus runtime assets must resolve to local package files.",
      "Install both packages locally before rendering.",
    );
  }

  let manifestFile: string;
  try {
    manifestFile = await realpath(fileURLToPath(manifestUrl));
    throwIfAssetLoadCancelled(signal);
  } catch (cause) {
    if (cause instanceof AssetLoadCancelled) throw cause;
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest could not be opened.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  }
  const packageRoot = dirname(manifestFile);
  const materializerFile = await requireContainedRealPath(
    packageRoot,
    fileURLToPath(materializerUrl),
    "The public browser materializer",
    signal,
  );
  let manifestBytes: Buffer;
  try {
    manifestBytes = await readBoundedStableFile(
      manifestFile,
      RUNTIME_ASSET_GRAPH_MAX_BYTES,
      signal,
    );
  } catch (cause) {
    if (cause instanceof AssetLoadCancelled) throw cause;
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest could not be read.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  }
  const manifest = parseManifest(manifestBytes);
  throwIfAssetLoadCancelled(signal);
  let companionPackage: CompanionPackageManifest;
  try {
    const packageBytes = await readBoundedStableFile(
      fileURLToPath(new URL("../package.json", import.meta.url)),
      PACKAGE_MANIFEST_MAX_BYTES,
      signal,
    );
    companionPackage = strictJsonParse(
      decodeStrictUtf8(packageBytes, "The @docxodus/export package manifest"),
    ) as CompanionPackageManifest;
  } catch (cause) {
    if (cause instanceof AssetLoadCancelled) throw cause;
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The @docxodus/export package identity could not be read.",
      "Reinstall the published companion package.",
      { cause },
    );
  }
  if (!companionPackage || typeof companionPackage !== "object"
    || Array.isArray(companionPackage)
    || companionPackage.name !== "@docxodus/export"
    || typeof companionPackage.version !== "string"
    || companionPackage.version !== manifest.packageVersion) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed docxodus and @docxodus/export versions do not match exactly.",
      "Install the same explicit version of both packages.",
      {
        detail: `docxodus=${manifest.packageVersion}; @docxodus/export=${companionPackage?.version ?? "invalid"}`,
      },
    );
  }
  const served = new Map<string, ServedAsset>();
  const seen = new Set<string>();

  for (const entry of manifest.assets) {
    throwIfAssetLoadCancelled(signal);
    if (!entry || typeof entry.path !== "string" || typeof entry.mediaType !== "string"
      || !Number.isSafeInteger(entry.byteLength) || entry.byteLength < 0
      || typeof entry.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(entry.sha256)
      || seen.has(entry.path)) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        "The installed Docxodus runtime asset graph is malformed or contains duplicates.",
        "Reinstall matching published packages.",
      );
    }
    seen.add(entry.path);
    const file = await requireContainedRealPath(
      packageRoot,
      safeAssetFile(packageRoot, entry.path),
      `Runtime asset ${entry.path}`,
      signal,
    );
    let bytes: Buffer;
    try {
      const fileInfo = await stat(file);
      if (!fileInfo.isFile() || fileInfo.size !== entry.byteLength) {
        throw new Error("not a regular file of the declared length");
      }
      bytes = await readBoundedStableFile(file, entry.byteLength, signal);
    } catch (cause) {
      if (cause instanceof AssetLoadCancelled) throw cause;
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        `A declared Docxodus runtime asset is missing: ${entry.path}`,
        "Reinstall matching published packages.",
        { cause },
      );
    }
    if (bytes.byteLength !== entry.byteLength || sha256(bytes) !== entry.sha256) {
      exportError(
        "unsupported_runtime",
        "wasm_initialization",
        `A Docxodus runtime asset failed digest verification: ${entry.path}`,
        "Reinstall matching published packages; do not mix build outputs.",
      );
    }
    served.set(`/${entry.path.slice(2)}`, {
      body: bytes,
      contentType: entry.mediaType,
    });
  }
  throwIfAssetLoadCancelled(signal);

  const declaredMaterializer = await requireContainedRealPath(
    packageRoot,
    safeAssetFile(packageRoot, "./export-browser.bundle.js"),
    "The declared browser materializer",
    signal,
  );
  if (materializerFile !== declaredMaterializer) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The public browser export entry does not match its runtime asset graph.",
      "Install matching versions of docxodus and @docxodus/export.",
    );
  }

  served.set("/export-assets.json", {
    body: manifestBytes,
    contentType: "application/json",
  });
  served.set("/index.html", {
    body: Buffer.from(BOOTSTRAP_HTML),
    contentType: "text/html; charset=utf-8",
    headers: {
      "Content-Security-Policy": "default-src 'none'; script-src 'self'; worker-src 'self'; connect-src 'self'; frame-src 'self'; img-src data:; media-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'; navigate-to 'none'",
    },
  });
  served.set("/bootstrap.js", {
    body: Buffer.from(BOOTSTRAP_JS),
    contentType: "text/javascript; charset=utf-8",
  });
  throwIfAssetLoadCancelled(signal);

  return Object.freeze({
    assets: served,
    packageVersion: manifest.packageVersion,
    // Match the browser materializer's canonical graph identity exactly;
    // pretty-printing or key order in export-assets.json is not semantic.
    manifestDigest: runtimeAssetGraphDigest(manifest),
    // These Node-only coordinator resources are served outside the package
    // manifest. Bind their exact bytes into the final Node fingerprint too.
    coordinatorDigest: sha256(Buffer.from(canonicalJson({
      schemaVersion: 1,
      indexHtmlSha256: sha256(Buffer.from(BOOTSTRAP_HTML, "utf8")),
      bootstrapJsSha256: sha256(Buffer.from(BOOTSTRAP_JS, "utf8")),
    }), "utf8")),
  });
}

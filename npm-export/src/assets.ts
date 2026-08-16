import { createHash } from "node:crypto";
import { readFile, realpath, stat } from "node:fs/promises";
import { dirname, isAbsolute, relative, resolve, sep } from "node:path";
import { fileURLToPath } from "node:url";
import { DocxodusExportError, exportError } from "./contracts.js";

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
}

export interface VerifiedAssetGraph {
  assets: ReadonlyMap<string, ServedAsset>;
  packageVersion: string;
  manifestDigest: string;
}

const BOOTSTRAP_HTML = `<!doctype html>
<html><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self'; worker-src 'self'; connect-src 'self'; frame-src 'self'; img-src data:; font-src data:; style-src 'unsafe-inline'; object-src 'none'; base-uri 'none'; form-action 'none'">
<title>Docxodus export runtime</title></head><body>
<script type="module" src="/bootstrap.js"></script></body></html>`;

const BOOTSTRAP_JS = `import {
  awaitFinalPrintReadiness,
  convertDocxToPaginatedHtml,
  DocxodusExportError,
  PrintReadinessError,
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
      if (stagePdf) globalThis.__docxodusPdfHtml = result.html;
      return {
        ok: true,
        result: {
          ...(includeHtml ? { html: result.html } : {}),
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
  async activatePdfDocument(timeoutMs) {
    try {
      const html = globalThis.__docxodusPdfHtml;
      delete globalThis.__docxodusPdfHtml;
      if (typeof html !== "string") throw new Error("No finalized HTML is staged for PDF output.");
      document.open();
      document.write(html);
      document.close();
      const readiness = await awaitFinalPrintReadiness(document, { timeoutMs });
      return {
        ok: true,
        readiness: {
          pageCount: readiness.pageTree.pageCount,
          signature: readiness.pageTree.signature,
          quietIntervalMs: readiness.pageTree.quietIntervalMs,
          animationFrames: readiness.pageTree.animationFrames,
        },
      };
    } catch (error) {
      return {
        ok: false,
        error: {
          phase: error instanceof PrintReadinessError ? error.phase : "page_tree_stability",
          message: error instanceof Error ? error.message : String(error),
          pending: error instanceof PrintReadinessError ? Array.from(error.pending) : [],
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

function parseManifest(bytes: Buffer): ExportAssetManifest {
  let value: unknown;
  try {
    value = JSON.parse(bytes.toString("utf8"));
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
  if (record.schema !== "https://docxodus.dev/schemas/export/export-assets/v1"
    || record.schemaVersion !== 1
    || typeof record.packageVersion !== "string"
    || !Array.isArray(record.assets)
    || record.assets.length === 0) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest has an unsupported shape.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
    );
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

function requireContainedRealPath(packageRoot: string, candidate: string, label: string): Promise<string> {
  return realpath(candidate).then((resolvedPath) => {
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

let cachedGraph: Promise<VerifiedAssetGraph> | undefined;

export function loadVerifiedAssetGraph(): Promise<VerifiedAssetGraph> {
  cachedGraph ??= loadVerifiedAssetGraphCore().catch((error) => {
    cachedGraph = undefined;
    throw error;
  });
  return cachedGraph;
}

async function loadVerifiedAssetGraphCore(): Promise<VerifiedAssetGraph> {
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
  } catch (cause) {
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
  );
  let manifestBytes: Buffer;
  try {
    manifestBytes = await readFile(manifestFile);
  } catch (cause) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed Docxodus export asset manifest could not be read.",
      "Reinstall matching versions of docxodus and @docxodus/export.",
      { cause },
    );
  }
  const manifest = parseManifest(manifestBytes);
  let companionPackage: CompanionPackageManifest;
  try {
    companionPackage = JSON.parse(
      await readFile(fileURLToPath(new URL("../package.json", import.meta.url)), "utf8"),
    ) as CompanionPackageManifest;
  } catch (cause) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The @docxodus/export package identity could not be read.",
      "Reinstall the published companion package.",
      { cause },
    );
  }
  if (companionPackage.name !== "@docxodus/export"
    || typeof companionPackage.version !== "string"
    || companionPackage.version !== manifest.packageVersion) {
    exportError(
      "unsupported_runtime",
      "wasm_initialization",
      "The installed docxodus and @docxodus/export versions do not match exactly.",
      "Install the same explicit version of both packages.",
      {
        detail: `docxodus=${manifest.packageVersion}; @docxodus/export=${companionPackage.version}`,
      },
    );
  }
  const served = new Map<string, ServedAsset>();
  const seen = new Set<string>();

  for (const entry of manifest.assets) {
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
    );
    let bytes: Buffer;
    try {
      const fileInfo = await stat(file);
      if (!fileInfo.isFile()) throw new Error("not a regular file");
      bytes = await readFile(file);
    } catch (cause) {
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

  const declaredMaterializer = await requireContainedRealPath(
    packageRoot,
    safeAssetFile(packageRoot, "./export-browser.bundle.js"),
    "The declared browser materializer",
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
  });
  served.set("/bootstrap.js", {
    body: Buffer.from(BOOTSTRAP_JS),
    contentType: "text/javascript; charset=utf-8",
  });

  return Object.freeze({
    assets: served,
    packageVersion: manifest.packageVersion,
    manifestDigest: sha256(manifestBytes),
  });
}

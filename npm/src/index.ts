import type {
  ConversionOptions,
  CompareOptions,
  Revision,
  VersionInfo,
  ErrorResponse,
  CompareResult,
  DocxodusWasmExports,
} from "./types.js";

import {
  CommentRenderMode,
  RevisionType,
  isInsertion,
  isDeletion,
} from "./types.js";

export type {
  ConversionOptions,
  CompareOptions,
  Revision,
  VersionInfo,
  ErrorResponse,
  CompareResult,
};

export { CommentRenderMode, RevisionType, isInsertion, isDeletion };

let wasmExports: DocxodusWasmExports | null = null;
let initPromise: Promise<void> | null = null;

/**
 * Derive the WASM base path from this module's URL.
 * Works whether loaded from node_modules, CDN, or bundled.
 */
function getDefaultWasmBasePath(): string {
  try {
    // import.meta.url gives us the URL of this module
    // e.g., "https://cdn.jsdelivr.net/npm/docxodus@3.1.1/dist/index.js"
    // or "file:///path/to/node_modules/docxodus/dist/index.js"
    const moduleUrl = import.meta.url;

    // Remove the filename to get the directory
    const baseDir = moduleUrl.substring(0, moduleUrl.lastIndexOf('/') + 1);

    // WASM files are in ./wasm/ relative to dist/
    return baseDir + "wasm/";
  } catch {
    // Fallback if import.meta.url is not available
    return "";
  }
}

/**
 * Current base path for WASM files.
 * Empty string means auto-detect from module URL.
 */
export let wasmBasePath = "";

/**
 * Set custom base path for WASM files.
 * Pass empty string or don't call this to auto-detect from module location.
 *
 * @param path - Custom path to WASM files, or empty string for auto-detection
 */
export function setWasmBasePath(path: string): void {
  wasmBasePath = path && !path.endsWith("/") ? path + "/" : path;
}

/**
 * Initialize the Docxodus WASM runtime.
 * Must be called before using any conversion/comparison functions.
 * Safe to call multiple times - will only initialize once.
 *
 * By default, WASM files are auto-detected from the module's location
 * (works with CDN, npm, or local hosting).
 * Pass a basePath to load from a custom location instead.
 *
 * @param basePath - Optional custom path to WASM files. Leave empty for auto-detection.
 */
export async function initialize(basePath?: string): Promise<void> {
  if (wasmExports) return;

  if (initPromise) {
    return initPromise;
  }

  if (basePath !== undefined) {
    setWasmBasePath(basePath);
  }

  initPromise = loadWasm();
  return initPromise;
}

/**
 * Try to load WASM from a specific base path
 */
async function tryLoadFromPath(basePath: string): Promise<boolean> {
  try {
    const dotnetPath = basePath + "_framework/dotnet.js";
    const { dotnet } = await import(/* webpackIgnore: true */ /* @vite-ignore */ dotnetPath);

    const { getAssemblyExports, getConfig } = await dotnet
      .withDiagnosticTracing(false)
      .create();

    const config = getConfig();
    const exports = await getAssemblyExports(config.mainAssemblyName);

    wasmExports = {
      DocumentConverter: exports.DocxodusWasm.DocumentConverter,
      DocumentComparer: exports.DocxodusWasm.DocumentComparer,
    };
    return true;
  } catch {
    return false;
  }
}

async function loadWasm(): Promise<void> {
  // If a custom path is set, use it directly
  if (wasmBasePath) {
    const success = await tryLoadFromPath(wasmBasePath);
    if (success) return;
    throw new Error(
      `Failed to load WASM from custom path: ${wasmBasePath}. ` +
      `Ensure the WASM files are served at this location.`
    );
  }

  // Try to auto-detect from module URL (works for CDN and local imports)
  const autoDetectedPath = getDefaultWasmBasePath();
  if (autoDetectedPath) {
    const success = await tryLoadFromPath(autoDetectedPath);
    if (success) {
      wasmBasePath = autoDetectedPath;
      return;
    }
  }

  // Auto-detection failed
  throw new Error(
    `Failed to load WASM files. ` +
    `Auto-detected path: ${autoDetectedPath || "(none)"}. ` +
    `You can specify a custom path by calling initialize("/path/to/wasm/").`
  );
}

function ensureInitialized(): DocxodusWasmExports {
  if (!wasmExports) {
    throw new Error(
      "Docxodus not initialized. Call initialize() first and await it."
    );
  }
  return wasmExports;
}

function isErrorResponse(result: string): result is string {
  try {
    const parsed = JSON.parse(result);
    return typeof parsed === "object" && "Error" in parsed;
  } catch {
    return false;
  }
}

function parseError(result: string): ErrorResponse {
  const parsed = JSON.parse(result);
  return {
    error: parsed.Error || parsed.error,
    type: parsed.Type || parsed.type,
    stackTrace: parsed.StackTrace || parsed.stackTrace,
  };
}

/**
 * Convert a File or Uint8Array to Uint8Array
 */
async function toBytes(input: File | Uint8Array): Promise<Uint8Array> {
  if (input instanceof Uint8Array) {
    return input;
  }
  const buffer = await input.arrayBuffer();
  return new Uint8Array(buffer);
}

/**
 * Convert a DOCX document to HTML.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param options - Conversion options
 * @returns HTML string
 * @throws Error if conversion fails
 */
export async function convertDocxToHtml(
  document: File | Uint8Array,
  options?: ConversionOptions
): Promise<string> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = options
    ? exports.DocumentConverter.ConvertDocxToHtmlWithOptions(
        bytes,
        options.pageTitle ?? "Document",
        options.cssPrefix ?? "docx-",
        options.fabricateClasses ?? true,
        options.additionalCss ?? "",
        options.commentRenderMode ?? CommentRenderMode.Disabled,
        options.commentCssClassPrefix ?? "comment-"
      )
    : exports.DocumentConverter.ConvertDocxToHtml(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Conversion failed: ${error.error}`);
  }

  return result;
}

/**
 * Compare two DOCX documents and return the redlined result as a DOCX.
 *
 * @param original - Original DOCX document
 * @param modified - Modified DOCX document
 * @param options - Comparison options
 * @returns Redlined DOCX as Uint8Array
 * @throws Error if comparison fails
 */
export async function compareDocuments(
  original: File | Uint8Array,
  modified: File | Uint8Array,
  options?: CompareOptions
): Promise<Uint8Array> {
  const exports = ensureInitialized();
  const originalBytes = await toBytes(original);
  const modifiedBytes = await toBytes(modified);

  let result: Uint8Array;

  if (options?.detailThreshold !== undefined || options?.caseInsensitive) {
    result = exports.DocumentComparer.CompareDocumentsWithOptions(
      originalBytes,
      modifiedBytes,
      options?.authorName ?? "Docxodus",
      options?.detailThreshold ?? 0.15,
      options?.caseInsensitive ?? false
    );
  } else {
    result = exports.DocumentComparer.CompareDocuments(
      originalBytes,
      modifiedBytes,
      options?.authorName ?? "Docxodus"
    );
  }

  if (result.length === 0) {
    throw new Error("Comparison failed - empty result");
  }

  return result;
}

/**
 * Compare two DOCX documents and return the result as HTML.
 *
 * @param original - Original DOCX document
 * @param modified - Modified DOCX document
 * @param options - Comparison options
 * @returns HTML string with redlined content
 * @throws Error if comparison fails
 */
export async function compareDocumentsToHtml(
  original: File | Uint8Array,
  modified: File | Uint8Array,
  options?: CompareOptions
): Promise<string> {
  const exports = ensureInitialized();
  const originalBytes = await toBytes(original);
  const modifiedBytes = await toBytes(modified);

  // Use the new options method if renderTrackedChanges is explicitly set
  const renderTrackedChanges = options?.renderTrackedChanges ?? true;

  const result = exports.DocumentComparer.CompareDocumentsToHtmlWithOptions(
    originalBytes,
    modifiedBytes,
    options?.authorName ?? "Docxodus",
    renderTrackedChanges
  );

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Comparison failed: ${error.error}`);
  }

  return result;
}

/**
 * Get revisions from a compared document.
 *
 * @param document - A document that has been through comparison (has tracked changes)
 * @returns Array of revisions
 * @throws Error if operation fails
 */
export async function getRevisions(
  document: File | Uint8Array
): Promise<Revision[]> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = exports.DocumentComparer.GetRevisionsJson(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to get revisions: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return (parsed.Revisions || parsed.revisions || []).map((r: any) => ({
    author: r.Author || r.author,
    date: r.Date || r.date,
    revisionType: r.RevisionType || r.revisionType,
    text: r.Text || r.text,
  }));
}

/**
 * Get version information about the library.
 */
export function getVersion(): VersionInfo {
  const exports = ensureInitialized();
  const result = exports.DocumentConverter.GetVersion();
  const parsed = JSON.parse(result);
  return {
    library: parsed.Library || parsed.library,
    dotnetVersion: parsed.DotnetVersion || parsed.dotnetVersion,
    platform: parsed.Platform || parsed.platform,
  };
}

/**
 * Check if the WASM runtime is initialized.
 */
export function isInitialized(): boolean {
  return wasmExports !== null;
}

import type {
  ConversionOptions,
  CompareOptions,
  Revision,
  VersionInfo,
  ErrorResponse,
  CompareResult,
  DocxodusWasmExports,
  GetRevisionsOptions,
  FormatChangeDetails,
  Annotation,
  AddAnnotationRequest,
  AddAnnotationResponse,
  RemoveAnnotationResponse,
  AnnotationOptions,
  DocumentStructure,
  DocumentElement,
  TableColumnInfo,
  AnnotationTarget,
  AddAnnotationWithTargetRequest,
  DocumentMetadata,
  SectionMetadata,
  RenderPageRangeOptions,
} from "./types.js";

import {
  CommentRenderMode,
  PaginationMode,
  AnnotationLabelMode,
  RevisionType,
  DocumentElementType,
  isInsertion,
  isDeletion,
  isMove,
  isMoveSource,
  isMoveDestination,
  findMovePair,
  isFormatChange,
  findElementById,
  findElementsByType,
  getParagraphs,
  getTables,
  getTableColumns,
  targetElement,
  targetParagraph,
  targetParagraphRange,
  targetRun,
  targetTable,
  targetTableRow,
  targetTableCell,
  targetTableColumn,
  targetSearch,
  targetSearchInElement,
} from "./types.js";

// Re-export pagination types and engine
export type {
  PageDimensions,
  MeasuredBlock,
  PageInfo,
  PaginationResult,
  PaginationOptions,
} from "./pagination.js";

export { PaginationEngine, paginateHtml } from "./pagination.js";

export type {
  ConversionOptions,
  CompareOptions,
  Revision,
  VersionInfo,
  ErrorResponse,
  CompareResult,
  GetRevisionsOptions,
  FormatChangeDetails,
  Annotation,
  AddAnnotationRequest,
  AddAnnotationResponse,
  RemoveAnnotationResponse,
  AnnotationOptions,
  DocumentStructure,
  DocumentElement,
  TableColumnInfo,
  AnnotationTarget,
  AddAnnotationWithTargetRequest,
  // Lazy loading / Phase 3 types
  DocumentMetadata,
  SectionMetadata,
  RenderPageRangeOptions,
};

export {
  CommentRenderMode,
  PaginationMode,
  AnnotationLabelMode,
  RevisionType,
  DocumentElementType,
  isInsertion,
  isDeletion,
  isMove,
  isMoveSource,
  isMoveDestination,
  findMovePair,
  isFormatChange,
  // Document structure helpers
  findElementById,
  findElementsByType,
  getParagraphs,
  getTables,
  getTableColumns,
  // Annotation target factory functions
  targetElement,
  targetParagraph,
  targetParagraphRange,
  targetRun,
  targetTable,
  targetTableRow,
  targetTableCell,
  targetTableColumn,
  targetSearch,
  targetSearchInElement,
};

let wasmExports: DocxodusWasmExports | null = null;
let initPromise: Promise<void> | null = null;

/**
 * Yields to the browser's main thread, allowing pending UI updates to render.
 *
 * This is critical for WASM operations: since WASM runs synchronously on the
 * main thread, React state updates (like loading spinners) won't paint unless
 * we yield before the blocking work begins.
 *
 * Uses requestAnimationFrame which fires just before the next paint, ensuring
 * any queued state updates are committed to the DOM.
 *
 * @internal
 */
async function yieldToMain(): Promise<void> {
  // In non-browser environments (SSR, tests), skip yielding
  if (typeof requestAnimationFrame === "undefined") {
    return;
  }

  // Double-rAF ensures the browser has fully painted before we continue
  // First rAF: scheduled for next frame
  // Second rAF: ensures first frame actually painted
  await new Promise<void>((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(() => resolve());
    });
  });
}

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
 *
 * @example
 * ```typescript
 * // Basic conversion
 * const html = await convertDocxToHtml(docxFile);
 *
 * // With pagination (PDF.js-style page view)
 * const html = await convertDocxToHtml(docxFile, {
 *   paginationMode: PaginationMode.Paginated,
 *   paginationScale: 0.8
 * });
 *
 * // With annotations rendered
 * const html = await convertDocxToHtml(docxFile, {
 *   renderAnnotations: true,
 *   annotationLabelMode: AnnotationLabelMode.Above
 * });
 *
 * // With footnotes and endnotes
 * const html = await convertDocxToHtml(docxFile, {
 *   renderFootnotesAndEndnotes: true
 * });
 *
 * // With headers and footers
 * const html = await convertDocxToHtml(docxFile, {
 *   renderHeadersAndFooters: true
 * });
 *
 * // With tracked changes (redlines visible)
 * const html = await convertDocxToHtml(docxFile, {
 *   renderTrackedChanges: true,
 *   showDeletedContent: true,
 *   renderMoveOperations: true
 * });
 * ```
 */
export async function convertDocxToHtml(
  document: File | Uint8Array,
  options?: ConversionOptions
): Promise<string> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before heavy WASM work - allows loading states to render
  await yieldToMain();

  let result: string;

  // Check if any of the new complete options are specified
  const needsCompleteMethod = options?.renderFootnotesAndEndnotes !== undefined ||
    options?.renderHeadersAndFooters !== undefined ||
    options?.renderTrackedChanges !== undefined ||
    options?.showDeletedContent !== undefined ||
    options?.renderMoveOperations !== undefined;

  // Use complete method when any new options are specified (most comprehensive)
  if (needsCompleteMethod || options?.renderAnnotations) {
    result = exports.DocumentConverter.ConvertDocxToHtmlComplete(
      bytes,
      options?.pageTitle ?? "Document",
      options?.cssPrefix ?? "docx-",
      options?.fabricateClasses ?? true,
      options?.additionalCss ?? "",
      options?.commentRenderMode ?? CommentRenderMode.Disabled,
      options?.commentCssClassPrefix ?? "comment-",
      options?.paginationMode ?? PaginationMode.None,
      options?.paginationScale ?? 1.0,
      options?.paginationCssClassPrefix ?? "page-",
      options?.renderAnnotations ?? false,
      options?.annotationLabelMode ?? AnnotationLabelMode.Above,
      options?.annotationCssClassPrefix ?? "annot-",
      options?.renderFootnotesAndEndnotes ?? false,
      options?.renderHeadersAndFooters ?? false,
      options?.renderTrackedChanges ?? false,
      options?.showDeletedContent ?? true,
      options?.renderMoveOperations ?? true
    );
  }
  // Use pagination-aware method when pagination is requested
  else if (options?.paginationMode !== undefined && options.paginationMode !== PaginationMode.None) {
    result = exports.DocumentConverter.ConvertDocxToHtmlWithPagination(
      bytes,
      options.pageTitle ?? "Document",
      options.cssPrefix ?? "docx-",
      options.fabricateClasses ?? true,
      options.additionalCss ?? "",
      options.commentRenderMode ?? CommentRenderMode.Disabled,
      options.commentCssClassPrefix ?? "comment-",
      options.paginationMode,
      options.paginationScale ?? 1.0,
      options.paginationCssClassPrefix ?? "page-"
    );
  } else if (options) {
    result = exports.DocumentConverter.ConvertDocxToHtmlWithOptions(
      bytes,
      options.pageTitle ?? "Document",
      options.cssPrefix ?? "docx-",
      options.fabricateClasses ?? true,
      options.additionalCss ?? "",
      options.commentRenderMode ?? CommentRenderMode.Disabled,
      options.commentCssClassPrefix ?? "comment-"
    );
  } else {
    result = exports.DocumentConverter.ConvertDocxToHtml(bytes);
  }

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

  // Yield to browser before heavy WASM work - allows loading states to render
  await yieldToMain();

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

  // Yield to browser before heavy WASM work - allows loading states to render
  await yieldToMain();

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
 * @param options - Optional move detection configuration
 * @returns Array of revisions
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * // Default settings (move detection enabled, 80% threshold)
 * const revisions = await getRevisions(comparedDoc);
 *
 * // Custom move detection settings
 * const revisions = await getRevisions(comparedDoc, {
 *   detectMoves: true,
 *   moveSimilarityThreshold: 0.9,  // Require 90% word overlap
 *   moveMinimumWordCount: 5,       // Only consider phrases of 5+ words
 *   caseInsensitive: true          // Ignore case when matching
 * });
 *
 * // Disable move detection entirely
 * const revisions = await getRevisions(comparedDoc, { detectMoves: false });
 * ```
 */
export async function getRevisions(
  document: File | Uint8Array,
  options?: GetRevisionsOptions
): Promise<Revision[]> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  // Apply defaults for move detection options
  const detectMoves = options?.detectMoves ?? true;
  const moveSimilarityThreshold = options?.moveSimilarityThreshold ?? 0.8;
  const moveMinimumWordCount = options?.moveMinimumWordCount ?? 3;
  const caseInsensitive = options?.caseInsensitive ?? false;

  const result = exports.DocumentComparer.GetRevisionsJsonWithOptions(
    bytes,
    detectMoves,
    moveSimilarityThreshold,
    moveMinimumWordCount,
    caseInsensitive
  );

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
    moveGroupId: r.MoveGroupId ?? r.moveGroupId,
    isMoveSource: r.IsMoveSource ?? r.isMoveSource,
    formatChange: (r.FormatChange || r.formatChange) ? {
      oldProperties: r.FormatChange?.OldProperties || r.formatChange?.oldProperties,
      newProperties: r.FormatChange?.NewProperties || r.formatChange?.newProperties,
      changedPropertyNames: r.FormatChange?.ChangedPropertyNames || r.formatChange?.changedPropertyNames,
    } : undefined,
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

/**
 * Get all annotations from a document.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns Array of annotations
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const annotations = await getAnnotations(docxFile);
 * for (const annot of annotations) {
 *   console.log(`${annot.label}: "${annot.annotatedText}"`);
 * }
 * ```
 */
export async function getAnnotations(
  document: File | Uint8Array
): Promise<Annotation[]> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = exports.DocumentConverter.GetAnnotations(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to get annotations: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return (parsed.Annotations || parsed.annotations || []).map((a: any) => ({
    id: a.Id || a.id,
    labelId: a.LabelId || a.labelId,
    label: a.Label || a.label,
    color: a.Color || a.color,
    author: a.Author || a.author,
    created: a.Created || a.created,
    bookmarkName: a.BookmarkName || a.bookmarkName,
    startPage: a.StartPage ?? a.startPage,
    endPage: a.EndPage ?? a.endPage,
    annotatedText: a.AnnotatedText || a.annotatedText,
    metadata: a.Metadata || a.metadata,
  }));
}

/**
 * Add an annotation to a document.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param request - Annotation details including search text or paragraph indices
 * @returns Response with modified document bytes and annotation info
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * // Annotate by searching for text
 * const result = await addAnnotation(docxFile, {
 *   id: "annot-1",
 *   labelId: "CLAUSE_A",
 *   label: "Important Clause",
 *   color: "#FFEB3B",
 *   searchText: "shall not be liable",
 *   occurrence: 1
 * });
 *
 * // Annotate by paragraph range
 * const result = await addAnnotation(docxFile, {
 *   id: "annot-2",
 *   labelId: "SECTION_1",
 *   label: "Introduction",
 *   color: "#4CAF50",
 *   startParagraphIndex: 0,
 *   endParagraphIndex: 2
 * });
 *
 * // Get modified document
 * const modifiedDocBytes = base64ToBytes(result.documentBytes);
 * ```
 */
export async function addAnnotation(
  document: File | Uint8Array,
  request: AddAnnotationRequest
): Promise<AddAnnotationResponse> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const requestJson = JSON.stringify({
    Id: request.id,
    LabelId: request.labelId,
    Label: request.label,
    Color: request.color ?? "#FFEB3B",
    Author: request.author,
    SearchText: request.searchText,
    Occurrence: request.occurrence ?? 1,
    StartParagraphIndex: request.startParagraphIndex,
    EndParagraphIndex: request.endParagraphIndex,
    Metadata: request.metadata,
  });

  const result = exports.DocumentConverter.AddAnnotation(bytes, requestJson);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to add annotation: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  const annotation = parsed.Annotation || parsed.annotation;

  return {
    success: parsed.Success ?? parsed.success ?? true,
    documentBytes: parsed.DocumentBytes || parsed.documentBytes,
    annotation: annotation ? {
      id: annotation.Id || annotation.id,
      labelId: annotation.LabelId || annotation.labelId,
      label: annotation.Label || annotation.label,
      color: annotation.Color || annotation.color,
      author: annotation.Author || annotation.author,
      created: annotation.Created || annotation.created,
      bookmarkName: annotation.BookmarkName || annotation.bookmarkName,
      annotatedText: annotation.AnnotatedText || annotation.annotatedText,
    } : undefined,
  };
}

/**
 * Remove an annotation from a document.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param annotationId - The ID of the annotation to remove
 * @returns Response with modified document bytes
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const result = await removeAnnotation(docxFile, "annot-1");
 * const modifiedDocBytes = base64ToBytes(result.documentBytes);
 * ```
 */
export async function removeAnnotation(
  document: File | Uint8Array,
  annotationId: string
): Promise<RemoveAnnotationResponse> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = exports.DocumentConverter.RemoveAnnotation(bytes, annotationId);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to remove annotation: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return {
    success: parsed.Success ?? parsed.success ?? true,
    documentBytes: parsed.DocumentBytes || parsed.documentBytes,
  };
}

/**
 * Check if a document has any annotations.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns true if the document has annotations
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * if (await hasAnnotations(docxFile)) {
 *   const annotations = await getAnnotations(docxFile);
 *   console.log(`Document has ${annotations.length} annotations`);
 * }
 * ```
 */
export async function hasAnnotations(
  document: File | Uint8Array
): Promise<boolean> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = exports.DocumentConverter.HasAnnotations(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to check annotations: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return parsed.HasAnnotations ?? parsed.hasAnnotations ?? false;
}

/**
 * Get the document structure for element-based annotation targeting.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns Document structure with element tree
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const structure = await getDocumentStructure(docxFile);
 *
 * // Navigate the structure tree
 * console.log(`Document has ${structure.root.children.length} top-level elements`);
 *
 * // Find all paragraphs
 * const paragraphs = getParagraphs(structure);
 * console.log(`Found ${paragraphs.length} paragraphs`);
 *
 * // Find all tables
 * const tables = getTables(structure);
 * for (const table of tables) {
 *   const columns = getTableColumns(structure, table.id);
 *   console.log(`Table ${table.id} has ${columns.length} columns`);
 * }
 *
 * // Look up element by ID
 * const element = findElementById(structure, "doc/p-0");
 * if (element) {
 *   console.log(`First paragraph: "${element.textPreview}"`);
 * }
 * ```
 */
export async function getDocumentStructure(
  document: File | Uint8Array
): Promise<DocumentStructure> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const result = exports.DocumentConverter.GetDocumentStructure(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to get document structure: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  // Convert from PascalCase to camelCase
  const convertElement = (el: any): DocumentElement => ({
    id: el.Id || el.id,
    type: el.Type || el.type,
    textPreview: el.TextPreview || el.textPreview,
    index: el.Index ?? el.index,
    rowIndex: el.RowIndex ?? el.rowIndex,
    columnIndex: el.ColumnIndex ?? el.columnIndex,
    rowSpan: el.RowSpan ?? el.rowSpan,
    columnSpan: el.ColumnSpan ?? el.columnSpan,
    children: (el.Children || el.children || []).map(convertElement),
  });

  const convertTableColumn = (col: any): TableColumnInfo => ({
    tableId: col.TableId || col.tableId,
    columnIndex: col.ColumnIndex ?? col.columnIndex,
    cellIds: col.CellIds || col.cellIds || [],
    rowCount: col.RowCount ?? col.rowCount,
  });

  const root = convertElement(parsed.Root || parsed.root);

  // Convert elementsById dictionary
  const elementsById: Record<string, DocumentElement> = {};
  const rawElementsById = parsed.ElementsById || parsed.elementsById || {};
  for (const [key, el] of Object.entries(rawElementsById)) {
    elementsById[key] = convertElement(el);
  }

  // Convert tableColumns dictionary
  const tableColumns: Record<string, TableColumnInfo> = {};
  const rawTableColumns = parsed.TableColumns || parsed.tableColumns || {};
  for (const [key, col] of Object.entries(rawTableColumns)) {
    tableColumns[key] = convertTableColumn(col);
  }

  return {
    root,
    elementsById,
    tableColumns,
  };
}

/**
 * Get document metadata for lazy loading pagination.
 * This is a fast operation that extracts structure information without full HTML rendering.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns Document metadata including sections, dimensions, and content counts
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const metadata = await getDocumentMetadata(docxFile);
 *
 * // Check document overview
 * console.log(`Document has ${metadata.totalParagraphs} paragraphs`);
 * console.log(`Document has ${metadata.sections.length} sections`);
 * console.log(`Estimated ${metadata.estimatedPageCount} pages`);
 *
 * // Check section properties
 * for (const section of metadata.sections) {
 *   console.log(`Section ${section.sectionIndex}: ${section.pageWidthPt}x${section.pageHeightPt}pt`);
 *   console.log(`  Paragraphs: ${section.paragraphCount}, Tables: ${section.tableCount}`);
 *   console.log(`  Has header: ${section.hasHeader}, Has footer: ${section.hasFooter}`);
 * }
 *
 * // Check document features
 * if (metadata.hasTrackedChanges) {
 *   console.log("Document has tracked changes");
 * }
 * if (metadata.hasFootnotes) {
 *   console.log("Document has footnotes");
 * }
 * ```
 */
export async function getDocumentMetadata(
  document: File | Uint8Array
): Promise<DocumentMetadata> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const result = exports.DocumentConverter.GetDocumentMetadata(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to get document metadata: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  // Convert from PascalCase to camelCase
  const convertSection = (s: any): SectionMetadata => ({
    sectionIndex: s.SectionIndex ?? s.sectionIndex,
    pageWidthPt: s.PageWidthPt ?? s.pageWidthPt,
    pageHeightPt: s.PageHeightPt ?? s.pageHeightPt,
    marginTopPt: s.MarginTopPt ?? s.marginTopPt,
    marginRightPt: s.MarginRightPt ?? s.marginRightPt,
    marginBottomPt: s.MarginBottomPt ?? s.marginBottomPt,
    marginLeftPt: s.MarginLeftPt ?? s.marginLeftPt,
    contentWidthPt: s.ContentWidthPt ?? s.contentWidthPt,
    contentHeightPt: s.ContentHeightPt ?? s.contentHeightPt,
    headerPt: s.HeaderPt ?? s.headerPt,
    footerPt: s.FooterPt ?? s.footerPt,
    paragraphCount: s.ParagraphCount ?? s.paragraphCount,
    tableCount: s.TableCount ?? s.tableCount,
    hasHeader: s.HasHeader ?? s.hasHeader,
    hasFooter: s.HasFooter ?? s.hasFooter,
    hasFirstPageHeader: s.HasFirstPageHeader ?? s.hasFirstPageHeader,
    hasFirstPageFooter: s.HasFirstPageFooter ?? s.hasFirstPageFooter,
    hasEvenPageHeader: s.HasEvenPageHeader ?? s.hasEvenPageHeader,
    hasEvenPageFooter: s.HasEvenPageFooter ?? s.hasEvenPageFooter,
    startParagraphIndex: s.StartParagraphIndex ?? s.startParagraphIndex,
    endParagraphIndex: s.EndParagraphIndex ?? s.endParagraphIndex,
    startTableIndex: s.StartTableIndex ?? s.startTableIndex,
    endTableIndex: s.EndTableIndex ?? s.endTableIndex,
  });

  return {
    sections: (parsed.Sections || parsed.sections || []).map(convertSection),
    totalParagraphs: parsed.TotalParagraphs ?? parsed.totalParagraphs,
    totalTables: parsed.TotalTables ?? parsed.totalTables,
    hasFootnotes: parsed.HasFootnotes ?? parsed.hasFootnotes,
    hasEndnotes: parsed.HasEndnotes ?? parsed.hasEndnotes,
    hasTrackedChanges: parsed.HasTrackedChanges ?? parsed.hasTrackedChanges,
    hasComments: parsed.HasComments ?? parsed.hasComments,
    estimatedPageCount: parsed.EstimatedPageCount ?? parsed.estimatedPageCount,
  };
}

/**
 * Render a specific page range for lazy loading/virtual scrolling.
 * Use getDocumentMetadata() first to get page count and dimensions for placeholder sizing.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param startPage - 1-based start page number
 * @param endPage - 1-based end page number (inclusive)
 * @param options - Rendering options
 * @returns HTML string containing only the requested page range
 * @throws Error if rendering fails
 *
 * @example
 * ```typescript
 * // First get metadata to understand the document
 * const metadata = await getDocumentMetadata(docxFile);
 * console.log(`Document has ${metadata.estimatedPageCount} estimated pages`);
 *
 * // Render just pages 1-3
 * const html = await renderPageRange(docxFile, 1, 3);
 *
 * // Render with options
 * const html = await renderPageRange(docxFile, 5, 10, {
 *   paginationScale: 0.8,
 *   renderHeadersAndFooters: true
 * });
 *
 * // The returned HTML contains data attributes for the page range:
 * // - data-start-page, data-end-page: requested range
 * // - data-total-pages: total estimated pages
 * // - data-start-block, data-end-block: content block indices
 * ```
 */
export async function renderPageRange(
  document: File | Uint8Array,
  startPage: number,
  endPage: number,
  options?: RenderPageRangeOptions
): Promise<string> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before heavy WASM work - allows loading states to render
  await yieldToMain();

  let result: string;

  // Check if any advanced options are specified
  const needsFullMethod = options?.renderTrackedChanges !== undefined ||
    options?.showDeletedContent !== undefined ||
    options?.renderComments !== undefined ||
    options?.additionalCss !== undefined;

  if (needsFullMethod) {
    result = exports.DocumentConverter.RenderPageRangeFull(
      bytes,
      startPage,
      endPage,
      options?.pageTitle ?? "Document",
      options?.cssPrefix ?? "docx-",
      options?.fabricateClasses ?? true,
      options?.additionalCss ?? "",
      options?.paginationScale ?? 1.0,
      options?.paginationCssClassPrefix ?? "page-",
      options?.renderFootnotesAndEndnotes ?? false,
      options?.renderHeadersAndFooters ?? false,
      options?.renderTrackedChanges ?? false,
      options?.showDeletedContent ?? true,
      options?.renderComments ?? false,
      options?.commentRenderMode ?? CommentRenderMode.EndnoteStyle
    );
  } else {
    result = exports.DocumentConverter.RenderPageRange(
      bytes,
      startPage,
      endPage,
      options?.pageTitle ?? "Document",
      options?.cssPrefix ?? "docx-",
      options?.fabricateClasses ?? true,
      options?.paginationScale ?? 1.0,
      options?.paginationCssClassPrefix ?? "page-",
      options?.renderFootnotesAndEndnotes ?? false,
      options?.renderHeadersAndFooters ?? false
    );
  }

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Page range rendering failed: ${error.error}`);
  }

  return result;
}

/**
 * Add an annotation using flexible targeting (element ID, indices, or text search).
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param request - Annotation details with target specification
 * @returns Response with modified document bytes and annotation info
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * // First get the document structure to find target elements
 * const structure = await getDocumentStructure(docxFile);
 *
 * // Annotate a specific paragraph by element ID
 * const result1 = await addAnnotationWithTarget(docxFile, {
 *   id: "annot-1",
 *   labelId: "INTRO",
 *   label: "Introduction",
 *   color: "#4CAF50",
 *   target: targetElement("doc/p-0")
 * });
 *
 * // Annotate a table cell
 * const result2 = await addAnnotationWithTarget(docxFile, {
 *   id: "annot-2",
 *   labelId: "CELL_HIGHLIGHT",
 *   label: "Important Cell",
 *   color: "#FFEB3B",
 *   target: targetTableCell(0, 1, 2)  // Table 0, Row 1, Cell 2
 * });
 *
 * // Annotate a table column
 * const result3 = await addAnnotationWithTarget(docxFile, {
 *   id: "annot-3",
 *   labelId: "COLUMN_DATA",
 *   label: "Values Column",
 *   color: "#2196F3",
 *   target: targetTableColumn(0, 1)  // Table 0, Column 1
 * });
 *
 * // Search for text within a specific element
 * const result4 = await addAnnotationWithTarget(docxFile, {
 *   id: "annot-4",
 *   labelId: "KEYWORD",
 *   label: "Keyword",
 *   color: "#FF5722",
 *   target: targetSearchInElement("doc/p-2", "important", 1)
 * });
 * ```
 */
export async function addAnnotationWithTarget(
  document: File | Uint8Array,
  request: AddAnnotationWithTargetRequest
): Promise<AddAnnotationResponse> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const requestJson = JSON.stringify({
    Id: request.id,
    LabelId: request.labelId,
    Label: request.label,
    Color: request.color ?? "#FFEB3B",
    Author: request.author,
    Metadata: request.metadata,
    ElementId: request.target.elementId,
    ElementType: request.target.elementType,
    ParagraphIndex: request.target.paragraphIndex,
    RunIndex: request.target.runIndex,
    TableIndex: request.target.tableIndex,
    RowIndex: request.target.rowIndex,
    CellIndex: request.target.cellIndex,
    ColumnIndex: request.target.columnIndex,
    SearchText: request.target.searchText,
    Occurrence: request.target.occurrence ?? 1,
    RangeEndParagraphIndex: request.target.rangeEndParagraphIndex,
  });

  const result = exports.DocumentConverter.AddAnnotationWithTarget(bytes, requestJson);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to add annotation: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  const annotation = parsed.Annotation || parsed.annotation;

  return {
    success: parsed.Success ?? parsed.success ?? true,
    documentBytes: parsed.DocumentBytes || parsed.documentBytes,
    annotation: annotation ? {
      id: annotation.Id || annotation.id,
      labelId: annotation.LabelId || annotation.labelId,
      label: annotation.Label || annotation.label,
      color: annotation.Color || annotation.color,
      author: annotation.Author || annotation.author,
      created: annotation.Created || annotation.created,
      bookmarkName: annotation.BookmarkName || annotation.bookmarkName,
      annotatedText: annotation.AnnotatedText || annotation.annotatedText,
      metadata: annotation.Metadata || annotation.metadata,
    } : undefined,
  };
}

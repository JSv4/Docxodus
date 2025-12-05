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
  // OpenContracts export types
  OpenContractDocExport,
  PawlsPage,
  PawlsPageBoundary,
  PawlsToken,
  OpenContractsAnnotation,
  OpenContractsSinglePageAnnotation,
  BoundingBox,
  TokenId,
  TextSpan,
  OpenContractsRelationship,
  // External annotation types
  AnnotationLabel,
  ExternalAnnotationSet,
  ExternalAnnotationValidationResult,
  ExternalAnnotationValidationIssue,
  ExternalAnnotationProjectionSettings,
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
  // Document metadata (useful for info)
  DocumentMetadata,
  SectionMetadata,
  // OpenContracts export types
  OpenContractDocExport,
  PawlsPage,
  PawlsPageBoundary,
  PawlsToken,
  OpenContractsAnnotation,
  OpenContractsSinglePageAnnotation,
  BoundingBox,
  TokenId,
  TextSpan,
  OpenContractsRelationship,
  // External annotation types
  AnnotationLabel,
  ExternalAnnotationSet,
  ExternalAnnotationValidationResult,
  ExternalAnnotationValidationIssue,
  ExternalAnnotationProjectionSettings,
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
 * Export document to OpenContracts format.
 *
 * This provides complete document text, structure, and layout information
 * compatible with the OpenContracts ecosystem for document analysis.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns OpenContractDocExport with complete document data
 * @throws Error if export fails
 *
 * @example
 * ```typescript
 * const result = await exportToOpenContract(docxFile);
 *
 * // Access complete document text
 * console.log(`Content length: ${result.content.length} characters`);
 *
 * // Get document structure
 * console.log(`Pages: ${result.pageCount}`);
 * console.log(`Structural annotations: ${result.labelledText.filter(a => a.structural).length}`);
 *
 * // Access PAWLS layout data
 * for (const page of result.pawlsFileContent) {
 *   console.log(`Page ${page.page.index}: ${page.tokens.length} tokens`);
 * }
 * ```
 */
export async function exportToOpenContract(
  document: File | Uint8Array
): Promise<OpenContractDocExport> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const result = exports.DocumentConverter.ExportToOpenContract(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to export to OpenContract format: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  // Convert from PascalCase to camelCase
  const convertPawlsPage = (p: any): PawlsPage => ({
    page: {
      width: p.Page?.Width ?? p.page?.width,
      height: p.Page?.Height ?? p.page?.height,
      index: p.Page?.Index ?? p.page?.index,
    },
    tokens: (p.Tokens || p.tokens || []).map((t: any) => ({
      x: t.X ?? t.x,
      y: t.Y ?? t.y,
      width: t.Width ?? t.width,
      height: t.Height ?? t.height,
      text: t.Text ?? t.text,
    })),
  });

  const convertAnnotation = (a: any): OpenContractsAnnotation => ({
    id: a.Id ?? a.id,
    annotationLabel: a.AnnotationLabel ?? a.annotationLabel,
    rawText: a.RawText ?? a.rawText,
    page: a.Page ?? a.page,
    annotationJson: convertAnnotationJson(a.AnnotationJson ?? a.annotationJson),
    parentId: a.ParentId ?? a.parentId,
    annotationType: a.AnnotationType ?? a.annotationType,
    structural: a.Structural ?? a.structural,
  });

  const convertAnnotationJson = (json: any): TextSpan | Record<string, OpenContractsSinglePageAnnotation> | undefined => {
    if (!json) return undefined;

    // Check if it's a TextSpan
    if (json.Start !== undefined || json.start !== undefined) {
      return {
        id: json.Id ?? json.id,
        start: json.Start ?? json.start,
        end: json.End ?? json.end,
        text: json.Text ?? json.text,
      };
    }

    // Otherwise it's a dictionary of single-page annotations
    const result: Record<string, OpenContractsSinglePageAnnotation> = {};
    for (const [key, value] of Object.entries(json)) {
      const v = value as any;
      result[key] = {
        bounds: {
          top: v.Bounds?.Top ?? v.bounds?.top,
          bottom: v.Bounds?.Bottom ?? v.bounds?.bottom,
          left: v.Bounds?.Left ?? v.bounds?.left,
          right: v.Bounds?.Right ?? v.bounds?.right,
        },
        tokensJsons: (v.TokensJsons || v.tokensJsons || []).map((t: any) => ({
          pageIndex: t.PageIndex ?? t.pageIndex,
          tokenIndex: t.TokenIndex ?? t.tokenIndex,
        })),
        rawText: v.RawText ?? v.rawText,
      };
    }
    return result;
  };

  const convertRelationship = (r: any): OpenContractsRelationship => ({
    id: r.Id ?? r.id,
    relationshipLabel: r.RelationshipLabel ?? r.relationshipLabel,
    sourceAnnotationIds: r.SourceAnnotationIds ?? r.sourceAnnotationIds ?? [],
    targetAnnotationIds: r.TargetAnnotationIds ?? r.targetAnnotationIds ?? [],
    structural: r.Structural ?? r.structural,
  });

  return {
    title: parsed.Title ?? parsed.title,
    content: parsed.Content ?? parsed.content,
    description: parsed.Description ?? parsed.description,
    pageCount: parsed.PageCount ?? parsed.pageCount,
    pawlsFileContent: (parsed.PawlsFileContent || parsed.pawlsFileContent || []).map(convertPawlsPage),
    docLabels: parsed.DocLabels ?? parsed.docLabels ?? [],
    labelledText: (parsed.LabelledText || parsed.labelledText || []).map(convertAnnotation),
    relationships: (parsed.Relationships || parsed.relationships)?.map(convertRelationship),
  };
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

// ============================================================================
// External Annotation Functions (Issue #57)
// ============================================================================

/**
 * Compute the SHA256 hash of a document for integrity validation.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @returns SHA256 hash as lowercase hex string
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const hash = await computeDocumentHash(docxFile);
 * console.log(`Document hash: ${hash}`);
 *
 * // Later, verify the document hasn't changed
 * const currentHash = await computeDocumentHash(docxFile);
 * if (currentHash !== storedHash) {
 *   console.log("Document has been modified");
 * }
 * ```
 */
export async function computeDocumentHash(
  document: File | Uint8Array
): Promise<string> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  const result = exports.DocumentConverter.ComputeDocumentHash(bytes);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to compute document hash: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return parsed.Hash ?? parsed.hash;
}

/**
 * Create an ExternalAnnotationSet from a document.
 * This extracts the document structure and computes the hash for integrity validation.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param documentId - Unique identifier for the document (filename, UUID, etc.)
 * @returns ExternalAnnotationSet ready for adding annotations
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * // Create an annotation set
 * const set = await createExternalAnnotationSet(docxFile, "contract-v1.0");
 *
 * // Access document text for searching
 * console.log(`Document length: ${set.content.length} chars`);
 *
 * // Add label definitions
 * set.textLabels["IMPORTANT"] = {
 *   id: "IMPORTANT",
 *   text: "Important",
 *   color: "#FF0000",
 *   description: "Important text",
 *   icon: "",
 *   labelType: "text"
 * };
 *
 * // Create annotations using the content
 * const annotation = createAnnotationFromSearch(
 *   "ann-001", "IMPORTANT", set.content, "shall not be liable"
 * );
 * if (annotation) {
 *   set.labelledText.push(annotation);
 * }
 *
 * // Serialize for storage
 * const json = JSON.stringify(set);
 * ```
 */
export async function createExternalAnnotationSet(
  document: File | Uint8Array,
  documentId: string
): Promise<ExternalAnnotationSet> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const result = exports.DocumentConverter.CreateExternalAnnotationSet(bytes, documentId);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to create external annotation set: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  // Convert from PascalCase to camelCase
  return convertExternalAnnotationSet(parsed);
}

/**
 * Validate an external annotation set against a document.
 * Checks hash match and verifies each annotation's text still matches.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param annotationSet - The annotation set to validate
 * @returns Validation result with any issues found
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const result = await validateExternalAnnotations(docxFile, annotationSet);
 *
 * if (!result.isValid) {
 *   if (result.hashMismatch) {
 *     console.log("Document has been modified since annotations were created");
 *   }
 *   for (const issue of result.issues) {
 *     console.log(`${issue.issueType}: ${issue.description}`);
 *   }
 * }
 * ```
 */
export async function validateExternalAnnotations(
  document: File | Uint8Array,
  annotationSet: ExternalAnnotationSet
): Promise<ExternalAnnotationValidationResult> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const annotationSetJson = JSON.stringify(annotationSet);
  const result = exports.DocumentConverter.ValidateExternalAnnotations(bytes, annotationSetJson);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to validate external annotations: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  return {
    isValid: parsed.IsValid ?? parsed.isValid,
    hashMismatch: parsed.HashMismatch ?? parsed.hashMismatch,
    issues: (parsed.Issues || parsed.issues || []).map((i: any) => ({
      annotationId: i.AnnotationId ?? i.annotationId,
      issueType: i.IssueType ?? i.issueType,
      description: i.Description ?? i.description,
      expectedText: i.ExpectedText ?? i.expectedText,
      actualText: i.ActualText ?? i.actualText,
    })),
  };
}

/**
 * Convert a DOCX document to HTML with external annotations projected.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param annotationSet - The external annotation set to project
 * @param conversionOptions - HTML conversion options
 * @param projectionOptions - Annotation projection options
 * @returns HTML string with annotations projected
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * // Basic usage
 * const html = await convertDocxToHtmlWithExternalAnnotations(
 *   docxFile,
 *   annotationSet
 * );
 *
 * // With custom options
 * const html = await convertDocxToHtmlWithExternalAnnotations(
 *   docxFile,
 *   annotationSet,
 *   { pageTitle: "Annotated Document" },
 *   { labelMode: AnnotationLabelMode.Inline, cssClassPrefix: "my-annot-" }
 * );
 * ```
 */
export async function convertDocxToHtmlWithExternalAnnotations(
  document: File | Uint8Array,
  annotationSet: ExternalAnnotationSet,
  conversionOptions?: ConversionOptions,
  projectionOptions?: ExternalAnnotationProjectionSettings
): Promise<string> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const annotationSetJson = JSON.stringify(annotationSet);
  const result = exports.DocumentConverter.ConvertDocxToHtmlWithExternalAnnotations(
    bytes,
    annotationSetJson,
    conversionOptions?.pageTitle ?? "Document",
    conversionOptions?.cssPrefix ?? "docx-",
    conversionOptions?.fabricateClasses ?? true,
    conversionOptions?.additionalCss ?? "",
    projectionOptions?.cssClassPrefix ?? "ext-annot-",
    projectionOptions?.labelMode ?? AnnotationLabelMode.Above
  );

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to convert with external annotations: ${error.error}`);
  }

  const parsed = JSON.parse(result);
  return parsed.Html ?? parsed.html;
}

/**
 * Search for text in a document and return character offsets.
 * Useful for finding text locations to create annotations.
 *
 * @param document - DOCX file as File object or Uint8Array
 * @param searchText - Text to search for
 * @param maxResults - Maximum number of results (default: 100)
 * @returns Array of TextSpan objects with offsets
 * @throws Error if operation fails
 *
 * @example
 * ```typescript
 * const occurrences = await searchTextOffsets(docxFile, "liability");
 * console.log(`Found ${occurrences.length} occurrences`);
 *
 * for (const span of occurrences) {
 *   console.log(`"${span.text}" at offset ${span.start}-${span.end}`);
 * }
 * ```
 */
export async function searchTextOffsets(
  document: File | Uint8Array,
  searchText: string,
  maxResults: number = 100
): Promise<TextSpan[]> {
  const exports = ensureInitialized();
  const bytes = await toBytes(document);

  // Yield to browser before WASM work - allows loading states to render
  await yieldToMain();

  const result = exports.DocumentConverter.SearchTextOffsets(bytes, searchText, maxResults);

  if (isErrorResponse(result)) {
    const error = parseError(result);
    throw new Error(`Failed to search text: ${error.error}`);
  }

  const parsed = JSON.parse(result);

  return (parsed.Results || parsed.results || []).map((r: any) => ({
    id: r.Id ?? r.id,
    start: r.Start ?? r.start,
    end: r.End ?? r.end,
    text: r.Text ?? r.text,
  }));
}

/**
 * Create an annotation from character offsets.
 * This is a client-side helper - no WASM call needed.
 *
 * @param id - Unique identifier for the annotation
 * @param labelId - Label/category ID for the annotation
 * @param documentText - Full document text (from annotationSet.content)
 * @param startOffset - Start character offset (0-indexed, inclusive)
 * @param endOffset - End character offset (exclusive)
 * @returns OpenContractsAnnotation ready to add to an annotation set
 * @throws Error if offsets are invalid
 *
 * @example
 * ```typescript
 * const set = await createExternalAnnotationSet(docxFile, "doc-1");
 * const annotation = createAnnotation("ann-001", "IMPORTANT", set.content, 100, 150);
 * set.labelledText.push(annotation);
 * ```
 */
export function createAnnotation(
  id: string,
  labelId: string,
  documentText: string,
  startOffset: number,
  endOffset: number
): OpenContractsAnnotation {
  if (startOffset < 0) {
    throw new Error("Start offset must be non-negative");
  }
  if (endOffset < startOffset) {
    throw new Error("End offset must be >= start offset");
  }
  if (endOffset > documentText.length) {
    throw new Error("End offset exceeds document length");
  }

  const rawText = documentText.substring(startOffset, endOffset);

  return {
    id,
    annotationLabel: labelId,
    rawText,
    page: 0,
    annotationJson: {
      id,
      start: startOffset,
      end: endOffset,
      text: rawText,
    },
    annotationType: "text",
    structural: false,
  };
}

/**
 * Create an annotation by searching for text in the document.
 * This is a client-side helper - no WASM call needed.
 *
 * @param id - Unique identifier for the annotation
 * @param labelId - Label/category ID for the annotation
 * @param documentText - Full document text (from annotationSet.content)
 * @param searchText - Text to search for
 * @param occurrence - Which occurrence to use (1-based, default: 1)
 * @returns OpenContractsAnnotation, or null if text not found
 *
 * @example
 * ```typescript
 * const set = await createExternalAnnotationSet(docxFile, "doc-1");
 *
 * // Find first occurrence
 * const ann1 = createAnnotationFromSearch("ann-001", "LIABILITY", set.content, "shall not be liable");
 * if (ann1) set.labelledText.push(ann1);
 *
 * // Find second occurrence
 * const ann2 = createAnnotationFromSearch("ann-002", "LIABILITY", set.content, "shall not be liable", 2);
 * if (ann2) set.labelledText.push(ann2);
 * ```
 */
export function createAnnotationFromSearch(
  id: string,
  labelId: string,
  documentText: string,
  searchText: string,
  occurrence: number = 1
): OpenContractsAnnotation | null {
  if (occurrence < 1) {
    throw new Error("Occurrence must be >= 1");
  }

  const offsets = findTextOccurrences(documentText, searchText);

  if (occurrence > offsets.length) {
    return null;
  }

  const { start, end } = offsets[occurrence - 1];
  return createAnnotation(id, labelId, documentText, start, end);
}

/**
 * Find all occurrences of a text string in the document.
 * This is a client-side helper - no WASM call needed.
 *
 * @param documentText - Full document text
 * @param searchText - Text to search for
 * @param maxResults - Maximum number of results (default: 100)
 * @returns Array of { start, end } offsets
 *
 * @example
 * ```typescript
 * const occurrences = findTextOccurrences(set.content, "the");
 * console.log(`Found ${occurrences.length} occurrences of "the"`);
 * ```
 */
export function findTextOccurrences(
  documentText: string,
  searchText: string,
  maxResults: number = 100
): Array<{ start: number; end: number }> {
  if (!searchText) return [];

  const results: Array<{ start: number; end: number }> = [];
  let index = 0;

  while (results.length < maxResults) {
    index = documentText.indexOf(searchText, index);
    if (index < 0) break;

    results.push({ start: index, end: index + searchText.length });
    index += 1; // Move past start to find overlapping matches
  }

  return results;
}

// Helper function to convert PascalCase response to camelCase ExternalAnnotationSet
function convertExternalAnnotationSet(parsed: any): ExternalAnnotationSet {
  const convertLabel = (l: any): AnnotationLabel => ({
    id: l.Id ?? l.id,
    color: l.Color ?? l.color,
    description: l.Description ?? l.description ?? "",
    icon: l.Icon ?? l.icon ?? "",
    text: l.Text ?? l.text,
    labelType: l.LabelType ?? l.labelType ?? "text",
  });

  const convertPawlsPage = (p: any): PawlsPage => ({
    page: {
      width: p.Page?.Width ?? p.page?.width,
      height: p.Page?.Height ?? p.page?.height,
      index: p.Page?.Index ?? p.page?.index,
    },
    tokens: (p.Tokens || p.tokens || []).map((t: any) => ({
      x: t.X ?? t.x,
      y: t.Y ?? t.y,
      width: t.Width ?? t.width,
      height: t.Height ?? t.height,
      text: t.Text ?? t.text,
    })),
  });

  const convertAnnotation = (a: any): OpenContractsAnnotation => ({
    id: a.Id ?? a.id,
    annotationLabel: a.AnnotationLabel ?? a.annotationLabel,
    rawText: a.RawText ?? a.rawText,
    page: a.Page ?? a.page,
    annotationJson: convertAnnotationJson(a.AnnotationJson ?? a.annotationJson),
    parentId: a.ParentId ?? a.parentId,
    annotationType: a.AnnotationType ?? a.annotationType,
    structural: a.Structural ?? a.structural,
  });

  const convertAnnotationJson = (json: any): TextSpan | Record<string, OpenContractsSinglePageAnnotation> | undefined => {
    if (!json) return undefined;

    // Check if it's a TextSpan
    if (json.Start !== undefined || json.start !== undefined) {
      return {
        id: json.Id ?? json.id,
        start: json.Start ?? json.start,
        end: json.End ?? json.end,
        text: json.Text ?? json.text,
      };
    }

    // Otherwise it's a dictionary of single-page annotations
    const result: Record<string, OpenContractsSinglePageAnnotation> = {};
    for (const [key, value] of Object.entries(json)) {
      const v = value as any;
      result[key] = {
        bounds: {
          top: v.Bounds?.Top ?? v.bounds?.top,
          bottom: v.Bounds?.Bottom ?? v.bounds?.bottom,
          left: v.Bounds?.Left ?? v.bounds?.left,
          right: v.Bounds?.Right ?? v.bounds?.right,
        },
        tokensJsons: (v.TokensJsons || v.tokensJsons || []).map((t: any) => ({
          pageIndex: t.PageIndex ?? t.pageIndex,
          tokenIndex: t.TokenIndex ?? t.tokenIndex,
        })),
        rawText: v.RawText ?? v.rawText,
      };
    }
    return result;
  };

  const convertRelationship = (r: any): OpenContractsRelationship => ({
    id: r.Id ?? r.id,
    relationshipLabel: r.RelationshipLabel ?? r.relationshipLabel,
    sourceAnnotationIds: r.SourceAnnotationIds ?? r.sourceAnnotationIds ?? [],
    targetAnnotationIds: r.TargetAnnotationIds ?? r.targetAnnotationIds ?? [],
    structural: r.Structural ?? r.structural,
  });

  // Convert label dictionaries
  const textLabels: Record<string, AnnotationLabel> = {};
  const rawTextLabels = parsed.TextLabels || parsed.textLabels || {};
  for (const [key, value] of Object.entries(rawTextLabels)) {
    textLabels[key] = convertLabel(value);
  }

  const docLabelDefinitions: Record<string, AnnotationLabel> = {};
  const rawDocLabelDefs = parsed.DocLabelDefinitions || parsed.docLabelDefinitions || {};
  for (const [key, value] of Object.entries(rawDocLabelDefs)) {
    docLabelDefinitions[key] = convertLabel(value);
  }

  return {
    documentId: parsed.DocumentId ?? parsed.documentId,
    documentHash: parsed.DocumentHash ?? parsed.documentHash,
    createdAt: parsed.CreatedAt ?? parsed.createdAt,
    updatedAt: parsed.UpdatedAt ?? parsed.updatedAt,
    version: parsed.Version ?? parsed.version,
    title: parsed.Title ?? parsed.title,
    content: parsed.Content ?? parsed.content,
    description: parsed.Description ?? parsed.description,
    pageCount: parsed.PageCount ?? parsed.pageCount,
    pawlsFileContent: (parsed.PawlsFileContent || parsed.pawlsFileContent || []).map(convertPawlsPage),
    docLabels: parsed.DocLabels ?? parsed.docLabels ?? [],
    labelledText: (parsed.LabelledText || parsed.labelledText || []).map(convertAnnotation),
    relationships: (parsed.Relationships || parsed.relationships)?.map(convertRelationship),
    textLabels,
    docLabelDefinitions,
  };
}


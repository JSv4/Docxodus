/**
 * Docxodus Web Worker
 *
 * This worker runs the WASM runtime in a separate thread, keeping the main
 * thread free for UI updates and user interactions.
 *
 * Communication is via postMessage with structured request/response types.
 * Document bytes are transferred (not copied) for efficiency.
 */

import type {
  WorkerRequest,
  WorkerResponse,
  WorkerInitRequest,
  WorkerConvertRequest,
  WorkerCompareRequest,
  WorkerCompareToHtmlRequest,
  WorkerGetRevisionsRequest,
  WorkerGetDocumentMetadataRequest,
  DocxodusWasmExports,
  ConversionOptions,
  CompareOptions,
  GetRevisionsOptions,
  Revision,
  RevisionType,
  DocumentMetadata,
  SectionMetadata,
} from "./types.js";

// Worker-local state
let wasmExports: DocxodusWasmExports | null = null;
let initPromise: Promise<void> | null = null;

/**
 * Initialize the WASM runtime in the worker.
 */
async function initializeWasm(basePath: string): Promise<void> {
  if (wasmExports) {
    return; // Already initialized
  }

  if (initPromise) {
    return initPromise; // Initialization in progress
  }

  initPromise = (async () => {
    try {
      // Normalize base path
      const normalizedPath = basePath.endsWith("/") ? basePath : basePath + "/";

      // Import dotnet.js from the WASM directory
      // In a worker, we need to use importScripts or dynamic import
      const dotnetModule = await import(
        /* webpackIgnore: true */ `${normalizedPath}_framework/dotnet.js`
      );

      const { getAssemblyExports, getConfig } = await dotnetModule.dotnet
        .withDiagnosticTracing(false)
        .create();

      const config = getConfig();
      const exports = await getAssemblyExports(config.mainAssemblyName);

      // Map exports to the expected structure (same as index.ts)
      wasmExports = {
        DocumentConverter: (exports as any).DocxodusWasm.DocumentConverter,
        DocumentComparer: (exports as any).DocxodusWasm.DocumentComparer,
      };
    } catch (error) {
      initPromise = null;
      throw error;
    }
  })();

  return initPromise;
}

/**
 * Ensure WASM is initialized, throwing if not.
 */
function ensureInitialized(): DocxodusWasmExports {
  if (!wasmExports) {
    throw new Error("Worker not initialized. Call init first.");
  }
  return wasmExports;
}

/**
 * Check if a result string is an error response.
 */
function isErrorResponse(result: string): boolean {
  return result.startsWith("{") && result.includes('"Error"');
}

/**
 * Parse an error response.
 */
function parseError(result: string): { error: string } {
  try {
    const parsed = JSON.parse(result);
    return { error: parsed.Error || parsed.error || "Unknown error" };
  } catch {
    return { error: result };
  }
}

/**
 * Handle convertDocxToHtml request.
 */
function handleConvert(
  request: WorkerConvertRequest
): { html?: string; error?: string } {
  const exports = ensureInitialized();
  const options = request.options;

  try {
    let result: string;

    // Check if any of the complete options are specified
    const needsCompleteMethod =
      options?.renderFootnotesAndEndnotes !== undefined ||
      options?.renderHeadersAndFooters !== undefined ||
      options?.renderTrackedChanges !== undefined ||
      options?.showDeletedContent !== undefined ||
      options?.renderMoveOperations !== undefined;

    if (needsCompleteMethod || options?.renderAnnotations) {
      result = exports.DocumentConverter.ConvertDocxToHtmlComplete(
        request.documentBytes,
        options?.pageTitle ?? "Document",
        options?.cssPrefix ?? "docx-",
        options?.fabricateClasses ?? true,
        options?.additionalCss ?? "",
        options?.commentRenderMode ?? -1,
        options?.commentCssClassPrefix ?? "comment-",
        options?.paginationMode ?? 0,
        options?.paginationScale ?? 1.0,
        options?.paginationCssClassPrefix ?? "page-",
        options?.renderAnnotations ?? false,
        options?.annotationLabelMode ?? 0,
        options?.annotationCssClassPrefix ?? "annot-",
        options?.renderFootnotesAndEndnotes ?? false,
        options?.renderHeadersAndFooters ?? false,
        options?.renderTrackedChanges ?? false,
        options?.showDeletedContent ?? true,
        options?.renderMoveOperations ?? true
      );
    } else if (
      options?.paginationMode !== undefined &&
      options.paginationMode !== 0
    ) {
      result = exports.DocumentConverter.ConvertDocxToHtmlWithPagination(
        request.documentBytes,
        options.pageTitle ?? "Document",
        options.cssPrefix ?? "docx-",
        options.fabricateClasses ?? true,
        options.additionalCss ?? "",
        options.commentRenderMode ?? -1,
        options.commentCssClassPrefix ?? "comment-",
        options.paginationMode,
        options.paginationScale ?? 1.0,
        options.paginationCssClassPrefix ?? "page-"
      );
    } else if (options) {
      result = exports.DocumentConverter.ConvertDocxToHtmlWithOptions(
        request.documentBytes,
        options.pageTitle ?? "Document",
        options.cssPrefix ?? "docx-",
        options.fabricateClasses ?? true,
        options.additionalCss ?? "",
        options.commentRenderMode ?? -1,
        options.commentCssClassPrefix ?? "comment-"
      );
    } else {
      result = exports.DocumentConverter.ConvertDocxToHtml(
        request.documentBytes
      );
    }

    if (isErrorResponse(result)) {
      return parseError(result);
    }

    return { html: result };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Handle compareDocuments request.
 */
function handleCompare(
  request: WorkerCompareRequest
): { documentBytes?: Uint8Array; error?: string } {
  const exports = ensureInitialized();
  const options = request.options;

  try {
    let result: Uint8Array;

    if (options?.detailThreshold !== undefined || options?.caseInsensitive) {
      result = exports.DocumentComparer.CompareDocumentsWithOptions(
        request.originalBytes,
        request.modifiedBytes,
        options?.authorName ?? "Docxodus",
        options?.detailThreshold ?? 0.15,
        options?.caseInsensitive ?? false
      );
    } else {
      result = exports.DocumentComparer.CompareDocuments(
        request.originalBytes,
        request.modifiedBytes,
        options?.authorName ?? "Docxodus"
      );
    }

    if (result.length === 0) {
      return { error: "Comparison returned empty result" };
    }

    return { documentBytes: result };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Handle compareDocumentsToHtml request.
 */
function handleCompareToHtml(
  request: WorkerCompareToHtmlRequest
): { html?: string; error?: string } {
  const exports = ensureInitialized();
  const options = request.options;

  try {
    const renderTrackedChanges = options?.renderTrackedChanges ?? true;

    const result = exports.DocumentComparer.CompareDocumentsToHtmlWithOptions(
      request.originalBytes,
      request.modifiedBytes,
      options?.authorName ?? "Docxodus",
      renderTrackedChanges
    );

    if (isErrorResponse(result)) {
      return parseError(result);
    }

    return { html: result };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Handle getRevisions request.
 */
function handleGetRevisions(
  request: WorkerGetRevisionsRequest
): { revisions?: Revision[]; error?: string } {
  const exports = ensureInitialized();
  const options = request.options;

  try {
    const detectMoves = options?.detectMoves ?? true;
    const moveSimilarityThreshold = options?.moveSimilarityThreshold ?? 0.8;
    const moveMinimumWordCount = options?.moveMinimumWordCount ?? 3;
    const caseInsensitive = options?.caseInsensitive ?? false;

    const result = exports.DocumentComparer.GetRevisionsJsonWithOptions(
      request.documentBytes,
      detectMoves,
      moveSimilarityThreshold,
      moveMinimumWordCount,
      caseInsensitive
    );

    if (isErrorResponse(result)) {
      return parseError(result);
    }

    const parsed = JSON.parse(result);
    const revisions = (parsed.Revisions || parsed.revisions || []).map(
      (r: any): Revision => ({
        author: r.Author || r.author,
        date: r.Date || r.date,
        revisionType: r.RevisionType || r.revisionType,
        text: r.Text || r.text,
        moveGroupId: r.MoveGroupId ?? r.moveGroupId,
        isMoveSource: r.IsMoveSource ?? r.isMoveSource,
        formatChange: r.FormatChange || r.formatChange
          ? {
              oldProperties:
                r.FormatChange?.OldProperties ||
                r.formatChange?.oldProperties,
              newProperties:
                r.FormatChange?.NewProperties ||
                r.formatChange?.newProperties,
              changedPropertyNames:
                r.FormatChange?.ChangedPropertyNames ||
                r.formatChange?.changedPropertyNames,
            }
          : undefined,
      })
    );

    return { revisions };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Handle getDocumentMetadata request.
 */
function handleGetDocumentMetadata(
  request: WorkerGetDocumentMetadataRequest
): { metadata?: DocumentMetadata; error?: string } {
  const exports = ensureInitialized();

  try {
    const result = exports.DocumentConverter.GetDocumentMetadata(
      request.documentBytes
    );

    if (isErrorResponse(result)) {
      return parseError(result);
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

    const metadata: DocumentMetadata = {
      sections: (parsed.Sections || parsed.sections || []).map(convertSection),
      totalParagraphs: parsed.TotalParagraphs ?? parsed.totalParagraphs,
      totalTables: parsed.TotalTables ?? parsed.totalTables,
      hasFootnotes: parsed.HasFootnotes ?? parsed.hasFootnotes,
      hasEndnotes: parsed.HasEndnotes ?? parsed.hasEndnotes,
      hasTrackedChanges: parsed.HasTrackedChanges ?? parsed.hasTrackedChanges,
      hasComments: parsed.HasComments ?? parsed.hasComments,
      estimatedPageCount: parsed.EstimatedPageCount ?? parsed.estimatedPageCount,
    };

    return { metadata };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Handle getVersion request.
 */
function handleGetVersion(): {
  version?: { library: string; dotnetVersion: string; platform: string };
  error?: string;
} {
  const exports = ensureInitialized();

  try {
    const result = exports.DocumentConverter.GetVersion();
    const parsed = JSON.parse(result);

    return {
      version: {
        library: parsed.Library || parsed.library,
        dotnetVersion: parsed.DotnetVersion || parsed.dotnetVersion,
        platform: parsed.Platform || parsed.platform,
      },
    };
  } catch (error) {
    return { error: String(error) };
  }
}

/**
 * Main message handler.
 */
self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  const request = event.data;

  try {
    let response: WorkerResponse;

    switch (request.type) {
      case "init": {
        const initRequest = request as WorkerInitRequest;
        try {
          await initializeWasm(initRequest.wasmBasePath);
          response = {
            id: request.id,
            type: "init",
            success: true,
          };
        } catch (error) {
          response = {
            id: request.id,
            type: "init",
            success: false,
            error: String(error),
          };
        }
        break;
      }

      case "convertDocxToHtml": {
        const convertRequest = request as WorkerConvertRequest;
        const result = handleConvert(convertRequest);
        response = {
          id: request.id,
          type: "convertDocxToHtml",
          success: !result.error,
          html: result.html,
          error: result.error,
        };
        break;
      }

      case "compareDocuments": {
        const compareRequest = request as WorkerCompareRequest;
        const result = handleCompare(compareRequest);
        response = {
          id: request.id,
          type: "compareDocuments",
          success: !result.error,
          documentBytes: result.documentBytes,
          error: result.error,
        };
        // Transfer the bytes back (zero-copy)
        if (result.documentBytes) {
          self.postMessage(response, { transfer: [result.documentBytes.buffer as ArrayBuffer] });
          return;
        }
        break;
      }

      case "compareDocumentsToHtml": {
        const compareToHtmlRequest = request as WorkerCompareToHtmlRequest;
        const result = handleCompareToHtml(compareToHtmlRequest);
        response = {
          id: request.id,
          type: "compareDocumentsToHtml",
          success: !result.error,
          html: result.html,
          error: result.error,
        };
        break;
      }

      case "getRevisions": {
        const getRevisionsRequest = request as WorkerGetRevisionsRequest;
        const result = handleGetRevisions(getRevisionsRequest);
        response = {
          id: request.id,
          type: "getRevisions",
          success: !result.error,
          revisions: result.revisions,
          error: result.error,
        };
        break;
      }

      case "getDocumentMetadata": {
        const getMetadataRequest = request as WorkerGetDocumentMetadataRequest;
        const result = handleGetDocumentMetadata(getMetadataRequest);
        response = {
          id: request.id,
          type: "getDocumentMetadata",
          success: !result.error,
          metadata: result.metadata,
          error: result.error,
        };
        break;
      }

      case "getVersion": {
        const result = handleGetVersion();
        response = {
          id: request.id,
          type: "getVersion",
          success: !result.error,
          version: result.version,
          error: result.error,
        };
        break;
      }

      default: {
        // This should never happen due to type narrowing, but handle gracefully
        const unknownRequest = request as { id: string; type: string };
        self.postMessage({
          id: unknownRequest.id,
          type: unknownRequest.type,
          success: false,
          error: `Unknown request type: ${unknownRequest.type}`,
        });
        return;
      }
    }

    self.postMessage(response);
  } catch (error) {
    // Catch-all for unexpected errors
    self.postMessage({
      id: request.id,
      type: request.type,
      success: false,
      error: String(error),
    } as WorkerResponse);
  }
};

// Signal that the worker is ready
self.postMessage({ type: "ready" });

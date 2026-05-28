/**
 * Worker Proxy - Main thread interface for the Docxodus Web Worker
 *
 * This module provides a Promise-based API that mirrors the main API but
 * executes all WASM operations in a Web Worker, keeping the main thread free.
 *
 * @example
 * ```typescript
 * import { createWorkerDocxodus } from 'docxodus/worker';
 *
 * // Create worker instance
 * const docxodus = await createWorkerDocxodus();
 *
 * // Use the same API as the main module, but non-blocking!
 * const html = await docxodus.convertDocxToHtml(docxFile);
 *
 * // Clean up when done
 * docxodus.terminate();
 * ```
 */

import type {
  WorkerRequest,
  WorkerResponse,
  WorkerConvertResponse,
  WorkerCompareResponse,
  WorkerCompareToHtmlResponse,
  WorkerGetRevisionsResponse,
  WorkerGetDocumentMetadataResponse,
  WorkerGetVersionResponse,
  WorkerSessionOpenResponse,
  WorkerSessionEditResponse,
  WorkerDocxodusOptions,
  ConversionOptions,
  CompareOptions,
  GetRevisionsOptions,
  Revision,
  VersionInfo,
  DocumentMetadata,
  DocxSessionSettings,
  DocumentAnnotation,
  AnnotationUpdate,
  CharSpan,
  EditResult,
} from "./types.js";

/**
 * Generate a unique request ID.
 */
function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 11)}`;
}

/**
 * Convert a File or Uint8Array to Uint8Array.
 */
async function toBytes(document: File | Uint8Array): Promise<Uint8Array> {
  if (document instanceof Uint8Array) {
    return document;
  }
  const buffer = await document.arrayBuffer();
  return new Uint8Array(buffer);
}

/**
 * Derive the WASM base path from the current module URL.
 */
function deriveWasmBasePath(): string {
  // Try to get the base path from the current script URL
  if (typeof document !== "undefined") {
    // Browser: look for docxodus script tag or use current location
    const scripts = document.querySelectorAll('script[src*="docxodus"]');
    if (scripts.length > 0) {
      const src = (scripts[0] as HTMLScriptElement).src;
      const base = src.substring(0, src.lastIndexOf("/") + 1);
      return base + "wasm/";
    }
  }

  // Default fallback
  return "/wasm/";
}

/**
 * A worker-proxied DocxSession. Mirrors the main-thread {@link DocxSession}
 * annotation write surface but each call returns a Promise, since the actual
 * work happens inside the Web Worker.
 *
 * Acquire via {@link WorkerDocxodus.openDocxSession}; always call
 * {@link close} when finished to free the in-worker handle.
 */
export interface WorkerDocxSession {
  /**
   * Add an annotation to the document at the given anchor.
   * @param anchorId - Markdown-projection anchor id of the target block
   * @param span - Character span within the block, or null for the whole block
   * @param annotation - Annotation data (id auto-generated if omitted)
   * @returns EditResult indicating success and any created/modified anchors
   */
  addAnnotation(
    anchorId: string,
    span: CharSpan | null,
    annotation: DocumentAnnotation
  ): Promise<EditResult>;

  /**
   * Remove an existing annotation by its id.
   * @param annotationId - The annotation id to remove
   * @returns EditResult indicating success
   */
  removeAnnotation(annotationId: string): Promise<EditResult>;

  /**
   * Partially update an annotation's metadata without moving it.
   * @param annotationId - The annotation id to update
   * @param update - Fields to change (omitted fields are left unchanged)
   * @returns EditResult indicating success
   */
  updateAnnotation(
    annotationId: string,
    update: AnnotationUpdate
  ): Promise<EditResult>;

  /**
   * Move an annotation to a new anchor/span position.
   * @param annotationId - The annotation id to move
   * @param newAnchorId - Target anchor id
   * @param newSpan - New character span, or null for the whole block
   * @returns EditResult indicating success
   */
  moveAnnotation(
    annotationId: string,
    newAnchorId: string,
    newSpan: CharSpan | null
  ): Promise<EditResult>;

  /**
   * Close the session and release its in-worker handle.
   * After calling this, the instance cannot be used anymore.
   */
  close(): Promise<void>;
}

/**
 * A worker-based Docxodus instance.
 *
 * Provides the same API as the main module but executes all operations
 * in a Web Worker for non-blocking UI.
 */
export interface WorkerDocxodus {
  /**
   * Convert a DOCX document to HTML.
   * @param document - DOCX file as File object or Uint8Array
   * @param options - Conversion options
   * @returns HTML string
   */
  convertDocxToHtml(
    document: File | Uint8Array,
    options?: ConversionOptions
  ): Promise<string>;

  /**
   * Compare two DOCX documents and return the redlined result.
   * @param original - Original DOCX document
   * @param modified - Modified DOCX document
   * @param options - Comparison options
   * @returns Redlined DOCX as Uint8Array
   */
  compareDocuments(
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ): Promise<Uint8Array>;

  /**
   * Compare two DOCX documents and return the result as HTML.
   * @param original - Original DOCX document
   * @param modified - Modified DOCX document
   * @param options - Comparison options
   * @returns HTML string with redlined content
   */
  compareDocumentsToHtml(
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ): Promise<string>;

  /**
   * Get revisions from a compared document.
   * @param document - A document that has tracked changes
   * @param options - Revision extraction options
   * @returns Array of revisions
   */
  getRevisions(
    document: File | Uint8Array,
    options?: GetRevisionsOptions
  ): Promise<Revision[]>;

  /**
   * Get document metadata for lazy loading pagination.
   * This is a fast operation that extracts structure without full HTML rendering.
   * @param document - DOCX file as File object or Uint8Array
   * @returns Document metadata including sections, dimensions, and content counts
   */
  getDocumentMetadata(document: File | Uint8Array): Promise<DocumentMetadata>;

  /**
   * Get version information about the library.
   * @returns Version information
   */
  getVersion(): Promise<VersionInfo>;

  /**
   * Open a {@link WorkerDocxSession} for surgical annotation editing inside
   * the worker. The document bytes are transferred to the worker (zero-copy).
   *
   * Always call {@link WorkerDocxSession.close} when you are done to release
   * the in-worker session handle.
   *
   * @param document - DOCX file as File or Uint8Array
   * @param settings - Optional session settings
   * @returns A proxied session whose methods are off-main-thread
   */
  openDocxSession(
    document: File | Uint8Array,
    settings?: DocxSessionSettings
  ): Promise<WorkerDocxSession>;

  /**
   * Terminate the worker.
   * After calling this, the instance cannot be used anymore.
   */
  terminate(): void;

  /**
   * Check if the worker is still active.
   */
  isActive(): boolean;
}

/**
 * Create a worker-based Docxodus instance.
 *
 * This function spawns a Web Worker that loads the WASM runtime independently.
 * All operations are executed in the worker, keeping the main thread responsive.
 *
 * @param options - Configuration options
 * @returns A Promise that resolves to a WorkerDocxodus instance
 *
 * @example
 * ```typescript
 * // Basic usage
 * const docxodus = await createWorkerDocxodus();
 * const html = await docxodus.convertDocxToHtml(docxFile);
 *
 * // With custom WASM path
 * const docxodus = await createWorkerDocxodus({
 *   wasmBasePath: '/assets/wasm/'
 * });
 * ```
 */
export async function createWorkerDocxodus(
  options?: WorkerDocxodusOptions
): Promise<WorkerDocxodus> {
  // Determine WASM base path
  const wasmBasePath = options?.wasmBasePath ?? deriveWasmBasePath();

  // Determine worker script path
  // The worker bundle should be in the same directory as this module
  let workerUrl: string;

  // Try to create worker from bundled script or blob
  // For now, we'll use a blob URL to inline the worker path
  const workerScriptPath = new URL("./docxodus.worker.js", import.meta.url)
    .href;

  // Create the worker
  const worker = new Worker(workerScriptPath, { type: "module" });

  // Track pending requests
  const pendingRequests = new Map<
    string,
    {
      resolve: (value: any) => void;
      reject: (error: Error) => void;
    }
  >();

  // Track if worker is active
  let isWorkerActive = true;

  // Handle worker messages
  worker.onmessage = (event: MessageEvent<WorkerResponse | { type: "ready" }>) => {
    const response = event.data;

    // Handle ready signal
    if (response.type === "ready") {
      return;
    }

    // Handle normal responses
    const pending = pendingRequests.get(response.id);
    if (pending) {
      pendingRequests.delete(response.id);

      if (response.success) {
        pending.resolve(response);
      } else {
        pending.reject(new Error(response.error || "Unknown error"));
      }
    }
  };

  // Handle worker errors
  worker.onerror = (error) => {
    // Reject all pending requests
    for (const pending of pendingRequests.values()) {
      pending.reject(new Error(`Worker error: ${error.message}`));
    }
    pendingRequests.clear();
    isWorkerActive = false;
  };

  /**
   * Send a request to the worker and wait for response.
   */
  function sendRequest<T extends WorkerResponse>(
    request: WorkerRequest,
    transfer?: Transferable[]
  ): Promise<T> {
    return new Promise((resolve, reject) => {
      if (!isWorkerActive) {
        reject(new Error("Worker has been terminated"));
        return;
      }

      pendingRequests.set(request.id, { resolve, reject });

      if (transfer && transfer.length > 0) {
        worker.postMessage(request, transfer);
      } else {
        worker.postMessage(request);
      }
    });
  }

  // Initialize the worker
  const initResponse = await sendRequest({
    id: generateId(),
    type: "init",
    wasmBasePath,
  });

  if (!initResponse.success) {
    worker.terminate();
    throw new Error(`Failed to initialize worker: ${initResponse.error}`);
  }

  // Return the WorkerDocxodus instance
  return {
    async convertDocxToHtml(
      document: File | Uint8Array,
      options?: ConversionOptions
    ): Promise<string> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerConvertResponse>(
        {
          id: generateId(),
          type: "convertDocxToHtml",
          documentBytes: bytes,
          options,
        },
        [bytes.buffer]
      );
      return response.html!;
    },

    async compareDocuments(
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ): Promise<Uint8Array> {
      const originalBytes = await toBytes(original);
      const modifiedBytes = await toBytes(modified);
      const response = await sendRequest<WorkerCompareResponse>(
        {
          id: generateId(),
          type: "compareDocuments",
          originalBytes,
          modifiedBytes,
          options,
        },
        [originalBytes.buffer, modifiedBytes.buffer]
      );
      return response.documentBytes!;
    },

    async compareDocumentsToHtml(
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ): Promise<string> {
      const originalBytes = await toBytes(original);
      const modifiedBytes = await toBytes(modified);
      const response = await sendRequest<WorkerCompareToHtmlResponse>(
        {
          id: generateId(),
          type: "compareDocumentsToHtml",
          originalBytes,
          modifiedBytes,
          options,
        },
        [originalBytes.buffer, modifiedBytes.buffer]
      );
      return response.html!;
    },

    async getRevisions(
      document: File | Uint8Array,
      options?: GetRevisionsOptions
    ): Promise<Revision[]> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerGetRevisionsResponse>(
        {
          id: generateId(),
          type: "getRevisions",
          documentBytes: bytes,
          options,
        },
        [bytes.buffer]
      );
      return response.revisions!;
    },

    async getDocumentMetadata(
      document: File | Uint8Array
    ): Promise<DocumentMetadata> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerGetDocumentMetadataResponse>(
        {
          id: generateId(),
          type: "getDocumentMetadata",
          documentBytes: bytes,
        },
        [bytes.buffer]
      );
      return response.metadata!;
    },

    async getVersion(): Promise<VersionInfo> {
      const response = await sendRequest<WorkerGetVersionResponse>({
        id: generateId(),
        type: "getVersion",
      });
      return response.version!;
    },

    async openDocxSession(
      document: File | Uint8Array,
      settings?: DocxSessionSettings
    ): Promise<WorkerDocxSession> {
      const bytes = await toBytes(document);
      const settingsJson = settings ? JSON.stringify(settings) : "";
      const openResponse = await sendRequest<WorkerSessionOpenResponse>(
        {
          id: generateId(),
          type: "sessionOpen",
          documentBytes: bytes,
          settingsJson,
        },
        [bytes.buffer]
      );

      if (!openResponse.success || openResponse.handle === undefined) {
        throw new Error(
          `Failed to open worker DocxSession: ${openResponse.error ?? "unknown error"}`
        );
      }

      const handle = openResponse.handle;

      return {
        async addAnnotation(
          anchorId: string,
          span: CharSpan | null,
          annotation: DocumentAnnotation
        ): Promise<EditResult> {
          const res = await sendRequest<WorkerSessionEditResponse>({
            id: generateId(),
            type: "sessionAddAnnotation",
            handle,
            anchorId,
            spanJson: span ? JSON.stringify(span) : "",
            annotationJson: JSON.stringify(annotation),
          });
          if (!res.success) {
            throw new Error(res.error ?? "sessionAddAnnotation failed");
          }
          return res.result!;
        },

        async removeAnnotation(annotationId: string): Promise<EditResult> {
          const res = await sendRequest<WorkerSessionEditResponse>({
            id: generateId(),
            type: "sessionRemoveAnnotation",
            handle,
            annotationId,
          });
          if (!res.success) {
            throw new Error(res.error ?? "sessionRemoveAnnotation failed");
          }
          return res.result!;
        },

        async updateAnnotation(
          annotationId: string,
          update: AnnotationUpdate
        ): Promise<EditResult> {
          const res = await sendRequest<WorkerSessionEditResponse>({
            id: generateId(),
            type: "sessionUpdateAnnotation",
            handle,
            annotationId,
            updateJson: JSON.stringify(update),
          });
          if (!res.success) {
            throw new Error(res.error ?? "sessionUpdateAnnotation failed");
          }
          return res.result!;
        },

        async moveAnnotation(
          annotationId: string,
          newAnchorId: string,
          newSpan: CharSpan | null
        ): Promise<EditResult> {
          const res = await sendRequest<WorkerSessionEditResponse>({
            id: generateId(),
            type: "sessionMoveAnnotation",
            handle,
            annotationId,
            newAnchorId,
            newSpanJson: newSpan ? JSON.stringify(newSpan) : "",
          });
          if (!res.success) {
            throw new Error(res.error ?? "sessionMoveAnnotation failed");
          }
          return res.result!;
        },

        async close(): Promise<void> {
          await sendRequest({
            id: generateId(),
            type: "sessionClose",
            handle,
          });
        },
      };
    },

    terminate(): void {
      isWorkerActive = false;
      worker.terminate();

      // Reject any pending requests
      for (const pending of pendingRequests.values()) {
        pending.reject(new Error("Worker terminated"));
      }
      pendingRequests.clear();
    },

    isActive(): boolean {
      return isWorkerActive;
    },
  };
}

/**
 * Check if Web Workers are supported in the current environment.
 */
export function isWorkerSupported(): boolean {
  return typeof Worker !== "undefined";
}

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
  WorkerGeneratePackageManifestResponse,
  WorkerVerifyDeliverableResponse,
  WorkerProjectReviewProfileResponse,
  WorkerCompareResponse,
  WorkerCompareToHtmlResponse,
  WorkerGetSemanticChangesResponse,
  WorkerGetRevisionsResponse,
  WorkerGetDocumentMetadataResponse,
  WorkerGetVersionResponse,
  WorkerPrepareResponse,
  WorkerSessionOpenResponse,
  WorkerSessionGetPackageManifestResponse,
  WorkerSessionGetSemanticChangesResponse,
  WorkerSessionVerifyDeliverableResponse,
  WorkerSessionEditResponse,
  WorkerDocxodusOptions,
  ConversionOptions,
  CompareOptions,
  DocxDiffSettings,
  GetRevisionsOptions,
  Revision,
  VersionInfo,
  DocumentMetadata,
  DocxSessionSettings,
  DocumentAnnotation,
  AnnotationUpdate,
  CharSpan,
  EditResult,
  PackageManifest,
  DeliverableVerificationResult,
  SemanticChangeSet,
  PackageManifestInspectionLimits,
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
    // Transfer only an exact-view clone. Transferring the caller's buffer would detach it,
    // and transferring a subarray's backing buffer would also send unrelated prefix/suffix bytes.
    return new Uint8Array(document);
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
  /** Generate a deterministic manifest of the session's current logical checkpoint. */
  getPackageManifest(): Promise<PackageManifest>;

  /** Run the default deliverable gate over the session's clean-save checkpoint. */
  verifyDeliverable(): Promise<DeliverableVerificationResult>;

  /** Compare the current logical checkpoint with the exact opening package. */
  getSemanticChanges(): Promise<SemanticChangeSet>;

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
  /** Generate a deterministic, non-mutating verification manifest. */
  generatePackageManifest(
    document: File | Uint8Array,
    limits?: PackageManifestInspectionLimits,
  ): Promise<PackageManifest>;

  /** Generate the exact canonical manifest JSON for strict boundary validation. */
  generatePackageManifestJson(
    document: File | Uint8Array,
    limits?: PackageManifestInspectionLimits,
  ): Promise<string>;

  /** Derive an isolated final/original package; the source bytes remain caller-owned. */
  projectReviewProfile(
    document: File | Uint8Array,
    profile: "final" | "original",
    maximumOutputBytes?: number,
  ): Promise<Uint8Array>;

  /** Run the default bounded deliverable gate, optionally against an exact baseline. */
  verifyDeliverable(
    document: File | Uint8Array,
    baseline?: File | Uint8Array
  ): Promise<DeliverableVerificationResult>;

  /** Compare two DOCX packages into the stable, versioned semantic-change schema. */
  getSemanticChanges(
    left: File | Uint8Array,
    right: File | Uint8Array,
    settings?: DocxDiffSettings
  ): Promise<SemanticChangeSet>;

  /**
   * Convert a DOCX document to HTML.
   * @param document - DOCX file as File object or Uint8Array
   * @param options - Conversion options
   * @returns HTML string
   */
  convertDocxToHtml(
    document: File | Uint8Array,
    options?: ConversionOptions,
    maximumOutputBytes?: number,
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
   * Pre-warm the comparison code path.
   *
   * The 10s runtime warmup paid by {@link createWorkerDocxodus} does not load
   * the comparison assemblies — the .NET WASM runtime defers
   * `Docxodus.*.wasm` and its `System.*.wasm` dependents until the first
   * {@link compareDocuments} call, which then costs ~3s of pure assembly-load
   * latency. Call `prepare()` after creating the worker to pay that cost ahead
   * of any user action; once it resolves, the next {@link compareDocuments}
   * (or {@link compareDocumentsToHtml}) triggers no further `.wasm` fetches.
   *
   * Semantics:
   * - **Idempotent.** Repeated calls share one in-flight warmup and resolve
   *   immediately once it has completed.
   * - **No caller IO.** No seed files to fetch, no inputs to construct — the
   *   seed documents are built inside the worker.
   * - **Concurrent-safe.** `prepare()` and `compareDocuments()` may be called
   *   in any order; a `compareDocuments()` issued while a `prepare()` is in
   *   flight does not double-load assemblies.
   *
   * @returns A Promise that resolves when the comparison path is fully hot.
   */
  prepare(): Promise<void>;

  /**
   * Open a {@link WorkerDocxSession} for surgical annotation editing inside
   * the worker. Uint8Array inputs are copied once so the caller's buffer remains attached and
   * subarray boundaries are preserved; the private copy is then transferred.
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
  if (options?.signal?.aborted) {
    throw new Error("Worker creation aborted");
  }
  // Determine WASM base path
  const wasmBasePath = options?.wasmBasePath ?? deriveWasmBasePath();

  // The worker bundle is packaged beside this module.
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
  let abortListener: (() => void) | undefined;

  const stopWorker = (message: string): void => {
    if (!isWorkerActive) return;
    isWorkerActive = false;
    worker.terminate();
    for (const pending of pendingRequests.values()) {
      pending.reject(new Error(message));
    }
    pendingRequests.clear();
    if (abortListener && options?.signal) {
      options.signal.removeEventListener("abort", abortListener);
      abortListener = undefined;
    }
  };
  if (options?.signal) {
    abortListener = () => stopWorker("Worker creation or operation aborted");
    options.signal.addEventListener("abort", abortListener, { once: true });
  }

  // Cached warmup promise. Set on the first prepare() and reused thereafter so
  // repeated/concurrent calls share a single in-flight (or completed) warmup.
  // Reset to null on failure so a later prepare() can retry.
  let preparePromise: Promise<void> | null = null;

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
    stopWorker(`Worker error: ${error.message}`);
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

      try {
        if (transfer && transfer.length > 0) {
          worker.postMessage(request, transfer);
        } else {
          worker.postMessage(request);
        }
      } catch (error) {
        pendingRequests.delete(request.id);
        reject(error instanceof Error ? error : new Error(String(error)));
      }
    });
  }

  // Initialize the worker
  try {
    await sendRequest({
      id: generateId(),
      type: "init",
      wasmBasePath,
    });
  } catch (error) {
    stopWorker("Worker initialization failed");
    throw error;
  }

  // Return the WorkerDocxodus instance
  return {
    async generatePackageManifest(
      document: File | Uint8Array,
      limits?: PackageManifestInspectionLimits,
    ): Promise<PackageManifest> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerGeneratePackageManifestResponse>(
        {
          id: generateId(),
          type: "generatePackageManifest",
          documentBytes: bytes,
          limits,
        },
        [bytes.buffer]
      );
      return response.manifest!;
    },

    async verifyDeliverable(
      document: File | Uint8Array,
      baseline?: File | Uint8Array
    ): Promise<DeliverableVerificationResult> {
      const bytes = await toBytes(document);
      const baselineBytes = baseline === undefined ? undefined : await toBytes(baseline);
      const transfer: Transferable[] = [bytes.buffer];
      if (baselineBytes !== undefined) transfer.push(baselineBytes.buffer);
      const response = await sendRequest<WorkerVerifyDeliverableResponse>(
        {
          id: generateId(),
          type: "verifyDeliverable",
          documentBytes: bytes,
          baselineBytes,
        },
        transfer
      );
      if (!response.success || !response.verification) {
        throw new Error(response.error ?? "verifyDeliverable failed");
      }
      return response.verification;
    },

    async getSemanticChanges(
      left: File | Uint8Array,
      right: File | Uint8Array,
      settings?: DocxDiffSettings
    ): Promise<SemanticChangeSet> {
      const leftBytes = await toBytes(left);
      const rightBytes = await toBytes(right);
      const response = await sendRequest<WorkerGetSemanticChangesResponse>(
        {
          id: generateId(),
          type: "getSemanticChanges",
          leftBytes,
          rightBytes,
          settings,
        },
        [leftBytes.buffer, rightBytes.buffer]
      );
      return response.semanticChanges!;
    },

    async generatePackageManifestJson(
      document: File | Uint8Array,
      limits?: PackageManifestInspectionLimits,
    ): Promise<string> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerGeneratePackageManifestResponse>(
        {
          id: generateId(),
          type: "generatePackageManifest",
          documentBytes: bytes,
          limits,
        },
        [bytes.buffer],
      );
      if (response.manifestJson === undefined) {
        throw new Error("Package manifest worker response omitted canonical JSON");
      }
      return response.manifestJson;
    },

    async projectReviewProfile(
      document: File | Uint8Array,
      profile: "final" | "original",
      maximumOutputBytes?: number,
    ): Promise<Uint8Array> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerProjectReviewProfileResponse>(
        {
          id: generateId(),
          type: "projectReviewProfile",
          documentBytes: bytes,
          profile,
          maximumOutputBytes,
        },
        [bytes.buffer],
      );
      if (!response.documentBytes || response.documentBytes.byteLength === 0) {
        throw new Error(`Failed to derive the ${profile} review profile`);
      }
      return response.documentBytes;
    },

    async convertDocxToHtml(
      document: File | Uint8Array,
      options?: ConversionOptions,
      maximumOutputBytes?: number,
    ): Promise<string> {
      const bytes = await toBytes(document);
      const response = await sendRequest<WorkerConvertResponse>(
        {
          id: generateId(),
          type: "convertDocxToHtml",
          documentBytes: bytes,
          options,
          maximumOutputBytes,
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

    prepare(): Promise<void> {
      // Idempotent: hand back the existing warmup if one is in flight or done.
      if (preparePromise) {
        return preparePromise;
      }
      preparePromise = sendRequest<WorkerPrepareResponse>({
        id: generateId(),
        type: "prepare",
      }).then(() => undefined);
      // On failure, clear the cache so a subsequent prepare() can retry.
      preparePromise.catch(() => {
        preparePromise = null;
      });
      return preparePromise;
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
        async getPackageManifest(): Promise<PackageManifest> {
          const res = await sendRequest<WorkerSessionGetPackageManifestResponse>({
            id: generateId(),
            type: "sessionGetPackageManifest",
            handle,
          });
          if (!res.success || !res.manifest) {
            throw new Error(res.error ?? "sessionGetPackageManifest failed");
          }
          return res.manifest;
        },

        async getSemanticChanges(): Promise<SemanticChangeSet> {
          const res = await sendRequest<WorkerSessionGetSemanticChangesResponse>({
            id: generateId(),
            type: "sessionGetSemanticChanges",
            handle,
          });
          if (!res.success || !res.semanticChanges) {
            throw new Error(res.error ?? "sessionGetSemanticChanges failed");
          }
          return res.semanticChanges;
        },

        async verifyDeliverable(): Promise<DeliverableVerificationResult> {
          const res = await sendRequest<WorkerSessionVerifyDeliverableResponse>({
            id: generateId(),
            type: "sessionVerifyDeliverable",
            handle,
          });
          if (!res.success || !res.verification) {
            throw new Error(res.error ?? "sessionVerifyDeliverable failed");
          }
          return res.verification;
        },

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
      stopWorker("Worker terminated");
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

import React, { useState, useEffect, useCallback, useRef, createElement } from "react";
import type { CSSProperties, ReactElement } from "react";
import {
  initialize,
  convertDocxToHtml,
  compareDocuments,
  compareDocumentsToHtml,
  getRevisions,
  isInitialized,
} from "./index.js";
import type {
  ConversionOptions,
  CompareOptions,
  Revision,
} from "./types.js";
import {
  PaginationEngine,
  type PaginationOptions,
  type PaginationResult,
} from "./pagination.js";

export type { ConversionOptions, CompareOptions, Revision, PaginationOptions, PaginationResult };

export interface UseDocxodusResult {
  /** Whether the WASM runtime is loaded and ready */
  isReady: boolean;
  /** Whether the runtime is currently loading */
  isLoading: boolean;
  /** Error that occurred during initialization, if any */
  error: Error | null;
  /** Convert DOCX to HTML */
  convertToHtml: (
    document: File | Uint8Array,
    options?: ConversionOptions
  ) => Promise<string>;
  /** Compare two documents and return redlined DOCX */
  compare: (
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ) => Promise<Uint8Array>;
  /** Compare two documents and return HTML */
  compareToHtml: (
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ) => Promise<string>;
  /** Get revisions from a compared document */
  getRevisions: (document: File | Uint8Array) => Promise<Revision[]>;
}

/**
 * React hook for using Docxodus WASM functionality.
 * Automatically initializes the WASM runtime on mount.
 *
 * WASM files are auto-detected from the module's location (works with CDN, npm, or local hosting).
 * Pass a custom path only if you need to host files at a different location.
 *
 * @param wasmBasePath - Optional custom path to WASM files. Leave empty for auto-detection.
 * @returns Object with ready state and document functions
 *
 * @example
 * ```tsx
 * function App() {
 *   // Auto-detects WASM location - no configuration needed!
 *   const { isReady, isLoading, error, convertToHtml } = useDocxodus();
 *
 *   const handleFile = async (file: File) => {
 *     if (!isReady) return;
 *     const html = await convertToHtml(file);
 *     setHtml(html);
 *   };
 *
 *   if (isLoading) return <div>Loading WASM...</div>;
 *   if (error) return <div>Error: {error.message}</div>;
 *
 *   return <input type="file" onChange={e => handleFile(e.target.files[0])} />;
 * }
 * ```
 */
export function useDocxodus(wasmBasePath?: string): UseDocxodusResult {
  const [isReady, setIsReady] = useState(isInitialized());
  const [isLoading, setIsLoading] = useState(!isInitialized());
  const [error, setError] = useState<Error | null>(null);

  useEffect(() => {
    if (isInitialized()) {
      setIsReady(true);
      setIsLoading(false);
      return;
    }

    let cancelled = false;

    const init = async () => {
      try {
        setIsLoading(true);
        await initialize(wasmBasePath);
        if (!cancelled) {
          setIsReady(true);
          setError(null);
        }
      } catch (err) {
        if (!cancelled) {
          setError(err instanceof Error ? err : new Error(String(err)));
        }
      } finally {
        if (!cancelled) {
          setIsLoading(false);
        }
      }
    };

    init();

    return () => {
      cancelled = true;
    };
  }, [wasmBasePath]);

  const convertToHtml = useCallback(
    async (document: File | Uint8Array, options?: ConversionOptions) => {
      if (!isReady) {
        throw new Error("Docxodus not initialized");
      }
      return convertDocxToHtml(document, options);
    },
    [isReady]
  );

  const compare = useCallback(
    async (
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ) => {
      if (!isReady) {
        throw new Error("Docxodus not initialized");
      }
      return compareDocuments(original, modified, options);
    },
    [isReady]
  );

  const compareToHtml = useCallback(
    async (
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ) => {
      if (!isReady) {
        throw new Error("Docxodus not initialized");
      }
      return compareDocumentsToHtml(original, modified, options);
    },
    [isReady]
  );

  const getRevisionsCallback = useCallback(
    async (document: File | Uint8Array) => {
      if (!isReady) {
        throw new Error("Docxodus not initialized");
      }
      return getRevisions(document);
    },
    [isReady]
  );

  return {
    isReady,
    isLoading,
    error,
    convertToHtml,
    compare,
    compareToHtml,
    getRevisions: getRevisionsCallback,
  };
}

export interface UseConversionResult {
  /** The converted HTML output */
  html: string | null;
  /** Whether a conversion is in progress */
  isConverting: boolean;
  /** Error from the last conversion attempt */
  error: Error | null;
  /** Convert a DOCX file to HTML */
  convert: (document: File | Uint8Array, options?: ConversionOptions) => Promise<void>;
  /** Clear the current result */
  clear: () => void;
}

/**
 * React hook for DOCX to HTML conversion with state management.
 * WASM files are auto-detected from the module's location.
 *
 * @param wasmBasePath - Optional custom path to WASM files. Leave empty for auto-detection.
 *
 * @example
 * ```tsx
 * function Converter() {
 *   // Auto-detects WASM location - no configuration needed!
 *   const { html, isConverting, error, convert } = useConversion();
 *
 *   return (
 *     <div>
 *       <input
 *         type="file"
 *         accept=".docx"
 *         onChange={e => e.target.files?.[0] && convert(e.target.files[0])}
 *         disabled={isConverting}
 *       />
 *       {isConverting && <p>Converting...</p>}
 *       {error && <p>Error: {error.message}</p>}
 *       {html && <div dangerouslySetInnerHTML={{ __html: html }} />}
 *     </div>
 *   );
 * }
 * ```
 */
export function useConversion(wasmBasePath?: string): UseConversionResult {
  const { isReady, convertToHtml } = useDocxodus(wasmBasePath);
  const [html, setHtml] = useState<string | null>(null);
  const [isConverting, setIsConverting] = useState(false);
  const [error, setError] = useState<Error | null>(null);

  const convert = useCallback(
    async (document: File | Uint8Array, options?: ConversionOptions) => {
      if (!isReady) {
        setError(new Error("Docxodus not initialized"));
        return;
      }

      setIsConverting(true);
      setError(null);

      try {
        const result = await convertToHtml(document, options);
        setHtml(result);
      } catch (err) {
        setError(err instanceof Error ? err : new Error(String(err)));
      } finally {
        setIsConverting(false);
      }
    },
    [isReady, convertToHtml]
  );

  const clear = useCallback(() => {
    setHtml(null);
    setError(null);
  }, []);

  return { html, isConverting, error, convert, clear };
}

export interface UseComparisonResult {
  /** The comparison result as a Uint8Array (redlined DOCX) */
  result: Uint8Array | null;
  /** The comparison result as HTML (if compareToHtml was used) */
  html: string | null;
  /** Revisions extracted from the comparison */
  revisions: Revision[] | null;
  /** Whether a comparison is in progress */
  isComparing: boolean;
  /** Error from the last comparison attempt */
  error: Error | null;
  /** Compare two documents and get redlined DOCX */
  compare: (
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ) => Promise<void>;
  /** Compare two documents and get HTML */
  compareToHtml: (
    original: File | Uint8Array,
    modified: File | Uint8Array,
    options?: CompareOptions
  ) => Promise<void>;
  /** Clear all results */
  clear: () => void;
  /** Download the result as a DOCX file */
  downloadResult: (filename?: string) => void;
}

/**
 * React hook for document comparison with state management.
 * WASM files are auto-detected from the module's location.
 *
 * @param wasmBasePath - Optional custom path to WASM files. Leave empty for auto-detection.
 *
 * @example
 * ```tsx
 * function Comparer() {
 *   // Auto-detects WASM location - no configuration needed!
 *   const { html, isComparing, error, compareToHtml, downloadResult } = useComparison();
 *   const [original, setOriginal] = useState<File | null>(null);
 *   const [modified, setModified] = useState<File | null>(null);
 *
 *   const handleCompare = () => {
 *     if (original && modified) {
 *       compareToHtml(original, modified, { authorName: 'User' });
 *     }
 *   };
 *
 *   return (
 *     <div>
 *       <input type="file" onChange={e => setOriginal(e.target.files?.[0] ?? null)} />
 *       <input type="file" onChange={e => setModified(e.target.files?.[0] ?? null)} />
 *       <button onClick={handleCompare} disabled={isComparing}>Compare</button>
 *       {html && <div dangerouslySetInnerHTML={{ __html: html }} />}
 *     </div>
 *   );
 * }
 * ```
 */
export function useComparison(wasmBasePath?: string): UseComparisonResult {
  const docxodus = useDocxodus(wasmBasePath);
  const [result, setResult] = useState<Uint8Array | null>(null);
  const [html, setHtml] = useState<string | null>(null);
  const [revisions, setRevisions] = useState<Revision[] | null>(null);
  const [isComparing, setIsComparing] = useState(false);
  const [error, setError] = useState<Error | null>(null);

  const compare = useCallback(
    async (
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ) => {
      if (!docxodus.isReady) {
        setError(new Error("Docxodus not initialized"));
        return;
      }

      setIsComparing(true);
      setError(null);

      try {
        const docResult = await docxodus.compare(original, modified, options);
        setResult(docResult);
        setHtml(null);

        // Also get revisions
        const revs = await docxodus.getRevisions(docResult);
        setRevisions(revs);
      } catch (err) {
        setError(err instanceof Error ? err : new Error(String(err)));
      } finally {
        setIsComparing(false);
      }
    },
    [docxodus]
  );

  const compareToHtmlCallback = useCallback(
    async (
      original: File | Uint8Array,
      modified: File | Uint8Array,
      options?: CompareOptions
    ) => {
      if (!docxodus.isReady) {
        setError(new Error("Docxodus not initialized"));
        return;
      }

      setIsComparing(true);
      setError(null);

      try {
        const htmlResult = await docxodus.compareToHtml(original, modified, options);
        setHtml(htmlResult);
        setResult(null);
        setRevisions(null);
      } catch (err) {
        setError(err instanceof Error ? err : new Error(String(err)));
      } finally {
        setIsComparing(false);
      }
    },
    [docxodus]
  );

  const clear = useCallback(() => {
    setResult(null);
    setHtml(null);
    setRevisions(null);
    setError(null);
  }, []);

  const downloadResult = useCallback(
    (filename = "comparison-result.docx") => {
      if (!result) return;

      const blob = new Blob([result as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    },
    [result]
  );

  return {
    result,
    html,
    revisions,
    isComparing,
    error,
    compare,
    compareToHtml: compareToHtmlCallback,
    clear,
    downloadResult,
  };
}

/**
 * Props for the PaginatedDocument component.
 */
export interface PaginatedDocumentProps {
  /** HTML string with pagination metadata (from convertDocxToHtml with PaginationMode.Paginated) */
  html: string;
  /** Scale factor for page rendering (1.0 = 100%). Default: 1 */
  scale?: number;
  /** Whether to show page numbers. Default: true */
  showPageNumbers?: boolean;
  /** Gap between pages in pixels. Default: 20 */
  pageGap?: number;
  /** Background color for the viewer. Default: "#525659" */
  backgroundColor?: string;
  /** CSS class prefix used in the HTML. Default: "page-" */
  cssPrefix?: string;
  /** Callback when pagination completes */
  onPaginationComplete?: (result: PaginationResult) => void;
  /** Callback when a page becomes visible (for tracking current page) */
  onPageVisible?: (pageNumber: number) => void;
  /** Additional CSS class for the container */
  className?: string;
  /** Additional inline styles for the container */
  style?: CSSProperties;
}

/**
 * Result of the usePagination hook.
 */
export interface UsePaginationResult {
  /** Pagination result after processing */
  result: PaginationResult | null;
  /** Whether pagination is in progress */
  isPaginating: boolean;
  /** Error that occurred during pagination */
  error: Error | null;
  /** Manually trigger pagination */
  paginate: () => void;
}

/**
 * React hook for pagination state management.
 *
 * @param html - HTML string with pagination metadata
 * @param containerRef - Ref to the container element
 * @param options - Pagination options
 * @returns Pagination state and controls
 *
 * @example
 * ```tsx
 * function Viewer({ html }: { html: string }) {
 *   const containerRef = useRef<HTMLDivElement>(null);
 *   const { result, isPaginating, error, paginate } = usePagination(html, containerRef);
 *
 *   return (
 *     <div ref={containerRef} style={{ minHeight: '100vh' }}>
 *       {isPaginating && <div>Paginating...</div>}
 *       {result && <div>Total pages: {result.totalPages}</div>}
 *     </div>
 *   );
 * }
 * ```
 */
export function usePagination(
  html: string,
  containerRef: React.RefObject<HTMLElement | null>,
  options: PaginationOptions = {}
): UsePaginationResult {
  const [result, setResult] = useState<PaginationResult | null>(null);
  const [isPaginating, setIsPaginating] = useState(false);
  const [error, setError] = useState<Error | null>(null);

  const paginate = useCallback(() => {
    if (!containerRef.current || !html) {
      return;
    }

    setIsPaginating(true);
    setError(null);

    try {
      const container = containerRef.current;

      // Insert HTML into container
      container.innerHTML = html;

      // Find staging and page container
      const cssPrefix = options.cssPrefix ?? "page-";
      const staging =
        container.querySelector<HTMLElement>("#pagination-staging") ||
        container.querySelector<HTMLElement>(`.${cssPrefix}staging`);
      const pageContainer =
        container.querySelector<HTMLElement>("#pagination-container") ||
        container.querySelector<HTMLElement>(`.${cssPrefix}container`);

      if (!staging || !pageContainer) {
        throw new Error(
          "Pagination elements not found. Ensure HTML was generated with PaginationMode.Paginated"
        );
      }

      const engine = new PaginationEngine(staging, pageContainer, options);
      const paginationResult = engine.paginate();
      setResult(paginationResult);
    } catch (err) {
      setError(err instanceof Error ? err : new Error(String(err)));
    } finally {
      setIsPaginating(false);
    }
  }, [html, containerRef, options]);

  // Auto-paginate when HTML changes
  useEffect(() => {
    if (html && containerRef.current) {
      paginate();
    }
  }, [html, paginate]);

  return { result, isPaginating, error, paginate };
}

/**
 * React component for displaying a paginated document view (PDF.js style).
 *
 * @example
 * ```tsx
 * import { useState, useEffect } from 'react';
 * import { useDocxodus, PaginatedDocument, PaginationMode } from 'docxodus/react';
 *
 * function DocumentViewer() {
 *   const { isReady, convertToHtml } = useDocxodus();
 *   const [html, setHtml] = useState<string | null>(null);
 *
 *   const handleFile = async (file: File) => {
 *     const result = await convertToHtml(file, {
 *       paginationMode: PaginationMode.Paginated,
 *       paginationScale: 0.8
 *     });
 *     setHtml(result);
 *   };
 *
 *   return (
 *     <div>
 *       <input type="file" accept=".docx" onChange={e => e.target.files?.[0] && handleFile(e.target.files[0])} />
 *       {html && (
 *         <PaginatedDocument
 *           html={html}
 *           scale={0.8}
 *           onPaginationComplete={result => console.log(`${result.totalPages} pages`)}
 *         />
 *       )}
 *     </div>
 *   );
 * }
 * ```
 */
export function PaginatedDocument({
  html,
  scale = 1,
  showPageNumbers = true,
  pageGap = 20,
  backgroundColor = "#525659",
  cssPrefix = "page-",
  onPaginationComplete,
  onPageVisible,
  className,
  style,
}: PaginatedDocumentProps): ReactElement {
  const containerRef = useRef<HTMLDivElement>(null);
  const { result, isPaginating, error } = usePagination(html, containerRef, {
    scale,
    showPageNumbers,
    pageGap,
    cssPrefix,
  });

  // Notify when pagination completes
  useEffect(() => {
    if (result && onPaginationComplete) {
      onPaginationComplete(result);
    }
  }, [result, onPaginationComplete]);

  // Set up intersection observer for page visibility tracking
  useEffect(() => {
    if (!result || !onPageVisible || !containerRef.current) {
      return;
    }

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            const pageNum = parseInt(
              (entry.target as HTMLElement).dataset.pageNumber || "1",
              10
            );
            onPageVisible(pageNum);
          }
        });
      },
      { threshold: 0.5 }
    );

    const pages = containerRef.current.querySelectorAll(`.${cssPrefix}box`);
    pages.forEach((page) => observer.observe(page));

    return () => observer.disconnect();
  }, [result, cssPrefix, onPageVisible]);

  // Use createElement to avoid TSX dependency
  const containerStyle: CSSProperties = {
    backgroundColor,
    minHeight: "100vh",
    overflow: "auto",
    ...style,
  };

  if (error) {
    return createElement("div", {
      style: { color: "red", padding: "20px", backgroundColor },
    }, `Pagination error: ${error.message}`);
  }

  return createElement("div", {
    ref: containerRef,
    className,
    style: containerStyle,
  }, isPaginating ? "Loading..." : null);
}

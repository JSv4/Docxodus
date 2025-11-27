import { useState, useEffect, useCallback } from "react";
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

export type { ConversionOptions, CompareOptions, Revision };

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
 * @param wasmBasePath - Optional base path to WASM files
 * @returns Object with ready state and document functions
 *
 * @example
 * ```tsx
 * function App() {
 *   const { isReady, isLoading, error, convertToHtml } = useDocxodus('/wasm/');
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
 *
 * @param wasmBasePath - Optional base path to WASM files
 *
 * @example
 * ```tsx
 * function Converter() {
 *   const { html, isConverting, error, convert } = useConversion('/wasm/');
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
 *
 * @param wasmBasePath - Optional base path to WASM files
 *
 * @example
 * ```tsx
 * function Comparer() {
 *   const { html, isComparing, error, compareToHtml, downloadResult } = useComparison('/wasm/');
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

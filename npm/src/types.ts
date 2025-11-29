/**
 * Comment render mode
 */
export enum CommentRenderMode {
  /** Render comments at the end of the document with bidirectional links (like footnotes) */
  EndnoteStyle = 0,
  /** Render comments as inline tooltips with data attributes */
  Inline = 1,
  /** Render comments in a margin column (CSS-positioned) */
  Margin = 2,
}

/**
 * Options for DOCX to HTML conversion
 */
export interface ConversionOptions {
  /** Title for the HTML document (default: "Document") */
  pageTitle?: string;
  /** CSS class prefix for generated styles (default: "docx-") */
  cssPrefix?: string;
  /** Whether to generate CSS classes (default: true) */
  fabricateClasses?: boolean;
  /** Additional CSS to include in the output */
  additionalCss?: string;
  /** Whether to render document comments (default: false) */
  renderComments?: boolean;
  /** How to render comments (default: EndnoteStyle) */
  commentRenderMode?: CommentRenderMode;
  /** CSS class prefix for comment elements (default: "comment-") */
  commentCssClassPrefix?: string;
  /** Whether to include author/date metadata in comment HTML (default: true) */
  includeCommentMetadata?: boolean;
}

/**
 * Options for document comparison
 */
export interface CompareOptions {
  /** Author name for tracked changes (default: "Docxodus") */
  authorName?: string;
  /** Detail threshold 0.0-1.0 (default: 0.15, lower = more detailed) */
  detailThreshold?: number;
  /** Whether comparison is case-insensitive (default: false) */
  caseInsensitive?: boolean;
  /**
   * Whether to render tracked changes visually in HTML output (default: true)
   * If true: insertions shown with <ins>, deletions with <del>, styled with colors
   * If false: changes are accepted, output shows final "clean" document
   */
  renderTrackedChanges?: boolean;
}

/**
 * Information about a document revision
 */
export interface Revision {
  /** Author who made the revision */
  author: string;
  /** ISO date string of the revision */
  date: string;
  /** Type of revision: "Insertion", "Deletion", etc. */
  revisionType: string;
  /** Text content of the revision */
  text: string;
}

/**
 * Version information for the library
 */
export interface VersionInfo {
  library: string;
  dotnetVersion: string;
  platform: string;
}

/**
 * Error response from WASM operations
 */
export interface ErrorResponse {
  error: string;
  type?: string;
  stackTrace?: string;
}

/**
 * Result of a comparison operation
 */
export interface CompareResult {
  /** The redlined document as a Uint8Array */
  document: Uint8Array;
  /** List of revisions found */
  revisions: Revision[];
}

/**
 * Internal WASM exports structure
 */
export interface DocxodusWasmExports {
  DocumentConverter: {
    ConvertDocxToHtml: (bytes: Uint8Array) => string;
    ConvertDocxToHtmlWithOptions: (
      bytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string
    ) => string;
    ConvertDocxToHtmlAdvanced: (
      bytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string,
      renderComments: boolean,
      commentRenderMode: number,
      commentCssClassPrefix: string,
      includeCommentMetadata: boolean
    ) => string;
    GetVersion: () => string;
  };
  DocumentComparer: {
    CompareDocuments: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string
    ) => Uint8Array;
    CompareDocumentsToHtml: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string
    ) => string;
    CompareDocumentsToHtmlWithOptions: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      renderTrackedChanges: boolean
    ) => string;
    CompareDocumentsWithOptions: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      detailThreshold: number,
      caseInsensitive: boolean
    ) => Uint8Array;
    GetRevisionsJson: (comparedDocBytes: Uint8Array) => string;
  };
}

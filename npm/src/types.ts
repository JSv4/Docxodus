/**
 * Revision type enum matching the .NET WmlComparerRevisionType
 */
export enum RevisionType {
  /** Text or content that was added/inserted */
  Inserted = "Inserted",
  /** Text or content that was removed/deleted */
  Deleted = "Deleted",
  /** Text or content that was relocated within the document */
  Moved = "Moved",
  /** Text content unchanged but formatting (bold, italic, etc.) changed */
  FormatChanged = "FormatChanged",
}

/**
 * Comment render mode
 * Use -1 (Disabled) to not render comments, or a positive value to enable with that mode
 */
export enum CommentRenderMode {
  /** Do not render comments (default) */
  Disabled = -1,
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
  /** Comment rendering mode: Disabled (-1), EndnoteStyle (0), Inline (1), or Margin (2). Default: Disabled */
  commentRenderMode?: CommentRenderMode;
  /** CSS class prefix for comment elements (default: "comment-") */
  commentCssClassPrefix?: string;
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
 * Information about a document revision extracted from a compared document.
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(comparedDoc);
 * for (const rev of revisions) {
 *   if (rev.revisionType === RevisionType.Inserted) {
 *     console.log(`${rev.author} added: "${rev.text}"`);
 *   } else if (rev.revisionType === RevisionType.Deleted) {
 *     console.log(`${rev.author} removed: "${rev.text}"`);
 *   }
 * }
 * ```
 */
export interface Revision {
  /**
   * Author who made the revision.
   * This comes from the Word document's tracked changes author attribute.
   * May be empty string if the document doesn't specify an author.
   */
  author: string;
  /**
   * ISO 8601 date string when the revision was made.
   * Format: "YYYY-MM-DDTHH:mm:ssZ" (e.g., "2024-01-15T10:30:00Z")
   * May be empty string if the document doesn't specify a date.
   */
  date: string;
  /**
   * Type of revision - "Inserted", "Deleted", or "Moved".
   * Use the RevisionType enum for type-safe comparisons.
   */
  revisionType: RevisionType | string;
  /**
   * Text content of the revision.
   * For paragraph breaks, this will be a newline character.
   * May be empty string for non-text elements (e.g., images, math equations).
   */
  text: string;
  /**
   * For Moved revisions, this ID links the source and destination.
   * Both the "from" and "to" revisions share the same moveGroupId.
   * Undefined for non-move revisions.
   */
  moveGroupId?: number;
  /**
   * For Moved revisions: true = source (content moved FROM here),
   * false = destination (content moved TO here).
   * Undefined for non-move revisions.
   */
  isMoveSource?: boolean;
  /**
   * For FormatChanged revisions: details about what formatting changed.
   * Undefined for non-format-change revisions.
   */
  formatChange?: FormatChangeDetails;
}

/**
 * Details about formatting changes for FormatChanged revisions.
 */
export interface FormatChangeDetails {
  /**
   * Dictionary of old property names and values.
   * Keys are friendly property names like "bold", "italic", "fontSize".
   */
  oldProperties?: Record<string, string>;
  /**
   * Dictionary of new property names and values.
   */
  newProperties?: Record<string, string>;
  /**
   * List of property names that changed (e.g., "bold", "italic", "fontSize").
   */
  changedPropertyNames?: string[];
}

/**
 * Type guard to check if a revision is an insertion.
 * @param revision - The revision to check
 * @returns true if the revision is an insertion
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const insertions = revisions.filter(isInsertion);
 * ```
 */
export function isInsertion(revision: Revision): boolean {
  return revision.revisionType === RevisionType.Inserted;
}

/**
 * Type guard to check if a revision is a deletion.
 * @param revision - The revision to check
 * @returns true if the revision is a deletion
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const deletions = revisions.filter(isDeletion);
 * ```
 */
export function isDeletion(revision: Revision): boolean {
  return revision.revisionType === RevisionType.Deleted;
}

/**
 * Type guard to check if a revision is a move operation.
 * @param revision - The revision to check
 * @returns true if the revision is part of a move
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const moves = revisions.filter(isMove);
 * ```
 */
export function isMove(revision: Revision): boolean {
  return revision.revisionType === RevisionType.Moved;
}

/**
 * Type guard to check if a revision is a format change.
 * @param revision - The revision to check
 * @returns true if the revision is a format change
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const formatChanges = revisions.filter(isFormatChange);
 * for (const rev of formatChanges) {
 *   console.log(`Format changed: ${rev.formatChange?.changedPropertyNames?.join(", ")}`);
 * }
 * ```
 */
export function isFormatChange(revision: Revision): boolean {
  return revision.revisionType === RevisionType.FormatChanged;
}

/**
 * Type guard to check if a revision is a move source (content moved FROM here).
 * @param revision - The revision to check
 * @returns true if the revision is the source of a move
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const moveSources = revisions.filter(isMoveSource);
 * ```
 */
export function isMoveSource(revision: Revision): boolean {
  return isMove(revision) && revision.isMoveSource === true;
}

/**
 * Type guard to check if a revision is a move destination (content moved TO here).
 * @param revision - The revision to check
 * @returns true if the revision is the destination of a move
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * const moveDestinations = revisions.filter(isMoveDestination);
 * ```
 */
export function isMoveDestination(revision: Revision): boolean {
  return isMove(revision) && revision.isMoveSource === false;
}

/**
 * Find the matching pair for a move revision.
 * @param revision - A move revision
 * @param allRevisions - All revisions from the document
 * @returns The matching move revision, or undefined if not found
 *
 * @example
 * ```typescript
 * const revisions = await getRevisions(doc);
 * for (const rev of revisions.filter(isMoveSource)) {
 *   const destination = findMovePair(rev, revisions);
 *   console.log(`"${rev.text}" moved to become "${destination?.text}"`);
 * }
 * ```
 */
export function findMovePair(
  revision: Revision,
  allRevisions: Revision[]
): Revision | undefined {
  if (!isMove(revision) || revision.moveGroupId === undefined) {
    return undefined;
  }
  return allRevisions.find(
    (r) =>
      r.moveGroupId === revision.moveGroupId &&
      r.isMoveSource !== revision.isMoveSource
  );
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
      additionalCss: string,
      commentRenderMode: number,
      commentCssClassPrefix: string
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
    GetRevisionsJsonWithOptions: (
      comparedDocBytes: Uint8Array,
      detectMoves: boolean,
      moveSimilarityThreshold: number,
      moveMinimumWordCount: number,
      caseInsensitive: boolean
    ) => string;
  };
}

/**
 * Options for revision extraction with move detection configuration.
 */
export interface GetRevisionsOptions {
  /**
   * Whether to detect and mark moved content.
   * When enabled, deletions and insertions with similar text are linked as move pairs.
   * @default true
   */
  detectMoves?: boolean;

  /**
   * Jaccard similarity threshold for move detection (0.0 to 1.0).
   * Higher values require more exact word overlap between deletion and insertion.
   * @default 0.8
   */
  moveSimilarityThreshold?: number;

  /**
   * Minimum word count for content to be considered for move detection.
   * Short phrases below this threshold are excluded to avoid false positives.
   * @default 3
   */
  moveMinimumWordCount?: number;

  /**
   * Whether similarity matching ignores case differences.
   * @default false
   */
  caseInsensitive?: boolean;
}

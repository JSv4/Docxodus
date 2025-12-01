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
 * Pagination mode for HTML output
 */
export enum PaginationMode {
  /** No pagination - content flows continuously (default) */
  None = 0,
  /**
   * Paginated view - outputs page containers with document dimensions
   * and content with data attributes for client-side pagination.
   * Creates a PDF.js-style page preview experience.
   */
  Paginated = 1,
}

/**
 * Annotation label display mode
 */
export enum AnnotationLabelMode {
  /** Floating label positioned above the highlight */
  Above = 0,
  /** Label displayed inline at start of highlight */
  Inline = 1,
  /** Label shown only on hover (tooltip) */
  Tooltip = 2,
  /** No labels displayed, only highlights */
  None = 3,
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
  /** Pagination mode: None (0) or Paginated (1). Default: None */
  paginationMode?: PaginationMode;
  /** Scale factor for page rendering in paginated mode (1.0 = 100%). Default: 1.0 */
  paginationScale?: number;
  /** CSS class prefix for pagination elements. Default: "page-" */
  paginationCssClassPrefix?: string;
  /** Whether to render custom annotations (default: false) */
  renderAnnotations?: boolean;
  /** How to display annotation labels (default: Above) */
  annotationLabelMode?: AnnotationLabelMode;
  /** CSS class prefix for annotation elements (default: "annot-") */
  annotationCssClassPrefix?: string;
  /** Whether to render footnotes and endnotes sections at the end of the document (default: false) */
  renderFootnotesAndEndnotes?: boolean;
  /** Whether to render document headers and footers (default: false) */
  renderHeadersAndFooters?: boolean;
  /** Whether to render tracked changes visually (insertions/deletions) (default: false) */
  renderTrackedChanges?: boolean;
  /** Whether to show deleted content with strikethrough (only when renderTrackedChanges=true, default: true) */
  showDeletedContent?: boolean;
  /** Whether to distinguish move operations from regular insert/delete (only when renderTrackedChanges=true, default: true) */
  renderMoveOperations?: boolean;
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
    ConvertDocxToHtmlWithPagination: (
      bytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string,
      commentRenderMode: number,
      commentCssClassPrefix: string,
      paginationMode: number,
      paginationScale: number,
      paginationCssClassPrefix: string
    ) => string;
    ConvertDocxToHtmlFull: (
      bytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string,
      commentRenderMode: number,
      commentCssClassPrefix: string,
      paginationMode: number,
      paginationScale: number,
      paginationCssClassPrefix: string,
      renderAnnotations: boolean,
      annotationLabelMode: number,
      annotationCssClassPrefix: string
    ) => string;
    ConvertDocxToHtmlComplete: (
      bytes: Uint8Array,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string,
      commentRenderMode: number,
      commentCssClassPrefix: string,
      paginationMode: number,
      paginationScale: number,
      paginationCssClassPrefix: string,
      renderAnnotations: boolean,
      annotationLabelMode: number,
      annotationCssClassPrefix: string,
      renderFootnotesAndEndnotes: boolean,
      renderHeadersAndFooters: boolean,
      renderTrackedChanges: boolean,
      showDeletedContent: boolean,
      renderMoveOperations: boolean
    ) => string;
    GetAnnotations: (bytes: Uint8Array) => string;
    AddAnnotation: (bytes: Uint8Array, requestJson: string) => string;
    AddAnnotationWithTarget: (bytes: Uint8Array, requestJson: string) => string;
    RemoveAnnotation: (bytes: Uint8Array, annotationId: string) => string;
    HasAnnotations: (bytes: Uint8Array) => string;
    GetDocumentStructure: (bytes: Uint8Array) => string;
    GetDocumentMetadata: (bytes: Uint8Array) => string;
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

/**
 * A custom annotation on a document range.
 */
export interface Annotation {
  /** Unique annotation ID */
  id: string;
  /** Label category/type identifier (e.g., "CLAUSE_TYPE_A", "DATE_REF") */
  labelId: string;
  /** Human-readable label text */
  label: string;
  /** Highlight color in hex format (e.g., "#FFEB3B") */
  color: string;
  /** Author who created the annotation */
  author?: string;
  /** Creation timestamp (ISO 8601) */
  created?: string;
  /** Internal bookmark name */
  bookmarkName?: string;
  /** Start page number (if computed) */
  startPage?: number;
  /** End page number (if computed) */
  endPage?: number;
  /** The annotated text content */
  annotatedText?: string;
  /** Custom metadata key-value pairs */
  metadata?: Record<string, string>;
}

/**
 * Request to add an annotation to a document.
 */
export interface AddAnnotationRequest {
  /** Unique annotation ID */
  id: string;
  /** Label category/type identifier */
  labelId: string;
  /** Human-readable label text */
  label: string;
  /** Highlight color in hex format (default: "#FFEB3B") */
  color?: string;
  /** Author who created the annotation */
  author?: string;
  /** Text to search for and annotate */
  searchText?: string;
  /** Which occurrence to annotate (1-based, default: 1) */
  occurrence?: number;
  /** Start paragraph index (0-based) */
  startParagraphIndex?: number;
  /** End paragraph index (0-based, inclusive) */
  endParagraphIndex?: number;
  /** Custom metadata key-value pairs */
  metadata?: Record<string, string>;
}

/**
 * Response from adding an annotation.
 */
export interface AddAnnotationResponse {
  /** Whether the operation succeeded */
  success: boolean;
  /** The modified document as base64 string */
  documentBytes: string;
  /** The added annotation details */
  annotation?: Annotation;
}

/**
 * Response from removing an annotation.
 */
export interface RemoveAnnotationResponse {
  /** Whether the operation succeeded */
  success: boolean;
  /** The modified document as base64 string */
  documentBytes: string;
}

/**
 * Options for annotation rendering in HTML output.
 */
export interface AnnotationOptions {
  /** Whether to render annotations (default: false) */
  renderAnnotations?: boolean;
  /** How to display annotation labels (default: Above) */
  annotationLabelMode?: AnnotationLabelMode;
  /** CSS class prefix for annotation elements (default: "annot-") */
  annotationCssClassPrefix?: string;
}

// ============================================================================
// Document Structure Types (for element-based annotation targeting)
// ============================================================================

/**
 * Document element types that can be annotated.
 */
export enum DocumentElementType {
  /** Root document element */
  Document = "Document",
  /** A paragraph (w:p) */
  Paragraph = "Paragraph",
  /** A run within a paragraph (w:r) */
  Run = "Run",
  /** A table (w:tbl) */
  Table = "Table",
  /** A table row (w:tr) */
  TableRow = "TableRow",
  /** A table cell (w:tc) */
  TableCell = "TableCell",
  /** A virtual table column (not a real OOXML element) */
  TableColumn = "TableColumn",
  /** A hyperlink (w:hyperlink) */
  Hyperlink = "Hyperlink",
  /** An image/drawing (w:drawing) */
  Image = "Image",
}

/**
 * A document element in the structure tree.
 */
export interface DocumentElement {
  /** Unique element ID (path-based, e.g., "doc/tbl-0/tr-1/tc-2") */
  id: string;
  /** Element type */
  type: DocumentElementType | string;
  /** Preview of text content (first ~100 characters) */
  textPreview?: string;
  /** Position index within parent element */
  index: number;
  /** Child elements */
  children: DocumentElement[];
  /** For table rows/cells: the row index */
  rowIndex?: number;
  /** For table cells: the column index */
  columnIndex?: number;
  /** For table cells: number of rows this cell spans */
  rowSpan?: number;
  /** For table cells: number of columns this cell spans */
  columnSpan?: number;
}

/**
 * Information about a table column.
 */
export interface TableColumnInfo {
  /** ID of the table this column belongs to */
  tableId: string;
  /** Zero-based column index */
  columnIndex: number;
  /** IDs of all cells in this column */
  cellIds: string[];
  /** Total number of rows in this column */
  rowCount: number;
}

/**
 * Document structure analysis result.
 */
export interface DocumentStructure {
  /** Root document element */
  root: DocumentElement;
  /** All elements indexed by ID for quick lookup */
  elementsById: Record<string, DocumentElement>;
  /** Table column information indexed by column ID */
  tableColumns: Record<string, TableColumnInfo>;
}

/**
 * Target specification for element-based annotation.
 * Supports multiple targeting modes: element ID, indices, or text search.
 */
export interface AnnotationTarget {
  /** Target by element ID (e.g., "doc/p-0/r-1") */
  elementId?: string;
  /** Element type for index-based targeting */
  elementType?: DocumentElementType | string;
  /** Paragraph index (0-based) */
  paragraphIndex?: number;
  /** Run index within paragraph (0-based) */
  runIndex?: number;
  /** Table index (0-based) */
  tableIndex?: number;
  /** Row index within table (0-based) */
  rowIndex?: number;
  /** Cell index within row (0-based) */
  cellIndex?: number;
  /** Column index for table column targeting (0-based) */
  columnIndex?: number;
  /** Text to search for (global or within elementId) */
  searchText?: string;
  /** Which occurrence of searchText to target (1-based, default: 1) */
  occurrence?: number;
  /** End paragraph index for range targeting */
  rangeEndParagraphIndex?: number;
}

/**
 * Request to add an annotation using flexible targeting.
 */
export interface AddAnnotationWithTargetRequest {
  /** Unique annotation ID */
  id: string;
  /** Label category/type identifier */
  labelId: string;
  /** Human-readable label text */
  label: string;
  /** Highlight color in hex format (default: "#FFEB3B") */
  color?: string;
  /** Author who created the annotation */
  author?: string;
  /** Custom metadata key-value pairs */
  metadata?: Record<string, string>;
  /** Target specification */
  target: AnnotationTarget;
}

// ============================================================================
// Helper functions for document structure navigation
// ============================================================================

/**
 * Find an element by ID in the document structure.
 * @param structure - The document structure
 * @param elementId - The element ID to find
 * @returns The element or undefined if not found
 */
export function findElementById(
  structure: DocumentStructure,
  elementId: string
): DocumentElement | undefined {
  return structure.elementsById[elementId];
}

/**
 * Find all elements of a specific type in the document structure.
 * @param structure - The document structure
 * @param type - The element type to find
 * @returns Array of matching elements
 */
export function findElementsByType(
  structure: DocumentStructure,
  type: DocumentElementType | string
): DocumentElement[] {
  return Object.values(structure.elementsById).filter(
    (el) => el.type === type
  );
}

/**
 * Get all paragraphs from the document structure.
 * @param structure - The document structure
 * @returns Array of paragraph elements
 */
export function getParagraphs(
  structure: DocumentStructure
): DocumentElement[] {
  return findElementsByType(structure, DocumentElementType.Paragraph);
}

/**
 * Get all tables from the document structure.
 * @param structure - The document structure
 * @returns Array of table elements
 */
export function getTables(structure: DocumentStructure): DocumentElement[] {
  return findElementsByType(structure, DocumentElementType.Table);
}

/**
 * Get column information for a specific table.
 * @param structure - The document structure
 * @param tableId - The table ID
 * @returns Array of column info objects sorted by column index
 */
export function getTableColumns(
  structure: DocumentStructure,
  tableId: string
): TableColumnInfo[] {
  return Object.values(structure.tableColumns)
    .filter((col) => col.tableId === tableId)
    .sort((a, b) => a.columnIndex - b.columnIndex);
}

/**
 * Create an annotation target for an element by ID.
 * @param elementId - The element ID (e.g., "doc/p-0", "doc/tbl-0/tr-1/tc-2")
 * @returns AnnotationTarget object
 */
export function targetElement(elementId: string): AnnotationTarget {
  return { elementId };
}

/**
 * Create an annotation target for a paragraph by index.
 * @param paragraphIndex - Zero-based paragraph index
 * @returns AnnotationTarget object
 */
export function targetParagraph(paragraphIndex: number): AnnotationTarget {
  return {
    elementType: DocumentElementType.Paragraph,
    paragraphIndex,
  };
}

/**
 * Create an annotation target for a range of paragraphs.
 * @param startIndex - Zero-based start paragraph index
 * @param endIndex - Zero-based end paragraph index
 * @returns AnnotationTarget object
 */
export function targetParagraphRange(
  startIndex: number,
  endIndex: number
): AnnotationTarget {
  return {
    elementType: DocumentElementType.Paragraph,
    paragraphIndex: startIndex,
    rangeEndParagraphIndex: endIndex,
  };
}

/**
 * Create an annotation target for a specific run within a paragraph.
 * @param paragraphIndex - Zero-based paragraph index
 * @param runIndex - Zero-based run index within the paragraph
 * @returns AnnotationTarget object
 */
export function targetRun(
  paragraphIndex: number,
  runIndex: number
): AnnotationTarget {
  return {
    elementType: DocumentElementType.Run,
    paragraphIndex,
    runIndex,
  };
}

/**
 * Create an annotation target for a table by index.
 * @param tableIndex - Zero-based table index
 * @returns AnnotationTarget object
 */
export function targetTable(tableIndex: number): AnnotationTarget {
  return {
    elementType: DocumentElementType.Table,
    tableIndex,
  };
}

/**
 * Create an annotation target for a table row.
 * @param tableIndex - Zero-based table index
 * @param rowIndex - Zero-based row index within the table
 * @returns AnnotationTarget object
 */
export function targetTableRow(
  tableIndex: number,
  rowIndex: number
): AnnotationTarget {
  return {
    elementType: DocumentElementType.TableRow,
    tableIndex,
    rowIndex,
  };
}

/**
 * Create an annotation target for a table cell.
 * @param tableIndex - Zero-based table index
 * @param rowIndex - Zero-based row index
 * @param cellIndex - Zero-based cell index within the row
 * @returns AnnotationTarget object
 */
export function targetTableCell(
  tableIndex: number,
  rowIndex: number,
  cellIndex: number
): AnnotationTarget {
  return {
    elementType: DocumentElementType.TableCell,
    tableIndex,
    rowIndex,
    cellIndex,
  };
}

/**
 * Create an annotation target for a table column (all cells in that column).
 * @param tableIndex - Zero-based table index
 * @param columnIndex - Zero-based column index
 * @returns AnnotationTarget object
 */
export function targetTableColumn(
  tableIndex: number,
  columnIndex: number
): AnnotationTarget {
  return {
    elementType: DocumentElementType.TableColumn,
    tableIndex,
    columnIndex,
  };
}

/**
 * Create an annotation target by text search.
 * @param searchText - Text to search for
 * @param occurrence - Which occurrence to target (1-based, default: 1)
 * @returns AnnotationTarget object
 */
export function targetSearch(
  searchText: string,
  occurrence: number = 1
): AnnotationTarget {
  return { searchText, occurrence };
}

/**
 * Create an annotation target to search text within a specific element.
 * @param elementId - The element ID to search within
 * @param searchText - Text to search for
 * @param occurrence - Which occurrence to target (1-based, default: 1)
 * @returns AnnotationTarget object
 */
export function targetSearchInElement(
  elementId: string,
  searchText: string,
  occurrence: number = 1
): AnnotationTarget {
  return { elementId, searchText, occurrence };
}

// ============================================================================
// Document Metadata Types (Phase 3: Lazy Loading)
// ============================================================================

/**
 * Metadata for a single section in the document.
 * All dimension values are in points (1/72 inch).
 */
export interface SectionMetadata {
  /** Section index (0-based) */
  sectionIndex: number;
  /** Page width in points */
  pageWidthPt: number;
  /** Page height in points */
  pageHeightPt: number;
  /** Top margin in points */
  marginTopPt: number;
  /** Right margin in points */
  marginRightPt: number;
  /** Bottom margin in points */
  marginBottomPt: number;
  /** Left margin in points */
  marginLeftPt: number;
  /** Content width (page minus margins) in points */
  contentWidthPt: number;
  /** Content height (page minus margins) in points */
  contentHeightPt: number;
  /** Header distance from top in points */
  headerPt: number;
  /** Footer distance from bottom in points */
  footerPt: number;
  /** Number of paragraphs in this section */
  paragraphCount: number;
  /** Number of tables in this section */
  tableCount: number;
  /** Whether this section has a default header */
  hasHeader: boolean;
  /** Whether this section has a default footer */
  hasFooter: boolean;
  /** Whether this section has a first page header (titlePg enabled) */
  hasFirstPageHeader: boolean;
  /** Whether this section has a first page footer (titlePg enabled) */
  hasFirstPageFooter: boolean;
  /** Whether this section has an even page header */
  hasEvenPageHeader: boolean;
  /** Whether this section has an even page footer */
  hasEvenPageFooter: boolean;
  /** Start paragraph index (0-based, global across document) */
  startParagraphIndex: number;
  /** End paragraph index (exclusive, global across document) */
  endParagraphIndex: number;
  /** Start table index (0-based, global across document) */
  startTableIndex: number;
  /** End table index (exclusive, global across document) */
  endTableIndex: number;
}

/**
 * Document metadata for lazy loading pagination.
 * Provides fast access to document structure without full HTML rendering.
 */
export interface DocumentMetadata {
  /** List of sections with their metadata */
  sections: SectionMetadata[];
  /** Total number of paragraphs in the document */
  totalParagraphs: number;
  /** Total number of tables in the document */
  totalTables: number;
  /** Whether the document has any footnotes */
  hasFootnotes: boolean;
  /** Whether the document has any endnotes */
  hasEndnotes: boolean;
  /** Whether the document has tracked changes */
  hasTrackedChanges: boolean;
  /** Whether the document has comments */
  hasComments: boolean;
  /** Estimated total page count (rough estimate based on content) */
  estimatedPageCount: number;
}

// ============================================================================
// Web Worker Types (Phase 2: Non-blocking WASM operations)
// ============================================================================

/**
 * Message types sent from main thread to worker.
 */
export type WorkerRequestType =
  | "init"
  | "convertDocxToHtml"
  | "compareDocuments"
  | "compareDocumentsToHtml"
  | "getRevisions"
  | "getDocumentMetadata"
  | "getVersion";

/**
 * Base structure for worker requests.
 */
export interface WorkerRequestBase {
  /** Unique request ID for correlating responses */
  id: string;
  /** The operation type */
  type: WorkerRequestType;
}

/**
 * Initialize the worker with WASM base path.
 */
export interface WorkerInitRequest extends WorkerRequestBase {
  type: "init";
  /** Base URL for loading WASM files (e.g., "/wasm/") */
  wasmBasePath: string;
}

/**
 * Convert DOCX to HTML request.
 */
export interface WorkerConvertRequest extends WorkerRequestBase {
  type: "convertDocxToHtml";
  /** Document bytes (transferred, not copied) */
  documentBytes: Uint8Array;
  /** Conversion options */
  options?: ConversionOptions;
}

/**
 * Compare two documents request.
 */
export interface WorkerCompareRequest extends WorkerRequestBase {
  type: "compareDocuments";
  /** Original document bytes */
  originalBytes: Uint8Array;
  /** Modified document bytes */
  modifiedBytes: Uint8Array;
  /** Comparison options */
  options?: CompareOptions;
}

/**
 * Compare documents and return HTML request.
 */
export interface WorkerCompareToHtmlRequest extends WorkerRequestBase {
  type: "compareDocumentsToHtml";
  /** Original document bytes */
  originalBytes: Uint8Array;
  /** Modified document bytes */
  modifiedBytes: Uint8Array;
  /** Comparison options */
  options?: CompareOptions;
}

/**
 * Get revisions from a document request.
 */
export interface WorkerGetRevisionsRequest extends WorkerRequestBase {
  type: "getRevisions";
  /** Document bytes */
  documentBytes: Uint8Array;
  /** Revision extraction options */
  options?: GetRevisionsOptions;
}

/**
 * Get document metadata for lazy loading request.
 */
export interface WorkerGetDocumentMetadataRequest extends WorkerRequestBase {
  type: "getDocumentMetadata";
  /** Document bytes */
  documentBytes: Uint8Array;
}

/**
 * Get library version request.
 */
export interface WorkerGetVersionRequest extends WorkerRequestBase {
  type: "getVersion";
}

/**
 * Union type of all possible worker requests.
 */
export type WorkerRequest =
  | WorkerInitRequest
  | WorkerConvertRequest
  | WorkerCompareRequest
  | WorkerCompareToHtmlRequest
  | WorkerGetRevisionsRequest
  | WorkerGetDocumentMetadataRequest
  | WorkerGetVersionRequest;

/**
 * Base structure for worker responses.
 */
export interface WorkerResponseBase {
  /** Request ID this response corresponds to */
  id: string;
  /** Whether the operation succeeded */
  success: boolean;
  /** Error message if success is false */
  error?: string;
}

/**
 * Response from init request.
 */
export interface WorkerInitResponse extends WorkerResponseBase {
  type: "init";
}

/**
 * Response from convertDocxToHtml request.
 */
export interface WorkerConvertResponse extends WorkerResponseBase {
  type: "convertDocxToHtml";
  /** The converted HTML string */
  html?: string;
}

/**
 * Response from compareDocuments request.
 */
export interface WorkerCompareResponse extends WorkerResponseBase {
  type: "compareDocuments";
  /** The redlined document bytes */
  documentBytes?: Uint8Array;
}

/**
 * Response from compareDocumentsToHtml request.
 */
export interface WorkerCompareToHtmlResponse extends WorkerResponseBase {
  type: "compareDocumentsToHtml";
  /** The HTML string with redlines */
  html?: string;
}

/**
 * Response from getRevisions request.
 */
export interface WorkerGetRevisionsResponse extends WorkerResponseBase {
  type: "getRevisions";
  /** Array of revisions */
  revisions?: Revision[];
}

/**
 * Response from getDocumentMetadata request.
 */
export interface WorkerGetDocumentMetadataResponse extends WorkerResponseBase {
  type: "getDocumentMetadata";
  /** Document metadata */
  metadata?: DocumentMetadata;
}

/**
 * Response from getVersion request.
 */
export interface WorkerGetVersionResponse extends WorkerResponseBase {
  type: "getVersion";
  /** Version information */
  version?: VersionInfo;
}

/**
 * Union type of all possible worker responses.
 */
export type WorkerResponse =
  | WorkerInitResponse
  | WorkerConvertResponse
  | WorkerCompareResponse
  | WorkerCompareToHtmlResponse
  | WorkerGetRevisionsResponse
  | WorkerGetDocumentMetadataResponse
  | WorkerGetVersionResponse;

/**
 * Options for creating a worker-based Docxodus instance.
 */
export interface WorkerDocxodusOptions {
  /**
   * Base URL for loading WASM files.
   * Defaults to auto-detection from module URL.
   */
  wasmBasePath?: string;
}

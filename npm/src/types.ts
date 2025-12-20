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
 * Types of content that cannot be fully converted to HTML.
 * Used for placeholder rendering when unsupported content is encountered.
 */
export enum UnsupportedContentType {
  /** Windows Metafile image format (legacy vector graphics) */
  WmfImage = "WmfImage",
  /** Enhanced Metafile image format (legacy vector graphics) */
  EmfImage = "EmfImage",
  /** SVG image format (not yet supported) */
  SvgImage = "SvgImage",
  /** Office Math Markup Language equations */
  MathEquation = "MathEquation",
  /** Form field elements (checkboxes, text inputs, dropdowns) */
  FormField = "FormField",
  /** Ruby annotations for East Asian text */
  RubyAnnotation = "RubyAnnotation",
  /** Embedded OLE objects */
  OleObject = "OleObject",
  /** Other unsupported content */
  Other = "Other",
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
  /**
   * Whether to render placeholders for unsupported content (default: false)
   * When enabled, unsupported content (WMF/EMF images, math equations, form fields, etc.)
   * will display as styled placeholder spans instead of being silently dropped.
   */
  renderUnsupportedContentPlaceholders?: boolean;
  /**
   * Override the document's default language for the HTML lang attribute.
   * If not specified, the language is auto-detected from document settings
   * (w:themeFontLang or default paragraph style), falling back to "en-US".
   * Examples: "en-US", "fr-FR", "de-DE", "ja-JP"
   */
  documentLanguage?: string;
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
      renderMoveOperations: boolean,
      renderUnsupportedContentPlaceholders: boolean,
      documentLanguage: string | null
    ) => string;
    GetAnnotations: (bytes: Uint8Array) => string;
    AddAnnotation: (bytes: Uint8Array, requestJson: string) => string;
    AddAnnotationWithTarget: (bytes: Uint8Array, requestJson: string) => string;
    RemoveAnnotation: (bytes: Uint8Array, annotationId: string) => string;
    HasAnnotations: (bytes: Uint8Array) => string;
    GetDocumentStructure: (bytes: Uint8Array) => string;
    GetDocumentMetadata: (bytes: Uint8Array) => string;
    ExportToOpenContract: (bytes: Uint8Array) => string;
    GetVersion: () => string;
    // Profiling methods
    ConvertDocxToHtmlProfiled: (bytes: Uint8Array) => string;
    ProfileConversionDetailed: (bytes: Uint8Array) => string;
    // External annotation methods
    ComputeDocumentHash: (bytes: Uint8Array) => string;
    CreateExternalAnnotationSet: (
      bytes: Uint8Array,
      documentId: string
    ) => string;
    ValidateExternalAnnotations: (
      bytes: Uint8Array,
      annotationSetJson: string
    ) => string;
    ConvertDocxToHtmlWithExternalAnnotations: (
      bytes: Uint8Array,
      annotationSetJson: string,
      pageTitle: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      additionalCss: string,
      extAnnotCssClassPrefix: string,
      extAnnotLabelMode: number
    ) => string;
    SearchTextOffsets: (
      bytes: Uint8Array,
      searchText: string,
      maxResults: number
    ) => string;
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
 *
 * ## Units
 * All dimension values are in **points** (1 point = 1/72 inch).
 *
 * Common page sizes in points:
 * - **US Letter**: 612 × 792 pt (8.5" × 11")
 * - **A4**: 595 × 842 pt (210mm × 297mm)
 * - **Legal**: 612 × 1008 pt (8.5" × 14")
 *
 * To convert points to other units:
 * - Points to inches: `pt / 72`
 * - Points to mm: `pt / 72 * 25.4`
 * - Points to pixels (96 DPI): `pt * 96 / 72` (or `pt * 1.333...`)
 *
 * @example
 * ```typescript
 * const section = metadata.sections[0];
 * // US Letter: pageWidthPt = 612, pageHeightPt = 792
 *
 * // Convert to inches
 * const widthInches = section.pageWidthPt / 72; // 8.5
 *
 * // Convert to pixels at 96 DPI
 * const widthPx = section.pageWidthPt * 96 / 72; // 816
 * ```
 */
export interface SectionMetadata {
  /** Section index (0-based, sequential across document) */
  sectionIndex: number;
  /** Page width in points (1 pt = 1/72 inch). US Letter = 612pt, A4 ≈ 595pt */
  pageWidthPt: number;
  /** Page height in points (1 pt = 1/72 inch). US Letter = 792pt, A4 ≈ 842pt */
  pageHeightPt: number;
  /** Top margin in points. Default is typically 72pt (1 inch) */
  marginTopPt: number;
  /** Right margin in points. Default is typically 72pt (1 inch) */
  marginRightPt: number;
  /** Bottom margin in points. Default is typically 72pt (1 inch) */
  marginBottomPt: number;
  /** Left margin in points. Default is typically 72pt (1 inch) */
  marginLeftPt: number;
  /** Content width in points (pageWidthPt - marginLeftPt - marginRightPt) */
  contentWidthPt: number;
  /** Content height in points (pageHeightPt - marginTopPt - marginBottomPt) */
  contentHeightPt: number;
  /** Header distance from page top in points. Default is typically 36pt (0.5 inch) */
  headerPt: number;
  /** Footer distance from page bottom in points. Default is typically 36pt (0.5 inch) */
  footerPt: number;
  /** Number of paragraphs in this section (includes paragraphs inside tables) */
  paragraphCount: number;
  /** Number of top-level tables in this section */
  tableCount: number;
  /** Whether this section has a default header (w:headerReference type="default") */
  hasHeader: boolean;
  /** Whether this section has a default footer (w:footerReference type="default") */
  hasFooter: boolean;
  /** Whether this section has a first page header (requires titlePg element) */
  hasFirstPageHeader: boolean;
  /** Whether this section has a first page footer (requires titlePg element) */
  hasFirstPageFooter: boolean;
  /** Whether this section has an even page header (for different even/odd headers) */
  hasEvenPageHeader: boolean;
  /** Whether this section has an even page footer (for different even/odd footers) */
  hasEvenPageFooter: boolean;
  /** Start paragraph index (0-based, global across document). Inclusive. */
  startParagraphIndex: number;
  /** End paragraph index (global across document). Exclusive - use `endParagraphIndex - startParagraphIndex` for count. */
  endParagraphIndex: number;
  /** Start table index (0-based, global across document). Inclusive. */
  startTableIndex: number;
  /** End table index (global across document). Exclusive - use `endTableIndex - startTableIndex` for count. */
  endTableIndex: number;
}

/**
 * Document metadata for lazy loading pagination.
 * Provides fast access to document structure without full HTML rendering.
 *
 * This is significantly faster than full HTML conversion and is designed for:
 * - Lazy loading / virtual scrolling of large documents
 * - Pre-calculating pagination layouts
 * - Document feature detection before rendering
 *
 * @example
 * ```typescript
 * const metadata = await getDocumentMetadata(docxFile);
 *
 * // Check document features before rendering
 * if (metadata.hasTrackedChanges) {
 *   console.log('Document has tracked changes');
 * }
 *
 * // Calculate total content for lazy loading
 * const totalSections = metadata.sections.length;
 * const firstPageWidth = metadata.sections[0].pageWidthPt;
 *
 * // Paragraph count includes paragraphs inside tables
 * console.log(`${metadata.totalParagraphs} paragraphs, ${metadata.totalTables} tables`);
 * ```
 *
 * @remarks
 * **Limitations:**
 * - Section breaks inside tables or text boxes are not detected (see GitHub issue #51)
 * - Estimated page count is heuristic-based and may not match actual rendered pages
 * - Maximum document size is 100MB
 */
export interface DocumentMetadata {
  /** List of sections with their metadata. Documents always have at least one section. */
  sections: SectionMetadata[];
  /** Total number of paragraphs in the document (includes paragraphs inside tables) */
  totalParagraphs: number;
  /** Total number of top-level tables in the document */
  totalTables: number;
  /** Whether the document has any footnotes (excludes separator/continuationSeparator) */
  hasFootnotes: boolean;
  /** Whether the document has any endnotes (excludes separator/continuationSeparator) */
  hasEndnotes: boolean;
  /** Whether the document has tracked changes (w:ins, w:del, w:moveFrom, w:moveTo) */
  hasTrackedChanges: boolean;
  /** Whether the document has comments (w:comment elements with content) */
  hasComments: boolean;
  /** Estimated total page count (heuristic based on content volume and page sizes) */
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

// ============================================================================
// OpenContracts Export Types (Issue #56)
// ============================================================================

/**
 * OpenContracts document export format.
 * Compatible with the OpenContracts ecosystem for document analysis.
 *
 * @example
 * ```typescript
 * const export = await exportToOpenContract(docxFile);
 * console.log(`Title: ${export.title}`);
 * console.log(`Content length: ${export.content.length} characters`);
 * console.log(`Pages: ${export.pageCount}`);
 * console.log(`Structural annotations: ${export.labelledText.filter(a => a.structural).length}`);
 * ```
 */
export interface OpenContractDocExport {
  /** Document title (from core properties or filename) */
  title: string;
  /** Complete document text content - ALL text from the document */
  content: string;
  /** Optional document description */
  description?: string;
  /** Estimated page count */
  pageCount: number;
  /** PAWLS-format page layout information with token positions */
  pawlsFileContent: PawlsPage[];
  /** Document-level labels (categories applied to the whole document) */
  docLabels: string[];
  /** Annotations/labeled text spans in the document */
  labelledText: OpenContractsAnnotation[];
  /** Relationships between annotations */
  relationships?: OpenContractsRelationship[];
}

/**
 * PAWLS page containing page boundary and token information.
 * PAWLS (Page-Aware Layout Segmentation) is a format for document layout data.
 */
export interface PawlsPage {
  /** Page boundary information (dimensions and index) */
  page: PawlsPageBoundary;
  /** Tokens on this page with position information */
  tokens: PawlsToken[];
}

/**
 * Page boundary information for PAWLS format.
 */
export interface PawlsPageBoundary {
  /** Page width in points (1pt = 1/72 inch) */
  width: number;
  /** Page height in points */
  height: number;
  /** Zero-based page index */
  index: number;
}

/**
 * Token with position information for PAWLS format.
 * Each token represents a word or text fragment with its bounding box.
 */
export interface PawlsToken {
  /** X coordinate (left edge) in points */
  x: number;
  /** Y coordinate (top edge) in points */
  y: number;
  /** Token width in points */
  width: number;
  /** Token height in points */
  height: number;
  /** The text content of this token */
  text: string;
}

/**
 * OpenContracts annotation format.
 * Used for both user annotations and structural elements.
 */
export interface OpenContractsAnnotation {
  /** Unique annotation identifier */
  id?: string;
  /** Label/category for this annotation (e.g., "SECTION", "PARAGRAPH", "CLAUSE_TYPE_A") */
  annotationLabel: string;
  /** The raw text content of the annotation */
  rawText: string;
  /** Starting page number (0-indexed) */
  page: number;
  /**
   * Position data for the annotation. Can be either:
   * - A TextSpan with start/end character offsets
   * - A dictionary of page indices to single-page annotation data
   */
  annotationJson?: TextSpan | Record<string, OpenContractsSinglePageAnnotation>;
  /** Parent annotation ID for hierarchical annotations */
  parentId?: string;
  /** Type of annotation (e.g., "text", "structural") */
  annotationType?: string;
  /** Whether this is a structural element (section, heading, table, etc.) */
  structural: boolean;
}

/**
 * Per-page annotation position data.
 * Used when an annotation spans multiple pages.
 */
export interface OpenContractsSinglePageAnnotation {
  /** Bounding box for the annotation on this page */
  bounds: BoundingBox;
  /** Token indices that make up this annotation on this page */
  tokensJsons: TokenId[];
  /** Raw text content on this page */
  rawText: string;
}

/**
 * Bounding box coordinates in points.
 */
export interface BoundingBox {
  /** Top edge coordinate */
  top: number;
  /** Bottom edge coordinate */
  bottom: number;
  /** Left edge coordinate */
  left: number;
  /** Right edge coordinate */
  right: number;
}

/**
 * Token identifier referencing a specific token on a specific page.
 */
export interface TokenId {
  /** Zero-based page index */
  pageIndex: number;
  /** Zero-based token index within the page */
  tokenIndex: number;
}

/**
 * Text span with character offsets for annotation positioning.
 * This is the simpler form of annotation_json for single-page annotations.
 */
export interface TextSpan {
  /** Optional span identifier */
  id?: string;
  /** Start character offset (0-indexed, inclusive) */
  start: number;
  /** End character offset (exclusive) */
  end: number;
  /** The text content of this span */
  text: string;
}

/**
 * Relationship between annotations.
 * Used to express hierarchical or semantic connections between annotations.
 */
export interface OpenContractsRelationship {
  /** Unique relationship identifier */
  id?: string;
  /** Label describing the relationship type (e.g., "CONTAINS", "REFERENCES") */
  relationshipLabel: string;
  /** IDs of source annotations */
  sourceAnnotationIds: string[];
  /** IDs of target annotations */
  targetAnnotationIds: string[];
  /** Whether this is a structural relationship */
  structural: boolean;
}

// ============================================================================
// External Annotation Types (Issue #57)
// ============================================================================

/**
 * Annotation label definition matching OpenContracts AnnotationLabelPythonType.
 */
export interface AnnotationLabel {
  /** Unique label identifier */
  id: string;
  /** Color in hex format (e.g., "#FFEB3B") */
  color: string;
  /** Description of what this label represents */
  description: string;
  /** Optional icon name */
  icon: string;
  /** Display name for the label */
  text: string;
  /** Type of label: "text", "doc", or "metadata" */
  labelType: "text" | "doc" | "metadata";
}

/**
 * External annotation set - extends OpenContractDocExport with binding/validation.
 * This allows storing annotations externally (in JSON/database) without modifying the DOCX.
 *
 * @example
 * ```typescript
 * // Create an annotation set from a document
 * const set = await createExternalAnnotationSet(docxFile, "my-doc-123");
 *
 * // Add a label definition
 * set.textLabels["IMPORTANT"] = {
 *   id: "IMPORTANT",
 *   text: "Important",
 *   color: "#FF0000",
 *   description: "Important text that needs attention",
 *   icon: "",
 *   labelType: "text"
 * };
 *
 * // Create an annotation
 * const annotation = createAnnotationFromSearch(
 *   "ann-001", "IMPORTANT", set.content, "contract term"
 * );
 * if (annotation) {
 *   set.labelledText.push(annotation);
 * }
 *
 * // Validate and project onto HTML
 * const result = await validateExternalAnnotations(docxFile, set);
 * if (result.isValid) {
 *   const html = await convertDocxToHtmlWithExternalAnnotations(docxFile, set);
 * }
 * ```
 */
export interface ExternalAnnotationSet extends OpenContractDocExport {
  /** Unique identifier for the source document (filename, UUID, or external reference) */
  documentId: string;
  /** SHA256 hash of the source document for integrity validation */
  documentHash: string;
  /** ISO 8601 timestamp when this annotation set was created */
  createdAt: string;
  /** ISO 8601 timestamp when this annotation set was last modified */
  updatedAt: string;
  /** Version of the external annotation format (for future migrations) */
  version: string;
  /** Text label definitions keyed by label ID */
  textLabels: Record<string, AnnotationLabel>;
  /** Document label definitions keyed by label ID */
  docLabelDefinitions: Record<string, AnnotationLabel>;
}

/**
 * Result of validating an external annotation set against a document.
 */
export interface ExternalAnnotationValidationResult {
  /** True if the annotation set is valid for the document */
  isValid: boolean;
  /** True if the document hash doesn't match, indicating the document may have been modified */
  hashMismatch: boolean;
  /** List of specific issues found during validation */
  issues: ExternalAnnotationValidationIssue[];
}

/**
 * A single validation issue found when validating an external annotation set.
 */
export interface ExternalAnnotationValidationIssue {
  /** ID of the annotation with the issue */
  annotationId: string;
  /** Type of issue: "TextMismatch", "OutOfBounds", or "MissingLabel" */
  issueType: "TextMismatch" | "OutOfBounds" | "MissingLabel";
  /** Human-readable description of the issue */
  description: string;
  /** For TextMismatch: the text that was expected (stored in annotation) */
  expectedText?: string;
  /** For TextMismatch: the actual text found at the annotation's offsets */
  actualText?: string;
}

/**
 * Settings for projecting external annotations onto HTML.
 */
export interface ExternalAnnotationProjectionSettings {
  /** CSS class prefix for annotation elements (default: "ext-annot-") */
  cssClassPrefix?: string;
  /** How to display annotation labels (default: Above) */
  labelMode?: AnnotationLabelMode;
  /** Whether to include annotation metadata as data attributes (default: true) */
  includeMetadata?: boolean;
  /** Whether to validate annotations before projection (default: true) */
  validateBeforeProjection?: boolean;
}


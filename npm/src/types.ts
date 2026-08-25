/**
 * Which parts of the OOXML package to include in the markdown projection.
 * Bitflags mirroring the .NET `Docxodus.ProjectionScopes` enum.
 */
export enum ProjectionScopes {
  Body = 1,
  Headers = 2,
  Footers = 4,
  Footnotes = 8,
  Endnotes = 16,
  Comments = 32,
  All = 63,
}

/**
 * How anchor markers are rendered in the markdown projection.
 * Mirrors the .NET `Docxodus.AnchorRenderMode` enum.
 */
export enum AnchorRenderMode {
  /** Anchor appears on its own line before each block element (default). */
  Block = 0,
  /** Block anchors plus inline `{#…}` markers for spans (comments, hyperlinks). */
  BlockAndInline = 1,
  /** No anchor markers in the output (projection only, no addressing). */
  None = 2,
}

/**
 * Strategy for rendering `w:tbl` elements that don't fit GFM pipe-table constraints.
 * Mirrors the .NET `Docxodus.TableRenderMode` enum.
 */
export enum TableRenderMode {
  /** Emit GFM pipe tables when possible, opaque anchor blocks otherwise (default). */
  GfmWithOpaqueFallback = 0,
  /** Always emit GFM pipe tables, flattening complex structure with possible loss. */
  AlwaysGfm = 1,
  /** Always emit opaque anchor blocks. */
  AlwaysOpaque = 2,
}

/**
 * How tracked changes are handled in the markdown projection.
 * Mirrors the .NET `Docxodus.TrackedChangeMode` enum.
 */
export enum TrackedChangeMode {
  /** Accept all revisions before conversion (default). */
  Accept = 0,
  /** Render insertions and deletions inline as `{+ins+}` / `{-del-}`. */
  RenderInline = 1,
  /** Accept insertions, drop deletions. */
  StripDeletions = 2,
}

/**
 * How empty paragraphs are rendered. Mirrors the .NET `Docxodus.EmptyParagraphMode` enum.
 */
export enum EmptyParagraphMode {
  /** Default: emit the anchor on its own line (`{#p:body:UNID}\n`). */
  AnchorOnly = 0,
  /** Tag the empty paragraph visibly so agents can pattern-match (`{#p:body:UNID} ∅`). */
  MarkedEmpty = 1,
  /** Skip empty paragraphs entirely — they don't appear in the markdown or the anchor index. */
  Suppress = 2,
}

/**
 * How anchor ids are rendered inside `{#…}` tokens (and keyed in
 * {@link MarkdownProjection.anchorIndex}). Mirrors the .NET
 * `Docxodus.AnchorIdRendering` enum.
 *
 * `Anchor.token` (and any `AnchorRef` returned by `DocxSession` mutations)
 * always carries the full Unid regardless of rendering — the choice only
 * affects the markdown text. The returned `anchorIndex` is dual-keyed so
 * lookups by either the rendered id or the full Unid work for the
 * `Abbreviated`/`Sequential` modes.
 */
export enum AnchorIdRendering {
  /** Full 32-char hex Unid (default; e.g. `{#h:body:a1b2c3d4e5f6789012345678901234ab}`). */
  FullUnid = 0,
  /** Shortest unique prefix per (kind, scope) bucket, 4-char floor (e.g. `{#h:body:a1b2}`).
   *  Saves 5-10% of projection-token budget for LLM consumption. */
  Abbreviated = 1,
  /** Sequential numeric ids per (kind, scope) bucket in document order (e.g. `{#h:body:1}`).
   *  Maximally token-efficient for one-shot LLM contexts. NOT stable across
   *  `project()` calls and must NOT be persisted. */
  Sequential = 2,
}

/**
 * How far below the target anchor to include in
 * {@link DocxSession.projectAnchor}. Mirrors the .NET
 * `Docxodus.ProjectionDepth` enum.
 */
export enum ProjectionDepth {
  /** Just the target block itself (its anchor + its own text). For headings,
   *  returns only the heading paragraph, not the section under it. */
  SelfOnly = 0,
  /** Self + descendants. Most useful for `tbl` anchors (returns the whole table);
   *  for paragraphs it's the same as `SelfOnly`. */
  Subtree = 1,
  /** Self + descendants + following siblings up to (but not including) the next
   *  sibling at the same or higher heading level. For non-heading anchors,
   *  equivalent to `Subtree`. Dominant "give me this section" case; the default. */
  SubtreeAndFollowingSiblings = 2,
}

/**
 * Settings controlling the markdown projection. Mirrors the .NET
 * `WmlToMarkdownConverterSettings` class — see `docs/architecture/markdown_projection.md`.
 */
export interface MarkdownProjectionSettings {
  scopes?: ProjectionScopes;
  headingLevelOffset?: number;
  anchorMode?: AnchorRenderMode;
  tableMode?: TableRenderMode;
  tableInlineCellMax?: number;
  trackedChanges?: TrackedChangeMode;
  resolveNumbering?: boolean;
  emptyParagraphs?: EmptyParagraphMode;
  /**
   * How anchor ids are rendered in markdown output. Default `FullUnid`.
   * Set to `Abbreviated` for terse LLM-friendly ids; `Sequential` for
   * 1-based per-scope counters (best for replay logs / human review).
   * `Anchor.token` (and any `AnchorRef`) always reflects the full Unid
   * regardless; this only affects the markdown text. Use the returned
   * {@link MarkdownProjection.anchorIndex} (dual-keyed for Abbreviated /
   * Sequential) to translate rendered ids back to full Unids.
   */
  anchorIdRendering?: AnchorIdRendering;
}

/**
 * Resolved location of an anchor in the underlying OOXML package — sufficient to walk
 * back to the source element via the .NET API.
 */
export interface MarkdownAnchorTarget {
  id: string;
  kind: string;
  scope: string;
  unid: string;
  partUri: string;
  /** First ~80 characters of the element's flat text — for previewing/picking anchors. */
  textPreview: string;
  /** Resolved auto-numbering prefix (e.g. "1.", "First") for paragraphs/headings/list
   *  items whose style or `w:numPr` produces numbering. Absent when the element has
   *  no numbering. The prefix is NOT included in {@link textPreview} because
   *  textPreview reflects only the run text; this field gives callers the value
   *  Word actually renders before the element's text. */
  autoNumberPrefix?: string;
}

/**
 * Output of the markdown projection: rendered text plus the anchor index mapping
 * each `{#…}` token back to a location in the OOXML package.
 */
export interface MarkdownProjection {
  markdown: string;
  anchorIndex: Record<string, MarkdownAnchorTarget>;
}

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
 * Which comparison engine {@link compareDocuments} / {@link compareDocumentsToHtml}
 * use to redline two documents.
 *
 * The numeric values are the contract shared with the .NET `ComparisonEngine` enum
 * and the WASM boundary (marshalled as an int); `0` remains
 * {@link ComparisonEngine.WmlComparer} for wire compatibility. Omitted selectors use
 * {@link ComparisonEngine.DocxDiff}.
 *
 * `DocxDiff` is the default comparison engine. `WmlComparer` remains available for
 * callers that explicitly require its historical behavior.
 */
export enum ComparisonEngine {
  /** The legacy WmlComparer engine (retained at wire value 0). */
  WmlComparer = 0,
  /** The default DocxDiff IR diff engine. */
  DocxDiff = 1,
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
  /**
   * Stamp block-level elements (p, h1-h6, li, table) with a `data-anchor`
   * attribute carrying the block's stable Unid. Required for the editor to
   * address blocks in the DOM and drive incremental per-block re-render via
   * `renderBlockHtml`. Default: false.
   */
  stampAnchors?: boolean;
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
  /**
   * Which comparison engine to use (default: {@link ComparisonEngine.DocxDiff}).
   * Pass {@link ComparisonEngine.WmlComparer} only when its historical behavior is required.
   */
  engine?: ComparisonEngine;
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
 * Which property container a FormatChanged revision describes.
 * `run` (the default) is an rPr-grade report (bold/italic/fontSize/…); the others are the
 * block-and-above scopes tracked by the block-format-change family: `paragraph` (pPr),
 * `tableCell`/`tableRow`/`table` (tcPr/trPr/tblPr+tblGrid), and `section` (sectPr).
 * Non-`run` scopes are only reported under Fine revision granularity.
 */
export type FormatChangeScope =
  | "run"
  | "paragraph"
  | "tableCell"
  | "tableRow"
  | "table"
  | "section";

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
   * For the table/section scopes this is a digest-grade marker (`["shell"]`, `["grid"]`).
   */
  changedPropertyNames?: string[];
  /**
   * Which property container this change describes (default `"run"`).
   */
  scope?: FormatChangeScope;
}

// ─── DocxDiff (IR diff engine) ──────────────────────────────────────────────
//
// The structure-aware default comparison engine. Its specialized APIs expose
// anchor-addressed revisions (`leftAnchor`/`rightAnchor`) and the diff-as-data
// edit script (`docxDiffGetEditScript`).

/**
 * How `docxDiffGetRevisions` projects the edit script to revisions. Integer
 * values match the .NET `DocxDiffRevisionGranularity` enum positions.
 */
export enum DocxDiffRevisionGranularity {
  /** The engine's native one-revision-per-token-span grain (the default). */
  Fine = 0,
  /** Coalesced to counts/texts comparable to the shipped WmlComparer's. */
  WmlComparerCompatible = 1,
}

/**
 * How DocxDiff compares run formatting. Integer values match the .NET
 * `DocxDiffFormatComparison` enum positions.
 */
export enum DocxDiffFormatComparison {
  /** Compare only the modeled rPr fields (the default). */
  ModeledOnly = 0,
  /** Compare the full run format including the unmodeled rPr digest. */
  Full = 1,
}

/**
 * Settings for the `docxDiff*` functions. Mirrors the .NET `DocxDiffSettings`;
 * every field is optional and an omitted field uses the engine default.
 */
export interface DocxDiffSettings {
  /** Author stamped on revisions and markup (default "Open-Xml-PowerTools"). */
  authorForRevisions?: string;
  /** Pin revision dates to a fixed epoch for byte-identical output (default true). */
  deterministic?: boolean;
  /** Explicit ISO-8601 revision date; overrides `deterministic` when set. */
  dateTimeForRevisions?: string;
  /**
   * Accept every pre-existing revision on both inputs before comparing (default false).
   * This flattens prior authorship; `preserveInputRevisions` takes precedence when both are set.
   */
  preAcceptInputRevisions?: boolean;
  /**
   * Preserve the inputs' existing tracked revisions Word-style (default false).
   * This takes precedence over `preAcceptInputRevisions` when both are set.
   */
  preserveInputRevisions?: boolean;
  /** Case-fold word match keys (default false). */
  caseInsensitive?: boolean;
  /** Culture name (e.g. "tr-TR") for case folding when `caseInsensitive` is true; default invariant/ordinal. */
  culture?: string;
  /** Fold NBSP (U+00A0) to ordinary space in match keys (default true). */
  conflateBreakingAndNonbreakingSpaces?: boolean;
  /** Override the word/separator split characters (default: engine's set). */
  wordSeparators?: string;
  /** Report relocations as native move pairs (default true). */
  detectMoves?: boolean;
  /** Jaccard similarity threshold for a fuzzy move 0.0-1.0 (default 0.8). */
  moveSimilarityThreshold?: number;
  /** Minimum word tokens for a fuzzy move (default 3). */
  moveMinimumWordCount?: number;
  /** Revision projection grain (default Fine). */
  revisionGranularity?: DocxDiffRevisionGranularity;
  /** Run-format comparison policy (default ModeledOnly). */
  formatComparison?: DocxDiffFormatComparison;
  /**
   * Compare header/footer stories (default true — Word Compare's own default).
   * Changed stories get native tracked-changes markup inside their parts;
   * Fine-mode revisions carry `hdr`/`ftr`-scoped anchors; the edit script
   * carries `headerFooterOps`. Set false to ignore header/footer scopes (the
   * pre-campaign behavior: left's headers/footers carried verbatim).
   */
  compareHeadersFooters?: boolean;

  /**
   * Track paragraph-and-above property changes (pPr/tcPr/trPr/tblPr/tblGrid/tblPrEx/sectPr) as native
   * Word markup. Default true. Set false to restore the pre-campaign untracked-right-apply behavior.
   * (Consolidate ignores block-format changes regardless.)
   */
  trackBlockFormatChanges?: boolean;

  /**
   * Render a run of >=2 adjacent word-matched modified paragraph pairs via a single cross-paragraph
   * word+pilcrow token-stream diff — the within-run flat-stream shape decoded from Word's compare
   * output (retained words may cross the pilcrow; paragraph marks are ins/del stream tokens; the
   * output paragraph count follows the token-level interleave). Markup (`docxDiffCompare`) only —
   * the revision list and edit-script JSON are unaffected. Default false.
   */
  crossParagraphTokenDiff?: boolean;
}

/**
 * One revision from `docxDiffGetRevisions`. Mirrors the consumer shape of
 * {@link Revision} and ADDS the block anchors the revision derives from — the
 * IR engine's differentiator.
 *
 * Anchor presence by `revisionType` — each type's PRIMARY anchor is ALWAYS
 * present; the opposite anchor MAY also be present for a token-level revision.
 * Inserted → `rightAnchor` always (plus `leftAnchor` when it is a token-level
 * insert inside a modified block); Deleted → `leftAnchor` always (plus
 * `rightAnchor` when token-level); FormatChanged → both; Moved is EXCLUSIVE:
 * source → `leftAnchor` only, destination → `rightAnchor` only. A token-level
 * revision (an insert/delete WITHIN a modified paragraph that exists on both
 * sides) carries both enclosing-block anchors; a whole-block insert/delete
 * carries only its primary anchor.
 */
export interface DocxDiffRevision {
  /** "Inserted" | "Deleted" | "Moved" | "FormatChanged". */
  revisionType: RevisionType | string;
  /** The affected text. */
  text: string;
  /** Author stamped on the revision. */
  author: string;
  /** ISO-8601 revision date. */
  date: string;
  /** For Moved revisions, links source and destination. */
  moveGroupId?: number;
  /** For Moved revisions: true = source, false = destination. */
  isMoveSource?: boolean;
  /** For FormatChanged revisions: what formatting changed. */
  formatChange?: FormatChangeDetails;
  /** The LEFT-document block anchor (`kind:scope:unid`). Always set for Deleted/FormatChanged/Moved-source;
   *  also set for a token-level insert inside a modified block; undefined for a whole-block insertion. */
  leftAnchor?: string;
  /** The RIGHT-document block anchor (`kind:scope:unid`). Always set for Inserted/FormatChanged/Moved-dest;
   *  also set for a token-level delete inside a modified block; undefined for a whole-block deletion. */
  rightAnchor?: string;
}

/** Stable v1 operation names emitted by the semantic-change schema. */
export type SemanticChangeOperation = "insert" | "delete" | "move" | "modify";

/** Stable v1 semantic families. New schema versions may append families. */
export type SemanticChangeFamily =
  | "text"
  | "block_structure"
  | "run_formatting"
  | "paragraph_formatting"
  | "style"
  | "numbering"
  | "list"
  | "table"
  | "table_row"
  | "table_cell"
  | "table_span"
  | "table_width"
  | "table_style"
  | "section"
  | "page_setup"
  | "header"
  | "footer"
  | "field"
  | "footnote"
  | "endnote"
  | "comment"
  | "hyperlink"
  | "bookmark"
  | "content_control"
  | "image"
  | "media"
  | "relationship"
  | "revision"
  | "annotation"
  | "opaque_package_part";

/**
 * Closed typed value used in {@link SemanticChange.before} and `after`.
 * Schema v1 integers stay within `Number.MIN_SAFE_INTEGER..Number.MAX_SAFE_INTEGER`; a document
 * value outside that range arrives as a decimal `string` rather than a rounded `integer`.
 */
export type SemanticValue =
  | { kind: "absent" }
  | { kind: "string"; value: string }
  | { kind: "boolean"; value: boolean }
  | { kind: "integer"; value: number }
  | { kind: "digest"; algorithm: string; profile: string | null; value: string }
  | { kind: "object"; value: Record<string, SemanticValue> }
  | { kind: "array"; value: SemanticValue[] };

/** One deterministic, anchor-addressed semantic change. */
export interface SemanticChange {
  id: string;
  operation: SemanticChangeOperation;
  family: SemanticChangeFamily;
  partUri: string;
  path: string;
  leftAnchor: string | null;
  rightAnchor: string | null;
  leftScope: string | null;
  rightScope: string | null;
  moveId: string | null;
  before: SemanticValue;
  after: SemanticValue;
}

/** Public `docxodus.semantic-changes` schema returned by semantic comparison APIs. */
export interface SemanticChangeSet {
  schema: "docxodus.semantic-changes";
  schemaVersion: 1;
  changeCount: number;
  changes: SemanticChange[];
}

// ─── DocxDiff consolidate (composite N-way) ─────────────────────────────────
//
// The composite layer over the IR diff engine: merge several reviewers' edits
// against one shared base DOCX. Mirrors the .NET consolidate surface
// (`DocxDiff.Consolidate` / `GetConflicts` / `GetConsolidatedRevisions` /
// `GetConsolidatedEditScriptJson`) — see `docs/architecture/ir_diff_engine.md`.

/**
 * How overlapping reviewer edits at the same base token span are resolved.
 * Integer values match the .NET `ConflictResolution` enum positions.
 */
export enum ConflictResolution {
  /** Keep the base text where reviewers disagree (the default). */
  BaseWins = 0,
  /** The first reviewer (by input order) wins each contested span. */
  FirstReviewerWins = 1,
  /** Stack every reviewer's variant so a human can pick. */
  StackAll = 2,
}

/** One reviewer's edited copy of the shared base document, plus their name. */
export interface DocxDiffReviewer {
  /** The reviewer's edited DOCX bytes. */
  document: Uint8Array;
  /** Author name stamped on this reviewer's revisions in the consolidated output. */
  author: string;
}

/**
 * Settings for the `docxDiffConsolidate*` functions. Extends {@link DocxDiffSettings}
 * with the conflict-resolution policy; mirrors the .NET `DocxDiffConsolidateSettings`.
 */
export interface DocxDiffConsolidateSettings extends DocxDiffSettings {
  /** Policy for overlapping reviewer edits (default {@link ConflictResolution.BaseWins}). */
  conflictResolution?: ConflictResolution;
}

/** One reviewer's contested variant at a conflict span. */
export interface DocxDiffConflictCompetitor {
  /** The reviewer whose variant this is. */
  author: string;
  /** The text this reviewer would produce for the contested span. */
  resultText: string;
}

/**
 * A single conflict: two or more reviewers edited the same base token span
 * incompatibly. Returned by {@link docxDiffGetConflicts}.
 */
export interface DocxDiffConflict {
  /** Stable id for this conflict within the consolidation. */
  id: number;
  /** The base-document block anchor (`kind:scope:unid`) the conflict sits in. */
  baseAnchor: string;
  /** First contested base token index (inclusive). */
  tokenStart: number;
  /** Last contested base token index (exclusive). */
  tokenEnd: number;
  /** The {@link ConflictResolution} policy that was applied to this conflict. */
  policy: ConflictResolution | number;
  /** The competing reviewer variants. */
  competitors: DocxDiffConflictCompetitor[];
}

/**
 * One revision from {@link docxDiffGetConsolidatedRevisions}. A
 * {@link DocxDiffRevision} plus the id of the conflict it participates in
 * (when it sits in a contested span).
 */
export interface DocxDiffConsolidatedRevision extends DocxDiffRevision {
  /** The {@link DocxDiffConflict.id} this revision belongs to, if any. */
  conflictId?: number;
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

/** Algorithm-labelled digest in a verification artifact. */
export interface VerificationDigest {
  algorithm: string;
  /** Lower-case hexadecimal digest bytes. */
  value: string;
}

/** Stable package location attached to a verification finding. */
export interface ChangeLocation {
  entryUri: string | null;
  ownerUri: string | null;
  relationshipId: string | null;
  targetUri: string | null;
  propertyPath: string | null;
}

export type VerificationFindingSeverity = "info" | "warning" | "error";

/** Machine-readable package validation or safety finding. */
export interface VerificationFinding {
  code: string;
  severity: VerificationFindingSeverity;
  message: string;
  location: ChangeLocation | null;
}

/** One physical ZIP entry. Duplicate names remain separate occurrences. */
export interface PackageManifestEntry {
  uri: string;
  occurrence: number;
  contentType: string | null;
  contentTypeSource: "override" | "default" | "implicit" | "unresolved";
  /** Exact declared uncompressed byte length as a base-10 integer string. */
  size: string;
  /** Exact compressed ZIP byte length as a base-10 integer string. */
  compressedSize: string;
  rawBytesDigest: VerificationDigest | null;
  normalizedXmlDigest: VerificationDigest | null;
  isXml: boolean;
  /** null when central-directory encryption flags could not be parsed authoritatively. */
  isEncrypted: boolean | null;
}

export interface PackageContentTypeDeclaration {
  kind: "default" | "override";
  key: string;
  contentType: string;
  occurrence: number;
}

export interface PackageRelationship {
  ownerUri: string;
  id: string;
  type: string;
  target: string;
  targetMode: "Internal" | "External";
  resolvedTargetUri: string | null;
  isTargetPresent: boolean | null;
}

export interface PackageRevisionCounts {
  insertions: number;
  deletions: number;
  moveFrom: number;
  moveTo: number;
  propertyChanges: number;
  /** The `rPrChange` subset of `propertyChanges`; not added into `total`. */
  runPropertyChanges: number;
  structuralChanges: number;
  otherChanges: number;
  total: number;
}

export interface PackageAnnotationCounts {
  comments: number;
  commentReplies: number;
  threadedCommentMetadata: number;
  resolvedComments: number;
  people: number;
  docxodusAnnotations: number;
}

export interface PackageManifestFacts {
  mainDocumentUri: string | null;
  isStrictOoxml: boolean;
  isMacroEnabled: boolean;
  hasCoreProperties: boolean;
  hasExtendedProperties: boolean;
  hasCustomProperties: boolean;
  sectionCount: number;
  paragraphCount: number;
  tableCount: number;
  headerPartCount: number;
  footerPartCount: number;
  footnoteCount: number;
  endnoteCount: number;
  styleCount: number;
  numberingDefinitionCount: number;
  themePartCount: number;
  mediaPartCount: number;
  customXmlPartCount: number;
  drawingCount: number;
  altChunkCount: number;
  fieldCount: number;
  revisions: PackageRevisionCounts;
  annotations: PackageAnnotationCounts;
}

/** Deterministic schema-v1 description of a DOCX/OPC package. */
export interface PackageManifest {
  schema: "https://docxodus.dev/schemas/verification/package-manifest/v1";
  schemaVersion: 1;
  packageKind: "opc" | "zip" | "zip-encrypted" | "ole-encrypted" | "ole" | "malformed";
  isValid: boolean;
  rawPackageBytesDigest: VerificationDigest;
  orderedOpcContentDigest: VerificationDigest | null;
  normalizedSemanticDigest: VerificationDigest | null;
  entries: readonly PackageManifestEntry[];
  contentTypes: readonly PackageContentTypeDeclaration[];
  relationships: readonly PackageRelationship[];
  facts: PackageManifestFacts;
  findings: readonly VerificationFinding[];
}

// ============================================================================
// Deliverable verification
// ============================================================================

/** Policy preset used by the default deliverable gate. */
export type DeliverableVerificationMode = "standard" | "strict" | "reportOnly";

/** Final policy decision in a deliverable-verification report. */
export type DeliverableVerificationDecision =
  | "passed"
  | "passedWithPreExistingFindings"
  | "failed"
  | "notEvaluated";

/** How a delivered finding relates to the exact opening/baseline package. */
export type DeliverableFindingDisposition =
  | "new"
  | "preExisting"
  | "resolved"
  | "unclassified";

export type DeliverableFindingCategory =
  | "package"
  | "openXml"
  | "relationship"
  | "structure"
  | "workflow"
  | "delta"
  | "render"
  | "artifact";

export type DeliverableCheckStatus =
  | "completed"
  | "skippedPrerequisiteFailed"
  | "unavailableEvidence";

export type DeliverablePackageChangeKind =
  | "entryAdded"
  | "entryRemoved"
  | "entryModified"
  | "relationshipAdded"
  | "relationshipRemoved"
  | "relationshipModified";

export type DeliverableArtifactRole =
  | "html"
  | "pdf"
  | "pageMap"
  | "pageImage"
  | "renderReport"
  | "other";

export type DeliverableArtifactAvailability = "available" | "unavailable";

/** Camel-case enum values used by the deliverable report's semantic summary. */
export type DeliverableSemanticChangeFamily =
  | "text"
  | "blockStructure"
  | "runFormatting"
  | "paragraphFormatting"
  | "style"
  | "numbering"
  | "list"
  | "table"
  | "tableRow"
  | "tableCell"
  | "tableSpan"
  | "tableWidth"
  | "tableStyle"
  | "section"
  | "pageSetup"
  | "header"
  | "footer"
  | "field"
  | "footnote"
  | "endnote"
  | "comment"
  | "hyperlink"
  | "bookmark"
  | "contentControl"
  | "image"
  | "media"
  | "relationship"
  | "revision"
  | "annotation"
  | "opaquePackagePart";

export interface DeliverablePackageIdentity {
  packageKind: string;
  manifestValid: boolean;
  rawPackageBytesDigest: VerificationDigest;
  orderedOpcContentDigest: VerificationDigest | null;
  normalizedSemanticDigest: VerificationDigest | null;
}

export interface DeliverableCheckResult {
  check: string;
  status: DeliverableCheckStatus;
  findingCount: number;
  diagnostic: string | null;
}

export interface DeliverableFinding {
  findingId: string;
  code: string;
  category: DeliverableFindingCategory;
  severity: VerificationFindingSeverity;
  disposition: DeliverableFindingDisposition;
  blocksDelivery: boolean;
  message: string;
  owningPartUri: string;
  location: ChangeLocation | null;
  anchorId: string | null;
  scope: string | null;
  xPath: string | null;
  remediation: string;
}

export interface DeliverablePackageChange {
  changeId: string;
  kind: DeliverablePackageChangeKind;
  location: ChangeLocation;
  beforeDigest: VerificationDigest | null;
  afterDigest: VerificationDigest | null;
  beforeValue: string | null;
  afterValue: string | null;
}

export interface DeliverableSemanticChange {
  changeId: string;
  fingerprint: string;
  operation: SemanticChangeOperation;
  family: DeliverableSemanticChangeFamily;
  partUri: string;
  path: string;
  leftAnchor: string | null;
  rightAnchor: string | null;
}

export interface DeliverableSemanticDelta {
  schema: "docxodus.semantic-changes";
  schemaVersion: 1;
  changeCount: number;
  canonicalDigest: VerificationDigest;
  changes: readonly DeliverableSemanticChange[];
}

export interface DeliverableArtifactMetadata {
  artifactId: string;
  role: DeliverableArtifactRole;
  mediaType: string;
  availability: DeliverableArtifactAvailability;
  byteLength: number | null;
  digest: VerificationDigest | null;
  unavailableReason: string | null;
  pageCount: number | null;
  rendererFingerprint: string | null;
  sourcePackageDigest: VerificationDigest | null;
  pageMapDigest: VerificationDigest | null;
  renderDiagnosticCount: number;
}

/** Canonical schema-v1 report returned by every default verification transport. */
export interface DeliverableVerificationResult {
  schema: "https://docxodus.dev/schemas/verification/deliverable-verification/v1";
  schemaVersion: 1;
  mode: DeliverableVerificationMode;
  decision: DeliverableVerificationDecision;
  analysisCompleted: boolean;
  baselineCompared: boolean;
  baselinePackage: DeliverablePackageIdentity | null;
  deliverablePackage: DeliverablePackageIdentity;
  checks: readonly DeliverableCheckResult[];
  findings: readonly DeliverableFinding[];
  resolvedFindings: readonly DeliverableFinding[];
  semanticDelta: DeliverableSemanticDelta | null;
  packageChanges: readonly DeliverablePackageChange[];
  companionArtifacts: readonly DeliverableArtifactMetadata[];
}

/** How a revision in a redline relates to the selected baseline. */
export type RedlineRevisionDisposition =
  | "preExisting"
  | "intendedFinalPreExisting"
  | "generated"
  | "conflicted";

/** Which of the two proof paths a result or finding belongs to. */
export type RedlineProofDirection = "acceptToFinal" | "rejectToBaseline";

/** How a package entry differs from a path's expected document. */
export type RedlinePackageDivergenceKind = "added" | "removed" | "modified";

/** Fail-closed resolution status of one native revision — the session registry's vocabulary. */
export type RedlineRevisionResolutionStatus = RevisionResolutionStatus;

/** Coarse family of one native revision. */
export type RedlineRevisionFamily =
  | "contentInsert"
  | "contentDelete"
  | "move"
  | "paragraphMark"
  | "rowInsert"
  | "rowDelete"
  | "cellInsert"
  | "cellDelete"
  | "cellMerge"
  | "contentControlInsert"
  | "contentControlDelete"
  | "numberingPropertiesInsert"
  | "numberingChange"
  | "propertiesChange"
  | "unsupported";

/** Why a revision could not be resolved — the session registry's shape. */
export type RedlineRevisionDiagnostic = RevisionDiagnostic;

/** Input or output package identity recorded by the proof. */
export interface RedlineProofPackageIdentity {
  rawPackageBytesDigest: VerificationDigest;
  orderedOpcContentDigest: VerificationDigest | null;
  normalizedWholePackageDigest: VerificationDigest | null;
}

/** A stable, part-qualified identity for one native Word revision. */
export interface RedlineRevisionIdentity {
  id: string;
  partUri: string;
  scope: string;
  type: string;
  family: RedlineRevisionFamily;
  constituentIds: readonly string[];
  constituentKeys: readonly string[];
  author: string;
  date: string | null;
  dateUtc: string | null;
  text: string;
  anchorId: string | null;
  affectedAnchorIds: readonly string[];
  resolutionStatus: RedlineRevisionResolutionStatus;
  diagnostic: RedlineRevisionDiagnostic | null;
}

/** Classification of a baseline/intended-final/redline revision identity triple. */
export interface RedlineRevisionClassification {
  disposition: RedlineRevisionDisposition;
  baseline: RedlineRevisionIdentity | null;
  intendedFinal: RedlineRevisionIdentity | null;
  redline: RedlineRevisionIdentity | null;
  reason: string;
}

/**
 * Modeled semantic comparison for one path. `available` and `changeCount` are explicit so an
 * empty modeled change set is never mistaken for complete package equality.
 */
export interface RedlineModeledSemanticComparison {
  available: boolean;
  equivalent: boolean | null;
  schema: string | null;
  changeCount: number | null;
  diagnostic: string | null;
}

/** One added, removed, or modified package entry on a proof path. */
export interface RedlinePackageDivergence {
  kind: RedlinePackageDivergenceKind;
  partUri: string;
  occurrence: number;
  anchorId: string | null;
  applicableRevisionIds: readonly string[];
  expectedRawDigest: VerificationDigest | null;
  actualRawDigest: VerificationDigest | null;
  expectedNormalizedDigest: VerificationDigest | null;
  actualNormalizedDigest: VerificationDigest | null;
  /** Whether the semantic change set reports a modeled change for this part. */
  hasModeledSemanticChange: boolean;
  /**
   * Conservatively true when the normalized difference may contain content outside the modeled
   * semantic projection. A modeled change in one part never proves every change in it was modeled.
   */
  unknownOrUnmodeled: boolean;
}

/** A structured, actionable proof finding. */
export interface RedlineProofFinding {
  code: string;
  severity: VerificationFindingSeverity;
  message: string;
  direction: RedlineProofDirection | null;
  location: ChangeLocation | null;
  anchorId: string | null;
  revisionIds: readonly string[];
  remediation: string | null;
}

/** Result of accepting or rejecting only the generated revision set. */
export interface RedlineProofPathResult {
  direction: RedlineProofDirection;
  completed: boolean;
  equivalent: boolean;
  requestedRevisionIds: readonly string[];
  resolvedRevisionIds: readonly string[];
  implicitlyResolvedRevisionIds: readonly string[];
  survivingPreExistingRevisions: readonly RedlineRevisionIdentity[];
  preExistingRevisionsPreserved: boolean;
  modeledSemantic: RedlineModeledSemanticComparison;
  normalizedWholePackageEquivalent: boolean;
  orderedOpcContentEquivalent: boolean;
  exactPackageBytesEquivalent: boolean;
  /**
   * Whether the complete bounded package delta was available. A false value never exposes a
   * potentially misleading prefix of `divergences`.
   */
  divergenceAnalysisCompleted: boolean;
  expectedPackage: RedlineProofPackageIdentity;
  actualPackage: RedlineProofPackageIdentity | null;
  firstDivergence: RedlinePackageDivergence | null;
  divergences: readonly RedlinePackageDivergence[];
  findings: readonly RedlineProofFinding[];
}

/**
 * Canonical schema-v1 proof that a redline's generated changes accept to the intended final and
 * reject to the baseline without consuming pre-existing review state.
 */
export interface RedlineReversibilityProof {
  schema: "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1";
  schemaVersion: 1;
  success: boolean;
  requireExactPackageBytes: boolean;
  baselinePackage: RedlineProofPackageIdentity;
  intendedFinalPackage: RedlineProofPackageIdentity;
  redlinePackage: RedlineProofPackageIdentity;
  revisionClassifications: readonly RedlineRevisionClassification[];
  acceptToFinal: RedlineProofPathResult | null;
  rejectToBaseline: RedlineProofPathResult | null;
  findings: readonly RedlineProofFinding[];
}

/** Effective #493 inspection limits applied while the manifest is being generated. */
export interface PackageManifestInspectionLimits {
  opcEntries: number;
  expandedOpcBytes: number;
  xmlPartBytes: number;
  opcUriCharacters: number;
  opcCompressionRatio: number;
}

/**
 * Internal WASM exports structure
 */
export interface DocxodusWasmExports {
  DocumentConverter: {
    GeneratePackageManifest: (bytes: Uint8Array) => string;
    VerifyDeliverable: (bytes: Uint8Array) => string;
    VerifyDeliverableWithBaseline: (
      baselineBytes: Uint8Array,
      bytes: Uint8Array
    ) => string;
    ProveRedlineReversibility: (
      baselineBytes: Uint8Array,
      intendedFinalBytes: Uint8Array,
      redlineBytes: Uint8Array
    ) => string;
    GeneratePackageManifestWithOptions: (
      bytes: Uint8Array,
      maxEntryCount: number,
      maxTotalUncompressedBytes: number,
      maxXmlPartBytes: number,
      maxCompressionRatio: number,
      maxUriLength: number,
    ) => string;
    ConvertDocxToHtml: (bytes: Uint8Array) => string;
    RenderBlockHtml: (
      bytes: Uint8Array,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean
    ) => string;
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
      documentLanguage: string | null,
      stampAnchors: boolean
    ) => string;
    GetAnnotations: (bytes: Uint8Array) => string;
    AddAnnotation: (bytes: Uint8Array, requestJson: string) => string;
    AddAnnotationWithTarget: (bytes: Uint8Array, requestJson: string) => string;
    RemoveAnnotation: (bytes: Uint8Array, annotationId: string) => string;
    HasAnnotations: (bytes: Uint8Array) => string;
    GetDocumentStructure: (bytes: Uint8Array) => string;
    GetDocumentMetadata: (bytes: Uint8Array) => string;
    ExportToOpenContract: (bytes: Uint8Array) => string;
    ConvertWmlToMarkdown: (bytes: Uint8Array, settingsJson: string) => string;
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
    // Incremental annotation overlay methods
    ProjectAnnotationsOntoHtml: (
      html: string,
      annotationSetJson: string,
      extAnnotCssClassPrefix: string,
      extAnnotLabelMode: number
    ) => string;
    AddAnnotationToHtml: (
      html: string,
      annotationJson: string,
      labelJson: string,
      extAnnotCssClassPrefix: string,
      extAnnotLabelMode: number
    ) => string;
    RemoveAnnotationFromHtml: (
      html: string,
      annotationId: string,
      extAnnotCssClassPrefix: string
    ) => string;
    GenerateAnnotationVisibilityCss: (
      hiddenLabelIdsJson: string,
      extAnnotCssClassPrefix: string
    ) => string;
    GenerateAnnotationCss: (
      labelsJson: string,
      extAnnotCssClassPrefix: string,
      extAnnotLabelMode: number
    ) => string;
  };
  DocumentComparer: {
    /**
     * Force the comparison code path hot by running a real comparison against
     * tiny in-memory seed documents. Returns "ok" or a JSON error object.
     * Idempotent — assemblies load only once.
     */
    Warmup: () => string;
    CompareDocuments: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      engine: number
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
      renderTrackedChanges: boolean,
      engine: number
    ) => string;
    CompareDocumentsToHtmlFull: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      detailThreshold: number,
      caseInsensitive: boolean,
      renderTrackedChanges: boolean,
      engine: number
    ) => string;
    CompareDocumentsWithOptions: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      detailThreshold: number,
      caseInsensitive: boolean,
      engine: number
    ) => Uint8Array;
    GetRevisionsJson: (comparedDocBytes: Uint8Array) => string;
    GetRevisionsJsonWithOptions: (
      comparedDocBytes: Uint8Array,
      detectMoves: boolean,
      moveSimilarityThreshold: number,
      moveMinimumWordCount: number,
      caseInsensitive: boolean
    ) => string;
    CompareDocumentsWithLog: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      detailThreshold: number,
      caseInsensitive: boolean
    ) => string;
    CompareDocumentsToHtmlWithLog: (
      originalBytes: Uint8Array,
      modifiedBytes: Uint8Array,
      authorName: string,
      detailThreshold: number,
      caseInsensitive: boolean,
      renderTrackedChanges: boolean
    ) => string;
  };
  DocxDiffBridge: {
    /** Redlined DOCX bytes (native markup), or empty array on error. */
    Compare: (
      leftBytes: Uint8Array,
      rightBytes: Uint8Array,
      settingsJson: string
    ) => Uint8Array;
    /** `{"revisions":[…]}` JSON, or a JSON error object. */
    GetRevisionsJson: (
      leftBytes: Uint8Array,
      rightBytes: Uint8Array,
      settingsJson: string
    ) => string;
    /** Edit-script JSON (diff-as-data), or a JSON error object. */
    GetEditScriptJson: (
      leftBytes: Uint8Array,
      rightBytes: Uint8Array,
      settingsJson: string
    ) => string;
    /** Canonical `docxodus.semantic-changes` JSON, or a JSON error object. */
    GetSemanticChangesJson: (
      leftBytes: Uint8Array,
      rightBytes: Uint8Array,
      settingsJson: string
    ) => string;
    /** Accept all tracked revisions in a redlined DOCX → "right"-side bytes, or empty array on error. */
    AcceptRevisions: (bytes: Uint8Array) => Uint8Array;
    /** Reject all tracked revisions in a redlined DOCX → "left"-side bytes, or empty array on error. */
    RejectRevisions: (bytes: Uint8Array) => Uint8Array;
    /** Consolidated redlined DOCX bytes (native markup), or empty array on error. */
    Consolidate: (
      baseBytes: Uint8Array,
      reviewersJson: string,
      settingsJson: string
    ) => Uint8Array;
    /** `{"revisions":[…]}` JSON for the consolidation, or a JSON error object. */
    GetConsolidatedRevisionsJson: (
      baseBytes: Uint8Array,
      reviewersJson: string,
      settingsJson: string
    ) => string;
    /** Consolidated edit-script JSON (diff-as-data), or a JSON error object. */
    GetConsolidatedEditScriptJson: (
      baseBytes: Uint8Array,
      reviewersJson: string,
      settingsJson: string
    ) => string;
    /** `{"conflicts":[…]}` JSON for the consolidation, or a JSON error object. */
    GetConflictsJson: (
      baseBytes: Uint8Array,
      reviewersJson: string,
      settingsJson: string
    ) => string;
  };
  DocxSessionBridge: {
    OpenSession: (bytes: Uint8Array, settingsJson: string) => number;
    OpenPreviewSession?: (liveHandle: number) => number;
    CloseSession: (handle: number) => void;
    CreateBlankDocx: () => Uint8Array;
    Project: (handle: number) => string;
    GetVersion: (handle: number) => string;
    RegisterPageMap: (handle: number, pageMapJson: string, expectedRendererFingerprint: string) => string;
    GetPageMapStatus: (handle: number, requestJson: string) => string;
    GetPageCitation: (handle: number, anchorId: string, requestJson: string) => string;
    GetPackageContentHash?: (handle: number) => string;
    GetPackageManifest: (handle: number) => string;
    RenderPreviewHtml?: (handle: number) => string;
    RenderPreviewBlockHtml?: (handle: number, anchorId: string) => string;
    CheckPreconditions: (handle: number, preconditionsJson: string) => string;
    BeginTransaction: (handle: number) => number;
    CommitTransaction: (transactionHandle: number) => void;
    RollbackTransaction: (transactionHandle: number) => void;
    ProjectAnchor: (handle: number, anchorId: string, depth: number) => string;
    ProjectAnchorWithCitations: (
      handle: number,
      anchorId: string,
      depth: number,
      requestJson: string,
    ) => string;
    /** Ordered top-level render units per scope container (JSON {@link RenderPlan}) —
     *  what the editor's incremental reconciler diffs its DOM against. Optional:
     *  absent on older WASM bundles. */
    ListBlocks?: (handle: number) => string;
    ListRenderedBlocks?: (handle: number, renderTrackedChanges: boolean) => string;
    /** Citation-ordered footnotes/endnotes (JSON {@link NoteListEntry}[]) — the
     *  id↔ordinal authority for renumbering rendered note chrome. Optional:
     *  absent on older WASM bundles. */
    ListNotes?: (handle: number, endnotes: boolean) => string;
    /** The anchor index alone as `{"anchorIndex":{…}}` — the editor's per-op
     *  anchor-map refresh without marshaling the whole markdown projection.
     *  Optional: absent on older WASM bundles. */
    ListAnchors?: (handle: number) => string;
    /** Batch block render: `anchorIdsJson` is a JSON string array; returns a JSON
     *  object mapping each id to its HTML element (null when unresolvable), or
     *  `{"error": …}` on total failure — parse and check for `error`. One throwaway
     *  doc + one converter run for the whole batch, with real sibling context and
     *  true list-marker numbers. Optional: absent on older WASM bundles. */
    RenderBlocksHtml?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean
    ) => string;
    RenderBlocksHtmlForReview?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean
    ) => string;
    RenderBlockHtml: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean
    ) => string;
    RenderBlockHtmlForReview?: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean
    ) => string;
    RenderHtml: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number
    ) => string;
    RenderHtmlForReview?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
      renderTrackedChanges: boolean
    ) => string;
    ReplaceText: (handle: number, anchor: string, md: string) => string;
    DeleteBlock: (handle: number, anchor: string) => string;
    MoveBlock: (handle: number, sourceAnchor: string, targetAnchor: string, pos: string) => string;
    /** JSON `{anchorId, before, after}[]` — the blocks a block may legally be moved next to and
     *  on which side, which is what a drag UI gates its drop targets on. The two sides are
     *  reported separately because a cross-block range or a section break between the blocks can
     *  make one legal and the other not. Optional: absent on older WASM bundles. */
    ValidMoveTargets?: (handle: number, sourceAnchor: string) => string;
    DeleteRange: (handle: number, fromAnchorId: string, toAnchorIdExclusive: string) => string;
    DeleteSection: (handle: number, headingAnchorId: string) => string;
    InsertParagraph: (handle: number, anchor: string, pos: string, md: string) => string;
    SplitParagraph: (handle: number, anchor: string, offset: number) => string;
    MergeParagraphs: (handle: number, first: string, second: string) => string;
    InsertHorizontalRule: (handle: number, anchor: string, pos: string, ruleJson: string) => string;
    InsertTable: (handle: number, anchor: string, pos: string, rows: number, cols: number, optionsJson: string) => string;
    GetTableMetadata: (handle: number, tableAnchor: string) => string;
    ResolveTableCellAnchor: (handle: number, cellAnchor: string) => string;
    ResolveTableCellCoordinate: (handle: number, tableAnchor: string, rowIndex: number, columnIndex: number) => string;
    InsertTableRow: (handle: number, cellAnchor: string, pos: string) => string;
    InsertTableColumn: (handle: number, cellAnchor: string, pos: string) => string;
    DeleteTableRow: (handle: number, cellAnchor: string) => string;
    DeleteTableColumn: (handle: number, cellAnchor: string) => string;
    MergeCells: (handle: number, cellAnchor: string, rowSpan: number, colSpan: number, content: string) => string;
    UnmergeCells: (handle: number, cellAnchor: string) => string;
    SetColumnWidths: (handle: number, cellAnchor: string, widthsJson: string) => string;
    SetTableBorders: (handle: number, cellAnchor: string, specJson: string) => string;
    SetCellShading: (handle: number, cellAnchor: string, fill: string, scope: string) => string;
    SetRepeatHeaderRow: (handle: number, cellAnchor: string, repeat: boolean) => string;
    SetTableRowOptions: (handle: number, cellAnchor: string, repeatHeader: boolean | null,
      allowBreakAcrossPages: boolean | null, heightTwips: number | null, heightRule: string) => string;
    SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    InsertPageNumberField: (handle: number, anchor: string, field: string, format: string) => string;
    SetPageNumbering: (handle: number, anchor: string, opJson: string) => string;
    ClearPageNumbering: (handle: number, anchor: string) => string;
    EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
    InsertFootnote: (handle: number, anchor: string, characterOffset: number, markdown: string) => string;
    InsertEndnote: (handle: number, anchor: string, characterOffset: number, markdown: string) => string;
    // Native Word comment authoring (issue #300). spanJson: "" = whole block, else
    // {"start":n,"length":n}; initials/date: "" = absent (date is ISO-8601).
    AddComment: (
      handle: number,
      anchor: string,
      spanJson: string,
      author: string,
      initials: string,
      date: string,
      markdown: string,
    ) => string;
    AddCommentToRevision: (
      handle: number,
      revisionId: string,
      author: string,
      initials: string,
      date: string,
      markdown: string,
    ) => string;
    AddCommentReply: (
      handle: number,
      parentCommentAnchor: string,
      author: string,
      initials: string,
      date: string,
      markdown: string,
    ) => string;
    UpdateComment: (handle: number, commentAnchor: string, markdown: string) => string;
    SetCommentResolved: (handle: number, commentAnchor: string, resolved: boolean) => string;
    RemoveComment: (handle: number, commentAnchor: string) => string;
    ListComments: (handle: number) => string;
    ListHyperlinks: (handle: number, scopes: number) => string;
    AddHyperlink: (handle: number, anchor: string, start: number, length: number, kind: string, target: string) => string;
    UpdateHyperlink: (handle: number, hyperlinkId: string, kind: string, target: string) => string;
    RemoveHyperlink: (handle: number, hyperlinkId: string) => string;
    GetImageCapabilities: () => string;
    ListImages: (handle: number, scopes: number) => string;
    InsertImage: (handle: number, anchor: string, characterOffset: number, imageBase64: string, optionsJson: string) => string;
    ReplaceImage: (handle: number, imageId: string, imageBase64: string) => string;
    SetImageDimensions: (handle: number, imageId: string, dimensionsJson: string) => string;
    SetImageMetadata: (handle: number, imageId: string, altText: string | null, title: string | null) => string;
    SetImageFloatingLayout: (handle: number, imageId: string, layoutJson: string) => string;
    RemoveImage: (handle: number, imageId: string) => string;
    ListContentControls: (handle: number, scopes: number) => string;
    FillContentControlText: (handle: number, anchorId: string, text: string, optionsJson: string) => string;
    FillContentControlRichText: (handle: number, anchorId: string, markdown: string, optionsJson: string) => string;
    SetContentControlChecked: (handle: number, anchorId: string, isChecked: boolean, optionsJson: string) => string;
    SetContentControlDate: (handle: number, anchorId: string, value: string, displayText: string | null, optionsJson: string) => string;
    SelectContentControlItem: (handle: number, anchorId: string, value: string, optionsJson: string) => string;
    FillContentControlPicture: (handle: number, anchorId: string, imageBase64: string, optionsJson: string) => string;
    AddRepeatingSectionItem: (handle: number, sectionAnchorId: string, afterItemAnchorId: string, optionsJson: string) => string;
    RemoveRepeatingSectionItem: (handle: number, itemAnchorId: string) => string;
    ListBookmarks: (handle: number, scopes: number) => string;
    AddBookmark: (handle: number, name: string, startAnchor: string, startOffset: number, endAnchor: string, endOffset: number) => string;
    RenameBookmark: (handle: number, name: string, newName: string) => string;
    MoveBookmark: (handle: number, name: string, startAnchor: string, startOffset: number, endAnchor: string, endOffset: number) => string;
    RemoveBookmark: (handle: number, name: string) => string;
    ListRevisions: (handle: number) => string;
    AcceptRevision: (handle: number, revisionId: string) => string;
    RejectRevision: (handle: number, revisionId: string) => string;
    AcceptAllRevisions: (handle: number) => string;
    RejectAllRevisions: (handle: number) => string;
    ApplyFormat: (handle: number, anchor: string, spanJson: string, opJson: string) => string;
    ApplyFormatBySubstring: (handle: number, anchor: string, substring: string, opJson: string) => string;
    SetParagraphStyle: (handle: number, anchor: string, styleId: string) => string;
    SetParagraphFormat: (handle: number, anchor: string, opJson: string) => string;
    SetListLevel: (handle: number, anchor: string, delta: number) => string;
    RemoveListMembership: (handle: number, anchor: string) => string;
    ApplyListFormat: (handle: number, anchor: string, kind: string) => string;
    ApplyListFormatRange: (handle: number, firstAnchor: string, lastAnchor: string, kind: string) => string;
    SetListStartOverride: (handle: number, anchor: string, value: number) => string;
    ClearListStartOverride: (handle: number, anchor: string) => string;
    ReplaceCellContent: (handle: number, anchor: string, md: string) => string;
    RawGetXml: (handle: number, anchor: string) => string;
    RawInsertXml: (handle: number, anchor: string, pos: string, xml: string) => string;
    RawReplaceXml: (handle: number, anchor: string, xml: string) => string;
    Grep: (handle: number, pattern: string, optionsJson: string) => string;
    GrepCrossBlock: (handle: number, pattern: string, optionsJson: string) => string;
    ReplaceTextRange: (handle: number, anchor: string, find: string, replace: string, optionsJson: string) => string;
    ReplaceTextAtSpan: (handle: number, anchor: string, spanStart: number, spanLength: number, replace: string) => string;
    ReplaceInner: (
      handle: number,
      matchText: string,
      anchor: string,
      spanStart: number,
      spanLength: number,
      newInner: string,
    ) => string;
    FindPlaceholders: (handle: number, kinds: number, scope: number, contextChars: number, boundary: number) => string;
    FindPlaceholdersWithCitations: (
      handle: number,
      kinds: number,
      scope: number,
      contextChars: number,
      boundary: number,
      requestJson: string,
    ) => string;
    GetEditSummary: (handle: number) => string;
    RemainingPlaceholders: (handle: number, kinds: number) => string;
    GetDiff: (handle: number, format: number) => string;
    GetSemanticChanges: (handle: number) => string;
    VerifyDeliverable: (handle: number) => string;
    FindByAnnotation: (handle: number, annotationId: string) => string;
    FindByAnnotationWithCitations: (handle: number, annotationId: string, requestJson: string) => string;
    FindByLabel: (handle: number, labelId: string) => string;
    FindByLabelWithCitations: (handle: number, labelId: string, requestJson: string) => string;
    FindByBookmark: (handle: number, bookmarkName: string) => string;
    FindByBookmarkWithCitations: (handle: number, bookmarkName: string, requestJson: string) => string;
    Exists: (handle: number, anchorId: string) => boolean;
    FindByText: (handle: number, needle: string, optionsJson: string) => string;
    FindAllByText: (handle: number, needle: string, optionsJson: string) => string;
    FindByRegex: (handle: number, pattern: string, regexOptions: number, optionsJson: string) => string;
    FindByKind: (handle: number, kind: string, scope: string) => string;
    FindByKindWithCitations: (handle: number, kind: string, scope: string, requestJson: string) => string;
    GetAnchorInfo: (handle: number, anchorId: string) => string;
    GetAnchorInfos: (handle: number, anchorIdsJson: string) => string;
    GetBlockMetadata: (handle: number, anchorId: string) => string;
    GetBlockMetadatas: (handle: number, anchorIdsJson: string) => string;
    GetListMembership: (handle: number, anchorId: string) => string;
    GetSectionInfo: (handle: number, anchorId: string) => string;
    ListStyles: (handle: number) => string;
    GetFormatting: (handle: number, anchorId: string) => string;
    ListInlineSpans: (handle: number, anchorId: string) => string;
    ListAnnotations: (handle: number) => string;
    // Session annotation write surface
    AddAnnotation: (
      handle: number,
      anchorId: string,
      spanJson: string,
      annotationJson: string
    ) => string;
    SessionRemoveAnnotation: (handle: number, annotationId: string) => string;
    UpdateAnnotation: (
      handle: number,
      annotationId: string,
      updateJson: string
    ) => string;
    MoveAnnotation: (
      handle: number,
      annotationId: string,
      newAnchorId: string,
      newSpanJson: string
    ) => string;
    Undo: (handle: number) => boolean;
    Redo: (handle: number) => boolean;
    SetTrackedChanges: (handle: number, mode: number) => void;
    SetRevisionAuthor: (handle: number, author: string) => void;
    Save: (handle: number) => Uint8Array;
    /** Save KEEPING the projector's `PtOpenXml:Unid` bookkeeping. Exists for the in-browser
     *  editor's remount, which re-renders these bytes and needs the anchors to survive the hop.
     *  Roughly 6x the document's size and read by no renderer — never hand it to a user; a
     *  save-to-disk wants {@link Save}. */
    SaveWithAnchorIds: (handle: number) => Uint8Array;
  };
}

// ─── DocxSession types ────────────────────────────────────────────────────

export type EditErrorCode =
  | "anchor_not_found"
  | "anchor_wrong_kind"
  | "anchors_not_adjacent"
  | "session_disposed"
  | "malformed_markdown"
  | "unsupported_markdown_syntax"
  | "table_insert_not_supported"
  | "footnote_ref_not_supported"
  | "comment_marker_not_supported"
  | "image_insert_not_supported"
  | "anchor_token_in_payload"
  | "offset_out_of_range"
  | "invalid_position"
  | "unknown_style"
  | "invalid_list_level"
  | "invalid_list_start_value"
  | "invalid_page_numbering"
  | "invalid_paragraph_format"
  | "invalid_table_styling"
  | "invalid_table_merge"
  | "table_anchor_migration_required"
  | "malformed_xml"
  | "disallowed_namespace"
  | "incompatible_element_type"
  | "validation_failed"
  | "nothing_to_undo"
  | "nothing_to_redo"
  | "duplicate_annotation_id"
  | "annotation_not_found"
  | "empty_annotation_span"
  | "empty_comment_span"
  | "revision_not_found"
  | "precondition_failed"
  | "invalid_batch_step"
  | "invalid_transaction"
  | "transaction_conflict"
  | "transaction_result_evicted"
  | "transaction_incomplete"
  | "hyperlink_not_found"
  | "bookmark_not_found"
  | "duplicate_bookmark_name"
  | "invalid_bookmark_name"
  | "invalid_hyperlink_target"
  | "missing_bookmark_target"
  | "bookmark_in_use"
  | "managed_bookmark"
  | "empty_hyperlink_span"
  | "unsupported_inline_boundary"
  | "revision_unsupported"
  | "revision_malformed"
  | "revision_ambiguous"
  | "tracked_operation_unsupported"
  | "unresolved_structural_revision"
  | "content_control_not_found"
  | "content_control_malformed"
  | "content_control_unsupported"
  | "content_control_locked"
  | "content_control_bound"
  | "content_control_wrong_type"
  | "invalid_content_control_value"
  | "content_control_placement_unsupported"
  | "content_control_nested_fill_unsupported"
  | "repeating_section_constraint"
  | "image_not_found"
  | "invalid_image_data"
  | "unsupported_image_format"
  | "image_too_large"
  | "invalid_image_dimensions"
  | "unsupported_image_markup"
  | "linked_image_read_only"
  | "invalid_image_layout"
  | "internal_error";

export interface AnchorRef {
  id: string;
  kind: string;
  scope: string;
  unid: string;
}

export interface EditError {
  code: EditErrorCode;
  message: string;
  anchorId?: string;
  precondition?: PreconditionFailure;
}

export interface PreconditionTarget {
  exists: boolean;
  anchorId?: string;
  kind?: string;
  scope?: string;
  contentHash?: string;
  visibleText?: string;
}

export interface PreconditionFailure {
  condition: string;
  expected: unknown;
  actual: unknown;
  currentVersion: number;
  currentTarget?: PreconditionTarget;
}

export interface TextRangePrecondition {
  start: number;
  length: number;
  text: string;
}

/** Optimistic guards checked immediately before a mutation. */
export interface MutationPreconditions {
  expectedVersion?: number;
  /** Optional explicit target; target-addressed methods infer their own anchor when omitted. */
  anchorId?: string;
  expectedContentHash?: string;
  expectedText?: string;
  expectedTextRange?: TextRangePrecondition;
  expectedKind?: string;
  expectedScope?: string;
  expectedMatchCount?: number;
}

export type MutationBatchMode = "atomic" | "best_effort";

/** One synchronous npm batch step. Atomic is the default execution mode. */
export interface MutationBatchStep {
  tool: string;
  action: string;
  mutation: () => EditResult | readonly EditResult[];
  /** Optional read-only validation: all run up front in atomic mode, per-step in best-effort. */
  preflight?: () => EditError | undefined;
}

/** A preview callback receives the isolated shadow session it must mutate/read. */
export interface MutationBatchPreviewStep {
  tool: string;
  action: string;
  mutation: (shadow: import("./session.js").DocxSession) => EditResult | readonly EditResult[];
  preflight?: (shadow: import("./session.js").DocxSession) => EditError | undefined;
}

export interface MutationBatchPreviewOptions {
  html?: "none" | "scoped" | "full";
  /** Required for scoped HTML. */
  htmlAnchorId?: string;
}

export interface MutationBatchStepResult {
  index: number;
  tool: string;
  action: string;
  success: boolean;
  rolledBack: boolean;
  results: readonly EditResult[];
}

export interface MutationBatchFailure {
  index: number;
  tool: string;
  action: string;
  error: EditError;
  rolledBack: boolean;
}

export interface MutationBatchChangeSet<T> {
  added: readonly T[];
  removed: readonly T[];
  modified: readonly T[];
}

export interface MutationBatchResult {
  mode: MutationBatchMode;
  status: "ok" | "failed" | "partial";
  preview: boolean;
  success: boolean;
  rolledBack: boolean;
  baseVersion: number;
  resultVersion: number;
  /**
   * Canonical SHA-256 of this result package, or `null` when it could not be computed.
   * Exact replay equality is guaranteed only for deterministic batches, so consult
   * {@link MutationBatchResult.warnings} before asserting on it — and note that `null`
   * never equals `null` for the purposes of a replay assertion: an absent hash proves
   * nothing and must be handled explicitly rather than compared.
   */
  packageHash: string | null;
  steps: readonly MutationBatchStepResult[];
  failure?: MutationBatchFailure;
  revisionChanges: MutationBatchChangeSet<RevisionListEntry>;
  commentChanges: MutationBatchChangeSet<CommentListEntry>;
  annotationChanges: MutationBatchChangeSet<DocumentAnnotation>;
  warnings: readonly string[];
  /** Shadow-only preview HTML when requested; null otherwise. */
  html: string | null;
}

export interface MarkdownPatch {
  scopeAnchorId: string;
  markdown: string;
}

export interface EditResult {
  success: boolean;
  error?: EditError;
  created: AnchorRef[];
  removed: AnchorRef[];
  modified: AnchorRef[];
  /** Deterministic structural identity map for table shape mutations. */
  tableAnchors?: TableAnchorMapping;
  patch?: MarkdownPatch;
  /** Set by the annotation ops (addAnnotation/removeAnnotation/updateAnnotation/
   * moveAnnotation) with the affected annotation id; absent for every other op. */
  annotationId?: string;
  hyperlinkId?: string;
  bookmarkName?: string;
  imageId?: string;
}

export type HyperlinkKind = "external" | "internal";

export interface HyperlinkInfo {
  id: string;
  kind: HyperlinkKind;
  owningPartUri: string;
  scope: string;
  anchorId: string;
  span: CharSpan;
  text: string;
  target?: string;
  relationshipId?: string;
  relationshipIsExternal?: boolean;
  isBroken: boolean;
}

export type ImageBinaryFormat = "unknown" | "png" | "jpeg" | "gif" | "bmp" | "tiff" | "webp";
export type ImageMarkupKind = "modern_drawing" | "legacy_vml" | "unsupported_drawing";
export type ImagePlacement = "inline" | "floating";
export type ImageWrapMode = "none" | "square" | "tight" | "through" | "top_and_bottom" | "unknown";
export type ImageWrapSide = "both_sides" | "left" | "right" | "largest" | "unknown";
export type ImageHorizontalReference = "page" | "margin" | "column" | "character" | "unknown";
export type ImageVerticalReference = "page" | "margin" | "paragraph" | "line" | "unknown";
export type ImageHorizontalAlignment = "left" | "center" | "right" | "inside" | "outside" | "unknown";
export type ImageVerticalAlignment = "top" | "center" | "bottom" | "inside" | "outside" | "unknown";

export interface FloatingImageLayout {
  horizontalRelativeFrom?: ImageHorizontalReference;
  horizontalOffsetEmu?: number | null;
  horizontalAlignment?: ImageHorizontalAlignment | null;
  verticalRelativeFrom?: ImageVerticalReference;
  verticalOffsetEmu?: number | null;
  verticalAlignment?: ImageVerticalAlignment | null;
  wrapMode?: ImageWrapMode;
  wrapSide?: ImageWrapSide;
  distanceTopEmu?: number;
  distanceBottomEmu?: number;
  distanceLeftEmu?: number;
  distanceRightEmu?: number;
  relativeHeight?: number;
  behindDocument?: boolean;
  locked?: boolean;
  layoutInCell?: boolean;
  allowOverlap?: boolean;
  rawHorizontalReference?: string;
  rawVerticalReference?: string;
  rawHorizontalPosition?: string;
  rawVerticalPosition?: string;
  rawWrapMode?: string;
  rawWrapSide?: string;
  rawRelativeSizeHorizontal?: string;
  rawRelativeSizeVertical?: string;
  rawFlagTokens?: Record<string, string>;
}

export interface ImageInsertOptions {
  placement?: ImagePlacement;
  widthPoints?: number;
  heightPoints?: number;
  preserveAspect?: boolean;
  altText?: string | null;
  title?: string | null;
  floatingLayout?: FloatingImageLayout;
}

export interface ImageDimensions {
  widthPoints?: number;
  heightPoints?: number;
  preserveAspect?: boolean;
}

export interface ImageOccurrence {
  id: string; markupKind: ImageMarkupKind; placement?: ImagePlacement; canMutate: boolean;
  unsupportedReason?: string; owningPartUri: string; scope: string; anchorId: string; span: CharSpan;
  relationshipId?: string; targetPartUri?: string; linkedRelationshipId?: string; linkedTarget?: string;
  isEmbedded: boolean; isLinked: boolean; isBroken: boolean; mediaFileName?: string; contentType?: string;
  format: ImageBinaryFormat; contentTypeMatchesBytes?: boolean; intrinsicWidthPixels?: number;
  intrinsicHeightPixels?: number; renderedWidthPoints?: number; renderedHeightPoints?: number;
  altText?: string; title?: string; floatingLayout?: FloatingImageLayout; floatingLayoutSupported: boolean;
}

export interface ImageFormatCapability {
  format: ImageBinaryFormat; contentType: string; canInspect: boolean; canInsert: boolean;
  canReplace: boolean; limitation?: string;
}

export interface ImageCapabilities {
  schemaVersion: number; runtime: string; formats: ImageFormatCapability[]; operations: string[];
  mutableWrapModes: ImageWrapMode[]; horizontalReferences: ImageHorizontalReference[];
  verticalReferences: ImageVerticalReference[]; maxInputBytes: number; maxRenderedPoints: number;
  defaultDpi: number; usesHeaderParsingOnly: boolean; acceptsBinaryBytes: boolean;
  supportsNetworkFetch: boolean; supportsFileIo: boolean;
}

export type ContentControlType = "plain_text" | "rich_text" | "checkbox" | "date"
  | "drop_down_list" | "combo_box" | "picture" | "repeating_section"
  | "repeating_section_item" | "unsupported";
export type ContentControlPlacement = "inline" | "block" | "row" | "cell" | "unknown";
export type ContentControlBindingPolicy = "preserve" | "detach_target";

export interface ContentControlFillOptions {
  bindingPolicy?: ContentControlBindingPolicy;
}

export interface ContentControlBindingInfo {
  storeItemId?: string;
  xpath?: string;
  prefixMappings?: string;
}

export interface ContentControlInfo {
  anchorId: string;
  type: ContentControlType;
  placement: ContentControlPlacement;
  nativeId?: string;
  tag?: string;
  alias?: string;
  lock?: string;
  isShowingPlaceholder: boolean;
  isBound: boolean;
  binding?: ContentControlBindingInfo;
  owningPartUri: string;
  scope: string;
  parentAnchorId?: string;
  depth: number;
  hasValidNativeId: boolean;
  hasDuplicateNativeId: boolean;
  canMutate: boolean;
  canDetachTargetBinding: boolean;
  unsupportedReason?: string;
  text: string;
  itemValues: string[];
}

export interface DocumentRange {
  startAnchorId: string;
  startOffset: number;
  endAnchorId: string;
  endOffset: number;
}

export interface BookmarkRangeSegment {
  owningPartUri: string;
  scope: string;
  anchorId: string;
  span: CharSpan;
  text: string;
}

export interface BookmarkInfo {
  name: string;
  bookmarkId: string;
  startPartUri: string;
  startScope: string;
  endPartUri?: string;
  endScope?: string;
  range?: DocumentRange;
  segments: BookmarkRangeSegment[];
  text: string;
  isPaired: boolean;
  isManaged: boolean;
  isValid: boolean;
  validationError?: string;
}

/**
 * One native Word comment, in comments-part order — see {@link DocxSession.listComments}.
 * `anchorId` addresses the definition (kind `cmt`) for updateComment/removeComment;
 * `date` is the raw `w:date` attribute string; `text` is the flattened body.
 * `parentAnchorId` and `resolved` are present when the comment has a
 * `commentsExtended.xml` entry; their absence distinguishes a legacy/flat comment from an
 * explicitly reopened one. The numeric `w:id` is deliberately not surfaced — comments are
 * addressed by anchor everywhere.
 */
export interface CommentListEntry {
  anchorId: string;
  author: string;
  initials?: string;
  date?: string;
  text: string;
  parentAnchorId?: string;
  resolved?: boolean;
}

/** Revision kind in a markup-native revision listing. A `move` entry is a linked
 * move pair — both sides resolve together. */
export type SessionRevisionType = "insert" | "delete" | "move" | "format" | "structure";
export type RevisionFamily =
  | "content_insert" | "content_delete" | "move" | "paragraph_mark"
  | "row_insert" | "row_delete" | "cell_insert" | "cell_delete" | "cell_merge"
  | "content_control_insert" | "content_control_delete"
  | "numbering_properties_insert" | "numbering_change" | "properties_change"
  | "unsupported";
export type RevisionResolutionStatus = "supported" | "unsupported" | "malformed" | "ambiguous";

export interface RevisionDiagnostic {
  code: string;
  message: string;
}

/**
 * One part-qualified atomic revision from the live registry. `id` is an opaque,
 * deterministic `rev2-…` identity; `constituentIds` exposes the native Word ids.
 * `family` identifies the exact operation, while `type` is its coarse display class.
 * Unsafe native topology remains listed through `resolutionStatus` and `diagnostic`.
 */
export interface RevisionListEntry {
  id: string;
  type: SessionRevisionType;
  family: RevisionFamily;
  constituentIds: string[];
  /**
   * QName-qualified native carrier identities. Unlike `constituentIds`, these
   * distinguish revision roles which legally use the same numeric `w:id` value.
   */
  constituentKeys: string[];
  author: string;
  date?: string;
  /** The `w16du:dateUtc` timestamp, when the markup carries one. */
  dateUtc?: string;
  text: string;
  partUri: string;
  scope: string;
  anchorId?: string;
  affectedAnchors: AnchorRef[];
  resolutionStatus: RevisionResolutionStatus;
  diagnostic?: RevisionDiagnostic;
}

export interface CharSpan {
  start: number;
  length: number;
}

export interface FormatOp {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  code?: boolean;
  color?: string;
  runStyle?: string;
  /** Vertical alignment: "superscript" | "subscript" | "" (clear). Omit to leave unchanged. */
  vertAlign?: string;
  /**
   * Font size in **points** (maps to `w:sz`/`w:szCs`, stored as half-points). Omit to leave
   * unchanged; a value &lt;= 0 clears the explicit size. Fractional points round to a half-point.
   */
  fontSizePts?: number;
  /**
   * Run font family (maps to `w:rFonts` ascii/hAnsi/cs). Omit to leave unchanged; `""` clears
   * the explicit font so the run inherits the style/default. Lets a run match a serif filing.
   */
  fontFamily?: string;
}

/** One edge of a paragraph border (`w:pBdr` top/bottom) — drives S-1 horizontal rules. */
export interface ParagraphBorderEdge {
  /** Border line style (`w:val`): "single","double","thick","dotted","dashed",… Default "single". */
  style?: string;
  /** Weight in eighths of a point (`w:sz`). Default 6 (≈0.75pt); a heavy rule ≈ 18–24. */
  size?: number;
  /** Hex color without '#', or "auto". Default "auto". */
  color?: string;
  /** Padding between border and text, in points (`w:space`). Default 1. */
  space?: number;
}

/** List format for `DocxSession.applyListFormat` / `applyListFormatRange`. The plain numbered
 * formats render `1.` / `a.` / `i.` level text; the `*Parenthesis` variants render `(1)` /
 * `(a)` / `(i)` — same `w:numFmt`, different `w:lvlText` (the legal-drafting presets). */
export type ListFormat =
  | "none"
  | "bullet"
  | "decimal"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman"
  | "decimalParenthesis"
  | "lowerLetterParenthesis"
  | "upperLetterParenthesis"
  | "lowerRomanParenthesis"
  | "upperRomanParenthesis";

/** How `ParagraphFormatOp.lineSpacing` is interpreted (`w:spacing/@w:lineRule`): under `"auto"`
 * the value is in 240ths of a line (240 = single, 360 = 1.5×, 480 = double); under
 * `"exact"`/`"atLeast"` it is a height in twips (20 twips = 1pt). */
export type LineSpacingRule = "auto" | "exact" | "atLeast";

/** Paragraph-level formatting for `DocxSession.setParagraphFormat`. Omit a field to leave it unchanged. */
export interface ParagraphFormatOp {
  /** Paragraph alignment. */
  alignment?: "left" | "center" | "right" | "justify";
  /** Adjust the left indent by this many twips (1440 = 1 inch); clamped at 0. */
  indentDelta?: number;
  /** First-line indent in twips (`w:ind/@w:firstLine`; 1440 = 1 inch). Absolute, not a delta;
   * 0 writes an explicit "none". Mutually exclusive with `hangingIndent` (Word's `w:ind` holds
   * one or the other) — setting this removes any hanging indent, and an op carrying both is
   * rejected with `invalid_paragraph_format`. Negatives are invalid. */
  firstLineIndent?: number;
  /** Hanging indent in twips (`w:ind/@w:hanging`; 1440 = 1 inch) — every line EXCEPT the first
   * starts this far right of the paragraph's left edge. Mutually exclusive with
   * `firstLineIndent`; setting this removes any first-line indent. Negatives are invalid. */
  hangingIndent?: number;
  /** Space above the paragraph in twips (`w:spacing/@w:before`; 20 twips = 1pt, so 240 = 12pt).
   * Absolute; negatives are invalid. */
  spacingBefore?: number;
  /** Space below the paragraph in twips (`w:spacing/@w:after`; 20 twips = 1pt). Absolute;
   * negatives are invalid. */
  spacingAfter?: number;
  /** Line spacing (`w:spacing/@w:line`). Units depend on `lineSpacingRule`: 240ths of a line
   * under `"auto"` (the default when the rule is omitted), twips under `"exact"`/`"atLeast"`.
   * Negatives are invalid. */
  lineSpacing?: number;
  /** How `lineSpacing` is interpreted. Only meaningful alongside `lineSpacing` — set without
   * it, the op is rejected with `invalid_paragraph_format`. */
  lineSpacingRule?: LineSpacingRule;
  /** Page-break-before: true to add, false to remove. */
  pageBreakBefore?: boolean;
  /** Top paragraph border (`w:pBdr/w:top`). Omit to leave unchanged. */
  topBorder?: ParagraphBorderEdge;
  /** Bottom paragraph border (`w:pBdr/w:bottom`) — what an S-1 horizontal rule is. Omit to leave unchanged. */
  bottomBorder?: ParagraphBorderEdge;
  /** Remove all paragraph borders before applying any top/bottom border in this op. */
  clearBorders?: boolean;
}

/**
 * Which header/footer story `DocxSession.setHeaderText`/`setFooterText` targets, mapping to the
 * OOXML `w:type`: `"default"` (all pages), `"first"` (first-page-only; sets `w:titlePg`),
 * `"even"` (even pages; sets `w:evenAndOddHeaders`).
 */
export type HeaderFooterKind = "default" | "first" | "even";

/** Which page-number field `DocxSession.insertPageNumberField` emits: `"currentPage"` → PAGE,
 * `"totalPages"` → NUMPAGES. */
export type PageNumberField = "currentPage" | "totalPages";

/**
 * Section-level page-numbering setup for {@link DocxSession.setPageNumbering} — the `w:pgNumType`
 * element, which is what Word's *Format Page Numbers…* dialog writes. Each field is tri-state: an
 * omitted field leaves that attribute exactly as it is. Use
 * {@link DocxSession.clearPageNumbering} to remove them.
 *
 * This governs how a **plain** PAGE field renders anywhere in the section, which is the normal way
 * to number pages: set the section once, insert unswitched fields. It is distinct from the
 * per-field `\*` switch {@link DocxSession.insertPageNumberField} can stamp, which overrides the
 * section for that one field.
 */
export interface PageNumberingOp {
  /** The page number this section starts at (`w:start`) — e.g. `1` to restart at a section break.
   *  Omitted leaves it unchanged; absent means the section continues the previous one. */
  start?: number;
  /** This section's page-number format (`w:fmt`) — e.g. `"lowerRoman"` for `i, ii, iii` front
   *  matter. Omitted leaves it unchanged; absent means Word's default `1, 2, 3`. `"bullet"` is
   *  rejected — pages cannot be bulleted. */
  format?: NumberFormat;
}

/** Options for `DocxSession.insertTable`. */
export interface TableInsertOptions {
  /** Emit an invisible layout table (explicit "none" borders) — the S-1 multi-column blocks. */
  borderless?: boolean;
  /** Row-major markdown for each cell (row 0 left→right, then row 1, …). Short/omitted ⇒ empty cells. */
  cellContents?: string[];
  /** Alignment applied to every cell paragraph (S-1 columns are centered). */
  cellAlignment?: "left" | "center" | "right" | "justify";
  /** Per-column widths in twips (one per column, left→right). Omit for equal columns; a list
   *  whose length != the column count is rejected. Drives unequal layouts like the S-1's
   *  wide-left / narrow-right filing-header row. */
  columnWidths?: number[];
}

export type TableVerticalMergeRole = "none" | "restart" | "continue";
export type TableAnchorEntityKind = "table" | "row" | "column" | "cell";

export interface TableCellMetadata {
  anchor: AnchorRef;
  tableAnchorId: string;
  rowAnchorId: string;
  rowIndex: number;
  columnIndex: number;
  rowSpan: number;
  columnSpan: number;
  verticalMerge: TableVerticalMergeRole;
  /** Direct cell paragraphs only; nested-table paragraphs belong to their own cells. */
  paragraphAnchors: AnchorRef[];
}

export interface TableRowMetadata {
  anchor: AnchorRef;
  tableAnchorId: string;
  rowIndex: number;
  gridBefore: number;
  gridAfter: number;
  cells: TableCellMetadata[];
}

export interface TableColumnMetadata {
  anchor: AnchorRef;
  tableAnchorId: string;
  columnIndex: number;
  widthTwips: number;
  /** True when an absent/underspecified tblGrid required a read-only coordinate identity. */
  isVirtual: boolean;
  cellAnchorIds: string[];
}

export interface TableMetadata {
  anchor: AnchorRef;
  columns: TableColumnMetadata[];
  rows: TableRowMetadata[];
}

export interface TableMetadataResult {
  success: boolean;
  error?: EditError;
  metadata?: TableMetadata;
}

export interface TableCellResolutionResult {
  success: boolean;
  error?: EditError;
  cell?: TableCellMetadata;
}

export interface TableAnchorLocation {
  anchor: AnchorRef;
  entityKind: TableAnchorEntityKind;
  rowIndex?: number;
  columnIndex?: number;
  rowSpan?: number;
  columnSpan?: number;
  isVirtual?: boolean;
}

export interface TableAnchorMapping {
  retained: { before: TableAnchorLocation; after: TableAnchorLocation }[];
  added: TableAnchorLocation[];
  invalidated: TableAnchorLocation[];
}

export type TableRowHeightRule = "auto" | "atLeast" | "exact";

export interface TableRowOptions {
  repeatHeader?: boolean;
  allowBreakAcrossPages?: boolean;
  /** Zero removes an explicit height. */
  heightTwips?: number;
  heightRule?: TableRowHeightRule;
}

/** Which table edges `DocxSession.setTableBorders` targets: `"outside"` = top/left/bottom/right,
 * `"inside"` = the inner grid lines (`w:insideH`/`w:insideV`), `"all"` = both. */
export type TableBorderScope = "all" | "outside" | "inside";

/** Shading granularity for `DocxSession.setCellShading`: the one cell the anchor sits in, or
 * every cell of its row (header-row banding). */
export type TableShadingScope = "cell" | "row";

/** What `DocxSession.mergeCells` does with the content of the cells a merge absorbs:
 * `"append"` (default) moves their non-empty blocks into the surviving cell — lossless;
 * `"discard"` drops them; `"reject"` fails the merge (`invalid_table_merge`) when any absorbed
 * cell is non-empty. */
export type TableMergeContent = "append" | "discard" | "reject";

/** Border specification for `DocxSession.setTableBorders`. Written as explicit `w:tblBorders`
 * edges (overriding style-inherited borders); edges outside `scope` are left untouched. */
export interface TableBorderSpec {
  /** Which edges to write. Default `"all"`. */
  scope?: TableBorderScope;
  /** Border line style (`w:val`): `"single"`, `"double"`, `"thick"`, `"dotted"`, `"dashed"`, … —
   * or `"none"` to remove the targeted edges. Default `"single"`. */
  style?: string;
  /** Border weight in eighths of a point (`w:sz`). Default 4 (= 0.5pt). */
  size?: number;
  /** Border color as a hex RRGGBB triplet without '#', or `"auto"`. Default `"auto"`. */
  color?: string;
}

export interface DocxSessionSettings {
  /**
   * Maximum undo steps retained. Default 20 (was 50).
   *
   * Each step is a full snapshot of every snapshot-scoped part, so this is a STEP count and not
   * a memory bound — the cost of one step scales with the document. Use `undoMemoryBudgetBytes`
   * to bound the heap.
   */
  undoDepth?: number;
  /**
   * Approximate ceiling, in bytes, on memory held by undo/redo snapshots. Default 134217728
   * (128 MiB). When exceeded the oldest history is discarded, so on a large document undo may
   * not reach the full `undoDepth`; one step is always retained. Set to 0 to bound by depth
   * alone (the pre-9.10 behavior).
   */
  undoMemoryBudgetBytes?: number;
  validateRawOps?: boolean;
  trackedChanges?: "accept" | "render_inline" | "strip_deletions";
  revisionAuthor?: string;
  /**
   * When false (default), `save()` strips the projector's internal `PtOpenXml:Unid`
   * attributes before serializing — they aren't OOXML schema, and persisting them
   * bloats large documents significantly (~700 KB on a 100-page DOCX). Set to true
   * only when anchor ids must survive a save/reopen round trip.
   */
  persistAnchorIds?: boolean;
  /**
   * When true, ReplaceText / ReplaceTextRange / ReplaceMatch payloads have ASCII
   * `"` and `'` converted to typographic curly quotes (`“ ” ‘ ’`) based on
   * context — open at start / after whitespace / after open-bracket, close
   * elsewhere. Avoids the cosmetic regression where a replacement lands as
   * straight-quoted text adjacent to surrounding already-curly text. Default false.
   */
  smartQuotes?: boolean;
  /**
   * When false, mutation EditResults omit the markdown patch and skip the per-op
   * scope re-projection that builds it — the fast path for clients that re-render
   * from HTML (the browser editor) rather than consuming markdown patches.
   * Default true.
   */
  emitMarkdownPatch?: boolean;
  /**
   * When `true` (default), the session projects the document at construction
   * time so {@link DocxSession.getDiff} can compare initial vs. current, and
   * retains the exact opening package for `getSemanticChanges()`.
   * Set to `false` to skip the upfront projection plus package-copy cost if you
   * do not plan to call either comparison API.
   */
  captureInitialProjection?: boolean;
}

/** One top-level render unit in a {@link RenderPlan}: a body block (`p`/`h`/`li`),
 *  one whole table (`tbl`), or one footnote/endnote definition (`fn`/`en`). */
export interface RenderUnit {
  id: string;
  kind: string;
}

/** Ordered top-level render units per scope container — the authority for "what
 *  blocks exist, in what order" that the incremental reconciler diffs against. */
export interface RenderPlan {
  body: RenderUnit[];
  footnotes: RenderUnit[];
  endnotes: RenderUnit[];
}

/** One footnote/endnote in citation order. `id` is the note's `w:id`; `ordinal`
 *  is its 1-based citation position — which IS its displayed number. */
export interface NoteListEntry {
  id: string;
  defAnchorId: string;
  ordinal: number;
}

export interface DocxSessionProjection {
  markdown: string;
  anchorIndex: Record<string, {
    partUri: string;
    unid: string;
    kind: string;
    scope: string;
    textPreview: string;
  }>;
  /** Present only when projectAnchor requested citations. */
  pageCitations?: Record<string, PageCitation>;
}

export interface PageCitationRequest {
  documentVersion: number;
  rendererFingerprint: string;
}

export type PageCitationUnavailableReason =
  | "no_page_map"
  | "continuous_mode"
  | "stale_document_version"
  | "renderer_fingerprint_mismatch"
  | "anchor_not_mapped";

export interface PageCitationFragment {
  fragmentId: string;
  anchorId: string;
  fragmentIndex: number;
  pageNumber: number;
  geometry: { x: number; y: number; width: number; height: number };
  story: "body" | "header" | "footer" | "footnote" | "endnote" | "comment";
  inTableCell: boolean;
}

export interface PageCitationPage {
  pageNumber: number;
  pageInSection: number;
  width: number;
  height: number;
  sectionIndex?: number;
  pageName: string;
}

export interface PageCitation {
  anchorId: string;
  availability: "available" | "unavailable";
  unavailableReason?: PageCitationUnavailableReason;
  documentVersion: number;
  rendererFingerprint: string;
  pages: PageCitationPage[];
  fragments: PageCitationFragment[];
}

export interface PageMapRegistrationResult {
  success: boolean;
  error?: "unsupported_schema_version" | "stale_document_version" | "renderer_fingerprint_mismatch" | "invalid_map";
  message?: string;
}

export interface PageMapStatus {
  availability: "available" | "unavailable";
  unavailableReason?: PageCitationUnavailableReason;
  documentVersion: number;
  rendererFingerprint?: string;
  mode?: "paginated" | "continuous";
}

/**
 * Per-fragment visible formatting reported by {@link DocxSession.grep}.
 */
export interface RunFormatting {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  code: boolean;
  color?: string;
  hyperlinkUrl?: string;
  runStyle?: string;
}

/**
 * One piece of a {@link TextMatch} that came from a single `<w:r>` run.
 */
export interface RunFragment {
  /** PtOpenXml:Unid of the `w:r` element this fragment came from. */
  unid: string;
  /** The text from this run that participates in the match. */
  text: string;
  /** Character offset + length of this fragment inside the run's flat text. */
  spanInElement: CharSpan;
  /** Visible formatting of the run this fragment came from. */
  formatting: RunFormatting;
}

/**
 * A single match returned by {@link DocxSession.grep}. The match always lives
 * within one block-level element.
 */
export interface TextMatch {
  text: string;
  enclosingAnchor: AnchorRef;
  span: CharSpan;
  fragments: RunFragment[];
  contextBefore: string;
  contextAfter: string;
  /** Regex capture groups; index 0 is always the whole match. */
  groups: string[];
  /** Present only when grep requested a citation for this exact render. */
  citation?: PageCitation;
}

/**
 * One block's contribution to a {@link CrossBlockMatch}. The slice's `fragments`
 * list is empty when the match touches an empty paragraph — the slice is still
 * recorded so callers can see the match crossed the empty block.
 */
export interface BlockSlice {
  anchor: AnchorRef;
  /** Character offset + length of the slice within the block's own flat text. */
  spanInBlock: CharSpan;
  /** Run fragments contributing to this slice, in document order. */
  fragments: RunFragment[];
}

/**
 * A single match returned by {@link DocxSession.grepCrossBlock}. The match may
 * span multiple adjacent block-level elements (paragraphs/headings/list items)
 * under the same parent container. `slices` is the per-block breakdown;
 * `enclosingAnchors` lists every block the match touches, in document order.
 *
 * Block boundaries appear in `text` / `contextBefore` / `contextAfter` as
 * single `\n` characters.
 */
export interface CrossBlockMatch {
  text: string;
  enclosingAnchors: AnchorRef[];
  slices: BlockSlice[];
  contextBefore: string;
  contextAfter: string;
  /** Regex capture groups; index 0 is always the whole match. */
  groups: string[];
  /** One per enclosingAnchors entry when requested. */
  citations?: PageCitation[];
}

/**
 * Options for {@link DocxSession.replaceTextRange}.
 */
export interface ReplaceOptions {
  /** Case-insensitive matching for the literal `find` needle. */
  ignoreCase?: boolean;
  /** Cap the number of replacements; omitted = unlimited. */
  maxReplacements?: number;
  /** Require exactly this many occurrences before applying any replacement. */
  expectedMatchCount?: number;
  /** Optional document/anchor guards evaluated before searching. */
  preconditions?: MutationPreconditions;
}

/**
 * Categories of bracketed placeholders {@link DocxSession.findPlaceholders} recognizes.
 *
 * - `blank_fill` — `[___]` or `$[___]` value slots
 * - `alternative_clause` — `[entire clause text]`
 * - `instruction` — `[insert X]`, `[specify Y]`, `[*italicized hint*]`
 */
export type PlaceholderKind = "blank_fill" | "alternative_clause" | "instruction";

/**
 * Numeric flag layout matching the .NET `PlaceholderKinds` enum. Combine with bitwise OR.
 */
export const PlaceholderKinds = {
  BlankFill: 1,
  AlternativeClause: 2,
  Instruction: 4,
  All: 7,
} as const;

/**
 * Numeric flag layout matching the .NET `DiffFormat` enum. Use with
 * {@link DocxSession.getDiff}.
 *
 * - `Json` (default) — anchor-keyed structured diff. Returns a `DiffEntry[]`.
 * - `Unified` — `patch(1)`-compatible unified diff over the markdown projection.
 *   Returns a single string (`""` when nothing changed).
 * - `SideBySide` — two-column human-review diff (`diff -y` style) over the
 *   markdown projection. Returns a single string.
 */
export const DiffFormat = {
  Json: 0,
  Unified: 1,
  SideBySide: 2,
} as const;

/**
 * A single anchor-keyed change in the diff between an initial and current projection.
 */
export interface DiffEntry {
  op: "delete" | "insert" | "modify";
  anchorId: string;
  /** Pre-change text content for delete/modify; absent for insert. */
  before?: string;
  /** Post-change text content for insert/modify; absent for delete. */
  after?: string;
}

/**
 * Aggregate snapshot of edit-state introspection signals returned by
 * {@link DocxSession.getEditSummary}.
 */
export interface EditSummary {
  totalAnchors: number;
  remainingPlaceholders: TemplatePlaceholder[];
  bareUnderscoreRuns: TextMatch[];
  footnoteCount: number;
  inlineFootnoteRefCount: number;
  commentCount: number;
}

/**
 * Numeric flag layout matching the .NET `ContextBoundary` enum. Controls
 * how `Grep` / `GrepCrossBlock` / `FindPlaceholders` decide where to stop
 * walking outward when computing `TextMatch.contextBefore` / `contextAfter`.
 *
 * - `Char` (default) — truncate at `contextChars`. Matches legacy behavior.
 * - `Bracket` — stop at `[` or `]`. Use for template fills: each placeholder's
 *   context is unambiguously its own even when multiple placeholders crowd
 *   into one sentence.
 * - `Sentence` — stop at `.`, `!`, `?`, `:`, `;`.
 * - `Comma` — stop at `,`. For matches inside enumerations.
 */
export const ContextBoundary = {
  Char: 0,
  Bracket: 1,
  Sentence: 2,
  Comma: 3,
} as const;

export interface TemplatePlaceholder {
  kind: PlaceholderKind;
  /** For `instruction` placeholders: the inner text with surrounding brackets/asterisks stripped. */
  hint?: string;
  match: TextMatch;
  /**
   * Additional plausible classifications when the primary `kind` is borderline.
   * Empty by default. The classic case is a long bracketed clause that happens
   * to contain a `_______` blank: primary `kind` stays `"blank_fill"`
   * (back-compat) and `alternativeKinds` contains `"alternative_clause"`.
   */
  alternativeKinds: PlaceholderKind[];
}

/**
 * Options for {@link DocxSession.fillPlaceholders}.
 */
export interface FillOptions {
  /** Which placeholder kinds to fill. Defaults to `PlaceholderKinds.All` so the
   *  picker is invoked for every kind in the doc. Narrow with e.g.
   *  `PlaceholderKinds.BlankFill | PlaceholderKinds.Instruction` to ignore
   *  bracketed alternative clauses. */
  kinds?: number;
  /** Which package parts to scan. Defaults to body (1). */
  scope?: number;
  /** Max iteration passes for multi-pass nested-bracket scenarios. Default 8. */
  maxPasses?: number;
  /** When the match starts with `$` and the picker's return value doesn't,
   *  preserve the `$` by prepending it. Default true. */
  preserveDollarPrefix?: boolean;
  /** Cap on `contextBefore` / `contextAfter` length on each side. Default 80. */
  contextChars?: number;
  /** Where to stop walking outward when computing context. Numeric layout
   *  matching {@link ContextBoundary}. Default `Char` (0). */
  boundary?: number;
  /** When the picker returns an empty string (and after `$`-prefix preservation
   *  has been applied), look at the chars immediately adjacent to the placeholder
   *  span and absorb surrounding whitespace / leading-space-before-punctuation /
   *  matched-brackets so the dropped placeholder doesn't leave cosmetic
   *  artifacts. Default `false` (preserve the literal-delete behavior).
   *
   *  Rules: whitespace on both sides collapses to one space; whitespace before
   *  and clause-terminating punctuation (`. , ; : ! ?`) after drops the leading
   *  space; matched open/close brackets (`()` `[]` `{}`) on either side drop
   *  both. NBSP / narrow NBSP / thin space are treated as whitespace.
   *
   *  Caveat: `$`-prefix preservation runs first, so a picker returning `""` for
   *  `$[xxx]` with `preserveDollarPrefix: true` (default) ends up replacing with
   *  `"$"` and coalescing is skipped. Set `preserveDollarPrefix: false` when you
   *  want the `$` to drop along with the brackets. */
  coalesceWhitespaceAroundEmptyFill?: boolean;
}

/**
 * Aggregate result returned by {@link DocxSession.fillPlaceholders}.
 *
 * `skipped` counts placeholders the picker returned null for in the first pass
 * that saw them — it stays > 0 even if later passes finished those same
 * placeholders. Use `stillPresent` (post-loop document state) for the
 * trustworthy "is the template done?" check; `skipped > 0 && stillPresent === 0`
 * means "picker said no the first time but later passes resolved it."
 */
export interface BulkEditResult {
  filled: number;
  skipped: number;
  /** Number of placeholders matching `options.kinds` in `options.scope` that
   *  remain in the document after the final pass. `0` means the template is
   *  fully filled for the requested kinds/scope. */
  stillPresent: number;
  passes: number;
  unfilled: TemplatePlaceholder[];
  errors: EditError[];
}

/**
 * Options for {@link DocxSession.grep}.
 *
 * `regexOptions` and `scope` use the numeric flag layouts of the .NET
 * `System.Text.RegularExpressions.RegexOptions` and `ProjectionScopes` enums.
 * Common values:
 *   - `RegexOptions.IgnoreCase = 1`
 *   - `RegexOptions.Multiline = 2`
 *   - `ProjectionScopes.Body = 1`, `Headers = 2`, `Footers = 4`,
 *     `Footnotes = 8`, `Endnotes = 16`, `Comments = 32`, `All = 63`.
 */
export interface GrepOptions {
  regexOptions?: number;
  scope?: number;
  contextChars?: number;
  /**
   * Whitespace handling. Numeric layout matching the .NET `WhitespaceMode` enum:
   *   - 0 = Preserve (default; match against the document's original characters)
   *   - 1 = Normalize (fold NBSP / narrow-NBSP / thin-space to ASCII space before matching)
   */
  whitespace?: number;
  /**
   * Where to stop walking outward when computing `TextMatch.contextBefore` /
   * `contextAfter`. Numeric layout matching the .NET `ContextBoundary` enum;
   * use the {@link ContextBoundary} const. Default `Char` (0) — truncate at
   * `contextChars`.
   */
  boundary?: number;
  /** Attach citations only if this exact registered layout is still valid. */
  citation?: PageCitationRequest;
}

/**
 * Options that tune the `findBy*` helpers on {@link DocxSession}. Mirrors the
 * .NET `FindOptions` record; wire keys are already camelCase, so this object is
 * serialized straight to the bridge. Omit it (or pass `{}`) for the defaults
 * (case-sensitive, whitespace-preserving, all scopes, no kind filter).
 */
export interface FindOptions {
  /** Case-insensitive matching. */
  ignoreCase?: boolean;
  /** Fold NBSP / narrow-NBSP / thin-space to ASCII space before matching. */
  ignoreWhitespace?: boolean;
  /** Only return anchors of this kind (e.g. `"h"` for headings, `"p"` for paragraphs). */
  kindFilter?: string;
  /**
   * Coarse-grained scope flag set (Body / Headers / Footers / Footnotes /
   * Endnotes / Comments). Numeric layout matching the .NET `ProjectionScopes`
   * flags enum; use {@link ProjectionScopes}. Defaults to all scopes.
   */
  scopes?: number;
  /**
   * Target a single named package part (e.g. `"hdr1"`). Prefer {@link scopes}
   * for whole-category filtering; this is for the rare single-part case.
   */
  scopeFilter?: string;
  /** Attach citations only if this exact registered layout is still valid. */
  citation?: PageCitationRequest;
}

/**
 * Resolved location of an anchor — what {@link DocxSession.findByAnnotation} and
 * the other discovery helpers return. The shape is {@link AnchorRef} plus the
 * `partUri` of the package part the element lives in (useful for callers that
 * walk the underlying OOXML directly).
 */
export interface AnchorTargetRef extends AnchorRef {
  partUri: string;
  /** First ~80 characters of the element's flat text — for previewing/picking anchors. */
  textPreview: string;
  /** Resolved auto-numbering prefix (e.g. "1.", "First") when the element carries
   *  numbering. Absent otherwise. See {@link MarkdownAnchorTarget.autoNumberPrefix}. */
  autoNumberPrefix?: string;
  /** Present only when the discovery call requested an exact page citation. */
  citation?: PageCitation;
}

/**
 * The shape returned by {@link DocxSession.getAnchorInfo}.
 * Use {@link MarkdownAnchorTarget} when iterating a full projection — it
 * includes the same fields plus `unid` and `partUri`.
 */
export interface AnchorInfo {
  id: string;
  kind: string;
  scope: string;
  textPreview: string;
  /** Exact live subtree hash suitable for expectedContentHash. */
  contentHash: string;
  /** Exact (untruncated) reader-visible text suitable for expectedText. */
  visibleText: string;
  /** Resolved auto-numbering prefix (e.g. "1.", "First") when the element carries
   *  numbering. Absent for un-numbered paragraphs or non-paragraph kinds. */
  autoNumberPrefix?: string;
}

/** Six list formats supported by the list write surface (decimal, upperLetter,
 *  lowerLetter, upperRoman, lowerRoman, bullet). Surfaced on
 *  {@link ListMembership.format} as a string union (mirrors the JSON wire format). */
export type NumberFormat =
  | "decimal"
  | "upperLetter"
  | "lowerLetter"
  | "upperRoman"
  | "lowerRoman"
  | "bullet";

/** Numbering facts for a list-item paragraph. Returned by
 *  {@link DocxSession.getListMembership} and surfaced as {@link BlockMetadata.list}. */
export interface ListMembership {
  /** Stable paragraph anchor accepted unchanged by every list mutation method. */
  anchorId: string;
  /** The w:numId the paragraph belongs to (the w:num instance). */
  numId: number;
  /** The w:abstractNumId the paragraph's w:num points at. */
  abstractNumId: number;
  /** The paragraph's level (w:ilvl), 0-8. */
  level: number;
  /** Resolved format for this level. */
  format: NumberFormat;
  /** Always true for a paragraph carrying w:numPr (inline or via style). */
  isAutoNumbered: boolean;
  /** True when the w:numPr is inherited from the paragraph's style chain. */
  fromStyle: boolean;
  /** Start-override from w:lvlOverride/w:startOverride for this level, if any. */
  startOverride?: number;
  /** Level definition's w:start value (1 when omitted). */
  start: number;
  /** Marker template such as "%1." or "(%2)". */
  levelText?: string;
  leftIndentTwips?: number;
  rightIndentTwips?: number;
  firstLineIndentTwips?: number;
  hangingIndentTwips?: number;
  /** Resolved label (e.g. "1.", "(a)") — same value surfaced via AnchorInfo.autoNumberPrefix. */
  generatedLabel?: string;
}

/** Block-level structural metadata. Returned by {@link DocxSession.getBlockMetadata}. */
export interface BlockMetadata {
  anchorId: string;
  kind: string;
  scope: string;
  styleId?: string;
  styleName?: string;
  /** 0-based outline level (Word convention). */
  outlineLevel?: number;
  list?: ListMembership;
  /** True when any descendant w:r carries a non-empty w:rPr. */
  hasInlineFormatting: boolean;
}

/**
 * A `w:headerReference`/`w:footerReference` on a section: the story kind it supplies and the
 * URI of the part holding it. Maps a {@link HeaderFooterKind} to a part — and thence to that
 * part's projection anchors, which carry the same `partUri` — instead of guessing from
 * part-collection order, which carries no kind information.
 */
export interface HeaderFooterRef {
  kind: HeaderFooterKind;
  partUri: string;
  /** True when this section declares no reference of `kind` and the story is INHERITED from the
   *  nearest preceding section that does (ECMA-376 §17.6.17). Editing it edits the shared part. */
  inherited: boolean;
}

/** Page-layout snapshot for the w:sectPr that governs an anchor.
 *  Returned by {@link DocxSession.getSectionInfo}. */
export interface SectionInfo {
  /** Body anchor used for the lookup; accepted unchanged by section mutation methods. */
  anchorId: string;
  sectionUnid: string;
  pageWidthTwips: number;
  pageHeightTwips: number;
  landscape: boolean;
  marginTopTwips: number;
  marginBottomTwips: number;
  marginLeftTwips: number;
  marginRightTwips: number;
  columns: number;
  headerPartUris: string[];
  footerPartUris: string[];
  /** Header references in declaration order, each with its `w:type`. Describes exactly the
   *  parts {@link SectionInfo.headerPartUris} lists, plus the kind each supplies. */
  headerRefs: HeaderFooterRef[];
  /** Footer references in declaration order, each with its `w:type`. Describes exactly the
   *  parts {@link SectionInfo.footerPartUris} lists, plus the kind each supplies. */
  footerRefs: HeaderFooterRef[];
  /** The page number this section starts at (`w:pgNumType/@w:start`). Absent when the section
   *  continues the previous section's numbering. */
  pageNumberStart?: number;
  /** This section's page-number format (`w:pgNumType/@w:fmt`). Absent means Word's default
   *  `1, 2, 3` — deliberately not reported as `"decimal"`, so a UI can tell "inherits" from
   *  "explicitly decimal" and avoid writing an attribute the document never had. */
  pageNumberFormat?: NumberFormat;
}

/** High-signal paragraph properties. Optional fields are deliberately absent when a
 * direct formatting layer did not write them; effective layers include schema defaults. */
export interface ParagraphFormatting {
  styleId?: string;
  alignment?: "left" | "center" | "right" | "justify";
  leftIndentTwips?: number;
  rightIndentTwips?: number;
  firstLineIndentTwips?: number;
  hangingIndentTwips?: number;
  spacingBeforeTwips?: number;
  spacingAfterTwips?: number;
  lineSpacing?: number;
  lineSpacingRule?: LineSpacingRule;
  keepNext?: boolean;
  keepLines?: boolean;
  pageBreakBefore?: boolean;
  outlineLevel?: number;
  shadingFill?: string;
  topBorder?: ParagraphBorderEdge;
  bottomBorder?: ParagraphBorderEdge;
}

/** High-signal character properties. Nullable-at-source fields are optional on the wire so
 * an absent direct property remains distinguishable from an explicit false/zero. */
export interface RunFormattingInfo {
  styleId?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  underlineStyle?: string;
  strike?: boolean;
  code?: boolean;
  color?: string;
  highlight?: string;
  vertAlign?: string;
  fontSizePts?: number;
  fontFamily?: string;
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
}

export interface TableStyleFormatting {
  alignment?: string;
  widthTwips?: number;
  indentTwips?: number;
  layout?: string;
  hasBorders?: boolean;
  cellShadingFill?: string;
}

/** One explicit document style. `id` is accepted unchanged by paragraph/run style mutations. */
export interface StyleInfo {
  id: string;
  name: string;
  type: "paragraph" | "character" | "table" | "numbering" | string;
  basedOn?: string;
  next?: string;
  isDefault: boolean;
  isCustom: boolean;
  hasLatentException: boolean;
  uiPriority?: number;
  semiHidden?: boolean;
  unhideWhenUsed?: boolean;
  quickFormat?: boolean;
  locked?: boolean;
  resolvedParagraph?: ParagraphFormatting;
  resolvedRun?: RunFormattingInfo;
  resolvedTable?: TableStyleFormatting;
}

/** One text-bearing run. `anchorId` + `span` can be passed unchanged to applyFormat. */
export interface InlineSpan {
  anchorId: string;
  runUnid: string;
  span: CharSpan;
  text: string;
  direct: RunFormattingInfo;
  effective: RunFormattingInfo;
  /** Outer-to-inner native content controls containing this run. */
  contentControlAnchorIds: string[];
}

/** Explicitly separated direct and effective formatting for one paragraph anchor. */
export interface FormattingInspection {
  anchorId: string;
  directParagraph: ParagraphFormatting;
  effectiveParagraph: ParagraphFormatting;
  runs: InlineSpan[];
}

/**
 * A custom annotation persisted in the document via Docxodus' annotation system.
 * Returned by {@link DocxSession.listAnnotations}; mirrors the wire-relevant
 * fields of the .NET `DocumentAnnotation` type. The page-info cache fields
 * (`startPage`/`endPage`/`pageInfoStale`/`pageInfoComputedAt`) are omitted to
 * keep the JSON payload compact — callers that need them can use the .NET API
 * directly. The `metadata` bag is emitted only when non-empty.
 *
 * See `docs/architecture/custom_annotations.md` for the persistence design.
 */
export interface DocumentAnnotation {
  /** Unique annotation identifier (caller-supplied at add time). */
  id: string;
  /** Label category/type (e.g. `"INDEMNIFICATION"`, `"CLAUSE_TYPE_A"`). */
  labelId: string;
  /** Human-readable label text displayed in the UI. */
  label: string;
  /** Highlight color in hex (e.g. `"#FFEB3B"`). */
  color: string;
  /** Internal bookmark name in the DOCX (`_Docxodus_Ann_{id}` for managed annotations). */
  bookmarkName: string;
  /** Author who created the annotation, if recorded. */
  author?: string;
  /** Creation timestamp in ISO-8601 (round-trip) format, if recorded. */
  created?: string;
  /** The text content covered by the annotation's bookmark, populated when reading. */
  annotatedText?: string;
  /** Arbitrary string→string metadata bag persisted with the annotation. */
  metadata?: Record<string, string>;
}

/**
 * Partial-update payload for {@link DocxSession.updateAnnotation}.
 * Null/missing fields leave the existing value unchanged. `metadataPatch`
 * is a per-key merge: a non-null value sets the key, an explicit `null`
 * removes it, a missing key leaves it unchanged.
 */
export interface AnnotationUpdate {
  labelId?: string;
  label?: string;
  color?: string;
  author?: string;
  metadataPatch?: Record<string, string | null>;
}

/**
 * Severity level for comparison log entries.
 */
export enum ComparisonLogLevel {
  /** Informational message about the comparison process */
  Info = "Info",
  /** Warning about a potential issue that didn't prevent comparison */
  Warning = "Warning",
  /** Error that may affect comparison results but didn't stop processing */
  Error = "Error",
}

/**
 * A single log entry from the comparison process.
 */
export interface ComparisonLogEntry {
  /** Severity level: "Info", "Warning", or "Error" */
  level: ComparisonLogLevel | string;
  /**
   * Machine-readable code identifying the type of issue.
   * Examples: "ORPHANED_FOOTNOTE_REFERENCE", "MISSING_STYLE"
   */
  code: string;
  /** Human-readable description of the issue */
  message: string;
  /** Additional context or technical details (optional) */
  details?: string;
  /**
   * Location in the document where the issue occurred (optional).
   * Format: "part/xpath" e.g., "document.xml/w:footnoteReference[@w:id='3']"
   */
  location?: string;
}

/**
 * Result from comparison operations that includes a log of warnings/errors.
 */
export interface CompareResultWithLog {
  /** Whether the comparison succeeded */
  success: boolean;
  /** The redlined document as a Uint8Array (only if success is true) */
  document?: Uint8Array;
  /** Error message if success is false */
  error?: string;
  /** Log entries from the comparison process */
  log: ComparisonLogEntry[];
  /** Whether the log contains any warnings */
  hasWarnings: boolean;
  /** Whether the log contains any errors */
  hasErrors: boolean;
}

/**
 * Result from HTML comparison operations that includes a log of warnings/errors.
 */
export interface CompareToHtmlResultWithLog {
  /** Whether the comparison succeeded */
  success: boolean;
  /** The HTML output (only if success is true) */
  html?: string;
  /** Error message if success is false */
  error?: string;
  /** Log entries from the comparison process */
  log: ComparisonLogEntry[];
  /** Whether the log contains any warnings */
  hasWarnings: boolean;
  /** Whether the log contains any errors */
  hasErrors: boolean;
}

/**
 * Well-known log entry codes used by the comparison engine.
 */
export const ComparisonLogCodes = {
  /** A footnote reference in the document body has no corresponding footnote definition */
  OrphanedFootnoteReference: "ORPHANED_FOOTNOTE_REFERENCE",
  /** An endnote reference in the document body has no corresponding endnote definition */
  OrphanedEndnoteReference: "ORPHANED_ENDNOTE_REFERENCE",
  /** A style referenced in the document is not defined in styles.xml */
  MissingStyle: "MISSING_STYLE",
  /** A numbering definition referenced in the document is missing */
  MissingNumberingDefinition: "MISSING_NUMBERING_DEFINITION",
  /** A relationship referenced in the document is missing */
  MissingRelationship: "MISSING_RELATIONSHIP",
  /** An image or media file referenced in the document is missing */
  MissingMedia: "MISSING_MEDIA",
  /** The document structure contains unexpected or malformed XML */
  MalformedXml: "MALFORMED_XML",
  /** A bookmark reference has no corresponding bookmark start/end */
  OrphanedBookmark: "ORPHANED_BOOKMARK",
} as const;

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
  /** Canonical session anchor when this element is addressable. */
  anchorId?: string;
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
  /** Canonical `col` anchor. */
  anchorId: string;
  /** Canonical owning `tbl` anchor. */
  tableAnchorId: string;
  isVirtual: boolean;
  /** Zero-based column index */
  columnIndex: number;
  /** IDs of all cells in this column */
  cellIds: string[];
  /** Canonical `tc` anchors for cells covering this column. */
  cellAnchorIds: string[];
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
  /** Explicit provenance: always "heuristic"; use PageMap for authoritative pages. */
  estimatedPageCountSource: "heuristic";
}

// ============================================================================
// Web Worker Types (Phase 2: Non-blocking WASM operations)
// ============================================================================

/**
 * Message types sent from main thread to worker.
 */
export type WorkerRequestType =
  | "init"
  | "generatePackageManifest"
  | "verifyDeliverable"
  | "proveRedlineReversibility"
  | "projectReviewProfile"
  | "convertDocxToHtml"
  | "compareDocuments"
  | "compareDocumentsToHtml"
  | "getSemanticChanges"
  | "getRevisions"
  | "getDocumentMetadata"
  | "getVersion"
  | "prepare"
  | "sessionOpen"
  | "sessionGetPackageManifest"
  | "sessionGetSemanticChanges"
  | "sessionVerifyDeliverable"
  | "sessionClose"
  | "sessionAddAnnotation"
  | "sessionRemoveAnnotation"
  | "sessionUpdateAnnotation"
  | "sessionMoveAnnotation";

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
  /** Private exact-view copy of the caller's document bytes, transferred to the worker. */
  documentBytes: Uint8Array;
  /** Conversion options */
  options?: ConversionOptions;
  /** Optional main-thread admission ceiling for the UTF-8 response. */
  maximumOutputBytes?: number;
}

/** Generate a deterministic package manifest without opening a live session. */
export interface WorkerGeneratePackageManifestRequest extends WorkerRequestBase {
  type: "generatePackageManifest";
  documentBytes: Uint8Array;
  /** When present, these lower ceilings constrain #493 inspection itself. */
  limits?: PackageManifestInspectionLimits;
  /**
   * Which representation the caller needs. A manifest near the entry ceiling is multi-megabyte,
   * so returning both costs a parse plus a structured clone nobody reads. Defaults to `"both"`.
   */
  representation?: "object" | "json" | "both";
}

/** Derive exact final/original package bytes before conversion. */
export interface WorkerProjectReviewProfileRequest extends WorkerRequestBase {
  type: "projectReviewProfile";
  documentBytes: Uint8Array;
  profile: "final" | "original";
  /** Do not transfer a derived package larger than this many bytes. */
  maximumOutputBytes?: number;
}

/** Run the default deliverable gate directly over exact supplied bytes. */
export interface WorkerVerifyDeliverableRequest extends WorkerRequestBase {
  type: "verifyDeliverable";
  documentBytes: Uint8Array;
  baselineBytes?: Uint8Array;
}

/**
 * Prove redline accept/reject reversibility off the main thread. Three packages are inspected and
 * two are rebuilt, so this is the heaviest verification request the worker serves.
 */
export interface WorkerProveRedlineReversibilityRequest extends WorkerRequestBase {
  type: "proveRedlineReversibility";
  baselineBytes: Uint8Array;
  intendedFinalBytes: Uint8Array;
  redlineBytes: Uint8Array;
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

/** Compare two packages into the stable, versioned semantic-change schema. */
export interface WorkerGetSemanticChangesRequest extends WorkerRequestBase {
  type: "getSemanticChanges";
  leftBytes: Uint8Array;
  rightBytes: Uint8Array;
  settings?: DocxDiffSettings;
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
 * Warm up the comparison code path so the next compare triggers no further
 * WASM assembly fetches. Carries no payload.
 */
export interface WorkerPrepareRequest extends WorkerRequestBase {
  type: "prepare";
}

/**
 * Open a DocxSession in the worker.
 */
export interface WorkerSessionOpenRequest extends WorkerRequestBase {
  type: "sessionOpen";
  /** Private exact-view copy of the caller's document bytes, transferred to the worker. */
  documentBytes: Uint8Array;
  /** Session settings as JSON */
  settingsJson?: string;
}

/** Generate a manifest from the current logical checkpoint of a worker session. */
export interface WorkerSessionGetPackageManifestRequest extends WorkerRequestBase {
  type: "sessionGetPackageManifest";
  handle: number;
}

/** Compare a worker session's current checkpoint with its opening package. */
export interface WorkerSessionGetSemanticChangesRequest extends WorkerRequestBase {
  type: "sessionGetSemanticChanges";
  handle: number;
}

/** Run the default deliverable gate over a worker session's clean-save checkpoint. */
export interface WorkerSessionVerifyDeliverableRequest extends WorkerRequestBase {
  type: "sessionVerifyDeliverable";
  handle: number;
}

/**
 * Close a worker DocxSession.
 */
export interface WorkerSessionCloseRequest extends WorkerRequestBase {
  type: "sessionClose";
  /** Session handle returned by sessionOpen */
  handle: number;
}

/**
 * Add an annotation via a worker DocxSession.
 */
export interface WorkerSessionAddAnnotationRequest extends WorkerRequestBase {
  type: "sessionAddAnnotation";
  handle: number;
  anchorId: string;
  /** CharSpan as JSON, or empty string for block-level */
  spanJson: string;
  annotationJson: string;
}

/**
 * Remove an annotation via a worker DocxSession.
 */
export interface WorkerSessionRemoveAnnotationRequest extends WorkerRequestBase {
  type: "sessionRemoveAnnotation";
  handle: number;
  annotationId: string;
}

/**
 * Update an annotation via a worker DocxSession.
 */
export interface WorkerSessionUpdateAnnotationRequest extends WorkerRequestBase {
  type: "sessionUpdateAnnotation";
  handle: number;
  annotationId: string;
  updateJson: string;
}

/**
 * Move an annotation via a worker DocxSession.
 */
export interface WorkerSessionMoveAnnotationRequest extends WorkerRequestBase {
  type: "sessionMoveAnnotation";
  handle: number;
  annotationId: string;
  newAnchorId: string;
  /** CharSpan as JSON, or empty string for block-level */
  newSpanJson: string;
}

/**
 * Union type of all possible worker requests.
 */
export type WorkerRequest =
  | WorkerInitRequest
  | WorkerGeneratePackageManifestRequest
  | WorkerVerifyDeliverableRequest
  | WorkerProveRedlineReversibilityRequest
  | WorkerProjectReviewProfileRequest
  | WorkerConvertRequest
  | WorkerCompareRequest
  | WorkerCompareToHtmlRequest
  | WorkerGetSemanticChangesRequest
  | WorkerGetRevisionsRequest
  | WorkerGetDocumentMetadataRequest
  | WorkerGetVersionRequest
  | WorkerPrepareRequest
  | WorkerSessionOpenRequest
  | WorkerSessionGetPackageManifestRequest
  | WorkerSessionGetSemanticChangesRequest
  | WorkerSessionVerifyDeliverableRequest
  | WorkerSessionCloseRequest
  | WorkerSessionAddAnnotationRequest
  | WorkerSessionRemoveAnnotationRequest
  | WorkerSessionUpdateAnnotationRequest
  | WorkerSessionMoveAnnotationRequest;

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
  /**
   * Machine-readable cause when success is false. Callers classify failures from this
   * rather than by matching `error`, whose wording is not a contract.
   */
  errorCode?: WorkerErrorCode;
}

/** Closed set of machine-readable worker failure causes. */
export type WorkerErrorCode = "resource_limit";

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

export interface WorkerGeneratePackageManifestResponse extends WorkerResponseBase {
  type: "generatePackageManifest";
  manifest?: PackageManifest;
  /** Exact canonical JSON, retained for strict duplicate-property/schema validation. */
  manifestJson?: string;
}

export interface WorkerProjectReviewProfileResponse extends WorkerResponseBase {
  type: "projectReviewProfile";
  documentBytes?: Uint8Array;
}

export interface WorkerVerifyDeliverableResponse extends WorkerResponseBase {
  type: "verifyDeliverable";
  verification?: DeliverableVerificationResult;
}

export interface WorkerProveRedlineReversibilityResponse extends WorkerResponseBase {
  type: "proveRedlineReversibility";
  proof?: RedlineReversibilityProof;
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

/** Response containing the public semantic-change schema. */
export interface WorkerGetSemanticChangesResponse extends WorkerResponseBase {
  type: "getSemanticChanges";
  semanticChanges?: SemanticChangeSet;
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
 * Response from prepare request. Carries no payload beyond success/error.
 */
export interface WorkerPrepareResponse extends WorkerResponseBase {
  type: "prepare";
}

/**
 * Response from sessionOpen request.
 */
export interface WorkerSessionOpenResponse extends WorkerResponseBase {
  type: "sessionOpen";
  /** Integer handle identifying the session in the worker */
  handle?: number;
}

/** Response containing the current worker-session package manifest. */
export interface WorkerSessionGetPackageManifestResponse extends WorkerResponseBase {
  type: "sessionGetPackageManifest";
  manifest?: PackageManifest;
}

/** Response containing a session's public semantic-change schema. */
export interface WorkerSessionGetSemanticChangesResponse extends WorkerResponseBase {
  type: "sessionGetSemanticChanges";
  semanticChanges?: SemanticChangeSet;
}

export interface WorkerSessionVerifyDeliverableResponse extends WorkerResponseBase {
  type: "sessionVerifyDeliverable";
  verification?: DeliverableVerificationResult;
}

/**
 * Response from sessionClose request.
 */
export interface WorkerSessionCloseResponse extends WorkerResponseBase {
  type: "sessionClose";
}

/**
 * Response from session annotation write operations.
 * The `result` field is the serialised EditResult from the WASM bridge.
 */
export interface WorkerSessionEditResponse extends WorkerResponseBase {
  type:
    | "sessionAddAnnotation"
    | "sessionRemoveAnnotation"
    | "sessionUpdateAnnotation"
    | "sessionMoveAnnotation";
  /** EditResult returned by the session operation */
  result?: EditResult;
}

/**
 * Union type of all possible worker responses.
 */
export type WorkerResponse =
  | WorkerInitResponse
  | WorkerGeneratePackageManifestResponse
  | WorkerVerifyDeliverableResponse
  | WorkerProveRedlineReversibilityResponse
  | WorkerProjectReviewProfileResponse
  | WorkerConvertResponse
  | WorkerCompareResponse
  | WorkerCompareToHtmlResponse
  | WorkerGetSemanticChangesResponse
  | WorkerGetRevisionsResponse
  | WorkerGetDocumentMetadataResponse
  | WorkerGetVersionResponse
  | WorkerPrepareResponse
  | WorkerSessionOpenResponse
  | WorkerSessionGetPackageManifestResponse
  | WorkerSessionGetSemanticChangesResponse
  | WorkerSessionVerifyDeliverableResponse
  | WorkerSessionCloseResponse
  | WorkerSessionEditResponse;

/**
 * Options for creating a worker-based Docxodus instance.
 */
export interface WorkerDocxodusOptions {
  /**
   * Base URL for loading WASM files.
   * Defaults to auto-detection from module URL.
   */
  wasmBasePath?: string;
  /** Abort the owned worker, including an initialization that has not completed. */
  signal?: AbortSignal;
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

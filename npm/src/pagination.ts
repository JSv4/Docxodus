/**
 * Pagination engine for creating a PDF.js-style paginated view from HTML output.
 *
 * This module provides client-side pagination that measures rendered content
 * and flows it across fixed-size page containers based on document dimensions.
 */

import { formatPageNumber } from "./page-number-format.js";
import {
  DEFAULT_MARGIN,
  DEFAULT_PAGE_HEIGHT,
  DEFAULT_PAGE_WIDTH,
  parseSectionDimensions,
  ptToPx,
  pxToPt,
  resolvePageBands,
} from "./page-geometry.js";
import type { PageBands, PageDimensions } from "./page-geometry.js";

export type { PageBands, PageDimensions } from "./page-geometry.js";

/**
 * Whether two sections describe the same physical page box. A continuous section
 * break is only honorable when it does — a changed page size or margin set forces
 * a real page break, which is Word's behavior too.
 */
function samePageBox(a: PageDimensions, b: PageDimensions): boolean {
  return (
    a.pageWidth === b.pageWidth &&
    a.pageHeight === b.pageHeight &&
    a.marginTop === b.marginTop &&
    a.marginRight === b.marginRight &&
    a.marginBottom === b.marginBottom &&
    a.marginLeft === b.marginLeft
  );
}

/**
 * Headers and footers for a specific section.
 */
export interface SectionHeaderFooter {
  /** Default header (used for odd pages or all pages) */
  headerDefault?: HTMLElement;
  /** First page header */
  headerFirst?: HTMLElement;
  /** Even page header */
  headerEven?: HTMLElement;
  /** Default footer (used for odd pages or all pages) */
  footerDefault?: HTMLElement;
  /** First page footer */
  footerFirst?: HTMLElement;
  /** Even page footer */
  footerEven?: HTMLElement;

  // Pre-measured heights (populated during registry parsing for lazy-loading compatibility)
  /** Measured height of default header in points */
  headerDefaultHeight?: number;
  /** Measured height of first page header in points */
  headerFirstHeight?: number;
  /** Measured height of even page header in points */
  headerEvenHeight?: number;
  /** Measured height of default footer in points */
  footerDefaultHeight?: number;
  /** Measured height of first page footer in points */
  footerFirstHeight?: number;
  /** Measured height of even page footer in points */
  footerEvenHeight?: number;
}

/**
 * Registry of headers and footers by section index.
 */
export type HeaderFooterRegistry = Map<number, SectionHeaderFooter>;

/**
 * A measured content block with metadata for pagination decisions.
 */
export interface MeasuredBlock {
  /** The DOM element */
  element: HTMLElement;
  /** Section whose body owns this block after a continuous section transition. */
  sectionIndex: number;
  /** Measured height in points (content + padding + border, excluding margins) */
  heightPt: number;
  /** Top margin in points */
  marginTopPt: number;
  /** Bottom margin in points */
  marginBottomPt: number;
  /** Whether to keep this block with the next one */
  keepWithNext: boolean;
  /** Whether to keep all lines of this block together */
  keepLines: boolean;
  /** Whether to force a page break before this block */
  pageBreakBefore: boolean;
  /** Whether this is a page break marker */
  isPageBreak: boolean;
  /** Whether the block is a Word paragraph whose top margin represents paragraph space-before */
  isWordParagraph?: boolean;
}

/**
 * Information about a rendered page.
 */
export interface PageInfo {
  /** 1-based page number */
  pageNumber: number;
  /** Section index this page belongs to */
  sectionIndex: number;
  /** Page dimensions */
  dimensions: PageDimensions;
  /** The page container element */
  element: HTMLElement;
}

/**
 * Result of pagination operation.
 */
export interface PaginationResult {
  /** Explicit completion result consumed by deterministic export barriers. */
  readiness: PaginationReadyResult;
  /** Total number of pages */
  totalPages: number;
  /** Array of page information */
  pages: PageInfo[];
  /** Present only when the caller supplied an exact layoutToken. */
  pageMap?: PageMap;
}

export interface PaginationDiagnostic {
  code:
    | "sections_processed"
    | "page_runs_processed"
    | "source_anchors_inventoried"
    | "note_references_inventoried";
  severity: "info";
  message: string;
  count: number;
}

export interface PaginationReadyResult {
  status: "ready";
  pageCount: number;
  diagnostics: PaginationDiagnostic[];
}

export type PageMapMode = "paginated" | "continuous";
export type PageMapAvailability = "available" | "unavailable";
export type PageMapStory = "body" | "header" | "footer" | "footnote" | "endnote" | "comment";

export interface PageMapRect {
  /** Page-relative points, independent of viewer zoom/transform. */
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PageMapPage {
  pageNumber: number;
  pageInSection: number;
  width: number;
  height: number;
  sectionIndex?: number;
  pageName: string;
}

export interface PageMapFragment {
  fragmentId: string;
  /** Canonical collision-safe `kind:scope:unid`, never the bare editor Unid. */
  anchorId: string;
  fragmentIndex: number;
  pageNumber: number;
  geometry: PageMapRect;
  story: PageMapStory;
  /** Table-cell ownership is orthogonal to story (e.g. a body or footnote table cell). */
  inTableCell: boolean;
}

/** Versioned portable layout contract consumed by DocxSession and remote agent surfaces. */
export interface PageMap {
  schemaVersion: 1;
  mode: PageMapMode;
  availability: PageMapAvailability;
  documentVersion: number;
  rendererFingerprint: string;
  pages: PageMapPage[];
  fragments: PageMapFragment[];
}

/** Explicit no-pages contract for a continuous viewer. It never estimates page numbers. */
export function createUnavailablePageMap(
  documentVersion: number,
  rendererFingerprint: string,
  mode: "continuous" = "continuous",
): PageMap {
  if (!rendererFingerprint) throw new Error("rendererFingerprint must be non-empty");
  return {
    schemaVersion: 1,
    mode,
    availability: "unavailable",
    documentVersion,
    rendererFingerprint,
    pages: [],
    fragments: [],
  };
}

export interface PageCitationNavigation {
  navigated: boolean;
  target?: HTMLElement;
  pageNumber?: number;
  fragmentId?: string;
  unavailableReason?: "citation_unavailable" | "fragment_not_found";
}

interface PageCitationHighlightState {
  target: HTMLElement;
  highlightClass?: string;
  addedClass?: boolean;
  inlineStyles?: Array<{ name: string; value: string; priority: string }>;
}

const activePageCitationHighlights = new WeakMap<ParentNode, PageCitationHighlightState>();

/** Remove the citation highlight previously applied within this paginated root. */
export function clearPageCitationHighlight(root: ParentNode): void {
  const active = activePageCitationHighlights.get(root);
  if (!active) return;
  if (active.highlightClass && active.addedClass) {
    active.target.classList.remove(active.highlightClass);
  }
  for (const style of active.inlineStyles ?? []) {
    if (style.value) {
      active.target.style.setProperty(style.name, style.value, style.priority);
    } else {
      active.target.style.removeProperty(style.name);
    }
  }
  activePageCitationHighlights.delete(root);
}

/**
 * Navigate an exact citation over an already-paginated DOM. The page-qualified fragment id is
 * authoritative; page + canonical source identity is a compatibility fallback for older v1 DOMs.
 */
export function navigateToPageCitation(
  root: ParentNode,
  citation: {
    availability: PageMapAvailability;
    anchorId: string;
    fragments: Array<{ fragmentId: string; pageNumber: number }>;
  },
  options: {
    highlightClass?: string;
    /** Apply a visible inline highlight when no class is supplied. Default true. */
    highlight?: boolean;
    behavior?: ScrollBehavior;
    block?: ScrollLogicalPosition;
  } = {},
): PageCitationNavigation {
  clearPageCitationHighlight(root);
  if (citation.availability !== "available" || citation.fragments.length === 0) {
    return { navigated: false, unavailableReason: "citation_unavailable" };
  }

  const byAttribute = (name: string, value: string, within: ParentNode = root): HTMLElement | null => {
    for (const node of Array.from(within.querySelectorAll<HTMLElement>(`[${name}]`))) {
      if (node.getAttribute(name) === value) return node;
    }
    return null;
  };

  const fragment = citation.fragments[0];
  let target = byAttribute("data-page-fragment-id", fragment.fragmentId);
  if (!target) {
    const page = byAttribute("data-page-number", String(fragment.pageNumber));
    if (page) target = byAttribute("data-source-anchor-id", citation.anchorId, page) ?? page;
  }
  if (!target) {
    return {
      navigated: false,
      pageNumber: fragment.pageNumber,
      fragmentId: fragment.fragmentId,
      unavailableReason: "fragment_not_found",
    };
  }

  if (options.highlightClass) {
    const addedClass = !target.classList.contains(options.highlightClass);
    target.classList.add(options.highlightClass);
    activePageCitationHighlights.set(root, {
      target,
      highlightClass: options.highlightClass,
      addedClass,
    });
  } else if (options.highlight !== false) {
    const names = ["outline", "outline-offset", "background-color"];
    const inlineStyles = names.map((name) => ({
      name,
      value: target.style.getPropertyValue(name),
      priority: target.style.getPropertyPriority(name),
    }));
    target.style.setProperty("outline", "3px solid #f4b400", "important");
    target.style.setProperty("outline-offset", "2px", "important");
    target.style.setProperty("background-color", "rgba(255, 235, 59, .18)", "important");
    activePageCitationHighlights.set(root, { target, inlineStyles });
  }
  target.scrollIntoView({
    behavior: options.behavior ?? "smooth",
    block: options.block ?? "center",
  });
  return {
    navigated: true,
    target,
    pageNumber: fragment.pageNumber,
    fragmentId: fragment.fragmentId,
  };
}

/**
 * Options for the pagination engine.
 */
export interface PaginationOptions {
  /** Scale factor for rendering (1.0 = 100%). Default: 1 */
  scale?: number;
  /** CSS class prefix used in the HTML. Default: "page-" */
  cssPrefix?: string;
  /** Whether to show page numbers. Default: true */
  showPageNumbers?: boolean;
  /** Gap between pages in pixels. Default: 20 */
  pageGap?: number;
  /**
   * Whether ordinary paragraphs may be fragmented across page boundaries.
   * Defaults to false for direct PaginationEngine callers; read-only viewer
   * entry points opt in explicitly.
   */
  fragmentParagraphs?: boolean;
  /** Cooperative checkpoint for bounded non-yielding browser layout work. */
  checkCancellation?: () => void;
  /**
   * Incremental admission check invoked immediately before each physical page is allocated.
   * Export callers use this to enforce `finalPages` without constructing an over-limit DOM first.
   */
  checkPageCount?: (prospectivePageCount: number) => void;
  /** Exact invalidation tokens used to materialize an authoritative PageMap with the result. */
  layoutToken?: { documentVersion: number; rendererFingerprint: string };
}

// Default letter size in points (612 x 792 = 8.5" x 11")
// Maximum percentage of content height that footnotes can occupy
// This allows footnotes to expand upward into body content space when needed
const MAX_FOOTNOTE_AREA_RATIO = 0.6; // 60% of content height
// Hidden measurement and final absolutely-positioned note bands differ by sub-pixel border/margin
// rounding in Chromium. Reserve a small physical-unit guard so the last baseline stays visible.
const FOOTNOTE_MEASUREMENT_GUARD_PT = 2;
/** Internal provenance retained only while a note element waits between pages. */
const FOOTNOTE_SOURCE_POSITION_ATTR = "data-pagination-footnote-source-position";

/**
 * Pagination engine that converts HTML with pagination metadata
 * into a paginated view with fixed-size page containers.
 */
/**
 * Registry of footnotes by ID for per-page distribution.
 */
export type FootnoteRegistry = Map<string, HTMLElement>;

/**
 * Tracks footnote content that needs to continue on the next page.
 */
interface FootnoteContinuation {
  /** The footnote ID being continued */
  footnoteId: string;
  /** Canonical fn:* definition identity from the registry wrapper. */
  sourceAnchorId?: string;
  /** Remaining paragraphs/elements that didn't fit */
  remainingElements: HTMLElement[];
}

/**
 * Tracks a partial footnote that was split on a page.
 */
interface PartialFootnote {
  /** The footnote ID */
  footnoteId: string;
  /** Elements that fit on this page */
  fittingElements: HTMLElement[];
}

/** Existing page payload that a newly split note must fit beside. */
interface FootnotePackingContext {
  footnoteIds: readonly string[];
  continuation?: FootnoteContinuation | null;
  partialFootnotes?: readonly PartialFootnote[];
}

type ParagraphSplitAttempt =
  | { kind: "split"; head: HTMLElement; tail: HTMLElement }
  | { kind: "no-fit" }
  | { kind: "indivisible" };

interface ParagraphFragmentEndpoint {
  node: Node;
  offset: number;
  /** UTF-16 offset in the paragraph's flattened text. */
  textOffset: number;
  /** Prefer a DOM boundary outside an atomic field over its last text-node boundary. */
  priority: number;
}

/**
 * A section's `w:pgNumType`, as stamped on its wrapper by the converter. Both fields are optional
 * because both attributes are: an absent `start` means the section continues the previous section's
 * numbering, and an absent `format` means Word's default `1, 2, 3`.
 */
interface SectionPageNumbering {
  start?: number;
  format?: string;
}

export class PaginationEngine {
  private stagingElement: HTMLElement;
  private containerElement: HTMLElement;
  private document: Document;
  private view: Window & typeof globalThis;
  private scale: number;
  private cssPrefix: string;
  private showPageNumbers: boolean;
  private pageGap: number;
  private fragmentParagraphs: boolean;
  private cancellationCheckpoint?: () => void;
  private pageCountCheckpoint?: (prospectivePageCount: number) => void;
  private createdPageCount = 0;
  private layoutToken?: { documentVersion: number; rendererFingerprint: string };
  private hfRegistry: HeaderFooterRegistry;
  private footnoteRegistry: FootnoteRegistry;
  private footnoteSeparator: HTMLElement | null = null;
  private footnoteContinuationSeparator: HTMLElement | null = null;
  private commentMarginRegistry: Map<string, HTMLElement>;
  private footnoteLayoutLikeWord8 = false;
  private pendingFootnoteContinuation: FootnoteContinuation | null = null;
  /** Per-section `w:pgNumType` (start / format), read off the section wrappers. */
  private pageNumbering: Map<number, SectionPageNumbering> = new Map();
  private lastPages: PageInfo[] = [];
  private expectedPageMapAnchorIds: Set<string> = new Set();
  private state: "ready" | "running" | "complete" | "failed" = "ready";

  /**
   * Creates a new pagination engine.
   *
   * @param staging - The staging element or its ID containing the content to paginate
   * @param container - The container element or its ID where pages will be rendered
   * @param options - Pagination options
   */
  constructor(
    staging: HTMLElement | string,
    container: HTMLElement | string,
    options: PaginationOptions = {}
  ) {
    const ownerDocument =
      typeof staging !== "string"
        ? staging.ownerDocument
        : typeof container !== "string"
          ? container.ownerDocument
          : globalThis.document;
    this.stagingElement =
      typeof staging === "string"
        ? (ownerDocument.getElementById(staging) as HTMLElement)
        : staging;
    this.containerElement =
      typeof container === "string"
        ? (ownerDocument.getElementById(container) as HTMLElement)
        : container;

    if (!this.stagingElement) {
      throw new Error("Staging element not found");
    }
    if (!this.containerElement) {
      throw new Error("Container element not found");
    }
    if (this.stagingElement.ownerDocument !== this.containerElement.ownerDocument) {
      throw new Error("Staging and container elements must belong to the same document");
    }
    const view = ownerDocument.defaultView;
    if (!view) {
      throw new Error("Pagination requires an attached document with a defaultView");
    }
    this.document = ownerDocument;
    this.view = view as Window & typeof globalThis;

    this.scale = options.scale ?? 1;
    this.cssPrefix = options.cssPrefix ?? "page-";
    this.showPageNumbers = options.showPageNumbers ?? true;
    this.pageGap = options.pageGap ?? 20;
    this.fragmentParagraphs = options.fragmentParagraphs ?? false;
    this.cancellationCheckpoint = options.checkCancellation;
    this.pageCountCheckpoint = options.checkPageCount;
    this.layoutToken = options.layoutToken;
    this.hfRegistry = new Map();
    this.footnoteRegistry = new Map();
    this.commentMarginRegistry = new Map();
  }

  /**
   * Runs the pagination process.
   *
   * @returns PaginationResult with page information
   */
  paginate(): PaginationResult {
    this.checkpoint();
    if (this.state !== "ready") {
      throw new Error(`PaginationEngine is one-shot and is already ${this.state}`);
    }
    if (!this.stagingElement.isConnected || !this.containerElement.isConnected) {
      throw new Error("Pagination requires staging and container elements attached to their document");
    }
    if (this.document.documentElement.getBoundingClientRect().width <= 0) {
      throw new Error("Pagination requires a browsing context with non-zero layout");
    }
    this.state = "running";

    try {
      const pages: PageInfo[] = [];
      let pageNumber = 1;

    // Parse the header/footer registry if present
    this.hfRegistry = this.parseHeaderFooterRegistry();

    // Parse the footnote registry if present
    this.footnoteRegistry = this.parseFootnoteRegistry();
    ({
      normal: this.footnoteSeparator,
      continuation: this.footnoteContinuationSeparator,
    } = this.parseFootnoteSeparators());

    // Parse the margin-comment registry if present. Its entries are cloned into
    // the side substrate of pages that contain the corresponding range marker.
    this.commentMarginRegistry = this.parseCommentMarginRegistry();
    this.footnoteLayoutLikeWord8 =
      this.stagingElement.dataset.footnoteLayoutLikeWord8 === "true";

    // Find all section containers
    const sections = this.stagingElement.querySelectorAll<HTMLElement>(
      "[data-section-index]"
    );

    this.pageNumbering = this.parsePageNumbering(sections);

    // If no sections found, treat the entire staging content as one section
    const sectionsToProcess =
      sections.length > 0 ? Array.from(sections) : [this.stagingElement];

    // Snapshot the addressable SOURCE inventory before flow moves nodes out of staging. PageMap
    // completeness cannot be inferred from whatever survives into page boxes: that would let a
    // dropped block silently disappear from both the DOM and the supposedly authoritative map.
    // Running-story variants and cited notes are inventoried from their source registries.
    this.expectedPageMapAnchorIds = new Set();
    const referencedFootnoteIds = new Set<string>();
    const referencedCommentIds = new Set<string>();
    for (const section of sectionsToProcess) {
      this.checkpoint();
      this.collectExpectedSourceAnchors(section, this.expectedPageMapAnchorIds, true);
      for (const reference of Array.from(section.querySelectorAll<HTMLElement>("[data-footnote-id]"))) {
        this.checkpoint();
        if (reference.closest("#pagination-footnote-registry, #pagination-hf-registry")) continue;
        const id = reference.dataset.footnoteId;
        if (id) referencedFootnoteIds.add(id);
      }
      for (const reference of Array.from(section.querySelectorAll<HTMLElement>("[data-comment-id]"))) {
        this.checkpoint();
        if (reference.closest(
          "#pagination-comment-margin-registry, #pagination-footnote-registry, #pagination-hf-registry",
        )) continue;
        const id = reference.dataset.commentId;
        if (id) referencedCommentIds.add(id);
      }
    }
    for (const id of referencedFootnoteIds) {
      this.checkpoint();
      const source = this.footnoteRegistry.get(id);
      if (!source) continue;
      this.collectExpectedSourceAnchors(source, this.expectedPageMapAnchorIds);
      // Margin comments are presentation stories selected by markers. Markers
      // inside a cited footnote are just as visible/addressable as body markers.
      for (const reference of Array.from(
        source.querySelectorAll<HTMLElement>("[data-comment-id]"),
      )) {
        this.checkpoint();
        const commentId = reference.dataset.commentId;
        if (commentId) referencedCommentIds.add(commentId);
      }
    }
    for (const id of referencedCommentIds) {
      this.checkpoint();
      const source = this.commentMarginRegistry.get(id);
      if (source) this.collectExpectedSourceAnchors(source, this.expectedPageMapAnchorIds);
    }

    // Group adjacent sections into page runs. A `w:type="continuous"` section keeps
    // filling the page its predecessor started rather than opening a fresh one, so it
    // joins the previous run — provided the page box (size and margins) is unchanged,
    // which is also Word's own condition for honoring a continuous break. Pages of a
    // merged run begins under the leading section's running stories. Each measured block retains
    // its own section, though: once a continuous section spills to a later physical page, that
    // page must switch to the new section's headers, footers, and page-numbering state.
    interface PageRun {
      sections: HTMLElement[];
      sectionIndex: number;
      sectionType: string;
      dims: PageDimensions;
      sectionDimensions: Map<number, PageDimensions>;
    }
    const runs: PageRun[] = [];
    for (const section of sectionsToProcess) {
      this.checkpoint();
      const dims = parseSectionDimensions(section);
      const previous = runs[runs.length - 1];
      if (
        previous &&
        section.dataset.sectionType === "continuous" &&
        samePageBox(previous.dims, dims)
      ) {
        previous.sections.push(section);
        previous.sectionDimensions.set(
          parseInt(section.dataset.sectionIndex || "0", 10),
          dims,
        );
      } else {
        const sectionIndex = parseInt(section.dataset.sectionIndex || "0", 10);
        runs.push({
          sections: [section],
          sectionIndex,
          sectionType: section.dataset.sectionType ?? "nextPage",
          dims,
          sectionDimensions: new Map([[sectionIndex, dims]]),
        });
      }
    }

    for (const run of runs) {
      this.checkpoint();
      // Odd/even section breaks begin the new section on the requested physical side. When the
      // following page has the opposite parity, Word inserts one intentionally blank filler page.
      // It belongs to the preceding section, advances physical/logical numbering, retains that
      // section's paper geometry, and carries no running story.
      const needsParityFiller = pages.length > 0
        && ((run.sectionType === "oddPage" && pageNumber % 2 === 0)
          || (run.sectionType === "evenPage" && pageNumber % 2 === 1));
      if (needsParityFiller) {
        const precedingPage = pages[pages.length - 1];
        const precedingPageInSection = parseInt(
          precedingPage.element.dataset.pageInSection ?? "1",
          10,
        );
        const precedingDisplayedPageNumber = parseInt(
          precedingPage.element.dataset.displayedPageNumber ?? String(precedingPage.pageNumber),
          10,
        );
        const filler = this.createPage(
          precedingPage.dimensions,
          pageNumber,
          precedingPage.sectionIndex,
          precedingDisplayedPageNumber + 1,
          [],
          precedingPageInSection + 1,
          [],
          0,
          null,
          undefined,
          true,
        );
        pages.push(filler);
        pageNumber++;
      }
      // Make staging visible for measurement
      this.stagingElement.style.visibility = "hidden";
      this.stagingElement.style.position = "absolute";
      this.stagingElement.style.left = "-9999px";
      this.stagingElement.style.display = "block";

      const blocks: MeasuredBlock[] = [];
      for (const section of run.sections) {
        this.checkpoint();
        const sectionIndex = parseInt(section.dataset.sectionIndex || "0", 10);
        const sectionDims = run.sectionDimensions.get(sectionIndex) ?? run.dims;
        // Set width for accurate line wrapping
        section.style.width = `${sectionDims.contentWidth}pt`;

        const columnCount = parseInt(section.dataset.cols || "1", 10);
        if (columnCount > 1) {
          const gap = parseFloat(section.dataset.colGap || "");
          blocks.push(...this.buildColumnBlocks(
            section,
            sectionDims,
            sectionIndex,
            columnCount,
            Number.isFinite(gap) ? gap : 36
          ));
        } else {
          blocks.push(...this.measureBlocks(section, sectionDims, sectionIndex));
        }
      }

      // Flow blocks into pages
      const sectionPages = this.flowToPages(
        blocks,
        run.dims,
        pageNumber,
        run.sectionIndex,
        run.sectionDimensions,
      );
      this.checkpoint();
      pages.push(...sectionPages);
      pageNumber += sectionPages.length;
    }

    // Hide staging after measurement
    this.stagingElement.style.display = "none";

    // Every page box exists now, so NUMPAGES has an answer and each PAGE marker knows its page.
    this.substitutePageNumberFields(pages.length);

    // Only running-story variants selected by a real page are expected to materialize. Read IDs
    // from registry sources, not presentation clones, so a failed clone remains detectable.
    for (const page of pages) {
      this.checkpoint();
      if (page.element.dataset.sectionFiller === "true") continue;
      const pageInSection = parseInt(page.element.dataset.pageInSection || "1", 10);
      const header = this.selectHeader(page.sectionIndex, pageInSection);
      const footer = this.selectFooter(page.sectionIndex, pageInSection);
      if (header) this.collectExpectedSourceAnchors(header, this.expectedPageMapAnchorIds);
      if (footer) this.collectExpectedSourceAnchors(footer, this.expectedPageMapAnchorIds);
    }

    // Establish one active editor anchor and page-qualify every presentation fragment.
    // Full canonical source identities remain on all clones, including table cells.
    this.qualifyPageFragments(pages);
    this.transferVisibleFragmentTargets();
    this.normalizeVisiblePageFragments(pages);
    this.lastPages = pages;

    const result = {
      readiness: {
        status: "ready" as const,
        pageCount: pages.length,
        diagnostics: [
          {
            code: "sections_processed" as const,
            severity: "info" as const,
            message: "Document sections processed by the paginator.",
            count: sectionsToProcess.length,
          },
          {
            code: "page_runs_processed" as const,
            severity: "info" as const,
            message: "Physical page runs processed after continuous-section grouping.",
            count: runs.length,
          },
          {
            code: "source_anchors_inventoried" as const,
            severity: "info" as const,
            message: "Canonical source anchors inventoried for PageMap completeness.",
            count: this.expectedPageMapAnchorIds.size,
          },
          {
            code: "note_references_inventoried" as const,
            severity: "info" as const,
            message: "Footnote and margin-comment references inventoried before layout.",
            count: referencedFootnoteIds.size + referencedCommentIds.size,
          },
        ],
      },
      totalPages: pages.length,
      pages,
      pageMap: this.layoutToken
        ? this.materializePageMap(
            this.layoutToken.documentVersion,
            this.layoutToken.rendererFingerprint,
          )
        : undefined,
    };
    this.state = "complete";
    return result;
    } catch (error) {
      this.state = "failed";
      throw error;
    }
  }

  private checkpoint(): void {
    this.cancellationCheckpoint?.();
  }

  /**
   * Normalize visible fragment identities after a caller applies final standalone styles.  This
   * deliberately runs before the stability barrier; materializePageMap is read-only so PageMap
   * measurement cannot mutate a tree after it was declared stable.
   */
  normalizePageMapFragmentIdentities(): void {
    if (this.lastPages.length === 0) {
      throw new Error("paginate() must complete before fragment identities can be normalized");
    }
    this.normalizeVisiblePageFragments(this.lastPages);
  }

  /**
   * Materialize the last completed browser layout as portable page-relative point geometry.
   * The caller supplies both invalidation tokens; this engine never guesses a document version
   * or renderer fingerprint.
   */
  materializePageMap(documentVersion: number, rendererFingerprint: string): PageMap {
    if (!Number.isSafeInteger(documentVersion) || documentVersion < 0) {
      throw new Error("documentVersion must be a non-negative safe integer");
    }
    if (!rendererFingerprint) throw new Error("rendererFingerprint must be non-empty");
    if (this.lastPages.length === 0) throw new Error("paginate() must complete before materializePageMap()");

    const pages: PageMapPage[] = this.lastPages.map((page) => ({
      pageNumber: page.pageNumber,
      pageInSection: parseInt(page.element.dataset.pageInSection || "1", 10),
      width: page.dimensions.pageWidth,
      height: page.dimensions.pageHeight,
      sectionIndex: page.sectionIndex,
      pageName: `docxodus-section-${page.sectionIndex}`,
    }));

    const fragments: PageMapFragment[] = [];
    const requiredAnchorIds = new Set(this.expectedPageMapAnchorIds);
    if (requiredAnchorIds.size === 0) {
      throw new Error("cannot publish an available PageMap without canonical source inventory");
    }
    const measuredAnchorIds = new Set<string>();
    const emittedFragmentCounts = new Map<string, number>();
    for (const page of this.lastPages) {
      this.checkpoint();
      const pageRect = page.element.getBoundingClientRect();
      if (pageRect.width <= 0 || pageRect.height <= 0) {
        throw new Error(`page ${page.pageNumber} has no measurable geometry`);
      }
      // Ratio-to-known-page-size removes CSS px, zoom, and transform from the contract.
      const pointPerRenderedX = page.dimensions.pageWidth / pageRect.width;
      const pointPerRenderedY = page.dimensions.pageHeight / pageRect.height;
      const nodes = page.element.querySelectorAll<HTMLElement>("[data-source-anchor-id]");
      for (const element of Array.from(nodes)) {
        this.checkpoint();
        // Preserve the source-side exclusion contract on presentation clones as well. This
        // covers the node itself and any excluded/hidden/aria-hidden ancestor within the page.
        if (this.isDeliberatelyUnrenderedSource(element, page.element)) continue;
        const anchorId = element.dataset.sourceAnchorId;
        if (!anchorId || !element.dataset.pageFragmentId
          || !Number.isInteger(parseInt(element.dataset.fragmentIndex || "", 10))) {
          throw new Error(`page ${page.pageNumber} contains an unqualified source anchor`);
        }

        const rect = element.getBoundingClientRect();
        const style = this.view.getComputedStyle(element);
        const deliberatelyHidden = style.display === "none" || style.visibility === "hidden";
        if (deliberatelyHidden) continue;
        requiredAnchorIds.add(anchorId);
        const visibleRect = this.intersectWithClippingAncestors(element, page.element, pageRect, rect);
        const left = visibleRect.left;
        const top = visibleRect.top;
        const right = visibleRect.right;
        const bottom = visibleRect.bottom;
        if (rect.width <= 0 || rect.height <= 0 || right <= left || bottom <= top) {
          // A continued note/story clone can contain children clipped off this page which become
          // measurable on its next clone. Enforce completeness once every page has been inspected.
          continue;
        }

        measuredAnchorIds.add(anchorId);
        // Clipped descendants in repeated note/story clones are deliberately omitted from the
        // portable map. Re-number only the visible fragments so the emitted contract remains
        // contiguous even when an earlier DOM clone carried no visible geometry on its page.
        const fragmentIndex = emittedFragmentCounts.get(anchorId) ?? 0;
        emittedFragmentCounts.set(anchorId, fragmentIndex + 1);
        const fragmentId = `p${page.pageNumber}-f${fragmentIndex}-${anchorId}`;
        if (element.dataset.fragmentIndex !== String(fragmentIndex)
          || element.dataset.pageFragmentId !== fragmentId
          || element.dataset.pageNumber !== String(page.pageNumber)) {
          throw new Error(
            `page ${page.pageNumber} fragment identity changed after final-tree normalization`,
          );
        }
        fragments.push({
          fragmentId,
          anchorId,
          fragmentIndex,
          pageNumber: page.pageNumber,
          geometry: {
            x: (left - pageRect.left) * pointPerRenderedX,
            y: (top - pageRect.top) * pointPerRenderedY,
            width: (right - left) * pointPerRenderedX,
            height: (bottom - top) * pointPerRenderedY,
          },
          story: this.storyForCanonicalAnchor(anchorId),
          inTableCell: element.matches("td,th") || element.closest("td,th") !== null,
        });
      }
    }

    const missingAnchor = Array.from(requiredAnchorIds).find((id) => !measuredAnchorIds.has(id));
    if (missingAnchor) {
      throw new Error(`source anchor ${missingAnchor} has no measurable fragment in the paginated layout`);
    }

    return {
      schemaVersion: 1,
      mode: "paginated",
      availability: "available",
      documentVersion,
      rendererFingerprint,
      pages,
      fragments,
    };
  }

  private normalizeVisiblePageFragments(pages: PageInfo[]): void {
    const emittedFragmentCounts = new Map<string, number>();
    for (const page of pages) {
      this.checkpoint();
      const pageRect = page.element.getBoundingClientRect();
      for (const element of Array.from(
        page.element.querySelectorAll<HTMLElement>("[data-source-anchor-id]"),
      )) {
        this.checkpoint();
        if (this.isDeliberatelyUnrenderedSource(element, page.element)) continue;
        const anchorId = element.dataset.sourceAnchorId;
        if (!anchorId) continue;
        const style = this.view.getComputedStyle(element);
        if (style.display === "none" || style.visibility === "hidden") continue;
        const rect = element.getBoundingClientRect();
        const visible = this.intersectWithClippingAncestors(element, page.element, pageRect, rect);
        if (rect.width <= 0 || rect.height <= 0
          || visible.right <= visible.left || visible.bottom <= visible.top) continue;
        const fragmentIndex = emittedFragmentCounts.get(anchorId) ?? 0;
        emittedFragmentCounts.set(anchorId, fragmentIndex + 1);
        element.dataset.pageNumber = String(page.pageNumber);
        element.dataset.fragmentIndex = String(fragmentIndex);
        element.dataset.pageFragmentId = `p${page.pageNumber}-f${fragmentIndex}-${anchorId}`;
      }
    }
  }

  /**
   * Intersect an element with every ancestor that establishes an overflow clip before the page
   * root. getBoundingClientRect() reports layout outside those clips, which is not rendered and
   * therefore must not satisfy PageMap completeness or inflate portable geometry.
   */
  private intersectWithClippingAncestors(
    element: HTMLElement,
    page: HTMLElement,
    pageRect: DOMRect,
    rect: DOMRect,
  ): { left: number; top: number; right: number; bottom: number } {
    let left = Math.max(rect.left, pageRect.left);
    let top = Math.max(rect.top, pageRect.top);
    let right = Math.min(rect.right, pageRect.right);
    let bottom = Math.min(rect.bottom, pageRect.bottom);
    const clips = (value: string) =>
      value === "hidden" || value === "clip" || value === "scroll" || value === "auto";

    for (let ancestor = element.parentElement;
      ancestor && ancestor !== page;
      ancestor = ancestor.parentElement) {
      const style = this.view.getComputedStyle(ancestor);
      const clipsX = clips(style.overflowX);
      const clipsY = clips(style.overflowY);
      if (!clipsX && !clipsY) continue;
      const ancestorRect = ancestor.getBoundingClientRect();
      if (clipsX) {
        left = Math.max(left, ancestorRect.left);
        right = Math.min(right, ancestorRect.right);
      }
      if (clipsY) {
        top = Math.max(top, ancestorRect.top);
        bottom = Math.min(bottom, ancestorRect.bottom);
      }
    }
    return { left, top, right, bottom };
  }

  private storyForCanonicalAnchor(anchorId: string): PageMapStory {
    const first = anchorId.indexOf(":");
    const second = first < 0 ? -1 : anchorId.indexOf(":", first + 1);
    const scope = first >= 0 && second > first ? anchorId.slice(first + 1, second) : "body";
    if (scope.startsWith("hdr")) return "header";
    if (scope.startsWith("ftr")) return "footer";
    if (scope === "fn") return "footnote";
    if (scope === "en") return "endnote";
    if (scope === "cmt") return "comment";
    return "body";
  }

  /**
   * Add canonical IDs from an addressable source subtree to the pre-pagination inventory.
   * Registry wrappers are excluded when scanning staging because selectable registry contents are
   * inventoried separately. Producers may explicitly mark content that has no visual substrate
   * with `data-page-map-exclude="true"`; native hidden semantics carry the same signal.
   */
  private collectExpectedSourceAnchors(
    source: HTMLElement,
    destination: Set<string>,
    excludeRegistries = false,
  ): void {
    const candidates: HTMLElement[] = source.matches("[data-source-anchor-id]") ? [source] : [];
    candidates.push(...Array.from(source.querySelectorAll<HTMLElement>("[data-source-anchor-id]")));
    for (const element of candidates) {
      if (excludeRegistries && element.closest(
        "#pagination-hf-registry, #pagination-footnote-registry, #pagination-comment-margin-registry",
      )) {
        continue;
      }
      if (this.isZeroHeightExplicitBreakCarrier(element)) {
        // Flow consumes the following break marker and clones this otherwise-empty paragraph.
        // Persist the exclusion so that the clone cannot reintroduce the anchor during PageMap
        // measurement after the marker itself has disappeared.
        element.dataset.pageMapExclude = "true";
      }
      if (this.isDeliberatelyUnrenderedSource(element, source)) continue;
      const anchorId = element.dataset.sourceAnchorId;
      if (anchorId) destination.add(anchorId);
    }
  }

  private isZeroHeightExplicitBreakCarrier(element: HTMLElement): boolean {
    const next = element.nextElementSibling as HTMLElement | null;
    const bounds = element.getBoundingClientRect();
    return element.hasAttribute("data-source-anchor-id")
      && (element.textContent ?? "").replace(/\u00a0/g, "").trim() === ""
      // Exclude only the converter's zero-height carrier. A blank paragraph with height from
      // padding, borders, a background, or authored sizing still has a visible page substrate and
      // must remain addressable under the authoritative PageMap contract.
      && bounds.height <= 0.01
      && !element.querySelector("[data-source-anchor-id]")
      && !element.querySelector("img,svg,canvas,table,hr,input,textarea,select")
      && (next?.dataset.pageBreak === "true"
        || next?.classList.contains(`${this.cssPrefix}break`) === true);
  }

  private isDeliberatelyUnrenderedSource(element: HTMLElement, sourceRoot: HTMLElement): boolean {
    for (let current: HTMLElement | null = element; current; current = current.parentElement) {
      const zeroHeightExplicitBreakCarrier = current === element
        && this.isZeroHeightExplicitBreakCarrier(current);
      if (
        current.dataset.pageMapExclude === "true"
        // An explicit page-break marker controls flow but has no painted substrate. The converter
        // intentionally emits it as an empty div, so requiring point geometry for its source
        // identity would make every otherwise-valid document with w:br[type=page] fail closed.
        || current.dataset.pageBreak === "true"
        || current.classList.contains(`${this.cssPrefix}break`)
        || zeroHeightExplicitBreakCarrier
        || current.hidden
        || current.getAttribute("aria-hidden") === "true"
        || current.style.display === "none"
        || current.style.visibility === "hidden"
      ) {
        return true;
      }
      if (current === sourceRoot) break;
    }
    return false;
  }

  /**
   * Keep exactly one active bare-Unid editor anchor per source block. Presentation clones use
   * canonical source identity plus page/fragment qualification instead.
   */
  private qualifyPageFragments(pages: PageInfo[]): void {
    const fragmentCounts = new Map<string, number>();
    const activeCanonicalIds = new Set<string>();
    const activeBareAnchorIds = new Set<string>();

    const makeInactive = (element: HTMLElement): void => {
      element.removeAttribute("data-anchor");
      element.removeAttribute("data-committed-text");
      if (element.hasAttribute("contenteditable")) element.setAttribute("contenteditable", "false");
    };

    for (const page of pages) {
      const nodes = page.element.querySelectorAll<HTMLElement>("[data-source-anchor-id]");
      for (const element of Array.from(nodes)) {
        const anchorId = element.dataset.sourceAnchorId;
        if (!anchorId) continue;
        const fragmentIndex = fragmentCounts.get(anchorId) ?? 0;
        fragmentCounts.set(anchorId, fragmentIndex + 1);
        element.dataset.pageNumber = String(page.pageNumber);
        element.dataset.fragmentIndex = String(fragmentIndex);
        element.dataset.pageFragmentId = `p${page.pageNumber}-f${fragmentIndex}-${anchorId}`;

        const story = this.storyForCanonicalAnchor(anchorId);
        const mayOwnActiveEditorAnchor =
          story === "body" || story === "comment" || story === "footnote" || story === "endnote";
        if (element.hasAttribute("data-anchor")) {
          const bareAnchorId = element.dataset.anchor!;
          if (mayOwnActiveEditorAnchor
            && !activeCanonicalIds.has(anchorId)
            && !activeBareAnchorIds.has(bareAnchorId)) {
            activeCanonicalIds.add(anchorId);
            activeBareAnchorIds.add(bareAnchorId);
          } else {
            makeInactive(element);
          }
        }
      }
    }

    // Body/comment page nodes are the editable copies. A repeated header/footer registry entry can
    // also render the same source story once per section/variant, so retain at most one active
    // staging node per canonical source and make every presentation duplicate inert.
    const activeStagingCanonicalIds = new Set<string>();
    for (const element of Array.from(
      this.stagingElement.querySelectorAll<HTMLElement>("[data-source-anchor-id][data-anchor]"),
    )) {
      const anchorId = element.dataset.sourceAnchorId;
      if (!anchorId) continue;
      const bareAnchorId = element.dataset.anchor!;
      if (activeCanonicalIds.has(anchorId)
        || activeStagingCanonicalIds.has(anchorId)
        || activeBareAnchorIds.has(bareAnchorId)) {
        makeInactive(element);
      } else {
        activeStagingCanonicalIds.add(anchorId);
        activeBareAnchorIds.add(bareAnchorId);
      }
    }
  }

  /**
   * Page flow clones source blocks while the hidden staging tree stays in the document. Any HTML
   * fragment target copied into a visible page would therefore resolve to its earlier hidden
   * source. Transfer target ownership to the page presentation after flow is complete; registry
   * and wrapper IDs that have no visible counterpart remain available to pagination internals.
   */
  private transferVisibleFragmentTargets(): void {
    const visibleIds = new Set(Array.from(
      this.containerElement.querySelectorAll<HTMLElement>("[id]"),
    ).map((element) => element.id).filter(Boolean));
    if (visibleIds.size === 0) return;
    for (const source of Array.from(
      this.stagingElement.querySelectorAll<HTMLElement>("[id]"),
    )) {
      if (visibleIds.has(source.id)) source.removeAttribute("id");
    }
  }

  /** Read each section's `w:pgNumType` off its wrapper (see {@link SectionPageNumbering}). */
  private parsePageNumbering(sections: ArrayLike<HTMLElement>): Map<number, SectionPageNumbering> {
    const map = new Map<number, SectionPageNumbering>();
    for (const section of Array.from(sections)) {
      const index = parseInt(section.dataset.sectionIndex || "0", 10);
      const rawStart = section.dataset.pageNumStart;
      const start = rawStart === undefined ? undefined : parseInt(rawStart, 10);
      map.set(index, {
        start: start !== undefined && Number.isFinite(start) ? start : undefined,
        format: section.dataset.pageNumFmt,
      });
    }
    return map;
  }

  /**
   * Fill in the page-number fields inside every page's cloned header/footer.
   *
   * A header/footer is authored once and cloned onto each page, so a PAGE field's single cached
   * result would otherwise show the same number on every page — the whole reason the converter
   * marks these. `data-field-format` (the field's own `\*` switch) wins over the section's format
   * when present, which is exactly how Word resolves the two.
   *
   * Runs after layout because NUMPAGES cannot be known before the last page exists. The
   * substituted text can therefore be marginally wider than the cached result the header was
   * measured with; the header band clips, so the failure mode is a hair of overflow rather than
   * a layout that disagrees with itself.
   *
   * Scoped to the CLONED header/footer regions on purpose. A page-number field in body text is
   * ordinary run content that the editor may make editable, and committing an edited block writes
   * back whatever text the DOM holds — rewriting it here would mean a body field commits a number
   * the document never contained. Body content is also not cloned, so it does not have the problem
   * this method exists to solve.
   */
  private substitutePageNumberFields(totalPages: number): void {
    const boxes = this.containerElement.querySelectorAll<HTMLElement>(`.${this.cssPrefix}box`);
    for (const box of Array.from(boxes)) {
      const markers = box.querySelectorAll<HTMLElement>(
        `.${this.cssPrefix}header [data-field], .${this.cssPrefix}footer [data-field]`,
      );
      if (markers.length === 0) continue;

      const sectionIndex = parseInt(box.dataset.sectionIndex || "0", 10);
      const pageNumber = parseInt(box.dataset.pageNumber || "1", 10);
      const numbering = this.pageNumbering.get(sectionIndex) ?? {};
      const displayed = parseInt(box.dataset.displayedPageNumber || String(pageNumber), 10);

      for (const marker of Array.from(markers)) {
        const kind = marker.dataset.field;
        if (kind !== "PAGE" && kind !== "NUMPAGES") continue;
        const format = marker.dataset.fieldFormat ?? numbering.format;
        marker.textContent = formatPageNumber(kind === "PAGE" ? displayed : totalPages, format);
      }
    }
  }

  /**
   * Measures all content blocks in a section.
   */
  private measureBlocks(
    section: HTMLElement,
    dims: PageDimensions,
    sectionIndex: number,
  ): MeasuredBlock[] {
    const blocks: MeasuredBlock[] = [];

    // Get direct children (paragraphs, tables, divs, etc.)
    const children = Array.from(section.children) as HTMLElement[];

    for (const child of children) {
      this.checkpoint();
      // Skip section dividers that are just wrappers
      if (child.dataset.sectionIndex !== undefined) {
        // Recursively get blocks from nested sections
        const nestedSectionIndex = parseInt(child.dataset.sectionIndex || String(sectionIndex), 10);
        const nestedBlocks = this.measureBlocks(child, dims, nestedSectionIndex);
        blocks.push(...nestedBlocks);
        continue;
      }

      // Converter-shaped endnotes are a safe nested block structure whose outer
      // section/list wrappers must not make the complete endnote collection one
      // indivisible page block. Flatten only the exact shape we understand; any
      // richer author HTML retains the conservative whole-block fallback below.
      if (child.matches("section.endnotes")) {
        const endnoteBlocks = this.measureSafeEndnoteBlocks(child, dims, sectionIndex);
        if (endnoteBlocks) {
          blocks.push(...endnoteBlocks);
          continue;
        }
      }

      // Measure height and margins separately for proper margin collapsing calculation
      // getBoundingClientRect() returns content+padding+border, not margins
      const rect = child.getBoundingClientRect();
      const style = this.view.getComputedStyle(child);
      const marginTopPx = parseFloat(style.marginTop) || 0;
      const marginBottomPx = parseFloat(style.marginBottom) || 0;
      const heightPt = pxToPt(rect.height);
      const marginTopPt = pxToPt(marginTopPx);
      const marginBottomPt = pxToPt(marginBottomPx);

      const isPageBreak =
        child.dataset.pageBreak === "true" ||
        child.classList.contains(`${this.cssPrefix}break`);

      blocks.push({
        element: child,
        sectionIndex,
        heightPt,
        marginTopPt,
        marginBottomPt,
        keepWithNext: child.dataset.keepWithNext === "true",
        keepLines: child.dataset.keepLines === "true",
        pageBreakBefore: child.dataset.pageBreakBefore === "true",
        isPageBreak,
        isWordParagraph: this.isWordParagraphElement(child),
      });
    }

    return blocks;
  }

  /**
   * Flatten the converter's `section.endnotes > ol > li > p` presentation into
   * ordinary paragraph blocks. This preserves paragraph formatting and canonical
   * p:en/en:en identities while allowing the existing paragraph fragmenter to
   * split a long endnote across page boundaries.
   */
  private measureSafeEndnoteBlocks(
    section: HTMLElement,
    dims: PageDimensions,
    sectionIndex: number,
  ): MeasuredBlock[] | null {
    const sectionChildren = Array.from(section.children) as HTMLElement[];
    const list = sectionChildren.find((child) => child.tagName === "OL");
    if (
      !list
      || sectionChildren.some((child) => child.tagName !== "HR" && child !== list)
      || Array.from(list.children).some((child) => child.tagName !== "LI")
    ) {
      return null;
    }

    const items = Array.from(list.children) as HTMLElement[];
    if (items.length === 0 || items.some((item) =>
      item.children.length === 0
      || Array.from(item.children).some((child) => child.tagName !== "P")
    )) {
      return null;
    }

    const blocks: MeasuredBlock[] = [];
    const sectionStyle = this.view.getComputedStyle(section);
    for (const rule of sectionChildren.filter((child) => child.tagName === "HR")) {
      this.checkpoint();
      const clonedRule = rule.cloneNode(true) as HTMLElement;
      if (blocks.length === 0) clonedRule.style.marginTop = sectionStyle.marginTop;
      blocks.push(this.measureElement(clonedRule, dims, sectionIndex));
    }

    const listStyle = this.view.getComputedStyle(list);
    for (let itemIndex = 0; itemIndex < items.length; itemIndex++) {
      this.checkpoint();
      const item = items[itemIndex];
      const ownerAnchorId = item.dataset.sourceAnchorId;
      const paragraphs = Array.from(item.children) as HTMLElement[];
      for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
        this.checkpoint();
        const paragraph = paragraphs[paragraphIndex];
        // The flattened clone no longer has the section/ol/li ancestors that supplied the
        // source's computed layout. Validate while the real paragraph is still attached; the
        // marker below records that completed check for canFragmentParagraph(), whose detached
        // clone cannot obtain meaningful computed styles. A richer custom endnote falls back to
        // the established indivisible section path instead of being range-split incorrectly.
        if (!this.hasRangeFragmentSafeLayout(paragraph)) {
          return null;
        }

        const clone = paragraph.cloneNode(true) as HTMLElement;
        clone.dataset.paginationSafeEndnote = "true";
        clone.style.fontSize ||= sectionStyle.fontSize;
        clone.style.lineHeight ||= sectionStyle.lineHeight;
        clone.style.paddingLeft ||= listStyle.paddingLeft;

        // Older/custom producers may put the endnote identity only on the li.
        // Mirror it into the visible paragraph without displacing the paragraph's
        // own identity, matching current converter output.
        if (ownerAnchorId && !Array.from(
          clone.querySelectorAll<HTMLElement>("[data-source-anchor-id]"),
        ).some((node) => node.dataset.sourceAnchorId === ownerAnchorId)) {
          const owner = this.document.createElement("span");
          owner.dataset.sourceAnchorId = ownerAnchorId;
          while (clone.firstChild) owner.appendChild(clone.firstChild);
          clone.appendChild(owner);
        }

        if (paragraphIndex === 0) {
          // Flattening must preserve both ends of the converter's endnote link and the list's
          // numbering format. The outer id survives only on the leading range fragment; normal
          // continuation cleanup removes it from every later fragment.
          const itemId = item.id;
          if (itemId) {
            if (clone.id && clone.id !== itemId) return null;
            clone.id = itemId;
          }
          const value = parseInt(item.getAttribute("value") || String(itemIndex + 1), 10);
          const marker = this.formatOrderedListMarker(
            Number.isFinite(value) ? value : itemIndex + 1,
            listStyle.listStyleType,
          );
          clone.insertBefore(this.document.createTextNode(`${marker}. `), clone.firstChild);
        }
        blocks.push(this.measureElement(clone, dims, sectionIndex));
      }
    }

    return blocks;
  }

  /** Render the CSS ordered-list formats emitted by the converter after an endnote is flattened. */
  private formatOrderedListMarker(value: number, listStyleType: string): string {
    const format = (() => {
      switch (listStyleType) {
        case "lower-roman": return "lowerRoman";
        case "upper-roman": return "upperRoman";
        case "lower-alpha":
        case "lower-latin": return "lowerLetter";
        case "upper-alpha":
        case "upper-latin": return "upperLetter";
        default: return "decimal";
      }
    })();
    return formatPageNumber(value, format);
  }

  /**
   * Flows a multi-column (`w:cols`) section's children into CSS-multicol container
   * blocks. Word lays such a section out as N columns inside the same body extent;
   * a balanced `column-count` container reproduces that geometry, and the paginator
   * then places each container as one ordinary measured block. A container grows
   * greedily until its balanced height would exceed the smallest page body available
   * to the section, so a long columned section still splits across pages at block
   * boundaries. Each child lands in exactly one container, so anchors never
   * duplicate. An explicit page break child passes through as its own block, which
   * ends the current container and lets the normal flow logic turn the page.
   */
  private buildColumnBlocks(
    section: HTMLElement,
    dims: PageDimensions,
    sectionIndex: number,
    columnCount: number,
    columnGapPt: number
  ): MeasuredBlock[] {
    const children = Array.from(section.children) as HTMLElement[];
    const blocks: MeasuredBlock[] = [];
    const maxFragmentHeight = this.smallestEffectiveContentHeight(dims, sectionIndex);

    const isBreak = (child: HTMLElement) =>
      child.dataset.pageBreak === "true" ||
      child.classList.contains(`${this.cssPrefix}break`);

    const makeContainer = (slice: HTMLElement[]): HTMLElement => {
      const container = this.document.createElement("div");
      container.style.columnCount = String(columnCount);
      container.style.columnGap = `${columnGapPt}pt`;
      for (const child of slice) {
        container.appendChild(child.cloneNode(true));
      }
      return container;
    };

    let start = 0;
    while (start < children.length) {
      this.checkpoint();
      if (isBreak(children[start])) {
        blocks.push(this.measureElement(children[start], dims, sectionIndex));
        start++;
        continue;
      }

      // Grow the container one child at a time. Even a single oversized child is
      // emitted alone, preserving the established oversized-block fallback.
      let end = start + 1;
      let container = makeContainer(children.slice(start, end));
      let measured = this.measureElement(container, dims, sectionIndex);
      while (end < children.length && !isBreak(children[end])) {
        this.checkpoint();
        const candidate = makeContainer(children.slice(start, end + 1));
        const candidateMeasured = this.measureElement(candidate, dims, sectionIndex);
        if (candidateMeasured.heightPt > maxFragmentHeight) break;
        container = candidate;
        measured = candidateMeasured;
        end++;
      }

      blocks.push({
        element: container,
        sectionIndex,
        heightPt: measured.heightPt,
        marginTopPt: measured.marginTopPt,
        marginBottomPt: measured.marginBottomPt,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
        isWordParagraph: false,
      });
      start = end;
    }

    return blocks;
  }

  /**
   * Measures one element in the same hidden staging context used for the source blocks.
   * This is intentionally DOM-based: table row heights cannot be inferred from individual
   * rows because wrapping and collapsed borders change the height of a fragment.
   */
  private measureElement(
    element: HTMLElement,
    dims: PageDimensions,
    sectionIndex: number,
  ): MeasuredBlock {
    const measurementHost = this.document.createElement("div");
    measurementHost.style.position = "absolute";
    measurementHost.style.visibility = "hidden";
    measurementHost.style.left = "-9999px";
    measurementHost.style.width = `${dims.contentWidth}pt`;

    const measuredElement = element.cloneNode(true) as HTMLElement;
    measurementHost.appendChild(measuredElement);
    this.stagingElement.appendChild(measurementHost);

    const rect = measuredElement.getBoundingClientRect();
    const style = this.view.getComputedStyle(measuredElement);
    const measured: MeasuredBlock = {
      element,
      sectionIndex,
      heightPt: pxToPt(rect.height),
      marginTopPt: pxToPt(parseFloat(style.marginTop) || 0),
      marginBottomPt: pxToPt(parseFloat(style.marginBottom) || 0),
      keepWithNext: element.dataset.keepWithNext === "true",
      keepLines: element.dataset.keepLines === "true",
      pageBreakBefore: element.dataset.pageBreakBefore === "true",
      isPageBreak:
        element.dataset.pageBreak === "true" ||
        element.classList.contains(`${this.cssPrefix}break`),
      isWordParagraph: this.isWordParagraphElement(element),
    };

    this.stagingElement.removeChild(measurementHost);
    return measured;
  }

  /**
   * Returns the contiguous keep-with-next chain beginning at a block.
   *
   * A hard page break or a page-break-before directive is stronger than a
   * keep-with-next directive, so it terminates the chain. The caller only
   * keeps a chain together when the whole chain can fit on a fresh page.
   */
  private getKeepWithNextChain(blocks: MeasuredBlock[], startIndex: number): MeasuredBlock[] {
    const firstBlock = blocks[startIndex];
    if (!firstBlock) return [];

    const chain = [firstBlock];
    let lastIndex = startIndex;

    while (blocks[lastIndex].keepWithNext) {
      const nextBlock = blocks[lastIndex + 1];
      if (!nextBlock || nextBlock.isPageBreak || nextBlock.pageBreakBefore) {
        break;
      }

      chain.push(nextBlock);
      lastIndex++;
    }

    return chain;
  }

  /**
   * Measures the visible body height of a keep-with-next chain using the same
   * collapsed-margin rules as normal block placement. The trailing margin is
   * intentionally excluded, matching the individual block fit check.
   */
  private measureKeepWithNextChainBodyHeight(
    chain: MeasuredBlock[],
    previousMarginBottomPt: number,
    isFirstOnPage: boolean,
    pageInSection: number
  ): number {
    const firstBlock = chain[0];
    if (!firstBlock) return 0;

    const firstMarginTop = this.effectiveBlockMarginTop(
      firstBlock,
      previousMarginBottomPt,
      isFirstOnPage,
      pageInSection
    );
    let bodyHeight = firstMarginTop + firstBlock.heightPt;

    for (let index = 1; index < chain.length; index++) {
      const previousBlock = chain[index - 1];
      const block = chain[index];
      bodyHeight += Math.max(previousBlock.marginBottomPt, block.marginTopPt) + block.heightPt;
    }

    return bodyHeight;
  }

  /** Word paragraphs render as `p`, or as `h1`–`h6` when their style has an outline level. */
  private isWordParagraphElement(element: HTMLElement): boolean {
    return element.tagName === "P" || /^H[1-6]$/.test(element.tagName);
  }

  private shouldSuppressPageTopSpacing(
    block: MeasuredBlock,
    isFirstOnPage: boolean,
    pageInSection: number
  ): boolean {
    return isFirstOnPage && pageInSection > 1 && block.isWordParagraph === true;
  }

  /**
   * Resolves the part of a block's top margin that consumes the current page.
   *
   * In Word's native DOCX layout, paragraph space-before is suppressed when a paragraph is the
   * first body block on a later page of the SAME section. The first page of a document/section is
   * the exception and keeps its spacing. Tables and other block margins are not paragraph spacing,
   * so they continue to use the ordinary CSS collapsing rule.
   */
  private effectiveBlockMarginTop(
    block: MeasuredBlock,
    previousMarginBottomPt: number,
    isFirstOnPage: boolean,
    pageInSection: number
  ): number {
    if (this.shouldSuppressPageTopSpacing(block, isFirstOnPage, pageInSection)) {
      return 0;
    }
    return isFirstOnPage
      ? block.marginTopPt
      : Math.max(block.marginTopPt, previousMarginBottomPt) - previousMarginBottomPt;
  }

  /** Clone a source block with the same page-top spacing decision used by the height budget. */
  private cloneBlockForPage(
    block: MeasuredBlock,
    isFirstOnPage: boolean,
    pageInSection: number
  ): HTMLElement {
    const clone = block.element.cloneNode(true) as HTMLElement;
    if (this.shouldSuppressPageTopSpacing(block, isFirstOnPage, pageInSection)) {
      clone.style.setProperty("margin-top", "0", "important");
    }
    return clone;
  }

  /**
   * Finds footnote references introduced by a sequence of blocks, preserving
   * document order and excluding references already assigned to the page.
   */
  private collectNewFootnoteIds(
    blocks: MeasuredBlock[],
    existingFootnoteIds: string[]
  ): string[] {
    const knownIds = new Set(existingFootnoteIds);
    const newIds: string[] = [];

    for (const block of blocks) {
      this.checkpoint();
      for (const id of this.extractFootnoteRefs(block.element)) {
        if (!knownIds.has(id)) {
          knownIds.add(id);
          newIds.push(id);
        }
      }
    }

    return newIds;
  }

  /**
   * The shortest body available to a section's first, default, or even page.
   * A row fragment must fit every variant, otherwise a later header/footer could
   * send it through the oversized-block fallback again.
   */
  private smallestEffectiveContentHeight(dims: PageDimensions, sectionIndex: number): number {
    return Math.min(
      this.getPageBands(dims, sectionIndex, 1).bodyHeight,
      this.getPageBands(dims, sectionIndex, 2).bodyHeight,
      this.getPageBands(dims, sectionIndex, 3).bodyHeight
    );
  }

  /**
   * Builds a clone of a simple table wrapper containing a contiguous run of rows.
   * Complex table features are deliberately rejected by the caller: a split across
   * merged cells, nested tables, or footnotes cannot be made correct by cloning rows.
   */
  private createSimpleTableFragment(
    wrapper: HTMLElement,
    table: HTMLTableElement,
    body: HTMLTableSectionElement,
    rows: HTMLTableRowElement[],
    retainAnchor: boolean
  ): HTMLElement {
    const wrapperClone = wrapper.cloneNode(false) as HTMLElement;
    const tableClone = table.cloneNode(false) as HTMLTableElement;

    // The eligibility gate permits only colgroups alongside the body. Keep each
    // colgroup so fixed and proportional column widths remain stable per fragment.
    for (const child of Array.from(table.children)) {
      if (child !== body) {
        tableClone.appendChild(child.cloneNode(true));
      }
    }

    const bodyClone = body.cloneNode(false) as HTMLTableSectionElement;
    for (const row of rows) {
      bodyClone.appendChild(row.cloneNode(true));
    }
    tableClone.appendChild(bodyClone);

    if (!retainAnchor) {
      wrapperClone.removeAttribute("data-anchor");
      tableClone.removeAttribute("data-anchor");
    }

    wrapperClone.appendChild(tableClone);
    return wrapperClone;
  }

  /**
   * Splits an oversized, ordinary table at row boundaries. This only participates
   * in the existing oversized-block fallback; unsupported tables keep the previous
   * overflow behavior rather than risking broken table semantics.
   */
  private trySplitSimpleOversizedTable(
    block: MeasuredBlock,
    dims: PageDimensions,
    sectionIndex: number
  ): MeasuredBlock[] | null {
    const wrapper = block.element;
    if (
      wrapper.tagName !== "DIV" ||
      wrapper.children.length !== 1 ||
      block.keepWithNext ||
      block.keepLines ||
      block.pageBreakBefore ||
      block.isPageBreak
    ) {
      return null;
    }

    const tableElement = wrapper.firstElementChild;
    if (!tableElement || tableElement.localName !== "table") {
      return null;
    }
    const table = tableElement as HTMLTableElement;

    const body = table.tBodies.length === 1 ? table.tBodies[0] : null;
    if (
      !body ||
      table.tHead ||
      table.tFoot ||
      body.rows.length < 2 ||
      Array.from(table.children).some(child => child !== body && child.tagName !== "COLGROUP") ||
      table.querySelector("table, [rowspan], [colspan], [data-footnote-id]") ||
      wrapper.querySelector("[data-footnote-id]")
    ) {
      return null;
    }

    const rows = Array.from(body.rows);
    const minimumContentHeight = this.smallestEffectiveContentHeight(dims, sectionIndex);
    // Use the source's full vertical margins while forming groups. Continuation
    // fragments later clear their joining margins, so this conservative bound
    // cannot create a fragment that overflows a header/footer variant.
    const maximumFragmentHeight = minimumContentHeight - block.marginTopPt - block.marginBottomPt;
    if (maximumFragmentHeight <= 0) {
      return null;
    }

    const groups: HTMLTableRowElement[][] = [];
    let start = 0;
    while (start < rows.length) {
      this.checkpoint();
      let end = start;
      while (end < rows.length) {
        this.checkpoint();
        const candidate = this.createSimpleTableFragment(
          wrapper,
          table,
          body,
          rows.slice(start, end + 1),
          start === 0
        );
        const measured = this.measureElement(candidate, dims, block.sectionIndex);
        if (measured.heightPt > maximumFragmentHeight) {
          break;
        }
        end++;
      }

      // Even a one-row fragment cannot fit. Preserve the established overflow
      // fallback rather than looping or clipping a partially split row.
      if (end === start) {
        return null;
      }

      groups.push(rows.slice(start, end));
      start = end;
    }

    if (groups.length < 2) {
      return null;
    }

    const fragments: MeasuredBlock[] = [];
    for (let index = 0; index < groups.length; index++) {
      const isFirst = index === 0;
      const isLast = index === groups.length - 1;
      const fragment = this.createSimpleTableFragment(
        wrapper,
        table,
        body,
        groups[index],
        isFirst
      );

      // Keep the source's outer spacing only at the table's real boundaries.
      // Continuation margins would otherwise add blank space at the top/bottom
      // of every paginated fragment.
      if (!isFirst) {
        fragment.style.setProperty("margin-top", "0", "important");
      }
      if (!isLast) {
        fragment.style.setProperty("margin-bottom", "0", "important");
      }

      const measured = this.measureElement(fragment, dims, block.sectionIndex);
      if (
        measured.heightPt + measured.marginTopPt + measured.marginBottomPt >
        minimumContentHeight
      ) {
        return null;
      }

      fragments.push({
        ...measured,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
      });
    }

    return fragments;
  }

  /**
   * DOM endpoints that can finish a paragraph fragment. The flattened UTF-16
   * offsets are checked against the browser's Unicode grapheme segmenter, so a
   * formatting-run boundary can never bisect a surrogate pair, combining
   * sequence, or joined emoji. NBSP/word-joiner boundaries remain indivisible.
   * PAGE/NUMPAGES field results are atomic because splitting their marker would
   * make later substitution duplicate or replace only half of the field.
   */
  private paragraphFragmentEndpoints(
    paragraph: HTMLElement,
    emergencyGraphemeBreaks = false,
  ): ParagraphFragmentEndpoint[] {
    const textNodes: Text[] = [];
    const textStarts = new Map<Text, number>();
    const lastTextByField = new Map<HTMLElement, Text>();
    const walker = this.document.createTreeWalker(paragraph, this.view.NodeFilter.SHOW_TEXT);
    const flattenedChunks: string[] = [];
    let flattenedLength = 0;
    let textNode: Text | null;
    while ((textNode = walker.nextNode() as Text | null)) {
      this.checkpoint();
      textStarts.set(textNode, flattenedLength);
      textNodes.push(textNode);
      flattenedChunks.push(textNode.data);
      flattenedLength += textNode.data.length;
      let field = textNode.parentElement?.closest<HTMLElement>("[data-field]") ?? null;
      while (field?.parentElement?.closest<HTMLElement>("[data-field]")) {
        field = field.parentElement.closest<HTMLElement>("[data-field]");
      }
      if (field && paragraph.contains(field)) lastTextByField.set(field, textNode);
    }
    const flattenedText = flattenedChunks.join("");
    if (!this.hasVisibleText(flattenedText)) return [];

    const invisible = /[\s\u200B-\u200F\uFEFF]/;
    let firstVisibleOffset = -1;
    let lastVisibleOffset = -1;
    for (let index = 0; index < flattenedText.length; index++) {
      if (index % 4096 === 0) this.checkpoint();
      if (invisible.test(flattenedText[index])) continue;
      if (firstVisibleOffset < 0) firstVisibleOffset = index;
      lastVisibleOffset = index;
    }

    const graphemeBoundaries = this.graphemeBoundaryOffsets(flattenedText);
    const candidates: ParagraphFragmentEndpoint[] = [];
    const addCandidate = (candidate: ParagraphFragmentEndpoint): void => {
      const { textOffset } = candidate;
      if (textOffset <= 0 || textOffset >= flattenedText.length) return;
      if (textOffset <= firstVisibleOffset || textOffset > lastVisibleOffset) return;
      if (this.isNonBreakingTextBoundary(flattenedText, textOffset)) return;
      if (!this.isLegalParagraphLineBoundary(paragraph, flattenedText, textOffset)) return;
      if (graphemeBoundaries) {
        if (!graphemeBoundaries.has(textOffset)) return;
      } else if (!this.isConservativeFallbackTextBoundary(flattenedText, textOffset)) {
        return;
      }
      candidates.push(candidate);
    };

    for (const node of textNodes) {
      this.checkpoint();
      const start = textStarts.get(node)!;
      // Field results and converter list markers are semantic atoms. A field
      // receives one boundary outside its wrapper below; a list marker stays
      // with the first real text fragment and is never emitted by itself.
      const atomic = node.parentElement?.closest<HTMLElement>(
        "[data-field], [data-list-marker], a[data-comment-id]",
      );
      if (atomic && paragraph.contains(atomic)) continue;

      // Only CSS-collapsible ASCII whitespace is a wrapping opportunity. JS's
      // broader `\s` class includes NBSP and other explicitly non-breaking text.
      const whitespace = /[\u0009-\u000D\u0020]+/g;
      let match: RegExpExecArray | null;
      let matchesSinceCheckpoint = 0;
      while ((match = whitespace.exec(node.data)) !== null) {
        if (++matchesSinceCheckpoint % 256 === 0) this.checkpoint();
        const offset = match.index + match[0].length;
        addCandidate({ node, offset, textOffset: start + offset, priority: 0 });
      }
      addCandidate({
        node,
        offset: node.data.length,
        textOffset: start + node.data.length,
        priority: 0,
      });
      if (emergencyGraphemeBreaks) {
        for (let offset = 1; offset < node.data.length; offset++) {
          if (offset % 256 === 0) this.checkpoint();
          addCandidate({
            node,
            offset,
            textOffset: start + offset,
            priority: -1,
          });
        }
      }
    }

    // A field is splittable only immediately after its complete outer wrapper.
    // Starting the tail at the parent's child boundary prevents Range from
    // cloning an empty `[data-field]` shell that substitution could repopulate.
    for (const field of Array.from(paragraph.querySelectorAll<HTMLElement>("[data-field]"))) {
      this.checkpoint();
      if (field.parentElement?.closest("[data-field]")) continue;
      const last = lastTextByField.get(field);
      const parent = field.parentNode;
      if (!last || !parent) continue;
      const childIndex = Array.prototype.indexOf.call(parent.childNodes, field);
      if (childIndex < 0) continue;
      addCandidate({
        node: parent,
        offset: childIndex + 1,
        textOffset: textStarts.get(last)! + last.data.length,
        priority: 1,
      });
    }

    // Several DOM boundaries can represent the same flattened offset. Prefer
    // the outer atomic boundary, then keep one stable document-order candidate.
    candidates.sort((left, right) =>
      left.textOffset - right.textOffset || right.priority - left.priority);
    const endpoints: ParagraphFragmentEndpoint[] = [];
    for (const candidate of candidates) {
      if (endpoints[endpoints.length - 1]?.textOffset !== candidate.textOffset) {
        endpoints.push(candidate);
      }
    }
    return endpoints;
  }

  /** Browser-native UAX #29 boundaries; null keeps older runtimes conservative. */
  private graphemeBoundaryOffsets(text: string): Set<number> | null {
    type Segment = { segment: string; index: number };
    type Segmenter = new (
      locale?: string,
      options?: { granularity: "grapheme" },
    ) => { segment(value: string): Iterable<Segment> };
    const SegmenterConstructor = (this.view.Intl as typeof Intl & {
      Segmenter?: Segmenter;
    }).Segmenter;
    if (!SegmenterConstructor) return null;

    const boundaries = new Set<number>([0, text.length]);
    let segmentCount = 0;
    for (const part of new SegmenterConstructor("en", { granularity: "grapheme" }).segment(text)) {
      if (++segmentCount % 256 === 0) this.checkpoint();
      boundaries.add(part.index);
      boundaries.add(part.index + part.segment.length);
    }
    return boundaries;
  }

  /**
   * Conservative UAX #14/CSS wrapping opportunities used for synthetic page
   * boundaries. In particular, a formatting-run boundary is not itself a word
   * boundary, and Japanese opening/closing punctuation stays with its pair.
   * Arbitrary grapheme breaks are admitted only when the paragraph's CSS asks
   * the browser to wrap anywhere.
   */
  private isLegalParagraphLineBoundary(
    paragraph: HTMLElement,
    text: string,
    offset: number,
  ): boolean {
    const lastCodeUnit = text.charCodeAt(offset - 1);
    const beforeOffset = lastCodeUnit >= 0xDC00 && lastCodeUnit <= 0xDFFF
      ? Math.max(0, offset - 2)
      : offset - 1;
    const beforeCodePoint = text.codePointAt(beforeOffset);
    const afterCodePoint = text.codePointAt(offset);
    const before = beforeCodePoint === undefined ? "" : String.fromCodePoint(beforeCodePoint);
    const after = afterCodePoint === undefined ? "" : String.fromCodePoint(afterCodePoint);
    if (/^[\u0009-\u000D\u0020]$/.test(before)
        || /^[\u0009-\u000D\u0020]$/.test(after)) return true;
    if (/[-\u00AD\u2010]$/u.test(before)) return true;

    const style = this.view.getComputedStyle(paragraph);
    if (style.wordBreak === "break-all"
        || style.overflowWrap === "anywhere"
        || style.overflowWrap === "break-word") return true;

    const eastAsian = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u;
    if (!eastAsian.test(before) && !eastAsian.test(after)) return false;
    const opening = "([{<\u2018\u201C\u3008\u300A\u300C\u300E\u3010\u3014\u3016\u3018\u301A\uFF08\uFF3B\uFF5B";
    const closing = ")]}>\u2019\u201D\u3001\u3002\u3009\u300B\u300D\u300F\u3011\u3015\u3017\u3019\u301B\uFF01\uFF05\uFF09\uFF0C\uFF0E\uFF1A\uFF1B\uFF1F\uFF3D\uFF5D";
    return !opening.includes(before) && !closing.includes(after);
  }

  /** Never fragment across characters whose line-break meaning is explicitly non-breaking. */
  private isNonBreakingTextBoundary(text: string, offset: number): boolean {
    const codePointBefore = (index: number): number | undefined => {
      if (index <= 0) return undefined;
      const last = text.charCodeAt(index - 1);
      const start = last >= 0xDC00 && last <= 0xDFFF ? index - 2 : index - 1;
      return text.codePointAt(Math.max(0, start));
    };
    const nonBreaking = new Set([
      0x00A0, 0x200E, 0x200F, 0x2011, 0x202A, 0x202B, 0x202C,
      0x202D, 0x202E, 0x202F, 0x2060, 0x2066, 0x2067, 0x2068,
      0x2069, 0xFEFF,
    ]);
    return nonBreaking.has(codePointBefore(offset) ?? -1)
      || nonBreaking.has(text.codePointAt(offset) ?? -1);
  }

  /**
   * Without Intl.Segmenter, admit only an all-ASCII boundary. This still
   * fragments ordinary prose while refusing to guess about Unicode clusters.
   */
  private isConservativeFallbackTextBoundary(text: string, offset: number): boolean {
    return text.charCodeAt(offset - 1) <= 0x7F && text.charCodeAt(offset) <= 0x7F;
  }

  /**
   * Whether a range contains visible text after ignoring bidi/zero-width marks.
   * Paragraph fragmentation deliberately excludes non-textual descendants, so
   * this is enough to reject empty head or tail fragments.
   */
  private hasVisibleFragmentText(fragment: DocumentFragment): boolean {
    return this.hasVisibleText(fragment.textContent || "");
  }

  private hasVisibleText(text: string): boolean {
    return text
      .replace(/[\u200B-\u200F\uFEFF]/g, "")
      .trim()
      .length > 0;
  }

  /**
   * Phase one only handles text-like paragraph descendants. Objects, explicit
   * line breaks, list markers, notes, and out-of-flow/inline-block content all
   * require their own line-layout rules and retain the established whole-block
   * fallback instead of risking broken content or duplicate anchors.
   */
  private canFragmentParagraph(block: MeasuredBlock): boolean {
    if (!this.fragmentParagraphs) {
      return false;
    }

    const paragraph = block.element;
    if (
      block.keepWithNext ||
      block.keepLines ||
      block.pageBreakBefore ||
      block.isPageBreak ||
      paragraph.dataset.widowControl === "true" ||
      paragraph.hasAttribute("contenteditable")
    ) {
      return false;
    }

    return this.canRangeFragmentParagraph(paragraph);
  }

  /**
   * Shared structural gate for body and note paragraph fragmentation. Footnote
   * first paragraphs are inline beside their marker, so that one known layout
   * context may opt into an inline root while retaining every descendant and
   * break-safety restriction used for body text.
   */
  private canRangeFragmentParagraph(
    paragraph: HTMLElement,
    allowInlineRoot: boolean = false,
  ): boolean {
    if (
      paragraph.tagName !== "P" ||
      paragraph.dataset.keepWithNext === "true" ||
      paragraph.dataset.keepLines === "true" ||
      paragraph.dataset.pageBreakBefore === "true" ||
      paragraph.dataset.widowControl === "true" ||
      paragraph.hasAttribute("contenteditable")
    ) {
      return false;
    }

    // Non-textual/out-of-flow descendants still need dedicated fragmenters.
    // Inline ids/editor anchors are reconciled after Range cloning, and list
    // markers are kept atomically with the leading text fragment.
    const unsupportedDescendants = [
      "br",
      "img",
      "picture",
      "svg",
      "math",
      "canvas",
      "video",
      "audio",
      "iframe",
      "object",
      "embed",
      "input",
      "button",
      "select",
      "textarea",
      "table",
      "ol",
      "ul",
      "li",
      "dl",
      "div",
      "p",
      "section",
      "article",
      "aside",
      "figure",
      "fieldset",
      "details",
      "[data-footnote-id]",
      "[contenteditable]",
    ].join(", ");
    if (paragraph.querySelector(unsupportedDescendants)) {
      return false;
    }

    const isValidatedEndnote = paragraph.dataset.paginationSafeEndnote === "true";
    return isValidatedEndnote || this.hasRangeFragmentSafeLayout(paragraph, allowInlineRoot);
  }

  /**
   * A range clone preserves nested inline formatting exactly. Anything that establishes its own
   * box/layout context is deferred until a future fragmenter can model it accurately. Callers
   * must invoke this while the paragraph is attached to the styled document.
   */
  private hasRangeFragmentSafeLayout(
    paragraph: HTMLElement,
    allowInlineRoot: boolean = false,
  ): boolean {
    const paragraphStyle = this.view.getComputedStyle(paragraph);
    if (
      (paragraphStyle.display !== "block" &&
        !(allowInlineRoot && paragraphStyle.display === "inline")) ||
      paragraphStyle.position !== "static" ||
      paragraphStyle.float !== "none" ||
      (paragraphStyle.whiteSpace !== "normal" && paragraphStyle.whiteSpace !== "pre-wrap") ||
      paragraphStyle.breakBefore !== "auto" ||
      paragraphStyle.breakAfter !== "auto" ||
      paragraphStyle.breakInside === "avoid" ||
      paragraphStyle.pageBreakBefore !== "auto" ||
      paragraphStyle.pageBreakAfter !== "auto" ||
      paragraphStyle.pageBreakInside === "avoid"
    ) {
      return false;
    }

    for (const descendant of Array.from(paragraph.querySelectorAll<HTMLElement>("*"))) {
      const style = this.view.getComputedStyle(descendant);
      if (
        style.display !== "inline" ||
        style.position !== "static" ||
        style.float !== "none" ||
        (style.whiteSpace !== "normal" && style.whiteSpace !== "pre-wrap")
      ) {
        return false;
      }
    }
    return true;
  }

  /**
   * Builds one range-cloned paragraph fragment. Only the leading fragment keeps
   * the source paragraph's addressability; continuations must not duplicate an
   * id/data-anchor in the rendered document.
   */
  private createParagraphFragment(
    paragraph: HTMLElement,
    range: Range,
    retainSourceIdentity: boolean,
    isFinalFragment: boolean
  ): HTMLElement {
    const fragment = paragraph.cloneNode(false) as HTMLElement;
    const contents = range.cloneContents();
    fragment.appendChild(contents);

    if (!retainSourceIdentity) {
      fragment.removeAttribute("id");
      fragment.removeAttribute("data-anchor");
      // A continuation starts with a normal line rather than repeating a first-
      // line/hanging indent. Keep the side margin so paragraph alignment remains.
      fragment.style.setProperty("margin-top", "0", "important");
      fragment.style.setProperty("text-indent", "0", "important");
    }
    if (!isFinalFragment) {
      // The original bottom spacing belongs after the complete paragraph, not
      // between synthetic page fragments.
      fragment.style.setProperty("margin-bottom", "0", "important");
    }

    return fragment;
  }

  /**
   * Range clones repeat an inline ancestor when the split lands inside it. Keep
   * semantic wrappers (links, comments, formatting) on both sides, but retain a
   * duplicated HTML/editor identity only on the leading fragment. Targets that
   * occur wholly after the split are absent from `head` and remain on `tail`.
   */
  private reconcileParagraphFragmentIdentities(
    head: HTMLElement,
    tail: HTMLElement,
  ): void {
    const elements = (root: HTMLElement, selector: string): HTMLElement[] => [
      ...(root.matches(selector) ? [root] : []),
      ...Array.from(root.querySelectorAll<HTMLElement>(selector)),
    ];

    const headIds = new Set(elements(head, "[id]").map((element) => element.id));
    for (const element of elements(tail, "[id]")) {
      if (headIds.has(element.id)) element.removeAttribute("id");
    }

    const headAnchors = new Set(elements(head, "[data-anchor]")
      .map((element) => element.dataset.anchor)
      .filter((anchor): anchor is string => Boolean(anchor)));
    for (const element of elements(tail, "[data-anchor]")) {
      if (!element.dataset.anchor || !headAnchors.has(element.dataset.anchor)) continue;
      element.removeAttribute("data-anchor");
      element.removeAttribute("data-committed-text");
    }
  }

  /**
   * Range-clone the largest safe prefix accepted by `fits`. The measurement
   * policy stays with the caller, allowing body blocks and note bands to share
   * one DOM fragmenter while measuring in their respective layout contexts.
   */
  private splitParagraphAtLargestFit(
    paragraph: HTMLElement,
    fits: (head: HTMLElement) => boolean,
    options: {
      emergencyGraphemeBreaks?: boolean;
      forceFirstOnNoFit?: boolean;
    } = {},
  ): ParagraphSplitAttempt {
    const largestFit = (endpoints: ParagraphFragmentEndpoint[]) => {
      // Prefix height is usually monotone, but selector-sensitive inline CSS
      // (`:last-child`, `:only-child`) can make a longer Range clone shorter.
      // Probe the largest legal endpoint first so the common non-monotone case
      // cannot be discarded by the binary search below.
      const lastEndpoint = endpoints[endpoints.length - 1];
      if (lastEndpoint) {
        const lastRange = this.document.createRange();
        lastRange.setStart(paragraph, 0);
        lastRange.setEnd(lastEndpoint.node, lastEndpoint.offset);
        const lastContents = lastRange.cloneContents();
        if (this.hasVisibleFragmentText(lastContents)) {
          const lastHead = this.createParagraphFragment(paragraph, lastRange, true, false);
          if (fits(lastHead)) return { endpoint: lastEndpoint, head: lastHead };
        }
      }

      let low = 0;
      let high = endpoints.length - 2;
      let best: { endpoint: ParagraphFragmentEndpoint; head: HTMLElement } | null = null;
      while (low <= high) {
        this.checkpoint();
        const middle = Math.floor((low + high) / 2);
        const endpoint = endpoints[middle];
        const headRange = this.document.createRange();
        headRange.setStart(paragraph, 0);
        headRange.setEnd(endpoint.node, endpoint.offset);
        const headContents = headRange.cloneContents();
        if (!this.hasVisibleFragmentText(headContents)) {
          low = middle + 1;
          continue;
        }

        const head = this.createParagraphFragment(paragraph, headRange, true, false);
        if (fits(head)) {
          best = { endpoint, head };
          low = middle + 1;
        } else {
          high = middle - 1;
        }
      }
      return best;
    };

    let endpoints = this.paragraphFragmentEndpoints(paragraph);
    let best = largestFit(endpoints);
    if (!best && options.emergencyGraphemeBreaks) {
      endpoints = this.paragraphFragmentEndpoints(paragraph, true);
      best = largestFit(endpoints);
    }
    if (!best && options.forceFirstOnNoFit && endpoints.length > 0) {
      const endpoint = endpoints[0];
      const headRange = this.document.createRange();
      headRange.setStart(paragraph, 0);
      headRange.setEnd(endpoint.node, endpoint.offset);
      best = {
        endpoint,
        head: this.createParagraphFragment(paragraph, headRange, true, false),
      };
    }
    if (!best) {
      return endpoints.length > 0 ? { kind: "no-fit" } : { kind: "indivisible" };
    }

    const tailRange = this.document.createRange();
    tailRange.setStart(best.endpoint.node, best.endpoint.offset);
    tailRange.setEnd(paragraph, paragraph.childNodes.length);
    const tailContents = tailRange.cloneContents();
    if (!this.hasVisibleFragmentText(tailContents)) return { kind: "indivisible" };

    const tail = this.createParagraphFragment(paragraph, tailRange, false, true);
    this.reconcileParagraphFragmentIdentities(best.head, tail);

    return {
      kind: "split",
      head: best.head,
      tail,
    };
  }

  /**
   * Splits a simple paragraph at the largest DOM Range endpoint that fits the
   * currently available body space. The caller then processes the tail normally,
   * allowing it to fragment again on later pages when necessary.
   */
  private tryFragmentParagraph(
    block: MeasuredBlock,
    dims: PageDimensions,
    availableHeightPt: number,
    effectiveMarginTopPt: number
  ): MeasuredBlock[] | null {
    if (!this.canFragmentParagraph(block) || availableHeightPt <= effectiveMarginTopPt) {
      return null;
    }

    const split = this.splitParagraphAtLargestFit(block.element, (head) => {
      const measured = this.measureElement(head, dims, block.sectionIndex);
      return effectiveMarginTopPt + measured.heightPt <= availableHeightPt;
    });
    if (split.kind !== "split") return null;

    const headMeasured = this.measureElement(split.head, dims, block.sectionIndex);
    const tailMeasured = this.measureElement(split.tail, dims, block.sectionIndex);

    return [
      {
        ...headMeasured,
        element: split.head,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
      },
      {
        ...tailMeasured,
        element: split.tail,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
      },
    ];
  }

  /**
   * Parses the header/footer registry from the staging element.
   * Also measures heights during parsing for lazy-loading compatibility.
   */
  private parseHeaderFooterRegistry(): HeaderFooterRegistry {
    const registry: HeaderFooterRegistry = new Map();
    const registryEl = this.stagingElement.querySelector("#pagination-hf-registry");

    if (!registryEl) return registry;

    // Build a map of section index -> content width for measurement
    const sectionWidths = new Map<number, number>();
    const sections = Array.from(this.stagingElement.querySelectorAll<HTMLElement>("[data-section-index]"));
    for (const section of sections) {
      const idx = parseInt(section.dataset.sectionIndex || "0", 10);
      const contentWidth = parseFloat(section.dataset.contentWidth || "") || DEFAULT_PAGE_WIDTH - 2 * DEFAULT_MARGIN;
      sectionWidths.set(idx, contentWidth);
    }
    // Fallback content width if no sections found
    const defaultContentWidth = sectionWidths.get(0) || DEFAULT_PAGE_WIDTH - 2 * DEFAULT_MARGIN;

    const entries = Array.from(registryEl.querySelectorAll<HTMLElement>("[data-section][data-hf-type]"));

    for (const entry of entries) {
      const sectionIndex = parseInt(entry.dataset.section || "0", 10);
      const hfType = entry.dataset.hfType as string;

      if (!registry.has(sectionIndex)) {
        registry.set(sectionIndex, {});
      }

      const section = registry.get(sectionIndex)!;
      // Clone the first child element (the actual header/footer content)
      const content = entry.cloneNode(true) as HTMLElement;

      // Get content width for this section (for accurate measurement)
      const contentWidth = sectionWidths.get(sectionIndex) || defaultContentWidth;

      // Measure height during parsing (one-time cost, enables lazy loading)
      const measuredHeight = this.measureHeaderFooterHeight(content, contentWidth);

      switch (hfType) {
        case "header-default":
          section.headerDefault = content;
          section.headerDefaultHeight = measuredHeight;
          break;
        case "header-first":
          section.headerFirst = content;
          section.headerFirstHeight = measuredHeight;
          break;
        case "header-even":
          section.headerEven = content;
          section.headerEvenHeight = measuredHeight;
          break;
        case "footer-default":
          section.footerDefault = content;
          section.footerDefaultHeight = measuredHeight;
          break;
        case "footer-first":
          section.footerFirst = content;
          section.footerFirstHeight = measuredHeight;
          break;
        case "footer-even":
          section.footerEven = content;
          section.footerEvenHeight = measuredHeight;
          break;
      }
    }

    return registry;
  }

  /**
   * Parses the footnote registry from the staging element.
   */
  private parseFootnoteRegistry(): FootnoteRegistry {
    const registry: FootnoteRegistry = new Map();
    const registryEl = this.stagingElement.querySelector("#pagination-footnote-registry");

    if (!registryEl) return registry;

    // Only direct registry entries are definitions. Custom separator stories may
    // legitimately contain arbitrary converted markup, including stale semantic
    // attributes that must never be mistaken for another note definition.
    const entries = Array.from(registryEl.children)
      .filter((entry): entry is HTMLElement =>
        entry instanceof this.view.HTMLElement && entry.hasAttribute("data-footnote-id"));

    for (const entry of entries) {
      const footnoteId = entry.dataset.footnoteId;
      if (footnoteId) {
        // Clone the footnote element for later use
        registry.set(footnoteId, entry.cloneNode(true) as HTMLElement);
      }
    }

    return registry;
  }

  /** Read Word's optional normal and continuation separator stories. */
  private parseFootnoteSeparators(): {
    normal: HTMLElement | null;
    continuation: HTMLElement | null;
  } {
    const registry = this.stagingElement.querySelector("#pagination-footnote-registry");
    const clone = (kind: "normal" | "continuation") => {
      const source = Array.from(registry?.children ?? []).find((entry) =>
        entry instanceof this.view.HTMLElement
        && entry.getAttribute("data-footnote-separator") === kind,
      );
      return source?.cloneNode(true) as HTMLElement | undefined;
    };
    return {
      normal: clone("normal") ?? null,
      continuation: clone("continuation") ?? null,
    };
  }

  /** Append the exact separator story selected for this initial/continued note band. */
  private appendFootnoteSeparator(container: HTMLElement, continuation: boolean): void {
    const kind = continuation ? "continuation" : "normal";
    const source = continuation ? this.footnoteContinuationSeparator : this.footnoteSeparator;
    if (!source) {
      const fallback = this.document.createElement("hr");
      fallback.dataset.footnoteSeparator = kind;
      container.appendChild(fallback);
      return;
    }

    const clone = source.cloneNode(true) as HTMLElement;
    for (const element of [clone, ...Array.from(clone.querySelectorAll<HTMLElement>("*"))]) {
      element.removeAttribute("id");
      element.removeAttribute("data-anchor");
      element.removeAttribute("data-committed-text");
      element.removeAttribute("data-source-anchor-id");
      element.removeAttribute("data-page-fragment-id");
      element.removeAttribute("data-fragment-index");
      element.removeAttribute("data-footnote-id");
      element.removeAttribute("data-comment-id");
      element.removeAttribute("name");
      if (element instanceof this.view.HTMLAnchorElement
          && element.getAttribute("href")?.startsWith("#")) {
        element.removeAttribute("href");
      }
      element.setAttribute("contenteditable", "false");
    }
    container.appendChild(clone);
  }

  /** Parses the hidden source notes used to render paginated margin comments. */
  private parseCommentMarginRegistry(): Map<string, HTMLElement> {
    const registry = new Map<string, HTMLElement>();
    const registryEl = this.stagingElement.querySelector(
      "#pagination-comment-margin-registry",
    );
    if (!registryEl) return registry;

    for (const entry of Array.from(
      registryEl.querySelectorAll<HTMLElement>("[data-comment-id]"),
    )) {
      const commentId = entry.dataset.commentId;
      if (commentId) registry.set(commentId, entry.cloneNode(true) as HTMLElement);
    }
    return registry;
  }

  /**
   * Extracts footnote reference IDs from an element.
   */
  private extractFootnoteRefs(element: HTMLElement): string[] {
    const refs = element.querySelectorAll<HTMLElement>("[data-footnote-id]");
    const ids: string[] = [];
    for (const ref of Array.from(refs)) {
      const id = ref.dataset.footnoteId;
      if (id && !ids.includes(id)) {
        ids.push(id);
      }
    }
    return ids;
  }

  /** Clone a note child for the continuation queue with its original selector position. */
  private cloneFootnoteElementForContinuation(
    element: HTMLElement,
    sourceIndex: number,
  ): HTMLElement {
    const clone = element.cloneNode(true) as HTMLElement;
    clone.setAttribute(
      FOOTNOTE_SOURCE_POSITION_ATTR,
      sourceIndex === 0 ? "first" : "later",
    );
    return clone;
  }

  /** Remove identities that belong only to the initial registry presentation shell. */
  private makeContinuationShellInert(element: HTMLElement): void {
    element.removeAttribute("id");
    element.removeAttribute("data-anchor");
    element.removeAttribute("data-committed-text");
  }

  /**
   * Build the exact continuation shape shared by measurement and paint.
   *
   * Keep the source `.footnote-item > .footnote-content` ancestry: document
   * CSS, inherited direction/language, and custom classes frequently target
   * those shells. Only the number and initial HTML/editor identities are
   * omitted. A hidden sentinel preserves `p:not(:first-of-type)` for a page
   * whose first carried element was a later source paragraph.
   */
  private createFootnoteContinuationWrapper(
    continuation: FootnoteContinuation,
    elements: readonly HTMLElement[] = continuation.remainingElements,
  ): HTMLElement {
    const source = this.footnoteRegistry.get(continuation.footnoteId);
    const wrapper = source
      ? source.cloneNode(false) as HTMLElement
      : this.document.createElement("div");
    this.makeContinuationShellInert(wrapper);
    wrapper.classList.add("footnote-continuation");
    wrapper.dataset.footnoteId = continuation.footnoteId;
    if (continuation.sourceAnchorId) {
      wrapper.dataset.sourceAnchorId = continuation.sourceAnchorId;
    }

    const sourceContent = source?.querySelector<HTMLElement>(".footnote-content");
    const content = sourceContent
      ? sourceContent.cloneNode(false) as HTMLElement
      : this.document.createElement("span");
    this.makeContinuationShellInert(content);
    if (!sourceContent) content.className = "footnote-content";

    const first = elements[0];
    if (
      first?.tagName === "P"
      && first.getAttribute(FOOTNOTE_SOURCE_POSITION_ATTR) === "later"
    ) {
      const sentinel = this.document.createElement("p");
      sentinel.className = "footnote-continuation-position-sentinel";
      sentinel.setAttribute("aria-hidden", "true");
      sentinel.style.setProperty("display", "none", "important");
      content.appendChild(sentinel);
    }
    for (let index = 0; index < elements.length; index++) {
      this.checkpoint();
      const element = elements[index];
      const clone = element.cloneNode(true) as HTMLElement;
      clone.removeAttribute(FOOTNOTE_SOURCE_POSITION_ATTR);
      if (index === 0 && clone.tagName === "P") {
        // Without the initial number, a carried paragraph always starts a new
        // line even when source CSS made the note's first paragraph inline.
        clone.style.setProperty("display", "block", "important");
      }
      content.appendChild(clone);
    }
    wrapper.appendChild(content);
    return wrapper;
  }

  /** Build the exact partial-note item shape shared by measurement and paint. */
  private createPartialFootnoteItem(
    footnote: HTMLElement,
    fittingElements: readonly HTMLElement[],
  ): HTMLElement {
    const item = footnote.cloneNode(false) as HTMLElement;
    const number = footnote.querySelector(".footnote-number");
    if (number) item.appendChild(number.cloneNode(true));

    const sourceContent = footnote.querySelector<HTMLElement>(".footnote-content");
    const content = sourceContent
      ? sourceContent.cloneNode(false) as HTMLElement
      : this.document.createElement("span");
    if (!sourceContent) content.className = "footnote-content";
    for (const element of fittingElements) {
      this.checkpoint();
      content.appendChild(element.cloneNode(true));
    }
    item.appendChild(content);
    return item;
  }

  /**
   * Measures the height of footnotes for given IDs (in points).
   * Creates a temporary container to measure the footnotes.
   * @param footnoteIds - IDs of footnotes to measure
   * @param contentWidth - Width for measurement
   * @param continuation - Optional continuation content to include first
   */
  private measureFootnotesHeight(
    footnoteIds: string[],
    contentWidth: number,
    continuation?: FootnoteContinuation | null,
    partialFootnotes?: readonly PartialFootnote[],
  ): number {
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if (footnoteIds.length === 0 && !hasContinuation) {
      return 0;
    }

    // Measure in the SAME styling context the notes render in: `.page-footnotes` carries
    // its own font-size and line-height, so measuring without the class sizes the note
    // block against body type and the reserve can never match what is drawn.
    // Create a temporary measurement container
    const measureContainer = this.document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    measureContainer.className = this.cssPrefix + "footnotes";

    // Add the same normal/continuation separator story that will be painted.
    this.appendFootnoteSeparator(measureContainer, Boolean(hasContinuation));

    // Add continuation content first (if any)
    if (hasContinuation) {
      measureContainer.appendChild(this.createFootnoteContinuationWrapper(continuation!));
    }

    // Add footnotes
    const partialById = new Map(
      partialFootnotes?.map((partial) => [partial.footnoteId, partial] as const) ?? [],
    );
    for (const id of footnoteIds) {
      this.checkpoint();
      const footnote = this.footnoteRegistry.get(id);
      if (footnote) {
        const partial = partialById.get(id);
        measureContainer.appendChild(partial
          ? this.createPartialFootnoteItem(footnote, partial.fittingElements)
          : footnote.cloneNode(true));
      }
    }

    // Append to staging for measurement
    this.stagingElement.appendChild(measureContainer);
    try {
      return pxToPt(measureContainer.getBoundingClientRect().height);
    } finally {
      measureContainer.remove();
    }
  }

  /**
   * Partition a continuation for one page's note band, preferring complete children
   * and range-fragmenting an eligible paragraph when necessary. Always advances by
   * at least one element so an indivisible oversized paragraph follows the established
   * clipped fallback without trapping pagination in a loop.
   */
  private splitContinuationForPage(
    continuation: FootnoteContinuation,
    availableHeightPt: number,
    contentWidth: number,
  ): { current: FootnoteContinuation; overflow: FootnoteContinuation | null } {
    const fitting: HTMLElement[] = [];
    let remaining: HTMLElement[] = [];

    // Keep one live measurement tree and append each candidate exactly once.
    // Rebuilding `[...fitting, candidate]` for every child cloned 1+2+...+N
    // descendants before a page/resource cap could run, which is quadratic for
    // producer-authored notes containing thousands of tiny paragraphs/runs.
    const measureContainer = this.document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    measureContainer.className = this.cssPrefix + "footnotes";
    this.appendFootnoteSeparator(measureContainer, true);
    const wrapper = this.createFootnoteContinuationWrapper(continuation, []);
    const content = wrapper.querySelector<HTMLElement>(":scope > .footnote-content");
    if (!content) {
      throw new Error("Footnote continuation is missing its content shell");
    }
    measureContainer.appendChild(wrapper);
    this.stagingElement.appendChild(measureContainer);

    try {
      for (let index = 0; index < continuation.remainingElements.length; index++) {
        this.checkpoint();
        const element = continuation.remainingElements[index];
        const sourcePosition = element.getAttribute(FOOTNOTE_SOURCE_POSITION_ATTR);
        if (
          fitting.length === 0
          && element.tagName === "P"
          && element.getAttribute(FOOTNOTE_SOURCE_POSITION_ATTR) === "later"
        ) {
          const sentinel = this.document.createElement("p");
          sentinel.className = "footnote-continuation-position-sentinel";
          sentinel.setAttribute("aria-hidden", "true");
          sentinel.style.setProperty("display", "none", "important");
          content.appendChild(sentinel);
        }

        const candidate = element.cloneNode(true) as HTMLElement;
        candidate.removeAttribute(FOOTNOTE_SOURCE_POSITION_ATTR);
        if (fitting.length === 0 && candidate.tagName === "P") {
          candidate.style.setProperty("display", "block", "important");
        }
        content.appendChild(candidate);
        const candidateHeight = pxToPt(measureContainer.getBoundingClientRect().height);
        if (candidateHeight <= availableHeightPt) {
          fitting.push(element);
          continue;
        }

        const canSplitCandidate = candidate.tagName === "P"
          && this.canRangeFragmentParagraph(candidate);
        candidate.remove();
        let split: ParagraphSplitAttempt | null = null;
        if (canSplitCandidate) {
          split = this.splitParagraphAtLargestFit(candidate, (head) => {
            content.appendChild(head);
            try {
              return pxToPt(measureContainer.getBoundingClientRect().height) <= availableHeightPt;
            } finally {
              head.remove();
            }
          }, {
            emergencyGraphemeBreaks: true,
            // A continuation page owns the full note band. If even one legal
            // grapheme is taller than it, clip only that unit and keep draining.
            forceFirstOnNoFit: true,
          });
        }
        if (split?.kind === "split") {
          if (sourcePosition) {
            split.head.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, sourcePosition);
            split.tail.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, sourcePosition);
          }
          fitting.push(split.head);
          remaining = [split.tail, ...continuation.remainingElements.slice(index + 1)];
        } else if (fitting.length === 0) {
          // A genuinely indivisible first element retains the established visible
          // clipped fallback, but only for itself. Its siblings continue later.
          fitting.push(element);
          remaining = continuation.remainingElements.slice(index + 1);
        } else {
          remaining = continuation.remainingElements.slice(index);
        }
        break;
      }
    } finally {
      measureContainer.remove();
    }

    return {
      current: {
        footnoteId: continuation.footnoteId,
        sourceAnchorId: continuation.sourceAnchorId,
        remainingElements: fitting,
      },
      overflow: remaining.length > 0 ? {
        footnoteId: continuation.footnoteId,
        sourceAnchorId: continuation.sourceAnchorId,
        remainingElements: remaining,
      } : null,
    };
  }

  /**
   * Splits a footnote element into parts that fit within the available height.
   * Returns the elements that fit and the elements that need to continue.
   */
  private splitFootnoteToFit(
    footnoteElement: HTMLElement,
    availableHeightPt: number,
    contentWidth: number,
    forceProgress = false,
    existingPayload?: FootnotePackingContext,
  ): { fits: HTMLElement[]; overflow: HTMLElement[] } {
    // Registry entries are detached clones, so getComputedStyle() cannot validate
    // their real paragraph layout. Attach an exact clone in the rendered note
    // context for the duration of the conservative range-fragmentation check.
    const layoutContext = this.document.createElement("div");
    layoutContext.style.position = "absolute";
    layoutContext.style.visibility = "hidden";
    layoutContext.style.width = `${contentWidth}pt`;
    layoutContext.style.left = "-9999px";
    layoutContext.className = this.cssPrefix + "footnotes";
    this.appendFootnoteSeparator(layoutContext, false);
    const attachedFootnote = footnoteElement.cloneNode(true) as HTMLElement;
    layoutContext.appendChild(attachedFootnote);
    this.stagingElement.appendChild(layoutContext);

    // A second live tree represents the exact already-packed page payload plus
    // an initially empty partial item. Candidates are appended once and layout
    // is read in place, avoiding cumulative prefix re-cloning.
    const measureContainer = this.document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    measureContainer.className = this.cssPrefix + "footnotes";
    const hasContinuation = Boolean(existingPayload?.continuation?.remainingElements.length);
    this.appendFootnoteSeparator(measureContainer, hasContinuation);
    if (hasContinuation) {
      measureContainer.appendChild(this.createFootnoteContinuationWrapper(
        existingPayload!.continuation!,
      ));
    }
    const partialById = new Map(
      existingPayload?.partialFootnotes?.map((partial) =>
        [partial.footnoteId, partial] as const) ?? [],
    );
    for (const id of existingPayload?.footnoteIds ?? []) {
      this.checkpoint();
      const source = this.footnoteRegistry.get(id);
      if (!source) continue;
      const partial = partialById.get(id);
      measureContainer.appendChild(partial
        ? this.createPartialFootnoteItem(source, partial.fittingElements)
        : source.cloneNode(true));
    }
    const measuredPartial = this.createPartialFootnoteItem(attachedFootnote, []);
    const measuredContent = measuredPartial.querySelector<HTMLElement>(":scope > .footnote-content");
    if (!measuredContent) {
      layoutContext.remove();
      throw new Error("Footnote is missing its content shell");
    }
    measureContainer.appendChild(measuredPartial);
    this.stagingElement.appendChild(measureContainer);

    // Get child elements (paragraphs) of the footnote content.
    //
    // `fits` is spliced into a freshly built `.footnote-item` > `.footnote-content` wrapper by
    // addPageFootnotes, so it must contain the note's CONTENT, never the note element itself.
    // Returning the whole `.footnote-item` here nested a complete item (number span and all)
    // inside another item's content span; the inner block-level div then broke the line, so the
    // note's number rendered alone above its text — the same visible symptom as the escaped-CSS
    // bug, from an unrelated cause, on the notes that happened to take a can't-split path.
    try {
      const footnoteContent = attachedFootnote.querySelector(".footnote-content");
      if (!footnoteContent) {
        // No content structure to split — hand back the element's own children.
        return {
          fits: Array.from(attachedFootnote.children).map((el) =>
            el.cloneNode(true) as HTMLElement),
          overflow: [],
        };
      }

      const children = Array.from(footnoteContent.children) as HTMLElement[];
      const fits: HTMLElement[] = [];
      let overflow: HTMLElement[] = [];

      for (let i = 0; i < children.length; i++) {
        this.checkpoint();
        const child = children[i];
        const measuredCandidate = child.cloneNode(true) as HTMLElement;
        measuredContent.appendChild(measuredCandidate);
        const candidateFits = pxToPt(measureContainer.getBoundingClientRect().height)
          <= availableHeightPt;
        if (candidateFits) {
          fits.push(child.cloneNode(true) as HTMLElement);
          continue;
        }

        measuredCandidate.remove();

        const split = this.canRangeFragmentParagraph(child, true)
          ? this.splitParagraphAtLargestFit(child, (head) => {
            measuredContent.appendChild(head);
            try {
              return pxToPt(measureContainer.getBoundingClientRect().height)
                <= availableHeightPt;
            } finally {
              head.remove();
            }
          }, {
            emergencyGraphemeBreaks: true,
            forceFirstOnNoFit: forceProgress,
          })
          : null;
        if (split?.kind === "split") {
          fits.push(split.head);
          split.tail.setAttribute(
            FOOTNOTE_SOURCE_POSITION_ATTR,
            i === 0 ? "first" : "later",
          );
          overflow = [split.tail, ...children.slice(i + 1).map((element, offset) =>
            this.cloneFootnoteElementForContinuation(element, i + 1 + offset))];
        } else if (split?.kind === "no-fit") {
          // The paragraph is splittable, but this citation page has less than
          // one line left. Defer it whole so a fresh note band can try again.
          overflow = children.slice(i).map((element, offset) =>
            this.cloneFootnoteElementForContinuation(element, i + offset));
        } else if (fits.length === 0 && forceProgress) {
          // Preserve the one-element clipping fallback for content that cannot be
          // range-fragmented only on a dedicated full note band, while allowing
          // every later sibling to continue. A residual citation-page band must
          // defer the whole element: it may fit untouched on the next page.
          fits.push(child.cloneNode(true) as HTMLElement);
          overflow = children.slice(i + 1).map((element, offset) =>
            this.cloneFootnoteElementForContinuation(element, i + 1 + offset));
        } else {
          overflow = children.slice(i).map((element, offset) =>
            this.cloneFootnoteElementForContinuation(element, i + offset));
        }
        break;
      }

      return { fits, overflow };
    } finally {
      measureContainer.remove();
      layoutContext.remove();
    }
  }

  /**
   * Adds footnotes to a page container, including continuation content.
   */
  private addPageFootnotes(
    pageBox: HTMLElement,
    footnoteIds: string[],
    dims: PageDimensions,
    bands: PageBands,
    footnoteHeight: number,
    continuation?: FootnoteContinuation | null,
    partialFootnotes?: PartialFootnote[]
  ): void {
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if (footnoteIds.length === 0 && !hasContinuation) {
      return;
    }
    if (this.footnoteRegistry.size === 0 && !hasContinuation) {
      return;
    }

    // Calculate max height for footnotes area (content height minus margin for body content)
    const maxFootnoteHeight = Math.min(
      footnoteHeight,
      bands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO
    );

    const footnotesDiv = this.document.createElement("div");
    footnotesDiv.className = `${this.cssPrefix}footnotes`;
    footnotesDiv.style.position = "absolute";
    // Notes sit at the FOOT OF THE BODY BAND, not at the bottom margin: a footer taller than
    // its margin raises that edge, and anchoring to the raw margin would draw notes over it.
    footnotesDiv.style.bottom = `${dims.pageHeight - (bands.bodyTop + bands.bodyHeight)}pt`;
    footnotesDiv.style.left = `${dims.marginLeft}pt`;
    footnotesDiv.style.width = `${dims.contentWidth}pt`;
    footnotesDiv.style.boxSizing = "border-box";
    // Constrain height and clip overflow to prevent footnotes covering body content
    footnotesDiv.style.maxHeight = `${maxFootnoteHeight}pt`;
    footnotesDiv.style.overflow = "hidden";

    this.appendFootnoteSeparator(footnotesDiv, Boolean(hasContinuation));

    // Add continuation content first (if any)
    if (hasContinuation) {
      footnotesDiv.appendChild(this.createFootnoteContinuationWrapper(continuation!));
    }

    // Clone footnotes in order of appearance
    for (const id of footnoteIds) {
      // Check if this is a partial footnote
      const partial = partialFootnotes?.find(p => p.footnoteId === id);
      if (partial) {
        // Render partial footnote (only the fitting elements)
        const footnote = this.footnoteRegistry.get(id);
        if (footnote) {
          footnotesDiv.appendChild(
            this.createPartialFootnoteItem(footnote, partial.fittingElements),
          );
        }
      } else {
        // Render full footnote from registry
        const footnote = this.footnoteRegistry.get(id);
        if (footnote) {
          footnotesDiv.appendChild(footnote.cloneNode(true));
        }
      }
    }

    pageBox.appendChild(footnotesDiv);
  }

  /** Select the section's first/odd/even header from its one-based page position. */
  private selectHeader(
    sectionIndex: number,
    pageInSection: number,
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first header if available
    if (pageInSection === 1 && sectionHf.headerFirst) {
      return sectionHf.headerFirst;
    }

    // OOXML counts odd/even pages from one inside each section, independently of a PAGE-field
    // restart or format. The displayed page number is deliberately irrelevant here.
    if (pageInSection % 2 === 0 && sectionHf.headerEven) {
      return sectionHf.headerEven;
    }

    // Default (odd) pages
    return sectionHf.headerDefault;
  }

  /** Select the section's first/odd/even footer from its one-based page position. */
  private selectFooter(
    sectionIndex: number,
    pageInSection: number,
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first footer if available
    if (pageInSection === 1 && sectionHf.footerFirst) {
      return sectionHf.footerFirst;
    }

    if (pageInSection % 2 === 0 && sectionHf.footerEven) {
      return sectionHf.footerEven;
    }

    // Default (odd) pages
    return sectionHf.footerDefault;
  }

  /**
   * The header, body, and footer bands for one page position.
   *
   * The single owner of "where does anything sit vertically on this page" — placement in
   * {@link createPage}, the body budget the flow loop spends, and the note area's anchor all
   * read it, so those three cannot disagree about where the body ends.
   *
   * Deterministic: it depends only on the section's page setup and the registry's pre-measured
   * story heights, never on the page's content, which is what keeps it lazy-loading compatible.
   */
  private getPageBands(
    dims: PageDimensions,
    sectionIndex: number,
    pageInSection: number,
  ): PageBands {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    return resolvePageBands(
      dims,
      this.selectStoryHeight(
        sectionHf?.headerFirstHeight,
        sectionHf?.headerEvenHeight,
        sectionHf?.headerDefaultHeight,
        pageInSection
      ),
      this.selectStoryHeight(
        sectionHf?.footerFirstHeight,
        sectionHf?.footerEvenHeight,
        sectionHf?.footerDefaultHeight,
        pageInSection
      )
    );
  }

  /**
   * The measured height of the running story this page position selects, mirroring
   * {@link selectHeader}/{@link selectFooter}. Zero when the page has no such story.
   */
  private selectStoryHeight(
    first: number | undefined,
    even: number | undefined,
    fallback: number | undefined,
    pageInSection: number,
  ): number {
    if (pageInSection === 1 && first != null) return first;
    if (pageInSection % 2 === 0 && even != null) return even;
    return fallback ?? 0;
  }

  /**
   * Measures the content height of a header or footer element.
   *
   * This is what tells {@link resolvePageBands} whether the story stays inside its margin or
   * pushes the body, so it must measure the story ALONE — any padding added here would have to
   * be added to the rendered band too, and the two drifting apart is exactly how a header
   * silently starts overlapping body text.
   */
  private measureHeaderFooterHeight(
    source: HTMLElement,
    contentWidth: number
  ): number {
    // Create a temporary measurement container
    const measureContainer = this.document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";

    // Clone and add the header/footer content
    for (const child of Array.from(source.childNodes)) {
      measureContainer.appendChild(child.cloneNode(true));
    }

    // Append to staging for measurement
    this.stagingElement.appendChild(measureContainer);

    // Measure
    const rect = measureContainer.getBoundingClientRect();
    const heightPt = pxToPt(rect.height);

    // Clean up
    this.stagingElement.removeChild(measureContainer);

    return heightPt;
  }

  /**
   * Flows measured blocks into page containers.
   * Implements a single-pass, forward-only algorithm that is compatible with future lazy loading.
   * Supports footnote continuation - long footnotes can split across pages.
   */
  private flowToPages(
    blocks: MeasuredBlock[],
    dims: PageDimensions,
    startPageNumber: number,
    sectionIndex: number,
    sectionDimensions: ReadonlyMap<number, PageDimensions>,
  ): PageInfo[] {
    const pages: PageInfo[] = [];
    let currentContent: HTMLElement[] = [];
    let pageNumber = startPageNumber;
    let pageSectionIndex = sectionIndex;
    // A continuous section can begin on a page owned by its predecessor. That shared physical
    // page is still page 1 of the later section, so its first independently owned page is page 2.
    const sectionStartPages = new Map<number, number>([[sectionIndex, startPageNumber]]);
    const pageInSection = (owner = pageSectionIndex, physicalPage = pageNumber) =>
      physicalPage - (sectionStartPages.get(owner) ?? physicalPage) + 1;
    const dimensionsFor = (owner = pageSectionIndex) =>
      sectionDimensions.get(owner) ?? dims;
    const markSectionPlaced = (owner: number) => {
      if (!sectionStartPages.has(owner)) sectionStartPages.set(owner, pageNumber);
    };
    const previousBox = this.containerElement.querySelector<HTMLElement>(
      `.${this.cssPrefix}box:last-of-type`,
    );
    let precedingDisplayedPageNumber = parseInt(
      previousBox?.dataset.displayedPageNumber ?? "0",
      10,
    );
    const displayedPageNumber = (
      owner = pageSectionIndex,
      ownerPageInSection = pageInSection(owner),
      physicalPage = pageNumber,
    ) => {
      const numbering = this.pageNumbering.get(owner) ?? {};
      if (numbering.start !== undefined) {
        return numbering.start + ownerPageInSection - 1;
      }
      const ownerNumbering = this.pageNumbering.get(pageSectionIndex) ?? {};
      const currentPhysicalNumber = ownerNumbering.start !== undefined
        ? ownerNumbering.start + pageInSection(pageSectionIndex) - 1
        : precedingDisplayedPageNumber + 1;
      return currentPhysicalNumber + physicalPage - pageNumber;
    };

    // Get effective content height for first page (accounts for header/footer sizes)
    let { bodyHeight: effectiveContentHeight } = this.getPageBands(
      dimensionsFor(), pageSectionIndex, pageInSection()
    );
    let remainingHeight = effectiveContentHeight;

    // Track the previous block's bottom margin for margin collapsing
    let prevMarginBottomPt = 0;
    // Track footnote IDs for the current page
    let currentFootnoteIds: string[] = [];
    let currentPageHasFootnoteReference = false;
    // Track height consumed by footnotes on current page
    let currentFootnoteHeight = 0;
    // Track footnote continuation for current page (from previous page)
    let currentContinuation: FootnoteContinuation | null = this.pendingFootnoteContinuation;
    // Track any new continuation that will carry to next page
    let nextPageContinuation: FootnoteContinuation | null = null;
    let currentContinuationPartitioned = false;
    let currentPageAdmitted = false;
    /**
     * Whole notes that could not be started on this page and must render on the next one.
     *
     * `nextPageContinuation` is a single slot, but the code that fills it runs once per note in a
     * page's citation list — so when two notes on the same page both failed to fit, the second
     * overwrote the first and that note was never rendered anywhere. On a 94-footnote document
     * four notes disappeared from the output entirely. A deferral queue keeps the single-slot
     * continuation for its real meaning (the tail of a note that was SPLIT) and carries
     * never-started notes forward as ordinary footnote ids, which the next page already knows how
     * to lay out.
     */
    let deferredFootnoteIds: string[] = [];
    // Track partial footnotes for current page (footnotes that were split)
    let currentPartialFootnotes: PartialFootnote[] = [];

    const adoptEmptyPageOwner = (owner: number) => {
      pageSectionIndex = owner;
      const bands = this.getPageBands(
        dimensionsFor(owner),
        owner,
        pageInSection(owner),
      );
      effectiveContentHeight = bands.bodyHeight;
      remainingHeight = effectiveContentHeight;
    };

    const admitCurrentPage = () => {
      if (currentPageAdmitted) return;
      this.admitPageAllocation();
      currentPageAdmitted = true;
    };

    const prepareCurrentContinuation = () => {
      if (currentContinuationPartitioned
          || (currentContinuation?.remainingElements.length ?? 0) === 0) return;
      // A continuation guarantees that this physical page will exist. Admit it
      // before touching the carried DOM, then partition once and retain that
      // exact head while body placement is decided.
      admitCurrentPage();
      const ownedDimensions = dimensionsFor();
      const pageBands = this.getPageBands(
        ownedDimensions,
        pageSectionIndex,
        pageInSection(),
      );
      const partition = this.splitContinuationForPage(
        currentContinuation!,
        pageBands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO,
        ownedDimensions.contentWidth,
      );
      currentContinuation = partition.current;
      if (partition.overflow) nextPageContinuation = partition.overflow;
      currentContinuationPartitioned = true;
      currentFootnoteHeight = this.measureFootnotesHeight(
        currentFootnoteIds,
        ownedDimensions.contentWidth,
        currentContinuation,
        currentPartialFootnotes,
      );
    };

    const finishPage = (
      nextPageSectionIndex = pageSectionIndex,
      forceEmptyPage = false,
    ) => {
      const hasCurrentContinuation =
        (currentContinuation?.remainingElements.length ?? 0) > 0;
      if (!forceEmptyPage && currentContent.length === 0 && currentFootnoteIds.length === 0
          && !hasCurrentContinuation) {
        adoptEmptyPageOwner(nextPageSectionIndex);
        return;
      }

      prepareCurrentContinuation();
      // Pages without a carried continuation are charged here. Continuation
      // pages were charged before their page-owned partition above.
      admitCurrentPage();

      let pageContinuation = currentContinuation;
      const ownedPageInSection = pageInSection();
      const ownedDimensions = dimensionsFor();
      const ownedDisplayedPageNumber = displayedPageNumber(
        pageSectionIndex,
        ownedPageInSection,
      );
      const pageBands = this.getPageBands(
        ownedDimensions,
        pageSectionIndex,
        ownedPageInSection,
      );
      const maxFootnoteHeight = pageBands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO;
      // `prepareCurrentContinuation` has already packed the exact head for this
      // page. Keeping that partition stable lets body fit decisions share the
      // remaining band without rescanning the entire tail.
      pageContinuation = currentContinuation;

      // Recompute the entire painted payload after continuation partitioning.
      // Measuring only the carried head here discarded the reserve for a newer
      // whole/partial note and let the clipped note band silently consume it.
      currentFootnoteHeight = this.measureFootnotesHeight(
        currentFootnoteIds,
        ownedDimensions.contentWidth,
        pageContinuation,
        currentPartialFootnotes,
      );

      // A final body page can seed the next page with several whole notes. Partition that queue
      // before materializing a note-only page: addPageFootnotes deliberately clips its band, so
      // placing the entire queue in one page would keep the DOM nodes while making later notes
      // invisible (and therefore absent from a geometry-clipped PageMap).
      if (currentContent.length === 0 && currentFootnoteIds.length > 0
          && currentPartialFootnotes.length === 0) {
        const fittingIds: string[] = [];
        for (let index = 0; index < currentFootnoteIds.length; index++) {
          const footnoteId = currentFootnoteIds[index];
          const candidateIds = [...fittingIds, footnoteId];
          const candidateHeight = this.measureFootnotesHeight(
            candidateIds, ownedDimensions.contentWidth, pageContinuation);
          const guardedCandidateHeight = candidateHeight + FOOTNOTE_MEASUREMENT_GUARD_PT;
          if (guardedCandidateHeight <= maxFootnoteHeight) {
            fittingIds.push(footnoteId);
            currentFootnoteHeight = guardedCandidateHeight;
            continue;
          }

          const hasPageContinuation =
            (pageContinuation?.remainingElements.length ?? 0) > 0;
          if (fittingIds.length === 0 && !hasPageContinuation) {
            // One note alone is taller than the note band. Split at the same safe paragraph
            // boundaries used during body flow; if it is indivisible, preserve the established
            // visible clipped fallback while still advancing the queue.
            const source = this.footnoteRegistry.get(footnoteId);
            const split = source
              ? this.splitFootnoteToFit(
                source,
                maxFootnoteHeight - FOOTNOTE_MEASUREMENT_GUARD_PT,
                ownedDimensions.contentWidth,
                true,
              )
              : null;
            fittingIds.push(footnoteId);
            if (source && split && split.fits.length > 0 && split.overflow.length > 0) {
              currentPartialFootnotes.push({ footnoteId, fittingElements: split.fits });
              nextPageContinuation = {
                footnoteId,
                sourceAnchorId: source.dataset.sourceAnchorId,
                remainingElements: split.overflow,
              };
              currentFootnoteHeight = maxFootnoteHeight;
            } else {
              currentFootnoteHeight = guardedCandidateHeight;
            }
            deferredFootnoteIds.push(...currentFootnoteIds.slice(index + 1));
          } else {
            deferredFootnoteIds.push(...currentFootnoteIds.slice(index));
          }
          break;
        }
        currentFootnoteIds = fittingIds;
      }

      const page = this.createPage(
        ownedDimensions,
        pageNumber,
        pageSectionIndex,
        ownedDisplayedPageNumber,
        currentContent,
        ownedPageInSection,
        currentFootnoteIds,
        currentFootnoteHeight,
        pageContinuation,
        currentPartialFootnotes.length > 0 ? currentPartialFootnotes : undefined,
        false,
        currentPageAdmitted,
      );
      pages.push(page);
      precedingDisplayedPageNumber = ownedDisplayedPageNumber;

      pageNumber++;
      currentContent = [];

      pageSectionIndex = nextPageSectionIndex;
      if (!sectionStartPages.has(pageSectionIndex)) {
        // A carried note can own pages before the first body block of a new
        // section lands. Count those physical pages in that section so first /
        // odd / even stories and restarted displayed numbering advance once.
        sectionStartPages.set(pageSectionIndex, pageNumber);
      }

      // Get effective content height for new page position
      const newBands = this.getPageBands(
        dimensionsFor(),
        pageSectionIndex,
        pageInSection(),
      );
      effectiveContentHeight = newBands.bodyHeight;
      remainingHeight = effectiveContentHeight;

      prevMarginBottomPt = 0; // Reset margin tracking for new page
      currentFootnoteIds = []; // Reset footnotes for new page
      currentPageHasFootnoteReference = false;
      currentPartialFootnotes = []; // Reset partial footnotes for new page

      // Carry over continuation to next page
      currentContinuation = nextPageContinuation;
      nextPageContinuation = null;
      currentContinuationPartitioned = false;
      currentPageAdmitted = false;

      // Notes that never got started land at the top of the new page's note area. They are
      // ordinary footnotes from here on, so the normal fitting path handles them — and because
      // this page is fresh, the space they were denied now exists.
      if (deferredFootnoteIds.length > 0) {
        currentFootnoteIds = [...deferredFootnoteIds];
        deferredFootnoteIds = [];
      }

      // Whole deferred notes have no carried DOM and can be measured directly.
      // A continuation is partitioned lazily after the prospective page has
      // passed its resource admission check, so never clone its entire tail here.
      const nextDimensions = dimensionsFor();
      currentFootnoteHeight = currentContinuation
        ? 0
        : this.measureFootnotesHeight(currentFootnoteIds, nextDimensions.contentWidth);
    };

    const pageHasPayload = () => currentContent.length > 0
      || currentFootnoteIds.length > 0
      || (currentContinuation?.remainingElements.length ?? 0) > 0;

    for (let i = 0; i < blocks.length; i++) {
      this.checkpoint();
      const block = blocks[i];
      const allBlockFootnoteIds = this.extractFootnoteRefs(block.element);

      // Word normally lets a same-box continuous section begin on its predecessor's page. By
      // default, a footnote reference before that boundary promotes it to a page break. The
      // footnoteLayoutLikeWW8 compatibility switch permits post-break paragraphs without their own
      // references to remain on the shared page; keep checking until the page turns so a later
      // referenced paragraph still begins on a fresh page.
      if (block.sectionIndex !== pageSectionIndex && currentPageHasFootnoteReference
        && (!this.footnoteLayoutLikeWord8 || allBlockFootnoteIds.length > 0)) {
        finishPage(block.sectionIndex);
      }

      const pageIsEmpty = currentContent.length === 0
        && currentFootnoteIds.length === 0
        && (currentContinuation?.remainingElements.length ?? 0) === 0;
      if (pageIsEmpty && block.sectionIndex !== pageSectionIndex) {
        adoptEmptyPageOwner(block.sectionIndex);
      }

      prepareCurrentContinuation();

      if (currentFootnoteHeight > 0) {
        const ownerDimensions = dimensionsFor();
        const ownerBands = this.getPageBands(
          ownerDimensions,
          pageSectionIndex,
          pageInSection(),
        );
        if (currentFootnoteHeight > ownerBands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO) {
          // Pack an oversized carried/queued payload before admitting another
          // body block. finishPage partitions it into visible note-only bands.
          finishPage(block.sectionIndex);
          i--;
          continue;
        }
      }
      const blockDimensions = dimensionsFor(block.sectionIndex);
      const nextBlockPageInSection = pageInSection(block.sectionIndex, pageNumber + 1);
      const freshBlockPageBodyHeight = this.getPageBands(
        blockDimensions,
        block.sectionIndex,
        nextBlockPageInSection,
      ).bodyHeight;

      // Handle explicit page breaks
      if (block.isPageBreak) {
        finishPage(block.sectionIndex);
        continue;
      }

      // Handle page break before
      if (block.pageBreakBefore && currentContent.length > 0) {
        finishPage(block.sectionIndex);
      }

      // A series of keep-with-next blocks is one indivisible placement unit
      // when it can fit a new page. The former one-block lookahead was never
      // applied to the placement decision, and it could not preserve a chain of
      // headings/paragraphs. Do not force an oversized chain to a fresh page:
      // its members retain the established greedy/overflow behavior instead.
      const previousBlock = blocks[i - 1];
      const startsKeepChain =
        !previousBlock ||
        !previousBlock.keepWithNext ||
        previousBlock.isPageBreak ||
        block.pageBreakBefore;
      // A footnote continuation also occupies this page even when its body has
      // no blocks yet. In that case a feasible chain may need to move past the
      // continuation as a unit.
      const pageHasOccupiedSpace = currentContent.length > 0 || currentFootnoteHeight > 0;
      if (pageHasOccupiedSpace && block.keepWithNext && startsKeepChain) {
        const keepChain = this.getKeepWithNextChain(blocks, i);
        if (keepChain.length > 1) {
          const newChainFootnoteIds = this.collectNewFootnoteIds(
            keepChain,
            currentFootnoteIds
          );
          let additionalChainFootnoteHeight = 0;
          if (newChainFootnoteIds.length > 0 && this.footnoteRegistry.size > 0) {
            const totalChainFootnoteHeight = this.measureFootnotesHeight(
              [...currentFootnoteIds, ...newChainFootnoteIds],
              blockDimensions.contentWidth,
              currentContinuation
            );
            additionalChainFootnoteHeight = Math.max(
              0,
              totalChainFootnoteHeight - currentFootnoteHeight
            );
          }

          const currentChainHeight =
            this.measureKeepWithNextChainBodyHeight(
              keepChain,
              prevMarginBottomPt,
              currentContent.length === 0,
              pageInSection(block.sectionIndex)
            ) +
            additionalChainFootnoteHeight;
          const currentAvailableHeight = remainingHeight - currentFootnoteHeight;

          if (currentChainHeight > currentAvailableHeight) {
            const nextPageBands = this.getPageBands(
              dimensionsFor(block.sectionIndex),
              block.sectionIndex,
              pageInSection(block.sectionIndex, pageNumber + 1),
            );
            const freshChainBodyHeight = this.measureKeepWithNextChainBodyHeight(
              keepChain,
              0,
              true,
              pageInSection(block.sectionIndex, pageNumber + 1)
            );
            // finishPage transfers this continuation to the new page's
            // currentContinuation state, so include it in the destination
            // page's footnote reservation before deciding to move the chain.
            const freshChainFootnoteHeight = this.measureFootnotesHeight(
              newChainFootnoteIds,
              blockDimensions.contentWidth,
              nextPageContinuation
            );

            if (
              freshChainBodyHeight + freshChainFootnoteHeight <=
              nextPageBands.bodyHeight
            ) {
              finishPage(block.sectionIndex);
            }
          }
        }
      }

      // Extract footnote references from this block
      // Only count new footnotes (not already on this page)
      const newFootnoteIds = this.collectNewFootnoteIds([block], currentFootnoteIds);

      // Calculate additional footnote height if this block is added
      let combinedFootnoteHeight = currentFootnoteHeight;
      let additionalFootnoteHeight = 0;
      if (newFootnoteIds.length > 0 && this.footnoteRegistry.size > 0) {
        // Measure the combined height of all footnotes that would be on this page
        // (including any continuation)
        const combinedFootnoteIds = [...currentFootnoteIds, ...newFootnoteIds];
        combinedFootnoteHeight = this.measureFootnotesHeight(
          combinedFootnoteIds,
          blockDimensions.contentWidth,
          currentContinuation,
          currentPartialFootnotes,
        );
        additionalFootnoteHeight = Math.max(
          0,
          combinedFootnoteHeight - currentFootnoteHeight,
        );
      }

      // Calculate the effective height this block will consume
      // Account for margin collapsing: the gap between blocks is max(prevBottom, currTop), not sum
      const isFirstOnPage = currentContent.length === 0;
      const effectiveMarginTop = this.effectiveBlockMarginTop(
        block,
        prevMarginBottomPt,
        isFirstOnPage,
        pageInSection(block.sectionIndex)
      );
      // Visible height = top margin gap + content + footnote space
      // Note: bottom margin is NOT included in the fit check because the last block's
      // bottom margin extends beyond the content area and is clipped by overflow:hidden.
      // It is still tracked in remainingHeight for correct margin collapsing with the next block.
      const blockSpace = effectiveMarginTop + block.heightPt + additionalFootnoteHeight;

      // Effective remaining height (content area minus footnotes already on page)
      const effectiveRemainingHeight = remainingHeight - currentFootnoteHeight;

      // Calculate maximum footnote area for this page (can expand into body content space)
      const bodyContentUsed = effectiveContentHeight - remainingHeight;
      const maxFootnoteArea = effectiveContentHeight * MAX_FOOTNOTE_AREA_RATIO;

      // A paragraph that cannot fit as a whole may still have a simple text-only
      // prefix that fits this page. Fragment before the ordinary next-page or
      // oversized fallback so the cloned head participates in the same margin
      // and footnote accounting as every other block.
      if (blockSpace > effectiveRemainingHeight) {
        const paragraphFragments = this.tryFragmentParagraph(
          block,
          blockDimensions,
          effectiveRemainingHeight,
          effectiveMarginTop
        );
        if (paragraphFragments) {
          blocks.splice(i, 1, ...paragraphFragments);
          i--;
          continue;
        }
      }

      // Check if block fits on current page (including its footnotes)
      if (
        blockSpace <= effectiveRemainingHeight
        && (combinedFootnoteHeight === 0
          || combinedFootnoteHeight + FOOTNOTE_MEASUREMENT_GUARD_PT <= maxFootnoteArea)
      ) {
        // Block fits with current footnote allocation
        markSectionPlaced(block.sectionIndex);
        currentContent.push(this.cloneBlockForPage(
          block,
          isFirstOnPage,
          pageInSection(block.sectionIndex),
        ));
        remainingHeight -= (effectiveMarginTop + block.heightPt + block.marginBottomPt);
        prevMarginBottomPt = block.marginBottomPt;
        // Add new footnotes to current page
        if (newFootnoteIds.length > 0) {
          currentFootnoteIds.push(...newFootnoteIds);
          currentFootnoteHeight = combinedFootnoteHeight;
        }
        currentPageHasFootnoteReference ||= allBlockFootnoteIds.length > 0;
      } else if (
        block.heightPt + this.effectiveBlockMarginTop(
          block,
          0,
          true,
          pageInSection(block.sectionIndex, pageNumber + 1),
        ) <= freshBlockPageBodyHeight
      ) {
        // Block doesn't fit with current allocation - try expanding footnote area
        const blockSpaceWithoutFootnotes = effectiveMarginTop + block.heightPt;

        if (newFootnoteIds.length > 0 && blockSpaceWithoutFootnotes <= effectiveRemainingHeight) {
          // Block itself fits, but footnotes don't - expand footnote area
          markSectionPlaced(block.sectionIndex);
          currentContent.push(this.cloneBlockForPage(
            block,
            isFirstOnPage,
            pageInSection(block.sectionIndex),
          ));
          remainingHeight -= (effectiveMarginTop + block.heightPt + block.marginBottomPt);
          prevMarginBottomPt = block.marginBottomPt;

          // Calculate EXPANDED space available for footnotes
          // Footnotes can take up to maxFootnoteArea or all remaining space, whichever is less
          const availableForFootnotes = Math.min(
            maxFootnoteArea,
            effectiveContentHeight - bodyContentUsed - blockSpaceWithoutFootnotes
          );

          // Try to fit as much of each new footnote as possible in expanded area
          for (let noteIndex = 0; noteIndex < newFootnoteIds.length; noteIndex++) {
            this.checkpoint();
            const footnoteId = newFootnoteIds[noteIndex];
            const footnote = this.footnoteRegistry.get(footnoteId);
            if (!footnote) continue;

            // One page can carry only one split tail. Once that slot is used,
            // keep every later note in source order on the deferral queue.
            if (nextPageContinuation) {
              deferredFootnoteIds.push(...newFootnoteIds.slice(noteIndex));
              break;
            }

            const candidateIds = [...currentFootnoteIds, footnoteId];
            const candidateHeight = this.measureFootnotesHeight(
              candidateIds,
              blockDimensions.contentWidth,
              currentContinuation,
              currentPartialFootnotes,
            );
            const spaceLeftForFootnotes = availableForFootnotes - currentFootnoteHeight;

            if (candidateHeight + FOOTNOTE_MEASUREMENT_GUARD_PT <= availableForFootnotes) {
              // Whole footnote fits in expanded area
              currentFootnoteIds.push(footnoteId);
              currentFootnoteHeight = candidateHeight;
            } else {
              // Footnote needs to be split - use all available expanded space
              if (spaceLeftForFootnotes > 20) { // Minimum space to start a footnote
                const { fits, overflow } = this.splitFootnoteToFit(
                  footnote,
                  availableForFootnotes,
                  blockDimensions.contentWidth,
                  false,
                  {
                    footnoteIds: currentFootnoteIds,
                    continuation: currentContinuation,
                    partialFootnotes: currentPartialFootnotes,
                  },
                );

                if (fits.length > 0) {
                  // Add partial footnote to current page
                  currentFootnoteIds.push(footnoteId);
                  currentPartialFootnotes.push({
                    footnoteId,
                    fittingElements: fits
                  });
                  if (overflow.length > 0) {
                    nextPageContinuation = {
                      footnoteId,
                      sourceAnchorId: footnote.dataset.sourceAnchorId,
                      remainingElements: overflow
                    };
                  }
                  currentFootnoteHeight = this.measureFootnotesHeight(
                    currentFootnoteIds,
                    blockDimensions.contentWidth,
                    currentContinuation,
                    currentPartialFootnotes,
                  );
                  if (overflow.length > 0) {
                    deferredFootnoteIds.push(...newFootnoteIds.slice(noteIndex + 1));
                    break;
                  }
                } else {
                  // Nothing of this note fits: defer the WHOLE note rather than assigning the
                  // single continuation slot, which a later note on this page would overwrite.
                  deferredFootnoteIds.push(...newFootnoteIds.slice(noteIndex));
                  break;
                }
              } else {
                // Not enough space to even start the note — same deferral.
                deferredFootnoteIds.push(...newFootnoteIds.slice(noteIndex));
                break;
              }
            }
          }
          currentPageHasFootnoteReference ||= allBlockFootnoteIds.length > 0;
        } else {
          // Block itself doesn't fit. Retry after the page turn so carried
          // continuations and deferred notes participate in the normal fit and
          // split decisions instead of being overwritten by forced placement.
          finishPage(block.sectionIndex, !pageHasPayload());
          i--;
          continue;
        }
      } else {
        // Block is taller than a page. Ordinary tables can be split at complete
        // row boundaries; every other block retains the established overflow path.
        const tableFragments = this.trySplitSimpleOversizedTable(
          block,
          dimensionsFor(block.sectionIndex),
          block.sectionIndex,
        );
        if (tableFragments) {
          if (currentContent.length > 0 || currentContinuation) {
            finishPage(block.sectionIndex);
          }
          blocks.splice(i, 1, ...tableFragments);
          i--;
          continue;
        }

        // Unsupported oversized blocks are intentionally left intact. Splitting
        // arbitrary HTML, merged tables, or footnote-bearing tables would be less
        // correct than the prior clipped fallback.
        if (pageHasPayload()) {
          finishPage(block.sectionIndex);
          i--;
          continue;
        }
        markSectionPlaced(block.sectionIndex);
        currentContent.push(this.cloneBlockForPage(
          block,
          true,
          pageInSection(block.sectionIndex),
        ));
        // An unsupported body block already occupies/clips its whole body band.
        // Preserve its notes losslessly on following note-only pages rather
        // than drawing them underneath that clipped fallback.
        deferredFootnoteIds.push(...newFootnoteIds);
        currentPageHasFootnoteReference ||= allBlockFootnoteIds.length > 0;
        finishPage(block.sectionIndex);
      }
    }

    // Finish last page
    finishPage();

    // A split created while finishing the final body page still needs a page substrate.
    // Drain all remaining note paragraphs into footnote-only continuation pages.
    while (
      currentFootnoteIds.length > 0
      || deferredFootnoteIds.length > 0
      || (currentContinuation?.remainingElements.length ?? 0) > 0
    ) {
      finishPage();
    }

    // Store any remaining continuation for next section
    this.pendingFootnoteContinuation = nextPageContinuation;

    return pages;
  }


  /**
   * Strip block addressing from a header/footer node cloned into a page box.
   *
   * A running story is authored ONCE and cloned onto every page, so the clones all carry the same
   * `data-anchor` — on this document, 42 page boxes claiming one footer paragraph. Left editable,
   * committing any one of them writes back through that single shared anchor, and the per-page
   * page-number substitution makes it worse: each clone shows a DIFFERENT number, so a commit
   * writes that page's number into the story as literal text and destroys the PAGE field.
   *
   * Page-box header/footer content is presentation. The docked editing bands
   * (`editor-headerfooter.ts`) are the addressable affordance, and they exist precisely because a
   * cloned node cannot be uniquely addressed.
   */
  private makeClonedStoryInert(root: HTMLElement): void {
    const nodes = [root, ...Array.from(root.querySelectorAll<HTMLElement>("*"))];
    for (const el of nodes) {
      el.removeAttribute("data-anchor");
      el.removeAttribute("data-committed-text");
      if (el.getAttribute("contenteditable") !== null) el.setAttribute("contenteditable", "false");
    }
  }

  /** A repeated margin note is presentation, not a second bookmark/link target. */
  private makeClonedMarginCommentInert(root: HTMLElement): void {
    this.makeClonedStoryInert(root);
    const nodes = [root, ...Array.from(root.querySelectorAll<HTMLElement>("*"))];
    for (const element of nodes) {
      element.removeAttribute("id");
      if (element.localName === "a" && element.getAttribute("href")?.startsWith("#")) {
        element.removeAttribute("href");
        element.setAttribute("aria-disabled", "true");
        element.tabIndex = -1;
      }
    }
  }

  /**
   * Resolves floating DrawingML objects after their anchor paragraphs have landed on a page.
   *
   * The converter deliberately carries the OOXML bases and offsets as data instead of flattening
   * them into CSS: page/margin/column bases belong to the page, while paragraph/line/character
   * bases belong to the laid-out anchor. Once both coordinate systems exist, promote the object
   * from the clipped text column into the page box and position it in one shared point space.
   */
  private positionDrawingAnchors(
    pageBox: HTMLElement,
    contentArea: HTMLElement,
    dims: PageDimensions,
    pageNumber: number,
  ): void {
    const anchors = Array.from(
      contentArea.querySelectorAll<HTMLElement>('[data-docx-drawing-anchor="true"]'),
    );
    if (anchors.length === 0) return;

    const pageRect = pageBox.getBoundingClientRect();
    const pixelsPerPoint = pageRect.width / dims.pageWidth;
    if (!(pixelsPerPoint > 0)) return;

    for (const anchor of anchors) {
      const staticRect = anchor.getBoundingClientRect();
      const paragraph = anchor.closest<HTMLElement>('p, h1, h2, h3, h4, h5, h6');
      const paragraphRect = paragraph?.getBoundingClientRect() ?? staticRect;
      const toPageX = (x: number) => (x - pageRect.left) / pixelsPerPoint;
      const toPageY = (y: number) => (y - pageRect.top) / pixelsPerPoint;

      const context = {
        staticLeft: toPageX(staticRect.left),
        staticTop: toPageY(staticRect.top),
        lineHeight: paragraph
          ? (parseFloat(this.view.getComputedStyle(paragraph).lineHeight) || staticRect.height) / pixelsPerPoint
          : staticRect.height / pixelsPerPoint,
        paragraphLeft: toPageX(paragraphRect.left),
        paragraphTop: toPageY(paragraphRect.top),
        paragraphWidth: paragraphRect.width / pixelsPerPoint,
        paragraphHeight: paragraphRect.height / pixelsPerPoint,
      };

      const widthReference = this.horizontalAnchorReference(
        anchor.getAttribute('data-docx-anchor-width-relative') ?? 'margin',
        dims,
        pageNumber,
        context,
      );
      const widthPercent = this.anchorNumber(anchor, 'width-percent');
      if (widthPercent !== undefined) {
        anchor.style.width = `${widthReference.size * widthPercent / 100}pt`;
      }

      const heightReference = this.verticalAnchorReference(
        anchor.getAttribute('data-docx-anchor-height-relative') ?? 'margin',
        dims,
        pageNumber,
        context,
      );
      const heightPercent = this.anchorNumber(anchor, 'height-percent');
      if (heightPercent !== undefined && anchor.dataset.docxAnchorAutofit !== 'true') {
        anchor.style.height = `${heightReference.size * heightPercent / 100}pt`;
      }

      // Read the final border-box size after relative sizing and before reparenting. The stored
      // wp:extent remains the fallback; wps:bodyPr insets stay inside it through border-box sizing.
      const sizedRect = anchor.getBoundingClientRect();
      const width = sizedRect.width / pixelsPerPoint;
      const height = sizedRect.height / pixelsPerPoint;
      const horizontal = this.horizontalAnchorReference(
        anchor.getAttribute('data-docx-anchor-h-relative') ?? 'column',
        dims,
        pageNumber,
        context,
      );
      const vertical = this.verticalAnchorReference(
        anchor.getAttribute('data-docx-anchor-v-relative') ?? 'paragraph',
        dims,
        pageNumber,
        context,
      );
      const left = this.resolveAnchorAxis(
        horizontal,
        width,
        anchor.getAttribute('data-docx-anchor-h-align'),
        this.anchorNumber(anchor, 'h-offset'),
        pageNumber,
      );
      const top = this.resolveAnchorAxis(
        vertical,
        height,
        anchor.getAttribute('data-docx-anchor-v-align'),
        this.anchorNumber(anchor, 'v-offset'),
        pageNumber,
      );

      pageBox.appendChild(anchor);
      anchor.style.position = 'absolute';
      anchor.style.left = `${left}pt`;
      anchor.style.top = `${top}pt`;
      anchor.style.margin = '0';
    }
  }

  private anchorNumber(anchor: HTMLElement, suffix: string): number | undefined {
    const raw = anchor.getAttribute(`data-docx-anchor-${suffix}`);
    if (raw === null) return undefined;
    const value = Number(raw);
    return Number.isFinite(value) ? value : undefined;
  }

  private horizontalAnchorReference(
    relativeFrom: string,
    dims: PageDimensions,
    pageNumber: number,
    context: {
      staticLeft: number;
      paragraphLeft: number;
      paragraphWidth: number;
    },
  ): { start: number; size: number } {
    const leftMargin = { start: 0, size: dims.marginLeft };
    const rightMargin = {
      start: dims.marginLeft + dims.contentWidth,
      size: dims.marginRight,
    };
    switch (relativeFrom) {
      case 'page':
        return { start: 0, size: dims.pageWidth };
      case 'character':
        return { start: context.staticLeft, size: 0 };
      case 'leftMargin':
        return leftMargin;
      case 'rightMargin':
        return rightMargin;
      case 'insideMargin':
        return pageNumber % 2 === 1 ? leftMargin : rightMargin;
      case 'outsideMargin':
        return pageNumber % 2 === 1 ? rightMargin : leftMargin;
      case 'paragraph':
        // Not part of ST_RelFromH, but accepting it gives malformed/legacy producers the
        // intuitive base without conflating it with the whole text column.
        return { start: context.paragraphLeft, size: context.paragraphWidth };
      case 'column':
      case 'margin':
      default:
        // The paginator currently lays out one text column per section, so the column and page
        // margin boxes coincide. Keeping the names distinct here leaves multi-column geometry
        // localized to this resolver when that layout support is added.
        return { start: dims.marginLeft, size: dims.contentWidth };
    }
  }

  private verticalAnchorReference(
    relativeFrom: string,
    dims: PageDimensions,
    pageNumber: number,
    context: {
      staticTop: number;
      lineHeight: number;
      paragraphTop: number;
      paragraphHeight: number;
    },
  ): { start: number; size: number } {
    const topMargin = { start: 0, size: dims.marginTop };
    const bottomMargin = {
      start: dims.marginTop + dims.contentHeight,
      size: dims.marginBottom,
    };
    switch (relativeFrom) {
      case 'page':
        return { start: 0, size: dims.pageHeight };
      case 'paragraph':
        return { start: context.paragraphTop, size: context.paragraphHeight };
      case 'line':
        return { start: context.staticTop, size: context.lineHeight };
      case 'topMargin':
        return topMargin;
      case 'bottomMargin':
        return bottomMargin;
      case 'insideMargin':
        return pageNumber % 2 === 1 ? topMargin : bottomMargin;
      case 'outsideMargin':
        return pageNumber % 2 === 1 ? bottomMargin : topMargin;
      case 'margin':
      default:
        return { start: dims.marginTop, size: dims.contentHeight };
    }
  }

  private resolveAnchorAxis(
    reference: { start: number; size: number },
    objectSize: number,
    alignment: string | null,
    offset: number | undefined,
    pageNumber: number,
  ): number {
    if (offset !== undefined) return reference.start + offset;

    let fraction = 0;
    if (alignment === 'center') fraction = 0.5;
    else if (alignment === 'right' || alignment === 'bottom') fraction = 1;
    else if (alignment === 'inside') fraction = pageNumber % 2 === 1 ? 0 : 1;
    else if (alignment === 'outside') fraction = pageNumber % 2 === 1 ? 1 : 0;
    return reference.start + (reference.size - objectSize) * fraction;
  }

  /** Charge a physical page before any page-owned partitioning or DOM allocation. */
  private admitPageAllocation(): void {
    const prospectivePageCount = this.createdPageCount + 1;
    this.pageCountCheckpoint?.(prospectivePageCount);
    this.createdPageCount = prospectivePageCount;
  }

  /**
   * Creates a page container element.
   */
  private createPage(
    dims: PageDimensions,
    pageNumber: number,
    sectionIndex: number,
    displayedPageNumber: number,
    content: HTMLElement[],
    pageInSection: number,
    footnoteIds: string[] = [],
    footnoteHeight: number = 0,
    continuation?: FootnoteContinuation | null,
    partialFootnotes?: PartialFootnote[],
    isSectionFiller = false,
    pageAlreadyAdmitted = false,
  ): PageInfo {
    if (!pageAlreadyAdmitted) this.admitPageAllocation();
    // Create page box at full size, then scale the entire box
    // This ensures proper clipping and consistent scaling of all elements
    const pageBox = this.document.createElement("div");
    pageBox.className = `${this.cssPrefix}box`;
    pageBox.style.width = `${dims.pageWidth}pt`;
    pageBox.style.height = `${dims.pageHeight}pt`;
    pageBox.style.overflow = "hidden";
    pageBox.style.position = "relative";
    // Use CSS zoom for better text rendering when supported, fall back to transform
    // Zoom affects layout (no negative margin hack needed) and renders text more crisply
    // Note: zoom is non-standard but supported in all major browsers
    if (this.scale !== 1) {
      if (this.view.CSS?.supports("zoom", "1")) {
        pageBox.style.zoom = String(this.scale);
      } else {
        pageBox.style.transform = `scale(${this.scale})`;
        pageBox.style.transformOrigin = "top left";
        // Transform does not affect layout, so compensate for the natural box dimensions.
        const heightReductionPt = dims.pageHeight * (1 - this.scale);
        const widthReductionPt = dims.pageWidth * (1 - this.scale);
        const heightReductionPx = ptToPx(heightReductionPt);
        const widthReductionPx = ptToPx(widthReductionPt);
        pageBox.style.marginRight = `-${widthReductionPx}px`;
        pageBox.style.marginBottom = `${this.pageGap - heightReductionPx}px`;
      }
    }
    // Hint browser for GPU compositing and layout isolation
    pageBox.style.willChange = "transform";
    pageBox.style.contain = "layout paint";
    pageBox.dataset.pageNumber = String(pageNumber);
    pageBox.dataset.sectionIndex = String(sectionIndex);
    pageBox.dataset.displayedPageNumber = String(displayedPageNumber);
    if (isSectionFiller) pageBox.dataset.sectionFiller = "true";
    // Needed by substitutePageNumberFields: a section that restarts numbering counts from its own
    // first page, not from the document's.
    pageBox.dataset.pageInSection = String(pageInSection);

    // Where the three bands sit on this page (no re-measurement needed)
    const bands = this.getPageBands(dims, sectionIndex, pageInSection);

    // Add header if available for this section/page
    const headerSource = isSectionFiller
      ? undefined
      : this.selectHeader(sectionIndex, pageInSection);

    if (headerSource) {
      const headerDiv = this.document.createElement("div");
      headerDiv.className = `${this.cssPrefix}header`;
      headerDiv.style.position = "absolute";
      // `w:header` is the distance to the TOP of the story, and the story grows downward from
      // there; `flex-start` therefore pins the edge the OOXML actually declares, and content
      // taller than the measurement clips at the bottom rather than sliding up the page.
      headerDiv.style.top = `${bands.headerTop}pt`;
      headerDiv.style.bottom = "auto";
      headerDiv.style.left = `${dims.marginLeft}pt`;
      headerDiv.style.width = `${dims.contentWidth}pt`;
      headerDiv.style.height = `${bands.headerHeight}pt`;
      headerDiv.style.overflow = "hidden";
      headerDiv.style.boxSizing = "border-box";
      headerDiv.style.display = "flex";
      headerDiv.style.flexDirection = "column";
      headerDiv.style.justifyContent = "flex-start";
      // Clone the header content (skip the wrapper div's data attributes)
      for (const child of Array.from(headerSource.childNodes)) {
        const clonedheaderDiv = child.cloneNode(true) as HTMLElement;
        if (clonedheaderDiv.nodeType === 1) this.makeClonedStoryInert(clonedheaderDiv);
        headerDiv.appendChild(clonedheaderDiv);
      }
      pageBox.appendChild(headerDiv);
    }

    // Create content area using pre-computed effective heights
    const contentAreaTop = bands.bodyTop;
    const contentAreaHeight = bands.bodyHeight;

    const contentArea = this.document.createElement("div");
    contentArea.className = `${this.cssPrefix}content`;
    contentArea.style.position = "absolute";
    contentArea.style.top = `${contentAreaTop}pt`;
    contentArea.style.left = `${dims.marginLeft}pt`;
    contentArea.style.width = `${dims.contentWidth}pt`;
    contentArea.style.height = `${contentAreaHeight}pt`;
    contentArea.style.overflow = "hidden";

    // Add content
    for (const el of content) {
      contentArea.appendChild(el);
    }

    pageBox.appendChild(contentArea);

    // Add footnotes if any references appear on this page (or continuation from previous)
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if (!isSectionFiller && (footnoteIds.length > 0 || hasContinuation)) {
      this.addPageFootnotes(pageBox, footnoteIds, dims, bands, footnoteHeight, continuation, partialFootnotes);

    }

    // Materialize margin comments after footnotes exist so markers inside a
    // note participate in the same page-owned comment story as body markers.
    const pageCommentIds: string[] = [];
    for (const marker of Array.from(
      pageBox.querySelectorAll<HTMLElement>("[data-comment-id]"),
    )) {
      if (marker.closest(`.${this.cssPrefix}comment-margin`)) continue;
      const id = marker.dataset.commentId;
      if (id && this.commentMarginRegistry.has(id) && !pageCommentIds.includes(id)) {
        pageCommentIds.push(id);
      }
    }
    if (pageCommentIds.length > 0) {
      const marginColumn = this.document.createElement("aside");
      marginColumn.className = `${this.cssPrefix}comment-margin`;
      marginColumn.style.position = "absolute";
      marginColumn.style.top = `${contentAreaTop}pt`;
      marginColumn.style.left = `${dims.marginLeft + dims.contentWidth + 3}pt`;
      marginColumn.style.width = `${Math.max(12, dims.marginRight - 6)}pt`;
      marginColumn.style.maxHeight = `${contentAreaHeight}pt`;
      marginColumn.style.overflow = "hidden";
      marginColumn.style.boxSizing = "border-box";
      for (const id of pageCommentIds) {
        const source = this.commentMarginRegistry.get(id);
        if (source) {
          const clone = source.cloneNode(true) as HTMLElement;
          this.makeClonedMarginCommentInert(clone);
          marginColumn.appendChild(clone);
        }
      }
      pageBox.appendChild(marginColumn);
    }

    // Add footer if available for this section/page
    const footerSource = isSectionFiller
      ? undefined
      : this.selectFooter(sectionIndex, pageInSection);
    if (footerSource) {
      const footerDiv = this.document.createElement("div");
      footerDiv.className = `${this.cssPrefix}footer`;
      footerDiv.style.position = "absolute";
      // `w:footer` is the distance to the BOTTOM of the story, and the story grows upward from
      // there — the mirror of the header, so `flex-end` and a bottom anchor.
      footerDiv.style.top = "auto";
      footerDiv.style.bottom = `${dims.footerDistance}pt`;
      footerDiv.style.left = `${dims.marginLeft}pt`;
      footerDiv.style.width = `${dims.contentWidth}pt`;
      footerDiv.style.height = `${bands.footerHeight}pt`;
      footerDiv.style.overflow = "hidden";
      footerDiv.style.boxSizing = "border-box";
      footerDiv.style.display = "flex";
      footerDiv.style.flexDirection = "column";
      footerDiv.style.justifyContent = "flex-end";
      // Clone the footer content (skip the wrapper div's data attributes)
      for (const child of Array.from(footerSource.childNodes)) {
        const clonedfooterDiv = child.cloneNode(true) as HTMLElement;
        if (clonedfooterDiv.nodeType === 1) this.makeClonedStoryInert(clonedfooterDiv);
        footerDiv.appendChild(clonedfooterDiv);
      }
      pageBox.appendChild(footerDiv);
    }

    // Add page number (will be hidden by CSS if document has footer)
    if (this.showPageNumbers && !isSectionFiller) {
      const pageNum = this.document.createElement("div");
      pageNum.className = `${this.cssPrefix}number`;
      pageNum.textContent = String(pageNumber);
      pageBox.appendChild(pageNum);
    }

    // Add to container
    this.containerElement.appendChild(pageBox);

    // Floating objects need the final page and anchor-paragraph boxes. Resolve them only after
    // append, when getBoundingClientRect() is meaningful, and lift them out of content clipping.
    this.positionDrawingAnchors(pageBox, contentArea, dims, pageNumber);

    // Body and notes must occupy DISJOINT bands. The note block is absolutely positioned against
    // the page bottom and grows upward, while the content area spans the whole text height — so
    // nothing but the layout loop's arithmetic keeps them apart, and any disagreement between the
    // height the loop reserved and the height the notes actually render at draws body text and note
    // text on top of each other (observed on a 94-footnote document: ~134pt of superimposed,
    // illegible glyphs). Shrinking the content area to the space the notes actually left removes
    // the failure mode by construction — the worst case becomes a clean clip by the content area's
    // existing `overflow: hidden`, which is obvious and recoverable rather than silent corruption.
    //
    // This runs AFTER the append on purpose: `getBoundingClientRect()` on a detached node is all
    // zeroes, so measuring while building the page silently did nothing.
    const notesEl = pageBox.querySelector<HTMLElement>(`.${this.cssPrefix}footnotes`);
    if (notesEl) {
      const notesHeightPt = pxToPt(notesEl.getBoundingClientRect().height);
      if (notesHeightPt > 0) {
        contentArea.style.height = `${Math.max(0, contentAreaHeight - notesHeightPt)}pt`;
      }
    }

    return {
      pageNumber,
      sectionIndex,
      dimensions: dims,
      element: pageBox,
    };
  }
}

/**
 * Convenience function to paginate HTML content.
 *
 * @param html - HTML string with pagination metadata
 * @param container - Container element or ID where pages will be rendered
 * @param options - Pagination options
 * @returns PaginationResult
 *
 * @example
 * ```typescript
 * const html = await convertDocxToHtml(docx, { paginationMode: PaginationMode.Paginated });
 *
 * // Create a container for the paginated view
 * const container = document.getElementById('viewer');
 *
 * // Parse and paginate
 * container.innerHTML = html;
 * const staging = document.getElementById('pagination-staging');
 * const pageContainer = document.getElementById('pagination-container');
 *
 * const engine = new PaginationEngine(staging, pageContainer, { scale: 0.8 });
 * const result = engine.paginate();
 *
 * console.log(`Document has ${result.totalPages} pages`);
 * ```
 */
export function paginateHtml(
  html: string,
  container: HTMLElement | string,
  options: PaginationOptions = {}
): PaginationResult {
  const ownerDocument =
    typeof container === "string" ? globalThis.document : container.ownerDocument;
  const containerEl =
    typeof container === "string"
      ? (ownerDocument.getElementById(container) as HTMLElement)
      : container;

  if (!containerEl) {
    throw new Error("Container element not found");
  }

  // Insert HTML into container
  containerEl.innerHTML = html;

  // Find staging and page container
  const cssPrefix = options.cssPrefix ?? "page-";
  const staging = containerEl.querySelector<HTMLElement>("#pagination-staging") ||
    containerEl.querySelector<HTMLElement>(`.${cssPrefix}staging`);
  const pageContainer = containerEl.querySelector<HTMLElement>("#pagination-container") ||
    containerEl.querySelector<HTMLElement>(`.${cssPrefix}container`);

  if (!staging) {
    throw new Error(
      "Pagination staging element not found. Make sure the HTML was generated with PaginationMode.Paginated"
    );
  }

  if (!pageContainer) {
    throw new Error("Pagination container element not found");
  }

  // This convenience function is the read-only HTML viewer path. Keep the
  // lower-level engine conservative by default, while enabling paragraph
  // fragmentation here unless a caller deliberately opts out.
  const engine = new PaginationEngine(staging, pageContainer, {
    ...options,
    fragmentParagraphs: options.fragmentParagraphs ?? true,
  });
  return engine.paginate();
}

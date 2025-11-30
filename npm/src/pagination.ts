/**
 * Pagination engine for creating a PDF.js-style paginated view from HTML output.
 *
 * This module provides client-side pagination that measures rendered content
 * and flows it across fixed-size page containers based on document dimensions.
 */

/**
 * Page dimensions extracted from HTML data attributes (in points).
 */
export interface PageDimensions {
  /** Page width in points */
  pageWidth: number;
  /** Page height in points */
  pageHeight: number;
  /** Content area width (page minus margins) in points */
  contentWidth: number;
  /** Content area height (page minus margins) in points */
  contentHeight: number;
  /** Top margin in points */
  marginTop: number;
  /** Right margin in points */
  marginRight: number;
  /** Bottom margin in points */
  marginBottom: number;
  /** Left margin in points */
  marginLeft: number;
  /** Header distance from top of page in points */
  headerHeight: number;
  /** Footer distance from bottom of page in points */
  footerHeight: number;
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
  /** Total number of pages */
  totalPages: number;
  /** Array of page information */
  pages: PageInfo[];
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
}

// Default letter size in points (612 x 792 = 8.5" x 11")
const DEFAULT_PAGE_WIDTH = 612;
const DEFAULT_PAGE_HEIGHT = 792;
const DEFAULT_MARGIN = 72; // 1 inch

// Maximum percentage of content height that footnotes can occupy
// This allows footnotes to expand upward into body content space when needed
const MAX_FOOTNOTE_AREA_RATIO = 0.6; // 60% of content height

// Minimum body content height per page (to avoid pages with only footnotes)
const MIN_BODY_CONTENT_HEIGHT = 72; // 1 inch minimum body content

/**
 * Converts pixels to points (assuming 96 DPI screen).
 */
function pxToPt(px: number): number {
  return px * 0.75; // 72 points / 96 pixels
}

/**
 * Converts points to pixels (assuming 96 DPI screen).
 */
function ptToPx(pt: number): number {
  return pt / 0.75;
}

// Default header/footer distance (0.5 inch)
const DEFAULT_HEADER_FOOTER_HEIGHT = 36;

/**
 * Parses page dimensions from a section element's data attributes.
 */
function parseDimensions(section: HTMLElement): PageDimensions {
  const pageWidth = parseFloat(section.dataset.pageWidth || "") || DEFAULT_PAGE_WIDTH;
  const pageHeight = parseFloat(section.dataset.pageHeight || "") || DEFAULT_PAGE_HEIGHT;
  const contentWidth = parseFloat(section.dataset.contentWidth || "") || pageWidth - 2 * DEFAULT_MARGIN;
  const contentHeight = parseFloat(section.dataset.contentHeight || "") || pageHeight - 2 * DEFAULT_MARGIN;
  const marginTop = parseFloat(section.dataset.marginTop || "") || DEFAULT_MARGIN;
  const marginRight = parseFloat(section.dataset.marginRight || "") || DEFAULT_MARGIN;
  const marginBottom = parseFloat(section.dataset.marginBottom || "") || DEFAULT_MARGIN;
  const marginLeft = parseFloat(section.dataset.marginLeft || "") || DEFAULT_MARGIN;
  const headerHeight = parseFloat(section.dataset.headerHeight || "") || DEFAULT_HEADER_FOOTER_HEIGHT;
  const footerHeight = parseFloat(section.dataset.footerHeight || "") || DEFAULT_HEADER_FOOTER_HEIGHT;

  return {
    pageWidth,
    pageHeight,
    contentWidth,
    contentHeight,
    marginTop,
    marginRight,
    marginBottom,
    marginLeft,
    headerHeight,
    footerHeight,
  };
}

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

export class PaginationEngine {
  private stagingElement: HTMLElement;
  private containerElement: HTMLElement;
  private scale: number;
  private cssPrefix: string;
  private showPageNumbers: boolean;
  private pageGap: number;
  private hfRegistry: HeaderFooterRegistry;
  private footnoteRegistry: FootnoteRegistry;
  private pendingFootnoteContinuation: FootnoteContinuation | null = null;

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
    this.stagingElement =
      typeof staging === "string"
        ? (document.getElementById(staging) as HTMLElement)
        : staging;
    this.containerElement =
      typeof container === "string"
        ? (document.getElementById(container) as HTMLElement)
        : container;

    if (!this.stagingElement) {
      throw new Error("Staging element not found");
    }
    if (!this.containerElement) {
      throw new Error("Container element not found");
    }

    this.scale = options.scale ?? 1;
    this.cssPrefix = options.cssPrefix ?? "page-";
    this.showPageNumbers = options.showPageNumbers ?? true;
    this.pageGap = options.pageGap ?? 20;
    this.hfRegistry = new Map();
    this.footnoteRegistry = new Map();
  }

  /**
   * Runs the pagination process.
   *
   * @returns PaginationResult with page information
   */
  paginate(): PaginationResult {
    const pages: PageInfo[] = [];
    let pageNumber = 1;

    // Parse the header/footer registry if present
    this.hfRegistry = this.parseHeaderFooterRegistry();

    // Parse the footnote registry if present
    this.footnoteRegistry = this.parseFootnoteRegistry();

    // Find all section containers
    const sections = this.stagingElement.querySelectorAll<HTMLElement>(
      "[data-section-index]"
    );

    // If no sections found, treat the entire staging content as one section
    const sectionsToProcess =
      sections.length > 0 ? Array.from(sections) : [this.stagingElement];

    for (const section of sectionsToProcess) {
      const sectionIndex = parseInt(section.dataset.sectionIndex || "0", 10);
      const dims = parseDimensions(section);

      // Make staging visible for measurement
      this.stagingElement.style.visibility = "hidden";
      this.stagingElement.style.position = "absolute";
      this.stagingElement.style.left = "-9999px";
      this.stagingElement.style.display = "block";

      // Set width for accurate line wrapping
      section.style.width = `${dims.contentWidth}pt`;

      // Measure all blocks in this section
      const blocks = this.measureBlocks(section, dims);

      // Flow blocks into pages
      const sectionPages = this.flowToPages(blocks, dims, pageNumber, sectionIndex);
      pages.push(...sectionPages);
      pageNumber += sectionPages.length;
    }

    // Hide staging after measurement
    this.stagingElement.style.display = "none";

    return { totalPages: pages.length, pages };
  }

  /**
   * Measures all content blocks in a section.
   */
  private measureBlocks(section: HTMLElement, dims: PageDimensions): MeasuredBlock[] {
    const blocks: MeasuredBlock[] = [];

    // Get direct children (paragraphs, tables, divs, etc.)
    const children = Array.from(section.children) as HTMLElement[];

    for (const child of children) {
      // Skip section dividers that are just wrappers
      if (child.dataset.sectionIndex !== undefined) {
        // Recursively get blocks from nested sections
        const nestedBlocks = this.measureBlocks(child, dims);
        blocks.push(...nestedBlocks);
        continue;
      }

      // Measure height and margins separately for proper margin collapsing calculation
      // getBoundingClientRect() returns content+padding+border, not margins
      const rect = child.getBoundingClientRect();
      const style = window.getComputedStyle(child);
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
        heightPt,
        marginTopPt,
        marginBottomPt,
        keepWithNext: child.dataset.keepWithNext === "true",
        keepLines: child.dataset.keepLines === "true",
        pageBreakBefore: child.dataset.pageBreakBefore === "true",
        isPageBreak,
      });
    }

    return blocks;
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

    const entries = Array.from(registryEl.querySelectorAll<HTMLElement>("[data-footnote-id]"));

    for (const entry of entries) {
      const footnoteId = entry.dataset.footnoteId;
      if (footnoteId) {
        // Clone the footnote element for later use
        registry.set(footnoteId, entry.cloneNode(true) as HTMLElement);
      }
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
    continuation?: FootnoteContinuation | null
  ): number {
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if ((footnoteIds.length === 0 && !hasContinuation) || this.footnoteRegistry.size === 0) {
      return 0;
    }

    // Create a temporary measurement container
    const measureContainer = document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";

    // Add separator line (same as will be rendered)
    const hr = document.createElement("hr");
    measureContainer.appendChild(hr);

    // Add continuation content first (if any)
    if (hasContinuation) {
      const contWrapper = document.createElement("div");
      contWrapper.className = "footnote-continuation";
      for (const el of continuation!.remainingElements) {
        contWrapper.appendChild(el.cloneNode(true));
      }
      measureContainer.appendChild(contWrapper);
    }

    // Add footnotes
    for (const id of footnoteIds) {
      const footnote = this.footnoteRegistry.get(id);
      if (footnote) {
        measureContainer.appendChild(footnote.cloneNode(true));
      }
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
   * Measures the height of just the continuation content (in points).
   */
  private measureContinuationHeight(
    continuation: FootnoteContinuation,
    contentWidth: number
  ): number {
    if (!continuation || continuation.remainingElements.length === 0) {
      return 0;
    }

    const measureContainer = document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";

    // Add separator line
    const hr = document.createElement("hr");
    measureContainer.appendChild(hr);

    // Add continuation content
    for (const el of continuation.remainingElements) {
      measureContainer.appendChild(el.cloneNode(true));
    }

    this.stagingElement.appendChild(measureContainer);
    const rect = measureContainer.getBoundingClientRect();
    const heightPt = pxToPt(rect.height);
    this.stagingElement.removeChild(measureContainer);

    return heightPt;
  }

  /**
   * Splits a footnote element into parts that fit within the available height.
   * Returns the elements that fit and the elements that need to continue.
   */
  private splitFootnoteToFit(
    footnoteElement: HTMLElement,
    availableHeightPt: number,
    contentWidth: number
  ): { fits: HTMLElement[]; overflow: HTMLElement[] } {
    // Get child elements (paragraphs) of the footnote content
    const footnoteContent = footnoteElement.querySelector(".footnote-content");
    if (!footnoteContent) {
      // No content structure, can't split - return whole footnote
      return { fits: [footnoteElement.cloneNode(true) as HTMLElement], overflow: [] };
    }

    const children = Array.from(footnoteContent.children) as HTMLElement[];
    if (children.length <= 1) {
      // Single paragraph, can't split at paragraph level
      return { fits: [footnoteElement.cloneNode(true) as HTMLElement], overflow: [] };
    }

    const fits: HTMLElement[] = [];
    const overflow: HTMLElement[] = [];
    let currentHeight = 0;

    // Measure separator line height
    const hrMeasure = document.createElement("div");
    hrMeasure.style.position = "absolute";
    hrMeasure.style.visibility = "hidden";
    hrMeasure.style.width = `${contentWidth}pt`;
    hrMeasure.style.left = "-9999px";
    const hr = document.createElement("hr");
    hrMeasure.appendChild(hr);
    this.stagingElement.appendChild(hrMeasure);
    const hrHeight = pxToPt(hrMeasure.getBoundingClientRect().height);
    this.stagingElement.removeChild(hrMeasure);

    currentHeight = hrHeight;

    // Also account for footnote number
    const footnoteNumber = footnoteElement.querySelector(".footnote-number");

    for (let i = 0; i < children.length; i++) {
      const child = children[i];

      // Measure this element
      const measureContainer = document.createElement("div");
      measureContainer.style.position = "absolute";
      measureContainer.style.visibility = "hidden";
      measureContainer.style.width = `${contentWidth}pt`;
      measureContainer.style.left = "-9999px";
      measureContainer.appendChild(child.cloneNode(true));
      this.stagingElement.appendChild(measureContainer);
      const childHeight = pxToPt(measureContainer.getBoundingClientRect().height);
      this.stagingElement.removeChild(measureContainer);

      if (currentHeight + childHeight <= availableHeightPt) {
        fits.push(child.cloneNode(true) as HTMLElement);
        currentHeight += childHeight;
      } else {
        // This and remaining elements overflow
        for (let j = i; j < children.length; j++) {
          overflow.push(children[j].cloneNode(true) as HTMLElement);
        }
        break;
      }
    }

    return { fits, overflow };
  }

  /**
   * Measures a single footnote's height.
   */
  private measureSingleFootnoteHeight(footnoteId: string, contentWidth: number): number {
    const footnote = this.footnoteRegistry.get(footnoteId);
    if (!footnote) return 0;

    const measureContainer = document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    measureContainer.appendChild(footnote.cloneNode(true));

    this.stagingElement.appendChild(measureContainer);
    const rect = measureContainer.getBoundingClientRect();
    const heightPt = pxToPt(rect.height);
    this.stagingElement.removeChild(measureContainer);

    return heightPt;
  }

  /**
   * Adds footnotes to a page container, including continuation content.
   */
  private addPageFootnotes(
    pageBox: HTMLElement,
    footnoteIds: string[],
    dims: PageDimensions,
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

    // Create a set of partial footnote IDs for quick lookup
    const partialFootnoteIds = new Set(partialFootnotes?.map(p => p.footnoteId) || []);

    // Calculate max height for footnotes area (content height minus margin for body content)
    const maxFootnoteHeight = Math.min(
      footnoteHeight,
      dims.contentHeight * MAX_FOOTNOTE_AREA_RATIO
    );

    const footnotesDiv = document.createElement("div");
    footnotesDiv.className = `${this.cssPrefix}footnotes`;
    footnotesDiv.style.position = "absolute";
    footnotesDiv.style.bottom = `${dims.marginBottom}pt`; // Above footer area
    footnotesDiv.style.left = `${dims.marginLeft}pt`;
    footnotesDiv.style.width = `${dims.contentWidth}pt`;
    footnotesDiv.style.boxSizing = "border-box";
    // Constrain height and clip overflow to prevent footnotes covering body content
    footnotesDiv.style.maxHeight = `${maxFootnoteHeight}pt`;
    footnotesDiv.style.overflow = "hidden";

    // Add separator line
    const hr = document.createElement("hr");
    footnotesDiv.appendChild(hr);

    // Add continuation content first (if any)
    if (hasContinuation) {
      const contWrapper = document.createElement("div");
      contWrapper.className = "footnote-continuation";
      for (const el of continuation!.remainingElements) {
        contWrapper.appendChild(el.cloneNode(true));
      }
      footnotesDiv.appendChild(contWrapper);
    }

    // Clone footnotes in order of appearance
    for (const id of footnoteIds) {
      // Check if this is a partial footnote
      const partial = partialFootnotes?.find(p => p.footnoteId === id);
      if (partial) {
        // Render partial footnote (only the fitting elements)
        const footnote = this.footnoteRegistry.get(id);
        if (footnote) {
          const partialDiv = document.createElement("div");
          partialDiv.className = "footnote-item";
          partialDiv.dataset.footnoteId = id;

          // Add footnote number
          const numberSpan = footnote.querySelector(".footnote-number");
          if (numberSpan) {
            partialDiv.appendChild(numberSpan.cloneNode(true));
          }

          // Add only the fitting content
          const contentSpan = document.createElement("span");
          contentSpan.className = "footnote-content";
          for (const el of partial.fittingElements) {
            contentSpan.appendChild(el.cloneNode(true));
          }
          partialDiv.appendChild(contentSpan);

          footnotesDiv.appendChild(partialDiv);
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

  /**
   * Selects the appropriate header for a page based on section, page position, and page number.
   */
  private selectHeader(
    sectionIndex: number,
    pageInSection: number,
    globalPageNumber: number
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first header if available
    if (pageInSection === 1 && sectionHf.headerFirst) {
      return sectionHf.headerFirst;
    }

    // Even pages use even header if available
    if (globalPageNumber % 2 === 0 && sectionHf.headerEven) {
      return sectionHf.headerEven;
    }

    // Default (odd) pages
    return sectionHf.headerDefault;
  }

  /**
   * Selects the appropriate footer for a page based on section, page position, and page number.
   */
  private selectFooter(
    sectionIndex: number,
    pageInSection: number,
    globalPageNumber: number
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first footer if available
    if (pageInSection === 1 && sectionHf.footerFirst) {
      return sectionHf.footerFirst;
    }

    // Even pages use even footer if available
    if (globalPageNumber % 2 === 0 && sectionHf.footerEven) {
      return sectionHf.footerEven;
    }

    // Default (odd) pages
    return sectionHf.footerDefault;
  }

  /**
   * Computes effective header, footer, and content heights for a specific page position.
   * Uses pre-measured header/footer heights from the registry.
   * This method is deterministic - same inputs always produce same outputs.
   * This enables lazy loading compatibility since available height can be computed
   * for any page position without knowing the page's content.
   */
  private getEffectiveHeights(
    dims: PageDimensions,
    sectionIndex: number,
    pageInSection: number,
    globalPageNumber: number
  ): { headerHeight: number; footerHeight: number; contentHeight: number } {
    const sectionHf = this.hfRegistry.get(sectionIndex);

    // Determine effective header height
    let headerHeight = dims.marginTop;
    if (sectionHf) {
      let measuredHeaderHeight: number | undefined;

      // Select the appropriate header height based on page position
      if (pageInSection === 1 && sectionHf.headerFirstHeight != null) {
        measuredHeaderHeight = sectionHf.headerFirstHeight;
      } else if (globalPageNumber % 2 === 0 && sectionHf.headerEvenHeight != null) {
        measuredHeaderHeight = sectionHf.headerEvenHeight;
      } else if (sectionHf.headerDefaultHeight != null) {
        measuredHeaderHeight = sectionHf.headerDefaultHeight;
      }

      if (measuredHeaderHeight != null) {
        // Use the larger of margin height or measured content height
        headerHeight = Math.max(dims.marginTop, measuredHeaderHeight);
      }
    }

    // Determine effective footer height
    let footerHeight = dims.marginBottom;
    if (sectionHf) {
      let measuredFooterHeight: number | undefined;

      // Select the appropriate footer height based on page position
      if (pageInSection === 1 && sectionHf.footerFirstHeight != null) {
        measuredFooterHeight = sectionHf.footerFirstHeight;
      } else if (globalPageNumber % 2 === 0 && sectionHf.footerEvenHeight != null) {
        measuredFooterHeight = sectionHf.footerEvenHeight;
      } else if (sectionHf.footerDefaultHeight != null) {
        measuredFooterHeight = sectionHf.footerDefaultHeight;
      }

      if (measuredFooterHeight != null) {
        // Use the larger of margin height or measured content height
        footerHeight = Math.max(dims.marginBottom, measuredFooterHeight);
      }
    }

    // Calculate effective content height
    // contentHeight from dims is: pageHeight - marginTop - marginBottom
    // We need to adjust for any header/footer expansion beyond margins
    const headerExpansion = headerHeight - dims.marginTop;
    const footerExpansion = footerHeight - dims.marginBottom;
    const contentHeight = dims.contentHeight - headerExpansion - footerExpansion;

    return { headerHeight, footerHeight, contentHeight };
  }

  /**
   * Measures the content height of a header or footer element.
   * This is needed because headers/footers can contain more content than fits in the margin area.
   */
  private measureHeaderFooterHeight(
    source: HTMLElement,
    contentWidth: number
  ): number {
    // Create a temporary measurement container
    const measureContainer = document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    // Add padding to match the actual rendering
    measureContainer.style.paddingBottom = "4pt";

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
    sectionIndex: number
  ): PageInfo[] {
    const pages: PageInfo[] = [];
    let currentContent: HTMLElement[] = [];
    let pageNumber = startPageNumber;
    // Track page number within this section for first-page header/footer selection
    let pageInSection = 1;

    // Get effective content height for first page (accounts for header/footer sizes)
    let { contentHeight: effectiveContentHeight } = this.getEffectiveHeights(
      dims, sectionIndex, pageInSection, pageNumber
    );
    let remainingHeight = effectiveContentHeight;

    // Track the previous block's bottom margin for margin collapsing
    let prevMarginBottomPt = 0;
    // Track footnote IDs for the current page
    let currentFootnoteIds: string[] = [];
    // Track height consumed by footnotes on current page
    let currentFootnoteHeight = 0;
    // Track footnote continuation for current page (from previous page)
    let currentContinuation: FootnoteContinuation | null = this.pendingFootnoteContinuation;
    // Track any new continuation that will carry to next page
    let nextPageContinuation: FootnoteContinuation | null = null;
    // Track partial footnotes for current page (footnotes that were split)
    let currentPartialFootnotes: PartialFootnote[] = [];

    // Account for any continuation from previous section/page
    if (currentContinuation && currentContinuation.remainingElements.length > 0) {
      currentFootnoteHeight = this.measureContinuationHeight(currentContinuation, dims.contentWidth);
    }

    const finishPage = () => {
      if (currentContent.length === 0 && !currentContinuation) return;

      const page = this.createPage(
        dims,
        pageNumber,
        sectionIndex,
        currentContent,
        pageInSection,
        currentFootnoteIds,
        currentFootnoteHeight,
        currentContinuation,
        currentPartialFootnotes.length > 0 ? currentPartialFootnotes : undefined
      );
      pages.push(page);

      pageNumber++;
      pageInSection++;
      currentContent = [];

      // Get effective content height for new page position
      const newHeights = this.getEffectiveHeights(dims, sectionIndex, pageInSection, pageNumber);
      effectiveContentHeight = newHeights.contentHeight;
      remainingHeight = effectiveContentHeight;

      prevMarginBottomPt = 0; // Reset margin tracking for new page
      currentFootnoteIds = []; // Reset footnotes for new page
      currentPartialFootnotes = []; // Reset partial footnotes for new page

      // Carry over continuation to next page
      currentContinuation = nextPageContinuation;
      nextPageContinuation = null;

      // Account for continuation height on new page
      if (currentContinuation && currentContinuation.remainingElements.length > 0) {
        currentFootnoteHeight = this.measureContinuationHeight(currentContinuation, dims.contentWidth);
      } else {
        currentFootnoteHeight = 0;
      }
    };

    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];
      const nextBlock = blocks[i + 1];

      // Handle explicit page breaks
      if (block.isPageBreak) {
        finishPage();
        continue;
      }

      // Handle page break before
      if (block.pageBreakBefore && currentContent.length > 0) {
        finishPage();
      }

      // Extract footnote references from this block
      const blockFootnoteIds = this.extractFootnoteRefs(block.element);
      // Only count new footnotes (not already on this page)
      const newFootnoteIds = blockFootnoteIds.filter(id => !currentFootnoteIds.includes(id));

      // Calculate additional footnote height if this block is added
      let additionalFootnoteHeight = 0;
      if (newFootnoteIds.length > 0 && this.footnoteRegistry.size > 0) {
        // Measure the combined height of all footnotes that would be on this page
        // (including any continuation)
        const combinedFootnoteIds = [...currentFootnoteIds, ...newFootnoteIds];
        const totalFootnoteHeight = this.measureFootnotesHeight(
          combinedFootnoteIds,
          dims.contentWidth,
          currentContinuation
        );
        additionalFootnoteHeight = totalFootnoteHeight - currentFootnoteHeight;
      }

      // Calculate the effective height this block will consume
      // Account for margin collapsing: the gap between blocks is max(prevBottom, currTop), not sum
      const isFirstOnPage = currentContent.length === 0;
      let effectiveMarginTop = block.marginTopPt;
      if (!isFirstOnPage) {
        // Margin collapsing: use the larger of the two adjacent margins
        effectiveMarginTop = Math.max(block.marginTopPt, prevMarginBottomPt) - prevMarginBottomPt;
      }
      // Total height = top margin gap + content + bottom margin + footnote space
      const blockSpace = effectiveMarginTop + block.heightPt + block.marginBottomPt + additionalFootnoteHeight;

      // Calculate needed height (including keepWithNext)
      let neededHeight = blockSpace;
      if (block.keepWithNext && nextBlock && !nextBlock.isPageBreak) {
        // For keepWithNext, include the next block with collapsed margins
        const collapsedMargin = Math.max(block.marginBottomPt, nextBlock.marginTopPt);
        neededHeight = effectiveMarginTop + block.heightPt + collapsedMargin +
                       nextBlock.heightPt + nextBlock.marginBottomPt + additionalFootnoteHeight;
      }

      // Effective remaining height (content area minus footnotes already on page)
      const effectiveRemainingHeight = remainingHeight - currentFootnoteHeight;

      // Calculate maximum footnote area for this page (can expand into body content space)
      const bodyContentUsed = effectiveContentHeight - remainingHeight;
      const maxFootnoteArea = effectiveContentHeight * MAX_FOOTNOTE_AREA_RATIO;
      const maxFootnoteExpansion = Math.max(0, maxFootnoteArea - currentFootnoteHeight);

      // Check if block fits on current page (including its footnotes)
      if (blockSpace <= effectiveRemainingHeight) {
        // Block fits with current footnote allocation
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        remainingHeight -= (effectiveMarginTop + block.heightPt + block.marginBottomPt);
        prevMarginBottomPt = block.marginBottomPt;
        // Add new footnotes to current page
        if (newFootnoteIds.length > 0) {
          currentFootnoteIds.push(...newFootnoteIds);
          currentFootnoteHeight += additionalFootnoteHeight;
        }
      } else if (block.heightPt + block.marginTopPt + block.marginBottomPt <= effectiveContentHeight) {
        // Block doesn't fit with current allocation - try expanding footnote area
        const blockSpaceWithoutFootnotes = effectiveMarginTop + block.heightPt + block.marginBottomPt;

        // Check if block fits if we expand footnote area
        // We can expand footnotes up to maxFootnoteArea, leaving room for body content
        const minBodySpaceNeeded = bodyContentUsed + blockSpaceWithoutFootnotes + MIN_BODY_CONTENT_HEIGHT;
        const canExpandFootnotes = minBodySpaceNeeded <= effectiveContentHeight;

        if (newFootnoteIds.length > 0 && blockSpaceWithoutFootnotes <= effectiveRemainingHeight) {
          // Block itself fits, but footnotes don't - expand footnote area
          currentContent.push(block.element.cloneNode(true) as HTMLElement);
          remainingHeight -= (effectiveMarginTop + block.heightPt + block.marginBottomPt);
          prevMarginBottomPt = block.marginBottomPt;

          // Calculate EXPANDED space available for footnotes
          // Footnotes can take up to maxFootnoteArea or all remaining space, whichever is less
          const availableForFootnotes = Math.min(
            maxFootnoteArea,
            effectiveContentHeight - bodyContentUsed - blockSpaceWithoutFootnotes
          );

          // Try to fit as much of each new footnote as possible in expanded area
          for (const footnoteId of newFootnoteIds) {
            const footnote = this.footnoteRegistry.get(footnoteId);
            if (!footnote) continue;

            const footnoteHeight = this.measureSingleFootnoteHeight(footnoteId, dims.contentWidth);
            const spaceLeftForFootnotes = availableForFootnotes - currentFootnoteHeight;

            if (footnoteHeight <= spaceLeftForFootnotes) {
              // Whole footnote fits in expanded area
              currentFootnoteIds.push(footnoteId);
              currentFootnoteHeight += footnoteHeight;
            } else {
              // Footnote needs to be split - use all available expanded space
              if (spaceLeftForFootnotes > 20) { // Minimum space to start a footnote
                const { fits, overflow } = this.splitFootnoteToFit(
                  footnote,
                  spaceLeftForFootnotes,
                  dims.contentWidth
                );

                if (fits.length > 0) {
                  // Add partial footnote to current page
                  currentFootnoteIds.push(footnoteId);
                  currentPartialFootnotes.push({
                    footnoteId,
                    fittingElements: fits
                  });
                  nextPageContinuation = {
                    footnoteId,
                    remainingElements: overflow
                  };
                  currentFootnoteHeight = availableForFootnotes;
                } else {
                  // Nothing fits, entire footnote continues to next page
                  nextPageContinuation = {
                    footnoteId,
                    remainingElements: Array.from(footnote.querySelectorAll(".footnote-content > *"))
                      .map(el => el.cloneNode(true) as HTMLElement)
                  };
                  if (nextPageContinuation.remainingElements.length === 0) {
                    nextPageContinuation.remainingElements = [footnote.cloneNode(true) as HTMLElement];
                  }
                }
              } else {
                // Not enough space to start footnote - continue whole thing
                nextPageContinuation = {
                  footnoteId,
                  remainingElements: Array.from(footnote.querySelectorAll(".footnote-content > *"))
                    .map(el => el.cloneNode(true) as HTMLElement)
                };
                if (nextPageContinuation.remainingElements.length === 0) {
                  nextPageContinuation.remainingElements = [footnote.cloneNode(true) as HTMLElement];
                }
              }
            }
          }
        } else if (canExpandFootnotes && newFootnoteIds.length > 0) {
          // Block doesn't fit with current layout, but might fit if we expand footnote area first
          // This handles the case where we need to give footnotes more space BEFORE adding the block

          // First, try to fit more of current footnotes by expanding the area
          // Then check if the block fits in reduced body space
          const expandedFootnoteSpace = Math.min(maxFootnoteArea, additionalFootnoteHeight + currentFootnoteHeight);
          const bodySpaceAfterExpansion = effectiveContentHeight - expandedFootnoteSpace;

          if (blockSpaceWithoutFootnotes <= bodySpaceAfterExpansion - bodyContentUsed) {
            // Block fits after expanding footnote area
            currentContent.push(block.element.cloneNode(true) as HTMLElement);
            remainingHeight = bodySpaceAfterExpansion - bodyContentUsed - blockSpaceWithoutFootnotes;
            prevMarginBottomPt = block.marginBottomPt;
            currentFootnoteIds.push(...newFootnoteIds);
            currentFootnoteHeight = expandedFootnoteSpace;
          } else {
            // Still doesn't fit - start new page
            finishPage();
            const newPageFootnoteHeight = blockFootnoteIds.length > 0
              ? this.measureFootnotesHeight(blockFootnoteIds, dims.contentWidth, currentContinuation)
              : (currentContinuation ? this.measureContinuationHeight(currentContinuation, dims.contentWidth) : 0);
            const newPageSpace = block.marginTopPt + block.heightPt + block.marginBottomPt;
            currentContent.push(block.element.cloneNode(true) as HTMLElement);
            remainingHeight = effectiveContentHeight - newPageSpace;
            prevMarginBottomPt = block.marginBottomPt;
            currentFootnoteIds = [...blockFootnoteIds];
            currentFootnoteHeight = newPageFootnoteHeight;
          }
        } else {
          // Block itself doesn't fit - start new page
          finishPage();
          // On new page, recalculate footnote height for just this block's footnotes
          // (plus any continuation from previous page)
          const newPageFootnoteHeight = blockFootnoteIds.length > 0
            ? this.measureFootnotesHeight(blockFootnoteIds, dims.contentWidth, currentContinuation)
            : (currentContinuation ? this.measureContinuationHeight(currentContinuation, dims.contentWidth) : 0);
          // Include full top margin
          const newPageSpace = block.marginTopPt + block.heightPt + block.marginBottomPt;
          currentContent.push(block.element.cloneNode(true) as HTMLElement);
          remainingHeight = effectiveContentHeight - newPageSpace;
          prevMarginBottomPt = block.marginBottomPt;
          currentFootnoteIds = [...blockFootnoteIds];
          currentFootnoteHeight = newPageFootnoteHeight;
        }
      } else {
        // Block is taller than a page - add it and let it overflow
        // (In a more sophisticated implementation, we would split the block)
        if (currentContent.length > 0) {
          finishPage();
        }
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        currentFootnoteIds = [...blockFootnoteIds];
        finishPage();
      }
    }

    // Finish last page
    finishPage();

    // Store any remaining continuation for next section
    this.pendingFootnoteContinuation = nextPageContinuation;

    return pages;
  }

  /**
   * Creates a page container element.
   */
  private createPage(
    dims: PageDimensions,
    pageNumber: number,
    sectionIndex: number,
    content: HTMLElement[],
    pageInSection: number,
    footnoteIds: string[] = [],
    footnoteHeight: number = 0,
    continuation?: FootnoteContinuation | null,
    partialFootnotes?: PartialFootnote[]
  ): PageInfo {
    // Create page box at full size, then scale the entire box
    // This ensures proper clipping and consistent scaling of all elements
    const pageBox = document.createElement("div");
    pageBox.className = `${this.cssPrefix}box`;
    pageBox.style.width = `${dims.pageWidth}pt`;
    pageBox.style.height = `${dims.pageHeight}pt`;
    pageBox.style.overflow = "hidden";
    pageBox.style.position = "relative";
    // Use CSS zoom for better text rendering when supported, fall back to transform
    // Zoom affects layout (no negative margin hack needed) and renders text more crisply
    // Note: zoom is non-standard but supported in all major browsers
    if (this.scale !== 1) {
      // Try zoom first (better text quality), with transform as fallback
      pageBox.style.zoom = String(this.scale);
      // For browsers that don't support zoom, also set transform
      // The zoom takes precedence in supporting browsers
      pageBox.style.transform = `scale(${this.scale})`;
      pageBox.style.transformOrigin = "top left";
      // Compensate for transform not affecting layout (only needed if zoom not supported)
      // Convert pt to px for consistent unit math
      const heightReductionPt = dims.pageHeight * (1 - this.scale);
      const widthReductionPt = dims.pageWidth * (1 - this.scale);
      const heightReductionPx = ptToPx(heightReductionPt);
      const widthReductionPx = ptToPx(widthReductionPt);
      pageBox.style.marginRight = `-${widthReductionPx}px`;
      pageBox.style.marginBottom = `${this.pageGap - heightReductionPx}px`;
    }
    // Hint browser for GPU compositing and layout isolation
    pageBox.style.willChange = "transform";
    pageBox.style.contain = "layout paint";
    pageBox.dataset.pageNumber = String(pageNumber);
    pageBox.dataset.sectionIndex = String(sectionIndex);

    // Get pre-computed effective heights for this page position (no re-measurement needed)
    const effectiveHeights = this.getEffectiveHeights(dims, sectionIndex, pageInSection, pageNumber);

    // Add header if available for this section/page
    const headerSource = this.selectHeader(sectionIndex, pageInSection, pageNumber);

    if (headerSource) {
      const headerDiv = document.createElement("div");
      headerDiv.className = `${this.cssPrefix}header`;
      headerDiv.style.position = "absolute";
      headerDiv.style.top = "0"; // Start at page top
      headerDiv.style.left = `${dims.marginLeft}pt`;
      headerDiv.style.width = `${dims.contentWidth}pt`;
      headerDiv.style.height = `${effectiveHeights.headerHeight}pt`; // Use pre-computed effective height
      headerDiv.style.overflow = "hidden";
      headerDiv.style.boxSizing = "border-box";
      headerDiv.style.display = "flex";
      headerDiv.style.flexDirection = "column";
      headerDiv.style.justifyContent = "flex-end"; // Align content to bottom of header area
      headerDiv.style.paddingBottom = "4pt"; // Small gap before content area
      // Clone the header content (skip the wrapper div's data attributes)
      for (const child of Array.from(headerSource.childNodes)) {
        headerDiv.appendChild(child.cloneNode(true));
      }
      pageBox.appendChild(headerDiv);
    }

    // Create content area using pre-computed effective heights
    const contentAreaTop = effectiveHeights.headerHeight;
    const contentAreaHeight = effectiveHeights.contentHeight;

    const contentArea = document.createElement("div");
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
    if (footnoteIds.length > 0 || hasContinuation) {
      this.addPageFootnotes(pageBox, footnoteIds, dims, footnoteHeight, continuation, partialFootnotes);
    }

    // Add footer if available for this section/page
    const footerSource = this.selectFooter(sectionIndex, pageInSection, pageNumber);
    if (footerSource) {
      const footerDiv = document.createElement("div");
      footerDiv.className = `${this.cssPrefix}footer`;
      footerDiv.style.position = "absolute";
      footerDiv.style.bottom = "0"; // Start at page bottom
      footerDiv.style.left = `${dims.marginLeft}pt`;
      footerDiv.style.width = `${dims.contentWidth}pt`;
      footerDiv.style.height = `${effectiveHeights.footerHeight}pt`; // Use pre-computed effective height
      footerDiv.style.overflow = "hidden";
      footerDiv.style.boxSizing = "border-box";
      footerDiv.style.display = "flex";
      footerDiv.style.flexDirection = "column";
      footerDiv.style.justifyContent = "flex-start"; // Align content to top of footer area
      footerDiv.style.paddingTop = "4pt"; // Small gap after content area
      // Clone the footer content (skip the wrapper div's data attributes)
      for (const child of Array.from(footerSource.childNodes)) {
        footerDiv.appendChild(child.cloneNode(true));
      }
      pageBox.appendChild(footerDiv);
    }

    // Add page number (will be hidden by CSS if document has footer)
    if (this.showPageNumbers) {
      const pageNum = document.createElement("div");
      pageNum.className = `${this.cssPrefix}number`;
      pageNum.textContent = String(pageNumber);
      pageBox.appendChild(pageNum);
    }

    // Add to container
    this.containerElement.appendChild(pageBox);

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
  const containerEl =
    typeof container === "string"
      ? (document.getElementById(container) as HTMLElement)
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

  const engine = new PaginationEngine(staging, pageContainer, options);
  return engine.paginate();
}

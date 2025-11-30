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
}

/**
 * A measured content block with metadata for pagination decisions.
 */
export interface MeasuredBlock {
  /** The DOM element */
  element: HTMLElement;
  /** Measured height in points */
  heightPt: number;
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

  return {
    pageWidth,
    pageHeight,
    contentWidth,
    contentHeight,
    marginTop,
    marginRight,
    marginBottom,
    marginLeft,
  };
}

/**
 * Pagination engine that converts HTML with pagination metadata
 * into a paginated view with fixed-size page containers.
 */
export class PaginationEngine {
  private stagingElement: HTMLElement;
  private containerElement: HTMLElement;
  private scale: number;
  private cssPrefix: string;
  private showPageNumbers: boolean;
  private pageGap: number;

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
  }

  /**
   * Runs the pagination process.
   *
   * @returns PaginationResult with page information
   */
  paginate(): PaginationResult {
    const pages: PageInfo[] = [];
    let pageNumber = 1;

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

      const rect = child.getBoundingClientRect();
      const heightPt = pxToPt(rect.height);

      const isPageBreak =
        child.dataset.pageBreak === "true" ||
        child.classList.contains(`${this.cssPrefix}break`);

      blocks.push({
        element: child,
        heightPt,
        keepWithNext: child.dataset.keepWithNext === "true",
        keepLines: child.dataset.keepLines === "true",
        pageBreakBefore: child.dataset.pageBreakBefore === "true",
        isPageBreak,
      });
    }

    return blocks;
  }

  /**
   * Flows measured blocks into page containers.
   */
  private flowToPages(
    blocks: MeasuredBlock[],
    dims: PageDimensions,
    startPageNumber: number,
    sectionIndex: number
  ): PageInfo[] {
    const pages: PageInfo[] = [];
    let currentContent: HTMLElement[] = [];
    let remainingHeight = dims.contentHeight;
    let pageNumber = startPageNumber;

    const finishPage = () => {
      if (currentContent.length === 0) return;

      const page = this.createPage(dims, pageNumber, sectionIndex, currentContent);
      pages.push(page);

      pageNumber++;
      currentContent = [];
      remainingHeight = dims.contentHeight;
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

      // Calculate needed height (including keepWithNext)
      let neededHeight = block.heightPt;
      if (block.keepWithNext && nextBlock && !nextBlock.isPageBreak) {
        neededHeight += nextBlock.heightPt;
      }

      // Check if block fits on current page
      if (block.heightPt <= remainingHeight) {
        // Block fits
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        remainingHeight -= block.heightPt;
      } else if (block.heightPt <= dims.contentHeight) {
        // Block doesn't fit but will fit on a new page
        finishPage();
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        remainingHeight = dims.contentHeight - block.heightPt;
      } else {
        // Block is taller than a page - add it and let it overflow
        // (In a more sophisticated implementation, we would split the block)
        if (currentContent.length > 0) {
          finishPage();
        }
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        finishPage();
      }
    }

    // Finish last page
    finishPage();

    return pages;
  }

  /**
   * Creates a page container element.
   */
  private createPage(
    dims: PageDimensions,
    pageNumber: number,
    sectionIndex: number,
    content: HTMLElement[]
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

    // Create content area at full page margins and dimensions
    const contentArea = document.createElement("div");
    contentArea.className = `${this.cssPrefix}content`;
    contentArea.style.position = "absolute";
    contentArea.style.top = `${dims.marginTop}pt`;
    contentArea.style.left = `${dims.marginLeft}pt`;
    contentArea.style.width = `${dims.contentWidth}pt`;
    contentArea.style.height = `${dims.contentHeight}pt`;
    contentArea.style.overflow = "hidden";

    // Add content
    for (const el of content) {
      contentArea.appendChild(el);
    }

    pageBox.appendChild(contentArea);

    // Add page number
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

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
  /**
   * Whether ordinary paragraphs may be fragmented across page boundaries.
   * Defaults to false for direct PaginationEngine callers; read-only viewer
   * entry points opt in explicitly.
   */
  fragmentParagraphs?: boolean;
}

// Default letter size in points (612 x 792 = 8.5" x 11")
// Maximum percentage of content height that footnotes can occupy
// This allows footnotes to expand upward into body content space when needed
const MAX_FOOTNOTE_AREA_RATIO = 0.6; // 60% of content height

// Minimum body content height per page (to avoid pages with only footnotes)
const MIN_BODY_CONTENT_HEIGHT = 72; // 1 inch minimum body content

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
  private scale: number;
  private cssPrefix: string;
  private showPageNumbers: boolean;
  private pageGap: number;
  private fragmentParagraphs: boolean;
  private hfRegistry: HeaderFooterRegistry;
  private footnoteRegistry: FootnoteRegistry;
  private pendingFootnoteContinuation: FootnoteContinuation | null = null;
  /** Per-section `w:pgNumType` (start / format), read off the section wrappers. */
  private pageNumbering: Map<number, SectionPageNumbering> = new Map();

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
    this.fragmentParagraphs = options.fragmentParagraphs ?? false;
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

    this.pageNumbering = this.parsePageNumbering(sections);

    // If no sections found, treat the entire staging content as one section
    const sectionsToProcess =
      sections.length > 0 ? Array.from(sections) : [this.stagingElement];

    for (const section of sectionsToProcess) {
      const sectionIndex = parseInt(section.dataset.sectionIndex || "0", 10);
      const dims = parseSectionDimensions(section);

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

    // Every page box exists now, so NUMPAGES has an answer and each PAGE marker knows its page.
    this.substitutePageNumberFields(pages.length);

    return { totalPages: pages.length, pages };
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
      const pageInSection = parseInt(box.dataset.pageInSection || "1", 10);
      const numbering = this.pageNumbering.get(sectionIndex) ?? {};

      // A section that restarts numbering counts from its own start; one that does not continues
      // the document-wide running number.
      const displayed =
        numbering.start !== undefined ? numbering.start + pageInSection - 1 : pageNumber;

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
   * Measures one element in the same hidden staging context used for the source blocks.
   * This is intentionally DOM-based: table row heights cannot be inferred from individual
   * rows because wrapping and collapsed borders change the height of a fragment.
   */
  private measureElement(element: HTMLElement, dims: PageDimensions): MeasuredBlock {
    const measurementHost = document.createElement("div");
    measurementHost.style.position = "absolute";
    measurementHost.style.visibility = "hidden";
    measurementHost.style.left = "-9999px";
    measurementHost.style.width = `${dims.contentWidth}pt`;

    const measuredElement = element.cloneNode(true) as HTMLElement;
    measurementHost.appendChild(measuredElement);
    this.stagingElement.appendChild(measurementHost);

    const rect = measuredElement.getBoundingClientRect();
    const style = window.getComputedStyle(measuredElement);
    const measured: MeasuredBlock = {
      element,
      heightPt: pxToPt(rect.height),
      marginTopPt: pxToPt(parseFloat(style.marginTop) || 0),
      marginBottomPt: pxToPt(parseFloat(style.marginBottom) || 0),
      keepWithNext: element.dataset.keepWithNext === "true",
      keepLines: element.dataset.keepLines === "true",
      pageBreakBefore: element.dataset.pageBreakBefore === "true",
      isPageBreak:
        element.dataset.pageBreak === "true" ||
        element.classList.contains(`${this.cssPrefix}break`),
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
    isFirstOnPage: boolean
  ): number {
    const firstBlock = chain[0];
    if (!firstBlock) return 0;

    const firstMarginTop = isFirstOnPage
      ? firstBlock.marginTopPt
      : Math.max(firstBlock.marginTopPt, previousMarginBottomPt) - previousMarginBottomPt;
    let bodyHeight = firstMarginTop + firstBlock.heightPt;

    for (let index = 1; index < chain.length; index++) {
      const previousBlock = chain[index - 1];
      const block = chain[index];
      bodyHeight += Math.max(previousBlock.marginBottomPt, block.marginTopPt) + block.heightPt;
    }

    return bodyHeight;
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
      this.getPageBands(dims, sectionIndex, 1, 1).bodyHeight,
      this.getPageBands(dims, sectionIndex, 2, 1).bodyHeight,
      this.getPageBands(dims, sectionIndex, 2, 2).bodyHeight
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

    const table = wrapper.firstElementChild;
    if (!(table instanceof HTMLTableElement)) {
      return null;
    }

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
      let end = start;
      while (end < rows.length) {
        const candidate = this.createSimpleTableFragment(
          wrapper,
          table,
          body,
          rows.slice(start, end + 1),
          start === 0
        );
        const measured = this.measureElement(candidate, dims);
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

      const measured = this.measureElement(fragment, dims);
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
   * A DOM endpoint that can finish a paragraph fragment. Endpoints are chosen
   * after whitespace or at a run boundary so the paginator never deliberately
   * cuts through an ordinary word merely to fill a little more of a page.
   */
  private paragraphFragmentEndpoints(paragraph: HTMLElement): Array<{ node: Text; offset: number }> {
    const endpoints: Array<{ node: Text; offset: number }> = [];
    const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
    let textNode: Text | null;

    while ((textNode = walker.nextNode() as Text | null)) {
      const text = textNode.data;
      if (text.length === 0) continue;

      // A whitespace boundary preserves normal word wrapping. Always retain the
      // end of a run too: adjacent runs can change formatting without containing
      // a whitespace character between them.
      const whitespace = /\s+/g;
      let match: RegExpExecArray | null;
      while ((match = whitespace.exec(text)) !== null) {
        endpoints.push({ node: textNode, offset: match.index + match[0].length });
      }
      if (endpoints.length === 0 || endpoints[endpoints.length - 1].node !== textNode ||
          endpoints[endpoints.length - 1].offset !== text.length) {
        endpoints.push({ node: textNode, offset: text.length });
      }
    }

    return endpoints;
  }

  /**
   * Whether a range contains visible text after ignoring bidi/zero-width marks.
   * Paragraph fragmentation deliberately excludes non-textual descendants, so
   * this is enough to reject empty head or tail fragments.
   */
  private hasVisibleFragmentText(fragment: DocumentFragment): boolean {
    return (fragment.textContent || "")
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
    if (!this.fragmentParagraphs || block.element.tagName !== "P") {
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

    // The outer paragraph may carry its source identity. Descendant identities
    // are not safe to duplicate in a continuation fragment, so reject them.
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
      "[data-list-marker]",
      "[data-anchor]",
      "[id]",
      "[contenteditable]",
    ].join(", ");
    if (paragraph.querySelector(unsupportedDescendants)) {
      return false;
    }

    const paragraphStyle = window.getComputedStyle(paragraph);
    if (
      paragraphStyle.display !== "block" ||
      paragraphStyle.position !== "static" ||
      paragraphStyle.float !== "none" ||
      paragraphStyle.whiteSpace !== "normal" ||
      paragraphStyle.breakBefore !== "auto" ||
      paragraphStyle.breakAfter !== "auto" ||
      paragraphStyle.breakInside === "avoid" ||
      paragraphStyle.pageBreakBefore !== "auto" ||
      paragraphStyle.pageBreakAfter !== "auto" ||
      paragraphStyle.pageBreakInside === "avoid"
    ) {
      return false;
    }

    // A range clone preserves nested inline formatting exactly. Anything that
    // establishes its own box/layout context is intentionally deferred until a
    // future fragmenter can model it accurately.
    for (const descendant of Array.from(paragraph.querySelectorAll<HTMLElement>("*"))) {
      const style = window.getComputedStyle(descendant);
      if (
        style.display !== "inline" ||
        style.position !== "static" ||
        style.float !== "none" ||
        style.whiteSpace !== "normal"
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

    const paragraph = block.element;
    const endpoints = this.paragraphFragmentEndpoints(paragraph);
    // The final endpoint is the full paragraph and cannot leave a tail. A
    // single text run with no earlier whitespace remains an overflow fallback.
    if (endpoints.length < 2) {
      return null;
    }

    let low = 0;
    let high = endpoints.length - 2;
    let best: { endpoint: { node: Text; offset: number }; element: HTMLElement; measured: MeasuredBlock } | null = null;

    // Fragment height is monotonic for the deliberately narrow eligible subset,
    // so binary search avoids measuring every word in long body paragraphs.
    while (low <= high) {
      const middle = Math.floor((low + high) / 2);
      const endpoint = endpoints[middle];
      const headRange = document.createRange();
      headRange.setStart(paragraph, 0);
      headRange.setEnd(endpoint.node, endpoint.offset);
      const headContents = headRange.cloneContents();

      if (!this.hasVisibleFragmentText(headContents)) {
        low = middle + 1;
        continue;
      }

      const head = this.createParagraphFragment(paragraph, headRange, true, false);
      const measured = this.measureElement(head, dims);
      if (effectiveMarginTopPt + measured.heightPt <= availableHeightPt) {
        best = { endpoint, element: head, measured };
        low = middle + 1;
      } else {
        high = middle - 1;
      }
    }

    if (!best) {
      return null;
    }

    const tailRange = document.createRange();
    tailRange.setStart(best.endpoint.node, best.endpoint.offset);
    tailRange.setEnd(paragraph, paragraph.childNodes.length);
    const tailContents = tailRange.cloneContents();
    if (!this.hasVisibleFragmentText(tailContents)) {
      return null;
    }

    const tail = this.createParagraphFragment(paragraph, tailRange, false, true);
    const tailMeasured = this.measureElement(tail, dims);

    return [
      {
        ...best.measured,
        element: best.element,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
      },
      {
        ...tailMeasured,
        element: tail,
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

    // Measure in the SAME styling context the notes render in: `.page-footnotes` carries
    // font-size 0.85em and line-height 1.4, so measuring without the class sizes the note
    // block against body type and the reserve can never match what is drawn.
    // Create a temporary measurement container
    const measureContainer = document.createElement("div");
    measureContainer.style.position = "absolute";
    measureContainer.style.visibility = "hidden";
    measureContainer.style.width = `${contentWidth}pt`;
    measureContainer.style.left = "-9999px";
    measureContainer.className = this.cssPrefix + "footnotes";

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
    measureContainer.className = this.cssPrefix + "footnotes";

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
    // Get child elements (paragraphs) of the footnote content.
    //
    // `fits` is spliced into a freshly built `.footnote-item` > `.footnote-content` wrapper by
    // addPageFootnotes, so it must contain the note's CONTENT, never the note element itself.
    // Returning the whole `.footnote-item` here nested a complete item (number span and all)
    // inside another item's content span; the inner block-level div then broke the line, so the
    // note's number rendered alone above its text — the same visible symptom as the escaped-CSS
    // bug, from an unrelated cause, on the notes that happened to take a can't-split path.
    const footnoteContent = footnoteElement.querySelector(".footnote-content");
    if (!footnoteContent) {
      // No content structure to split — hand back the element's own children.
      return {
        fits: Array.from(footnoteElement.children).map((el) => el.cloneNode(true) as HTMLElement),
        overflow: [],
      };
    }

    const children = Array.from(footnoteContent.children) as HTMLElement[];
    if (children.length <= 1) {
      // Single paragraph: can't split at paragraph level, but the whole content still fits.
      return {
        fits: children.map((el) => el.cloneNode(true) as HTMLElement),
        overflow: [],
      };
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
    hrMeasure.className = this.cssPrefix + "footnotes";
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
      measureContainer.className = this.cssPrefix + "footnotes";
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
    measureContainer.className = this.cssPrefix + "footnotes";
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

    // Create a set of partial footnote IDs for quick lookup
    const partialFootnoteIds = new Set(partialFootnotes?.map(p => p.footnoteId) || []);

    // Calculate max height for footnotes area (content height minus margin for body content)
    const maxFootnoteHeight = Math.min(
      footnoteHeight,
      bands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO
    );

    const footnotesDiv = document.createElement("div");
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
    globalPageNumber: number
  ): PageBands {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    return resolvePageBands(
      dims,
      this.selectStoryHeight(
        sectionHf?.headerFirstHeight,
        sectionHf?.headerEvenHeight,
        sectionHf?.headerDefaultHeight,
        pageInSection,
        globalPageNumber
      ),
      this.selectStoryHeight(
        sectionHf?.footerFirstHeight,
        sectionHf?.footerEvenHeight,
        sectionHf?.footerDefaultHeight,
        pageInSection,
        globalPageNumber
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
    globalPageNumber: number
  ): number {
    if (pageInSection === 1 && first != null) return first;
    if (globalPageNumber % 2 === 0 && even != null) return even;
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
    const measureContainer = document.createElement("div");
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
    sectionIndex: number
  ): PageInfo[] {
    const pages: PageInfo[] = [];
    let currentContent: HTMLElement[] = [];
    let pageNumber = startPageNumber;
    // Track page number within this section for first-page header/footer selection
    let pageInSection = 1;

    // Get effective content height for first page (accounts for header/footer sizes)
    let { bodyHeight: effectiveContentHeight } = this.getPageBands(
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

    // Account for any continuation from previous section/page
    if (currentContinuation && currentContinuation.remainingElements.length > 0) {
      currentFootnoteHeight = this.measureContinuationHeight(currentContinuation, dims.contentWidth);
    }

    const finishPage = () => {
      const hasCurrentContinuation =
        (currentContinuation?.remainingElements.length ?? 0) > 0;
      if (currentContent.length === 0 && !hasCurrentContinuation) return;

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
      const newBands = this.getPageBands(dims, sectionIndex, pageInSection, pageNumber);
      effectiveContentHeight = newBands.bodyHeight;
      remainingHeight = effectiveContentHeight;

      prevMarginBottomPt = 0; // Reset margin tracking for new page
      currentFootnoteIds = []; // Reset footnotes for new page
      currentPartialFootnotes = []; // Reset partial footnotes for new page

      // Carry over continuation to next page
      currentContinuation = nextPageContinuation;
      nextPageContinuation = null;

      // Notes that never got started land at the top of the new page's note area. They are
      // ordinary footnotes from here on, so the normal fitting path handles them — and because
      // this page is fresh, the space they were denied now exists.
      if (deferredFootnoteIds.length > 0) {
        currentFootnoteIds = [...deferredFootnoteIds];
        deferredFootnoteIds = [];
      }

      // Account for continuation height on new page
      if (currentContinuation && currentContinuation.remainingElements.length > 0) {
        currentFootnoteHeight = this.measureContinuationHeight(currentContinuation, dims.contentWidth);
      } else {
        currentFootnoteHeight = 0;
      }
      if (currentFootnoteIds.length > 0) {
        currentFootnoteHeight += this.measureFootnotesHeight(
          currentFootnoteIds, dims.contentWidth, null);
      }
    };

    for (let i = 0; i < blocks.length; i++) {
      const block = blocks[i];

      // Handle explicit page breaks
      if (block.isPageBreak) {
        finishPage();
        continue;
      }

      // Handle page break before
      if (block.pageBreakBefore && currentContent.length > 0) {
        finishPage();
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
              dims.contentWidth,
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
              currentContent.length === 0
            ) +
            additionalChainFootnoteHeight;
          const currentAvailableHeight = remainingHeight - currentFootnoteHeight;

          if (currentChainHeight > currentAvailableHeight) {
            const nextPageBands = this.getPageBands(
              dims,
              sectionIndex,
              pageInSection + 1,
              pageNumber + 1
            );
            const freshChainBodyHeight = this.measureKeepWithNextChainBodyHeight(
              keepChain,
              0,
              true
            );
            // finishPage transfers this continuation to the new page's
            // currentContinuation state, so include it in the destination
            // page's footnote reservation before deciding to move the chain.
            const freshChainFootnoteHeight = this.measureFootnotesHeight(
              newChainFootnoteIds,
              dims.contentWidth,
              nextPageContinuation
            );

            if (
              freshChainBodyHeight + freshChainFootnoteHeight <=
              nextPageBands.bodyHeight
            ) {
              finishPage();
            }
          }
        }
      }

      // Extract footnote references from this block
      const allBlockFootnoteIds = this.extractFootnoteRefs(block.element);
      // Only count new footnotes (not already on this page)
      const newFootnoteIds = this.collectNewFootnoteIds([block], currentFootnoteIds);

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
      const maxFootnoteExpansion = Math.max(0, maxFootnoteArea - currentFootnoteHeight);

      // A paragraph that cannot fit as a whole may still have a simple text-only
      // prefix that fits this page. Fragment before the ordinary next-page or
      // oversized fallback so the cloned head participates in the same margin
      // and footnote accounting as every other block.
      if (blockSpace > effectiveRemainingHeight) {
        const paragraphFragments = this.tryFragmentParagraph(
          block,
          dims,
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
      } else if (block.heightPt + block.marginTopPt <= effectiveContentHeight) {
        // Block doesn't fit with current allocation - try expanding footnote area
        const blockSpaceWithoutFootnotes = effectiveMarginTop + block.heightPt;

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
                  if (overflow.length > 0) {
                    nextPageContinuation = {
                      footnoteId,
                      remainingElements: overflow
                    };
                  }
                  currentFootnoteHeight = availableForFootnotes;
                } else {
                  // Nothing of this note fits: defer the WHOLE note rather than assigning the
                  // single continuation slot, which a later note on this page would overwrite.
                  deferredFootnoteIds.push(footnoteId);
                }
              } else {
                // Not enough space to even start the note — same deferral.
                deferredFootnoteIds.push(footnoteId);
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
            // Block fits after expanding footnote area.
            currentContent.push(block.element.cloneNode(true) as HTMLElement);
            // `remainingHeight` tracks BODY consumption only — every other branch maintains it
            // that way, and the footnote reserve is applied separately via `effectiveRemainingHeight`
            // at the top of each iteration. Assigning `bodySpaceAfterExpansion - …` here folded the
            // reserve in a second time, so a later block would see a body budget short by the whole
            // footnote area. Consistency fix: no measurable difference on the documents tested, but
            // the two meanings must not coexist or the next change here inherits a latent bug.
            remainingHeight = bodySpaceAfterExpansion - bodyContentUsed - blockSpaceWithoutFootnotes;
            prevMarginBottomPt = block.marginBottomPt;
            currentFootnoteIds.push(...newFootnoteIds);
            currentFootnoteHeight = expandedFootnoteSpace;
          } else {
            // Still doesn't fit - start new page
            finishPage();
            const newPageFootnoteHeight = allBlockFootnoteIds.length > 0
              ? this.measureFootnotesHeight(allBlockFootnoteIds, dims.contentWidth, currentContinuation)
              : (currentContinuation ? this.measureContinuationHeight(currentContinuation, dims.contentWidth) : 0);
            const newPageSpace = block.marginTopPt + block.heightPt + block.marginBottomPt;
            currentContent.push(block.element.cloneNode(true) as HTMLElement);
            remainingHeight = effectiveContentHeight - newPageSpace;
            prevMarginBottomPt = block.marginBottomPt;
            // Merge, never replace: finishPage() may have just seeded this page with notes deferred
          // from the previous one, and overwriting here dropped them from the document.
          currentFootnoteIds = [...currentFootnoteIds, ...allBlockFootnoteIds];
            currentFootnoteHeight = newPageFootnoteHeight;
          }
        } else {
          // Block itself doesn't fit - start new page
          finishPage();
          // On new page, recalculate footnote height for just this block's footnotes
          // (plus any continuation from previous page)
          const newPageFootnoteHeight = allBlockFootnoteIds.length > 0
            ? this.measureFootnotesHeight(allBlockFootnoteIds, dims.contentWidth, currentContinuation)
            : (currentContinuation ? this.measureContinuationHeight(currentContinuation, dims.contentWidth) : 0);
          // Include full top margin
          const newPageSpace = block.marginTopPt + block.heightPt + block.marginBottomPt;
          currentContent.push(block.element.cloneNode(true) as HTMLElement);
          remainingHeight = effectiveContentHeight - newPageSpace;
          prevMarginBottomPt = block.marginBottomPt;
          // Merge, never replace: finishPage() may have just seeded this page with notes deferred
          // from the previous one, and overwriting here dropped them from the document.
          currentFootnoteIds = [...currentFootnoteIds, ...allBlockFootnoteIds];
          currentFootnoteHeight = newPageFootnoteHeight;
        }
      } else {
        // Block is taller than a page. Ordinary tables can be split at complete
        // row boundaries; every other block retains the established overflow path.
        const tableFragments = this.trySplitSimpleOversizedTable(block, dims, sectionIndex);
        if (tableFragments) {
          if (currentContent.length > 0 || currentContinuation) {
            finishPage();
          }
          blocks.splice(i, 1, ...tableFragments);
          i--;
          continue;
        }

        // Unsupported oversized blocks are intentionally left intact. Splitting
        // arbitrary HTML, merged tables, or footnote-bearing tables would be less
        // correct than the prior clipped fallback.
        if (currentContent.length > 0) {
          finishPage();
        }
        currentContent.push(block.element.cloneNode(true) as HTMLElement);
        // Merge, never replace: finishPage() may have just seeded this page with notes deferred
          // from the previous one, and overwriting here dropped them from the document.
          currentFootnoteIds = [...currentFootnoteIds, ...allBlockFootnoteIds];
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
    // Needed by substitutePageNumberFields: a section that restarts numbering counts from its own
    // first page, not from the document's.
    pageBox.dataset.pageInSection = String(pageInSection);

    // Where the three bands sit on this page (no re-measurement needed)
    const bands = this.getPageBands(dims, sectionIndex, pageInSection, pageNumber);

    // Add header if available for this section/page
    const headerSource = this.selectHeader(sectionIndex, pageInSection, pageNumber);

    if (headerSource) {
      const headerDiv = document.createElement("div");
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
      this.addPageFootnotes(pageBox, footnoteIds, dims, bands, footnoteHeight, continuation, partialFootnotes);

    }

    // Add footer if available for this section/page
    const footerSource = this.selectFooter(sectionIndex, pageInSection, pageNumber);
    if (footerSource) {
      const footerDiv = document.createElement("div");
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
    if (this.showPageNumbers) {
      const pageNum = document.createElement("div");
      pageNum.className = `${this.cssPrefix}number`;
      pageNum.textContent = String(pageNumber);
      pageBox.appendChild(pageNum);
    }

    // Add to container
    this.containerElement.appendChild(pageBox);

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

  // This convenience function is the read-only HTML viewer path. Keep the
  // lower-level engine conservative by default, while enabling paragraph
  // fragmentation here unless a caller deliberately opts out.
  const engine = new PaginationEngine(staging, pageContainer, {
    ...options,
    fragmentParagraphs: options.fragmentParagraphs ?? true,
  });
  return engine.paginate();
}

/**
 * Pagination engine for creating a PDF.js-style paginated view from HTML output.
 *
 * This module provides client-side pagination that measures rendered content
 * and flows it across fixed-size page containers based on document dimensions.
 */

import { formatPageNumber } from "./page-number-format.js";
import { alignBaselinesToWord, alignExactLineBaselines, applyUnroundedNormalLineHeights, insetClipBounds, preserveExactLineInk } from "./line-metrics.js";
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
import { EXACT_ROW_CELL } from "./exact-row-text.js";
import { omitPaintedBorderSide, PAINT as PAINTED_BORDER } from "./paragraph-borders.js";
import {
  layoutMarkupBalloons,
  marginNoteRegistry,
  readMarkupBalloonComments,
  scopeMarkupBalloons,
  type MarkupBalloonComment,
  type MarkupBalloonResult,
} from "./markup-balloons.js";
import { BalloonAppendix } from "./balloon-appendix.js";

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
  /** A fresh-page attempt could not honor this paragraph's keep-lines constraint. */
  relaxedKeepLines?: boolean;
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
/** The balloon layout's result, with the pages its appendix added after the last (if any). */
export type MarkupBalloonLayout = MarkupBalloonResult & { appendedPages: PageInfo[] };

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
  /** Present only when the caller supplied an exact layoutToken. */
  pageMap?: PageMap;
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
   * Comment placement: physical margin for PaginationEngine, outside for paginateHtml, or
   * "balloons" for Word's All Markup print (issue #1300, markup-balloons.ts): pages shrink beside
   * a markup area holding comments and, with Word-presentation HTML, formatting changes. As with
   * "outside", comments then have no PageMap citations.
   */
  commentColumnPlacement?: "margin" | "outside" | "balloons";
  /**
   * Under "balloons", what becomes of the rest of a comment whose balloon is cut short to fit beside
   * its page. "truncate" (the default) prints it nowhere, as Word's Save as PDF does. "appendix"
   * numbers each cut comment's box and prints the comments in full on pages added after the last,
   * as Word's Print to PDF does; those pages join the result and the PageMap.
   */
  balloonOverflow?: "truncate" | "appendix";
  /** Outside comment-column width in points. Default: 144 (2 inches). */
  commentColumnWidthPt?: number;
  /** Gap between the page and an outside comment column, in points. Default: 12. */
  commentColumnGapPt?: number;
  /**
   * Whether ordinary paragraphs may be fragmented across page boundaries.
   * Defaults to false for direct PaginationEngine callers; read-only viewer
   * entry points opt in explicitly.
   */
  fragmentParagraphs?: boolean;
  /**
   * Put each line's glyphs on Word's baseline once the pages are laid out (issue #942), moving
   * them with relative positioning so no line box changes. Defaults to false for direct
   * PaginationEngine callers; the read-only viewer entry points opt in. An editable surface
   * keeps it off (typed text would land in unmoved boxes), and the standalone export runs the
   * same pass itself after its clipping check.
   */
  alignBaselinesToWord?: boolean;
  /** Cooperative checkpoint for bounded non-yielding browser layout work. */
  checkCancellation?: () => void;
  /**
   * Skip fragment-identity stamping at the end of paginate(). Callers that mutate the page tree
   * afterwards (running-story placement) and then call normalizePageMapFragmentIdentities() would
   * otherwise pay the whole forced-layout pass twice and discard the first result. Default: false.
   */
  deferFragmentIdentities?: boolean;
  /**
   * Skip the "balloons" layout at the end of paginate(); the caller runs layoutMarkupBalloons()
   * itself once it has finished changing the page tree (the standalone export). Default: false.
   */
  deferMarkupBalloons?: boolean;
  /**
   * Points to withhold from the body height available to specific physical pages (1-based page
   * number → pt). Only the fitting budget shrinks; the page's rendered bands do not move. The
   * standalone export uses it for its bounded corrective pass: when a finished page still clips
   * its body — its layout ran a few points past what the placement model predicted — the export
   * lays the document out again from pristine HTML with that page's budget reduced by the
   * overflow, so the offending block moves or splits onto the next page (issue #848).
   */
  bodyBudgetReductionsPt?: Readonly<Record<number, number>>;
  /**
   * Incremental admission check invoked immediately before each physical page is allocated.
   * Export callers use this to enforce `finalPages` without constructing an over-limit DOM first.
   */
  checkPageCount?: (prospectivePageCount: number) => void;
  /** Exact invalidation tokens used to materialize an authoritative PageMap with the result. */
  layoutToken?: { documentVersion: number; rendererFingerprint: string };
}

// The share of the body band the note area (separator included) may take while body text remains.
// Word reference PDFs show none held back (issue #1178): a note takes everything below the line
// that cites it, and a carried note that does not end on a page takes that whole page. The 60% this
// used to be reserved room for body text that Word does not reserve.
// The band never takes body text already placed on the page, before or after the last body block
// (LibreOffice 25.8.7.3 keeps a twentieth of the band for the body once it is placed; Word does
// not, docs/ooxml_corner_cases.md).
const MAX_FOOTNOTE_AREA_RATIO = 1;
// Hidden measurement and final absolutely-positioned note bands differ by sub-pixel border/margin
// rounding in Chromium. Reserve a small physical-unit guard so the last baseline stays visible.
const FOOTNOTE_MEASUREMENT_GUARD_PT = 2;
/** Note-band space below which the paginator does not try to start a note on the citing page. */
const MIN_FOOTNOTE_START_PT = 20;
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
/** Answers {@link PaginationEngine.notesCanStartOnPage} has given for one page payload. */
interface NoteStartMemo {
  packedHeights: Map<string, number>;
  startsFrom: Map<string, number>;
  failsUpTo: Map<string, number>;
}

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
  /** A hard line break becomes the page boundary, rather than an empty line on the tail. */
  consumeBreak?: boolean;
  /** UTF-16 offset in the paragraph's flattened text. */
  textOffset: number;
  /** Prefer a hard-break or atomic-field boundary over a text-node boundary. */
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

/**
 * The top and bottom margins a block occupies in its parent's flow, in CSS px, AFTER CSS margin
 * collapsing through its first and last in-flow children (CSS 2.1 §8.3.1). getBoundingClientRect()
 * excludes such a collapsed-through margin, and getComputedStyle() reports only the element's own:
 * the converter wraps a table in an unstyled `div`, so the table's own top margin collapses
 * through the wrapper and occupies page height that neither of those reads accounted for. The
 * paginator's budget then fell a few pixels to a line short of the laid-out page, and the export
 * refused the clipped page (issue #848).
 *
 * A margin collapses with a parent's while the parent is an ordinary block box (display block or
 * list-item, overflow visible, not floated or absolutely positioned — anything else establishes
 * its own formatting context) with no border or padding on that side, the child is the parent's
 * first (last) in-flow block-level child, and no line content precedes (follows) it. For the
 * bottom edge the parent must also have an automatic height. Positive margins combine by maximum
 * and negative ones by minimum, as CSS collapses them.
 */
export function collapsedBlockMarginsPx(view: Window, element: HTMLElement): { topPx: number; bottomPx: number } {
  return {
    topPx: collapseMarginChain(view, element, "top"),
    bottomPx: collapseMarginChain(view, element, "bottom"),
  };
}

function collapseMarginChain(view: Window, element: HTMLElement, edge: "top" | "bottom"): number {
  const margins: number[] = [];
  let current: HTMLElement | null = element;
  while (current) {
    const style = view.getComputedStyle(current);
    margins.push(parseFloat(edge === "top" ? style.marginTop : style.marginBottom) || 0);
    current = collapsesThroughTo(view, current, style, edge);
  }
  const positive = Math.max(0, ...margins);
  const negative = Math.min(0, ...margins);
  return positive + negative;
}

/** The child whose `edge` margin collapses through `parent`'s, or null when none does. */
function collapsesThroughTo(
  view: Window,
  parent: HTMLElement,
  style: CSSStyleDeclaration,
  edge: "top" | "bottom",
): HTMLElement | null {
  const ordinaryBlock = (style.display === "block" || style.display === "list-item")
    && style.overflowX === "visible" && style.overflowY === "visible"
    && style.float === "none" && style.position !== "absolute" && style.position !== "fixed";
  if (!ordinaryBlock) return null;
  // Column boxes establish a formatting context: child margins stay inside the multicol
  // box and are already included in its measured height (CSS Multi-column Layout §2).
  if (style.columnCount !== "auto" || style.columnWidth !== "auto") return null;
  const border = parseFloat(edge === "top" ? style.borderTopWidth : style.borderBottomWidth) || 0;
  const padding = parseFloat(edge === "top" ? style.paddingTop : style.paddingBottom) || 0;
  if (border !== 0 || padding !== 0) return null;
  if (edge === "bottom" && parent.style.height !== "") return null;

  const nodes = Array.from(parent.childNodes);
  if (edge === "bottom") nodes.reverse();
  for (const node of nodes) {
    if (node.nodeType === Node.TEXT_NODE) {
      if ((node.textContent ?? "").trim() !== "") return null; // line content comes first
      continue;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) continue;
    const child = node as HTMLElement;
    const childStyle = view.getComputedStyle(child);
    if (childStyle.display === "none") continue;
    if (childStyle.position === "absolute" || childStyle.position === "fixed" || childStyle.float !== "none") continue;
    const blockLevel = childStyle.display === "block" || childStyle.display === "list-item"
      || childStyle.display === "table" || childStyle.display === "flow-root"
      || childStyle.display === "flex" || childStyle.display === "grid";
    return blockLevel ? child : null; // an inline-level child starts a line box
  }
  return null;
}

/** A page-space box in client coordinates: left/top inclusive, right/bottom exclusive. */
export interface VisibleBox {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

/**
 * Set by the paginator on a floating drawing it promotes out of its paragraph into the page box;
 * the value is the `data-source-anchor-id` of the element that hosted it. It survives into the
 * standalone export, whose offline re-verification measures through the same function.
 */
export const DRAWING_HOST_ATTRIBUTE = "data-docx-anchor-host";

/**
 * Set on a source-anchored element a floating drawing was promoted out of, so
 * {@link measureRenderedSource} looks up promoted drawings only for elements that have them.
 */
export const HOSTS_PROMOTED_DRAWING_ATTRIBUTE = "data-docx-hosts-promoted-drawing";

/**
 * Marks a block taller than its page that the paginator placed alone on a page. Word starts an
 * oversized exact row on a new page and clips at the paper edge; a footnote band retains its own
 * protected boundary (issue #1202).
 */
export const OVERSIZED_BLOCK_ATTRIBUTE = "data-pagination-oversized";

/** An oversized exact row responsible for this element's intentional vertical overflow. */
export function findOversizedExactRow(element: Element): HTMLTableRowElement | null {
  const block = element.closest(`[${OVERSIZED_BLOCK_ATTRIBUTE}]`);
  const body = block?.parentElement;
  if (!block || !body) return null;
  const height = body.getBoundingClientRect().height;
  // Column boxes are siblings of rows, but their height includes every row in their table.
  const columnTable = element.matches("colgroup, col") ? element.closest("table") : null;
  for (const clip of Array.from(block.querySelectorAll(EXACT_ROW_CELL))) {
    const row = clip.closest<HTMLTableRowElement>("tr");
    if (row && clip.getBoundingClientRect().height > height + 1
      && (row.contains(element) || element.contains(row) || columnTable?.contains(row))) return row;
  }
  return null;
}

/**
 * Whether the document itself clips `element` away whole: its layout box (`rect`) lies entirely
 * outside the exact-height row cell it sits in, as Word hides whatever an exact row cannot hold
 * (issue #1168). Such content renders nowhere, so it has no fragment and is not missing either.
 * Only that clip counts: the body band and the paginator's own clips also hide overflow, but there
 * a hidden source is content the layout lost, which PageMap completeness exists to catch.
 */
function clippedAwayByExactRow(element: HTMLElement, rect: DOMRect): boolean {
  if (!(rect.width > 0 && rect.height > 0)) {
    // A row whose exact height is no more than its borders and padding, its cells, and a table
    // made only of such rows have no height of their own: they show nothing, and the rows' clip
    // boxes have no height either. Only the element's own cells and rows count, not a nested
    // table's.
    const zeroHeightClip = (cell: Element) => {
      const clip = cell.querySelector(`:scope > ${EXACT_ROW_CELL}`);
      return clip !== null && !(clip.getBoundingClientRect().height > 0);
    };
    const zeroHeightRow = (row: HTMLTableRowElement) => {
      const cells = Array.from(row.cells);
      return cells.length > 0 && cells.some(zeroHeightClip);
    };
    if (element.localName === "tr") return zeroHeightRow(element as HTMLTableRowElement);
    if (element.localName === "td" || element.localName === "th") return zeroHeightClip(element);
    if (element.localName === "table") {
      const rows = Array.from((element as HTMLTableElement).rows);
      return rows.length > 0 && rows.every(zeroHeightRow);
    }
    return false;
  }
  // An exact row taller than its page is clipped where the page ends (issue #1202).
  const body = element.closest(`[${OVERSIZED_BLOCK_ATTRIBUTE}]`)?.parentElement;
  if (body && findOversizedExactRow(element)) {
    const boundary = body.style.overflowY === "visible" && body.parentElement
      ? body.parentElement.getBoundingClientRect().bottom : body.getBoundingClientRect().bottom;
    if (rect.top >= boundary) return true;
  }
  // Every enclosing exact row clips: a nested table's row can itself lie below its outer row.
  for (let clip = element.parentElement?.closest<HTMLElement>(EXACT_ROW_CELL); clip;
    clip = clip.parentElement?.closest<HTMLElement>(EXACT_ROW_CELL)) {
    const box = clip.getBoundingClientRect();
    if (rect.bottom <= box.top || rect.top >= box.bottom || rect.right <= box.left || rect.left >= box.right) return true;
  }
  return false;
}

/** True when `element` shows no text (beyond white space) and no image, SVG or canvas. */
function rendersOnlyItsLine(element: HTMLElement): boolean {
  return (element.textContent ?? '').replace(/[\s\u00a0]/g, '') === ''
    && element.querySelector('img, svg, canvas, [data-docx-drawing-anchor]') === null;
}

/** Self-contained closure for trusted host evaluation in the script-disabled print context. */
function createRenderedSourceMeasurer(
  clipInset: typeof insetClipBounds,
  drawingHostAttribute: string,
  hostsPromotedDrawingAttribute: string,
) {
  /**
   * Intersect a box with the page and with every ancestor that establishes an overflow clip before
   * the page root. getBoundingClientRect() reports layout outside those clips, which is not rendered
   * and therefore must not satisfy PageMap completeness or inflate portable geometry.
   */
  function visibleWithinPage(
    view: Window,
    element: HTMLElement,
    page: HTMLElement,
    pageRect: DOMRect,
    rect: DOMRect,
  ): VisibleBox {
    let left = Math.max(rect.left, pageRect.left);
    let top = Math.max(rect.top, pageRect.top);
    let right = Math.min(rect.right, pageRect.right);
    let bottom = Math.min(rect.bottom, pageRect.bottom);
    const clips = (value: string) =>
      value === "hidden" || value === "clip" || value === "scroll" || value === "auto";

    for (let ancestor = element.parentElement;
      ancestor && ancestor !== page;
      ancestor = ancestor.parentElement) {
      const style = view.getComputedStyle(ancestor);
      const clipsX = clips(style.overflowX);
      const clipsY = clips(style.overflowY);
      if (!clipsX && !clipsY && style.clipPath === "none") continue;
      const ancestorRect = ancestor.getBoundingClientRect();
      const scale = ancestorRect.width / ancestor.offsetWidth || 1;
      const inset = clipInset(style.clipPath, ancestorRect, scale);
      if (inset) {
        left = Math.max(left, inset.left);
        top = Math.max(top, inset.top);
        right = Math.min(right, inset.right);
        bottom = Math.min(bottom, inset.bottom);
      }
      // The exact-row path keeps the original visible rectangle; its expanded overflow clip is
      // deliberately wider than the band, so only the path bounds constrain the map.
      const expanded = inset !== null && style.clipPath.startsWith("path(");
      if (clipsX && !expanded) {
        left = Math.max(left, ancestorRect.left);
        right = Math.min(right, ancestorRect.right);
      }
      if (clipsY && !expanded) {
        top = Math.max(top, ancestorRect.top);
        bottom = Math.min(bottom, ancestorRect.bottom);
      }
    }
    return { left, top, right, bottom };
  }

  return function measureRenderedSource(
    view: Window,
    element: HTMLElement,
    page: HTMLElement,
    pageRect: DOMRect = page.getBoundingClientRect(),
  ): { rect: DOMRect; visible: VisibleBox } {
    const own = element.getBoundingClientRect();
    const ownResult = () => ({ rect: own, visible: visibleWithinPage(view, element, page, pageRect, own) });
    const ownMeasurable = own.width > 0 && own.height > 0;
    const hostsPromotedDrawing = element.hasAttribute(hostsPromotedDrawingAttribute);
    if (ownMeasurable && !hostsPromotedDrawing) return ownResult();

    const hostAnchorId = element.dataset.sourceAnchorId;
    const promoted = hostAnchorId
      ? Array.from(page.querySelectorAll<HTMLElement>(`[${drawingHostAttribute}]`))
        .filter((drawing) => drawing.getAttribute(drawingHostAttribute) === hostAnchorId)
      : [];
    if (ownMeasurable && promoted.length === 0) return ownResult();
    const rendered = [
      ...(ownMeasurable ? [element] : Array.from(element.querySelectorAll<HTMLElement>("*"))),
      ...promoted.flatMap((drawing) => [drawing, ...Array.from(drawing.querySelectorAll<HTMLElement>("*"))]),
    ];

    let left = Infinity;
    let top = Infinity;
    let right = -Infinity;
    let bottom = -Infinity;
    for (const child of rendered) {
      const style = view.getComputedStyle(child);
      if (style.display === "none" || style.visibility === "hidden") continue;
      const childRect = child.getBoundingClientRect();
      if (childRect.width <= 0 || childRect.height <= 0) continue;
      const visible = visibleWithinPage(view, child, page, pageRect, childRect);
      if (visible.right <= visible.left || visible.bottom <= visible.top) continue;
      left = Math.min(left, visible.left);
      top = Math.min(top, visible.top);
      right = Math.max(right, visible.right);
      bottom = Math.max(bottom, visible.bottom);
    }
    if (!(left < right && top < bottom)) return ownResult();
    return { rect: new DOMRect(left, top, right - left, bottom - top), visible: { left, top, right, bottom } };
  };
}

/**
 * Where a source-anchored element renders on `page`: `rect` is its layout box and `visible` the
 * part of it the page shows. Normally that is the element's own border box. An element the
 * paginator promoted floating drawings out of measures as the union of what it renders — its own
 * box (or, when that is empty, its descendants) and the drawings promoted out of it (linked by
 * {@link DRAWING_HOST_ATTRIBUTE}) — each clipped along its OWN ancestor chain, since a promoted
 * drawing sits outside the text column that clips its host (issue #849). Such a paragraph keeps a
 * line of its own when the drawing was its only content (issue #880), and its fragment still
 * encloses the drawing. Returns the element's own degenerate box when none of that renders.
 *
 * Fragment-identity normalization, PageMap materialization and the standalone export's offline
 * re-verification all measure through this one function, so they agree on which elements are
 * measurable and where.
 */
export const measureRenderedSource = createRenderedSourceMeasurer(
  insetClipBounds, DRAWING_HOST_ATTRIBUTE, HOSTS_PROMOTED_DRAWING_ATTRIBUTE,
);

/** @internal Owned runtime code only; never inserts a script or enables document JavaScript. */
export function renderedSourceMeasurerExpression(): string {
  return `(${createRenderedSourceMeasurer.toString()})(${insetClipBounds.toString()},
    ${JSON.stringify(DRAWING_HOST_ATTRIBUTE)}, ${JSON.stringify(HOSTS_PROMOTED_DRAWING_ATTRIBUTE)})`;
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
  private commentColumnPlacement: "margin" | "outside" | "balloons";
  private deferMarkupBalloons: boolean;
  private balloonOverflow: "truncate" | "appendix";
  private markupBalloons?: { comments: MarkupBalloonComment[]; commentPrefix: string; result?: MarkupBalloonLayout };
  private commentColumnWidthPt: number;
  private commentColumnGapPt: number;
  private fragmentParagraphs: boolean;
  private alignBaselines: boolean;
  private deferFragmentIdentities: boolean;
  private bodyBudgetReductionsPt: Readonly<Record<number, number>>;
  private cancellationCheckpoint?: () => void;
  private pageCountCheckpoint?: (prospectivePageCount: number) => void;
  private createdPageCount = 0;
  private layoutToken?: { documentVersion: number; rendererFingerprint: string };
  private hfRegistry: HeaderFooterRegistry;
  private footnoteRegistry: FootnoteRegistry;
  private footnoteSeparator: HTMLElement | null = null;
  private footnoteContinuationSeparator: HTMLElement | null = null;
  private commentMarginRegistry: Map<string, HTMLElement>;
  private commentReturnLinks: Array<{
    link: HTMLAnchorElement;
    href: string;
    ariaDisabled: string | null;
    tabIndex: string | null;
  }> = [];
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
    this.commentColumnPlacement = options.commentColumnPlacement ?? "margin";
    this.commentColumnWidthPt = options.commentColumnWidthPt ?? 144;
    this.commentColumnGapPt = options.commentColumnGapPt ?? 12;
    if (!["margin", "outside", "balloons"].includes(this.commentColumnPlacement)) {
      throw new RangeError("commentColumnPlacement must be one of margin, outside, balloons");
    }
    this.deferMarkupBalloons = options.deferMarkupBalloons ?? false;
    this.balloonOverflow = options.balloonOverflow ?? "truncate";
    if (this.balloonOverflow !== "truncate" && this.balloonOverflow !== "appendix") {
      throw new RangeError("balloonOverflow must be one of truncate, appendix");
    }
    if (!Number.isFinite(this.commentColumnWidthPt) || this.commentColumnWidthPt <= 0) {
      throw new RangeError("commentColumnWidthPt must be a finite positive number");
    }
    if (!Number.isFinite(this.commentColumnGapPt) || this.commentColumnGapPt < 0) {
      throw new RangeError("commentColumnGapPt must be a finite non-negative number");
    }
    this.fragmentParagraphs = options.fragmentParagraphs ?? false;
    this.alignBaselines = options.alignBaselinesToWord ?? false;
    this.deferFragmentIdentities = options.deferFragmentIdentities ?? false;
    this.bodyBudgetReductionsPt = options.bodyBudgetReductionsPt ?? {};
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

    // The balloon layout hides comment reference marks, which changes line breaks: before measuring.
    // Any other placement clears what an earlier balloon pagination left on these containers.
    const balloons = this.commentColumnPlacement === "balloons";
    if (balloons) {
      this.markupBalloons = readMarkupBalloonComments(marginNoteRegistry(this.stagingElement));
    }
    scopeMarkupBalloons(this.document, [this.stagingElement, this.containerElement], balloons,
      this.markupBalloons?.commentPrefix, this.cssPrefix);

    // Chromium rounds a `line-height: normal` line to a whole pixel, so single-spaced text
    // advanced 12.75pt where Word advances 12.65pt for 11pt Times New Roman, and the error
    // accumulated down every page (issues #850, #1151). Resolve each line height before
    // anything here is measured.
    applyUnroundedNormalLineHeights(this.stagingElement);

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
    if (this.commentColumnPlacement !== "margin") {
      // The PageMap describes physical page substrates. Outside review panels remain readable
      // UI, with an explicit source exclusion rather than invented in-page comment geometry.
      for (const source of this.commentMarginRegistry.values()) source.dataset.pageMapExclude = "true";
    }
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
            Number.isFinite(gap) ? gap : 36,
            section.dataset.sectionType === "continuous"
              && section === sectionsToProcess[sectionsToProcess.length - 1],
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
      const displayedPageNumber = parseInt(
        page.element.dataset.displayedPageNumber || String(page.pageNumber),
        10,
      );
      const header = this.selectHeader(page.sectionIndex, pageInSection, displayedPageNumber);
      const footer = this.selectFooter(page.sectionIndex, pageInSection, displayedPageNumber);
      for (const story of [header, footer]) {
        if (!story) continue;
        this.collectExpectedSourceAnchors(story, this.expectedPageMapAnchorIds);
        // The story's margin comments materialize with it (issue #1310).
        for (const reference of Array.from(story.querySelectorAll<HTMLElement>("[data-comment-id]"))) {
          const source = this.commentMarginRegistry.get(reference.dataset.commentId ?? "");
          if (source) this.collectExpectedSourceAnchors(source, this.expectedPageMapAnchorIds);
        }
      }
    }

    // Establish one active editor anchor and page-qualify every presentation fragment.
    // Full canonical source identities remain on all clones, including table cells.
    this.qualifyPageFragments(pages);
    this.transferVisibleFragmentTargets();
    if (!this.deferFragmentIdentities) this.normalizeVisiblePageFragments(pages);
    // Glyph placement only: run once every page and running story is final.
    if (this.alignBaselines) alignBaselinesToWord(this.containerElement);
    this.activateCommentReturnLinks(pages);
    this.lastPages = pages;
    if (!this.deferMarkupBalloons) this.layoutMarkupBalloons();

      const result = {
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
   * With "balloons" placement, lays out Word's markup area and balloons on the finished pages
   * (once; later calls return the first result). Runs inside paginate() unless deferMarkupBalloons
   * is set. Returns what could not be printed in full; undefined under other placements.
   */
  layoutMarkupBalloons(): MarkupBalloonLayout | undefined {
    if (!this.markupBalloons || this.markupBalloons.result) return this.markupBalloons?.result;
    if (this.lastPages.length === 0) throw new Error("paginate() must complete before layoutMarkupBalloons()");
    const staging = this.stagingElement.dataset;
    const font = staging.balloonFontFamily && Number.parseFloat(staging.balloonFontSize ?? "") > 0
      ? { family: staging.balloonFontFamily, sizePt: Number.parseFloat(staging.balloonFontSize!) }
      : undefined;
    const appendix = this.balloonOverflow === "appendix";
    const layout = layoutMarkupBalloons({
      document: this.document,
      pages: this.lastPages,
      comments: this.markupBalloons.comments,
      commentPrefix: this.markupBalloons.commentPrefix,
      cssPrefix: this.cssPrefix,
      font,
      numberCutBoxes: appendix,
    });
    const appendedPages: PageInfo[] = [];
    if (appendix && layout.cutComments.length > 0) {
      const pages = new BalloonAppendix(this.document, layout.cutComments, font ?? { family: "Segoe UI", sizePt: 9 },
        this.cssPrefix);
      while (!pages.finished) pages.fill(this.appendBlankPage(appendedPages));
    }
    return this.markupBalloons.result = { ...layout, appendedPages };
  }

  /**
   * A page after the last, in the last page's section, with no running stories: one more page of the
   * pagination, its result and its PageMap.
   */
  private appendBlankPage(appended: PageInfo[]): PageInfo {
    const last = this.lastPages[this.lastPages.length - 1];
    const number = (key: string, fallback: number) => Number.parseInt(last.element.dataset[key] ?? "", 10) || fallback;
    const page = this.createPage(last.dimensions, last.pageNumber + 1, last.sectionIndex,
      number("displayedPageNumber", last.pageNumber) + 1, [], number("pageInSection", 1) + 1, [], 0, null, undefined, true);
    page.element.dataset.balloonAppendix = "true";
    this.lastPages.push(page);
    appended.push(page);
    return page;
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
    const clippedAnchorIds = new Set<string>();
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

        const style = this.view.getComputedStyle(element);
        const deliberatelyHidden = style.display === "none" || style.visibility === "hidden";
        if (deliberatelyHidden) continue;
        const { rect, visible: visibleRect } = measureRenderedSource(this.view, element, page.element, pageRect);
        requiredAnchorIds.add(anchorId);
        const left = visibleRect.left;
        const top = visibleRect.top;
        const right = visibleRect.right;
        const bottom = visibleRect.bottom;
        if (rect.width <= 0 || rect.height <= 0 || right <= left || bottom <= top) {
          // A continued note/story clone can contain children clipped off this page which become
          // measurable on its next clone. Enforce completeness once every page has been inspected.
          if (clippedAwayByExactRow(element, rect)) clippedAnchorIds.add(anchorId);
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
          inTableCell: element.matches("td,th") || element.closest('td,th,[data-docx-frame-in-table-cell="true"]') !== null,
        });
      }
    }

    const missingAnchor = Array.from(requiredAnchorIds)
      .find((id) => !measuredAnchorIds.has(id) && !clippedAnchorIds.has(id));
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
        const { rect, visible } = measureRenderedSource(this.view, element, page.element, pageRect);
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
   * Ordinary cached fields are scoped to the CLONED header/footer regions. A page-number field in body text is
   * ordinary run content that the editor may make editable, and committing an edited block writes
   * back whatever text the DOM holds — rewriting it here would mean a body field commits a number
   * the document never contained. Legacy w:pgNum blocks are the exception: they have no cached
   * text and are explicitly non-editable, so their marked projection follows the current page
   * in every story while ordinary body fields keep their cached results.
   */
  private substitutePageNumberFields(totalPages: number): void {
    const boxes = this.containerElement.querySelectorAll<HTMLElement>(`.${this.cssPrefix}box`);
    for (const box of Array.from(boxes)) {
      const markers = box.querySelectorAll<HTMLElement>(
        `.${this.cssPrefix}header [data-field], .${this.cssPrefix}footer [data-field], [data-legacy-page-number="true"]`,
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
    this.prepareShapeClearance(section, dims);

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
      const { topPx: marginTopPx, bottomPx: marginBottomPx } = collapsedBlockMarginsPx(this.view, child);
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
      // The rule leaves its section when flattened, so preserve the section-scoped
      // print styling before measuring and painting the detached clone.
      const ruleStyle = this.view.getComputedStyle(rule);
      for (const property of ["height", "width", "max-width", "border", "background-color", "color",
        "margin-top", "margin-right", "margin-bottom", "margin-left"]) {
        clonedRule.style.setProperty(property, ruleStyle.getPropertyValue(property));
      }
      if (blocks.length === 0) {
        clonedRule.style.marginTop = `calc(${sectionStyle.marginTop} + ${ruleStyle.marginTop})`;
      }
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
          if (!clone.querySelector('[data-endnote-number]')) {
            const value = parseInt(item.getAttribute("value") || String(itemIndex + 1), 10);
            const marker = this.formatOrderedListMarker(
              Number.isFinite(value) ? value : itemIndex + 1,
              listStyle.listStyleType,
            );
            clone.insertBefore(this.document.createTextNode(`${marker}. `), clone.firstChild);
          }
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
   * a `column-count` container reproduces that geometry, and the paginator
   * then places each container as one ordinary measured block. A container grows
   * greedily until it would exceed the smallest page body available to the section.
   * Final continuous sections fill that height sequentially; other sections retain
   * balancing. A long columned section still splits across pages at block
   * boundaries. Each child lands in exactly one container, so anchors never
   * duplicate. An explicit page break child passes through as its own block, which
   * ends the current container and lets the normal flow logic turn the page.
   */
  private buildColumnBlocks(
    section: HTMLElement,
    dims: PageDimensions,
    sectionIndex: number,
    columnCount: number,
    columnGapPt: number,
    sequential = false,
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
      if (sequential) {
        container.dataset.paginationSequentialColumns = "true";
        container.style.columnFill = "auto";
        container.style.maxHeight = `${maxFragmentHeight}pt`;
      }
      for (const child of slice) {
        container.appendChild(child.cloneNode(true));
      }
      if (sequential) this.liftColumnSpaceBefore(container, dims);
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
      if (sequential && this.columnContainerOverflows(container, dims)) {
        // An indivisible child that exceeds all available columns keeps the established
        // oversized-block path; a constrained box must not hide it in an extra column.
        delete container.dataset.paginationSequentialColumns;
        container.style.columnFill = "balance";
        container.style.maxHeight = "";
      }
      let measured = this.measureElement(container, dims, sectionIndex);
      while (end < children.length && !isBreak(children[end])) {
        this.checkpoint();
        const candidate = makeContainer(children.slice(start, end + 1));
        const candidateMeasured = this.measureElement(candidate, dims, sectionIndex);
        if (candidateMeasured.heightPt > maxFragmentHeight
          || (sequential && this.columnContainerOverflows(candidate, dims))) break;
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

  /** Keep Word's paragraph space-before in the ordinary inter-block spacing budget. */
  private liftColumnSpaceBefore(container: HTMLElement, dims: PageDimensions): void {
    container.style.marginTop = "";
    const first = container.firstElementChild as HTMLElement | null;
    if (!first || !this.isWordParagraphElement(first)) return;
    const host = this.document.createElement("div");
    host.style.width = `${dims.contentWidth}pt`;
    host.appendChild(container);
    this.stagingElement.appendChild(host);
    try {
      const spaceBefore = first.dataset.paginationColumnSpaceBefore
        ?? this.view.getComputedStyle(first).marginTop;
      first.dataset.paginationColumnSpaceBefore = spaceBefore;
      first.style.marginTop = "0";
      container.style.marginTop = spaceBefore;
    } finally {
      container.remove();
      host.remove();
    }
  }

  /** A height-constrained multicol box overflows horizontally when it needs another column. */
  private columnContainerOverflows(element: HTMLElement, dims: PageDimensions): boolean {
    const host = this.document.createElement("div");
    host.style.width = `${dims.contentWidth}pt`;
    const clone = element.cloneNode(true) as HTMLElement;
    host.appendChild(clone);
    this.stagingElement.appendChild(host);
    try {
      if (clone.scrollWidth > clone.clientWidth + 1) return true;
      // Keep-lines paragraphs must fit as a whole. A constrained column container
      // can report only its first partial box, even when scrollWidth still fits.
      // Move that source paragraph to the next page before print clips its text.
      return Array.from(clone.querySelectorAll<HTMLElement>('[data-keep-lines="true"]'))
        .some(paragraph => paragraph.getClientRects().length > 1);
    } finally {
      host.remove();
    }
  }

  /**
   * Reflow a final continuous column section against the actual remaining body band.
   * A preceding title, running story or note reservation can make it shorter than the
   * fresh-page band used while collecting source blocks. Split at source-child boundaries
   * when the narrower band would create an extra horizontal column; the tail is reconsidered
   * normally on the following page. No source child is duplicated or discarded.
   */
  private reflowSequentialColumns(
    block: MeasuredBlock,
    dims: PageDimensions,
    availableHeightPt: number,
  ): MeasuredBlock[] | null {
    if (block.element.dataset.paginationSequentialColumns !== "true"
      || availableHeightPt <= 0) return null;
    const maximumHeight = this.smallestEffectiveContentHeight(dims, block.sectionIndex);
    const height = Math.min(availableHeightPt, maximumHeight);
    if (Math.abs(parseFloat(block.element.style.maxHeight) - height) < 0.01) return null;
    const children = Array.from(block.element.children) as HTMLElement[];
    const make = (slice: HTMLElement[], limit: number): HTMLElement => {
      const container = block.element.cloneNode(false) as HTMLElement;
      container.style.maxHeight = `${limit}pt`;
      for (const child of slice) container.appendChild(child.cloneNode(true));
      this.liftColumnSpaceBefore(container, dims);
      return container;
    };
    const whole = make(children, height);
    if (!this.columnContainerOverflows(whole, dims)) {
      return [this.measureElement(whole, dims, block.sectionIndex)];
    }
    let end = 0;
    for (let count = 1; count < children.length; count++) {
      this.checkpoint();
      const candidate = make(children.slice(0, count), height);
      if (this.columnContainerOverflows(candidate, dims)) break;
      if (children[count - 1].dataset.keepWithNext !== "true") end = count;
    }
    if (end === 0) return null;
    return [
      this.measureElement(make(children.slice(0, end), height), dims, block.sectionIndex),
      this.measureElement(make(children.slice(end), maximumHeight), dims, block.sectionIndex),
    ];
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
    preceding: readonly HTMLElement[] = [],
  ): MeasuredBlock {
    const measurementHost = this.document.createElement("div");
    measurementHost.style.position = "absolute";
    measurementHost.style.visibility = "hidden";
    measurementHost.style.left = "-9999px";
    measurementHost.style.width = `${dims.contentWidth}pt`;

    const measuredElement = element.cloneNode(true) as HTMLElement;
    for (const previous of preceding) measurementHost.appendChild(previous.cloneNode(true));
    measurementHost.appendChild(measuredElement);
    this.stagingElement.appendChild(measurementHost);

    const rect = measuredElement.getBoundingClientRect();
    const margins = collapsedBlockMarginsPx(this.view, measuredElement);
    const measured: MeasuredBlock = {
      element,
      sectionIndex,
      heightPt: pxToPt(rect.height),
      marginTopPt: pxToPt(margins.topPx),
      marginBottomPt: pxToPt(margins.bottomPx),
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
   * keeps a chain together when the whole chain or a valid first fragment can fit on a fresh page.
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
      this.getPageBands(dims, sectionIndex, 1, 1).bodyHeight,
      this.getPageBands(dims, sectionIndex, 2, 2).bodyHeight,
      this.getPageBands(dims, sectionIndex, 3, 3).bodyHeight
    );
  }

  /**
   * Builds a clone of a simple table wrapper containing a contiguous run of rows.
   * The caller supplies complete vertical-merge groups: no rowspan crosses a
   * fragment boundary.
   * Horizontal merges (colspan) are confined to their own row. `repeatedRows` (the table's
   * header rows, on a continuation) are copied in first, marked as repeats.
   */
  private createSimpleTableFragment(
    wrapper: HTMLElement | null,
    table: HTMLTableElement,
    body: HTMLTableSectionElement,
    rows: HTMLTableRowElement[],
    retainAnchor: boolean,
    repeatedRows: HTMLTableRowElement[] = [],
  ): HTMLElement {
    const wrapperClone = wrapper ? wrapper.cloneNode(false) as HTMLElement : null;
    const tableClone = table.cloneNode(false) as HTMLTableElement;

    // The eligibility gate permits only colgroups alongside the body. Keep each
    // colgroup so fixed and proportional column widths remain stable per fragment.
    for (const child of Array.from(table.children)) {
      if (child !== body) {
        tableClone.appendChild(child.cloneNode(true));
      }
    }

    const bodyClone = body.cloneNode(false) as HTMLTableSectionElement;
    for (const row of repeatedRows) {
      // A presentation copy: the original row stays the editor anchor, the bookmark target and
      // the comment, so a comment in a header row is shown once, on the first page.
      const repeat = row.cloneNode(true) as HTMLTableRowElement;
      this.makePresentationCopyInert(repeat, true);
      repeat.dataset.docxRepeatedHeaderRow = "true";
      bodyClone.appendChild(repeat);
    }
    for (const row of rows) {
      bodyClone.appendChild(row.cloneNode(true));
    }
    tableClone.appendChild(bodyClone);

    if (!retainAnchor) {
      wrapperClone?.removeAttribute("data-anchor");
      tableClone.removeAttribute("data-anchor");
    }

    if (!wrapperClone) return tableClone;
    wrapperClone.appendChild(tableClone);
    return wrapperClone;
  }

  /**
   * The row counts after which a table may split: a boundary is safe only after every cell
   * starting above it has ended. Overlapping merges in different columns extend the same
   * indivisible group; HTML rowspan=0 extends through the remainder of the row section. The last
   * entry is always `rows.length`.
   */
  private tableRowGroupEnds(rows: HTMLTableRowElement[]): number[] {
    const ends: number[] = [];
    let mergedEnd = 0;
    for (let index = 0; index < rows.length; index++) {
      this.checkpoint();
      for (const cell of Array.from(rows[index].cells)) {
        mergedEnd = Math.max(mergedEnd, cell.rowSpan === 0
          ? rows.length : Math.min(rows.length, index + cell.rowSpan));
      }
      if (index + 1 >= mergedEnd) ends.push(index + 1);
    }
    return ends;
  }

  /**
   * How many leading rows are the table's header rows (`w:tblHeader`), which a split repeats at
   * the top of each continuation: the run must end at a safe boundary and leave rows to carry.
   */
  private leadingHeaderRowCount(rows: HTMLTableRowElement[], ends: number[]): number {
    let count = 0;
    while (count < rows.length && rows[count].dataset.docxHeaderRow === "true") count++;
    return count === rows.length || !ends.includes(count) ? 0 : count;
  }

  /**
   * Splits an oversized, ordinary table at row boundaries. This only participates
   * in the existing oversized-block fallback; unsupported tables keep the previous
   * overflow behavior rather than risking broken table semantics.
   *
   * Word repeats a table's leading run of header rows (`w:tblHeader`) at the top of every page
   * the table continues on, whatever their height rule (issue #1171), so each continuation
   * starts with copies of them. When the copies leave no room to split the table, it splits as
   * it did without them (`repeatHeaderRows` false).
   */
  private trySplitSimpleOversizedTable(
    block: MeasuredBlock,
    dims: PageDimensions,
    sectionIndex: number,
    repeatHeaderRows = true,
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
      Array.from(table.children).some(child => child !== body && child.tagName !== "COLGROUP") ||
      table.querySelector("[data-footnote-id]") ||
      wrapper.querySelector("[data-footnote-id]")
    ) {
      return null;
    }

    const rows = Array.from(body.rows);
    const rowGroupEnds = this.tableRowGroupEnds(rows);
    const minimumContentHeight = this.smallestEffectiveContentHeight(dims, sectionIndex);
    // Use the source's full vertical margins while forming groups. Continuation
    // fragments later clear their joining margins, so this conservative bound
    // cannot create a fragment that overflows a header/footer variant.
    const maximumFragmentHeight = minimumContentHeight - block.marginTopPt - block.marginBottomPt;
    if (maximumFragmentHeight <= 0) {
      return null;
    }

    // The header run must end at a safe boundary and leave body rows to carry.
    const headerCount = repeatHeaderRows ? this.leadingHeaderRowCount(rows, rowGroupEnds) : 0;
    const headerRows = rows.slice(0, headerCount);
    const withoutRepeats = (): MeasuredBlock[] | null => {
      if (headerCount === 0) return null;
      const fragments = this.trySplitSimpleOversizedTable(block, dims, sectionIndex, false);
      if (fragments && headerRows.some(row => row.querySelector(EXACT_ROW_CELL) !== null && !rowFits(row))) {
        // Word starts the text after a table with an indivisible oversized exact header on
        // a fresh page. finishPage ignores this marker when the last row already turned it.
        fragments.push({
          ...block,
          element: this.document.createElement("div"),
          heightPt: 0,
          marginTopPt: 0,
          marginBottomPt: 0,
          keepWithNext: false,
          keepLines: false,
          pageBreakBefore: false,
          isPageBreak: true,
        });
      }
      return fragments;
    };

    const groups: HTMLTableRowElement[][] = [];
    // Groups that are a page on their own however they are cut: one exact-height row, or the part
    // of a row that holds one block taller than a page. They are placed alone on a page and clipped.
    const oversizedGroups = new Set<HTMLTableRowElement[]>();
    const rowFits = (row: HTMLTableRowElement) => this.measureElement(
      this.createSimpleTableFragment(wrapper, table, body, [row], true), dims, block.sectionIndex,
    ).heightPt <= maximumFragmentHeight;
    let start = 0;
    let boundaryIndex = 0;
    while (start < rows.length) {
      this.checkpoint();
      let end = start;
      while (boundaryIndex < rowGroupEnds.length) {
        this.checkpoint();
        const candidateEnd = rowGroupEnds[boundaryIndex];
        const candidate = this.createSimpleTableFragment(
          wrapper,
          table,
          body,
          rows.slice(start, candidateEnd),
          start === 0,
          start === 0 ? [] : headerRows,
        );
        const measured = this.measureElement(candidate, dims, block.sectionIndex);
        if (measured.heightPt > maximumFragmentHeight) {
          break;
        }
        end = candidateEnd;
        boundaryIndex++;
      }

      // Header rows left alone at the foot of the first page do not count as its content, and a
      // header run that is not followed on the page by a complete row group is not repeated.
      if (headerCount > 0 && (end === start || (start === 0 && end <= headerCount))) {
        return withoutRepeats();
      }
      if (end === start) {
        // Even one complete row group cannot fit a page. Word breaks a row taller than a page
        // between the blocks its cells hold (issue #1202); a merged group keeps the established
        // overflow fallback rather than cutting a merged cell.
        const row = rows[start];
        if (rowGroupEnds[boundaryIndex] !== start + 1) return null;
        for (const piece of this.splitRowAtBlocks(row, rowFits)) {
          const group = [piece.row];
          groups.push(group);
          if (piece.oversized) oversizedGroups.add(group);
        }
        start++;
        boundaryIndex++;
        continue;
      }

      groups.push(rows.slice(start, end));
      start = end;
    }

    if (groups.length < 2) {
      return withoutRepeats();
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
        isFirst,
        isFirst ? [] : headerRows,
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
        !oversizedGroups.has(groups[index]) &&
        measured.heightPt + measured.marginTopPt + measured.marginBottomPt >
        minimumContentHeight
      ) {
        return withoutRepeats();
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
   * Cuts a row taller than a page into rows that each fit one (`fits`), as Word breaks such a row
   * (issue #1202): every cell keeps, in order, the blocks it holds, and a cell whose blocks are all
   * placed is left empty. A block taller than a page is a piece of its own, marked `oversized`. An
   * exact-height row is not cut: Word starts it on a new page and clips it there. Neither is a row
   * whose cells hold anything but blocks.
   */
  private splitRowAtBlocks(
    row: HTMLTableRowElement,
    fits: (row: HTMLTableRowElement) => boolean,
  ): Array<{ row: HTMLTableRowElement; oversized: boolean }> {
    const cells = Array.from(row.cells);
    const whole = [{ row, oversized: true }];
    if (cells.some(cell => cell.querySelector(`:scope > ${EXACT_ROW_CELL}`))) return whole;
    const pending = cells.map(cell => Array.from(cell.childNodes));
    if (pending.some(nodes => nodes.some(node => node.nodeType !== this.view.Node.ELEMENT_NODE
      && (node.nodeType !== this.view.Node.TEXT_NODE || (node.textContent ?? "").trim() !== "")))) return whole;
    const blocks = pending.map(nodes => nodes.filter(node => node.nodeType === this.view.Node.ELEMENT_NODE) as HTMLElement[]);

    const build = (parts: HTMLElement[][], first: boolean) => {
      const piece = row.cloneNode(false) as HTMLTableRowElement;
      cells.forEach((cell, index) => {
        const clone = cell.cloneNode(false) as HTMLElement;
        for (const element of parts[index]) clone.appendChild(element.cloneNode(true));
        if (!first) {
          clone.removeAttribute("id");
          clone.removeAttribute("data-anchor");
        }
        piece.appendChild(clone);
      });
      if (!first) {
        piece.removeAttribute("id");
        piece.removeAttribute("data-anchor");
      }
      return piece;
    };
    const pieces: Array<{ row: HTMLTableRowElement; oversized: boolean }> = [];
    while (blocks.some(list => list.length > 0)) {
      this.checkpoint();
      // A row is as tall as its tallest cell, so each cell takes the most of its blocks that fit.
      let take = blocks.map((list, index) => {
        let count = 0;
        while (count < list.length
          && fits(build(blocks.map((other, j) => j === index ? list.slice(0, count + 1) : []), true))) count++;
        return count;
      });
      const oversized = take.every(count => count === 0);
      if (oversized) take = blocks.map(list => Math.min(1, list.length));
      pieces.push({ row: build(blocks.map((list, index) => list.slice(0, take[index])), pieces.length === 0), oversized });
      blocks.forEach((list, index) => list.splice(0, take[index]));
    }
    return pieces.length > 0 ? pieces : whole;
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
    const lastTextByAtomic = new Map<HTMLElement, Text>();
    const breaks: Array<{ element: HTMLElement; textOffset: number }> = [];
    const walker = this.document.createTreeWalker(paragraph,
      this.view.NodeFilter.SHOW_TEXT | this.view.NodeFilter.SHOW_ELEMENT);
    const flattenedChunks: string[] = [];
    let flattenedLength = 0;
    let node: Node | null;
    while ((node = walker.nextNode())) {
      this.checkpoint();
      if (node.nodeType === this.view.Node.ELEMENT_NODE) {
        if ((node as Element).localName === "br") {
          breaks.push({ element: node as HTMLElement, textOffset: flattenedLength });
          flattenedChunks.push("\n");
          flattenedLength++;
        }
        continue;
      }
      const textNode = node as Text;
      textStarts.set(textNode, flattenedLength);
      textNodes.push(textNode);
      flattenedChunks.push(textNode.data);
      flattenedLength += textNode.data.length;
      let field = textNode.parentElement?.closest<HTMLElement>("[data-field], [data-docx-word-line=true]") ?? null;
      while (field?.parentElement?.closest<HTMLElement>("[data-field]")) {
        field = field.parentElement.closest<HTMLElement>("[data-field]");
      }
      if (field && paragraph.contains(field)) lastTextByAtomic.set(field, textNode);
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
      if (!candidate.consumeBreak && this.isNonBreakingTextBoundary(flattenedText, textOffset)) return;
      if (!this.isLegalParagraphLineBoundary(paragraph, flattenedText, textOffset)) return;
      if (graphemeBoundaries) {
        if (!graphemeBoundaries.has(textOffset)) return;
      } else if (!candidate.consumeBreak && !this.isConservativeFallbackTextBoundary(flattenedText, textOffset)) {
        return;
      }
      candidates.push(candidate);
    };

    for (const { element, textOffset } of breaks) {
      this.checkpoint();
      // Fields and marker wrappers remain atomic. Do not cut after an empty
      // line: a trailing BR alone does not create its final blank line in HTML.
      // Keeping a consecutive break sequence on the tail preserves those lines.
      if (element.closest("[data-field], [data-list-marker], a[data-comment-id]")
          || flattenedText[textOffset - 1] === "\n") continue;
      const parent = element.parentNode;
      if (!parent) continue;
      addCandidate({ node: parent, offset: Array.prototype.indexOf.call(parent.childNodes, element),
        textOffset, priority: 2, consumeBreak: true });
    }

    for (const node of textNodes) {
      this.checkpoint();
      const start = textStarts.get(node)!;
      // Field results and converter list markers are semantic atoms. A field
      // receives one boundary outside its wrapper below; a list marker stays
      // with the first real text fragment and is never emitted by itself.
      const atomic = node.parentElement?.closest<HTMLElement>(
        "[data-field], [data-list-marker], a[data-comment-id], [data-docx-word-line=true]",
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

    // Fields and fitted Word lines split only after their complete outer wrapper.
    // Starting the tail at the parent's child boundary prevents Range from
    // cloning an empty `[data-field]` shell that substitution could repopulate.
    for (const field of Array.from(paragraph.querySelectorAll<HTMLElement>("[data-field], [data-docx-word-line=true]"))) {
      this.checkpoint();
      if (field.parentElement?.closest("[data-field]")) continue;
      const last = lastTextByAtomic.get(field);
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
    // the consumed break, then an outer atomic boundary, in stable document order.
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
   * Phase one handles text-like paragraph descendants and hard line breaks.
   * Objects, notes, and out-of-flow/inline-block content require their own
   * line-layout rules and retain the established whole-block fallback instead
   * of risking broken content or duplicate anchors.
   */
  private canFragmentParagraph(block: MeasuredBlock, relaxKeepLines = false): boolean {
    if (!this.fragmentParagraphs) {
      return false;
    }

    const paragraph = block.element;
    if (
      block.keepWithNext ||
      (block.keepLines && !relaxKeepLines && !block.relaxedKeepLines) ||
      block.pageBreakBefore ||
      block.isPageBreak ||
      (paragraph.dataset.widowControl === "true" && !relaxKeepLines && !block.relaxedKeepLines) ||
      paragraph.hasAttribute("contenteditable")
    ) {
      return false;
    }

    // A body paragraph may cite notes: the paginator places a fragment's notes with the
    // fragment, so a citation never splits from its note (see notesCanStartOnPage).
    return this.canRangeFragmentParagraph(paragraph, false, true,
      relaxKeepLines || block.relaxedKeepLines === true);
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
    bodyParagraph: boolean = false,
    relaxKeepLines = false,
  ): boolean {
    // The converter writes a Word paragraph with an outline level (a heading style) as h1-h6;
    // it breaks across pages like any other paragraph unless it keeps its lines together.
    const wordParagraph = bodyParagraph
      ? this.isWordParagraphElement(paragraph)
      : paragraph.tagName === "P";
    if (
      !wordParagraph ||
      paragraph.dataset.keepWithNext === "true" ||
      (paragraph.dataset.keepLines === "true" && !relaxKeepLines) ||
      paragraph.dataset.pageBreakBefore === "true" ||
      (paragraph.dataset.widowControl === "true" && !relaxKeepLines) ||
      paragraph.hasAttribute("contenteditable")
    ) {
      return false;
    }

    // Non-textual/out-of-flow descendants still need dedicated fragmenters.
    // Inline ids/editor anchors are reconciled after Range cloning, and list
    // markers are kept atomically with the leading text fragment.
    const unsupportedDescendants = [
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
      ...(bodyParagraph ? [] : ["[data-footnote-id]"]),
      "[contenteditable]",
    ].join(", ");
    if (paragraph.querySelector(unsupportedDescendants)) {
      return false;
    }

    const isValidatedEndnote = paragraph.dataset.paginationSafeEndnote === "true";
    return isValidatedEndnote || this.hasRangeFragmentSafeLayout(paragraph, allowInlineRoot, relaxKeepLines);
  }

  /**
   * A range clone preserves nested inline formatting exactly. Anything that establishes its own
   * box/layout context is deferred until a future fragmenter can model it accurately. Callers
   * must invoke this while the paragraph is attached to the styled document.
   */
  private hasRangeFragmentSafeLayout(
    paragraph: HTMLElement,
    allowInlineRoot: boolean = false,
    relaxKeepLines = false,
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
      (paragraphStyle.breakInside === "avoid" && !relaxKeepLines) ||
      paragraphStyle.pageBreakBefore !== "auto" ||
      paragraphStyle.pageBreakAfter !== "auto" ||
      (paragraphStyle.pageBreakInside === "avoid" && !relaxKeepLines)
    ) {
      return false;
    }

    // A list item's number and the tab after a number render as aligned boxes ahead of the
    // text. No fragment endpoint falls inside them, so they always stay whole with the first
    // fragment and their inner layout is never cut.
    const leading = this.leadingAtomicBoxes(paragraph);
    for (const descendant of Array.from(paragraph.querySelectorAll<HTMLElement>("*"))) {
      if (leading.some((box) => box.contains(descendant))) continue;
      const style = this.view.getComputedStyle(descendant);
      const wordLine = descendant.dataset.docxWordLine === "true"
        && Number(descendant.parentElement!.dataset.docxWordFit) > 0
        && descendant.children.length === 0;
      if (
        (style.display !== "inline" && !(wordLine && style.display === "block")) ||
        style.position !== "static" ||
        style.float !== "none" ||
        (style.whiteSpace !== "normal" && style.whiteSpace !== "pre-wrap" && !(wordLine && style.whiteSpace === "pre"))
      ) {
        return false;
      }
      if (descendant.localName === "br" && (style.clear !== "none"
          || style.breakBefore !== "auto" || style.breakAfter !== "auto")) return false;
    }
    return true;
  }

  /**
   * Whether the text after `endpoint` begins a new visual line of the laid-out `paragraph`: the
   * nearest glyph after it sits lower than the nearest glyph before it by more than half a line.
   * A raised citation number in the middle of a line sits higher, so it stays on its line.
   */
  private startsVisualLine(paragraph: HTMLElement, endpoint: ParagraphFragmentEndpoint): boolean {
    if (endpoint.consumeBreak) return true;
    const range = this.document.createRange();
    range.selectNodeContents(paragraph);
    range.setEnd(endpoint.node, endpoint.offset);
    const before = this.rangeEdgeGlyph(range, false);
    range.selectNodeContents(paragraph);
    range.setStart(endpoint.node, endpoint.offset);
    const after = this.rangeEdgeGlyph(range, true);
    return !before || !after || after.top > before.top + before.height / 2;
  }

  /** The first/last painted text glyph in a range, excluding whitespace and zero-width marks. */
  private rangeEdgeGlyph(range: Range, forward: boolean): DOMRect | null {
    const root = range.commonAncestorContainer;
    const texts: Text[] = [];
    if (root.nodeType === this.view.Node.TEXT_NODE) texts.push(root as Text);
    else {
      const walker = this.document.createTreeWalker(root, this.view.NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        this.checkpoint();
        if (range.intersectsNode(node)) texts.push(node as Text);
      }
    }
    if (!forward) texts.reverse();
    for (const text of texts) {
      const start = text === range.startContainer ? range.startOffset : 0;
      const end = text === range.endContainer ? range.endOffset : text.data.length;
      for (let k = 0; k < end - start; k++) {
        this.checkpoint();
        const index = forward ? start + k : end - 1 - k;
        if (!this.hasVisibleText(text.data[index])) continue;
        const character = this.document.createRange();
        character.setStart(text, index);
        character.setEnd(text, index + 1);
        const rect = character.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) return rect;
      }
    }
    return null;
  }

  private rangeHasMultipleTextLines(range: Range): boolean {
    const first = this.rangeEdgeGlyph(range, true);
    const last = this.rangeEdgeGlyph(range, false);
    return !!first && !!last && last.top > first.top + first.height / 2;
  }

  /**
   * The boxes a paragraph opens with before its first text: its list marker, a note's number,
   * and any box that holds no text, such as the tab that follows either.
   */
  private leadingAtomicBoxes(paragraph: HTMLElement): HTMLElement[] {
    const boxes: HTMLElement[] = [];
    for (const node of Array.from(paragraph.childNodes)) {
      if (node.nodeType === this.view.Node.TEXT_NODE) {
        if (this.hasVisibleText(node.textContent ?? "")) break;
        continue;
      }
      if (!(node instanceof this.view.HTMLElement)) continue;
      if (node.dataset.listMarker === "true" || node.dataset.noteRun === "true"
          || node.classList.contains("footnote-number") || node.classList.contains("endnote-number")
          || !this.hasVisibleText(node.textContent ?? "")) {
        boxes.push(node);
        continue;
      }
      break;
    }
    return boxes;
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
      // Word can place the final natural line box against the page edge;
      // inter-line leading belongs before a following line, not after this head.
      const run = fragment.querySelector<HTMLElement>('[data-docx-word-fit]');
      const last = run?.lastElementChild as HTMLElement | null;
      const natural = Number(run?.dataset.docxWordFit);
      if (last && natural > 0 && natural < parseFloat(run!.style.lineHeight))
        last.style.lineHeight = `${natural}px`;
      else if (run?.dataset.docxWordFit === "0") {
        // An unchanged inline flow has no private line boxes. Measure its
        // complete prefix in the laid-out source, then trim only this head.
        const plain = run;
        const text = paragraph.firstElementChild?.firstChild;
        const height = parseFloat(plain.style.getPropertyValue("--docx-word-native-height"));
        const pitch = parseFloat(plain.style.lineHeight);
        if (!plain.children.length && text?.nodeType === 3 && range.endContainer === text && height > 0 && pitch > height) {
          let end = range.endOffset;
          const content = text.textContent!;
          while (end > 0 && /\s/.test(content[end - 1])) {
            if ((end & 255) === 0) this.checkpoint();
            end--;
          }
          if (end > 0) {
            const glyph = range.cloneRange();
            glyph.setStart(text, 0); glyph.setEnd(text, 1);
            const firstBox = glyph.getBoundingClientRect();
            glyph.setStart(text, end - 1); glyph.setEnd(text, end);
            const lastBox = glyph.getBoundingClientRect();
            if (firstBox.height > 0 && lastBox.height > 0 && lastBox.top >= firstBox.top) {
              const lines = Math.round((lastBox.top - firstBox.top) / pitch);
              fragment.style.setProperty("--docx-word-text-height", `${lines * pitch + height}px`);
              fragment.style.height = "var(--docx-word-text-height)";
            }
          }
        }
      }
      // A page break is not the end of the source paragraph: its last line
      // still takes the paragraph's justification rather than last-line alignment.
      if (this.view.getComputedStyle(paragraph).textAlign === "justify")
        fragment.style.setProperty("text-align-last", "justify", "important");
    }

    return fragment;
  }

  /**
   * Range clones repeat an inline ancestor when the split lands inside it. Keep
   * semantic wrappers (links, comments, formatting) on both sides, but retain a
   * duplicated HTML/editor identity only on the leading fragment. Targets that
   * occur wholly after the split are absent from `head` and remain on `tail`.
   */
  private reconcileFragmentIdentities(
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
      /** Break only where a visual line starts, never within a line. */
      lineStartsOnly?: boolean;
      /** Keep at least two text lines on each side when relaxing a kept paragraph. */
      preserveWidowLines?: boolean;
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
    const legalBreak = (endpoint: ParagraphFragmentEndpoint) => {
      if (options.lineStartsOnly && !this.startsVisualLine(paragraph, endpoint)) return false;
      if (!options.preserveWidowLines) return true;
      const range = this.document.createRange();
      range.selectNodeContents(paragraph);
      range.setEnd(endpoint.node, endpoint.offset);
      if (!this.rangeHasMultipleTextLines(range)) return false;
      range.selectNodeContents(paragraph);
      range.setStart(endpoint.node, endpoint.offset + (endpoint.consumeBreak ? 1 : 0));
      return this.rangeHasMultipleTextLines(range);
    };
    if (best && !legalBreak(best.endpoint)) {
      // Fit is monotone in the head's length, so the nearest earlier endpoint that ends a line
      // fits as well: it is the head without the partly-filled last line.
      let index = endpoints.indexOf(best.endpoint) - 1;
      while (index >= 0 && !legalBreak(endpoints[index])) index--;
      if (index < 0) {
        best = null;
      } else {
        const headRange = this.document.createRange();
        headRange.setStart(paragraph, 0);
        headRange.setEnd(endpoints[index].node, endpoints[index].offset);
        best = { endpoint: endpoints[index], head: this.createParagraphFragment(paragraph, headRange, true, false) };
      }
    }
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
    tailRange.setStart(best.endpoint.node, best.endpoint.offset + (best.endpoint.consumeBreak ? 1 : 0));
    tailRange.setEnd(paragraph, paragraph.childNodes.length);
    const tailContents = tailRange.cloneContents();
    if (!this.hasVisibleFragmentText(tailContents)) return { kind: "indivisible" };

    const tail = this.createParagraphFragment(paragraph, tailRange, false, true);
    this.reconcileFragmentIdentities(best.head, tail);

    return {
      kind: "split",
      head: best.head,
      tail,
    };
  }

  /**
   * The converter's box around a run of paragraphs that share one border (`w:pBdr`): a `div` whose
   * element children are all Word paragraphs and that draws a border. `element` must be laid out.
   */
  private isBorderGroup(element: HTMLElement): boolean {
    if (element.localName !== "div" || element.children.length === 0
      || !Array.from(element.children).every(child => this.isWordParagraphElement(child as HTMLElement))) return false;
    // The PDF export replaces declared-width borders with padding and paints them after pagination.
    const style = this.view.getComputedStyle(element);
    // A frame or other out-of-flow box is not a run of body paragraphs.
    if (style.position === "absolute" || style.position === "fixed" || style.float !== "none"
      || element.hasAttribute("data-docx-frame-container")) return false;
    if (element.hasAttribute(PAINTED_BORDER)) return true;
    return [style.borderTopStyle, style.borderBottomStyle, style.borderLeftStyle, style.borderRightStyle]
      .some(value => value !== "none" && value !== "hidden");
  }

  /**
   * Splits a bordered paragraph group where Word breaks it across pages (issue #1202): after the last
   * paragraph that fits, or inside the paragraph that does not, as an ordinary paragraph splits. Word
   * leaves the border open at the break: the first part has no bottom border and the rest no top
   * border, so each part keeps its side borders. A group that cites notes, or keeps its lines or its
   * next paragraph, keeps the whole-block fallback.
   */
  private tryFragmentBorderGroup(
    block: MeasuredBlock,
    dims: PageDimensions,
    availableHeightPt: number,
    effectiveMarginTopPt: number,
  ): MeasuredBlock[] | null {
    if (!this.fragmentParagraphs || block.keepWithNext || block.keepLines || block.pageBreakBefore
      || block.isPageBreak || block.element.querySelector("[data-footnote-id]")) return null;

    // Continuations are detached clones: lay the group out in the staging context first.
    const host = !block.element.isConnected ? this.document.createElement("div") : null;
    const group = host ? block.element.cloneNode(true) as HTMLElement : block.element;
    if (host) {
      host.style.cssText = `position:absolute;visibility:hidden;left:-9999px;width:${dims.contentWidth}pt`;
      host.appendChild(group);
      this.stagingElement.appendChild(host);
    }
    try {
      if (!this.isBorderGroup(group)) return null;
      const paragraphs = Array.from(group.children) as HTMLElement[];
      const part = (content: HTMLElement[], open: "bottom" | "top") => {
        const wrapper = group.cloneNode(false) as HTMLElement;
        wrapper.style.setProperty(`border-${open}`, "none", "important");
        wrapper.style.setProperty(`padding-${open}`, "0", "important");
        wrapper.style.setProperty(`margin-${open}`, "0", "important");
        omitPaintedBorderSide(wrapper, open);
        for (const paragraph of content) wrapper.appendChild(paragraph);
        return wrapper;
      };
      const fits = (content: HTMLElement[]) => effectiveMarginTopPt
        + this.measureElement(part(content, "bottom"), dims, block.sectionIndex).heightPt <= availableHeightPt;
      const clones = (from: number, to: number) => paragraphs.slice(from, to).map(p => p.cloneNode(true) as HTMLElement);

      // Whole paragraphs that fit, never ending after one that keeps with its next.
      let whole = 0;
      while (whole + 1 < paragraphs.length && fits(clones(0, whole + 1))) whole++;
      while (whole > 0 && paragraphs[whole - 1].dataset.keepWithNext === "true") whole--;

      let head: HTMLElement[] = clones(0, whole);
      let tail: HTMLElement[] = clones(whole, paragraphs.length);
      const next = paragraphs[whole];
      if (next.dataset.keepWithNext !== "true" && this.canRangeFragmentParagraph(next, false, true)) {
        const split = this.splitParagraphAtLargestFit(next, (part) => fits([...clones(0, whole), part]));
        if (split.kind === "split") {
          head = [...head, split.head];
          tail = [split.tail, ...clones(whole + 1, paragraphs.length)];
        }
      }
      if (head.length === 0) return null;
      // Word does not repeat a paragraph's space before at the top of the page the rest moves to.
      tail[0].style.setProperty("margin-top", "0", "important");
      return [part(head, "bottom"), part(tail, "top")].map(element => ({
        ...this.measureElement(element, dims, block.sectionIndex),
        element,
        keepWithNext: false,
        keepLines: false,
        pageBreakBefore: false,
        isPageBreak: false,
      }));
    } finally {
      host?.remove();
    }
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
    effectiveMarginTopPt: number,
    notesStart?: (footnoteIds: string[], headHeightPt: number) => boolean,
    preceding: readonly HTMLElement[] = [],
    relaxKeepLines = false,
  ): MeasuredBlock[] | null {
    if (availableHeightPt <= effectiveMarginTopPt) return null;
    if (block.element.localName === "div") {
      return this.tryFragmentBorderGroup(block, dims, availableHeightPt, effectiveMarginTopPt);
    }

    // Continuations are detached Range clones. Validate their computed layout
    // in the same staging context used for measurement before splitting again.
    // A preceding float still narrows this paragraph's first lines. Measure candidate heads
    // beside the actual page prefix, not in a full-width paragraph in isolation. Continuations
    // on another page do not inherit that page-owned exclusion.
    const floatSelector = '[data-docx-edge-picture],[data-docx-shape-exclusion]';
    const prefix = preceding.some(element => element.matches(floatSelector) || element.querySelector(floatSelector))
      ? preceding : [];
    const host = !block.element.isConnected || prefix.length > 0 ? this.document.createElement("div") : null;
    const paragraph = host ? block.element.cloneNode(true) as HTMLElement : block.element;
    if (host) {
      host.style.cssText = `position:absolute;visibility:hidden;left:-9999px;width:${dims.contentWidth}pt`;
      for (const previous of prefix) host.appendChild(previous.cloneNode(true));
      host.appendChild(paragraph);
      this.stagingElement.appendChild(host);
    }
    let split: ParagraphSplitAttempt;
    try {
      if (!this.canFragmentParagraph({ ...block, element: paragraph }, relaxKeepLines)) return null;

      // Word moves a citing line, never part of one, to the page where its note can begin:
      // stop the head above the first citation whose note cannot start beside it. A head
      // holding any part of that line is taller than the citation's top.
      let headLimitPt = availableHeightPt;
      if (notesStart) {
        const top = paragraph.getBoundingClientRect().top;
        const cited: string[] = [];
        for (const reference of Array.from(paragraph.querySelectorAll<HTMLElement>("[data-footnote-id]"))) {
          const footnoteId = reference.dataset.footnoteId;
          if (!footnoteId || cited.includes(footnoteId)) continue;
          const marker = this.document.createRange();
          marker.selectNodeContents(reference);
          const rect = marker.getBoundingClientRect();
          if (rect.height === 0) continue;
          if (notesStart([...cited, footnoteId], pxToPt(rect.bottom - top))) {
            cited.push(footnoteId);
            continue;
          }
          headLimitPt = Math.min(headLimitPt, effectiveMarginTopPt + pxToPt(rect.top - top));
          break;
        }
        if (headLimitPt <= effectiveMarginTopPt) return null;
      }

      // Below the last citation, the head's own lines still compete with its notes for the page.
      // A head whose notes do not start may still fit by ending earlier on the citing line, so a
      // note-aware split breaks only where a line starts: the citing line moves whole.
      split = this.splitParagraphAtLargestFit(paragraph, (head) => {
        const measured = this.measureElement(head, dims, block.sectionIndex, prefix);
        return effectiveMarginTopPt + measured.heightPt <= headLimitPt
          && (!notesStart || notesStart(this.extractFootnoteRefs(head), measured.heightPt));
      }, {
        lineStartsOnly: notesStart !== undefined,
        preserveWidowLines: (relaxKeepLines || block.relaxedKeepLines === true)
          && paragraph.dataset.widowControl !== "false",
      });
    } finally {
      host?.remove();
    }
    if (split.kind !== "split") return null;

    return [split.head, split.tail].map((element, index) => ({
      ...this.measureElement(element, dims, block.sectionIndex, index === 0 ? prefix : []),
      element,
      keepWithNext: false,
      keepLines: false,
      relaxedKeepLines: relaxKeepLines || block.relaxedKeepLines,
      pageBreakBefore: false,
      isPageBreak: false,
    }));
  }

  /** Keep complete row groups before a citation whose note cannot start on this page. */
  private tryFragmentCitingTable(
    block: MeasuredBlock,
    dims: PageDimensions,
    availableHeightPt: number,
    effectiveMarginTopPt: number,
    notesStart: (footnoteIds: string[], headHeightPt: number) => boolean,
  ): MeasuredBlock[] | null {
    if (block.keepLines || block.keepWithNext || block.pageBreakBefore) return null;
    const split = this.splitTableAtRows(block.element, (head) => {
      const measured = this.measureElement(head, dims, block.sectionIndex);
      return effectiveMarginTopPt + measured.heightPt <= availableHeightPt
        && notesStart(this.extractFootnoteRefs(head), measured.heightPt);
    }, false);
    if (!split) return null;
    split.head.style.setProperty("margin-bottom", "0", "important");
    split.tail.style.setProperty("margin-top", "0", "important");
    return [split.head, split.tail].map(element => ({
      ...this.measureElement(element, dims, block.sectionIndex),
      keepWithNext: false,
      pageBreakBefore: false,
    }));
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
    this.makePresentationCopyInert(clone, false);
    container.appendChild(clone);
  }

  /**
   * Strip every identity from a copy the paginator paints again: element ids and anchor names,
   * editor anchors, note, comment and annotation ids, and internal links, so the copy is not a
   * second bookmark target, comment, note citation or annotation and cannot be edited.
   * `keepSourceAnchors` keeps the canonical source identity, for a copy (a repeated table header
   * row) that the PageMap maps as one more fragment of its source.
   */
  private makePresentationCopyInert(root: HTMLElement, keepSourceAnchors: boolean): void {
    // A comment's reference mark (`[1]`) is the comment's presentation, not document text: the
    // original shows it once. The converter names every mark `comment-ref-<id>`.
    for (const mark of Array.from(root.querySelectorAll('a[id^="comment-ref-"]'))) mark.remove();
    for (const element of [root, ...Array.from(root.querySelectorAll<HTMLElement>("*"))]) {
      element.removeAttribute("id");
      element.removeAttribute("data-anchor");
      element.removeAttribute("data-committed-text");
      // Only the copied content's own identity is kept: an inline comment's story is not copied.
      if (!keepSourceAnchors || this.storyForCanonicalAnchor(element.dataset.sourceAnchorId ?? "") === "comment") {
        element.removeAttribute("data-source-anchor-id");
        element.removeAttribute("data-page-fragment-id");
        element.removeAttribute("data-fragment-index");
      }
      element.removeAttribute("data-footnote-id");
      element.removeAttribute("data-comment-id");
      element.removeAttribute("data-annotation-id");
      element.removeAttribute("name");
      if (element instanceof this.view.HTMLAnchorElement
          && element.getAttribute("href")?.startsWith("#")) {
        element.removeAttribute("href");
      }
      element.setAttribute("contenteditable", "false");
    }
  }

  /** Parses the hidden source notes used to render paginated margin comments. */
  private parseCommentMarginRegistry(): Map<string, HTMLElement> {
    const registry = new Map<string, HTMLElement>();
    const registryEl = marginNoteRegistry(this.stagingElement);
    if (!registryEl) return registry;

    for (const entry of Array.from(
      registryEl.querySelectorAll<HTMLElement>("[data-comment-id]"),
    )) {
      const commentId = entry.dataset.commentId;
      if (!commentId) continue;
      // A reply's note nests inside its thread root (issue #540), so every id in a thread
      // maps to the root's clone — a reply marker selects the whole thread, and the
      // placement loop dedupes by the clone's own id so the thread lands once per page.
      let root = entry;
      for (
        let parent = root.parentElement?.closest<HTMLElement>("[data-comment-id]");
        parent && registryEl.contains(parent);
        parent = root.parentElement?.closest<HTMLElement>("[data-comment-id]")
      ) {
        root = parent;
      }
      registry.set(commentId, root.cloneNode(true) as HTMLElement);
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
   * those shells. Only a number rendered beside the content and the initial
   * HTML/editor identities are omitted; the converter's number travels inside
   * the first paragraph (issue #1130). A hidden sentinel preserves
   * `p:not(:first-of-type)` in document CSS for a page whose first carried
   * element was a later source paragraph; the generated stylesheet has no
   * first-paragraph rule any more.
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
        // A carried paragraph always starts a new line, even when document CSS
        // makes the note's first paragraph inline (the generated stylesheet no
        // longer does, issue #1130).
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
    // Only a number rendered beside the content. The converter puts it inside the first
    // paragraph, as Word does (issue #1130), where it travels with that paragraph's leading
    // fragment in `fittingElements`; copying it here as well would draw it twice.
    const number = footnote.querySelector(":scope > .footnote-number");
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
   * Where a note measurement tree is attached.
   *
   * A note band is reserved from a measurement and then painted; if the two
   * happen under different inherited typography the painted notes overflow the
   * reserve, which is the invisible clipping this engine exists to avoid. The
   * registry lives in the staging tree but a band paints inside the output
   * container, so measure there. The host is never a page box, so a page's
   * `zoom` cannot scale a reserve that is accounted for in unscaled points.
   */
  private noteMeasurementHost(): HTMLElement {
    return this.containerElement.isConnected
      ? this.containerElement
      : this.stagingElement;
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
    // Mirror addPageFootnotes: with no registry and nothing carried, no band is
    // painted, so reserving the separator's height would shrink every page's
    // body for a document whose note references resolve to nothing.
    if (this.footnoteRegistry.size === 0 && !hasContinuation) {
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
    this.noteMeasurementHost().appendChild(measureContainer);
    try {
      return pxToPt(measureContainer.getBoundingClientRect().height);
    } finally {
      measureContainer.remove();
    }
  }

  private blockTable(element: HTMLElement): HTMLTableElement | null {
    return element.localName === "table"
      ? element as HTMLTableElement
      : element.childElementCount === 1 && element.firstElementChild?.localName === "table"
        ? element.firstElementChild as HTMLTableElement
        : null;
  }

  /**
   * Split a body or note table at the largest complete row-group boundary that fits, preserving
   * vertical merges and repeated headers. Null if no group fits, unless `forceFirst` requests
   * the note-only liveness fallback. A wrapper holding only the table is retained.
   */
  private splitTableAtRows(
    element: HTMLElement,
    fits: (head: HTMLElement) => boolean,
    forceFirst: boolean,
    allowNested = false,
  ): { head: HTMLElement; tail: HTMLElement } | null {
    const table = this.blockTable(element);
    if (!table || table.tBodies.length !== 1 || (!allowNested && table.querySelector("table"))
      || Array.from(table.children).some(child => child !== table.tBodies[0] && child.localName !== "colgroup")) {
      return null;
    }
    const body = table.tBodies[0];
    const rows = Array.from(body.rows);
    const allEnds = this.tableRowGroupEnds(rows);
    const headerRows = rows.slice(0, this.leadingHeaderRowCount(rows, allEnds));
    // A cut inside or right after the header run would leave a page of header rows alone.
    const ends = allEnds.filter((end) => end > headerRows.length && end < rows.length);
    if (ends.length === 0) return null;
    const wrapper = element === table ? null : element;
    // The table's header rows open the continuation again, as on a body table's later pages.
    const fragment = (from: number, to: number, first: boolean): HTMLElement =>
      this.createSimpleTableFragment(wrapper, table, body, rows.slice(from, to), first,
        first ? [] : headerRows);
    let low = 0;
    let high = ends.length - 1;
    let best = -1;
    while (low <= high) {
      this.checkpoint();
      const middle = Math.floor((low + high) / 2);
      if (fits(fragment(0, ends[middle], true))) {
        best = middle;
        low = middle + 1;
      } else {
        high = middle - 1;
      }
    }
    if (best < 0 && !forceFirst) return null;
    const cut = ends[Math.max(best, 0)];
    return { head: fragment(0, cut, true), tail: fragment(cut, rows.length, false) };
  }

  /** Split a note's table between rows, or inside the first row when no complete row fits. */
  private splitNoteTable(
    element: HTMLElement,
    fits: (head: HTMLElement) => boolean,
    forceFirst: boolean,
  ): { head: HTMLElement; tail: HTMLElement } | null {
    const wholeRows = this.splitTableAtRows(element, fits, false, true);
    if (wholeRows) return wholeRows;
    const fallback = () => forceFirst ? this.splitTableAtRows(element, fits, true, true) : null;
    const table = this.blockTable(element);
    if (!table || table.tBodies.length !== 1
      || Array.from(table.children).some(child => child !== table.tBodies[0] && child.localName !== "colgroup")) {
      return null;
    }
    const body = table.tBodies[0];
    const rows = Array.from(body.rows);
    const ends = this.tableRowGroupEnds(rows);
    const headers = rows.slice(0, this.leadingHeaderRowCount(rows, ends));
    const row = rows[headers.length];
    // Keep vertical merges, declared row heights and non-top cell alignment intact. These need
    // their own layout rules; ordinary rows can partition each cell in the same measured band.
    if (!row || !ends.includes(headers.length + 1) || row.dataset.docxCantSplit === "true"
      || row.dataset.docxHeaderRow === "true"
      || row.dataset.docxRowHeightRule || row.style.height || row.style.minHeight
      || Array.from(row.cells).some(cell => cell.rowSpan !== 1
        || this.view.getComputedStyle(cell).verticalAlign !== "top")) return fallback();

    // DOCX cells contain blocks. Direct HTML callers can also supply bare text; the block
    // partition below cannot represent it, so retain the whole row rather than dropping it.
    if (Array.from(row.cells).some(cell => Array.from(cell.childNodes).some(node =>
      node.nodeType !== this.view.Node.ELEMENT_NODE
      && (node.nodeType !== this.view.Node.TEXT_NODE || (node.textContent ?? "").trim() !== "")))) {
      return fallback();
    }

    const emptyRow = row.cloneNode(false) as HTMLTableRowElement;
    for (const cell of Array.from(row.cells)) emptyRow.appendChild(cell.cloneNode(false));
    const wrapper = element === table ? null : element;
    const head = this.createSimpleTableFragment(wrapper, table, body, [...headers, emptyRow], true);
    const tail = this.createSimpleTableFragment(wrapper, table, body,
      [emptyRow, ...rows.slice(headers.length + 1)], false, headers);
    const headCells = this.blockTable(head)!.tBodies[0].rows[headers.length].cells;
    const tailCells = this.blockTable(tail)!.tBodies[0].rows[headers.length].cells;
    let hasOverflow = rows.length > headers.length + 1;
    let hasContent = false;
    for (let column = 0; column < row.cells.length; column++) {
      this.checkpoint();
      const children = Array.from(row.cells[column].children) as HTMLElement[];
      const current = headCells[column];
      const remaining = tailCells[column];
      for (let index = 0; index < children.length; index++) {
        this.checkpoint();
        const child = children[index];
        // A keep-with-next chain is one candidate. It may move to the next band, but a prefix
        // must not leave a heading alone at the end of this cell fragment.
        let end = index + 1;
        while (end < children.length && children[end - 1].dataset.keepWithNext === "true") {
          this.checkpoint();
          end++;
        }
        const group = children.slice(index, end).map(node => node.cloneNode(true) as HTMLElement);
        for (const node of group) current.appendChild(node);
        if (fits(head)) {
          hasContent ||= group.some(node => this.hasVisibleText(node.textContent ?? ""));
          index = end - 1;
          continue;
        }
        group.forEach(node => node.remove());
        const fitsChild = (piece: HTMLElement) => {
          current.appendChild(piece);
          try { return fits(head); } finally { piece.remove(); }
        };
        let split: { head: HTMLElement; tail: HTMLElement } | null = null;
        if (end === index + 1) {
          if (this.canRangeFragmentParagraph(child)) {
            const paragraph = this.splitParagraphAtLargestFit(child, fitsChild);
            if (paragraph.kind === "split") split = paragraph;
          } else {
            split = this.splitNoteTable(child, fitsChild, false);
          }
        }
        if (split) {
          current.appendChild(split.head);
          remaining.appendChild(split.tail);
          hasContent = true;
          index++;
        }
        for (const rest of children.slice(index)) remaining.appendChild(rest.cloneNode(true));
        hasOverflow = true;
        break;
      }
    }
    if (!hasContent || !hasOverflow || !fits(head)) return fallback();
    this.reconcileFragmentIdentities(head, tail);
    return { head, tail };
  }

  /** Partition a carried note, advancing even when its first indivisible element is oversized. */
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
    this.noteMeasurementHost().appendChild(measureContainer);

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

        // A kept paragraph that reaches an otherwise empty continuation band has
        // exhausted its fresh-page retry. Word relaxes the keep here, preserving
        // the source metadata and the usual safe-inline fragmentation guard.
        const relaxKeepLines = this.fragmentParagraphs && fitting.length === 0
          && candidate.dataset.keepLines === "true";
        const canSplitCandidate = candidate.tagName === "P"
          && this.canRangeFragmentParagraph(candidate, false, false, relaxKeepLines);
        // Keep the source attached between probes so nested paragraphs retain their actual
        // table-cell styles when the shared fragmenter checks whether a range split is safe.
        const fitsCandidate = (head: HTMLElement) => {
          candidate.replaceWith(head);
          try {
            return pxToPt(measureContainer.getBoundingClientRect().height) <= availableHeightPt;
          } finally {
            head.replaceWith(candidate);
          }
        };
        let split: ParagraphSplitAttempt | null = null;
        if (canSplitCandidate) {
          split = this.splitParagraphAtLargestFit(candidate, fitsCandidate, {
            // Both fallbacks below cut where ordinary line breaking would not,
            // so they are only ever right for an element that owns the whole
            // band: if this page already carries earlier content, a paragraph
            // that will not start here belongs intact in the next note band.
            emergencyGraphemeBreaks: fitting.length === 0,
            // A continuation page owns the full note band. If even one legal
            // grapheme is taller than it, clip only that unit and keep draining.
            forceFirstOnNoFit: fitting.length === 0,
            preserveWidowLines: relaxKeepLines && candidate.dataset.widowControl !== "false",
          });
        }
        const tableSplit = split?.kind === "split" ? null
          : this.splitNoteTable(candidate, fitsCandidate, fitting.length === 0);
        candidate.remove();
        if (split?.kind === "split") {
          if (sourcePosition) {
            split.head.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, sourcePosition);
            split.tail.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, sourcePosition);
          }
          fitting.push(split.head);
          remaining = [split.tail, ...continuation.remainingElements.slice(index + 1)];
        } else if (tableSplit) {
          if (sourcePosition) tableSplit.tail.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, sourcePosition);
          fitting.push(tableSplit.head);
          remaining = [tableSplit.tail, ...continuation.remainingElements.slice(index + 1)];
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
   * Whether every note in `footnoteIds` can begin in a note band of `availableHeightPt` that
   * already holds the page's packed payload: each whole except the last, which may be split
   * provided at least its first line fits. Mirrors the placement loop's decisions without
   * changing any page state, so the paginator can move a citation to the next page instead of
   * separating it from its note (Word starts a note on the page that cites it).
   */
  private notesCanStartOnPage(
    footnoteIds: string[],
    availableHeightPt: number,
    contentWidth: number,
    payload: FootnotePackingContext,
    splitTailTaken: boolean,
    memo?: NoteStartMemo,
  ): boolean {
    // The answer only improves as the band grows, so one known answer settles every band on
    // its side. Paragraph fragmentation asks about the same citations at many head heights.
    const key = footnoteIds.join(" ");
    const startsFrom = memo?.startsFrom.get(key);
    if (startsFrom !== undefined && availableHeightPt >= startsFrom) return true;
    const failsUpTo = memo?.failsUpTo.get(key);
    if (failsUpTo !== undefined && availableHeightPt <= failsUpTo) return false;
    const starts = this.notesStart(footnoteIds, availableHeightPt, contentWidth, payload, splitTailTaken, memo);
    if (memo && starts) memo.startsFrom.set(key, Math.min(startsFrom ?? Infinity, availableHeightPt));
    if (memo && !starts) memo.failsUpTo.set(key, Math.max(failsUpTo ?? -Infinity, availableHeightPt));
    return starts;
  }

  private notesStart(
    footnoteIds: string[],
    availableHeightPt: number,
    contentWidth: number,
    payload: FootnotePackingContext,
    splitTailTaken: boolean,
    memo?: NoteStartMemo,
  ): boolean {
    // The packed height of whole notes does not depend on the band.
    const packedHeight = (ids: string[]) => {
      const key = ids.join(" ");
      let height = memo?.packedHeights.get(key);
      if (height === undefined) {
        height = this.measureFootnotesHeight(
          ids, contentWidth, payload.continuation, payload.partialFootnotes);
        memo?.packedHeights.set(key, height);
      }
      return height;
    };
    const placed = [...payload.footnoteIds];
    let placedHeight = packedHeight(placed);
    for (let index = 0; index < footnoteIds.length; index++) {
      this.checkpoint();
      const footnoteId = footnoteIds[index];
      const footnote = this.footnoteRegistry.get(footnoteId);
      if (!footnote) continue;
      if (splitTailTaken) return false;
      const candidateHeight = packedHeight([...placed, footnoteId]);
      if (candidateHeight + FOOTNOTE_MEASUREMENT_GUARD_PT <= availableHeightPt) {
        placed.push(footnoteId);
        placedHeight = candidateHeight;
        continue;
      }
      if (availableHeightPt - placedHeight <= MIN_FOOTNOTE_START_PT) return false;
      const { fits } = this.splitFootnoteToFit(footnote, availableHeightPt, contentWidth, false, {
        footnoteIds: placed,
        continuation: payload.continuation,
        partialFootnotes: payload.partialFootnotes,
      });
      // A split note takes the page's only continuation slot; a later note cannot start here.
      return fits.length > 0 && index === footnoteIds.length - 1;
    }
    return true;
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
  ): { fits: HTMLElement[]; overflow: HTMLElement[]; keptParagraph?: boolean } {
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
    this.noteMeasurementHost().appendChild(layoutContext);

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
    this.noteMeasurementHost().appendChild(measureContainer);

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
      let keptParagraph = false;

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

        const fitsCandidate = (head: HTMLElement) => {
          measuredContent.appendChild(head);
          try {
            return pxToPt(measureContainer.getBoundingClientRect().height) <= availableHeightPt;
          } finally {
            head.remove();
          }
        };
        keptParagraph = this.fragmentParagraphs && fits.length === 0
          && child.dataset.keepLines === "true"
          && this.canRangeFragmentParagraph(child, true, false, true);
        // Word mode 15 retries the whole kept note on its own page. Older modes
        // relax an oversized keep beside the citation; the caller probes a fresh
        // band before requesting forced progress. Complex inline content keeps
        // its existing conservative fallback in both modes.
        const relaxKeepLines = keptParagraph && forceProgress;
        const split = this.canRangeFragmentParagraph(child, true, false, relaxKeepLines)
          ? this.splitParagraphAtLargestFit(child, fitsCandidate, {
            // Both fallbacks cut where ordinary line breaking would not, so they
            // are only ever right for an element that owns the whole band: with
            // earlier siblings already packed here, a paragraph that will not
            // start belongs intact in the next note band.
            emergencyGraphemeBreaks: fits.length === 0 && forceProgress,
            forceFirstOnNoFit: fits.length === 0 && forceProgress,
            preserveWidowLines: relaxKeepLines && child.dataset.widowControl !== "false",
          })
          : null;
        const tableSplit = split ? null
          : this.splitNoteTable(child, fitsCandidate, fits.length === 0 && forceProgress);
        if (tableSplit) {
          fits.push(tableSplit.head);
          tableSplit.tail.setAttribute(FOOTNOTE_SOURCE_POSITION_ATTR, i === 0 ? "first" : "later");
          overflow = [tableSplit.tail, ...children.slice(i + 1).map((element, offset) =>
            this.cloneFootnoteElementForContinuation(element, i + 1 + offset))];
        } else if (split?.kind === "split") {
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

      return { fits, overflow, keptParagraph };
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
    partialFootnotes?: PartialFootnote[],
    noteBandLimitPt = bands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO,
  ): void {
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if (footnoteIds.length === 0 && !hasContinuation) {
      return;
    }
    if (this.footnoteRegistry.size === 0 && !hasContinuation) {
      return;
    }

    // Calculate max height for footnotes area (content height minus margin for body content)
    const maxFootnoteHeight = Math.min(footnoteHeight, noteBandLimitPt);

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
    displayedPageNumber: number,
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first header if available
    if (pageInSection === 1 && sectionHf.headerFirst) {
      return sectionHf.headerFirst;
    }

    // ECMA-376 §17.10.5 hangs the even story on "even numbered pages": the PAGE NUMBER,
    // which keeps counting across a section boundary unless w:pgNumType restarts it (and a
    // restart moves the parity with it). Word and LibreOffice both render DB001-Sections
    // this way; selecting by position-in-section flips every story of a section that
    // begins on an even page (issue #536).
    if (displayedPageNumber % 2 === 0 && sectionHf.headerEven) {
      return sectionHf.headerEven;
    }

    // Default (odd) pages
    return sectionHf.headerDefault;
  }

  /**
   * The story KIND (`"first"` / `"even"` / `"default"`) a page position selects — the same
   * decision {@link selectHeader} / {@link selectFooter} make, named rather than resolved to
   * an element, so a page can advertise which OOXML story it is showing.
   */
  private storyKindFor(
    which: "header" | "footer",
    sectionIndex: number,
    pageInSection: number,
    displayedPageNumber: number,
  ): "first" | "even" | "default" {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    const first = which === "header" ? sectionHf?.headerFirst : sectionHf?.footerFirst;
    const even = which === "header" ? sectionHf?.headerEven : sectionHf?.footerEven;
    if (pageInSection === 1 && first) return "first";
    if (displayedPageNumber % 2 === 0 && even) return "even";
    return "default";
  }

  /** Select the section's first/odd/even footer from its one-based page position. */
  private selectFooter(
    sectionIndex: number,
    pageInSection: number,
    displayedPageNumber: number,
  ): HTMLElement | undefined {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    if (!sectionHf) return undefined;

    // First page of section uses first footer if available
    if (pageInSection === 1 && sectionHf.footerFirst) {
      return sectionHf.footerFirst;
    }

    if (displayedPageNumber % 2 === 0 && sectionHf.footerEven) {
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
  /** The body height a physical page may fill: its body band less any caller-requested reduction. */
  private bodyBudgetPt(bodyHeightPt: number, physicalPageNumber: number): number {
    return Math.max(0, bodyHeightPt - (this.bodyBudgetReductionsPt[physicalPageNumber] ?? 0));
  }

  private getPageBands(
    dims: PageDimensions,
    sectionIndex: number,
    pageInSection: number,
    displayedPageNumber: number,
  ): PageBands {
    const sectionHf = this.hfRegistry.get(sectionIndex);
    return resolvePageBands(
      dims,
      this.selectStoryHeight(
        sectionHf?.headerFirstHeight,
        sectionHf?.headerEvenHeight,
        sectionHf?.headerDefaultHeight,
        pageInSection,
        displayedPageNumber
      ),
      this.selectStoryHeight(
        sectionHf?.footerFirstHeight,
        sectionHf?.footerEvenHeight,
        sectionHf?.footerDefaultHeight,
        pageInSection,
        displayedPageNumber
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
    displayedPageNumber: number,
  ): number {
    if (pageInSection === 1 && first != null) return first;
    if (displayedPageNumber % 2 === 0 && even != null) return even;
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
    let effectiveContentHeight = this.bodyBudgetPt(this.getPageBands(
      dimensionsFor(), pageSectionIndex, pageInSection(), displayedPageNumber()
    ).bodyHeight, pageNumber);
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
    // Set when the carried continuation does not end on this page: the page is the note's alone.
    let continuationRunsOn = false;
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
    // The note band a page offers: up to the cap, never over the body already placed on it.
    // An indivisible note part taller than that is clipped in the band, leaving the body visible.
    const noteBandLimit = (bodyHeight: number) => Math.max(0, Math.min(bodyHeight * MAX_FOOTNOTE_AREA_RATIO,
      bodyHeight - (effectiveContentHeight - remainingHeight - prevMarginBottomPt)));

    const adoptEmptyPageOwner = (owner: number) => {
      pageSectionIndex = owner;
      const bands = this.getPageBands(
        dimensionsFor(owner),
        owner,
        pageInSection(owner),
        displayedPageNumber(owner),
      );
      effectiveContentHeight = this.bodyBudgetPt(bands.bodyHeight, pageNumber);
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
        displayedPageNumber(),
      );
      const partition = this.splitContinuationForPage(
        currentContinuation!,
        noteBandLimit(pageBands.bodyHeight),
        ownedDimensions.contentWidth,
      );
      currentContinuation = partition.current;
      continuationRunsOn = partition.overflow !== null;
      if (partition.overflow) {
        nextPageContinuation = partition.overflow;
        // The carried note continues again, so it stays last here: notes deferred to this page
        // move on with it (issue #1179).
        deferredFootnoteIds = [...currentFootnoteIds, ...deferredFootnoteIds];
        currentFootnoteIds = [];
      }
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
        ownedDisplayedPageNumber,
      );
      const maxFootnoteHeight = noteBandLimit(pageBands.bodyHeight);
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
            // Fragment probes measure the exact painted payload, as carried
            // continuations do. The whole-note guard above must not also shrink
            // this physical band and discard a line that actually fits.
            const split = source
              ? this.splitFootnoteToFit(
                source,
                maxFootnoteHeight,
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
        maxFootnoteHeight,
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
        displayedPageNumber(),
      );
      effectiveContentHeight = this.bodyBudgetPt(newBands.bodyHeight, pageNumber);
      remainingHeight = effectiveContentHeight;

      prevMarginBottomPt = 0; // Reset margin tracking for new page
      currentFootnoteIds = []; // Reset footnotes for new page
      currentPageHasFootnoteReference = false;
      currentPartialFootnotes = []; // Reset partial footnotes for new page

      // Carry over continuation to next page
      currentContinuation = nextPageContinuation;
      nextPageContinuation = null;
      currentContinuationPartitioned = false;
      continuationRunsOn = false;
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
    // Word keeps a continued note last on its page and its notes in citation order (issue #1179):
    // once a note on this page continues onto the next, or a note cited here has been deferred
    // to it, no later note starts on this page. Its notes start on the next page instead, below
    // the continuation and after the deferred notes.
    const laterNotesClosed = () => nextPageContinuation !== null || deferredFootnoteIds.length > 0;
    // A keep chain gets one fresh-page retry. If its own note still cannot start there,
    // Word can leave the heading and move/split the citing paragraph rather than retrying
    // the same chain on every page beside a carried note.
    const retriedKeepChains = new Set<MeasuredBlock>();
    const releasedKeepPredecessors = new Set<MeasuredBlock>();

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

      // A carried note that does not end on this page fills it: Word reference PDFs show such a
      // page holding the note alone, and body text, with any new citation, resumes on the page
      // where the note ends (issue #1178).
      if (continuationRunsOn && currentContent.length === 0) {
        finishPage(block.sectionIndex);
        i--;
        continue;
      }

      if (currentFootnoteHeight > 0) {
        const ownerDimensions = dimensionsFor();
        const ownerBands = this.getPageBands(
          ownerDimensions,
          pageSectionIndex,
          pageInSection(),
          displayedPageNumber(),
        );
        const ownerMaxFootnoteArea = ownerBands.bodyHeight * MAX_FOOTNOTE_AREA_RATIO;
        if (currentFootnoteHeight > ownerMaxFootnoteArea) {
          // A queue of whole notes deferred to this page that is taller than the
          // band: keep the notes that fit here, defer the rest, and only fall back
          // to a note-owned page when not even one note fits — that case needs the
          // splitting finishPage performs.
          const fittingIds: string[] = [];
          let fittingHeight = 0;
          if ((currentContinuation?.remainingElements.length ?? 0) === 0
              && currentPartialFootnotes.length === 0) {
            for (const footnoteId of currentFootnoteIds) {
              this.checkpoint();
              const candidateHeight = this.measureFootnotesHeight(
                [...fittingIds, footnoteId],
                ownerDimensions.contentWidth,
              );
              if (candidateHeight + FOOTNOTE_MEASUREMENT_GUARD_PT > ownerMaxFootnoteArea) break;
              fittingIds.push(footnoteId);
              fittingHeight = candidateHeight;
            }
          }
          if (fittingIds.length > 0) {
            deferredFootnoteIds = [
              ...currentFootnoteIds.slice(fittingIds.length),
              ...deferredFootnoteIds,
            ];
            currentFootnoteIds = fittingIds;
            currentFootnoteHeight = fittingHeight;
          } else {
            finishPage(block.sectionIndex);
            i--;
            continue;
          }
        }
      }
      const blockDimensions = dimensionsFor(block.sectionIndex);
      const nextBlockPageInSection = pageInSection(block.sectionIndex, pageNumber + 1);
      const freshBlockPageBodyHeight = this.getPageBands(
        blockDimensions,
        block.sectionIndex,
        nextBlockPageInSection,
        displayedPageNumber(block.sectionIndex, nextBlockPageInSection, pageNumber + 1),
      ).bodyHeight;
      const freshBlockHeight = block.heightPt
        + this.effectiveBlockMarginTop(block, 0, true, nextBlockPageInSection);

      // Handle explicit page breaks
      if (block.isPageBreak) {
        finishPage(block.sectionIndex);
        continue;
      }

      // Handle page break before. The block then starts over on the fresh page, so that page's
      // carried note is partitioned before any placement decision reads its height.
      if (block.pageBreakBefore && currentContent.length > 0) {
        finishPage(block.sectionIndex);
        i--;
        continue;
      }

      // Keep the chain's headings with the following paragraph or a valid first fragment.
      // If neither fits here, move a chain that fits a fresh page as a unit. Oversized
      // indivisible chains retain the established greedy/overflow behavior.
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
      if (!pageHasOccupiedSpace && block.keepWithNext && startsKeepChain) {
        const chain = this.getKeepWithNextChain(blocks, i);
        const last = chain[chain.length - 1];
        if (chain.length > 1 && this.fragmentParagraphs && last.keepLines
          && this.isWordParagraphElement(last.element)
          && this.collectNewFootnoteIds(chain, currentFootnoteIds).length > 0) {
          retriedKeepChains.add(block);
          releasedKeepPredecessors.add(last);
        }
      }
      if (pageHasOccupiedSpace && block.keepWithNext && startsKeepChain && !retriedKeepChains.has(block)) {
        const keepChain = this.getKeepWithNextChain(blocks, i);
        if (keepChain.length > 1) {
          const newChainFootnoteIds = this.collectNewFootnoteIds(
            keepChain,
            currentFootnoteIds
          );
          // A note need only start beside the line that cites it; the rest may run on. Whether
          // the chain's notes can start is asked the way a citing paragraph asks it
          // (notesCanStartOnPage), here beside this page's notes and on a fresh page alike.
          const chainNotesStart = (bodyPt: number, bandPt: number, payload: FootnotePackingContext,
            closed: boolean, ids = newChainFootnoteIds) =>
            ids.length === 0 || this.footnoteRegistry.size === 0
            || this.notesCanStartOnPage(ids, bandPt - bodyPt,
              blockDimensions.contentWidth, payload, closed);
          const currentChainHeight =
            this.measureKeepWithNextChainBodyHeight(
              keepChain,
              prevMarginBottomPt,
              currentContent.length === 0,
              pageInSection(block.sectionIndex)
            );
          const currentAvailableHeight = remainingHeight - currentFootnoteHeight;
          // Once a note cited here has ended, Word may defer a later whole note while
          // keeping its short, kept citing chain here. A carried tail alone, or a note
          // already running on from this page, does not permit that deferral.
          const canDeferChainNote = this.fragmentParagraphs
            && keepChain[keepChain.length - 1].keepLines
            && this.isWordParagraphElement(keepChain[keepChain.length - 1].element)
            && newChainFootnoteIds.length > 0 && currentPageHasFootnoteReference
            && currentFootnoteIds.length > 0 && !laterNotesClosed();
          const chainFitsHere = currentChainHeight <= currentAvailableHeight
            && (canDeferChainNote || chainNotesStart(currentChainHeight, remainingHeight, {
              footnoteIds: currentFootnoteIds,
              continuation: currentContinuation,
              partialFootnotes: currentPartialFootnotes,
            }, laterNotesClosed()));

          if (!chainFitsHere) {
            const last = keepChain[keepChain.length - 1];
            const canRelaxChain = this.fragmentParagraphs && last.keepLines
              && this.isWordParagraphElement(last.element) && newChainFootnoteIds.length > 0;
            const precedingNoteIds = this.collectNewFootnoteIds(keepChain.slice(0, -1), currentFootnoteIds);
            const fragmentLast = (chainHeight: number, available: number, band: number,
              payload: FootnotePackingContext, closed: boolean, preceding: readonly HTMLElement[]) => {
              const beforeLastHeight = chainHeight - last.heightPt;
              return this.tryFragmentParagraph(last, blockDimensions, available - beforeLastHeight, 0,
                (ids, headHeight) => chainNotesStart(beforeLastHeight + headHeight, band, payload, closed,
                  [...new Set([...precedingNoteIds, ...ids])].filter(id => !currentFootnoteIds.includes(id))),
                [...preceding, ...keepChain.slice(0, -1).map(item => item.element)]);
            };
            const fragments = fragmentLast(currentChainHeight, currentAvailableHeight, remainingHeight, {
              footnoteIds: currentFootnoteIds,
              continuation: currentContinuation,
              partialFootnotes: currentPartialFootnotes,
            }, laterNotesClosed(), currentContent);
            if (fragments) {
              blocks.splice(i + keepChain.length - 1, 1, ...fragments);
              i--;
              continue;
            }

            const nextPageBands = this.getPageBands(
              dimensionsFor(block.sectionIndex),
              block.sectionIndex,
              pageInSection(block.sectionIndex, pageNumber + 1),
              displayedPageNumber(
                block.sectionIndex,
                pageInSection(block.sectionIndex, pageNumber + 1),
                pageNumber + 1,
              ),
            );
            const freshChainBodyHeight = this.measureKeepWithNextChainBodyHeight(
              keepChain,
              0,
              true,
              pageInSection(block.sectionIndex, pageNumber + 1)
            );
            // Whether the chain can fit a page at all, its notes started beside it. A carried note on
            // the next page does not count against it: Word reference PDFs show that note either
            // leaving the chain its room or keeping the page, the chain opening the page after
            // (issue #1180). Counting the note refused the move, and the chain's first block was
            // left alone at this page's foot.
            if (
              (freshChainBodyHeight <= nextPageBands.bodyHeight
                && (canRelaxChain || chainNotesStart(freshChainBodyHeight, nextPageBands.bodyHeight,
                  { footnoteIds: [], continuation: null, partialFootnotes: [] }, false)))
              || fragmentLast(freshChainBodyHeight, nextPageBands.bodyHeight, nextPageBands.bodyHeight,
                { footnoteIds: [], continuation: null, partialFootnotes: [] }, false, []) !== null
            ) {
              // Start the chain over on the fresh page, so that page's carried note is
              // partitioned first and the chain is judged again beside it.
              if (canRelaxChain) {
                retriedKeepChains.add(block);
                releasedKeepPredecessors.add(last);
              }
              finishPage(block.sectionIndex);
              i--;
              continue;
            }
          }
        }
      }

      // After a newly split note, body text resumes on the page where that note ends.
      // Keep-chain handling above records its retry before this page turn.
      if (this.fragmentParagraphs && nextPageContinuation && currentContent.length > 0) {
        finishPage(block.sectionIndex);
        i--;
        continue;
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
      const columnFragments = this.reflowSequentialColumns(
        block,
        blockDimensions,
        remainingHeight - currentFootnoteHeight - additionalFootnoteHeight - effectiveMarginTop,
      );
      if (columnFragments) {
        blocks.splice(i, 1, ...columnFragments);
        i--;
        continue;
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

      // The note band left beside `bodyPt` more of body on this page, and whether a paragraph
      // head of that height can start every note it cites in it. Word begins a note on the page
      // that cites it, so a fragment ends before a citation whose note cannot start here.
      const notePayload = {
        footnoteIds: currentFootnoteIds,
        continuation: currentContinuation,
        partialFootnotes: currentPartialFootnotes,
      };
      const noteBandWith = (bodyPt: number) => Math.min(
        maxFootnoteArea, effectiveContentHeight - bodyContentUsed - bodyPt);
      const noteStartMemo: NoteStartMemo = { packedHeights: new Map(), startsFrom: new Map(), failsUpTo: new Map() };
      const headNotesStart = (footnoteIds: string[], headHeightPt: number) => {
        const headIds = footnoteIds.filter((id) => !currentFootnoteIds.includes(id));
        return headIds.length === 0 || this.notesCanStartOnPage(headIds,
          noteBandWith(effectiveMarginTop + headHeightPt), blockDimensions.contentWidth,
          notePayload, laterNotesClosed(), noteStartMemo);
      };

      // A paragraph that cannot fit as a whole may still have a simple text-only
      // prefix that fits this page. Fragment before the ordinary next-page or
      // oversized fallback so the cloned head participates in the same margin
      // and footnote accounting as every other block. Only the paragraph's own lines decide
      // this: notes too tall for the page are split or moved by the note placement below.
      if (effectiveMarginTop + block.heightPt > effectiveRemainingHeight) {
        const paragraphFragments = this.tryFragmentParagraph(
          block,
          blockDimensions,
          effectiveRemainingHeight,
          effectiveMarginTop,
          headNotesStart,
          currentContent,
          block.keepLines && isFirstOnPage && block.heightPt > freshBlockPageBodyHeight
            && allBlockFootnoteIds.some(id => this.footnoteRegistry.has(id)),
        );
        if (paragraphFragments) {
          blocks.splice(i, 1, ...paragraphFragments);
          i--;
          continue;
        }
      }

      // Check if block fits on current page (including its footnotes). The note cap judges the
      // notes this block adds; notes already admitted to the page do not refuse a block that
      // cites none (a carried continuation can fill the capped area exactly).
      if (
        blockSpace <= effectiveRemainingHeight
        && (newFootnoteIds.length === 0
          || (!laterNotesClosed()
            && combinedFootnoteHeight + FOOTNOTE_MEASUREMENT_GUARD_PT <= maxFootnoteArea))
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
      } else if (freshBlockHeight <= freshBlockPageBodyHeight) {
        // Block doesn't fit with current allocation - try expanding footnote area
        const blockSpaceWithoutFootnotes = effectiveMarginTop + block.heightPt;

        if (newFootnoteIds.length > 0 && blockSpaceWithoutFootnotes <= effectiveRemainingHeight) {
          // Keep a paragraph's earlier lines or a table's earlier row groups here, moving the
          // citing line/group with its note. Paragraph splitting remains opt-in for the editor;
          // tables retain whole rows in either mode. A whole table moves only if its notes can
          // actually start beside it on a fresh page.
          const table = this.blockTable(block.element);
          const turnHelps = noteBandWith(blockSpaceWithoutFootnotes) - currentFootnoteHeight
            < freshBlockPageBodyHeight * MAX_FOOTNOTE_AREA_RATIO;
          if ((this.fragmentParagraphs || table) && turnHelps && !this.notesCanStartOnPage(newFootnoteIds,
            noteBandWith(blockSpaceWithoutFootnotes), blockDimensions.contentWidth, notePayload,
            laterNotesClosed(), noteStartMemo)) {
            const fragments = this.tryFragmentParagraph(
              block,
              blockDimensions,
              effectiveRemainingHeight,
              effectiveMarginTop,
              headNotesStart,
              currentContent,
              block.keepLines && isFirstOnPage && currentFootnoteHeight === 0 && !currentContinuation,
            ) ?? this.tryFragmentCitingTable(block, blockDimensions, effectiveRemainingHeight,
              effectiveMarginTop, headNotesStart);
            if (fragments) {
              blocks.splice(i, 1, ...fragments);
              i--;
              continue;
            }
            const canMoveWhole = this.isWordParagraphElement(block.element) || (table
              && this.notesCanStartOnPage(newFootnoteIds, freshBlockPageBodyHeight - freshBlockHeight,
              blockDimensions.contentWidth, { footnoteIds: [], continuation: null, partialFootnotes: [] }, false));
            // A preceding keep-with-next block belongs with the citation; the chain path owns it.
            if (currentContent.length > 0 && canMoveWhole
                && (!blocks[i - 1]?.keepWithNext || releasedKeepPredecessors.has(block))) {
              finishPage(block.sectionIndex);
              i--;
              continue;
            }
          }
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
          // Deferring a note costs a page turn, and it only buys anything if the
          // next page's band is larger than what this page can already offer.
          // When this page already offers the maximum band, a note that cannot
          // start here cannot start there either: keep the established clipped
          // fallback on the citing page rather than evacuating it.
          const nextPageMaxFootnoteArea =
            freshBlockPageBodyHeight * MAX_FOOTNOTE_AREA_RATIO;
          const deferralCannotHelp = availableForFootnotes >= nextPageMaxFootnoteArea;

          // Try to fit as much of each new footnote as possible in expanded area
          for (let noteIndex = 0; noteIndex < newFootnoteIds.length; noteIndex++) {
            this.checkpoint();
            const footnoteId = newFootnoteIds[noteIndex];
            const footnote = this.footnoteRegistry.get(footnoteId);
            if (!footnote) continue;

            // One page can carry only one split tail, and it stays last. Once that slot is
            // used, or an earlier note was deferred, keep every later note in source order
            // on the deferral queue.
            if (laterNotesClosed()) {
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
              if (spaceLeftForFootnotes > MIN_FOOTNOTE_START_PT) {
                const payload = {
                  footnoteIds: currentFootnoteIds,
                  continuation: currentContinuation,
                  partialFootnotes: currentPartialFootnotes,
                };
                let { fits, overflow, keptParagraph } = this.splitFootnoteToFit(
                  footnote,
                  availableForFootnotes,
                  blockDimensions.contentWidth,
                  deferralCannotHelp,
                  payload,
                );
                // Editor mode defers an explicitly kept first paragraph to its own page.
                // Other indivisible content retains its existing progress fallback here.
                const firstPart = footnote.querySelector<HTMLElement>(":scope > .footnote-content")?.firstElementChild;
                const deferKeptEditorNote = !this.fragmentParagraphs
                  && firstPart instanceof this.view.HTMLElement && firstPart.dataset.keepLines === "true";
                if (!deferKeptEditorNote
                  && !(keptParagraph && Number(this.stagingElement.dataset.docxCompatibilityMode) >= 15)
                  && fits.length === 0 && !deferralCannotHelp
                  && this.splitFootnoteToFit(footnote, keptParagraph
                    ? nextPageMaxFootnoteArea - freshBlockHeight
                    : nextPageMaxFootnoteArea, blockDimensions.contentWidth).fits.length === 0) {
                  ({ fits, overflow } = this.splitFootnoteToFit(
                    footnote, availableForFootnotes, blockDimensions.contentWidth, true, payload));
                }

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
          // Let normal fitting place the first fragment beside preceding content when it fits.
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
        const oversized = this.cloneBlockForPage(
          block,
          true,
          pageInSection(block.sectionIndex),
        );
        // A known exact row clips at the paper edge; other fallback overflow stays guarded.
        oversized.setAttribute(OVERSIZED_BLOCK_ATTRIBUTE, "true");
        currentContent.push(oversized);
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

  /** Keep repeated note identities inert; resolve their return links after every page exists. */
  private makeClonedMarginCommentInert(root: HTMLElement): void {
    this.makeClonedStoryInert(root);
    const nodes = [root, ...Array.from(root.querySelectorAll<HTMLElement>("*"))];
    for (const element of nodes) {
      element.removeAttribute("id");
      const href = element.getAttribute("href");
      if (element.localName === "a" && href?.startsWith("#")) {
        if (element instanceof this.view.HTMLAnchorElement && /^#comment-ref--?[0-9]+$/.test(href)
            && Array.from(element.classList).some(name => name.endsWith("margin-backref"))
            && element.getAttribute("aria-disabled") !== "true") {
          this.commentReturnLinks.push({ link: element, href,
            ariaDisabled: element.getAttribute("aria-disabled"),
            tabIndex: element.getAttribute("tabindex") });
        }
        element.removeAttribute("href");
        element.setAttribute("aria-disabled", "true");
        element.tabIndex = -1;
      }
    }
  }

  /** A return link may target only one actual rendered reference, including on a later page. */
  private activateCommentReturnLinks(pages: PageInfo[]): void {
    if (this.commentReturnLinks.length === 0) return;
    const wanted = new Set(this.commentReturnLinks.map(({ href }) => href.slice(1)));
    const targets = new Map<string, HTMLElement | null>();
    for (const element of Array.from(this.document.querySelectorAll<HTMLElement>('[id^="comment-ref-"]'))) {
      this.checkpoint();
      if (!wanted.has(element.id)) continue;
      targets.set(element.id, targets.has(element.id) ? null : element);
    }
    const pageElements = new Set<Element>(pages.map(page => page.element));
    const eligible = new Set<string>();
    for (const [id, target] of targets) {
      this.checkpoint();
      if (!target) continue;
      let page: Element | null = target;
      while (page && !pageElements.has(page)) page = page.parentElement;
      if (!page || target.closest(`.${this.cssPrefix}header, .${this.cssPrefix}footer, .${this.cssPrefix}comment-margin`)) continue;
      const sourceId = target.closest<HTMLElement>('[data-source-anchor-id]')?.dataset.sourceAnchorId;
      if (sourceId) {
        const story = this.storyForCanonicalAnchor(sourceId);
        if (story !== "body" && story !== "footnote" && story !== "endnote") continue;
      } else if (!target.closest(`.${this.cssPrefix}content, .${this.cssPrefix}footnotes`)) continue;
      const pageElement = page as HTMLElement;
      // A painted citation may be excluded from PageMap's authored source inventory.
      if (this.view.getComputedStyle(target).visibility !== "visible") continue;
      const { visible } = measureRenderedSource(this.view, target, pageElement);
      if (visible.right <= visible.left || visible.bottom <= visible.top) continue;
      let transparent = false;
      for (let ancestor: HTMLElement | null = target; ancestor; ancestor = ancestor.parentElement) {
        if (Number(this.view.getComputedStyle(ancestor).opacity) === 0) { transparent = true; break; }
        if (ancestor === pageElement) break;
      }
      if (!transparent) eligible.add(id);
    }
    for (const { link, href, ariaDisabled, tabIndex } of this.commentReturnLinks) {
      this.checkpoint();
      if (!eligible.has(href.slice(1))) continue;
      link.setAttribute("href", href);
      if (ariaDisabled === null) link.removeAttribute("aria-disabled");
      else link.setAttribute("aria-disabled", ariaDisabled);
      if (tabIndex === null) link.removeAttribute("tabindex");
      else link.setAttribute("tabindex", tabIndex);
    }
    this.commentReturnLinks.length = 0;
  }

  /**
   * Resolves floating DrawingML objects after their anchor paragraphs have landed on a page.
   *
   * The converter deliberately carries the OOXML bases and offsets as data instead of flattening
   * them into CSS: page/margin/column bases belong to the page, while paragraph/line/character
   * bases belong to the laid-out anchor. Once both coordinate systems exist, promote the object
   * from the clipped text column into the page box and position it in one shared point space.
   */
  /** Resolve legacy paragraph frames after their ordinary anchor paragraphs have flowed. */
  private positionTextFrames(pageBox: HTMLElement, dims: PageDimensions, pageNumber: number): HTMLElement[] {
    const frames = Array.from(pageBox.querySelectorAll<HTMLElement>('[data-docx-text-frame="true"]'));
    const pageRect = pageBox.getBoundingClientRect();
    const pixelsPerPoint = pageRect.width / dims.pageWidth;
    if (!(pixelsPerPoint > 0)) return [];
    for (const frame of frames) {
      const placeholder = frame.closest<HTMLElement>('[data-docx-frame-container="true"]');
      if (!placeholder) continue;
      let following = placeholder.nextElementSibling;
      while (following instanceof this.view.HTMLElement && following.dataset.docxFrameContainer === 'true') {
        following = following.nextElementSibling;
      }
      const paragraph = following instanceof this.view.HTMLElement
        ? (following.matches('p,h1,h2,h3,h4,h5,h6') ? following : following.querySelector<HTMLElement>('p,h1,h2,h3,h4,h5,h6'))
        : null;
      const anchorRect = (paragraph ?? placeholder).getBoundingClientRect();
      const rect = frame.getBoundingClientRect();
      const computed = this.view.getComputedStyle(frame);
      const context = {
        staticLeft: (anchorRect.left - pageRect.left) / pixelsPerPoint,
        staticTop: (anchorRect.top - pageRect.top) / pixelsPerPoint,
        paragraphLeft: (anchorRect.left - pageRect.left) / pixelsPerPoint,
        paragraphTop: (anchorRect.top - pageRect.top) / pixelsPerPoint,
        paragraphWidth: anchorRect.width / pixelsPerPoint,
        paragraphHeight: anchorRect.height / pixelsPerPoint,
        lineHeight: anchorRect.height / pixelsPerPoint,
      };
      const horizontal = this.horizontalAnchorReference(
        frame.dataset.docxFrameHAnchor === 'text' ? 'column' : frame.dataset.docxFrameHAnchor ?? 'column',
        dims, pageNumber, context,
      );
      const vertical = this.verticalAnchorReference(
        frame.dataset.docxFrameVAnchor === 'text' ? 'paragraph' : frame.dataset.docxFrameVAnchor ?? 'paragraph',
        dims, pageNumber, context,
      );
      const xAlign = frame.getAttribute('data-docx-frame-xalign');
      const yAlign = frame.getAttribute('data-docx-frame-yalign');
      const left = this.resolveAnchorAxis(horizontal, rect.width / pixelsPerPoint, xAlign,
        xAlign ? undefined : pxToPt(parseFloat(computed.left) || 0), pageNumber);
      const top = this.resolveAnchorAxis(vertical, rect.height / pixelsPerPoint, yAlign,
        yAlign ? undefined : pxToPt(parseFloat(computed.top) || 0), pageNumber);
      // Escape body/story clipping while keeping inherited presentation and canonical source ids.
      for (const property of ['font-family', 'font-size', 'font-style', 'font-weight', 'line-height',
        'color', 'direction', 'text-align', 'letter-spacing', 'word-spacing', 'white-space']) {
        frame.style.setProperty(property, computed.getPropertyValue(property));
      }
      if (frame.closest('td,th')) frame.dataset.docxFrameInTableCell = 'true';
      pageBox.appendChild(frame);
      frame.style.left = `${left}pt`;
      frame.style.top = `${top}pt`;
      frame.style.margin = '0';
    }
    return frames;
  }

  private sizeDrawingAnchor(anchor: HTMLElement, dims: PageDimensions, pageNumber: number,
    context: { staticLeft: number; staticTop: number; lineHeight: number; paragraphLeft: number;
      paragraphTop: number; paragraphWidth: number; paragraphHeight: number }): void {
    const widthReference = this.horizontalAnchorReference(
      anchor.getAttribute('data-docx-anchor-width-relative') ?? 'margin',
      dims,
      pageNumber,
      context,
    );
    const widthPercent = this.anchorNumber(anchor, 'width-percent');
    // Word supports strictly positive relative dimensions (MS-ODRAWXML 2.18.3.1/2).
    // Zero is schema-valid but must not collapse a visible object with a stored extent.
    if (widthPercent !== undefined && widthPercent > 0) {
      anchor.style.width = `${widthReference.size * widthPercent / 100}pt`;
    }

    const heightReference = this.verticalAnchorReference(
      anchor.getAttribute('data-docx-anchor-height-relative') ?? 'margin',
      dims,
      pageNumber,
      context,
    );
    const heightPercent = this.anchorNumber(anchor, 'height-percent');
    if (heightPercent !== undefined && heightPercent > 0 && anchor.dataset.docxAnchorAutofit !== 'true') {
      anchor.style.height = `${heightReference.size * heightPercent / 100}pt`;
    }
  }

  private prepareShapeClearance(section: HTMLElement, dims: PageDimensions): void {
    const context = {staticLeft: 0, staticTop: 0, lineHeight: 0, paragraphLeft: 0,
      paragraphTop: 0, paragraphWidth: dims.contentWidth, paragraphHeight: 0};
    const bottoms = new Map<HTMLElement, number>();
    for (const anchor of Array.from(section.querySelectorAll<HTMLElement>('[data-docx-shape-anchor="true"]'))) {
      // Page/margin percentages are independent of the eventual page number. Page-dependent
      // margin variants retain unsupported metadata from conversion until they can be laid out.
      this.sizeDrawingAnchor(anchor, dims, 1, context);
      const width = parseFloat(anchor.style.width) || this.anchorNumber(anchor, 'extent-width') || 0;
      const height = parseFloat(anchor.style.height) || this.anchorNumber(anchor, 'extent-height') || 0;
      const exclusion = anchor.parentElement?.closest<HTMLElement>('[data-docx-shape-exclusion]');
      if (exclusion) { exclusion.style.width = `${width}pt`; exclusion.style.height = `${height}pt`; }
      if (anchor.dataset.docxAnchorWrap !== 'wrapTopAndBottom') continue;
      const paragraph = anchor.closest<HTMLElement>('p,h1,h2,h3,h4,h5,h6');
      if (!paragraph) continue;
      const offset = anchor.dataset.docxAnchorVRelative === 'paragraph' ? this.anchorNumber(anchor, 'v-offset') ?? 0 : 0;
      bottoms.set(paragraph, Math.max(bottoms.get(paragraph) ?? 0,
        height + offset + (this.anchorNumber(anchor, 'wrap-bottom') ?? 0)));
    }
    for (const [paragraph, height] of bottoms) {
      const spacer = Array.from(paragraph.children).find(e => e.hasAttribute('data-docx-shape-clearance')) as HTMLElement | undefined;
      if (spacer) spacer.style.height = `${Math.max(0, height)}pt`;
    }
  }

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

      this.sizeDrawingAnchor(anchor, dims, pageNumber, context);

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

      // Leaving the paragraph must not orphan it: a paragraph whose only content was this drawing
      // renders as the drawing, and measureRenderedSource finds it through this link (issue #849).
      const hostElement = anchor.parentElement?.closest<HTMLElement>('[data-source-anchor-id]');
      const hostAnchorId = hostElement?.dataset.sourceAnchorId;
      if (hostAnchorId) anchor.setAttribute(DRAWING_HOST_ATTRIBUTE, hostAnchorId);
      pageBox.appendChild(anchor);
      // A host left with nothing but its paragraph-mark line (issue #880) renders as that line plus
      // the drawing; one with text of its own keeps measuring as its own box.
      if (hostElement && rendersOnlyItsLine(hostElement))
        hostElement.setAttribute(HOSTS_PROMOTED_DRAWING_ATTRIBUTE, 'true');
      anchor.style.position = 'absolute';
      anchor.style.left = `${left}pt`;
      anchor.style.top = `${top}pt`;
      anchor.style.margin = '0';
    }

    // relativeHeight is an unsigned 32-bit value; using it directly as a CSS z-index would
    // clamp distinct values above 2^31. Dense ranks also keep every behindDoc object below
    // the text, independently of its relativeHeight and XML order.
    const paintOrder = [...anchors].sort((a, b) =>
      (this.anchorNumber(a, 'order') ?? 0) - (this.anchorNumber(b, 'order') ?? 0));
    paintOrder.forEach((anchor, index) => {
      anchor.style.zIndex = String(anchor.dataset.docxAnchorBehind === 'true'
        ? index - paintOrder.length : index + 1);
    });
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
    noteBandLimitPt?: number,
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
    // Standalone export removes contain/will-change. Keep negative drawing layers above
    // this page's paper background after those performance hints are stripped.
    pageBox.style.isolation = "isolate";
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
    const bands = this.getPageBands(dims, sectionIndex, pageInSection, displayedPageNumber);

    // Add header if available for this section/page
    const headerSource = isSectionFiller
      ? undefined
      : this.selectHeader(sectionIndex, pageInSection, displayedPageNumber);

    if (headerSource) {
      const headerDiv = this.document.createElement("div");
      headerDiv.className = `${this.cssPrefix}header`;
      // Which story this page shows, so an editor can map the page back to the OOXML part it
      // edits in place (a first-page header is a different part from the default one).
      headerDiv.dataset.hfType = this.storyKindFor(
        "header", sectionIndex, pageInSection, displayedPageNumber,
      );
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
    // Clip only vertically: Word lets over-wide content (e.g. a table pulled left by a negative
    // w:tblInd, issue #827) run into the side margins, clipping it only at the paper edge.
    // `clip` rather than `hidden`, because `hidden` would force the x axis to `auto`.
    contentArea.style.overflowX = "visible";
    contentArea.style.overflowY = "clip";

    // Add content
    for (const el of content) {
      contentArea.appendChild(el);
    }

    pageBox.appendChild(contentArea);

    // Add footnotes if any references appear on this page (or continuation from previous)
    const hasContinuation = continuation && continuation.remainingElements.length > 0;
    if (!isSectionFiller && (footnoteIds.length > 0 || hasContinuation)) {
      this.addPageFootnotes(pageBox, footnoteIds, dims, bands, footnoteHeight, continuation, partialFootnotes,
        noteBandLimitPt);

    }

    // Materialize margin comments after footnotes exist so markers inside a
    // note participate in the same page-owned comment story as body markers. The page's header
    // and footer count too: a running story's comment prints on every page that story prints on,
    // as its balloon does (issue #1310). The footer is placed below, so read it from its source.
    const footerSource = isSectionFiller
      ? undefined
      : this.selectFooter(sectionIndex, pageInSection, displayedPageNumber);
    const pageCommentIds: string[] = [];
    for (const marker of [pageBox, footerSource].flatMap((story) =>
      Array.from(story?.querySelectorAll<HTMLElement>("[data-comment-id]") ?? []))) {
      if (marker.closest(`.${this.cssPrefix}comment-margin`)) continue;
      const id = marker.dataset.commentId;
      if (id && this.commentMarginRegistry.has(id) && !pageCommentIds.includes(id)) {
        pageCommentIds.push(id);
      }
    }
    let outsideCommentColumn: HTMLElement | undefined;
    if (this.commentColumnPlacement !== "balloons" && pageCommentIds.length > 0) {
      const marginColumn = this.document.createElement("aside");
      marginColumn.className = `${this.cssPrefix}comment-margin`;
      marginColumn.style.position = "absolute";
      marginColumn.style.top = `${contentAreaTop}pt`;
      marginColumn.style.left = `${dims.marginLeft + dims.contentWidth + 3}pt`;
      marginColumn.style.width = `${Math.max(12, dims.marginRight - 6)}pt`;
      marginColumn.style.maxHeight = `${contentAreaHeight}pt`;
      marginColumn.style.overflowY = "auto";
      marginColumn.style.overflowX = "hidden";
      marginColumn.tabIndex = 0;
      marginColumn.setAttribute("aria-label", `Comments for page ${displayedPageNumber}`);
      marginColumn.dataset.pageNumber = String(pageNumber);
      marginColumn.style.boxSizing = "border-box";
      const placedThreadRoots = new Set<string>();
      for (const id of pageCommentIds) {
        const source = this.commentMarginRegistry.get(id);
        if (source) {
          // Every id in a nested thread maps to the same root note; place it once even
          // when the root's and a reply's markers share the page.
          const rootId = source.dataset.commentId ?? id;
          if (placedThreadRoots.has(rootId)) continue;
          placedThreadRoots.add(rootId);
          const clone = source.cloneNode(true) as HTMLElement;
          this.makeClonedMarginCommentInert(clone);
          marginColumn.appendChild(clone);
        }
      }
      if (this.commentColumnPlacement === "outside") {
        marginColumn.style.top = `${contentAreaTop * this.scale}pt`;
        marginColumn.style.left = "0";
        marginColumn.style.width = `${this.commentColumnWidthPt}pt`;
        marginColumn.style.transform = `scale(${this.scale})`;
        marginColumn.style.transformOrigin = "top left";
        marginColumn.dataset.pageMapExclude = "true";
        outsideCommentColumn = marginColumn;
      } else pageBox.appendChild(marginColumn);
    }

    // Add footer if available for this section/page
    if (footerSource) {
      const footerDiv = this.document.createElement("div");
      footerDiv.className = `${this.cssPrefix}footer`;
      footerDiv.dataset.hfType = this.storyKindFor(
        "footer", sectionIndex, pageInSection, displayedPageNumber,
      );
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
    if (this.commentColumnPlacement === "outside" && this.commentMarginRegistry.size > 0) {
      const row = this.document.createElement("div");
      row.dataset.docxPageCommentsRow = "true";
      row.style.display = "flex";
      row.style.alignItems = "flex-start";
      row.style.width = "max-content";
      // Auto margins center a row when it fits, and become zero when a narrow viewer must pan.
      row.style.alignSelf = "flex-start";
      row.style.marginLeft = "auto";
      row.style.marginRight = "auto";
      row.style.gap = `${this.commentColumnGapPt * this.scale}pt`;
      pageBox.style.flexShrink = "0";
      const panel = this.document.createElement("div");
      panel.dataset.docxCommentPanel = "true";
      panel.style.position = "relative";
      panel.style.flexShrink = "0";
      panel.style.width = `${this.commentColumnWidthPt * this.scale}pt`;
      panel.style.height = `${dims.pageHeight * this.scale}pt`;
      if (outsideCommentColumn) panel.appendChild(outsideCommentColumn);
      else panel.setAttribute("aria-hidden", "true");
      row.append(pageBox, panel);
      this.containerElement.appendChild(row);
      if (!this.document.head.querySelector("style[data-docx-comment-column-print]")) {
        const printStyle = this.document.createElement("style");
        printStyle.dataset.docxCommentColumnPrint = "true";
        printStyle.textContent = "@media print { [data-docx-comment-panel] { display:none!important; } " +
          "[data-docx-page-comments-row] { display:block!important; width:auto!important; } }";
        this.document.head.appendChild(printStyle);
      }
    } else this.containerElement.appendChild(pageBox);

    // Floating objects need the final page and anchor-paragraph boxes. Resolve them only after
    // append, when getBoundingClientRect() is meaningful, and lift them out of content clipping.
    this.positionTextFrames(pageBox, dims, pageNumber);
    // One page-wide pass includes running stories and compares all drawings in the same
    // stacking order, including those nested in a promoted paragraph frame.
    this.positionDrawingAnchors(pageBox, pageBox, dims, pageNumber);

    // Body and notes must occupy DISJOINT bands. The note block is absolutely positioned against
    // the page bottom and grows upward, while the content area spans the whole text height — so
    // nothing but the layout loop's arithmetic keeps them apart, and any disagreement between the
    // height the loop reserved and the height the notes actually render at draws body text and note
    // text on top of each other (observed on a 94-footnote document: ~134pt of superimposed,
    // illegible glyphs). Shrinking the content area to the space the notes actually left removes
    // the failure mode by construction — the worst case becomes a clean clip by the content area's
    // vertical overflow clip, which is obvious and recoverable rather than silent corruption.
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

    // Word lets an oversized exact row paint through the bottom margin and even behind a
    // footer. The page box still clips at the paper edge; notes keep their disjoint body band.
    if ((!notesEl || notesEl.getBoundingClientRect().height === 0)
      && Array.from(contentArea.querySelectorAll(`[${OVERSIZED_BLOCK_ATTRIBUTE}]`))
        .some(block => findOversizedExactRow(block))) contentArea.style.overflowY = "visible";

    // Paint offsets must follow range fragmentation: positioned runs are deliberately excluded
    // from that structural algorithm, even though these offsets leave every line box unchanged.
    alignExactLineBaselines(contentArea);
    // Note paragraphs keep their own exact spacing too (issue #1120); the band is bottom-anchored,
    // so only glyphs move within their unchanged line boxes.
    if (notesEl) alignExactLineBaselines(notesEl);
    const paper = pageBox.getBoundingClientRect();
    const header = pageBox.querySelector<HTMLElement>(`.${this.cssPrefix}header`);
    const footer = pageBox.querySelector<HTMLElement>(`.${this.cssPrefix}footer`);
    preserveExactLineInk(contentArea,
      Math.max(paper.top, header?.getBoundingClientRect().bottom ?? paper.top),
      Math.min(paper.bottom, footer?.getBoundingClientRect().top ?? paper.bottom,
        notesEl?.getBoundingClientRect().top ?? paper.bottom), paper);

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
    alignBaselinesToWord: options.alignBaselinesToWord ?? true,
    commentColumnPlacement: options.commentColumnPlacement ?? "outside",
  });
  return engine.paginate();
}

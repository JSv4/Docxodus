/**
 * Page geometry — the document's own page setup, and the scaling a view applies to show it.
 *
 * A Word document's line breaking is authored against ONE width: the text column its
 * `w:sectPr` defines (page size minus margins). That width is a property of the document,
 * not of the device showing it, which is why the converter emits it on every section wrapper
 * (`data-page-width`, `data-content-width`, the four margins, …) in every render mode.
 *
 * A view that instead sizes the column from the viewport gets a different line-breaking
 * result on every screen, and — once a run is enlarged past what the narrowed column can
 * hold — content that no longer fits the page at all. Word and LibreOffice never do that:
 * they keep the column fixed and ZOOM the page to fit the window. This module is the shared
 * owner of both halves — reading the geometry, and computing/applying that zoom — so the
 * paginated view (`pagination.ts`) and the continuous view (`editor.ts` via `viewport.ts`)
 * agree on what a page is instead of each inventing its own.
 */

/** Page dimensions extracted from a section wrapper's data attributes (in points). */
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

/** US Letter portrait — Word's default, and this module's fallback for an unstamped section. */
export const DEFAULT_PAGE_WIDTH = 612;
export const DEFAULT_PAGE_HEIGHT = 792;
export const DEFAULT_MARGIN = 72; // 1 inch
/** Default header/footer distance (0.5 inch). */
export const DEFAULT_HEADER_FOOTER_HEIGHT = 36;

/** Converts pixels to points (assuming 96 DPI screen). */
export function pxToPt(px: number): number {
  return px * 0.75; // 72 points / 96 pixels
}

/** Converts points to pixels (assuming 96 DPI screen). */
export function ptToPx(pt: number): number {
  return pt / 0.75;
}

/** Parses page dimensions from a section wrapper's data attributes. */
export function parseSectionDimensions(section: HTMLElement): PageDimensions {
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

/** Every section wrapper under `root`, or `root` itself when the render stamped none. */
export function sectionWrappers(root: HTMLElement): HTMLElement[] {
  const sections = Array.from(root.querySelectorAll<HTMLElement>("[data-section-index]"));
  return sections.length > 0 ? sections : [root];
}

/**
 * The zoom that fits `naturalPt` of document into `availablePx` of window, never magnifying
 * past `max`. Below `MIN_FIT_SCALE` the text would be unreadable, so the view stops shrinking
 * and lets the surface scroll horizontally instead — the same trade a word processor makes.
 */
export const MIN_FIT_SCALE = 0.25;

export function fitScale(availablePx: number, naturalPt: number, max = 1): number {
  if (!(availablePx > 0) || !(naturalPt > 0)) return max;
  const naturalPx = ptToPx(naturalPt);
  if (naturalPx <= availablePx) return max;
  return Math.max(MIN_FIT_SCALE, Math.min(max, availablePx / naturalPx));
}

/**
 * Applies a zoom factor to `el`.
 *
 * Prefers CSS `zoom`, which participates in layout — so the element's own box shrinks, the
 * page keeps its scrollbars honest, and caret/hit-testing geometry stays correct inside a
 * `contenteditable`. `transform: scale()` is the fallback for engines without `zoom`; it does
 * not affect layout, so the caller must compensate for the space the unscaled box still
 * reserves, which is what `naturalPt` is for.
 */
export function applyZoom(
  el: HTMLElement,
  scale: number,
  naturalPt?: { width: number; height: number },
): void {
  if (scale === 1) {
    el.style.removeProperty("zoom");
    el.style.removeProperty("transform");
    el.style.removeProperty("transform-origin");
    el.style.removeProperty("margin-right");
    el.style.removeProperty("margin-bottom");
    return;
  }
  const zoomSupported =
    typeof CSS !== "undefined" && typeof CSS.supports === "function" && CSS.supports("zoom", "0.5");
  if (zoomSupported) {
    el.style.zoom = String(scale);
    return;
  }
  el.style.transform = `scale(${scale})`;
  el.style.transformOrigin = "top left";
  if (naturalPt) {
    el.style.marginRight = `-${ptToPx(naturalPt.width * (1 - scale))}px`;
    el.style.marginBottom = `-${ptToPx(naturalPt.height * (1 - scale))}px`;
  }
}

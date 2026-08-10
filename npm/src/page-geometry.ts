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
  /** `w:pgMar/@w:header` — page top edge to the TOP of the header story, in points. */
  headerDistance: number;
  /** `w:pgMar/@w:footer` — page bottom edge to the BOTTOM of the footer story, in points. */
  footerDistance: number;
  /** @deprecated Misnamed alias of {@link headerDistance} — it never held a height. */
  headerHeight: number;
  /** @deprecated Misnamed alias of {@link footerDistance} — it never held a height. */
  footerHeight: number;
}

/** US Letter portrait — Word's default, and this module's fallback for an unstamped section. */
export const DEFAULT_PAGE_WIDTH = 612;
export const DEFAULT_PAGE_HEIGHT = 792;
export const DEFAULT_MARGIN = 72; // 1 inch
/** Default header/footer distance from the paper edge (0.5 inch). */
export const DEFAULT_HEADER_FOOTER_DISTANCE = 36;
/** @deprecated Misnamed alias of {@link DEFAULT_HEADER_FOOTER_DISTANCE}. */
export const DEFAULT_HEADER_FOOTER_HEIGHT = DEFAULT_HEADER_FOOTER_DISTANCE;

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
  // `data-header-height`/`data-footer-height` carry `w:pgMar/@w:header` and `@w:footer`, which
  // are distances from the paper edge — the attribute names predate that being understood.
  const headerDistance =
    parseFloat(section.dataset.headerHeight || "") || DEFAULT_HEADER_FOOTER_DISTANCE;
  const footerDistance =
    parseFloat(section.dataset.footerHeight || "") || DEFAULT_HEADER_FOOTER_DISTANCE;

  return {
    pageWidth,
    pageHeight,
    contentWidth,
    contentHeight,
    marginTop,
    marginRight,
    marginBottom,
    marginLeft,
    headerDistance,
    footerDistance,
    headerHeight: headerDistance,
    footerHeight: footerDistance,
  };
}

/**
 * Where a page's three bands sit, in points measured from the page's TOP edge.
 *
 * Resolved by {@link resolvePageBands} from the section's page setup plus the measured
 * height of the running content actually selected for that page.
 */
export interface PageBands {
  /** Top of the header band; it grows downward from here. */
  headerTop: number;
  /** Height of the header band (the measured header story, clipped to the body's top). */
  headerHeight: number;
  /** Top of the body text band. */
  bodyTop: number;
  /** Height of the body text band. */
  bodyHeight: number;
  /** Top of the footer band; it grows upward, so its BOTTOM is the fixed edge. */
  footerTop: number;
  /** Height of the footer band (the measured footer story, clipped to the body's bottom). */
  footerHeight: number;
}

/**
 * A body band this short cannot hold a single line, and a page that can hold nothing sends
 * every block down the oversized-block path. Running content may push the body — Word does
 * that — but it may not delete it, so a pathological header/footer gives this much back.
 */
export const MIN_BODY_BAND_PT = 12;

/**
 * Word's page band model.
 *
 * `w:header`/`w:footer` are DISTANCES from the paper edge to the running content, and
 * `w:top`/`w:bottom` are distances from the paper edge to the body text — four independent
 * numbers, not two nested boxes. The header is laid out from `headerDistance` downward and
 * the footer from `footerDistance` upward; the body then starts at the top margin unless the
 * header has already run past it, and ends at the bottom margin unless the footer has
 * already climbed above it.
 *
 * Anchoring the bands to the margins instead — header bottom-aligned to `marginTop`, footer
 * top-aligned to `pageHeight - marginBottom` — collapses both toward the body by exactly
 * `margin - distance`, which is what this function replaced.
 */
export function resolvePageBands(
  dims: PageDimensions,
  headerContentHeight: number,
  footerContentHeight: number,
): PageBands {
  const headerTop = dims.headerDistance;
  const footerBottom = dims.pageHeight - dims.footerDistance;

  // A page with NO running story has nothing below the paper edge to push the body: the
  // distance alone reserves no space. (Word's own empty header is still a story — one empty
  // paragraph with a real line height — so zero here means absent, not blank.)
  const headerReach = headerContentHeight > 0 ? headerTop + headerContentHeight : 0;
  const footerReach = footerContentHeight > 0 ? footerBottom - footerContentHeight : dims.pageHeight;

  let bodyTop = Math.max(dims.marginTop, headerReach);
  let bodyBottom = Math.min(dims.pageHeight - dims.marginBottom, footerReach);

  // Hand back whatever the bands over-claimed, in proportion to how far each grew past its
  // own margin, so a tall header does not silently consume the whole page.
  const shortfall = MIN_BODY_BAND_PT - (bodyBottom - bodyTop);
  if (shortfall > 0) {
    const headerGrowth = bodyTop - dims.marginTop;
    const footerGrowth = dims.pageHeight - dims.marginBottom - bodyBottom;
    const growth = headerGrowth + footerGrowth;
    if (growth > 0) {
      bodyTop -= (shortfall * headerGrowth) / growth;
      bodyBottom += (shortfall * footerGrowth) / growth;
    }
  }

  return {
    headerTop,
    headerHeight: Math.max(0, Math.min(headerContentHeight, bodyTop - headerTop)),
    bodyTop,
    bodyHeight: Math.max(0, bodyBottom - bodyTop),
    footerTop: footerBottom - Math.max(0, Math.min(footerContentHeight, footerBottom - bodyBottom)),
    footerHeight: Math.max(0, Math.min(footerContentHeight, footerBottom - bodyBottom)),
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

/**
 * The document viewport — what a word processor calls the page view's zoom.
 *
 * The editor's continuous view used to have no page geometry at all: it injected the
 * converter's body into whatever box the host gave it, so the text column was the DEVICE's
 * width. Three things follow from that, and all three are wrong:
 *
 *  - Line breaking differs on every screen, and matches neither Word nor the editor's own
 *    paginated view of the same document.
 *  - A table authored at, say, 496pt cannot fit a 354pt phone column. A table box never
 *    shrinks below its content's minimum, so it overflows the sheet and is clipped by the
 *    window — the document visibly runs off the paper.
 *  - Enlarging a run's font size makes that minimum grow, so the overflow gets worse exactly
 *    when the user is editing.
 *
 * Word and LibreOffice hold the column at the section's authored width and ZOOM the page to
 * fit the window. This class does the same: it stamps each section wrapper with its own
 * `w:sectPr` geometry (content width + margins, so the section IS the sheet), then observes
 * the host and applies a fit-to-width zoom. Narrow screens get a smaller — never a reflowed —
 * page, so what the user sees is what the document says.
 */

import {
  applyZoom,
  fitScale,
  parseSectionDimensions,
  ptToPx,
  pxToPt,
  sectionWrappers,
} from "./page-geometry.js";

/**
 * How the continuous view sizes its text column.
 *
 * - `"section"` (default): the section's own content width, as Word lays it out. Fidelity.
 * - `"fluid"`: the host's width, the pre-geometry behavior. For embeds that want the document
 *   to reflow as ordinary web text and accept that line breaking will not match Word.
 */
export type ColumnWidth = "section" | "fluid";

export interface DocumentViewportOptions {
  /** Text-column sizing for the continuous view. Default `"section"`. */
  columnWidth?: ColumnWidth;
  /**
   * Scale the page down when it is wider than the host, the way a word processor's
   * fit-to-width zoom does. Default true. With it off, an oversized page simply overflows
   * (the host may scroll).
   */
  fitToWidth?: boolean;
  /** Author-pinned zoom, applied on top of fit-to-width (which never magnifies past it). Default 1. */
  scale?: number;
}

export class DocumentViewport {
  private readonly host: HTMLElement;
  private readonly options: Required<DocumentViewportOptions>;
  private root: HTMLElement | null = null;
  /** Natural page size in points — the widest section's PAGE box, which is what must fit. */
  private natural = { width: 0, height: 0 };
  private observer: ResizeObserver | null = null;

  constructor(host: HTMLElement, options: DocumentViewportOptions = {}) {
    this.host = host;
    this.options = {
      columnWidth: options.columnWidth ?? "section",
      fitToWidth: options.fitToWidth ?? true,
      scale: options.scale ?? 1,
    };
  }

  /**
   * Adopt a freshly mounted document root (a continuous flow, or the paginated page stack).
   * Safe to call on every remount; the previous root is released first.
   *
   * `applySectionGeometry` is false for the paginated view, which already builds real page
   * boxes at the section's dimensions — there the viewport contributes only the fit zoom.
   */
  attach(root: HTMLElement, applySectionGeometry: boolean): void {
    this.release();
    this.root = root;
    this.natural = applySectionGeometry ? this.stampSections(root) : this.measurePages(root);
    this.refresh();
    if (typeof ResizeObserver !== "undefined") {
      this.observer = new ResizeObserver(() => this.refresh());
      this.observer.observe(this.host);
    }
  }

  /** Recompute the fit zoom against the host's current width. */
  refresh(): void {
    if (!this.root) return;
    const scale = this.scale;
    applyZoom(this.root, scale, this.natural);
    // Publish the page's on-screen width so chrome docked OUTSIDE the zoomed sheet — the
    // header/footer bands — can line up with it at whatever zoom is in force.
    this.host.style.setProperty(
      "--docx-sheet-width",
      this.natural.width > 0 ? `${ptToPx(this.natural.width) * scale}px` : "100%",
    );
  }

  /**
   * Change the author-pinned zoom and re-apply it. Fit-to-width still caps it: a page wider
   * than the host never magnifies past what fits, so "100%" on a phone is the fit zoom.
   */
  setScale(scale: number): void {
    this.options.scale = Math.max(0.1, Math.min(4, scale));
    this.refresh();
  }

  /** The author-pinned zoom (the value a zoom control shows), before fit-to-width caps it. */
  get requestedScale(): number {
    return this.options.scale;
  }

  /** The zoom currently applied (1 = 100%). Reported by the ribbon's anchor rail. */
  get scale(): number {
    if (!this.root) return this.options.scale;
    return this.options.fitToWidth
      ? fitScale(this.availableWidthPx(), this.natural.width, this.options.scale)
      : this.options.scale;
  }

  dispose(): void {
    this.release();
    this.root = null;
  }

  private release(): void {
    this.observer?.disconnect();
    this.observer = null;
    if (this.root) applyZoom(this.root, 1);
    this.host.style.removeProperty("--docx-sheet-width");
  }

  /** The host's content box, which is the space a page has to fit into. */
  private availableWidthPx(): number {
    const style = typeof getComputedStyle === "function" ? getComputedStyle(this.host) : null;
    const padding = style
      ? (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0)
      : 0;
    return Math.max(0, this.host.clientWidth - padding);
  }

  /**
   * Give each section wrapper its `w:sectPr` geometry: the authored text column, guttered by
   * the authored margins. The wrapper then measures exactly one page wide, which is what the
   * sheet chrome paints and what the fit zoom scales.
   */
  private stampSections(root: HTMLElement): { width: number; height: number } {
    const sections = sectionWrappers(root);
    if (this.options.columnWidth === "fluid") {
      // The pre-geometry behavior: the host's width is the column. Nothing to fit.
      root.style.removeProperty("width");
      for (const section of sections) {
        section.style.removeProperty("width");
        section.style.removeProperty("padding-left");
        section.style.removeProperty("padding-right");
      }
      return { width: 0, height: 0 };
    }

    let widest = 0;
    for (const section of sections) {
      const dims = parseSectionDimensions(section);
      section.style.width = `${dims.contentWidth}pt`;
      section.style.paddingLeft = `${dims.marginLeft}pt`;
      section.style.paddingRight = `${dims.marginRight}pt`;
      section.style.boxSizing = "content-box";
      section.style.marginLeft = "auto";
      section.style.marginRight = "auto";
      widest = Math.max(widest, dims.pageWidth);
    }
    // The sheet is exactly as wide as the widest page it holds, so the chrome that paints it
    // paints a page rather than a full-bleed panel — and centering it is the host's business.
    if (widest > 0 && sections[0] !== root) root.style.width = `${widest}pt`;
    return { width: widest, height: 0 };
  }

  /**
   * The paginated view's page boxes are already page-sized — `pagination.ts` writes each
   * box's `width` in points — so the widest of those is the natural width. Reading the
   * inline width rather than the laid-out box keeps this independent of the per-box zoom
   * pagination may itself have applied.
   */
  private measurePages(root: HTMLElement): { width: number; height: number } {
    let widest = 0;
    for (const box of Array.from(root.children) as HTMLElement[]) {
      const declared = /^([\d.]+)pt$/.exec(box.style.width || "");
      widest = Math.max(widest, declared ? parseFloat(declared[1]) : pxToPt(box.offsetWidth));
    }
    return { width: widest, height: 0 };
  }
}

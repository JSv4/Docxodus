/**
 * Physical paragraph-border widths for the standalone export (issue #1116).
 *
 * CSS lays every border out in whole CSS pixels, never below one, so a declared 0.25pt
 * (`w:sz="2"`) paragraph border is both drawn and spaced as 0.75pt. The converter records each
 * explicitly sized single or dashed side as `--docx-border-{side}-width` on the border wrapper
 * (`data-docx-exact-border`). Before pagination, `prepareParagraphBorders` replaces those CSS
 * borders with padding of the declared width, so the box and the text it holds are laid out
 * exactly. After the page tree is final, `paintParagraphBorders` draws the recorded sides as
 * vector shapes from each page box's origin, which is the only coordinate space the PDF printer
 * does not snap. Omitted-size borders and other border styles keep their CSS rendering.
 *
 * The shapes are drawn over the page, outside the paragraph's own ancestors, so they do not inherit
 * the clips the paragraph sits under. An exact-height table row hides the borders of the content it
 * hides, as Word does (issue #1171), so each side is trimmed to the vertical span of the exact-row
 * cells around the paragraph. The row clips by height only: a left or right border in the cell
 * margin, beside the text, is drawn in full.
 */

import { exactRowClipSpan } from "./exact-row-text.js";

const EXACT = "data-docx-exact-border";
/** Set by `prepareParagraphBorders` on a border box whose recorded sides it paints after pagination. */
export const PAINT = "data-docx-border-paint";
const SVG_NS = "http://www.w3.org/2000/svg";
const SIDES = ["top", "right", "bottom", "left"] as const;
type Side = (typeof SIDES)[number];

interface PaintedSide {
  side: Side;
  style: "solid" | "dashed";
  color: string;
  /** CSS pixels. */
  width: number;
}

const PT_TO_PX = 4 / 3;

/** Lay out recorded paragraph borders at their declared width. Run before pagination measures. */
export function prepareParagraphBorders(root: HTMLElement): number {
  const view = root.ownerDocument.defaultView;
  if (!view) return 0;
  let prepared = 0;
  for (const element of Array.from(root.querySelectorAll<HTMLElement>(`[${EXACT}]`))) {
    if (element.hasAttribute(PAINT)) continue;
    const style = view.getComputedStyle(element);
    const painted: PaintedSide[] = [];
    for (const side of SIDES) {
      const declared = Number.parseFloat(style.getPropertyValue(`--docx-border-${side}-width`));
      const cssStyle = style.getPropertyValue(`border-${side}-style`);
      if (!(declared > 0) || (cssStyle !== "solid" && cssStyle !== "dashed")) continue;
      const exact = declared * PT_TO_PX;
      const laidOut = Number.parseFloat(style.getPropertyValue(`border-${side}-width`)) || 0;
      const padding = Number.parseFloat(style.getPropertyValue(`padding-${side}`)) || 0;
      painted.push({ side, style: cssStyle, color: style.getPropertyValue(`border-${side}-color`), width: exact });
      element.style.setProperty(`border-${side}-style`, "none");
      element.style.setProperty(`padding-${side}`, `${padding + exact}px`);
      // Horizontal borders sit outside the text edge (issue #1126): keep that edge where the
      // pixel-width layout put it.
      if (side === "left" || side === "right") {
        const margin = Number.parseFloat(style.getPropertyValue(`margin-${side}`)) || 0;
        element.style.setProperty(`margin-${side}`, `${margin + laidOut - exact}px`);
      }
    }
    if (painted.length === 0) continue;
    // CSSOM keeps more precision than it serializes. Reparse the values that offline HTML
    // will contain before pagination: a near-integer padding can otherwise gain a layout
    // unit on reopening, accumulating into different text positions across border boxes.
    element.style.cssText = element.style.cssText;
    element.setAttribute(PAINT, JSON.stringify(painted));
    prepared++;
  }
  return prepared;
}

/**
 * Leave one horizontal side of a prepared border box unpainted: the part of a bordered paragraph
 * group that a page break opens (issue #1202).
 */
export function omitPaintedBorderSide(element: HTMLElement, side: "top" | "bottom"): void {
  const recorded = element.getAttribute(PAINT);
  if (!recorded) return;
  const sides = (JSON.parse(recorded) as PaintedSide[]).filter(painted => painted.side !== side);
  element.setAttribute(PAINT, JSON.stringify(sides));
}

/** Paint prepared borders on the final page tree. Layout and the PageMap do not change here. */
export function paintParagraphBorders(container: HTMLElement, pageBoxSelector = ".page-box"): number {
  const document = container.ownerDocument;
  let painted = 0;
  const overlays = new Map<HTMLElement, SVGSVGElement>();
  const overlayFor = (page: HTMLElement): SVGSVGElement => {
    let svg = overlays.get(page);
    if (!svg) {
      svg = document.createElementNS(SVG_NS, "svg");
      svg.setAttribute("aria-hidden", "true");
      svg.setAttribute("width", String(page.offsetWidth));
      svg.setAttribute("height", String(page.offsetHeight));
      svg.style.cssText = "position:absolute;left:0;top:0;overflow:visible;pointer-events:none;z-index:1";
      page.append(svg);
      overlays.set(page, svg);
    }
    return svg;
  };
  for (const element of Array.from(container.querySelectorAll<HTMLElement>(`[${PAINT}]`))) {
    const page = element.closest<HTMLElement>(pageBoxSelector);
    if (!page) continue;
    let sides: PaintedSide[];
    try {
      sides = JSON.parse(element.getAttribute(PAINT)!) as PaintedSide[];
    } catch {
      continue;
    }
    const pageRect = page.getBoundingClientRect();
    const scale = page.offsetWidth > 0 ? pageRect.width / page.offsetWidth : 1;
    const box = element.getBoundingClientRect();
    if (!(box.width > 0) || !(box.height > 0) || !(scale > 0)) continue;
    const left = (box.left - pageRect.left) / scale;
    const top = (box.top - pageRect.top) / scale;
    const width = box.width / scale;
    const height = box.height / scale;
    // The vertical span the paragraph's exact rows show, in the overlay's coordinates; null when
    // no exact row cuts it.
    const span = exactRowClipSpan(element);
    const clip = !span || (span.top <= box.top && span.bottom >= box.bottom)
      ? null
      : { top: (span.top - pageRect.top) / scale, bottom: (span.bottom - pageRect.top) / scale };
    let paintedSides = 0;
    for (const { side, style, color, width: stroke } of sides) {
      const horizontal = side === "top" || side === "bottom";
      const x = side === "right" ? left + width - stroke : left;
      let y = side === "bottom" ? top + height - stroke : top;
      const w = horizontal ? width : stroke;
      let h = horizontal ? stroke : height;
      // How far a trimmed side starts into the untrimmed one, so a dashed side keeps its rhythm.
      let phase = 0;
      if (clip) {
        const y0 = Math.max(y, clip.top);
        const y1 = Math.min(y + h, clip.bottom);
        if (!(y1 > y0)) continue;
        if (!horizontal) phase = y0 - y;
        y = y0;
        h = y1 - y0;
      }
      const svg = overlayFor(page);
      if (style === "solid") {
        const rect = document.createElementNS(SVG_NS, "rect");
        for (const [name, value] of [["x", x], ["y", y], ["width", w], ["height", h]] as const)
          rect.setAttribute(name, String(value));
        rect.setAttribute("fill", color);
        svg.append(rect);
      } else {
        // Keep the dash rhythm the browser draws for a CSS dashed border of this side (at least one
        // CSS pixel wide); only the stroke width changes. Word's own pattern is not yet captured.
        const line = document.createElementNS(SVG_NS, "line");
        const cx = x + w / 2;
        const cy = y + h / 2;
        line.setAttribute("x1", String(horizontal ? x : cx));
        line.setAttribute("y1", String(horizontal ? cy : y));
        line.setAttribute("x2", String(horizontal ? x + w : cx));
        line.setAttribute("y2", String(horizontal ? cy : y + h));
        line.setAttribute("stroke", color);
        // A side the clip cuts across keeps only the stroke's visible part.
        line.setAttribute("stroke-width", String(horizontal ? h : w));
        const rhythm = Math.max(1, stroke);
        line.setAttribute("stroke-dasharray", `${3 * rhythm} ${2 * rhythm}`);
        if (phase > 0) line.setAttribute("stroke-dashoffset", String(phase));
        svg.append(line);
      }
      paintedSides++;
    }
    if (paintedSides > 0) painted++;
  }
  return painted;
}

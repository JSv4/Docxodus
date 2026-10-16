/**
 * Whole-text clipping for exact-height table rows in the standalone PDF export (issue #1140).
 *
 * The converter clips an exact row's cell content with `overflow: hidden`, as Word clips it. Chromium's
 * PDF writer (Skia's SkPDFDevice) keeps a glyph in the text layer only when the glyph's bounds meet the
 * bounding rectangle of the clip it is drawn under. A line the row cuts through therefore extracted as
 * broken pieces: the letters whose ascenders reach the row survived and the rest of every word was dropped.
 *
 * A clip whose painted area is the row but whose bounding rectangle covers the hidden content keeps every
 * glyph in the text layer, once, in place. Only a path can be that shape: the visible rectangle plus
 * zero-area spikes. In CSS a clip path makes its element a stacking context, which moves the cell's text
 * to the end of the page's content stream; an SVG `clipPath` inside an inline SVG root does not. So the
 * content of an overflowing clip box moves into a `foreignObject` clipped that way.
 *
 * Geometry. Chromium paints an SVG root at a whole pixel of the page. The root is placed on the pixel at
 * or before the clip box (negative margins keep the cell's layout) and the `foreignObject` takes back the
 * fraction, so the content lays out and paints exactly where it did. The visible rectangle is the
 * whole-pixel rectangle Chromium snaps the `overflow: hidden` clip to. The clip box keeps an overflow clip
 * (for the PageMap measurer, which intersects every overflow-clipping ancestor) with a margin wide enough
 * not to bound the glyphs. Page boxes print at their own origin, so both grids are the page box's.
 *
 * Paint order. The SVG root paints its content atomically, so the cell's text moves to whatever paint
 * phase the root is in. Within a stacking context, text in a positioned box (the relative offsets that
 * `alignBaselinesToWord` and exact line spacing put on runs) paints after all non-positioned content, in
 * tree order; other text paints earlier, with the in-flow content. The root takes `position: relative`
 * exactly when the clip box's text painted in the positioned phase, so the cell keeps its place in the
 * content stream. (A cell mixing both kinds of text was already split across the two phases; it paints in
 * the positioned phase, at its place in tree order.)
 *
 * Cost. The hidden glyphs are now drawn under the clip instead of being left out. They lie outside the
 * clip, but a PDF rasterizer can still tint the pixel row the clip edge crosses where a hidden glyph's
 * top touches that row (anti-aliased coverage, hinting). See docs/architecture/standalone_paginated_export.md.
 *
 * Left alone: a clip box whose content fits, and one holding an absolutely or fixed positioned element,
 * for which the `foreignObject` would become the containing block and move it.
 */

const EXACT_ROW_CELL = 'div[data-docx-exact-row-cell="true"]';
/** Marks the SVG root that now holds an exact-row cell's clipped content. */
export const EXACT_ROW_TEXT_ATTRIBUTE = 'data-docx-exact-row-text';
const SVG_NS = 'http://www.w3.org/2000/svg';
/** Reach of the zero-area spikes and of the overflow clip margin, in CSS pixels: beyond any page. */
const REACH = 100000;
/** Layout overflow below this many CSS pixels cannot hold a whole glyph outside the clip. */
const OVERFLOW_EPSILON = 1 / 64;

// Layout positions are multiples of 1/64 px, which six decimals hold exactly: a rounded margin and height
// that no longer cancel would grow the row by a layout unit.
const number = (value: number) => String(Number(value.toFixed(6)));
const px = (value: number) => `${number(value)}px`;

/** Whether the clip box's content extends past it on either axis. */
function overflows(clip: HTMLElement): boolean {
  if (clip.scrollHeight > clip.clientHeight || clip.scrollWidth > clip.clientWidth) return true;
  // scrollHeight/scrollWidth are whole pixels; measure the content itself for a fractional overflow.
  const box = clip.getBoundingClientRect();
  const range = clip.ownerDocument.createRange();
  range.selectNodeContents(clip);
  return Array.from(range.getClientRects()).some(rect => rect.width > 0 && rect.height > 0 && (
    rect.bottom > box.bottom + OVERFLOW_EPSILON || rect.top < box.top - OVERFLOW_EPSILON ||
    rect.right > box.right + OVERFLOW_EPSILON || rect.left < box.left - OVERFLOW_EPSILON));
}

/**
 * Whether the clip box's text painted in the positioned phase of its stacking context: some text node in
 * it has a positioned box between it and the clip box.
 */
function textPaintsPositioned(clip: HTMLElement, view: Window & typeof globalThis): boolean {
  const walker = clip.ownerDocument.createTreeWalker(clip, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!(node.nodeValue ?? '').trim()) continue;
    for (let element = node.parentElement; element && element !== clip; element = element.parentElement)
      if (view.getComputedStyle(element).position !== 'static') return true;
  }
  return false;
}

/** Whether the clip box holds an element that a `foreignObject` would re-anchor. */
function holdsOutOfFlowPositioned(clip: HTMLElement, view: Window & typeof globalThis): boolean {
  return Array.from(clip.querySelectorAll('*'))
    .some(element => ['absolute', 'fixed'].includes(view.getComputedStyle(element).position));
}

/**
 * Move every exact-row clip box whose content overflows it into an SVG clipped to the same visible
 * rectangle, so the clipped text stays whole in the PDF text layer. Run on the final page tree, after
 * the export's clipping check and the baseline offsets. Returns how many clip boxes moved.
 */
export function preserveClippedExactRowText(root: HTMLElement): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  let rewritten = 0;
  for (const clip of Array.from(root.querySelectorAll<HTMLElement>(EXACT_ROW_CELL))) {
    const page = clip.closest<HTMLElement>('.page-box');
    const parent = clip.parentElement;
    if (!page || !parent || !overflows(clip) || holdsOutOfFlowPositioned(clip, view)) continue;
    const pageRect = page.getBoundingClientRect();
    const box = clip.getBoundingClientRect();
    const left = box.left - pageRect.left;
    const top = box.top - pageRect.top;
    // The SVG root's origin: the whole pixel at or before the clip box.
    const fx = left - Math.floor(left);
    const fy = top - Math.floor(top);
    // The rectangle Chromium snaps the overflow clip to, in the SVG root's coordinates.
    const x0 = Math.round(left) - Math.floor(left);
    const y0 = Math.round(top) - Math.floor(top);
    const x1 = Math.round(left + box.width) - Math.floor(left);
    const y1 = Math.round(top + box.height) - Math.floor(top);

    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute(EXACT_ROW_TEXT_ATTRIBUTE, 'true');
    // The standalone document admits an inline SVG root only with a definite size.
    svg.setAttribute('width', number(box.width + fx));
    svg.setAttribute('height', number(box.height + fy));
    const position = textPaintsPositioned(clip, view) ? 'relative' : 'static';
    svg.setAttribute('style', `display:block;position:${position};overflow:visible;`
      + `width:${px(box.width + fx)};height:${px(box.height + fy)};margin:${px(-fy)} 0 0 ${px(-fx)}`);
    const clipPath = document.createElementNS(SVG_NS, 'clipPath');
    const id = `docx-exact-row-clip-${rewritten}`;
    clipPath.id = id;
    clipPath.setAttribute('clipPathUnits', 'userSpaceOnUse');
    const shape = document.createElementNS(SVG_NS, 'path');
    // The visible rectangle, then zero-area spikes that stretch the clip's bounds in every direction.
    shape.setAttribute('d', `M${x0} ${y0}H${x1}V${y1}H${x0}Z`
      + ` M${x0} ${y0}H${REACH} M${x0} ${y0}H${-REACH} M${x0} ${y0}V${REACH} M${x0} ${y0}V${-REACH}`);
    clipPath.appendChild(shape);
    const content = document.createElementNS(SVG_NS, 'foreignObject');
    content.setAttribute('x', number(fx));
    content.setAttribute('y', number(fy));
    content.setAttribute('width', number(box.width));
    content.setAttribute('height', number(box.height));
    content.setAttribute('clip-path', `url(#${id})`);
    content.setAttribute('style', 'overflow:visible');
    svg.append(clipPath, content);

    const display = view.getComputedStyle(clip).display;
    parent.replaceChild(svg, clip);
    content.appendChild(clip);
    clip.style.setProperty('width', px(box.width));
    // `overflow: hidden` made the box a block formatting context; keep its margins from collapsing.
    if (display === 'block') clip.style.setProperty('display', 'flow-root');
    clip.style.setProperty('overflow', 'clip');
    clip.style.setProperty('overflow-clip-margin', px(REACH));
    rewritten++;
  }
  return rewritten;
}

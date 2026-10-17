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
 * `alignBaselinesToWord` and exact line spacing put on runs) paints after all in-flow content, in tree
 * order; other text paints with the in-flow content. The root takes `position: relative` when any of the
 * clip box's text painted in the positioned phase, so the cell paints where that text did; a cell whose
 * text is all in-flow keeps the in-flow slot. Inside the root (a stacking context of its own) a cell with
 * text in both phases gives its in-flow text an offset-free relative box, so the cell reads in tree order.
 *
 * Cost. The hidden glyphs are now drawn under the clip instead of being left out. They lie outside the
 * clip, but a PDF rasterizer can still tint the pixel row the clip edge crosses where a hidden glyph's
 * top touches that row (anti-aliased coverage, hinting). See docs/architecture/standalone_paginated_export.md.
 *
 * Left alone: a clip box that hides no glyph (see `hidesText`), and one holding an absolutely or fixed
 * positioned element, for which the `foreignObject` would become the containing block and move it.
 */

/** The converter's clip box around an exact-height row cell's content (`ClipExactRowCells`). */
export const EXACT_ROW_CELL = 'div[data-docx-exact-row-cell="true"]';
/** Marks the SVG root that now holds an exact-row cell's clipped content. */
const EXACT_ROW_TEXT_ATTRIBUTE = 'data-docx-exact-row-text';
const SVG_NS = 'http://www.w3.org/2000/svg';
/** Reach of the zero-area spikes and of the overflow clip margin, in CSS pixels: beyond any page. */
const REACH = 100000;
// Layout positions are multiples of 1/64 px, which six decimals hold exactly: a rounded margin and height
// that no longer cancel would grow the row by a layout unit.
const number = (value: number) => String(Number(value.toFixed(6)));
const px = (value: number) => `${number(value)}px`;

/** The clip box's text nodes that hold more than white space. */
function textNodes(clip: HTMLElement): Text[] {
  const nodes: Text[] = [];
  const walker = clip.ownerDocument.createTreeWalker(clip, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode())
    if ((node.nodeValue ?? '').trim()) nodes.push(node as Text);
  return nodes;
}

/**
 * Whether the clip hides some of the clip box's text, in the sense that matters to the PDF writer: a glyph
 * wholly outside the clip. Layout overflow alone does not say so: a line box includes the font's whole
 * ascent and descent and a paragraph its spacing, so a one-line row a little shorter than its line, or a
 * paragraph's space after, overflows without hiding a glyph. Glyphs stand on their baseline, so a line
 * loses glyphs vertically only when the clip edge passes its baseline: below it, every glyph shows at
 * least the part above the baseline (a glyph drawn wholly below its baseline, such as an underscore, is
 * not looked for). Horizontally, a glyph is lost when its whole advance lies outside.
 */
function hidesText(clip: HTMLElement, view: Window & typeof globalThis): boolean {
  const document = clip.ownerDocument;
  const box = clip.getBoundingClientRect();
  const context = document.createElement('canvas').getContext('2d');
  const range = document.createRange();
  for (const node of textNodes(clip)) {
    const style = view.getComputedStyle(node.parentElement!);
    // The share of a line box's content area below the baseline; without it, the whole box counts.
    let descent = 0;
    if (context) {
      context.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
      const metrics = context.measureText('');
      const height = metrics.fontBoundingBoxAscent + metrics.fontBoundingBoxDescent;
      if (height > 0) descent = metrics.fontBoundingBoxDescent / height;
    }
    range.selectNodeContents(node);
    for (const line of Array.from(range.getClientRects())) {
      if (!(line.width > 0) || !(line.height > 0)) continue;
      const baseline = line.bottom - descent * line.height;
      if (baseline > box.bottom || baseline < box.top) return true;
      if (line.right <= box.right && line.left >= box.left) continue;
      // A line that crosses a side edge: look for a character wholly beyond it.
      const text = node.nodeValue ?? '';
      for (let index = 0; index < text.length; index++) {
        range.setStart(node, index);
        range.setEnd(node, index + 1);
        if (Array.from(range.getClientRects()).some(glyph => glyph.width > 0
          && (glyph.left >= box.right || glyph.right <= box.left))) return true;
      }
      range.selectNodeContents(node);
    }
  }
  return false;
}

/**
 * The clip box's text nodes that painted in the positioned phase of their stacking context (a positioned
 * box between them and the clip box), and those that painted with the in-flow content.
 */
function textByPaintPhase(clip: HTMLElement, view: Window & typeof globalThis): { positioned: Text[]; inFlow: Text[] } {
  const positioned: Text[] = [];
  const inFlow: Text[] = [];
  for (const node of textNodes(clip)) {
    let found = false;
    for (let element = node.parentElement; element && element !== clip && !found; element = element.parentElement)
      found = view.getComputedStyle(element).position !== 'static';
    (found ? positioned : inFlow).push(node);
  }
  return { positioned, inFlow };
}

/** Whether the clip box holds an element that a `foreignObject` would re-anchor. */
function holdsOutOfFlowPositioned(clip: HTMLElement, view: Window & typeof globalThis): boolean {
  return Array.from(clip.querySelectorAll('*'))
    .some(element => ['absolute', 'fixed'].includes(view.getComputedStyle(element).position));
}

/**
 * Move every exact-row clip box that hides some of its text into an SVG clipped to the same visible
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
    if (!page || !parent || !hidesText(clip, view) || holdsOutOfFlowPositioned(clip, view)) continue;
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
    const phases = textByPaintPhase(clip, view);
    const position = phases.positioned.length > 0 ? 'relative' : 'static';
    svg.setAttribute('style', `display:block;position:${position};overflow:visible;`
      + `width:${px(box.width + fx)};height:${px(box.height + fy)};margin:${px(-fy)} 0 0 ${px(-fx)}`);
    const clipPath = document.createElementNS(SVG_NS, 'clipPath');
    // A bookmark's name becomes an element id verbatim, and `url(#id)` takes the first element with it.
    let serial = rewritten;
    while (document.getElementById(`docx-exact-row-clip-${serial}`)) serial++;
    const id = `docx-exact-row-clip-${serial}`;
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
    // The foreignObject is a stacking context of its own. A cell whose text painted in both phases puts
    // its in-flow text in the positioned phase too (a relative box with no offset does not move), so the
    // cell's text paints in tree order inside it.
    if (phases.positioned.length > 0)
      for (const node of phases.inFlow) node.parentElement?.style.setProperty('position', 'relative');
    rewritten++;
  }
  return rewritten;
}

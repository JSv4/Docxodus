/**
 * Unrounded natural line heights for the paginated export (issue #850).
 *
 * Chromium lays a `line-height: normal` line out at the font's natural line height — ascent +
 * descent + line gap from its metrics — but first rounds that to a whole CSS pixel (0.75 pt). The
 * converter relies on `normal` for single spacing and on `calc(1lh * m)` for Word's `auto`
 * multiples, so every line inherits the rounding: 11 pt Carlito advances 18 px instead of
 * 17.90 px, and the error accumulates down the page and moves page breaks. An explicit length keeps
 * Chromium's 1/64-pixel layout precision, so this pass replaces each `normal` with the same font's
 * natural height measured where rounding is negligible.
 */

/** Font size the natural line height is probed at: rounding to a pixel there is < 0.0005 em. */
const PROBE_FONT_SIZE_PX = 1000;

/** Probe text for an element whose own text the primary font covers (Latin, through U+024F). */
const LATIN_PROBE = "Hg";

/** At most this many distinct characters of an element's own text go into its probe. */
const MAX_PROBE_CHARACTERS = 256;

/**
 * What to measure an element's natural line height with. Under `normal`, Chromium also grows a line
 * for the fallback fonts its text needs (a CJK run in a Calibri paragraph lays out taller than Latin
 * text), so an element whose own text goes beyond Latin is probed with those characters themselves.
 */
function probeText(element: Element): string {
  const characters = new Set<string>();
  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType !== 3) continue;
    for (const character of node.textContent ?? "") {
      if (/\s/.test(character)) continue;
      characters.add(character);
      if (characters.size >= MAX_PROBE_CHARACTERS) break;
    }
  }
  const beyondLatin = Array.from(characters).filter((character) => character.codePointAt(0)! > 0x024f);
  return beyondLatin.length === 0 ? LATIN_PROBE : LATIN_PROBE + beyondLatin.sort().join("");
}

/**
 * A cached measure of a font's natural line height as a fraction of its size, probed at
 * {@link PROBE_FONT_SIZE_PX} with the given text.
 */
function naturalLineHeightRatios(document: Document): (style: CSSStyleDeclaration, text: string) => number {
  const ratios = new Map<string, number>();
  return (style, text) => {
    const key = `${style.fontStyle}|${style.fontWeight}|${style.fontStretch}|${style.fontFamily}|${text}`;
    let ratio = ratios.get(key);
    if (ratio === undefined) {
      const probe = document.createElement("span");
      probe.setAttribute("aria-hidden", "true");
      probe.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;"
        + "display:inline-block;white-space:nowrap;line-height:normal;margin:0;padding:0;border:0;";
      probe.style.fontFamily = style.fontFamily;
      probe.style.fontStyle = style.fontStyle;
      probe.style.fontWeight = style.fontWeight;
      probe.style.fontStretch = style.fontStretch;
      probe.style.fontSize = `${PROBE_FONT_SIZE_PX}px`;
      probe.textContent = text;
      (document.body ?? document.documentElement).appendChild(probe);
      ratio = probe.getBoundingClientRect().height / PROBE_FONT_SIZE_PX;
      probe.remove();
      ratios.set(key, ratio);
    }
    return ratio;
  };
}

/**
 * Give every element under `root` whose computed `line-height` is `normal` an explicit line height
 * equal to its own font's unrounded natural line height (including any fallback font its own text
 * needs). Elements are collected before any is
 * changed, because an explicit value on a parent would otherwise be inherited by children that
 * should keep sizing their lines from their own font. Returns how many elements changed.
 */
export function applyUnroundedNormalLineHeights(root: Element): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const ratioFor = naturalLineHeightRatios(document);

  const targets: Array<{ element: HTMLElement; lineHeight: number }> = [];
  for (const element of [root, ...Array.from(root.querySelectorAll("*"))]) {
    if (!(element instanceof view.HTMLElement)) continue;
    const style = view.getComputedStyle(element);
    if (style.lineHeight !== "normal") continue;
    const fontSize = Number.parseFloat(style.fontSize);
    const ratio = ratioFor(style, probeText(element));
    if (!(fontSize > 0) || !(ratio > 0)) continue;
    targets.push({ element, lineHeight: ratio * fontSize });
  }
  for (const { element, lineHeight } of targets) {
    element.style.setProperty("line-height", `${lineHeight.toFixed(4)}px`);
  }
  return targets.length;
}

/** The elements the converter writes a Word paragraph as. */
const PARAGRAPH_SELECTOR = "p, h1, h2, h3, h4, h5, h6";

/** How far Chromium puts the first baseline below the top of a paragraph shaped like this one. */
function baselineOffsets(document: Document): (paragraph: CSSStyleDeclaration, child: CSSStyleDeclaration) => number {
  const offsets = new Map<string, number>();
  const host = document.body ?? document.documentElement;
  return (paragraph, child) => {
    const key = [paragraph.fontStyle, paragraph.fontWeight, paragraph.fontStretch, paragraph.fontFamily,
      paragraph.fontSize, paragraph.lineHeight, child.fontStyle, child.fontWeight, child.fontStretch,
      child.fontFamily, child.fontSize, child.lineHeight, child.verticalAlign, child.top].join("|");
    let offset = offsets.get(key);
    if (offset === undefined) {
      const probe = document.createElement("div");
      probe.setAttribute("aria-hidden", "true");
      probe.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;width:max-content;"
        + "margin:0;padding:0;border:0;white-space:nowrap;";
      for (const property of ["font-family", "font-style", "font-weight", "font-stretch", "font-size", "line-height"])
        probe.style.setProperty(property, paragraph.getPropertyValue(property));
      const span = document.createElement("span");
      for (const property of ["font-family", "font-style", "font-weight", "font-stretch", "font-size", "line-height",
        "vertical-align", "position", "top"])
        span.style.setProperty(property, child.getPropertyValue(property));
      const marker = document.createElement("span");
      marker.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0;";
      span.append(marker, LATIN_PROBE);
      probe.append(span);
      host.appendChild(probe);
      offset = marker.getBoundingClientRect().bottom - probe.getBoundingClientRect().top;
      probe.remove();
      offsets.set(key, offset);
    }
    return offset;
  };
}

/**
 * Word places an exact-spaced line's baseline at 80% of its height, independent of font size
 * (issue #882; fixtures/exact-line-spacing/word.json). Move each run by its own font's offset,
 * retaining the converter's top alignment so mixed font sizes cannot enlarge the line boxes.
 */
export function alignExactLineBaselines(root: Element): void {
  const view = root.ownerDocument.defaultView;
  if (!view) return;
  const offsetFor = baselineOffsets(root.ownerDocument);
  for (const run of Array.from(root.querySelectorAll<HTMLElement>("[data-docx-exact-run]"))) {
    const paragraph = run.closest(PARAGRAPH_SELECTOR);
    if (!paragraph || !movable(run, view)) continue;
    const style = view.getComputedStyle(paragraph);
    if (!style.getPropertyValue("--docx-exact-line-height")) continue;
    const height = Number.parseFloat(style.lineHeight);
    const runStyle = view.getComputedStyle(run);
    if (!(height > 0) || runStyle.display !== "inline" || runStyle.verticalAlign !== "top") continue;
    const top = Number.parseFloat(runStyle.top) || 0;
    const offset = offsetFor(style, runStyle) - top;
    const shift = height * 0.8 - offset;
    if (!Number.isFinite(shift)) continue;
    run.style.position = "relative";
    run.style.top = `${(top + shift).toFixed(4)}px`;
  }
}

const EXACT_INK_ATTRIBUTE = "data-docx-exact-ink";

/**
 * Exact line boxes still occupy their declared height when their glyphs extend outside them.
 * Word lets that ink enter the margins. Expand only the paint clip, after placement, for plain
 * inline text whose paragraph fits its band and whose ink cannot enter a running story or leave
 * the paper. Block geometry and pagination budgets remain unchanged.
 */
export function preserveExactLineInk(content: HTMLElement, topLimit: number, bottomLimit: number, paper: DOMRect): void {
  const view = content.ownerDocument.defaultView;
  if (!view) return;
  const band = content.getBoundingClientRect();
  const scale = band.width / content.offsetWidth || 1;
  let above = 0;
  let below = 0;
  for (const paragraph of Array.from(content.querySelectorAll<HTMLElement>(PARAGRAPH_SELECTOR))) {
    if (!view.getComputedStyle(paragraph).getPropertyValue("--docx-exact-line-height")) continue;
    const box = paragraph.getBoundingClientRect();
    if (box.top < band.top - 0.5 || box.bottom > band.bottom + 0.5) continue;
    if (!paragraph.querySelector("[data-docx-exact-run]") || paragraph.querySelector("img, svg, canvas, video, object, iframe"))
      continue;
    if (Array.from(paragraph.querySelectorAll<HTMLElement>("*")).some(element => {
      const style = view.getComputedStyle(element);
      return style.display !== "inline" || !["static", "relative"].includes(style.position);
    })) continue;
    const range = content.ownerDocument.createRange();
    range.selectNodeContents(paragraph);
    const rects = Array.from(range.getClientRects()).filter(rect => rect.width > 0 && rect.height > 0);
    if (rects.length === 0) continue;
    const top = Math.min(...rects.map(rect => rect.top));
    const bottom = Math.max(...rects.map(rect => rect.bottom));
    if (top < topLimit || bottom > bottomLimit) continue;
    const extra = Math.max(0, band.top - top, bottom - band.bottom);
    if (extra === 0) continue;
    paragraph.setAttribute(EXACT_INK_ATTRIBUTE, "true");
    above = Math.max(above, band.top - top);
    below = Math.max(below, bottom - band.bottom);
  }
  if (above > 0 || below > 0) {
    const up = Math.ceil(above / scale * 64) / 64;
    const down = Math.ceil(below / scale * 64) / 64;
    // overflow-clip-margin is ignored by Chromium when the other axis stays visible. A pixel
    // inset clips each vertical edge independently and keeps horizontal clipping at the paper.
    content.style.overflowY = "visible";
    content.style.clipPath = `inset(${-up}px ${(band.right - paper.right) / scale}px ${-down}px ${(paper.left - band.left) / scale}px)`;
  }
}

/** Read the pixel inset used for the exact-text paint clip, including CSS's 1-4 value shorthand. */
export function insetClipBounds(value: string, box: DOMRect, scale: number):
  { top: number; right: number; bottom: number; left: number } | null {
  const match = /^inset\(((?:-?\d+(?:\.\d+)?px\s*){1,4})\)$/.exec(value);
  if (!match) return null;
  const parts = match[1].trim().split(/\s+/).map(Number.parseFloat);
  const [top, right = top, bottom = top, left = right] = parts;
  return { top: box.top + top * scale, right: box.right - right * scale,
    bottom: box.bottom - bottom * scale, left: box.left + left * scale };
}

/** Only inline ink in a fitting exact-spaced paragraph may use the expanded paint clip. */
export function isPreservedExactLineInk(element: HTMLElement, content: HTMLElement): boolean {
  const paragraph = element.closest<HTMLElement>(`[${EXACT_INK_ATTRIBUTE}]`);
  const view = content.ownerDocument.defaultView;
  if (!paragraph || !view || element === paragraph || !content.contains(paragraph)) return false;
  if (view.getComputedStyle(element).display !== "inline") return false;
  const band = content.getBoundingClientRect();
  const box = paragraph.getBoundingClientRect();
  const clip = insetClipBounds(view.getComputedStyle(content).clipPath, band, band.width / content.offsetWidth || 1);
  const ink = element.getBoundingClientRect();
  return clip !== null && box.top >= band.top - 0.5 && box.bottom <= band.bottom + 0.5
    && ink.top >= clip.top - 0.02 && ink.bottom <= clip.bottom + 0.02;
}

/**
 * Put every exported baseline where Word puts it, to a fraction of a pixel (issue #942).
 *
 * Word sets a line's baseline its natural line height less the font's descent below the line's top: the
 * font's ascent, with any line gap above it. Chromium instead rounds the ascent and descent to whole pixels
 * and floors half the remaining leading, so the baseline lands up to a pixel off; for 11 pt Calibri and 12 pt
 * Arial it is one pixel high. Word also puts a font's line gap above the text where CSS splits it, which moves
 * Arial's baseline a little further. That offset is the same on every line of a paragraph, whatever its position, so
 * it is measured once per paragraph shape in a probe and taken back with relative positioning, which moves the
 * glyphs without changing any line box: pagination and line pitch are untouched.
 *
 * Each direct inline child of a paragraph is moved, composing with a `top` it already has. A child positioned
 * some other way, or holding an image, an SVG or an absolutely positioned element (for which a relative child
 * would become the containing block), is left alone. Run once the page tree is final: the moved boxes count
 * toward a page band's overflow, which the export's clipping check must not see.
 * Returns how many children moved.
 */
export function alignBaselinesToWord(root: Element): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const descents = new Map<string, number>();
  const ratioFor = naturalLineHeightRatios(document);
  const chromiumOffset = baselineOffsets(document);

  /** The font's descent as a fraction of its size, read where rounding is negligible. */
  const descentRatio = (style: CSSStyleDeclaration): number => {
    // The canvas font shorthand takes no percentage stretch (computed styles give one), and ignores a value it
    // cannot parse, so the stretch is left out and an assignment that did not take counts as unmeasurable.
    const font = `${style.fontStyle} ${style.fontWeight} ${PROBE_FONT_SIZE_PX}px ${style.fontFamily}`;
    let ratio = descents.get(font);
    if (ratio === undefined) {
      const context = document.createElement("canvas").getContext("2d");
      ratio = Number.NaN;
      if (context) {
        context.font = font;
        if (context.font.includes(`${PROBE_FONT_SIZE_PX}px`))
          ratio = context.measureText(LATIN_PROBE).fontBoundingBoxDescent / PROBE_FONT_SIZE_PX;
      }
      descents.set(font, ratio);
    }
    return ratio;
  };

  const moves: Array<{ element: HTMLElement; top: number }> = [];
  for (const paragraph of Array.from(root.querySelectorAll(PARAGRAPH_SELECTOR))) {
    if (!(paragraph instanceof view.HTMLElement)) continue;
    const style = view.getComputedStyle(paragraph);
    // Probe with a child that sits on the baseline (not a raised or lowered run), preferably in the
    // paragraph's own font rather than, say, a list marker's.
    const candidates = Array.from(paragraph.children).filter((child): child is HTMLElement =>
      child instanceof view.HTMLElement && (child.textContent ?? "").trim() !== "" && movable(child, view) &&
      view.getComputedStyle(child).display === "inline" && view.getComputedStyle(child).verticalAlign === "baseline");
    const first = candidates.find((child) => view.getComputedStyle(child).fontFamily === style.fontFamily) ??
      candidates[0];
    if (!first) continue;
    const lineHeight = Number.parseFloat(style.lineHeight);
    const fontSize = Number.parseFloat(style.fontSize);
    // Word's placement is recorded for natural line heights (single spacing and auto multiples, which the
    // converter builds on the paragraph's natural height), not for exact or at-least ones.
    if (!(lineHeight > 0) || !(fontSize > 0) ||
      Math.abs(lineHeight - ratioFor(style, LATIN_PROBE) * fontSize) > 0.01) continue;
    const descent = descentRatio(style);
    if (!(descent > 0)) continue;
    const childStyle = view.getComputedStyle(first);
    // Word: the extra height of a multiple goes below the text (issue #908), so the baseline is one natural
    // line height less the descent below the line's top, however tall the line is.
    const shift = (lineHeight - descent * fontSize) - chromiumOffset(style, childStyle);
    // The correction is Chromium's rounding plus the half of any line gap it puts below the text, which Word
    // puts above: a fraction of the line. Half a line or more means the font does not follow this model.
    if (!Number.isFinite(shift) || Math.abs(shift) < 1 / 64 || Math.abs(shift) >= lineHeight / 2) continue;
    for (const child of Array.from(paragraph.children))
      if (child instanceof view.HTMLElement && movable(child, view)) {
        // Compose with the offset the child already has (a run's w:position, or the lift of an auto multiple),
        // which may come from a stylesheet rule rather than its own style attribute.
        const top = Number.parseFloat(view.getComputedStyle(child).top);
        moves.push({ element: child, top: (Number.isFinite(top) ? top : 0) + shift });
      }
  }
  for (const { element, top } of moves) {
    element.style.setProperty("position", "relative");
    element.style.setProperty("top", `${top.toFixed(4)}px`);
  }
  return moves.length;
}

/** Whether relative positioning can move a paragraph child without disturbing anything laid out from it. */
function movable(element: HTMLElement, view: Window & typeof globalThis): boolean {
  const position = view.getComputedStyle(element).position;
  if (position !== "static" && position !== "relative") return false;
  // Inline boxes share the line's baseline; tab stops and leaders are inline-blocks on the same line.
  if (!["inline", "inline-block"].includes(view.getComputedStyle(element).display)) return false;
  if (element.querySelector("img, svg")) return false;
  for (const descendant of Array.from(element.querySelectorAll("*")))
    if (["absolute", "fixed"].includes(view.getComputedStyle(descendant).position)) return false;
  return true;
}

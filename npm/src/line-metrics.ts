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

/** Large-font probe reduces per-component pixel rounding before scaling back to the run's size. */
const PROBE_FONT_SIZE_PX = 10000;
const probeSizes = new WeakMap<Document, number>();

/** Firefox clamps large rendered fonts without changing their computed size (#1424). */
function fontProbeSize(document: Document): number {
  let size = probeSizes.get(document);
  if (size !== undefined) return size;
  const probe = document.createElement("span");
  probe.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;display:inline-block;"
    + `font:${PROBE_FONT_SIZE_PX / 2}px/normal serif;margin:0;padding:0;border:0;`;
  probe.textContent = "Hg";
  (document.body ?? document.documentElement).appendChild(probe);
  const height = probe.getBoundingClientRect().height;
  probe.style.fontSize = `${PROBE_FONT_SIZE_PX}px`;
  size = height > 0 && probe.getBoundingClientRect().height === height ? 1000 : PROBE_FONT_SIZE_PX;
  probe.remove();
  if (height > 0) probeSizes.set(document, size);
  return size;
}

/** Probe text for an element whose own text the primary font covers (Latin, through U+024F). */
const LATIN_PROBE = "Hg";

/** A character the Latin probe's font may not cover (beyond U+024F), other than white space. */
const BEYOND_LATIN = /[^\u0000-\u024f\s]/u;
const BEYOND_LATIN_CHARACTERS = /[^\u0000-\u024f\s]/gu;

/** At most this many distinct characters of an element's own text go into its probe. */
const MAX_PROBE_CHARACTERS = 256;

/**
 * The characters to measure an element's natural line height with, beyond {@link LATIN_PROBE}.
 * Under `normal`, Chromium also grows a line for the fallback fonts its text needs (a CJK run in a
 * Calibri paragraph lays out taller than Latin text), so an element whose own text goes beyond
 * Latin is probed with those characters themselves.
 */
function probeCharacters(element: Element): string[] {
  const characters = new Set<string>();
  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType !== 3) continue;
    // Most text is Latin throughout; test the whole string once before looking at characters.
    const text = node.textContent ?? "";
    if (!BEYOND_LATIN.test(text)) continue;
    for (const [character] of text.matchAll(BEYOND_LATIN_CHARACTERS)) {
      characters.add(character);
      if (characters.size >= MAX_PROBE_CHARACTERS) return Array.from(characters);
    }
  }
  return Array.from(characters);
}

/** The font properties a line-height probe copies from an element's computed style. */
type ProbeFont = Pick<CSSStyleDeclaration, "fontStyle" | "fontWeight" | "fontStretch" | "fontFamily">;

/**
 * A cached measure of a font's natural line height as a fraction of its size, probed at
 * {@link PROBE_FONT_SIZE_PX} with the given text.
 */
function naturalLineHeightRatios(document: Document): (style: ProbeFont, text: string) => number {
  const probeSize = fontProbeSize(document);
  const ratios = new Map<string, number>();
  const ratio = (style: ProbeFont, text: string): number => {
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
      probe.style.fontSize = `${probeSize}px`;
      probe.textContent = text;
      (document.body ?? document.documentElement).appendChild(probe);
      ratio = probe.getBoundingClientRect().height / probeSize;
      probe.remove();
      ratios.set(key, ratio);
    }
    return ratio;
  };
  return ratio;
}

/**
 * Give every element under `root` whose computed `line-height` is `normal` an explicit line height
 * equal to its own font's unrounded natural line height (including any fallback font its own text
 * needs). Elements are collected before any is
 * changed, because an explicit value on a parent would otherwise be inherited by children that
 * should keep sizing their lines from their own font. Returns how many elements changed.
 *
 * The paginator runs this on its staging tree, so every paginated surface lays lines out alike. A
 * caller that needs the heights earlier (the export measures text before it paginates) may run it
 * first: a resolved height is no longer `normal`, so the paginator's pass leaves it alone.
 */
export function applyUnroundedNormalLineHeights(root: Element): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const ratioFor = naturalLineHeightRatios(document);

  // Read every style before probing: each probe changes the document, so a style read after one
  // recomputes the styles of the whole tree.
  const candidates: Array<{ element: HTMLElement; font: ProbeFont; fontSize: number; characters: string[]; floor: boolean }> = [];
  const grids: Array<{ paragraph: HTMLElement; pitch: number }> = [];
  const minimums: HTMLElement[] = [];
  for (const element of [root, ...Array.from(root.querySelectorAll("*"))]) {
    if (!(element instanceof view.HTMLElement)) continue;
    const style = view.getComputedStyle(element);
    const paragraph = element.matches(PARAGRAPH_SELECTOR);
    const minimum = Number.parseFloat(style.getPropertyValue('--docx-min-line-height')) * 4 / 3;
    if (paragraph && minimum > 0) minimums.push(element);
    const pitch = Number.parseFloat(style.getPropertyValue("--docx-line-grid")) * 4 / 3;
    // On a grid, an at-least line height is only the floor of the text's grid lines (#1551):
    // its natural height is measured too, and the minimum stays until the grid pass replaces it.
    const floor = pitch > 0 && Math.abs(Number.parseFloat(style.lineHeight) - minimum) < .01;
    if (style.lineHeight !== "normal" && !floor) continue;
    const fontSize = Number.parseFloat(style.fontSize);
    if (!(fontSize > 0)) continue;
    const { fontStyle, fontWeight, fontStretch, fontFamily } = style;
    candidates.push({ element, font: { fontStyle, fontWeight, fontStretch, fontFamily }, fontSize,
      characters: probeCharacters(element), floor });
    if (paragraph && pitch > 0) grids.push({ paragraph: element, pitch });
  }
  const targets: Array<{ element: HTMLElement; lineHeight: number; floor: boolean }> = [];
  for (const { element, font, fontSize, characters, floor } of candidates) {
    // Fallback fonts are chosen per run of text, so the characters are probed together, as they
    // are laid out; a character probed alone can fall back to a different font.
    const ratio = ratioFor(font, LATIN_PROBE + [...characters].sort().join(""));
    if (ratio > 0) targets.push({ element, lineHeight: ratio * fontSize, floor });
  }
  const resolved = targets.filter(target => !target.floor);
  for (const { element, lineHeight } of resolved) {
    element.style.setProperty("line-height", `${lineHeight.toFixed(4)}px`);
  }
  let restored = 0;
  if (grids.length || minimums.length) {
    const naturalHeights = new Map(targets.map(({ element, lineHeight }) => [element, lineHeight]));
    if (grids.length) applyParagraphLineGrids(grids, naturalHeights);
    if (minimums.length) restored = applyUniformLineMinimums(minimums, ratioFor, naturalHeights);
  }
  positionRubyAnnotations(root);
  return resolved.length + restored;
}

const faceCanvases = new WeakMap<Document, CanvasRenderingContext2D | null>();

/**
 * Whether a run's face lays a line out as its paragraph's does: the same size and stretch, and the same
 * family, style and weight or the same ascent and descent. Word lines up a bold word with its text (#1551).
 */
function sameLineFace(document: Document, run: CSSStyleDeclaration, paragraph: CSSStyleDeclaration): boolean {
  if (run.fontSize !== paragraph.fontSize || run.fontStretch !== paragraph.fontStretch) return false;
  if (run.fontFamily === paragraph.fontFamily && run.fontStyle === paragraph.fontStyle
    && run.fontWeight === paragraph.fontWeight) return true;
  if (!faceCanvases.has(document)) faceCanvases.set(document, document.createElement("canvas").getContext("2d"));
  const canvas = faceCanvases.get(document), size = fontProbeSize(document);
  if (!canvas) return false;
  const extent = (face: CSSStyleDeclaration) => {
    // A face the canvas cannot parse leaves the font before it in place.
    canvas.font = "1px serif";
    canvas.font = `${face.fontStyle} ${face.fontWeight} ${size}px ${face.fontFamily}`;
    const { fontBoundingBoxAscent, fontBoundingBoxDescent } = canvas.measureText(LATIN_PROBE);
    return canvas.font.includes(`${size}px`) ? `${fontBoundingBoxAscent}|${fontBoundingBoxDescent}` : undefined;
  };
  const own = extent(paragraph);
  return own !== undefined && own === extent(run);
}

/** Ordinary inline runs whose metrics can be represented by their paragraph's font. */
function uniformParagraphRuns(paragraph: HTMLElement, allowStructuralWrappers = false,
  positioned?: (run: CSSStyleDeclaration) => boolean): HTMLElement[] | undefined {
  const view = paragraph.ownerDocument.defaultView!;
  const style = view.getComputedStyle(paragraph);
  if (style.direction !== "ltr" || style.writingMode !== "horizontal-tb") return;
  const children = Array.from(paragraph.children);
  const marker = allowStructuralWrappers && children[0]?.getAttribute('data-list-marker') === 'true'
    && view.getComputedStyle(children[0]).display === 'inline-flex' ? children[0] : null;
  // Paginated endnotes retain their note identity in one inline wrapper around
  // each paragraph's runs. Its own font and line height must move with them.
  const note = allowStructuralWrappers && paragraph.matches('section.endnotes > ol > li > p')
    && children.length === 1 && children[0].getAttribute('data-source-anchor-id')?.startsWith('en:en:') ? children[0] : null;
  // Tracked changes shown as markup wrap their runs in <ins> or <del>, which Word lays out as plain text (#1551).
  const wrappers = children.filter(child => child === note || allowStructuralWrappers && /^(ins|del)$/.test(child.localName));
  const runs = children.flatMap(child => child === marker ? [child, ...Array.from(child.querySelectorAll('*'))]
    : wrappers.includes(child) ? [child, ...Array.from(child.children)] : [child]);
  if (!children.length || [paragraph, ...wrappers].some(parent =>
    Array.from(parent.childNodes).some(n => n.nodeType === 3 && n.textContent?.trim()))
    || runs.some(child => {
      const inMarker = marker?.contains(child), wrapper = wrappers.includes(child);
      if (!(child instanceof view.HTMLElement) || (!wrapper && child.localName !== "span")
        // A condensed justified line (word-text.ts) narrows spaces inside its run, nothing else.
        || (!inMarker && ((!wrapper && child.querySelector(':scope > :not([data-docx-condensed-line])'))
          || !child.textContent?.trim()))) return true;
      const run = view.getComputedStyle(child);
      // Converter list prefixes are an inline flex box with two flex items: the
      // number and its tab suffix. Their uniform metrics must move together.
      const display = child === marker ? 'inline-flex' : child.parentElement === marker ? 'block'
        : inMarker && child.hasAttribute('data-docx-tab') ? 'inline-block' : 'inline';
      return !sameLineFace(paragraph.ownerDocument, run, style) || run.display !== display
        || (run.position !== "static" && !positioned?.(run)) || run.verticalAlign !== "baseline" || run.transform !== "none";
    })) return;
  return runs as HTMLElement[];
}

/** Uniform lines take the greater of their natural height and their authored minimum. */
function applyUniformLineMinimums(paragraphs: HTMLElement[], ratioFor: (style: ProbeFont, text: string) => number,
  naturalHeights: Map<HTMLElement, number>): number {
  const view = paragraphs[0].ownerDocument.defaultView!;
  let changed = 0;
  for (const paragraph of paragraphs) {
    if (paragraph.closest('a, aside, [data-comment-id], '
      + '[data-docx-textbox], [data-docx-text-frame], [data-docx-drawing-anchor]')
      || (paragraph.closest('ol, ul, li') && !paragraph.matches('section.endnotes > ol > li > p'))
      || BEYOND_LATIN.test(paragraph.textContent ?? '')) continue;
    const style = view.getComputedStyle(paragraph);
    const minimum = Number.parseFloat(style.getPropertyValue('--docx-min-line-height')) * 4 / 3;
    if (!(minimum > 0)) continue;
    const natural = ratioFor(style, LATIN_PROBE) * Number.parseFloat(style.fontSize);
    const height = Number.parseFloat(style.lineHeight);
    const restoreNatural = natural > minimum && Math.abs(height - minimum) < .01;
    // The converter omits fixed heights below 14pt. Recover only its normal-height
    // fallback, retaining explicit CSS overrides even when they equal the natural height.
    const applyMinimum = minimum < 14 * 4 / 3 && minimum > natural
      && naturalHeights.has(paragraph) && Math.abs(height - natural) < .01;
    if (!restoreNatural && !applyMinimum) continue;
    const runs = uniformParagraphRuns(paragraph, true);
    if (!runs || runs.some(run => view.getComputedStyle(run).lineHeight !== style.lineHeight
      || (applyMinimum && !naturalHeights.has(run)))) continue;
    for (const element of [paragraph, ...runs]) {
      element.style.lineHeight = `${Math.max(natural, minimum).toFixed(4)}px`;
      if (!naturalHeights.has(element)) changed++;
    }
    // Only a below-natural minimum is equivalent to single spacing. Enlarged
    // lines retain their authored minimum for the final baseline pass.
    if (restoreNatural) paragraph.style.setProperty('--docx-auto-line-spacing', '1');
  }
  return changed;
}

/** The converter's direction mark after a `<br>`: it holds an otherwise empty line's font. */
const BREAK_CARRIER = /^\u200e+$/u;

/** The converter's empty box holding a mapped symbol marker's original ascent (#1567). */
const GLYPH_ASCENT = 'span[data-docx-glyph-ascent]';

/**
 * Give mixed ordinary text one shared ascent, descent and spacing floor per visual line (#1565).
 * Word's line is its text's largest ascent plus largest descent, and an automatic multiple of that.
 * A generated list marker only raises its own line's ascent: Word adds neither its descent nor a
 * multiple of it. A break's direction mark sizes only a line that holds nothing else, so an empty
 * line made by `w:br` takes the break's own font.
 */
export function prepareMixedLineHeights(root: HTMLElement, checkpoint?: () => void): number {
  const document = root.ownerDocument, view = document.defaultView!;
  const ratioFor = naturalLineHeightRatios(document), probeSize = fontProbeSize(document);
  const canvas = document.createElement('canvas').getContext('2d');
  if (!canvas) return 0;
  const fontMetrics = new Map<string, { ascent: number; descent: number; offset: number; native: number }>();
  const measure = (font: ProbeFont, size: number) => {
    const key = `${font.fontStyle} ${font.fontWeight} ${font.fontStretch} ${size}px ${font.fontFamily}`;
    let result = fontMetrics.get(key);
    if (result) return result;
    const height = ratioFor(font, LATIN_PROBE) * size;
    canvas.font = `${font.fontStyle} ${font.fontWeight} ${probeSize}px ${font.fontFamily}`;
    const descent = canvas.measureText(LATIN_PROBE).fontBoundingBoxDescent / probeSize * size;
    const probe = document.createElement('span'), marker = document.createElement('span');
    probe.style.cssText = 'position:absolute;left:-100000px;visibility:hidden;display:inline-block;white-space:nowrap;line-height:normal';
    Object.assign(probe.style, { fontFamily: font.fontFamily, fontStyle: font.fontStyle,
      fontWeight: font.fontWeight, fontStretch: font.fontStretch, fontSize: `${size}px` });
    probe.textContent = LATIN_PROBE;
    marker.style.cssText = 'display:inline-block;width:0;height:0;padding:0;margin:0;border:0;vertical-align:baseline';
    probe.append(marker); document.body.append(probe);
    const range = document.createRange(); range.selectNodeContents(probe.firstChild!);
    result = { ascent: height - descent, descent, native: probe.getBoundingClientRect().height,
      offset: marker.getBoundingClientRect().top - range.getBoundingClientRect().top };
    probe.remove(); fontMetrics.set(key, result); return result;
  };
  const paragraphs = root.matches(PARAGRAPH_SELECTOR) ? [root]
    : Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPH_SELECTOR));
  let prepared = 0;
  for (const paragraph of paragraphs) {
    checkpoint?.();
    const owned = paragraph.hasAttribute('data-docx-mixed-lines');
    const width = paragraph.getBoundingClientRect().width;
    if (owned && Number(paragraph.dataset.docxMixedLines) === width) continue;
    if (!(width > 0) || paragraph.closest('table, ol, ul, li, aside, [data-hf-type], '
      + '#pagination-hf-registry, #pagination-footnote-registry, #pagination-comment-margin-registry, '
      + '[data-footnote-id], [data-comment-id], section.endnotes, '
      + '[data-docx-textbox], [data-docx-text-frame], [data-docx-drawing-anchor]')) continue;
    const style = view.getComputedStyle(paragraph);
    const minimum = Number.parseFloat(style.getPropertyValue('--docx-min-line-height')) * 4 / 3 || 0;
    const multiple = Number(style.getPropertyValue('--docx-auto-line-spacing')) || 1;
    const pitch = Number.parseFloat(style.getPropertyValue('--docx-line-grid')) * 4 / 3 || 0;
    const section = paragraph.closest<HTMLElement>('[data-section-index]');
    if (Number(section?.dataset.cols ?? 1) > 1 || section?.querySelector(
      '[data-docx-text-frame], [data-docx-wrap-text], [data-docx-edge-picture]')) continue;
    if (style.direction !== 'ltr' || style.writingMode !== 'horizontal-tb'
      || multiple < 1
      || Number.parseFloat(style.getPropertyValue('--docx-exact-line-height')) > 0
      || Array.from(paragraph.childNodes).some(n => n.nodeType === 3 && n.textContent?.trim())) continue;
    const marker = paragraph.firstElementChild?.getAttribute('data-list-marker') === 'true'
      ? paragraph.firstElementChild : null;
    // A bullet's glyph may lie beyond Latin: only its font's ascent counts, and its own line height is zero.
    if (/[^\u0020-\u024f\u200e]/u.test((paragraph.textContent ?? '').slice(marker?.textContent?.length ?? 0))) continue;
    if (paragraph.querySelector('[data-field], [data-footnote-id], [data-endnote-id], [data-comment-id], '
      + '[data-docx-fit-text-group], [data-docx-word-line], [data-docx-ruby]')) continue;
    // The converter's number and suffix tab occupy two baseline-aligned flex items.
    if (marker && (marker.children.length !== 2 || marker.children[0].children.length !== 1
      || !marker.children[0].firstElementChild?.matches('span[data-list-marker="true"]')
      // A mapped symbol glyph may carry the converter's ascent box for the font it replaced.
      || Array.from(marker.children[0].firstElementChild.children).some(child => !child.matches(GLYPH_ASCENT))
      || marker.children[1].children.length !== 1
      || !marker.children[1].firstElementChild?.matches('span[data-docx-tab]')
      || marker.children[1].firstElementChild.childNodes.length)) continue;
    // A replaced symbol font's ascent differs from the glyph's fallback face even when every run matches.
    const glyphAscent = !!marker?.querySelector(GLYPH_ASCENT);
    const runs = Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
      .filter(run => !run.hasAttribute('data-docx-line-metric') && !run.matches(GLYPH_ASCENT));
    const fonts = runs.map(run => {
      const css = view.getComputedStyle(run);
      return { run, size: Number.parseFloat(css.fontSize), family: css.fontFamily,
        weight: css.fontWeight, style: css.fontStyle, stretch: css.fontStretch,
        display: css.display, align: css.verticalAlign, position: css.position, top: css.top,
        height: css.lineHeight, transform: css.transform,
        sides: [css.left, css.right], bottom: css.bottom, float: css.cssFloat };
    });
    const paragraphFont = owned ? undefined : measure(style, Number.parseFloat(style.fontSize));
    // Converter line breaks stay inside their run, followed by the run's direction mark and, in an
    // indented paragraph, an empty span that carries the following line's horizontal offset.
    const contained = (child: Element) => child.matches('br[data-docx-break]')
      || child.localName === 'span' && !child.childNodes.length && !!child.previousElementSibling?.matches('br[data-docx-break]')
      || owned && child.hasAttribute('data-docx-line-metric');
    if (!runs.length || paragraph.querySelector('*:not(span):not(br[data-docx-break])')
      || Array.from(paragraph.children).some(child => child.localName === 'br') || fonts.some(f =>
      !(f.size > 0) || f.transform !== 'none' || !['static', 'relative'].includes(f.position)
      || f.float !== 'none' || f.sides.some(side => !['auto', '0px'].includes(side))
      || !['auto', '0px'].includes(f.bottom) && !(f.position === 'relative'
        && Math.abs(Number.parseFloat(f.bottom) + Number.parseFloat(f.top)) < .02)
      || f.position === 'relative' && !(multiple > 1
        && Math.abs(Number.parseFloat(f.top) + Number.parseFloat(f.height) * (1 - 1 / multiple) / 2) < .02)
      || !owned && f.run.style.lineHeight && !['normal', 'calc(1lh*var(--docx-auto-line-spacing))']
        .includes(f.run.style.lineHeight.replace(/\s/g, ''))
      || !owned && f.height !== 'normal' && Math.abs(Number.parseFloat(f.height)
        - (minimum >= 14 * 4 / 3 ? minimum : paragraphFont!.native * multiple)) > .02
      || f.display !== (f.run === marker ? 'inline-flex' : f.run.parentElement === marker ? 'block'
        : marker?.contains(f.run) && f.run.hasAttribute('data-docx-tab') ? 'inline-block' : 'inline')
      || f.align !== 'baseline'
      || !marker?.contains(f.run) && Array.from(f.run.children).some(child => !contained(child)))
      || !owned && style.lineHeight !== 'normal'
        && !(minimum >= 14 * 4 / 3 && Math.abs(Number.parseFloat(style.lineHeight) - minimum) < .02)
      || !owned && !glyphAscent && new Set(fonts.map(f => `${f.size}|${f.family}|${f.weight}|${f.style}`)).size < 2) continue;
    const metrics = new Map(fonts.map(f => [f.run, measure({ fontFamily: f.family,
      fontWeight: f.weight, fontStyle: f.style, fontStretch: f.stretch }, f.size)]));
    // Styling alone does not require vertical line ownership. Equal metrics keep
    // the ordinary flow available to the precise horizontal text fitter.
    if (!owned && !glyphAscent && new Set([...metrics.values()].map(m => `${m.ascent}|${m.descent}`)).size < 2) continue;
    const before = paragraph.cloneNode(true) as HTMLElement;
    const sourceText = paragraph.textContent;
    const words = () => {
      const pieces: Array<{ node: Text; start: number; end: number; left: number; width: number; top: number }> = [];
      const walker = document.createTreeWalker(paragraph, view.NodeFilter.SHOW_TEXT);
      for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null)
        // The marker is measured by its font alone: its glyph may render in a fallback face.
        for (const match of marker?.contains(node) ? [] : node.data.matchAll(/\S+/gu)) {
          const range = document.createRange();
          range.setStart(node, match.index); range.setEnd(node, match.index + match[0].length);
          if (range.getClientRects().length !== 1) return undefined;
          const box = range.getBoundingClientRect();
          pieces.push({ node, start: match.index, end: match.index + match[0].length,
            left: box.left, width: box.width, top: box.top });
        }
      return pieces;
    };
    let pristine = owned ? undefined : words();
    if (!owned && !pristine?.length) continue;
    for (const strut of Array.from(paragraph.querySelectorAll('[data-docx-line-metric]'))) strut.remove();
    // Removing old metrics reunites the original shaping runs before a width refit.
    paragraph.normalize();
    if (owned) pristine = words();
    for (const f of fonts) {
      const font = { fontFamily: f.family, fontWeight: f.weight, fontStyle: f.style, fontStretch: f.stretch };
      const metric = metrics.get(f.run)!;
      Object.assign(f.run.style, font, { fontSize: `${f.size}px`, lineHeight: `${metric.ascent + metric.descent}px`, position: 'static', top: 'auto' });
    }
    // Resolve font-relative paragraph geometry before removing its global strut.
    for (const property of ['text-indent', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right',
      'padding-top', 'padding-bottom', 'padding-left', 'padding-right'])
      paragraph.style.setProperty(property, style.getPropertyValue(property));
    paragraph.style.fontSize = '0'; paragraph.style.lineHeight = '0';
    type Piece = { node: Text; start: number; end: number; left: number; width: number; run: HTMLElement };
    // A marker or a break's direction mark keeps its own maxima: neither sizes a line of text.
    type Extent = { ascent: number; descent: number };
    type Line = Extent & { pieces: Piece[]; baseline: number; marker: number; carrier: Extent };
    const lines: Line[] = [];
    let simple = true;
    const walker = document.createTreeWalker(paragraph, view.NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode() as Text | null; node; node = walker.nextNode() as Text | null) {
      const run = node.parentElement!, metric = metrics.get(run);
      if (!metric || marker?.contains(run)) continue;
      for (const match of node.data.matchAll(/\S+/gu)) {
        checkpoint?.();
        const range = document.createRange();
        range.setStart(node, match.index); range.setEnd(node, match.index + match[0].length);
        const box = range.getBoundingClientRect();
        if (range.getClientRects().length !== 1) simple = false;
        const baseline = box.top + metric.offset;
        if (!lines.length || Math.abs(baseline - lines.at(-1)!.baseline) > .05)
          lines.push({ pieces: [], ascent: 0, descent: 0, baseline, marker: 0, carrier: { ascent: 0, descent: 0 } });
        const line = lines.at(-1)!;
        line.pieces.push({ node, start: match.index, end: match.index + match[0].length, left: box.left, width: box.width, run });
        const extent = BREAK_CARRIER.test(match[0]) ? line.carrier : line;
        extent.ascent = Math.max(extent.ascent, metric.ascent); extent.descent = Math.max(extent.descent, metric.descent);
      }
    }
    for (const run of runs) run.style.lineHeight = '0';
    // The marker opens the first line, which it raises by its font's ascent alone.
    const markerText = document.createTreeWalker(marker ?? paragraph, view.NodeFilter.SHOW_TEXT);
    for (let node = marker && markerText.nextNode(); node && lines.length; node = markerText.nextNode()) {
      if (!node.textContent?.trim()) continue;
      // A mapped symbol glyph counts with the replaced font's ascent, not its fallback face's (#1567).
      const span = node.parentElement!, box = span.querySelector<HTMLElement>(GLYPH_ASCENT);
      lines[0].marker = Math.max(lines[0].marker, box ? box.getBoundingClientRect().height : metrics.get(span)?.ascent ?? 0);
    }
    // A line holding only a break's direction mark takes that break's font.
    for (const line of lines) if (!line.descent) Object.assign(line, line.carrier);
    // Lengths keep only the six significant digits CSS serializes, so a saved page lays out
    // alike: 12.265638px reopened as 12.2656px is one 1/64px layout unit shorter.
    const css = (value: number) => Number(value.toPrecision(6));
    const inserted: HTMLElement[] = [];
    for (const line of [...lines].reverse()) {
      const last = line.pieces.at(-1)!;
      const text = line.ascent + line.descent;
      // The marker raises its line's ascent outside any multiple of the text's own height.
      const natural = css(Math.max(line.ascent, line.marker) + line.descent);
      const spaced = pitch ? gridLineHeight(natural, pitch, multiple) : natural + text * (multiple - 1);
      const height = css(Math.max(minimum, spaced));
      // Word puts minimum-height surplus above the baseline; automatic multiples
      // put their surplus below it, and a grid line centres its text in its grid lines
      // (#1551), surplus over them above. All retain the font's natural descent.
      line.descent = css(line.descent);
      const descent = css(line.descent + (pitch ? (spaced - natural) / 2 : minimum ? 0 : height - natural));
      const range = document.createRange(); range.setStart(last.node, last.end); range.collapse(true);
      const strut = document.createElement('span');
      strut.dataset.docxLineMetric = pitch ? 'grid' : minimum ? 'minimum' : 'auto';
      if (pitch) strut.dataset.docxLineGrid = String(css(spaced));
      strut.dataset.docxLineNatural = String(natural); strut.dataset.docxLineDescent = String(line.descent);
      strut.setAttribute('aria-hidden', 'true');
      strut.style.cssText = `display:inline-block;position:static;float:none;transform:none;width:0;height:${height}px;font-size:0;line-height:0;vertical-align:${-descent}px;margin:0;padding:0;border:0`;
      range.insertNode(strut); inserted.unshift(strut);
    }
    // Empty inline boxes may wrap separately. Require each metric to remain on
    // its own source line, and never accept a changed horizontal layout.
    const observed = (words() ?? []).map(piece => ({ ...piece,
      baseline: piece.top + metrics.get(piece.node.parentElement!)!.offset }));
    let nextTop = paragraph.getBoundingClientRect().top + Number.parseFloat(style.paddingTop) + Number.parseFloat(style.borderTopWidth);
    const ownsLines = inserted.every((strut, index) => {
      const box = strut.getBoundingClientRect(), expected = Number.parseFloat(strut.style.height);
      const valid = Math.abs(box.top - nextTop) < .05 + index / 64 && Math.abs(box.height - expected) < 1 / 32;
      nextTop += expected;
      return valid;
    });
    let pieceIndex = 0;
    const membership = lines.every((line, index) => {
      const strut = inserted[index], box = strut.getBoundingClientRect();
      const baseline = box.bottom + Number.parseFloat(strut.style.verticalAlign);
      return line.pieces.every(() => Math.abs(observed[pieceIndex++]?.baseline - baseline) < .05);
    });
    const markerBox = marker?.getBoundingClientRect(), firstLine = inserted[0]?.getBoundingClientRect();
    const markerOnFirstLine = !markerBox || !!firstLine
      && markerBox.top >= firstLine.top - .05 && markerBox.bottom <= firstLine.bottom + .05;
    const valid = pristine && simple && lines.length > 0 && lines.every(line => line.descent > 0) && markerOnFirstLine && ownsLines && membership && sourceText === paragraph.textContent && observed.length === pristine.length
      && pristine.every((piece, index) => Math.abs(piece.left - observed[index].left) < 1 / 32
        && Math.abs(piece.width - observed[index].width) < 1 / 32)
      && inserted.every((strut, index) => index === 0 || strut.getBoundingClientRect().top > inserted[index - 1].getBoundingClientRect().top);
    if (!valid) {
      paragraph.replaceChildren(...Array.from(before.childNodes));
      const saved = before.getAttribute('style');
      if (saved === null) paragraph.removeAttribute('style'); else paragraph.setAttribute('style', saved);
      continue;
    }
    paragraph.dataset.docxMixedLines = String(width);
    prepared++;
  }
  return prepared;
}

/** Only our empty zero-width metric boxes may bypass ordinary inline layout guards. */
export function isMixedLineMetric(element: HTMLElement): boolean {
  if (!element.closest('[data-docx-mixed-lines]') || element.localName !== 'span'
    || element.childNodes.length || element.id || element.hasAttribute('data-anchor')
    || element.hasAttribute('data-source-anchor-id') || !['auto', 'minimum', 'grid'].includes(element.dataset.docxLineMetric ?? '')) return false;
  const style = element.isConnected ? element.ownerDocument.defaultView!.getComputedStyle(element) : element.style;
  const natural = Number(element.dataset.docxLineNatural), descent = Number(element.dataset.docxLineDescent);
  const height = Number.parseFloat(style.height), kind = element.dataset.docxLineMetric;
  // The surplus below the text: an automatic multiple's, or half of a grid line's below its grid lines.
  const grid = kind === 'grid' ? Number(element.dataset.docxLineGrid) : natural;
  const extra = kind === 'auto' ? height - natural : (grid - natural) / 2;
  return natural > 0 && descent > 0 && descent < natural && Number.isFinite(natural)
    && grid >= natural && grid < height + .01
    && style.display === 'inline-block' && style.position === 'static' && style.transform === 'none'
    && style.float === 'none' && style.width === '0px' && height > 0 && Number.isFinite(height)
    && Math.abs(Number.parseFloat(style.verticalAlign) + descent + extra) < .01
    && ['margin-top', 'margin-right', 'margin-bottom', 'margin-left', 'padding-top', 'padding-right',
      'padding-bottom', 'padding-left', 'border-top-width', 'border-right-width', 'border-bottom-width', 'border-left-width']
      .every(property => Number.parseFloat(style.getPropertyValue(property)) === 0);
}

/**
 * A line's height on its section's grid. Word 2024 gives a visual line the whole grid slots its natural
 * height needs or its multiple of one pitch, whichever is taller (#1551): 1.5 lines on an 18pt pitch is
 * 27pt, on a 12pt pitch 24pt. An at-least minimum taller than those slots is the line instead.
 */
function gridLineHeight(natural: number, pitch: number, multiple: number): number {
  return Math.max(Math.ceil(natural / pitch), multiple) * pitch;
}

/**
 * Size uniform automatic and at-least lines on their section's grid, retaining their natural baseline.
 * A minimum taller than the grid lines is the line instead, which the converter already set.
 */
function applyParagraphLineGrids(grids: Array<{ paragraph: HTMLElement; pitch: number }>, naturalHeights: Map<HTMLElement, number>): void {
  const view = grids[0].paragraph.ownerDocument.defaultView!;
  for (const { paragraph, pitch } of grids) {
    const natural = naturalHeights.get(paragraph);
    if (!natural) continue;
    const style = view.getComputedStyle(paragraph);
    const multiple = Number(style.getPropertyValue('--docx-auto-line-spacing')) || 1;
    const height = gridLineHeight(natural, pitch, multiple);
    if (Number.parseFloat(style.getPropertyValue('--docx-min-line-height')) * 4 / 3 > height) continue;
    // The converter's own multiple: m natural lines, its text lifted to their top. The grid replaces both.
    const multiplied = (run: CSSStyleDeclaration) => multiple !== 1 && run.position === 'relative'
      && Math.abs(Number.parseFloat(run.lineHeight) - natural * multiple) < .01
      && Math.abs(Number.parseFloat(run.top) - Math.min(0, natural * (1 - multiple) / 2)) < .01;
    const runs = uniformParagraphRuns(paragraph, true, multiplied);
    // A tracked change's wrapper carries the line of the runs inside it and moves them with it.
    if (!runs || runs.some(run => run.parentElement === paragraph
      ? naturalHeights.get(run) !== natural && !multiplied(view.getComputedStyle(run))
      : view.getComputedStyle(run).lineHeight !== view.getComputedStyle(run.parentElement!).lineHeight)) continue;
    // The strut stays natural, also where the converter set an at-least minimum.
    paragraph.style.lineHeight = `${natural.toFixed(4)}px`;
    for (const child of runs) {
      if (child.parentElement !== paragraph) continue;
      // The paragraph strut remains natural, and the run's taller line box centres its ink
      // in the grid lines, as Word does (#1551). That replaces a multiple's lift to their top.
      child.style.lineHeight = `${height}px`;
      child.style.position = "relative";
      child.style.top = "0px";
      child.style.setProperty("--docx-word-native-height", `${natural}px`);
      child.style.setProperty("--docx-word-grid-height", `${height}px`);
    }
  }
}

/**
 * How far below the top of its line Word puts a grid run's natural line: it centres the text in its grid
 * lines (#1551), as the run's line box does. Zero for a run the grid pass has not sized.
 */
export function gridTextOffset(run: CSSStyleDeclaration): number {
  const grid = Number.parseFloat(run.getPropertyValue("--docx-word-grid-height"));
  const native = Number.parseFloat(run.getPropertyValue("--docx-word-native-height"));
  return grid > native ? (grid - native) / 2 : 0;
}

/** Word's ruby raise is a baseline distance. Keep native horizontal ruby layout, moving only
 * its annotation and reserving the same signed change in the surrounding line's height. */
function positionRubyAnnotations(root: Element): void {
  const document = root.ownerDocument, view = document.defaultView!;
  for (const wrapper of Array.from(root.querySelectorAll<HTMLElement>('[data-docx-ruby="true"]'))) {
    if (wrapper.dataset.docxRubyPositioned === "true") continue;
    const raise = Number(wrapper.dataset.docxRubyRaise);
    const annotation = wrapper.querySelector<HTMLElement>("ruby > rt");
    if (!annotation || !Number.isFinite(raise) || raise < 0) continue;
    // Notes can still be in a hidden registry. Copy resolved presentation into a measurable
    // clone, including every styled run, rather than reading their zero-sized rectangles.
    const clone = wrapper.cloneNode(true) as HTMLElement;
    const originals = [wrapper, ...Array.from(wrapper.querySelectorAll<HTMLElement>("*"))];
    const copies = [clone, ...Array.from(clone.querySelectorAll<HTMLElement>("*"))];
    originals.forEach((element, index) => {
      const style = view.getComputedStyle(element);
      for (const property of ["font-family", "font-size", "font-style", "font-weight", "font-stretch",
        "line-height", "letter-spacing", "word-spacing", "white-space", "direction"])
        copies[index].style.setProperty(property, style.getPropertyValue(property));
    });
    const base = clone.querySelector<HTMLElement>("[data-docx-ruby-base]");
    const rt = clone.querySelector<HTMLElement>("ruby > rt");
    if (!base || !rt) continue;
    const host = document.createElement("div");
    host.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;width:max-content;";
    host.append(clone);
    (document.body ?? document.documentElement).append(host);
    const marker = () => {
      const span = document.createElement("span");
      span.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0";
      return span;
    };
    const baseMarker = marker(), annotationMarker = marker();
    base.append(baseMarker); rt.append(annotationMarker);
    const shift = baseMarker.getBoundingClientRect().top - annotationMarker.getBoundingClientRect().top - raise * 96 / 72;
    host.remove();
    annotation.style.position = "relative";
    annotation.style.top = `${shift}px`;
    wrapper.style.marginTop = `${-shift}px`;
    wrapper.dataset.docxRubyPositioned = "true";
  }
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
    // A paginated note's number is rendered inside its first paragraph, as Word's reference-mark
    // run is (issue #1130), so it moves with that paragraph's line like any other exact run.
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

/** Read a rectangular paint clip: a pixel inset or an exact-row rectangle with zero-area spikes. */
export function insetClipBounds(value: string, box: DOMRect, scale: number):
  { top: number; right: number; bottom: number; left: number } | null {
  const n = '-?\\d+(?:\\.\\d+)?';
  const rectangle = new RegExp(`^path\\(["']M\\s*(${n})[ ,]+(${n})\\s*H\\s*(${n})\\s*V\\s*(${n})\\s*H\\s*(${n})\\s*Z`
    + `(?:\\s*M\\s*${n}[ ,]+${n}\\s*[HV]\\s*${n}){4}["']\\)$`).exec(value);
  if (rectangle && rectangle[1] === rectangle[5]) {
    const [, left, top, right, bottom] = rectangle.map(Number);
    return { left: box.left + left * scale, top: box.top + top * scale,
      right: box.left + right * scale, bottom: box.top + bottom * scale };
  }
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

/** Elements whose box, padding and border included, take part in line layout though they compute `display: inline`. */
const REPLACED_ELEMENTS = new Set(["img", "svg", "canvas", "video", "object", "embed", "iframe", "input", "textarea", "select"]);

/**
 * Measure each descendant's content bottom once for this layout snapshot (#1271, #1493).
 * Inline padding/borders can paint beyond a fitting line, including through zero-height
 * wrappers. Measure nested text/children without those decorations; atomic boxes keep
 * their full bounds, but a mixed line's metric box only sizes its line and paints nothing.
 * Unboxed generated content conservatively uses its boxed ancestor.
 */
export function measureLaidOutBottoms(band: HTMLElement): Map<Element, number> {
  const bottoms = new Map<Element, number>();
  const view = band.ownerDocument.defaultView;
  const scale = band.getBoundingClientRect().width / band.offsetWidth || 1;
  const range = band.ownerDocument.createRange();
  const visit = (element: Element, ancestorBottom: number): number => {
    if (element.hasAttribute("data-docx-line-metric") && isMixedLineMetric(element as HTMLElement)) return -Infinity;
    const style = view?.getComputedStyle(element);
    const transparent = style?.display === "contents";
    const inline = style?.display === "inline" && !REPLACED_ELEMENTS.has(element.localName);
    const boxBottom = transparent ? ancestorBottom : element.getBoundingClientRect().bottom
      - (inline ? (Number.parseFloat(style.paddingBottom) + Number.parseFloat(style.borderBottomWidth)) * scale : 0);
    const measureChildren = transparent || (inline && element.children.length > 0);
    const generated = measureChildren && ["::before", "::after"].some(pseudo =>
      !["none", "normal"].includes(view!.getComputedStyle(element, pseudo).content));
    let contentBottom = -Infinity;
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === 1) contentBottom = Math.max(contentBottom, visit(child as Element, boxBottom));
      else if (measureChildren && child.nodeType === 3) {
        range.selectNodeContents(child);
        for (const rect of Array.from(range.getClientRects())) contentBottom = Math.max(contentBottom, rect.bottom);
      }
    }
    const bottom = measureChildren && !generated && Number.isFinite(contentBottom) ? contentBottom : boxBottom;
    bottoms.set(element, bottom);
    return bottom;
  };
  const boundary = band.getBoundingClientRect().bottom;
  for (const element of Array.from(band.children)) visit(element, boundary);
  return bottoms;
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
 * some other way, or holding an image, SVG, MathML or an absolutely positioned element (for which a relative child
 * would become the containing block), is left alone. Run once the page tree is final: the moved boxes count
 * toward a page band's overflow, which the export's clipping check must not see.
 * Returns how many children moved.
 */
export function alignBaselinesToWord(root: Element): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const descents = new Map<string, number>();
  const probeSize = fontProbeSize(document);
  const ratioFor = naturalLineHeightRatios(document);
  const chromiumOffset = baselineOffsets(document);

  /** The font's descent as a fraction of its size, read where rounding is negligible. */
  const descentRatio = (style: CSSStyleDeclaration): number => {
    // The canvas font shorthand takes no percentage stretch (computed styles give one), and ignores a value it
    // cannot parse, so the stretch is left out and an assignment that did not take counts as unmeasurable.
    const font = `${style.fontStyle} ${style.fontWeight} ${probeSize}px ${style.fontFamily}`;
    let ratio = descents.get(font);
    if (ratio === undefined) {
      const context = document.createElement("canvas").getContext("2d");
      ratio = Number.NaN;
      if (context) {
        context.font = font;
        if (context.font.includes(`${probeSize}px`))
          ratio = context.measureText(LATIN_PROBE).fontBoundingBoxDescent / probeSize;
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
    const natural = ratioFor(style, LATIN_PROBE) * fontSize;
    const minimum = Number.parseFloat(style.getPropertyValue('--docx-min-line-height')) * 4 / 3;
    // The original Word 16pt-on-10pt capture places the baseline at H - descent,
    // putting the enlarged minimum's surplus above the text. Admit only uniform
    // authored minima, retaining explicit overrides and exact/mixed-run handling.
    const minimumRuns = minimum >= natural && Math.abs(lineHeight - minimum) < .01
      && !Number.parseFloat(style.getPropertyValue('--docx-exact-line-height'))
      && !BEYOND_LATIN.test(paragraph.textContent ?? '')
      ? uniformParagraphRuns(paragraph, true) : undefined;
    const uniformMinimum = minimumRuns?.every(run => view.getComputedStyle(run).lineHeight === style.lineHeight);
    if (!(lineHeight > 0) || !(fontSize > 0) ||
      !uniformMinimum && Math.abs(lineHeight - natural) > 0.01) continue;
    const descent = descentRatio(style);
    if (!(descent > 0)) continue;
    const childStyle = view.getComputedStyle(first);
    // Word: the extra height of a multiple goes below the text (issue #908), so the baseline is one natural
    // line height less the descent below the line's top, however tall the line is. A grid line centres its
    // text (#1551), and where an at-least minimum wins on the grid, half the grid lines' surplus stays below.
    const pitch = Number.parseFloat(style.getPropertyValue("--docx-line-grid")) * 4 / 3;
    const centred = uniformMinimum && pitch > 0 ? (natural - gridLineHeight(natural, pitch, 1)) / 2 : gridTextOffset(childStyle);
    const shift = (lineHeight + centred - descent * fontSize) - chromiumOffset(style, childStyle);
    // The correction is Chromium's rounding plus the half of any line gap it puts below the text, which Word
    // puts above: a fraction of the line. Half a line or more means the font does not follow this model.
    if (!Number.isFinite(shift) || Math.abs(shift) < 1 / 64 || Math.abs(shift) >= lineHeight / 2) continue;
    for (const child of Array.from(paragraph.children))
      if (child instanceof view.HTMLElement && (movable(child, view)
        || uniformMinimum && minimumRuns!.includes(child))) {
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
  if (element.querySelector("img, svg, math")) return false;
  for (const descendant of Array.from(element.querySelectorAll("*")))
    if (["absolute", "fixed"].includes(view.getComputedStyle(descendant).position)) return false;
  return true;
}

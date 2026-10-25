/**
 * Fractional baselines for uniform Word paragraphs in the standalone export.
 *
 * HTML print snaps each inline baseline to a whole CSS pixel, even with an
 * explicit fractional line height. An SVG retains the
 * measured line breaks, uses font advances for plain word positions, and paints baselines at the
 * font's unrounded ascent. Source paragraph identities and block geometry stay
 * on the original elements. Complex flow still uses the existing HTML renderer.
 */

const PREPARED = 'data-docx-word-text-prepared';
const FITTED_RUN = 'data-docx-word-fit';
const FITTED_LINE = 'data-docx-word-line';
const SVG_NS = 'http://www.w3.org/2000/svg';
const PARAGRAPHS = 'p, h1, h2, h3, h4, h5, h6';

function plainDocument(root: HTMLElement): boolean {
  if (root.querySelector('table, svg, math, canvas, video, audio, iframe, object, a, ol, ul, aside, template, section.footnotes, section.endnotes'))
    return false;
  if (Array.from(root.querySelectorAll('img')).some(img => !img.closest(PARAGRAPHS))) return false;
  // A document mixing plain text with shifted runs, exact spacing or rich
  // inline structure keeps one consistent legacy layout. Those contexts need
  // their own line model before they can share the precision renderer.
  // Running stories are ordinary paragraphs measured into their own bands, so
  // they qualify like body text (issue #1123). A right-to-left paragraph keeps
  // its HTML line layout, so it no longer disqualifies the left-to-right ones
  // when it is otherwise plain; it is held to every other condition here.
  const view = root.ownerDocument.defaultView;
  return Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPHS)).every(p => {
    const style = view?.getComputedStyle(p);
    const run = plainRun(p, style?.direction === 'rtl');
    return run !== null && Number(style?.getPropertyValue('--docx-auto-line-spacing')) > 0;
  });
}

function pictureSibling(element: Element, view: Window & typeof globalThis): boolean {
  if (!(element instanceof view.HTMLElement) || element.localName !== 'span' || element.children.length !== 1)
    return false;
  const wrapper = element.firstElementChild!;
  if (wrapper.localName !== 'span' || wrapper.children.length !== 1 || wrapper.firstElementChild?.localName !== 'img')
    return false;
  if ((wrapper as HTMLElement).dataset.docxEdgePicture !== 'true') return false;
  const style = view.getComputedStyle(wrapper);
  const outer = view.getComputedStyle(element);
  const image = view.getComputedStyle(wrapper.firstElementChild);
  return ['left', 'right'].includes(style.cssFloat) && style.position === 'static' && style.transform === 'none' &&
    outer.display === 'inline' && ['static', 'relative'].includes(outer.position) && outer.transform === 'none' &&
    image.transform === 'none' && image.position === 'static';
}

/**
 * The paragraph's single plain text run, or null. `rightToLeft` admits a right-to-left paragraph
 * whose structure is otherwise plain, for the document guard only: such a paragraph keeps the
 * HTML line layout, so its direction and script need no precision line model.
 */
function plainRun(paragraph: HTMLElement, rightToLeft = false): HTMLElement | null {
  const view = paragraph.ownerDocument.defaultView;
  if (!view) return null;
  const children = Array.from(paragraph.children);
  const content = children.filter(child => child.textContent !== '');
  if (content.length !== 1) return null;
  const run = content[0];
  if (children.length !== 1) {
    // The converter carries an explicit column break on the paragraph and
    // leaves an empty run where the break appeared. It has no text identity;
    // accept only empty spans with the same line metrics as the visible run.
    const runStyle = view.getComputedStyle(run);
    for (const empty of children.filter(child => child !== run)) {
      if (pictureSibling(empty, view)) {
        const pictureStyle = view.getComputedStyle(empty);
        if (['font-size', 'font-style', 'font-weight', 'vertical-align'].some(property =>
          pictureStyle.getPropertyValue(property) !== runStyle.getPropertyValue(property))) return null;
        continue;
      }
      if (paragraph.dataset.columnBreakBefore !== 'true') return null;
      if (!(empty instanceof view.HTMLElement) || empty.localName !== 'span' || empty.childNodes.length ||
        Array.from(empty.attributes).some(a => !['class', 'style', 'lang', 'dir'].includes(a.name))) return null;
      const emptyStyle = view.getComputedStyle(empty);
      if (['font-family', 'font-size', 'font-style', 'font-weight', 'font-stretch', 'line-height',
        'display', 'position', 'top', 'vertical-align'].some(property =>
        emptyStyle.getPropertyValue(property) !== runStyle.getPropertyValue(property))) return null;
    }
  }
  if (!(run instanceof view.HTMLElement) || run.localName !== 'span') return null;
  const fitted = Number(run.getAttribute(FITTED_RUN)) > 0;
  if (fitted) {
    if (!run.childNodes.length || Array.from(run.childNodes).some(node =>
      !(node instanceof view.HTMLElement) || node.localName !== 'span' ||
      node.getAttribute(FITTED_LINE) !== 'true' || node.children.length ||
      node.childNodes.length > 1 || (node.firstChild && node.firstChild.nodeType !== 3))) return null;
  } else if (run.children.length || run.childNodes.length !== 1 || run.firstChild?.nodeType !== 3) return null;
  const text = run.textContent ?? '';
  // Other scripts need their own fallback-font metrics, shaping and bidi model.
  if (!text.trim() || /^\s/.test(text) || (!rightToLeft && !/^[\u0020-\u024f]*$/u.test(text))) return null;
  if (Array.from(run.attributes).some(a => !['class', 'style', 'lang', 'dir', ...(fitted ? [FITTED_RUN] : [])].includes(a.name))) return null;
  const style = view.getComputedStyle(run);
  const parent = view.getComputedStyle(paragraph);
  if (['Top', 'Right', 'Bottom', 'Left'].some(side =>
    parent.getPropertyValue(`padding-${side.toLowerCase()}`) !== '0px' ||
    parent.getPropertyValue(`border-${side.toLowerCase()}-width`) !== '0px')) return null;
  if (parent.direction !== (rightToLeft ? 'rtl' : 'ltr') || parent.writingMode !== 'horizontal-tb' ||
      style.display !== 'inline' || !['static', 'relative'].includes(style.position) ||
      style.textTransform !== 'none' || style.fontVariantCaps !== 'normal' ||
      style.textDecorationLine !== 'none' || !['normal', '0px'].includes(style.letterSpacing) ||
      !['normal', '100%'].includes(style.fontStretch) ||
      style.backgroundColor !== 'rgba(0, 0, 0, 0)' || style.textShadow !== 'none' ||
      style.transform !== 'none' || style.opacity !== '1' || style.verticalAlign !== 'baseline') return null;
  return run;
}

/** Word excludes the contour across the native text band, with horizontal clearance. */
function contourRows(image: HTMLImageElement, paragraph: HTMLElement, run: HTMLElement, lineCount: number, checkpoint?: () => void):
  Map<number, Array<{ text: string; x: number }>> | null {
  const view = paragraph.ownerDocument.defaultView!;
  const wrapper = image.parentElement!;
  if (wrapper.dataset.docxWrapText !== 'bothSides' ||
      image.closest('.page-content') !== paragraph.closest('.page-content')) return null;
  const wrapStyle = view.getComputedStyle(wrapper);
  if (wrapStyle.marginTop !== '0px' || wrapStyle.marginBottom !== '0px') return null;
  const polygon = /^polygon\((.*)\) border-box$/.exec(wrapStyle.shapeOutside);
  if (!polygon) return null;
  const points = polygon[1].split(',').map(point => {
    const match = /^\s*([\d.+-]+)%\s+([\d.+-]+)%\s*$/.exec(point);
    return match ? { x: Number(match[1]) / 100, y: Number(match[2]) / 100 } : null;
  });
  if (points.length < 3 || points.some(p => !p || !Number.isFinite(p.x + p.y))) return null;
  const vertices = points as Array<{ x: number; y: number }>;
  // One convex contour has a single excluded interval at each height. More
  // complex contours keep the established CSS flow until they have a line model.
  let orientation = 0;
  for (let i = 0; i < vertices.length; i++) {
    const a = vertices[i], b = vertices[(i + 1) % vertices.length], c = vertices[(i + 2) % vertices.length];
    const cross = (b.x - a.x) * (c.y - b.y) - (b.y - a.y) * (c.x - b.x);
    if (Math.abs(cross) < 1e-8) continue;
    if (orientation && Math.sign(cross) !== orientation) return null;
    orientation = Math.sign(cross);
  }
  if (!orientation) return null;
  const left = Number.parseFloat(wrapStyle.getPropertyValue('--docx-wrap-distance-left')) * 4 / 3;
  const right = Number.parseFloat(wrapStyle.getPropertyValue('--docx-wrap-distance-right')) * 4 / 3;
  if (!Number.isFinite(left + right) || left < 0 || right < 0) return null;
  const box = paragraph.getBoundingClientRect(); const picture = image.getBoundingClientRect();
  if (box.top >= picture.bottom || box.bottom <= picture.top) return null;
  const style = view.getComputedStyle(run);
  if (!['start', 'left'].includes(view.getComputedStyle(paragraph).textAlign)) return null;
  const pitch = Number.parseFloat(style.lineHeight);
  const nativeHeight = Number.parseFloat(style.getPropertyValue('--docx-word-native-height'));
  if (!(nativeHeight > 0) || !(pitch > 0) || !(picture.width > 0) || !(picture.height > 0)) return null;
  const context = paragraph.ownerDocument.createElement('canvas').getContext('2d');
  if (!context) return null;
  const measureScale = Number.parseFloat(style.fontSize) / 10000;
  context.font = `${style.fontStyle} ${style.fontWeight} 10000px ${style.fontFamily}`;
  context.fontKerning = style.fontKerning as CanvasFontKerning;
  if (!context.font.includes('10000px')) return null;
  const words = Array.from(run.textContent!.matchAll(/\S+\s*/gu), (match, index) => {
    if ((index & 63) === 0) checkpoint?.();
    return { text: match[0], width: context.measureText(match[0].trimEnd()).width * measureScale,
      advance: context.measureText(match[0]).width * measureScale };
  });
  const rows = new Map<number, Array<{ text: string; x: number }>>();
  let next = 0;
  for (let line = 0; line < lineCount && next < words.length; line++) {
    checkpoint?.();
    const top = (box.top + line * pitch - picture.top) / picture.height;
    const bottom = top + nativeHeight / picture.height;
    const crossings: number[] = [];
    for (let i = 0; i < vertices.length; i++) {
      const a = vertices[i], b = vertices[(i + 1) % vertices.length];
      if (a.y >= top && a.y <= bottom) crossings.push(a.x);
      for (const y of [top, bottom])
        if (a.y !== b.y && y >= Math.min(a.y, b.y) && y <= Math.max(a.y, b.y))
          crossings.push(a.x + (b.x - a.x) * (y - a.y) / (b.y - a.y));
    }
    const slots = crossings.length ? [
      [box.left, Math.min(box.right, picture.left + Math.min(...crossings) * picture.width - left)],
      [Math.max(box.left, picture.left + Math.max(...crossings) * picture.width + right), box.right],
    ] : [[box.left, box.right]];
    const row: Array<{ text: string; x: number }> = [];
    for (const [start, end] of slots) {
      let x = start;
      while (next < words.length && x + words[next].width <= end + 1 / 64) {
        const word = words[next++]; row.push({ text: word.text, x: x - box.left }); x += word.advance;
      }
    }
    if (!row.length) return null;
    rows.set(line, row);
  }
  // Painting cannot change pagination or paragraph geometry. Use this model
  // only when its complete source inventory fits the measured number of lines.
  return next === words.length && rows.size === lineCount ? rows : null;
}

/** Measure advances at a large em to avoid CSS-pixel font hinting in line positions. */
function wordMeasure(run: HTMLElement): ((text: string) => number) | null {
  const style = run.ownerDocument.defaultView!.getComputedStyle(run);
  const context = run.ownerDocument.createElement('canvas').getContext('2d');
  if (!context) return null;
  context.font = `${style.fontStyle} ${style.fontWeight} 10000px ${style.fontFamily}`;
  context.fontKerning = style.fontKerning as CanvasFontKerning;
  if (!context.font.includes('10000px')) return null;
  const scale = Number.parseFloat(style.fontSize) / 10000;
  return text => context.measureText(text).width * scale;
}

/** Resolve plain line contents before pagination, using unrounded font advances. */
function fitWordLines(paragraph: HTMLElement, run: HTMLElement, modernJustification: boolean,
  checkpoint?: () => void): boolean {
  const view = run.ownerDocument.defaultView!;
  const style = view.getComputedStyle(paragraph);
  const runStyle = view.getComputedStyle(run);
  const text = run.textContent!;
  // Hyphens/slashes and nonordinary spaces have additional break opportunities.
  // Pictures and columns retain their separately measured flow.
  if (runStyle.whiteSpace !== 'pre-wrap' || /[^\u0020-\u007e\u00c0-\u024f]|[-/]/u.test(text) ||
      !/^-?[\d.]+px$/.test(style.textIndent) || style.wordBreak !== 'normal' ||
      (!['normal', '0px'].includes(runStyle.wordSpacing) &&
       !run.style.getPropertyValue('--docx-word-space-compression'))) return false;
  const measure = wordMeasure(run);
  const width = paragraph.getBoundingClientRect().width;
  const indent = Number.parseFloat(style.textIndent) || 0;
  if (!measure || !(width > 0) || !(width - indent > 0)) return false;
  const space = measure(' ');
  const words = Array.from(text.matchAll(/\S+ */gu), (match, index) => {
    if ((index & 63) === 0) checkpoint?.();
    const word = match[0].trimEnd();
    return { text: match[0], width: measure(word), spaces: match[0].length - word.length };
  });
  if (!words.length || words.some(w => w.width > width) || words[0].width > width - indent) return false;
  const rows: Array<{ text: string; spaces: number; available: number }> = [];
  let start = 0;
  while (start < words.length) {
    checkpoint?.();
    const available = width - (rows.length ? 0 : indent);
    let natural = 0, spaces = 0, best = start, bestCost = Infinity;
    for (let end = start; end < words.length; end++) {
      if ((end & 63) === 0) checkpoint?.();
      if (end > start) { spaces += words[end - 1].spaces; natural += words[end - 1].spaces * space; }
      natural += words[end].width;
      const minimum = natural - (modernJustification ? spaces * space / 4 : 0);
      if (minimum > available) break;
      // Native Word prefers less gap strain, with condensation twice as costly
      // as expansion. A final line that fits naturally needs no justification.
      // The 75% floor bounds which condensed alternatives are legal.
      const cost = !modernJustification ? -end : end === words.length - 1 && natural <= available ? -1 :
        spaces ? Math.abs(available - natural) / (spaces * space) * (natural > available ? 2 : 1) : Infinity;
      if (cost <= bestCost) { best = end; bestCost = cost; }
    }
    rows.push({ text: words.slice(start, best + 1).map(w => w.text).join(''),
      spaces: words.slice(start, best).reduce((n, w) => n + w.spaces, 0), available });
    start = best + 1;
  }
  const previousStyle = run.getAttribute('style');
  const fragment = run.ownerDocument.createDocumentFragment();
  const lines = rows.map((row, index) => {
    checkpoint?.();
    const line = run.ownerDocument.createElement('span');
    line.setAttribute(FITTED_LINE, 'true');
    line.style.cssText = 'display:block;white-space:pre;word-spacing:0px;text-indent:0';
    if (!index && indent) line.style.marginLeft = `${indent}px`;
    line.textContent = row.text;
    fragment.append(line);
    return line;
  });
  run.style.position = 'static';
  run.style.top = '0';
  run.style.wordSpacing = '0px';
  run.replaceChildren(fragment);
  for (let index = 0; index < lines.length; index++) {
    checkpoint?.();
    const range = run.ownerDocument.createRange();
    range.setStart(lines[index].firstChild!, 0);
    range.setEnd(lines[index].firstChild!, rows[index].text.trimEnd().length);
    const excess = range.getBoundingClientRect().width - rows[index].available;
    if (excess <= 0) continue;
    if (!rows[index].spaces) {
      run.replaceChildren(run.ownerDocument.createTextNode(text));
      run.removeAttribute(FITTED_RUN);
      if (previousStyle === null) run.removeAttribute('style'); else run.setAttribute('style', previousStyle);
      return false;
    }
    // The DOM supplies line boxes for pagination. Compensate its font rounding
    // only within this chosen line; vector paint uses the font's true advances.
    lines[index].style.wordSpacing = `${-excess / rows[index].spaces}px`;
  }
  // Carry the measured natural height with the owned line boxes, including
  // Range clones, so pagination need not resolve their font metrics again.
  run.setAttribute(FITTED_RUN, String(Number.parseFloat(run.style.getPropertyValue('--docx-word-native-height'))));
  if (modernJustification) run.style.setProperty('--docx-word-space-compression', `${-space / 4}px`);
  return true;
}

/** Space-aware positioning for an ordinary, unfragmented left-to-right line box. */
function positionWords(paragraph: HTMLElement, run: HTMLElement,
  rows: Map<number, Array<{ text: string; x: number }>>, checkpoint?: () => void): void {
  // NBSP, soft hyphens and other break controls need their own model. Retain
  // the DOM positions for them, including the source's nonbreaking semantics.
  if (/[^\u0020-\u007e\u00c0-\u024f]/u.test(run.textContent!)) return;
  const view = run.ownerDocument.defaultView!;
  const style = view.getComputedStyle(run);
  if (style.whiteSpace !== 'pre-wrap') return;
  if (!['normal', '0px'].includes(style.wordSpacing) &&
      !run.style.getPropertyValue('--docx-word-space-compression')) return;
  const measure = wordMeasure(run);
  if (!measure) return;
  const parent = view.getComputedStyle(paragraph);
  const width = paragraph.getBoundingClientRect().width;
  if (!/^-?[\d.]+px$/.test(parent.textIndent)) return;
  const indent = Number.parseFloat(parent.textIndent) || 0;
  const lines = Array.from(rows.values());
  lines.forEach((words, index) => {
    checkpoint?.();
    const start = index === 0 ? indent : 0;
    const available = width - start;
    const last = index === lines.length - 1;
    let align = parent.textAlign;
    if (last) align = parent.textAlignLast === 'auto' ? (align === 'justify' ? 'start' : align) : parent.textAlignLast;
    const advances = words.map((word, i) => {
      if ((i & 63) === 0) checkpoint?.();
      return measure(word.text);
    });
    const natural = advances.reduce((a, b) => a + b, 0) -
      (advances.at(-1)! - measure(words.at(-1)!.text.trimEnd()));
    const spaces = words.slice(0, -1).reduce((count, word) => count + (word.text.match(/ +$/)?.[0].length ?? 0), 0);
    const compressed = run.style.getPropertyValue('--docx-word-space-compression') && natural > available;
    const adjustment = spaces && (align === 'justify' || compressed) ? (available - natural) / spaces : 0;
    const remaining = available - natural - adjustment * spaces;
    let x = start + (['right', 'end'].includes(align) ? remaining :
      align === 'center' ? remaining / 2 : 0);
    words.forEach((word, i) => {
      word.x = x;
      x += advances[i] + adjustment * (word.text.match(/ +$/)?.[0].length ?? 0);
    });
  });
}

/** Run after fonts and natural line heights resolve, before pagination measures them. */
export function prepareWordText(root: HTMLElement, checkpoint?: () => void): number {
  if (!plainDocument(root)) return 0;
  const view = root.ownerDocument.defaultView;
  if (!view) return 0;
  // Pagination measures each section at this declared width. Line fitting must
  // use that same band before the engine starts collecting its block heights.
  for (const section of Array.from(root.querySelectorAll<HTMLElement>('[data-section-index][data-content-width]'))) {
    const width = Number(section.dataset.contentWidth);
    if (Number.isFinite(width) && width > 0) section.style.width = `${width}pt`;
  }
  let prepared = 0;
  for (const paragraph of Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPHS))) {
    checkpoint?.();
    const run = plainRun(paragraph);
    if (!run) continue;
    const style = view.getComputedStyle(paragraph);
    const multiple = Number(style.getPropertyValue('--docx-auto-line-spacing'));
    const runStyle = view.getComputedStyle(run);
    const kerning = runStyle.getPropertyValue('--docx-font-kerning').trim() || 'none';
    // A unit multiple leaves the run at `line-height: normal`; the export's unrounded
    // normal-line pass (line-metrics.ts) has already resolved it to a length.
    const height = Number.parseFloat(runStyle.lineHeight);
    if (!(multiple > 0) || !(height > 0) || !['none', 'normal'].includes(kerning)) continue;
    // Word's auto spacing advances on the document's twip grid. Resolve the
    // native single-line height to twips before applying its spacing multiple.
    const nativeTwips = Math.round(height / multiple * 15);
    const twips = Math.floor(nativeTwips * multiple);
    if (!(twips > 0)) continue;
    run.style.fontKerning = kerning;
    const mode = Number(root.closest('[data-docx-compatibility-mode]')?.getAttribute('data-docx-compatibility-mode') ??
      root.ownerDocument.body.getAttribute('data-docx-compatibility-mode') ?? 14);
    let column = false;
    for (let ancestor = paragraph.parentElement; ancestor && ancestor !== root; ancestor = ancestor.parentElement)
      if (Number(view.getComputedStyle(ancestor).columnCount) > 1) column = true;
    run.style.lineHeight = `${twips / 15}px`;
    run.style.setProperty('--docx-word-native-height', `${nativeTwips / 15}px`);
    if (!column && !root.querySelector('img'))
      fitWordLines(paragraph, run, mode >= 15 && style.textAlign === 'justify', checkpoint);
    // An empty drawing run participates in the first inline box. Give it the
    // same resolved height so its inherited auto multiplier is not applied twice.
    for (const picture of Array.from(paragraph.children).filter(child => pictureSibling(child, view))) {
      (picture as HTMLElement).style.lineHeight = run.style.lineHeight;
      (picture as HTMLElement).style.top = '0';
    }
    for (const empty of Array.from(paragraph.children).filter(child => child !== run && !child.children.length)) empty.remove();
    paragraph.style.lineHeight = run.style.lineHeight;
    paragraph.setAttribute(PREPARED, 'true');
    prepared++;
  }
  return prepared;
}

/** Paint prepared paragraphs after the page tree is final. No line or page breaks change here. */
export function paintWordText(root: HTMLElement, checkpoint?: () => void): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const paragraphs = Array.from(root.querySelectorAll<HTMLElement>(`[${PREPARED}]`));
  const pictures = Array.from(root.querySelectorAll('img'));
  const hasPictures = pictures.length > 0;
  const columnOf = (paragraph: HTMLElement): HTMLElement | null => {
    for (let parent = paragraph.parentElement; parent && parent !== root; parent = parent.parentElement)
      if (Number(view.getComputedStyle(parent).columnCount) > 1) return parent;
    return null;
  };
  // A column container with fragmented paragraphs keeps its original line flow.
  // Mixing fixed vector boxes with an HTML paragraph that crosses a column can
  // reflow that paragraph outside the body band and drop its printed text.
  const fragmentedColumns = new Set<HTMLElement>();
  for (const paragraph of paragraphs) {
    checkpoint?.();
    const column = columnOf(paragraph);
    if (column && paragraph.getClientRects().length !== 1) fragmentedColumns.add(column);
  }
  const metrics = new Map<string, { height: number; descent: number }>();
  function fontMetrics(style: CSSStyleDeclaration): { height: number; descent: number } | null {
    const font = `${style.fontStyle} ${style.fontWeight} 10000px ${style.fontFamily}`;
    const cached = metrics.get(font);
    if (cached) return cached;
    const probe = document.createElement('span');
    probe.setAttribute('aria-hidden', 'true');
    probe.style.cssText = `position:absolute;left:-100000px;top:0;visibility:hidden;display:inline-block;white-space:nowrap;margin:0;padding:0;border:0;font:${font};line-height:normal`;
    probe.textContent = 'Hg';
    document.body.append(probe);
    const height = probe.getBoundingClientRect().height / 10000;
    probe.remove();
    const context = document.createElement('canvas').getContext('2d');
    if (!context) return null;
    context.font = font;
    if (!context.font.includes('10000px')) return null;
    const descent = context.measureText('Hg').fontBoundingBoxDescent / 10000;
    if (!(height > 0) || !(descent > 0)) return null;
    const result = { height, descent };
    metrics.set(font, result);
    return result;
  }

  const pending: Array<{ paragraph: HTMLElement; run: HTMLElement; svg: SVGSVGElement; height: number }> = [];
  // Chromium lays a twip-grid line (19.333px for 14.5pt) out in 1/64-pixel units (19.328125px), so
  // each prepared line ends slightly early and later paragraphs drift up. Paint each paragraph at
  // its ideal place: its layout top plus what the prepared lines above it on the page lost.
  const drift = new Map<Element, number>();
  for (const paragraph of paragraphs) {
    checkpoint?.();
    const column = columnOf(paragraph);
    if (column && fragmentedColumns.has(column)) {
      paragraph.removeAttribute(PREPARED);
      continue;
    }
    const run = plainRun(paragraph);
    if (!run) continue;
    const style = view.getComputedStyle(run);
    const metric = fontMetrics(style);
    const fontSize = Number.parseFloat(style.fontSize);
    const pitch = Number.parseFloat(style.lineHeight);
    const box = paragraph.getBoundingClientRect();
    if (!metric || !(fontSize > 0) || !(pitch > 0) || !(box.width > 0) || !(box.height > 0)) continue;
    const columnBand = column || hasPictures ? paragraph.closest<HTMLElement>('.page-content') : null;
    const band = columnBand?.getBoundingClientRect() ?? box;
    let rows = new Map<number, Array<{ text: string; x: number }>>();
    let wordCount = 0;
    let splitWord = false;
    const walker = document.createTreeWalker(run, view.NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) for (const match of node.textContent!.matchAll(/\S+\s*/gu)) {
      if ((wordCount++ & 63) === 0) checkpoint?.();
      const range = document.createRange();
      range.setStart(node, match.index!);
      range.setEnd(node, match.index! + match[0].trimEnd().length);
      if (range.getClientRects().length !== 1) { splitWord = true; break; }
      const rect = range.getBoundingClientRect();
      const row = Math.round(rect.top * 4);
      const words = rows.get(row) ?? [];
      words.push({ text: match[0], x: rect.left - box.left });
      rows.set(row, words);
    }
    if (!rows.size || splitWord) continue;
    if (!hasPictures && !column) positionWords(paragraph, run, rows, checkpoint);
    if (pictures.length === 1) rows = contourRows(pictures[0], paragraph, run, rows.size, checkpoint) ?? rows;
    const pageBox = paragraph.closest('.page-box');
    const lost = column || !pageBox ? 0 : drift.get(pageBox) ?? 0;
    if (pageBox && !column) {
      const shortfall = rows.size * pitch - box.height;
      // Only layout-unit truncation (well under a pixel) is a drift; anything else is geometry.
      if (shortfall > 0 && shortfall < 1) drift.set(pageBox, lost + shortfall);
    }
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('width', String(band.width));
    svg.setAttribute('height', String(band.height));
    // A larger viewBox retains fractional glyph coordinates without increasing
    // the physical font size or changing the source font file.
    svg.setAttribute('viewBox', `0 0 ${band.width * 4} ${band.height * 4}`);
    // Chromium paints the SVG from its layout origin rounded to the nearest CSS pixel of its page
    // (a paragraph at 100.32px paints its vector from 100px). Keep that remainder inside the
    // vector, so a fractional indent or margin still puts every glyph at its layout position.
    const page = paragraph.closest<HTMLElement>('.page-box')?.getBoundingClientRect();
    const remainder = (offset: number) => offset - Math.round(offset);
    const dx = page ? remainder(band.left - page.left) : 0;
    const dy = page ? remainder(band.top - page.top) : 0;
    svg.setAttribute('role', 'group');
    svg.style.cssText = columnBand
      ? 'position:absolute;left:0;top:0;display:block;overflow:visible'
      : 'display:block;overflow:visible';
    const baseline = (metric.height - metric.descent) * fontSize;
    Array.from(rows.values()).forEach((words, index) => {
      const line = document.createElementNS(SVG_NS, 'text');
      line.setAttribute('y', String((box.top - band.top + dy + lost + baseline + index * pitch) * 4));
      line.setAttribute('xml:space', 'preserve');
      for (const property of ['font-family', 'font-weight', 'font-style', 'font-stretch', 'font-kerning'])
        line.style.setProperty(property, style.getPropertyValue(property));
      line.style.fontSize = `${fontSize * 4}px`;
      line.style.fill = style.color;
      for (const word of words) {
        const span = document.createElementNS(SVG_NS, 'tspan');
        span.setAttribute('x', String((box.left - band.left + dx + word.x) * 4));
        span.textContent = word.text;
        line.append(span);
      }
      svg.append(line);
    });
    pending.push({ paragraph, run, svg, height: box.height });
  }
  // Read every line before replacing any paragraph, so one vector box cannot
  // change the column coordinates used to measure the following paragraph.
  for (const { paragraph, run, svg, height } of pending) {
    checkpoint?.();
    if (Array.from(paragraph.children).some(child => pictureSibling(child, view))) run.replaceWith(svg);
    else paragraph.replaceChildren(svg);
    // CSSOM rounds a numeric height when serializing an inline declaration.
    // A slightly lower re-parsed value loses one layout unit on every paragraph
    // and shifts PageMap fragments on offline reopen. Custom-property tokens
    // retain the measured value verbatim through the standalone HTML round trip.
    paragraph.style.setProperty('--docx-word-text-height', `${height}px`);
    paragraph.style.height = 'var(--docx-word-text-height)';
    paragraph.removeAttribute(PREPARED);
  }
  return pending.length;
}

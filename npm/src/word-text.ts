/**
 * Fractional baselines for uniform Word paragraphs in the standalone export.
 *
 * HTML print snaps each inline baseline to a whole CSS pixel, even with an
 * explicit fractional line height. A paragraph-contained SVG retains the
 * measured line breaks and word positions while painting the baselines at the
 * font's unrounded ascent. Source paragraph identities and block geometry stay
 * on the original elements. Complex flow still uses the existing HTML renderer.
 */

const PREPARED = 'data-docx-word-text-prepared';
const SVG_NS = 'http://www.w3.org/2000/svg';
const PARAGRAPHS = 'p, h1, h2, h3, h4, h5, h6';

function plainDocument(root: HTMLElement): boolean {
  if (root.querySelector('table, img, svg, math, canvas, video, audio, iframe, object, a, ol, ul, aside, template, header, footer, [data-hf-type], section.footnotes, section.endnotes, [data-column-break-before]'))
    return false;
  if (Array.from(root.querySelectorAll('[data-cols]')).some(e => Number(e.getAttribute('data-cols')) > 1)) return false;
  // A document mixing plain text with shifted runs, exact spacing or rich
  // inline structure keeps one consistent legacy layout. Those contexts need
  // their own line model before they can share the precision renderer.
  return Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPHS)).every(p => {
    const run = plainRun(p);
    return run !== null && Number(root.ownerDocument.defaultView?.getComputedStyle(p)
      .getPropertyValue('--docx-auto-line-spacing')) > 0;
  });
}

function plainRun(paragraph: HTMLElement): HTMLElement | null {
  const view = paragraph.ownerDocument.defaultView;
  if (!view || paragraph.children.length !== 1) return null;
  const run = paragraph.firstElementChild;
  if (!(run instanceof view.HTMLElement) || run.localName !== 'span' || run.children.length ||
      run.childNodes.length !== 1 || run.firstChild?.nodeType !== 3) return null;
  const text = run.textContent ?? '';
  // Other scripts need their own fallback-font metrics, shaping and bidi model.
  if (!text.trim() || /^\s/.test(text) || !/^[\u0020-\u024f]*$/u.test(text)) return null;
  if (Array.from(run.attributes).some(a => !['class', 'style', 'lang', 'dir'].includes(a.name))) return null;
  const style = view.getComputedStyle(run);
  const parent = view.getComputedStyle(paragraph);
  if (['Top', 'Right', 'Bottom', 'Left'].some(side =>
    parent.getPropertyValue(`padding-${side.toLowerCase()}`) !== '0px' ||
    parent.getPropertyValue(`border-${side.toLowerCase()}-width`) !== '0px')) return null;
  if (parent.direction !== 'ltr' || parent.writingMode !== 'horizontal-tb' ||
      style.display !== 'inline' || !['static', 'relative'].includes(style.position) ||
      style.textTransform !== 'none' || style.fontVariantCaps !== 'normal' ||
      style.textDecorationLine !== 'none' || !['normal', '0px'].includes(style.letterSpacing) ||
      !['normal', '100%'].includes(style.fontStretch) ||
      style.backgroundColor !== 'rgba(0, 0, 0, 0)' || style.textShadow !== 'none' ||
      style.transform !== 'none' || style.opacity !== '1' || style.verticalAlign !== 'baseline') return null;
  return run;
}

/** Run after fonts and natural line heights resolve, before pagination measures them. */
export function prepareWordText(root: HTMLElement, checkpoint?: () => void): number {
  if (!plainDocument(root)) return 0;
  const view = root.ownerDocument.defaultView;
  if (!view) return 0;
  let prepared = 0;
  for (const paragraph of Array.from(root.querySelectorAll<HTMLElement>(PARAGRAPHS))) {
    checkpoint?.();
    const run = plainRun(paragraph);
    if (!run) continue;
    const style = view.getComputedStyle(paragraph);
    const multiple = Number(style.getPropertyValue('--docx-auto-line-spacing'));
    const runStyle = view.getComputedStyle(run);
    const kerning = runStyle.getPropertyValue('--docx-font-kerning').trim() || 'none';
    const height = Number.parseFloat(runStyle.lineHeight);
    if (!(multiple > 0) || !(height > 0) || !['none', 'normal'].includes(kerning)) continue;
    // Word's auto spacing advances on the document's twip grid. Resolve the
    // native single-line height to twips before applying its spacing multiple.
    const twips = Math.floor(Math.round(height / multiple * 15) * multiple);
    if (!(twips > 0)) continue;
    run.style.fontKerning = kerning;
    run.style.lineHeight = `${twips / 15}px`;
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

  let painted = 0;
  for (const paragraph of Array.from(root.querySelectorAll<HTMLElement>(`[${PREPARED}]`))) {
    checkpoint?.();
    const run = plainRun(paragraph);
    if (!run) continue;
    const style = view.getComputedStyle(run);
    const metric = fontMetrics(style);
    const fontSize = Number.parseFloat(style.fontSize);
    const pitch = Number.parseFloat(style.lineHeight);
    const box = paragraph.getBoundingClientRect();
    if (!metric || !(fontSize > 0) || !(pitch > 0) || !(box.width > 0) || !(box.height > 0)) continue;
    const text = run.textContent!;
    const rows = new Map<number, Array<{ text: string; x: number }>>();
    let wordCount = 0;
    for (const match of text.matchAll(/\S+\s*/gu)) {
      if ((wordCount++ & 63) === 0) checkpoint?.();
      const range = document.createRange();
      range.setStart(run.firstChild!, match.index!);
      range.setEnd(run.firstChild!, match.index! + match[0].trimEnd().length);
      const rect = range.getBoundingClientRect();
      const row = Math.round(rect.top * 4);
      const words = rows.get(row) ?? [];
      words.push({ text: match[0], x: rect.left - box.left });
      rows.set(row, words);
    }
    if (!rows.size) continue;
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('width', String(box.width));
    svg.setAttribute('height', String(box.height));
    // A larger viewBox retains fractional glyph coordinates without increasing
    // the physical font size or changing the source font file.
    svg.setAttribute('viewBox', `0 0 ${box.width * 4} ${box.height * 4}`);
    svg.setAttribute('role', 'group');
    svg.style.cssText = 'display:block;overflow:visible';
    const baseline = (metric.height - metric.descent) * fontSize;
    Array.from(rows.values()).forEach((words, index) => {
      const line = document.createElementNS(SVG_NS, 'text');
      line.setAttribute('y', String((baseline + index * pitch) * 4));
      line.setAttribute('xml:space', 'preserve');
      for (const property of ['font-family', 'font-weight', 'font-style', 'font-stretch', 'font-kerning'])
        line.style.setProperty(property, style.getPropertyValue(property));
      line.style.fontSize = `${fontSize * 4}px`;
      line.style.fill = style.color;
      for (const word of words) {
        const span = document.createElementNS(SVG_NS, 'tspan');
        span.setAttribute('x', String(word.x * 4));
        span.textContent = word.text;
        line.append(span);
      }
      svg.append(line);
    });
    paragraph.replaceChildren(svg);
    // CSSOM rounds a numeric height when serializing an inline declaration.
    // A slightly lower re-parsed value loses one layout unit on every paragraph
    // and shifts PageMap fragments on offline reopen. Custom-property tokens
    // retain the measured value verbatim through the standalone HTML round trip.
    paragraph.style.setProperty('--docx-word-text-height', `${box.height}px`);
    paragraph.style.height = 'var(--docx-word-text-height)';
    paragraph.removeAttribute(PREPARED);
    painted++;
  }
  return painted;
}

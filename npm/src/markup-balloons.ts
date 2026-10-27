/**
 * Word's All Markup print with comments and formatting in balloons (issue #1300), behind
 * `commentColumnPlacement: "balloons"`.
 *
 * Word keeps the paper size: it shrinks the page, together with a grey markup area on its right,
 * until both fit the paper. Each comment, reply and formatting change gets a balloon in that area,
 * level with its line where there is room and tied to it by a leader line; insertions and
 * deletions stay inline. Here the page box itself is left paper-sized, and only a layer inside it
 * is shrunk, so PageMap pages, print sizing and the export's audits all still see the paper, and
 * PageMap fragments land where the shrunken text prints.
 */
import type { PageDimensions } from "./page-geometry.js";

/**
 * Word's balloon geometry and colours, measured from Word-printed All Markup PDFs. Lengths are page
 * points before the page is shrunk, except where the name says paper.
 */
const WORD = {
  /** The markup area: Word's default 3.7in balloon width, counted from the text column's right edge. */
  areaWidth: 266.4,
  /** The grey starts this far right of the text column... */
  greyStart: 9,
  /** ...and the balloons this far. */
  balloonStart: 31.6,
  /** Room left between a balloon and the sheet's right edge. */
  balloonEnd: 4,
  /** A leader runs along its anchor's line to this far past the text column, then turns. */
  elbow: 4.6,
  /** A leader reaches its balloon this far below the balloon's top. */
  leaderDrop: 6,
  /** The shrunken page's foot stands this far above the paper's, in paper points. */
  paperFootGap: 10.8,
  border: 0.5,
  corner: 3,
  padBlock: 1.8,
  padInline: 4,
  leading: 1.15,
  /** Vertical space between neighbouring balloons. */
  between: 1,
  /** Balloons keep this far inside the page's top and foot. */
  pageInset: 2,
  hairline: 0.2,
  /**
   * A balloon cut short to its first line, at the 9pt balloon font and scaled with it: Word's height
   * for it, and the box at its lower right corner holding the ellipsis (its height and horizontal
   * padding), measured from Word's Save as PDF and Print to PDF of one page with 25 long comments.
   * The line's text ends at least this far before the box; Word's lines break a word earlier, so
   * the exact gap was not measurable.
   */
  cut: { height: 16.1, boxHeight: 11.6, boxPadInline: 4, textGap: 1 },
  markSize: 3,
  grey: "#F2F2F2",
  font: { family: "Segoe UI", sizePt: 9 },
  /**
   * The Word presentation's author colours (WmlToHtmlConverter.WordAuthorPalette), for HTML that does
   * not carry them. Only the first two are Word's; the rest are this implementation's choice.
   */
  palette: ["#0078D4", "#D13438", "#5C2E91", "#038387", "#CA5010", "#498205", "#C239B3", "#986F0B"],
  /** Word's pale comment fills for its first two author colours. */
  paleFills: { "#0078D4": "#D5EDFF", "#D13438": "#F8DCDD" } as Record<string, string>,
};

/** Anchors whose line tops are this close share a line. */
const SAME_LINE = 2;

/** One margin comment from the converter's registry. */
export interface MarkupBalloonComment {
  id: string;
  parentId?: string;
  author?: string;
  initials?: string;
  color?: string;
  fontFamily?: string;
  paragraphs: string[];
}

export interface MarkupBalloonPage {
  dimensions: PageDimensions;
  element: HTMLElement;
}

/**
 * Reads the margin registry: notes in document order, a thread's replies nested in (so after) its
 * root, plus the class prefix the converter gave comment markup.
 */
export function readMarkupBalloonComments(registry: Element | null): {
  comments: MarkupBalloonComment[];
  commentPrefix: string;
} {
  let commentPrefix = "comment-";
  const comments = Array.from(registry?.querySelectorAll<HTMLElement>("[data-comment-id]") ?? []).flatMap((note) => {
    const noteClass = Array.from(note.classList).find((name) => name.endsWith("margin-note"));
    if (!noteClass) return [];
    commentPrefix = noteClass.replace(/margin-note$/, "");
    const { commentId: id = "", author, initials, authorColor: color, fontFamily } = note.dataset;
    const body = note.querySelector(`:scope > .${commentPrefix}margin-note-body`);
    const parentId = note.parentElement?.closest<HTMLElement>("[data-comment-id]")?.dataset.commentId;
    const paragraphs = Array.from(body?.children ?? [], (paragraph) => paragraph.textContent?.trim() ?? "");
    return [{ id, parentId, author, initials, color, fontFamily, paragraphs }];
  });
  return { comments, commentPrefix };
}

/** The attribute that scopes the balloon rules to the pagination using them. */
const SCOPE = "data-docx-markup-balloons";

/**
 * Turns the balloon rules on or off for one pagination's staging and page containers. On, they go in
 * before anything is measured, because they change line breaks: Word prints no comment reference
 * mark, and paints the commented range itself (see `layoutMarkupBalloons`). The rules apply only
 * inside elements carrying the scope attribute, so a later pagination in the same document is not
 * affected, and turning them off for its containers clears any attribute left behind.
 */
export function scopeMarkupBalloons(
  document: Document, containers: HTMLElement[], on: boolean, commentPrefix = "comment-", cssPrefix = "page-",
): void {
  for (const container of containers) container.toggleAttribute(SCOPE, on);
  const key = `${commentPrefix} ${cssPrefix}`;
  if (!on || Array.from(document.querySelectorAll(`style[${SCOPE}]`)).some((style) => style.getAttribute(SCOPE) === key)) return;
  const rules = document.createElement("style");
  rules.setAttribute(SCOPE, key);
  rules.textContent = [
    `[${SCOPE}] a.${commentPrefix}marker { font-size: 0 !important; margin: 0 !important; padding: 0 !important; visibility: hidden !important; }`,
    `[${SCOPE}] span.${commentPrefix}highlight { background: none !important; border: 0 !important; }`,
    `[${SCOPE}] .${cssPrefix}balloon { position: absolute; box-sizing: border-box; border: ${WORD.border}pt solid;` +
      ` border-radius: ${WORD.corner}pt; padding: ${WORD.padBlock}pt ${WORD.padInline}pt; line-height: ${WORD.leading};` +
      " color: #000; text-align: left; overflow: hidden; overflow-wrap: anywhere; }",
    `[${SCOPE}] .${cssPrefix}balloon b { font-weight: bold; }`,
  ].join("\n");
  document.head.append(rules);
}

/** What a balloon layout could not print in full. */
export interface MarkupBalloonResult {
  /** Comments whose balloons were cut short to fit beside their page. */
  truncatedCommentIds: string[];
  /** Formatting balloons cut short to fit. */
  truncatedFormatting: number;
  /** Comments marked on a page that got no balloon at all. */
  unplacedCommentIds: string[];
}

/** A balloon to place: what it says, how it is drawn, and the line it belongs to. */
interface Entry {
  label: string;
  text: string[];
  color: string;
  fill: string;
  textFont?: string;
  /** Document order, for balloons of one line. */
  order: number;
  lineTop: number;
  /** Where the leader leaves the text. */
  from: { x: number; y: number };
  commentId?: string;
  element?: HTMLElement;
  height: number;
  top: number;
}

type Mark = { d: string; color: string; filled: boolean };

/** Lays out every page's markup area. Call once, on the finished page tree. */
export function layoutMarkupBalloons(options: {
  document: Document;
  pages: MarkupBalloonPage[];
  comments: MarkupBalloonComment[];
  commentPrefix: string;
  cssPrefix: string;
  font?: { family: string; sizePt: number };
}): MarkupBalloonResult {
  const { document, pages, comments, commentPrefix, cssPrefix } = options;
  const result: MarkupBalloonResult = { truncatedCommentIds: [], truncatedFormatting: 0, unplacedCommentIds: [] };
  const font = options.font && options.font.sizePt > 0 ? options.font : WORD.font;
  const anchored = pages.map((page) =>
    Array.from(page.element.querySelectorAll<HTMLElement>("[data-comment-id], [data-format-description]")));
  // Word prints a document without comments or formatting changes as it is.
  if (anchored.every((elements) => elements.length === 0)) return result;

  // The registry lists a root before its replies, so each reply's thread is already known.
  const thread = new Map<string, string>();
  for (const { id, parentId } of comments) thread.set(id, parentId !== undefined && thread.has(parentId) ? thread.get(parentId)! : id);
  const byId = new Map(comments.map((comment) => [comment.id, comment]));
  const authorsSeen: string[] = [];
  const colorOf = ({ color, author = "" }: MarkupBalloonComment) => {
    if (color) return color;
    if (!authorsSeen.includes(author)) authorsSeen.push(author);
    return WORD.palette[authorsSeen.indexOf(author) % WORD.palette.length];
  };
  const rangeSpans = (scope: ParentNode, id?: string) => Array.from(scope.querySelectorAll<HTMLElement>(
    `span.${commentPrefix}highlight[data-comment-id${id === undefined ? "" : `="${CSS.escape(id)}"`}]`));
  for (const { element: pageBox } of pages) {
    for (const span of rangeSpans(pageBox)) {
      const root = byId.get(thread.get(span.dataset.commentId!) ?? "");
      if (root) span.style.setProperty("background-color", paleFill(colorOf(root)), "important");
    }
  }
  // Headers and footers print on every page, and their balloons with them. In the body, a
  // paragraph's balloon goes on the page holding its mark (its last fragment), and a comment's on
  // the page where its range ends, as its leader starts there.
  const repeats = (element: Element) => element.closest(`.${cssPrefix}header, .${cssPrefix}footer`) !== null;
  const lastPage = new Map<string, number>();
  pages.forEach((page, index) => anchored[index].forEach((element) => {
    if (repeats(element)) return;
    if (isBlock(element) && element.dataset.sourceAnchorId) lastPage.set(`p:${element.dataset.sourceAnchorId}`, index);
    const root = thread.get(element.dataset.commentId ?? "");
    if (root !== undefined && element.matches(`span.${commentPrefix}highlight`)) lastPage.set(`c:${root}`, index);
  }));

  const seen = new Set<string>();
  const printed = new Set<string>();
  pages.forEach((page, pageIndex) => {
    const sheet = new MarkupSheet(document, page, cssPrefix);
    const entries: Entry[] = [];
    const marks: Mark[] = [];
    let order = 0;
    let lastRun: { element: HTMLElement; key: string } | undefined;
    const pendingBlocks: Array<{ block: HTMLElement; entry: Entry }> = [];
    const leaveBlocksBefore = (element?: HTMLElement) => {
      while (pendingBlocks.length > 0 && !(element && pendingBlocks[pendingBlocks.length - 1].block.contains(element))) {
        pendingBlocks.pop()!.entry.order = order++;
      }
    };

    const printedHere = new Set<string>();
    for (const element of anchored[pageIndex]) {
      // A paragraph's balloon follows the balloons of the runs inside it.
      leaveBlocksBefore(element);
      if (element.dataset.commentId !== undefined) {
        const rootId = thread.get(element.dataset.commentId);
        if (rootId === undefined) continue;
        const everyPage = repeats(element);
        if (!everyPage) seen.add(rootId);
        if (printedHere.has(rootId) || (!everyPage && (printed.has(rootId) || (lastPage.get(`c:${rootId}`) ?? pageIndex) !== pageIndex))) continue;
        const spans = rangeSpans(page.element, rootId).filter((span) => repeats(span) === everyPage);
        // A point or empty-range comment has only its (hidden) reference mark: anchor on its line.
        const first = sheet.box(spans.length > 0 ? lines(spans[0])[0] : markerLine(document, element));
        const last = sheet.box(spans.length > 0 ? lines(spans[spans.length - 1]).pop() : markerLine(document, element));
        if (!first || !last) continue;
        (everyPage ? printedHere : printed).add(rootId);
        for (const comment of comments.filter((candidate) => thread.get(candidate.id) === rootId)) {
          const color = colorOf(comment);
          const reply = comment.parentId !== undefined ? `R${comments.indexOf(byId.get(comment.parentId)!) + 1}` : "";
          entries.push({
            label: `Commented [${comment.initials || initialsOf(comment.author)}${comments.indexOf(comment) + 1}${reply}]:`,
            text: comment.paragraphs, color, fill: paleFill(color), textFont: comment.fontFamily, order: order++,
            lineTop: last.top, from: { x: last.right, y: last.bottom }, height: 0, top: 0, commentId: comment.id,
          });
        }
        if (spans.length > 0) {
          // Word brackets the commented range.
          const color = colorOf(byId.get(rootId)!);
          marks.push({ d: `M ${first.left} ${first.top} V ${first.bottom}`, color, filled: false });
          marks.push({ d: `M ${last.right} ${last.top} V ${last.bottom}`, color, filled: false });
        }
        continue;
      }

      const wording = element.dataset.formatDescription ?? "";
      const color = element.dataset.authorColor ?? WORD.palette[0];
      const base = { label: wording ? "Formatted:" : "Formatted", text: wording ? [wording] : [], color, fill: "#FFFFFF" };
      if (isBlock(element)) {
        const anchorId = element.dataset.sourceAnchorId;
        const line = sheet.box(lines(element, document).pop());
        if ((anchorId !== undefined && !repeats(element) && lastPage.get(`p:${anchorId}`) !== pageIndex) || !line) continue;
        const y = (line.top + line.bottom) / 2;
        // Marked at the text column's edge, level with the paragraph's last line.
        marks.push({ d: arrowhead(sheet.columnRight, y, "left"), color, filled: true });
        const entry: Entry = { ...base, order: 0, lineTop: line.top, from: { x: sheet.columnRight + WORD.markSize, y }, height: 0, top: 0 };
        entries.push(entry);
        pendingBlocks.push({ block: element, entry });
        continue;
      }
      // One formatting revision, one balloon, however many runs, list numbers and untracked
      // paragraph marks it spans.
      const key = [color, wording, element.dataset.formatKey].join("\n");
      if (lastRun?.key === key && onlyMarksBetween(document, lastRun.element, element)) {
        lastRun.element = element;
        continue;
      }
      lastRun = { element, key };
      const box = sheet.box(lines(element)[0]);
      if (!box) continue;
      // Marked under the change's first character.
      marks.push({ d: arrowhead(box.left, box.bottom, "up"), color, filled: true });
      entries.push({ ...base, order: order++, lineTop: box.top, from: { x: box.left, y: box.bottom }, height: 0, top: 0 });
    }
    leaveBlocksBefore();

    if (entries.length > 0) {
      orderByLine(entries);
      const left = sheet.columnRight + WORD.balloonStart;
      for (const entry of entries) {
        entry.element = balloon(document, entry, font, cssPrefix, left, sheet.width - WORD.balloonEnd - left);
        sheet.layer.append(entry.element);
        entry.height = sheet.height(entry.element);
      }
      const top = sheet.pageTop + WORD.pageInset;
      const bottom = sheet.pageTop + page.dimensions.pageHeight - WORD.pageInset;
      fit(entries, top, bottom, font.sizePt * WORD.leading,
        (entry) => cutShort(document, entry, font, cssPrefix, sheet), (entry) => restore(entry, cssPrefix, sheet));
      stack(entries, top, bottom);
      for (const entry of entries) {
        entry.element!.style.top = `${entry.top}pt`;
        // Past the page's foot a balloon is cut off by the paper's edge.
        if (entry.top + entry.height > bottom + 0.01) entry.element!.dataset.balloonTruncated = "true";
        if (entry.element!.dataset.balloonTruncated !== "true") continue;
        if (entry.commentId === undefined) result.truncatedFormatting++;
        else if (!result.truncatedCommentIds.includes(entry.commentId)) result.truncatedCommentIds.push(entry.commentId);
      }
      connect(document, sheet, entries, marks, left);
    }
    sheet.shrink();
  });
  result.unplacedCommentIds = comments.filter((comment) => seen.has(thread.get(comment.id)!) && !printed.has(thread.get(comment.id)!))
    .map((comment) => comment.id);
  return result;
}

/**
 * The page box's content, moved into one layer that also holds the markup area. The layer spans the
 * page up to its text column's right edge plus the markup area, and is shrunk to the paper's width
 * with the page standing at the paper's foot. Unscaled, it reaches up from the page box's foot rather
 * than down past it: print fragmentation would push anything laid out below the foot onto a new page.
 */
class MarkupSheet {
  readonly layer: HTMLElement;
  readonly columnRight: number;
  readonly width: number;
  readonly layerHeight: number;
  readonly pageTop: number;
  private readonly factor: number;

  constructor(document: Document, page: MarkupBalloonPage, cssPrefix: string) {
    const { pageWidth: paperWidth, pageHeight: paperHeight, marginRight } = page.dimensions;
    this.columnRight = paperWidth - marginRight;
    this.width = Math.max(paperWidth, this.columnRight + WORD.areaWidth);
    this.factor = paperWidth / this.width;
    this.layerHeight = paperHeight / this.factor;
    this.pageTop = this.layerHeight - paperHeight - WORD.paperFootGap / this.factor;

    const placed = (className: string, left: number, top: number, width: number, height: number) => {
      const element = document.createElement("div");
      element.className = className;
      element.style.cssText = `position:absolute;left:${left}pt;top:${top}pt;width:${width}pt;height:${height}pt`;
      return element;
    };
    this.layer = placed(`${cssPrefix}markup-layer`, 0, paperHeight - this.layerHeight, this.width, this.layerHeight);
    this.layer.style.transformOrigin = "left bottom";
    const pageHolder = placed(`${cssPrefix}markup-page`, 0, this.pageTop, paperWidth, paperHeight);
    pageHolder.append(...Array.from(page.element.childNodes));
    const area = placed(`${cssPrefix}markup-area`, this.columnRight + WORD.greyStart, this.pageTop,
      this.width - this.columnRight - WORD.greyStart, paperHeight);
    area.style.background = WORD.grey;
    area.dataset.pageMapExclude = "true";
    area.setAttribute("aria-hidden", "true");
    this.layer.append(pageHolder, area);
    page.element.append(this.layer);
    page.element.dataset.docxMarkupScale = String(this.factor);
  }

  /** Points per CSS pixel in the unscaled layer (the viewer may zoom the page box). */
  private get pointsPerPixel(): number {
    return this.width / this.layer.getBoundingClientRect().width;
  }

  /** A client rect in the layer's own points. */
  box(rect: DOMRect | undefined) {
    if (!rect) return undefined;
    const origin = this.layer.getBoundingClientRect();
    const k = this.pointsPerPixel;
    return {
      left: (rect.left - origin.left) * k, right: (rect.right - origin.left) * k,
      top: (rect.top - origin.top) * k, bottom: (rect.bottom - origin.top) * k,
    };
  }

  height(element: HTMLElement): number {
    return element.getBoundingClientRect().height * this.pointsPerPixel;
  }

  shrink(): void {
    this.layer.style.transform = `scale(${this.factor})`;
  }
}

function balloon(document: Document, entry: Entry, font: { family: string; sizePt: number }, cssPrefix: string,
  left: number, width: number): HTMLElement {
  const element = document.createElement("div");
  element.className = `${cssPrefix}balloon`;
  element.dataset.pageMapExclude = "true";
  // The label and any formatting wording are set in the balloon font; a comment keeps its own.
  element.style.cssText = `left:${left}pt;width:${width}pt;font-size:${font.sizePt}pt;border-color:${entry.color};` +
    `background-color:${entry.fill};font-family:${fontStack(entry.textFont ?? font.family)}`;
  const content = document.createElement("div");
  const label = document.createElement("b");
  label.style.fontFamily = fontStack(font.family);
  label.textContent = entry.label;
  content.append(label, ...entry.text.flatMap((paragraph, index) =>
    index === 0 ? [` ${paragraph}`] : [document.createElement("br"), paragraph]));
  element.append(content);
  return element;
}

/**
 * Which balloons to cut short when a page's cannot all stand level with their lines. A stack the
 * page's height can hold is never cut: `stack` pulls it up from the foot as far as the room above
 * allows, as Word does with a balloon near the foot. When even that cannot hold it, Word keeps
 * balloons whole from the top for as long as the rest, each cut to its first line, still follow at
 * or below their own lines and end above the page's foot; the first balloon that cannot stay whole,
 * and every one after it, is cut, and none is moved above its line to keep another whole. Measured
 * on a page of 25 long comments: 3 whole, 22 cut. A balloon of one line is never cut.
 */
function fit(entries: Entry[], top: number, bottom: number, lineHeight: number,
  shorten: (entry: Entry) => void, lengthen: (entry: Entry) => void): void {
  const whole = entries.map((entry) => entry.height);
  if (whole.reduce((sum, height) => sum + height + WORD.between, -WORD.between) <= bottom - top + 0.01) return;
  const oneLine = lineHeight + 2 * (WORD.padBlock + WORD.border);
  const cuttable = whole.map((height) => height > oneLine + lineHeight / 2);
  entries.forEach((entry, index) => { if (cuttable[index]) shorten(entry); });
  const cut = entries.map((entry) => entry.height);
  const stackedFoot = (heightOf: (index: number) => number) => entries.reduce((floor, entry, index) =>
    Math.max(floor, entry.lineTop) + heightOf(index) + WORD.between, top) - WORD.between;
  const fits = (keptWhole: number) => stackedFoot((index) => (index < keptWhole ? whole[index] : cut[index])) <= bottom + 0.01;
  let keptWhole = 0;
  while (keptWhole < entries.length && fits(keptWhole + 1)) keptWhole++;
  for (let index = 0; index < keptWhole; index++) if (cuttable[index]) lengthen(entries[index]);
}

/**
 * Cuts a balloon to its first line, as Word does: the line's text stops short of a small white box
 * at the balloon's lower right corner that holds an ellipsis, and the balloon takes Word's height for
 * a cut balloon. The balloon becomes a two-column grid, the line and the box, and the box's negative
 * margins carry its edges onto the balloon's padding edge. It stays in the flow: positioned
 * absolutely, it made Chromium shrink every printed page.
 */
function cutShort(document: Document, entry: Entry, font: { family: string; sizePt: number }, cssPrefix: string,
  sheet: MarkupSheet): void {
  const element = entry.element!;
  const scale = font.sizePt / WORD.font.sizePt;
  let box = element.querySelector<HTMLElement>(`:scope > .${cssPrefix}balloon-cut`);
  if (!box) {
    box = document.createElement("span");
    box.className = `${cssPrefix}balloon-cut`;
    box.textContent = "\u2026";
    element.append(box);
  }
  const boxHeight = WORD.cut.boxHeight * scale;
  box.style.cssText = `align-self:end;box-sizing:border-box;height:${boxHeight}pt;` +
    `margin:0 -${WORD.padInline}pt -${WORD.padBlock}pt 0;line-height:${boxHeight - 2 * WORD.hairline}pt;` +
    `padding:0 ${WORD.cut.boxPadInline * scale}pt;border:${WORD.hairline}pt solid #000;background:#fff;color:#000;` +
    `white-space:nowrap;font-size:${font.sizePt}pt;font-family:${fontStack(font.family)}`;
  element.style.display = "grid";
  element.style.gridTemplateColumns = "minmax(0, 1fr) auto";
  element.style.minHeight = `${WORD.cut.height * scale}pt`;
  // The text stops the gap short of the box; then the first line's own box, whose height the label's
  // font and the text's share, is all that stays visible.
  const content = element.firstElementChild as HTMLElement;
  content.style.cssText = `align-self:start;margin-right:${WORD.cut.textGap * scale}pt`;
  const rows = lines(content, document).map((rect) => sheet.box(rect)!);
  const onFirst = rows.filter((row) => row.top < rows[0].bottom - 0.5);
  const firstLine = Math.max(...onFirst.map((row) => row.bottom)) - Math.min(...onFirst.map((row) => row.top));
  content.style.cssText += `;max-height:${firstLine}pt;overflow:hidden`;
  element.dataset.balloonTruncated = "true";
  entry.height = sheet.height(element);
}

/** Undoes `cutShort`. */
function restore(entry: Entry, cssPrefix: string, sheet: MarkupSheet): void {
  const element = entry.element!;
  element.querySelector(`:scope > .${cssPrefix}balloon-cut`)?.remove();
  element.style.display = "";
  element.style.gridTemplateColumns = "";
  element.style.minHeight = "";
  (element.firstElementChild as HTMLElement).style.cssText = "";
  delete element.dataset.balloonTruncated;
  entry.height = sheet.height(element);
}

/**
 * Down the page: no balloon above its line or over the previous one. Then up from the foot as far as
 * the room above allows, and down once more so that balloons that still do not fit run off the foot
 * rather than overlap.
 */
function stack(entries: Entry[], top: number, bottom: number): void {
  const pushDown = () => entries.reduce((floor, entry) =>
    (entry.top = Math.max(entry.top, floor)) + entry.height + WORD.between, top);
  for (const entry of entries) entry.top = entry.lineTop;
  pushDown();
  entries.reduceRight((ceiling, entry) =>
    (entry.top = Math.max(top, Math.min(entry.top, ceiling - entry.height))) - WORD.between, bottom);
  pushDown();
}

/** Hairline leaders, one per balloon, and the range brackets and change marks. */
function connect(document: Document, sheet: MarkupSheet, entries: Entry[], marks: Mark[], balloonLeft: number): void {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", `0 0 ${sheet.width} ${sheet.layerHeight}`);
  svg.setAttribute("style",
    `position:absolute;left:0;top:0;width:${sheet.width}pt;height:${sheet.layerHeight}pt;overflow:visible;pointer-events:none`);
  for (const name of ["aria-hidden", "data-page-map-exclude"]) svg.setAttribute(name, "true");
  const elbow = sheet.columnRight + WORD.elbow;
  const paths = [
    ...entries.map(({ from, top, color }) => ({
      d: `M ${from.x} ${from.y} H ${elbow} L ${balloonLeft} ${top + WORD.leaderDrop}`, color, filled: false, dotted: true,
    })),
    ...marks.map((mark) => ({ ...mark, dotted: false })),
  ];
  for (const { d, color, filled, dotted } of paths) {
    const path = document.createElementNS(ns, "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", filled ? color : "none");
    path.setAttribute("stroke", color);
    path.setAttribute("stroke-width", String(WORD.hairline));
    if (dotted) path.setAttribute("stroke-dasharray", String(WORD.hairline));
    svg.append(path);
  }
  sheet.layer.append(svg);
}

/** A small filled triangle: under a character pointing up, or at the column's edge pointing left. */
function arrowhead(x: number, y: number, pointing: "up" | "left"): string {
  const s = WORD.markSize;
  return pointing === "up"
    ? `M ${x - s / 2} ${y} L ${x} ${y - s} L ${x + s / 2} ${y} Z`
    : `M ${x} ${y} L ${x + s} ${y + s / 2} V ${y - s / 2} Z`;
}

/** Line by line down the page; balloons of one line keep document order. */
function orderByLine(entries: Entry[]): void {
  const lineTops: number[] = [];
  const lineOf = (entry: Entry) => {
    const known = lineTops.find((top) => Math.abs(top - entry.lineTop) <= SAME_LINE);
    if (known !== undefined) return known;
    lineTops.push(entry.lineTop);
    return entry.lineTop;
  };
  const keyed = entries.map((entry) => ({ entry, line: lineOf(entry) }));
  keyed.sort((a, b) => a.line - b.line || a.entry.order - b.entry.order);
  keyed.forEach(({ entry }, index) => (entries[index] = entry));
}

function isBlock(element: Element): boolean {
  return ["p", "li", "h1", "h2", "h3", "h4", "h5", "h6"].includes(element.localName);
}

/**
 * Whether only list numbers and blank text lie between two changes, with no tracked paragraph mark:
 * then Word treats them as one revision.
 */
function onlyMarksBetween(document: Document, earlier: HTMLElement, later: HTMLElement): boolean {
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
  walker.currentNode = earlier;
  for (let node = walker.nextNode(); node && node !== later; node = walker.nextNode()) {
    if (earlier.contains(node)) continue;
    if (node instanceof Element) {
      if (isBlock(node) && /(^|\s)\S*para-(ins|del)(\s|$)/.test(node.className)) return false;
    } else if (node.textContent?.trim() && !node.parentElement?.closest("[data-list-marker]")) {
      return false;
    }
  }
  return true;
}

function initialsOf(author = ""): string {
  return (author.match(/(?:^|\s)(\S)/g) ?? []).map((initial) => initial.trim().toUpperCase()).join("");
}

/** A comment's fill: Word's own for its first two colours, otherwise the colour faded towards white. */
function paleFill(color: string): string {
  return WORD.paleFills[color.toUpperCase()] ?? `color-mix(in srgb, ${color} 17%, white)`;
}

function fontStack(family: string): string {
  return `${JSON.stringify(family)}, sans-serif`;
}

/**
 * The line a hidden reference mark sits on: the mark's own position, with the height of the text
 * line around it (the mark itself is drawn at zero size).
 */
function markerLine(document: Document, marker: Element): DOMRect | undefined {
  const at = marker.getBoundingClientRect();
  const block = marker.closest("p, li, h1, h2, h3, h4, h5, h6") ?? marker.parentElement;
  if (!block) return undefined;
  const middle = at.top + at.height / 2;
  const rows = lines(block, document);
  const row = rows.find((rect) => middle >= rect.top - 1 && middle <= rect.bottom + 1)
    ?? rows.reduce<DOMRect | undefined>((best, rect) =>
      !best || Math.abs(rect.top - middle) < Math.abs(best.top - middle) ? rect : best, undefined);
  return row && new DOMRect(at.left, row.top, 0, row.height);
}

/** The line boxes of an element, or of its text when given the document (a block's own box is one rect). */
function lines(element: Element, document?: Document): DOMRect[] {
  let target: Element | Range = element;
  if (document) {
    target = document.createRange();
    target.selectNodeContents(element);
  }
  return Array.from(target.getClientRects()).filter((rect) => rect.height > 0);
}

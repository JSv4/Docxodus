/**
 * The full text of comments whose balloons were cut short, printed on pages after the document's
 * last, as Word's Print to PDF does with All Markup balloons (`balloonOverflow: "appendix"`).
 *
 * Each comment is headed by a bar in its balloon's colours naming the page its balloon is on, the
 * number in its box, its label, its author and the date it was written; its paragraphs follow across
 * the text column, blank ones as blank lines. Comments run on from page to page: a paragraph splits
 * at a line break only where two lines stay on each side, and a bar may close a page alone. The
 * geometry below was measured from Word's Print to PDF of a page of 25 long comments, in points.
 */
import { DEFAULT_BALLOON_FONT, fontStack, type CutComment, type MarkupBalloonPage } from "./markup-balloons.js";

const APPENDIX = {
  /** The bar's outer edges stand this far outside the text column on either side. */
  barOutset: 4.92,
  barBorder: 0.48,
  /** Its shadow, cast right and down in the comment's colour. */
  barShadow: 0.96,
  /** Border box height, the borders included. */
  barHeight: 14.88,
  /** The date ends this far inside the text column. */
  dateInset: 4,
  /** From the end of one comment's text to the next bar, and from a bar to its comment's text. */
  spaceBeforeBar: 5.9,
  spaceAfterBar: 2.2,
  /** Word's Comment Text size. */
  textSizePt: 10,
  leading: 1.15,
  /** Widow and orphan control: lines a paragraph keeps on each side of a page break. */
  minLines: 2,
};

type Block = { comment: CutComment; number: number } & ({ kind: "bar" } | { kind: "text"; text: string });

/**
 * Lays the comments out a page at a time. The caller hands `fill` blank pages, in order, until
 * `finished` is true; every call places at least one block, so the loop ends.
 */
export class BalloonAppendix {
  private readonly queue: Block[];

  constructor(
    private readonly document: Document,
    comments: CutComment[],
    private readonly cssPrefix: string,
    private readonly font: { family: string; sizePt: number } = DEFAULT_BALLOON_FONT,
  ) {
    this.queue = comments.flatMap((comment, index): Block[] => [
      { kind: "bar", comment, number: index + 1 },
      ...comment.paragraphs.map((text): Block => ({ kind: "text", comment, number: index + 1, text })),
    ]);
  }

  get finished(): boolean {
    return this.queue.length === 0;
  }

  fill(page: MarkupBalloonPage): void {
    const { marginLeft, marginTop, contentWidth, contentHeight } = page.dimensions;
    const column = this.document.createElement("div");
    column.className = `${this.cssPrefix}balloon-appendix`;
    column.dataset.pageMapExclude = "true";
    column.style.cssText = `position:absolute;left:${marginLeft}pt;top:${marginTop}pt;width:${contentWidth}pt`;
    page.element.append(column);
    const origin = column.getBoundingClientRect();
    const k = contentWidth / origin.width;
    const bottomOf = (element: Element) => (element.getBoundingClientRect().bottom - origin.top) * k;

    while (this.queue.length > 0) {
      const block = this.queue[0];
      const previous = column.lastElementChild as HTMLElement | null;
      const element = block.kind === "bar" ? this.bar(block) : this.paragraph(block.comment, block.text);
      element.style.marginTop = `${!previous ? 0 : block.kind === "bar" ? APPENDIX.spaceBeforeBar
        : previous.dataset.appendixBar === "true" ? APPENDIX.spaceAfterBar : 0}pt`;
      column.append(element);
      if (bottomOf(element) <= contentHeight + 0.01) {
        this.queue.shift();
        continue;
      }
      if (block.kind === "bar") {
        // A bar that does not fit opens the next page, unless the page holds nothing else.
        if (previous) element.remove();
        else this.queue.shift();
        return;
      }
      const rows = lineBoxes(this.document, element);
      const room = rows.filter((row) => (row.bottom - origin.top) * k <= contentHeight + 0.01).length;
      // Two lines stay on each side of the break; alone on its page, a paragraph keeps what fits.
      let keep = Math.min(room, rows.length - APPENDIX.minLines);
      if (!previous) keep = Math.max(1, keep);
      let split = -1;
      // Rows start wherever Chromium broke the line, inside a word too (after a hyphen, or in a long
      // address). Cut there, then step back a row while the head still runs past the column.
      for (; keep >= (previous ? APPENDIX.minLines : 1); keep--) {
        element.textContent = block.text;
        split = rowStart(this.document, element, rows[keep]);
        if (split <= 0) continue;
        element.textContent = block.text.slice(0, split).trimEnd();
        if (bottomOf(element) <= contentHeight + 0.01) break;
        split = -1;
      }
      if (split > 0) {
        this.queue[0] = { ...block, text: block.text.slice(split).trimStart() };
      } else if (previous) {
        element.remove();
      } else {
        // One line taller than the page: it prints as it is.
        this.queue.shift();
      }
      return;
    }
  }

  /** The heading: page, box number and label at the left, the author centred, the date at the right. */
  private bar({ comment, number }: Block): HTMLElement {
    const bar = this.document.createElement("div");
    bar.className = `${this.cssPrefix}balloon-appendix-bar`;
    bar.dataset.appendixBar = "true";
    const inset = APPENDIX.barOutset - APPENDIX.barBorder;
    bar.style.cssText = `display:grid;grid-template-columns:minmax(0,1fr) auto minmax(0,1fr);align-items:center;` +
      `column-gap:${APPENDIX.dateInset}pt;box-sizing:border-box;height:${APPENDIX.barHeight}pt;` +
      `margin-left:-${APPENDIX.barOutset}pt;margin-right:-${APPENDIX.barOutset}pt;padding:0 ${inset}pt;` +
      `border:${APPENDIX.barBorder}pt solid ${comment.color};` +
      `box-shadow:${APPENDIX.barShadow}pt ${APPENDIX.barShadow}pt 0 ${comment.color};background:${comment.fill};` +
      `color:#000;font:bold ${this.font.sizePt}pt/1 ${fontStack(this.font.family)};white-space:nowrap`;
    const cell = (text: string, css: string) => {
      const span = this.document.createElement("span");
      span.textContent = text;
      span.style.cssText = `overflow:hidden;${css}`;
      return span;
    };
    bar.append(
      cell(`Page ${comment.pageNumber}: [${number}] ${comment.label}`, "justify-self:start"),
      cell(comment.author ?? "", ""),
      cell(printedDate(comment.dateAsWritten), `justify-self:end;padding-right:${APPENDIX.dateInset}pt`),
    );
    return bar;
  }

  private paragraph(comment: CutComment, text: string): HTMLElement {
    const paragraph = this.document.createElement("div");
    paragraph.className = `${this.cssPrefix}balloon-appendix-text`;
    const lineHeight = APPENDIX.textSizePt * APPENDIX.leading;
    paragraph.style.cssText = `min-height:${lineHeight}pt;color:#000;overflow-wrap:anywhere;` +
      `font-size:${APPENDIX.textSizePt}pt;line-height:${lineHeight}pt;font-family:${fontStack(comment.textFont ?? this.font.family)}`;
    paragraph.textContent = text;
    return paragraph;
  }
}

/** A paragraph's line boxes, top to bottom; a blank paragraph is one line. */
function lineBoxes(document: Document, element: HTMLElement): DOMRect[] {
  const range = document.createRange();
  range.selectNodeContents(element);
  const rows: DOMRect[] = [];
  for (const rect of Array.from(range.getClientRects())) {
    if (rect.height <= 0) continue;
    const row = rows.find((known) => Math.abs(known.top - rect.top) < rect.height / 2);
    if (!row) rows.push(rect);
  }
  rows.sort((a, b) => a.top - b.top);
  return rows.length > 0 ? rows : [element.getBoundingClientRect()];
}

/**
 * The offset in a paragraph's text of the first character laid out on `row` or below it: characters
 * run down the lines in order, so a binary search over them finds it. -1 when there is none.
 */
function rowStart(document: Document, element: HTMLElement, row: DOMRect | undefined): number {
  const text = element.firstChild;
  const length = text?.textContent?.length ?? 0;
  if (!text || !row) return -1;
  const range = document.createRange();
  const below = (offset: number) => {
    range.setStart(text, offset);
    range.setEnd(text, offset + 1);
    const rect = Array.from(range.getClientRects()).find((candidate) => candidate.width > 0) ?? range.getBoundingClientRect();
    return rect.top >= row.top - row.height / 2;
  };
  let low = 0;
  let high = length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (below(middle)) high = middle;
    else low = middle + 1;
  }
  return low < length ? low : -1;
}

/**
 * w:date as Word prints it in the heading, M/D/YYYY h:mm:ss AM/PM, its digits as written: Word
 * stores the author's clock time there with a Z suffix it does not honour.
 */
function printedDate(written = ""): string {
  const digits = written.replace(/(Z|[+-]\d{2}:\d{2})$/, "");
  const at = new Date(`${digits}Z`);
  if (!/^\d{4}-\d{2}-\d{2}T/.test(digits) || Number.isNaN(at.getTime())) return "";
  return at.toLocaleString("en-US", {
    timeZone: "UTC", year: "numeric", month: "numeric", day: "numeric", hour: "numeric", minute: "2-digit", second: "2-digit",
  }).replace(",", "").replace(/\s+/g, " ");
}

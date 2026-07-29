/**
 * DocxEditor — a framework-agnostic, in-browser DOCX block editor.
 *
 * Architecture (see docs/architecture/ir_editor_feasibility.md, "Option B"):
 *   - model-of-record: a live DocxSession in WASM (lossless save);
 *   - rendering: WmlToHtmlConverter HTML (faithful) stamped with data-anchor;
 *   - editing: each block is contenteditable; on commit, the edit goes through
 *     DocxSession by anchor, then ONLY that block is re-rendered from the live
 *     session (session-attached RenderBlockHtml) and patched into the DOM.
 *
 * The IR/anchor system is the addressing spine; the live OOXML is the truth.
 * This is the pure-TypeScript core; a React wrapper can sit on top.
 *
 * MVP scope: per-block, commit-on-blur editing of paragraphs/headings. An edited
 * block's content is replaced from its plain text (inline formatting within an
 * edited block is not preserved — a documented MVP limit); UNTOUCHED blocks keep
 * full fidelity, and save() is lossless for them.
 */

import { paginateHtml } from "./pagination.js";
import { HeaderFooterRegion } from "./editor-headerfooter.js";
import type { BandWhich } from "./editor-headerfooter.js";
import type { HeaderFooterKind } from "./types.js";

/** The subset of WASM bridge exports the editor needs (as exposed on `window.Docxodus`). */
export interface DocxEditorExports {
  DocxSessionBridge: {
    OpenSession: (bytes: Uint8Array, settingsJson: string) => number;
    CloseSession: (handle: number) => void;
    CreateBlankDocx: () => Uint8Array;
    Project: (handle: number) => string;
    ReplaceText: (handle: number, anchor: string, md: string) => string;
    ReplaceTextAtSpan: (
      handle: number,
      anchor: string,
      spanStart: number,
      spanLength: number,
      replace: string,
    ) => string;
    SplitParagraph: (handle: number, anchor: string, offset: number) => string;
    MergeParagraphs: (handle: number, first: string, second: string) => string;
    DeleteBlock: (handle: number, anchor: string) => string;
    InsertHorizontalRule: (handle: number, anchor: string, pos: string, ruleJson: string) => string;
    InsertTable: (
      handle: number,
      anchor: string,
      pos: string,
      rows: number,
      cols: number,
      optionsJson: string,
    ) => string;
    InsertTableRow: (handle: number, cellAnchor: string, pos: string) => string;
    InsertTableColumn: (handle: number, cellAnchor: string, pos: string) => string;
    DeleteTableRow: (handle: number, cellAnchor: string) => string;
    DeleteTableColumn: (handle: number, cellAnchor: string) => string;
    ApplyFormat: (handle: number, anchor: string, spanJson: string, opJson: string) => string;
    SetParagraphStyle: (handle: number, anchor: string, styleId: string) => string;
    SetParagraphFormat: (handle: number, anchor: string, opJson: string) => string;
    ApplyListFormat: (handle: number, anchor: string, kind: string) => string;
    SetListLevel: (handle: number, anchor: string, delta: number) => string;
    GetListMembership: (handle: number, anchor: string) => string;
    RenderBlockHtml: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    /** Session-attached full-document render (optional: older WASM bundles lack it). */
    RenderHtml?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
    ) => string;
    Save: (handle: number) => Uint8Array;
    Undo: (handle: number) => boolean;
    Redo: (handle: number) => boolean;
    /** Header/footer region: the section a body anchor belongs to (kind → part mapping). */
    GetSectionInfo: (handle: number, anchorId: string) => string;
    SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    InsertPageNumberField: (handle: number, anchor: string, field: string) => string;
    EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
  };
  DocumentConverter: {
    ConvertDocxToHtmlComplete: (...args: any[]) => string;
  };
}

export interface DocxEditorOptions {
  /** CSS class prefix for rendered HTML. Default "docx-". */
  cssPrefix?: string;
  /**
   * Fabricate CSS classes (vs inline styles). Default FALSE for the editor: a per-block
   * re-render must be self-contained, but fabricated class names are per-conversion and
   * have no matching stylesheet on the page, so re-rendered blocks would lose styling.
   * Inline styles keep every block's formatting intact on incremental re-render.
   */
  fabricateClasses?: boolean;
  /** Make paragraph/heading blocks editable. Default true. */
  editable?: boolean;
  /** Render block-flow pages (page boxes via pagination.ts) vs a continuous view. Default false. */
  paginated?: boolean;
  /** Page render scale for paginated mode (1.0 = 100%). Default 1. */
  scale?: number;
  /**
   * Render docked Header/Footer editing bands around the body flow. Default FALSE — with it off
   * the editor's DOM is unchanged, so existing consumers are unaffected. When on, the body flow
   * is wrapped in a `.docx-body-flow` element that becomes the edit root, keeping band blocks out
   * of the body's block list (which indexes remount focus).
   */
  headerFooter?: boolean;
  /** Called after a block edit commits (with the affected anchor). */
  onEdit?: (info: { anchorId: string; unid: string }) => void;
}

interface AnchorTargetLite {
  unid: string;
  kind: string;
  scope: string;
  textPreview?: string;
}

const EDITABLE_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6"]);

// ─── M1: inline HTML → markdown (preserve formatting on edit) ───────────────

interface InlineSeg {
  text: string;
  bold: boolean;
  italic: boolean;
  href: string | null;
}

function fontWeightIsBold(w: string): boolean {
  if (w === "bold" || w === "bolder") return true;
  const n = parseInt(w, 10);
  return !Number.isNaN(n) && n >= 600;
}

function escapeInlineMarkdown(text: string): string {
  // Escape the markdown the projector subset is sensitive to; keep it minimal.
  return text.replace(/([\\`*_[\]])/g, "\\$1");
}

function collectInlineSegments(node: Node, out: InlineSeg[]): void {
  node.childNodes.forEach((child) => {
    // Skip converter-generated chrome (list markers, note citation markers, note backrefs) —
    // it isn't part of the paragraph's content and must never be committed as text.
    if (isGeneratedChrome(child)) return;
    if (child.nodeType === 3 /* TEXT_NODE */) {
      const text = child.textContent ?? "";
      if (!text) return;
      const parent = child.parentElement;
      let bold = false;
      let italic = false;
      let href: string | null = null;
      if (parent && typeof getComputedStyle === "function") {
        const cs = getComputedStyle(parent);
        bold = fontWeightIsBold(cs.fontWeight);
        italic = cs.fontStyle === "italic" || cs.fontStyle === "oblique";
        const a = parent.closest("a");
        href = a ? a.getAttribute("href") : null;
      }
      out.push({ text, bold, italic, href });
    } else if (child.nodeType === 1 /* ELEMENT_NODE */) {
      const el = child as HTMLElement;
      if (el.tagName === "BR") {
        out.push({ text: "\n", bold: false, italic: false, href: null });
        return;
      }
      collectInlineSegments(el, out);
    }
  });
}

function segToMarkdown(seg: InlineSeg): string {
  // A <br> segment is a hard line break → the canonical GFM "  \n", which the
  // DocxSession markdown parser turns into a real w:br (Word's intra-paragraph
  // line break) instead of a literal newline in w:t.
  if (seg.text === "\n") return "  \n";
  let md = escapeInlineMarkdown(seg.text).replace(/[ \t]*\n/g, "  \n");
  if (/\S/.test(seg.text)) {
    // Don't wrap pure whitespace — `** **` is not valid emphasis.
    if (seg.bold && seg.italic) md = `***${md}***`;
    else if (seg.bold) md = `**${md}**`;
    else if (seg.italic) md = `*${md}*`;
  }
  if (seg.href) md = `[${md}](${seg.href})`;
  return md;
}

/**
 * Serialize a block's inline content to the projector's markdown subset, preserving
 * bold / italic / links (emphasis detected via computed style). Used so an edit keeps
 * the block's formatting instead of flattening it to plain text. Formatting the markdown
 * subset cannot express (font size/color) is still dropped on an edited block.
 */
export function serializeInlineMarkdown(block: HTMLElement): string {
  const segs: InlineSeg[] = [];
  collectInlineSegments(block, segs);
  // Merge adjacent segments with identical formatting to avoid `**a****b**`.
  const merged: InlineSeg[] = [];
  for (const s of segs) {
    const prev = merged[merged.length - 1];
    if (
      prev &&
      prev.text !== "\n" &&
      s.text !== "\n" &&
      prev.bold === s.bold &&
      prev.italic === s.italic &&
      prev.href === s.href
    ) {
      prev.text += s.text;
    } else {
      merged.push({ ...s });
    }
  }
  return merged.map(segToMarkdown).join("").trim();
}

// ─── M2: structural editing (split / merge) ─────────────────────────────────

interface AnchorRef {
  id: string;
  kind: string;
  scope: string;
  unid: string;
}

interface EditResultLite {
  success: boolean;
  created?: AnchorRef[];
  removed?: AnchorRef[];
  modified?: AnchorRef[];
  error?: { message?: string };
}

/** True if `block` renders as a list item (has a generated marker as its first child). */
function isListBlock(block: HTMLElement): boolean {
  return !!block.querySelector(":scope > [data-list-marker]");
}

/**
 * Inline chrome the CONVERTER generates that is not part of a paragraph's run text: list
 * number/bullet markers, footnote/endnote citation markers
 * (`<a class="footnote-ref"><sup>1</sup></a>`), and the note backrefs (`↩`).
 *
 * The session's run text contains none of it. A citation is a zero-width
 * `w:footnoteReference` — the displayed number is computed by the renderer from document order —
 * so every character of chrome the editor fails to exclude shifts its content-offset space away
 * from the session's. Each omission has its own failure mode: excluded from offsets but not from
 * serialization and the display number gets COMMITTED as literal text (destroying the citation
 * run); left editable and the user can delete a marker outright, orphaning the note.
 */
const GENERATED_CHROME_SELECTOR =
  '[data-list-marker], a.footnote-ref, a.endnote-ref, a[class$="-backref"]';

function isGeneratedChrome(node: Node | null | undefined): boolean {
  return node?.nodeType === 1 && !!(node as Element).matches?.(GENERATED_CHROME_SELECTOR);
}

/** True if `node` is, or is inside, generated chrome (not editable content). */
function isInMarker(node: Node | null): boolean {
  let el: Element | null = node && node.nodeType === 1 ? (node as Element) : node?.parentElement ?? null;
  while (el) {
    if (isGeneratedChrome(el)) return true;
    el = el.parentElement;
  }
  return false;
}

/**
 * Unicode bidi formatting marks the HTML converter injects to preserve visual order: LRM/RLM/ALM,
 * the embedding/override controls, and the isolates (see WmlToHtmlConverter — a paragraph/run gets
 * a leading U+200E or U+200F). They are presentation-only — NOT part of the paragraph's run text the
 * session holds — so the editor must exclude them from its content-offset space, the same way it
 * excludes generated list markers. Otherwise every caret offset is shifted by the leading mark and a
 * caret at end-of-line overshoots the session's text length, so SplitParagraph/ApplyFormat reject the
 * offset (symptom: Enter at the end of a Google-Docs-exported paragraph is silently dropped).
 */
// LRM, RLM, ALM; the embedding/override controls (LRE RLE PDF LRO RLO); the isolates (LRI RLI FSI PDI).
const BIDI_MARK_CLASS = "\u200E\u200F\u061C\u202A-\u202E\u2066-\u2069";
const BIDI_MARKS_RE_G = new RegExp(`[${BIDI_MARK_CLASS}]`, "g");
const BIDI_MARK_RE = new RegExp(`[${BIDI_MARK_CLASS}]`);
function stripBidi(s: string): string {
  return s.replace(BIDI_MARKS_RE_G, "");
}

/** Raw string index in `s` for content offset `n` (content = chars excluding bidi marks). */
function domOffsetForContentOffset(s: string, n: number): number {
  let content = 0;
  for (let i = 0; i < s.length; i++) {
    if (content >= n) return i;
    if (!BIDI_MARK_RE.test(s[i])) content++;
  }
  return s.length;
}

/**
 * Content-text offset of (container, offset) within `block`, EXCLUDING generated list-marker
 * text and injected bidi marks. This is the offset DocxSession ops expect (the paragraph's run
 * text, not the rendered number/bullet or bidi marks the converter injects).
 */
function contentOffsetOf(block: HTMLElement, container: Node, offset: number): number {
  let count = 0;
  let done = false;
  const walk = (node: Node): void => {
    if (done) return;
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (node === container) {
        if (!isInMarker(node)) count += stripBidi((node.textContent ?? "").slice(0, offset)).length;
        done = true; return;
      }
      if (!isInMarker(node)) count += stripBidi(node.textContent ?? "").length;
    } else {
      if (node === container) {
        // Element container: `offset` is a child index — count content up to that child.
        const kids = Array.from(node.childNodes);
        for (let i = 0; i < offset && i < kids.length; i++) walk(kids[i]);
        done = true;
        return;
      }
      node.childNodes.forEach(walk);
    }
  };
  walk(block);
  return count;
}

/** Content-text offset of the collapsed caret within `block` (excludes markers), or null. */
function caretOffsetIn(block: HTMLElement): number | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!block.contains(range.startContainer)) return null;
  return contentOffsetOf(block, range.startContainer, range.startOffset);
}

/** Visible content text of `block`, excluding generated list-marker text (the same content
 *  caretOffsetIn/contentOffsetOf count). */
function blockContentText(block: HTMLElement): string {
  let out = "";
  const walk = (node: Node): void => {
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (!isInMarker(node)) out += stripBidi(node.textContent ?? "");
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(block);
  return out;
}

/**
 * Map a DOM caret offset (from caretOffsetIn) into the run-text offset the session holds after a
 * commit. commitBlock/syncBlock commit `serializeInlineMarkdown(el)`, which `.trim()`s leading and
 * trailing whitespace, so the session's paragraph text is shorter than the DOM text whenever the
 * block has edge whitespace — e.g. a blank document renders its empty paragraph with a placeholder
 * space, and typing lands after it. Without this adjustment the caret offset overshoots the
 * committed length, SplitParagraph returns OffsetOutOfRange, and splitAtCaret silently drops the
 * Enter (no new paragraph). Subtracting the leading whitespace before the caret and clamping to the
 * trimmed length keeps the split offset consistent with what was committed.
 */
function trimmedSplitOffset(block: HTMLElement, domOffset: number): number {
  const { leading, trimmedLen } = trimBounds(block);
  return Math.max(0, Math.min(domOffset - Math.min(domOffset, leading), trimmedLen));
}

/** How far `block`'s DOM content text is offset from, and longer than, its committed form. */
function trimBounds(block: HTMLElement): { leading: number; trimmedLen: number } {
  const content = blockContentText(block);
  return {
    leading: content.length - content.replace(/^\s+/, "").length,
    trimmedLen: content.trim().length,
  };
}

/**
 * Map a DOM content SPAN into the run-text space the session holds after a commit — the span
 * analogue of {@link trimmedSplitOffset}, and for the same reason.
 *
 * A block rendered with edge whitespace produces a span longer than the text the commit stores:
 * an empty header/footer story renders as a lone NBSP placeholder, so typing into it leaves a
 * trailing NBSP that `serializeInlineMarkdown(...).trim()` removes (JS `trim()` treats U+00A0 as
 * whitespace). "Select all, then Bold" then asks `ApplyFormat` for [0, len+1) — one past the
 * committed end — and the op is REJECTED with OffsetOutOfRange, so the format silently does
 * nothing. The demo's format buttons preventDefault on mousedown to keep the selection alive,
 * which is exactly the path that computes the span before `syncBlock` commits, so this is the
 * ordinary case rather than an edge one.
 */
function trimmedSpan(
  block: HTMLElement,
  span: { start: number; length: number },
): { start: number; length: number } {
  const { leading, trimmedLen } = trimBounds(block);
  const start = Math.max(0, Math.min(span.start - Math.min(span.start, leading), trimmedLen));
  const end = Math.max(start, Math.min(span.start + span.length - leading, trimmedLen));
  return { start, length: end - start };
}

/**
 * DOM (node, offset) for content offset `offset` within `el` — the same content-offset
 * space as contentOffsetOf (marker text and injected bidi marks excluded). Clamps past-end
 * offsets to the end of the last text node (or the element itself when it has none), so a
 * caller can always build a Range from the result.
 */
function contentPositionIn(el: HTMLElement, offset: number): { node: Node; offset: number } {
  let remaining = offset;
  let result: { node: Node; offset: number } | null = null;
  let lastText: Node | null = null;
  const walk = (node: Node): void => {
    if (result) return;
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (isInMarker(node)) return;
      const raw = node.textContent ?? "";
      const len = stripBidi(raw).length;
      lastText = node;
      if (remaining <= len) {
        result = { node, offset: domOffsetForContentOffset(raw, remaining) };
        return;
      }
      remaining -= len;
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(el);
  if (result) return result;
  if (lastText) return { node: lastText, offset: ((lastText as Node).textContent ?? "").length };
  return { node: el, offset: el.childNodes.length };
}

/** Place the caret at content offset `offset` within `el`, skipping marker text. */
function placeCaretAtOffset(el: HTMLElement, offset: number): void {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel) return;
  el.focus();
  const range = document.createRange();
  let remaining = offset;
  let placed = false;
  const walk = (node: Node): void => {
    if (placed) return;
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (isInMarker(node)) return; // never land the caret in the marker
      const raw = node.textContent ?? "";
      const len = stripBidi(raw).length; // content length excludes injected bidi marks
      if (remaining <= len) {
        range.setStart(node, domOffsetForContentOffset(raw, remaining));
        placed = true;
      } else {
        remaining -= len;
      }
    } else {
      node.childNodes.forEach(walk);
    }
  };
  walk(el);
  if (!placed) {
    range.selectNodeContents(el);
    range.collapse(false);
  } else {
    range.collapse(true);
  }
  sel.removeAllRanges();
  sel.addRange(range);
}

// ─── M5: formatting controls ────────────────────────────────────────────────

export type FormatKey = "bold" | "italic" | "underline" | "strike" | "code" | "superscript" | "subscript";

/** Paragraph alignment passed to DocxEditor.setAlignment. */
export type EditorAlignment = "left" | "center" | "right" | "justify";

/** The selection's content-text {start,length} within `block` (excludes markers), or null. */
function selectionSpanIn(block: HTMLElement): { start: number; length: number } | null {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (range.collapsed) return null;
  if (!block.contains(range.startContainer) || !block.contains(range.endContainer)) return null;
  const start = contentOffsetOf(block, range.startContainer, range.startOffset);
  const end = contentOffsetOf(block, range.endContainer, range.endOffset);
  // Normalized into the committed run-text space: every consumer feeds this straight to a
  // DocxSession op, which rejects a span that overshoots the committed length.
  const span = trimmedSpan(block, { start: Math.min(start, end), length: Math.abs(end - start) });
  return span.length > 0 ? span : null;
}

/** Restore a content-text selection spanning [start, start+length) within `el` (skips markers). */
function selectRange(el: HTMLElement, start: number, length: number): void {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  // The block may have been swapped out of the document by a re-render before this runs
  // (e.g. a focus-stealing toolbar control firing twice). addRange on a detached range
  // throws "the given range isn't in document" — skip rather than warn.
  if (!sel || !el.isConnected) return;
  el.focus();
  const range = document.createRange();
  const end = start + length;
  let pos = 0;
  let startSet = false;
  const walk = (node: Node): boolean => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3 /* TEXT_NODE */) {
        if (isInMarker(child)) continue; // marker text isn't part of the content offset space
        const raw = child.textContent ?? "";
        const len = stripBidi(raw).length; // content length excludes injected bidi marks
        if (!startSet && pos + len >= start) {
          range.setStart(child, domOffsetForContentOffset(raw, start - pos));
          startSet = true;
        }
        if (startSet && pos + len >= end) {
          range.setEnd(child, domOffsetForContentOffset(raw, end - pos));
          return true;
        }
        pos += len;
      } else if (walk(child)) {
        return true;
      }
    }
    return false;
  };
  if (walk(el) || startSet) {
    sel.removeAllRanges();
    sel.addRange(range);
  }
}

/**
 * True when `el`'s immediate parent is a paragraph-border `<div>` the full render wrapped it in
 * (CreateBorderDivs groups visibly-bordered paragraphs into a div). The body wrapper div has no
 * border. Splitting/merging such a block must re-render the whole document so the converter can
 * re-group the border boxes — an in-place node swap would leave the new (often borderless)
 * paragraph stranded inside the stale border div, drawing the rule's line under its text.
 */
function inBorderWrapper(el: HTMLElement): boolean {
  const style = el.parentElement?.getAttribute("style") ?? "";
  return /border-(top|bottom|left|right):\s*(?!none)[^;]+/i.test(style);
}

/** Whether the current selection's start already carries `key`, read from computed style. */
function selectionHasFormat(key: FormatKey, fallback: HTMLElement): boolean {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  let el: HTMLElement | null = fallback;
  if (sel && sel.rangeCount > 0) {
    const n = sel.getRangeAt(0).startContainer;
    el = n.nodeType === 3 ? n.parentElement : (n as HTMLElement);
  }
  if (!el || typeof getComputedStyle !== "function") return false;
  const cs = getComputedStyle(el);
  switch (key) {
    case "bold": return fontWeightIsBold(cs.fontWeight);
    case "italic": return cs.fontStyle === "italic" || cs.fontStyle === "oblique";
    case "underline": return cs.textDecorationLine.includes("underline");
    case "strike": return cs.textDecorationLine.includes("line-through");
    case "code": return /mono|courier|consolas/i.test(cs.fontFamily);
    case "superscript": return cs.verticalAlign === "super" || !!el.closest("sup");
    case "subscript": return cs.verticalAlign === "sub" || !!el.closest("sub");
    default: return false;
  }
}

/** Build the full ConvertDocxToHtmlComplete arg list (stampAnchors = last arg). */
function completeArgs(
  bytes: Uint8Array,
  cssPrefix: string,
  fabricate: boolean,
  paginated: boolean,
  scale: number,
): any[] {
  return [
    bytes, "Document", cssPrefix, fabricate, "", -1, "comment-",
    /* paginationMode */ paginated ? 1 : 0, /* paginationScale */ scale, "page-",
    false, 0, "annot-",
    // Footnotes/endnotes ON: they are document content, and the editor makes the rendered note
    // paragraphs editable. Must stay in step with DocxSessionOps.RenderHtml (the remount path),
    // whose output has to match this first paint byte-for-byte.
    /* renderFootnotesAndEndnotes */ true, /* renderHeadersAndFooters */ paginated,
    false, true, true, false, null, /* stampAnchors */ true,
  ];
}

export class DocxEditor {
  private readonly exports: DocxEditorExports;
  private readonly container: HTMLElement;
  private readonly handle: number;
  private readonly options: Required<Omit<DocxEditorOptions, "onEdit">> & Pick<DocxEditorOptions, "onEdit">;
  /** Map a block's current bare unid → its full kind:scope:unid (DocxSession anchor). */
  private readonly unidToFullId = new Map<string, string>();
  /** The element whose [data-anchor] descendants are the editable blocks (container or page container). */
  private editRoot: HTMLElement;
  /** The most recently focused editable block — the target for ribbon/format commands. */
  private activeBlock: HTMLElement | null = null;
  private closed = false;
  /**
   * Re-entrancy guard for node replacement. Replacing a contenteditable block that still holds
   * focus removes the focused node, which fires a SYNCHRONOUS `blur` → re-enters commitBlock; the
   * interleaved second replaceWith then throws NotFoundError ("node ... no longer a child") and the
   * structural edit (split/merge/format) is lost. While this flag is set, commitBlock no-ops.
   */
  private replacing = false;

  /**
   * The last real (non-collapsed) text selection inside an editable block. A toolbar control that
   * must take focus to be used — the font-size combobox — blurs the block and collapses the live
   * selection, so without this an operation triggered from such a control could only target the
   * whole paragraph (S-1 smoke-test finding 3). Refreshed whenever a non-empty selection sits in a
   * block, and cleared when a caret is collapsed inside a block (so it never goes stale).
   */
  private lastSelection: { unid: string; span: { start: number; length: number } } | null = null;

  /** The docked header/footer bands, when `options.headerFooter` is on. */
  private region: HeaderFooterRegion | null = null;

  private constructor(
    container: HTMLElement,
    exports: DocxEditorExports,
    handle: number,
    options: DocxEditor["options"],
  ) {
    this.container = container;
    this.exports = exports;
    this.handle = handle;
    this.options = options;
    this.editRoot = container;
    if (typeof document !== "undefined")
      document.addEventListener("selectionchange", this.onSelectionChange);
  }

  /** Track the last meaningful selection so focus-stealing toolbar controls can still target it. */
  private readonly onSelectionChange = (): void => {
    if (this.closed) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0) return; // no selection info — keep the cache as-is
    const range = sel.getRangeAt(0);
    const block = this.editableBlockOf(range.commonAncestorContainer);
    if (!block) return; // selection is outside the editor (e.g. a toolbar field) — keep the cache
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    if (range.collapsed) {
      this.lastSelection = null; // an explicit caret in a block — drop any stale selection
      return;
    }
    const span = selectionSpanIn(block);
    if (span) this.lastSelection = { unid, span };
  };

  /** The editable block (contenteditable [data-anchor]) containing `node`, if any, within this editor.
   *  Fenced by `container`, not `editRoot`, so header/footer band blocks — which live outside the
   *  body edit root by design — also register. The fence still rejects other editors on the page. */
  private editableBlockOf(node: Node | null): HTMLElement | null {
    if (!node) return null;
    const start = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement;
    const block = start?.closest<HTMLElement>('[data-anchor][contenteditable="true"]') ?? null;
    return block && this.container.contains(block) ? block : null;
  }

  /**
   * The root owning `el`'s sibling block list: its header/footer band's story container, else the
   * body edit root. Keeps a multi-block selection from spanning a band and the body, whose block
   * lists belong to different OOXML parts.
   */
  private ownerRoot(el: HTMLElement): HTMLElement {
    return this.region?.blockRootOf(el) ?? this.editRoot;
  }

  /** True when `el` is a header/footer band block rather than a body block. */
  private isBandBlock(el: HTMLElement): boolean {
    return !!this.region?.contains(el);
  }

  /**
   * Repaint after an edit to `block` that would otherwise remount the whole document: a band
   * repaints only itself (a story is one to three paragraphs), leaving the body DOM — and the
   * user's place in it — untouched.
   */
  private refreshAfter(block: HTMLElement, focusIndex: number, caretAtEnd = false): void {
    const band = this.region?.bandOf(block);
    if (band) {
      this.region!.refresh(this.region!.whichOf(band));
      return;
    }
    this.remount(focusIndex, caretAtEnd);
  }

  /** Open a document, render it into `container`, and wire up editing. */
  static open(
    container: HTMLElement,
    bytes: Uint8Array,
    exports: DocxEditorExports,
    options: DocxEditorOptions = {},
  ): DocxEditor {
    const opts = {
      cssPrefix: options.cssPrefix ?? "docx-",
      fabricateClasses: options.fabricateClasses ?? false,
      editable: options.editable ?? true,
      paginated: options.paginated ?? false,
      scale: options.scale ?? 1,
      headerFooter: options.headerFooter ?? false,
      onEdit: options.onEdit,
    };
    // persistAnchorIds=true keeps PtOpenXml:Unid attributes in Save() output, so a remount's
    // full re-render keeps the SAME unids the live session uses (a content change like becoming
    // a list otherwise re-derives a fresh unid, leaving the block unwired). The cost is that
    // saved bytes carry the Unid attributes (Word ignores them).
    const handle = exports.DocxSessionBridge.OpenSession(bytes, '{"persistAnchorIds":true}');
    const editor = new DocxEditor(container, exports, handle, opts);
    editor.refreshAnchorMap();
    if (opts.headerFooter) editor.createRegion();
    const fullHtml = exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(bytes, opts.cssPrefix, opts.fabricateClasses, opts.paginated, opts.scale),
    );
    if (opts.paginated) editor.mountPaginated(fullHtml);
    else editor.mountHtml(fullHtml);
    editor.syncRegionToBody();
    return editor;
  }

  /**
   * Open a fresh, blank document (a "New document" — single empty paragraph, Normal style,
   * US-Letter section) and wire up editing. The seed bytes come from the WASM bridge so the
   * result opens cleanly in Word too.
   */
  static openBlank(
    container: HTMLElement,
    exports: DocxEditorExports,
    options: DocxEditorOptions = {},
  ): DocxEditor {
    return DocxEditor.open(container, exports.DocxSessionBridge.CreateBlankDocx(), exports, options);
  }

  /** Lossless DOCX bytes reflecting all edits. */
  save(): Uint8Array {
    this.assertOpen();
    return this.exports.DocxSessionBridge.Save(this.handle);
  }

  /** Release the underlying WASM session. The editor is unusable afterward. */
  close(): void {
    if (this.closed) return;
    this.closed = true;
    if (typeof document !== "undefined")
      document.removeEventListener("selectionchange", this.onSelectionChange);
    this.exports.DocxSessionBridge.CloseSession(this.handle);
  }

  /**
   * Switch between continuous and paginated rendering WITHOUT losing edits. Re-renders from the
   * LIVE session, so every committed edit (and the undo/redo history) survives the toggle — unlike
   * re-opening the original bytes, which silently discards session edits. No-op if already `value`.
   */
  setPaginated(value: boolean): void {
    this.assertOpen();
    if (this.options.paginated === value) return;
    this.options.paginated = value;
    this.remount();
  }

  /** The editor's current DOM (for inspection/tests). */
  get root(): HTMLElement {
    return this.container;
  }

  // ─── internals ───────────────────────────────────────────────────────

  private assertOpen(): void {
    if (this.closed) throw new Error("DocxEditor is closed");
  }

  /**
   * Rebuild unid → full-anchor-id from the live session projection.
   *
   * Unids are CONTENT-ADDRESSED, so blocks with identical content in DIFFERENT parts collide —
   * e.g. a document with empty default/first/even header stories has one unid for several
   * header parts. A collision must resolve to the BODY entry: body blocks carry only
   * `data-anchor` (the bare unid) and have nothing else to resolve through, whereas a
   * header/footer band block carries its full anchor in `data-hf-anchor` and is resolved from
   * that (see `anchorIdOf`). Letting a non-body scope win here would silently redirect a body
   * edit into a header part.
   */
  private refreshAnchorMap(): void {
    const proj = JSON.parse(this.exports.DocxSessionBridge.Project(this.handle)) as {
      anchorIndex: Record<string, AnchorTargetLite>;
    };
    this.unidToFullId.clear();
    const bodyOwned = new Set<string>();
    for (const [fullId, target] of Object.entries(proj.anchorIndex)) {
      const isBody = target.scope === "body";
      if (!isBody && bodyOwned.has(target.unid)) continue;
      this.unidToFullId.set(target.unid, fullId);
      if (isBody) bodyOwned.add(target.unid);
    }
  }

  /**
   * The full `kind:scope:unid` anchor for a rendered block. A header/footer band block carries
   * its own — the unid map cannot disambiguate one, since several parts' story paragraphs can
   * share a content-addressed unid (a real Word document with empty default/first/even stories
   * does exactly that, and a unid-keyed lookup would land the edit in the wrong header part).
   */
  private anchorIdOf(el: HTMLElement): string | undefined {
    const stamped = el.getAttribute("data-hf-anchor");
    if (stamped) return stamped;
    const unid = el.getAttribute("data-anchor");
    return unid ? this.unidToFullId.get(unid) : undefined;
  }

  /** Build the header/footer region (called once, before the first mount). */
  private createRegion(): void {
    this.region = new HeaderFooterRegion(
      this.exports.DocxSessionBridge,
      this.handle,
      { cssPrefix: this.options.cssPrefix, fabricateClasses: this.options.fabricateClasses },
      {
        wireBlock: (el) => this.wireBlock(el),
        refreshAnchorMap: () => this.refreshAnchorMap(),
      },
    );
  }

  /**
   * Point the bands at the section governing the current focus (or the first body block).
   * Called after every mount, and on focus of a body block, so a multi-section document shows
   * the stories that actually apply where the caret is.
   */
  private syncRegionToBody(fromBlock?: HTMLElement): void {
    if (!this.region) return;
    const block = fromBlock ?? this.editableList()[0];
    const unid = block?.getAttribute("data-anchor");
    const id = unid ? this.unidToFullId.get(unid) : undefined;
    this.region.syncToBody(id ?? null);
  }

  /**
   * Insert the bands around `bodyRoot` and make `bodyRoot` the edit root. Band blocks must stay
   * OUT of the edit root: `editableList()`/`blockIndex()` enumerate it to compute remount focus
   * indices, and band blocks in that list would shift every index.
   */
  private dockBands(bodyRoot: HTMLElement): void {
    if (!this.region) return;
    bodyRoot.before(this.region.headerBand);
    bodyRoot.after(this.region.footerBand);
    this.region.refreshAll();
  }

  /** Continuous (non-paginated) mount: inject the converter's styles + body, wire blocks. */
  private mountHtml(fullHtml: string): void {
    const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    if (!this.region) {
      this.container.innerHTML = styles + parsed.body.innerHTML;
      this.editRoot = this.container;
      if (this.options.editable) this.wireBlocks(this.container);
      return;
    }
    // With bands docked, the body flow needs its own wrapper to be the edit root.
    this.container.innerHTML = styles;
    const flow = document.createElement("div");
    flow.className = "docx-body-flow";
    flow.innerHTML = parsed.body.innerHTML;
    this.container.appendChild(flow);
    this.editRoot = flow;
    if (this.options.editable) this.wireBlocks(flow);
    this.dockBands(flow);
  }

  /** Paginated mount: flow blocks into page boxes via pagination.ts, wire the page clones. */
  private mountPaginated(fullHtml: string): void {
    // With bands docked, pagination writes into its own wrapper so the bands can sit outside the
    // page stack (and so pagination's innerHTML reset can never eat them).
    let target = this.container;
    if (this.region) {
      this.container.innerHTML = "";
      target = document.createElement("div");
      target.className = "docx-body-flow";
      this.container.appendChild(target);
    }
    // Fragmented paragraphs intentionally have only one addressable head and
    // are therefore unsuitable for the editor's one-block editing model.
    paginateHtml(fullHtml, target, {
      scale: this.options.scale,
      cssPrefix: "page-",
      fragmentParagraphs: false,
    });
    // pagination.ts measures the hidden #pagination-staging subtree ONCE, then flows CLONES of its
    // blocks into the visible page boxes. Leaving staging in the live DOM is a trap: every
    // data-anchor exists twice (staging + page-box copy), so document.querySelector('[data-anchor]')
    // is ambiguous (hits the hidden copy first), and the staging copy goes stale because edits land
    // only on the page-box copy — a future reflow-from-staging would silently revert them. Staging
    // is a transient measurement scaffold; drop it so the page-box copies are the single source of
    // truth. A remount (setPaginated, list/undo edits) rebuilds staging fresh from the live session.
    this.container.querySelector("#pagination-staging, .page-staging")?.remove();
    const pageRoot = target.querySelector<HTMLElement>("#pagination-container") ?? target;
    this.editRoot = pageRoot;
    if (this.options.editable) this.wireBlocks(pageRoot);
    // The page boxes render their own (read-only) header/footer margins; the editable bands dock
    // around the page stack, so there is still exactly one addressable node per story paragraph.
    if (this.region) this.dockBands(target);
  }

  private wireBlocks(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((el) => this.wireBlock(el));
  }

  private wireBlock(el: HTMLElement): void {
    if (!EDITABLE_TAGS.has(el.tagName)) return;
    const unid = el.getAttribute("data-anchor");
    // Only blocks the markdown projection addresses are editable via the text path. This INCLUDES
    // table-cell paragraphs (the projection indexes them), so cell text IS editable — but structural
    // keys are kept inert inside a cell (see onKeydown / GAP3) so single-block editing can't corrupt
    // table structure. Anything the projection does not index (unstamped content) stays read-only.
    // A band block is authoritative via its stamped `data-hf-anchor` even when the unid map
    // resolves that unid to a different part (content-addressed unids collide across parts).
    if (!unid || !this.anchorIdOf(el)) return;
    el.setAttribute("contenteditable", "true");
    // Generated chrome (list number/bullet, footnote/endnote citation markers, note backrefs) is
    // not editable content — keep the caret out so offsets stay aligned with the run text, and so
    // a citation marker can't be deleted directly (which would orphan its note definition).
    el.querySelectorAll<HTMLElement>(GENERATED_CHROME_SELECTOR)
      .forEach((m) => m.setAttribute("contenteditable", "false"));
    // Baseline for the commit diff: CONTENT text (list markers + injected bidi marks excluded),
    // matching the session's flat run-text offset space.
    el.dataset.committedText = blockContentText(el);
    el.addEventListener("focus", () => {
      this.activeBlock = el;
      // Follow the caret's section so a cover-page-plus-body document shows the stories that
      // actually apply. Focusing a BAND block must not re-sync: it has no governing section of
      // its own, and re-resolving would clobber the user's kind selection.
      if (this.region && !this.isBandBlock(el)) this.syncRegionToBody(el);
    });
    el.addEventListener("blur", () => this.commitBlock(el));
    el.addEventListener("keydown", (ev) => this.onKeydown(el, ev as KeyboardEvent));
    // A band block re-rendered by an incremental swap is a fresh DOM node; re-adopt it (with the
    // anchor its caller already stamped) so the band chrome can still address it.
    const stamped = el.getAttribute("data-hf-anchor");
    if (stamped && this.region?.contains(el)) this.region.adoptBlock(el, stamped);
  }

  /**
   * Replace `oldEl` with `newNodes`, suppressing the re-entrant blur→commit that removing a focused
   * block fires (see `replacing`), AND tolerating the case where a synchronous blur during focus
   * transfer detaches `oldEl` between the caller's checks and here. `replaceWith` then throws
   * NotFoundError ("node … no longer a child … moved in a blur event handler") — `isConnected`
   * alone doesn't catch this race. The session is already updated, so a skipped/failed visual swap
   * leaves correct content (the typed DOM); the next commit or remount reconciles it. This is why
   * the catch is silent rather than rethrowing — there's no lost data, only a deferred re-render.
   * Returns true if the swap happened.
   */
  private replaceNode(oldEl: HTMLElement, ...newNodes: Node[]): boolean {
    const prev = this.replacing;
    this.replacing = true;
    try {
      if (!oldEl.parentNode) return false;
      oldEl.replaceWith(...newNodes);
      return true;
    } catch {
      return false;
    } finally {
      this.replacing = prev;
    }
  }

  /** Commit a block edit on blur: diff → run-preserving session op → re-render only this block. */
  private commitBlock(el: HTMLElement): void {
    if (this.closed || this.replacing) return;
    const unid = el.getAttribute("data-anchor");
    if (!unid) return;
    const fullId = this.anchorIdOf(el);
    if (!fullId) return;

    const result = this.commitTextChange(el, fullId);
    if (!result) return; // no change
    if (!result.success) {
      // Session unchanged — re-render this block from truth to discard the rejected DOM edit.
      const fresh = this.renderInto(fullId);
      if (fresh && this.replaceNode(el, fresh)) {
        this.wireBlock(fresh);
        if (this.activeBlock === el) this.activeBlock = fresh;
      }
      return;
    }

    const newAnchor = result.modified?.[0]?.id ?? fullId;
    const newUnid = result.modified?.[0]?.unid ?? unid;

    // List items: do NOT re-render on a text commit. Re-rendering replaces the node *during* the
    // blur, cancelling the browser's in-flight focus transfer when the user clicks straight to
    // another bullet; numbering also needs whole-document context a single-block render lacks. The
    // DOM already shows what the user typed with the correct marker — sync the baseline only.
    if (el.querySelector(":scope > [data-list-marker]")) {
      el.dataset.committedText = blockContentText(el);
      this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
      return;
    }

    // Plain block: re-render ONLY this block from the live session for canonical HTML. Swapping the
    // just-blurred node here is safe (verified — focus stays on the newly-clicked block).
    const html = this.exports.DocxSessionBridge.RenderBlockHtml(
      this.handle,
      newAnchor,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) !== 0x7b /* not an error object */) {
      const fresh = new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
      const inBand = this.isBandBlock(el);
      if (fresh && this.replaceNode(el, fresh)) {
        // Band blocks resolve through `data-hf-anchor`; their unid can collide with another
        // part's, so writing it into the map would corrupt that entry (see anchorIdOf).
        if (inBand) {
          this.region!.adoptBlock(fresh, newAnchor);
        } else {
          this.unidToFullId.delete(unid);
          this.unidToFullId.set(newUnid, newAnchor);
        }
        this.wireBlock(fresh);
        if (this.activeBlock === el) this.activeBlock = fresh; // keep ribbon target valid
      }
    }

    this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
  }

  // ─── M2: structural editing ──────────────────────────────────────────

  private onKeydown(el: HTMLElement, ev: KeyboardEvent): void {
    if (this.closed) return;
    // Common formatting / history shortcuts.
    if ((ev.ctrlKey || ev.metaKey) && !ev.altKey) {
      const k = ev.key.toLowerCase();
      const fmt: Record<string, FormatKey> = { b: "bold", i: "italic", u: "underline" };
      if (fmt[k]) { ev.preventDefault(); this.format(fmt[k]); return; }
      if (k === "z") { ev.preventDefault(); ev.shiftKey ? this.redo() : this.undo(); return; }
      if (k === "y") { ev.preventDefault(); this.redo(); return; }
    }
    // Inside a table cell, structural ops that change the TABLE GRID (cross-cell merge,
    // list-nest, focus-jumping Tab) stay INERT — the single-block model can't give them
    // whole-table context. Tab is swallowed (no focus escape / literal tab); Backspace at
    // the cell's start does not merge across the cell boundary (mid-text Backspace still
    // deletes normally). Enter, however, splits the cell paragraph into two paragraphs
    // WITHIN the same cell — the engine keeps the new w:p in the w:tc, the grid is
    // unchanged, so it's safe (it's how a cell holds stacked lines: value over a smaller
    // label, multi-line addresses). (GAP3.)
    const inTableCell = !!el.closest("table");

    // Tab / Shift+Tab on a list item nests / un-nests it (changes list level).
    if (ev.key === "Tab") {
      if (inTableCell) { ev.preventDefault(); return; }
      if (isListBlock(el)) {
        ev.preventDefault();
        this.activeBlock = el;
        this.setListLevel(ev.shiftKey ? -1 : 1);
        return;
      }
    }
    // Shift+Enter inserts an intra-paragraph line break (a real w:br on commit),
    // not a paragraph split. Deterministic across browsers and allowed in cells
    // (a line break changes no table structure).
    if (ev.key === "Enter" && ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      this.insertLineBreakAtCaret();
      return;
    }
    if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      // Splits at the caret — in a cell this stacks a second paragraph within the same
      // w:tc (grid unchanged); in the body it splits the paragraph as before.
      this.splitAtCaret(el);
    } else if (ev.key === "Backspace") {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (sel && sel.isCollapsed && caretOffsetIn(el) === 0 && !inTableCell) {
        const prev = this.previousEditable(el);
        if (prev) {
          ev.preventDefault();
          this.mergeWithPrevious(prev, el);
        }
      }
    }
  }

  /** Shift+Enter: insert an intra-paragraph line break at the caret. Delegates to the
   *  native `insertLineBreak` command, which inserts a <br> AND positions the caret
   *  after it correctly (handling the browser's bogus trailing-<br> rule) so typing
   *  continues on the new line. Commits (on blur) as a w:br via the "  \n" hard break
   *  the serializer emits for a <br>. */
  private insertLineBreakAtCaret(): void {
    if (typeof document !== "undefined" && typeof document.execCommand === "function") {
      document.execCommand("insertLineBreak");
    }
  }

  /** Enter: split the block at the caret into two paragraphs. */
  private splitAtCaret(el: HTMLElement): void {
    const rawOffset = caretOffsetIn(el);
    const unid = el.getAttribute("data-anchor");
    if (rawOffset == null || !unid) return;
    let fullId = this.anchorIdOf(el);
    if (!fullId) return;
    // The session commits trimmed text, so map the DOM caret offset into the trimmed run-text
    // offset (else an overshoot — e.g. a placeholder leading space — is rejected and Enter is lost).
    const offset = trimmedSplitOffset(el, rawOffset);

    const idx = this.blockIndex(el); // capture before the op (for list remount focus)
    // Whether this paragraph is rendered inside a border <div> — captured before the DOM mutates.
    const wrappedInBorder = inBorderWrapper(el);
    fullId = this.syncBlock(el, fullId); // flush any uncommitted typing first
    const res = this.parseEdit(this.exports.DocxSessionBridge.SplitParagraph(this.handle, fullId, offset));
    if (!res.success) return;
    const first = res.modified?.[0];
    const second = res.created?.[0];
    if (!first || !second) return;

    // Splitting a list item makes a continuing list item, and splitting a bordered paragraph (e.g.
    // a horizontal rule) yields a new paragraph whose border status differs — both need a whole-
    // document re-render so numbering continues / border <div>s regroup correctly. An in-place node
    // swap would leave the new paragraph inside the old border div (the rule's line under its text).
    if (this.affectsList(res) || wrappedInBorder) {
      this.refreshAfter(el, idx + 1, false);
      this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
      return;
    }

    const firstEl = this.renderInto(first.id);
    const secondEl = this.renderInto(second.id);
    if (!firstEl || !secondEl) return;

    // el is the focused block — replaceNode guards the re-entrant blur→commit and tolerates a
    // node detached mid-focus-transfer; replacing with both new blocks at once keeps them adjacent.
    const inBand = this.isBandBlock(el);
    if (!this.replaceNode(el, firstEl, secondEl)) return;
    // Band blocks resolve through `data-hf-anchor` (their unids can collide across parts).
    if (inBand) {
      this.region!.adoptBlock(firstEl, first.id);
      this.region!.adoptBlock(secondEl, second.id);
    } else {
      this.unidToFullId.delete(unid);
      this.unidToFullId.set(first.unid, first.id);
      this.unidToFullId.set(second.unid, second.id);
    }
    this.wireBlock(firstEl);
    this.wireBlock(secondEl);
    placeCaretAtOffset(secondEl, 0);
    this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
  }

  /** Backspace at block start: merge this block into the previous one. */
  private mergeWithPrevious(prev: HTMLElement, el: HTMLElement): void {
    const prevUnid = prev.getAttribute("data-anchor");
    const thisUnid = el.getAttribute("data-anchor");
    if (!prevUnid || !thisUnid) return;
    let prevId = this.anchorIdOf(prev);
    let thisId = this.anchorIdOf(el);
    if (!prevId || !thisId) return;

    const prevIdx = this.blockIndex(prev); // capture before the op
    // Either side rendered inside a border <div> means the merge changes border grouping — captured
    // before the DOM mutates so the post-merge branch can force a full re-render.
    const wrappedInBorder = inBorderWrapper(prev) || inBorderWrapper(el);
    prevId = this.syncBlock(prev, prevId);
    thisId = this.syncBlock(el, thisId);
    const caret = (prev.textContent ?? "").length; // merge boundary

    const res = this.parseEdit(this.exports.DocxSessionBridge.MergeParagraphs(this.handle, prevId, thisId));
    if (!res.success) return;
    const merged = res.modified?.[0];
    if (!merged) return;

    // Merging list items renumbers the list, and merging across a border <div> boundary changes the
    // border grouping — both need a whole-document re-render (caret at the merge boundary).
    if (this.affectsList(res) || wrappedInBorder) {
      this.refreshAfter(el, prevIdx, true);
      this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
      return;
    }

    const mergedEl = this.renderInto(merged.id);
    if (!mergedEl) return;

    // prev may be focused — replaceNode guards re-entrancy and tolerates a detached node.
    const inBand = this.isBandBlock(prev);
    if (!this.replaceNode(prev, mergedEl)) return;
    el.remove();
    // Band blocks resolve through `data-hf-anchor` (their unids can collide across parts).
    if (inBand) {
      this.region!.adoptBlock(mergedEl, merged.id);
    } else {
      this.unidToFullId.delete(prevUnid);
      this.unidToFullId.delete(thisUnid);
      this.unidToFullId.set(merged.unid, merged.id);
    }
    this.wireBlock(mergedEl);
    placeCaretAtOffset(mergedEl, caret);
    this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
  }

  /**
   * Apply the block's pending text change to the session with full inline-formatting fidelity.
   * Diffs the committed content text (markers + bidi excluded) against the current content text
   * and rewrites only the changed span via ReplaceTextAtSpan — every untouched run keeps its exact
   * rPr, and typed text inherits the boundary run's formatting. Returns the parsed EditResult, or
   * null when there is no change. Empty/whitespace-only baselines (e.g. the placeholder space the
   * converter renders for an empty paragraph, whose DOM text doesn't line up with the session's
   * empty run text) are rebuilt via ReplaceText — there is no inline formatting to preserve there.
   */
  private commitTextChange(el: HTMLElement, fullId: string): EditResultLite | null {
    // `old` mirrors the session's flat run-text: strip bidi marks (blockContentText strips them,
    // but wireBlock may have stored textContent before this Task 2 change, and the bidi test
    // explicitly stores textContent). Using stripBidi keeps the baseline consistent with the
    // session's offset space regardless of how committedText was stored.
    const old = stripBidi(el.dataset.committedText ?? "");
    const next = blockContentText(el);
    if (old === next) return null;

    if (old.trim().length === 0) {
      return this.parseEdit(
        this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, serializeInlineMarkdown(el)),
      );
    }

    const minLen = Math.min(old.length, next.length);
    let p = 0;
    while (p < minLen && old[p] === next[p]) p++;
    let s = 0;
    while (s < minLen - p && old[old.length - 1 - s] === next[next.length - 1 - s]) s++;

    let start = p;
    let len = old.length - p - s;
    let middle = next.slice(p, next.length - s);

    // A pure insertion is a zero-length span, which resolves to no runs and is rejected. Anchor a
    // neighbor char so the span is non-empty and the inserted text inherits an adjacent run's rPr
    // (the LEFT run when there is one, matching contenteditable; the first run at the very start).
    if (len === 0) {
      if (start > 0) { start -= 1; len = 1; middle = old[start] + middle; }
      else { len = 1; middle = middle + old[0]; }
    }

    return this.parseEdit(
      this.exports.DocxSessionBridge.ReplaceTextAtSpan(this.handle, fullId, start, len, middle),
    );
  }

  /** Flush a block's current (uncommitted) text to the session; returns the live full id. */
  private syncBlock(el: HTMLElement, fullId: string): string {
    const result = this.commitTextChange(el, fullId);
    if (!result || !result.success) return fullId;
    el.dataset.committedText = blockContentText(el);
    return result.modified?.[0]?.id ?? fullId;
  }

  /** Render a block by anchor and parse it into a detached element (null on error). */
  private renderInto(anchorId: string): HTMLElement | null {
    const html = this.exports.DocxSessionBridge.RenderBlockHtml(
      this.handle,
      anchorId,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) === 0x7b /* error object */) return null;
    return new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
  }

  /** The editable block immediately before `el` within its own root, or null. */
  private previousEditable(el: HTMLElement): HTMLElement | null {
    const all = Array.from(
      this.ownerRoot(el).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    const i = all.indexOf(el);
    return i > 0 ? all[i - 1] : null;
  }

  private parseEdit(json: string): EditResultLite {
    try {
      return JSON.parse(json) as EditResultLite;
    } catch {
      return { success: false };
    }
  }

  // ─── M5: formatting commands (ribbon) ────────────────────────────────

  // ─── Multi-block selection helpers (format a whole stack of paragraphs at once) ──────

  /** Editable blocks the current selection covers, in document order. Uses Range.comparePoint
   *  (robust to a selection boundary that normalized onto a wrapper element rather than a block
   *  or text node — Range.intersectsNode misses the end block at a `(block, childCount)` boundary).
   *  A collapsed or single-block selection yields just the active block. */
  private selectedBlocks(): HTMLElement[] {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    // Enumerate the ACTIVE block's own root — a band's story container, else the body edit root —
    // so a selection can never span a band and the body (different OOXML parts).
    const root = this.activeBlock ? this.ownerRoot(this.activeBlock) : this.editRoot;
    const all = Array.from(
      root.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
      const range = sel.getRangeAt(0);
      const hit = all.filter((b) => {
        try {
          const startsAfterEnd = range.comparePoint(b, 0) > 0; // block begins after the selection ends
          const endsBeforeStart = range.comparePoint(b, b.childNodes.length) < 0; // block ends before it starts
          return !startsAfterEnd && !endsBeforeStart;
        } catch {
          return false;
        }
      });
      if (hit.length > 1) return hit;
    }
    return this.activeBlock ? [this.activeBlock] : [];
  }

  /** The selection's span within `block`, clipped to the block (for inline ops across blocks):
   *  the first block runs selection-start→end-of-block, middle blocks are whole, the last block
   *  runs start-of-block→selection-end. Returns null for a whole-block apply. */
  private blockSpanForSelection(block: HTMLElement): { start: number; length: number } | null {
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
    const range = sel.getRangeAt(0);
    const hasStart = block.contains(range.startContainer);
    const hasEnd = block.contains(range.endContainer);
    if (hasStart && hasEnd) return selectionSpanIn(block);
    // Partial/whole slices are normalized into the committed run-text space for the same reason
    // selectionSpanIn is (see trimmedSpan): a session op rejects a span past the committed end.
    const contentLen = blockContentText(block).length;
    if (hasStart) {
      const start = contentOffsetOf(block, range.startContainer, range.startOffset);
      return trimmedSpan(block, { start, length: Math.max(0, contentLen - start) });
    }
    if (hasEnd) {
      const end = contentOffsetOf(block, range.endContainer, range.endOffset);
      return trimmedSpan(block, { start: 0, length: end });
    }
    return trimmedSpan(block, { start: 0, length: contentLen }); // fully-spanned middle block
  }

  /** Apply an inline ApplyFormat op to each block's slice of the selection, then reconcile
   *  the DOM incrementally (see {@link finishMultiBlockOp} — a full remount costs a whole-
   *  document convert, seconds on a large doc, where per-block swaps are ~10 ms each).
   *  Returns false (caller falls back to the single-block path) for a 1-block selection. */
  private applyInlineOpAcrossBlocks(blocks: HTMLElement[], op: object): boolean {
    if (blocks.length <= 1) return false;
    const targets = this.multiBlockTargets(blocks);
    for (const t of targets) {
      if (!t.fullId || (t.span && t.span.length === 0)) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      const res = this.parseEdit(this.exports.DocxSessionBridge.ApplyFormat(
        this.handle, synced, t.span ? JSON.stringify(t.span) : "", JSON.stringify(op),
      ));
      if (res.success) t.res = res;
    }
    this.finishMultiBlockOp(targets, false);
    return true;
  }

  /** Apply a whole-block (paragraph-level) op to each selected block, then reconcile the DOM
   *  incrementally ({@link finishMultiBlockOp}). `forceRemount` is for ops whose rendering
   *  needs whole-document context (e.g. border-div regrouping after clearBorders). Returns
   *  false for a 1-block selection (caller uses the single-block path). */
  private applyParagraphOpAcrossBlocks(
    blocks: HTMLElement[],
    run: (fullId: string) => string,
    forceRemount = false,
  ): boolean {
    if (blocks.length <= 1) return false;
    const targets = this.multiBlockTargets(blocks);
    for (const t of targets) {
      if (!t.fullId) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      const res = this.parseEdit(run(synced));
      if (res.success) t.res = res;
    }
    this.finishMultiBlockOp(targets, forceRemount);
    return true;
  }

  /** Snapshot each selected block's identity + selection slice BEFORE any session op runs
   *  (ops never mutate the DOM, so spans captured here stay valid until the swap phase). */
  private multiBlockTargets(blocks: HTMLElement[]): Array<{
    block: HTMLElement;
    unid: string | null;
    fullId: string | undefined;
    span: { start: number; length: number } | null;
    res: EditResultLite | null;
  }> {
    return blocks.map((b) => {
      const unid = b.getAttribute("data-anchor");
      return {
        block: b,
        unid,
        fullId: this.anchorIdOf(b),
        span: this.blockSpanForSelection(b),
        res: null,
      };
    });
  }

  /**
   * Reconcile the DOM after a multi-block op. Fidelity-identical to the single-block path by
   * construction: each edited block is swapped for its own session-attached single-block
   * render — exactly what format()/setFontSize()/applyParagraphFormat() do for one block —
   * so a multi-block apply is N single-block applies, not one whole-document re-render
   * (which froze the UI for the full-document convert time on every ribbon action). Falls
   * back to ONE full remount when the op touched a list item (numbering continuation needs
   * whole-document context) or the caller forced it. Restores the cross-block selection so
   * consecutive ribbon actions (center, then bold) keep targeting the same range.
   */
  private finishMultiBlockOp(
    targets: Array<{
      block: HTMLElement;
      unid: string | null;
      span: { start: number; length: number } | null;
      res: EditResultLite | null;
    }>,
    forceRemount: boolean,
  ): void {
    const edited = targets.filter((t) => t.res && t.unid);
    if (edited.length === 0) return;
    // Paginated mode: a multi-block format change can alter block heights, and page boxes
    // need a reflow — keep the (pre-existing) full remount there until M4 lands a scoped
    // re-paginate. Continuous mode reconciles incrementally.
    if (forceRemount || this.options.paginated || edited.some((t) => this.affectsList(t.res!))) {
      this.refreshAfter(edited[0].block, -1, false);
      return;
    }
    const swapped: Array<{ el: HTMLElement; span: { start: number; length: number } | null }> = [];
    for (const t of edited) {
      const fresh = this.swapBlock(t.block, t.unid!, t.res!.modified?.[0]);
      swapped.push({ el: fresh ?? t.block, span: t.span });
    }
    const first = swapped[0];
    const last = swapped[swapped.length - 1];
    if (!first.el.isConnected || !last.el.isConnected) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel) return;
    const startOff = first.span?.start ?? 0;
    const endOff = last.span ? last.span.start + last.span.length : blockContentText(last.el).length;
    const startPos = contentPositionIn(first.el, startOff);
    const endPos = contentPositionIn(last.el, endOff);
    try {
      const range = document.createRange();
      range.setStart(startPos.node, startPos.offset);
      range.setEnd(endPos.node, endPos.offset);
      sel.removeAllRanges();
      sel.addRange(range);
    } catch {
      /* selection restore is best-effort — the edits themselves are already committed */
    }
  }

  /**
   * Toggle (or set) an inline format on the current selection in the active block.
   * A selection spanning multiple blocks applies to each. With no selection, applies to
   * the whole paragraph. Routes through DocxSession (`ApplyFormat`) so it is lossless and
   * supports underline/strike, not just markdown.
   */
  format(key: FormatKey, value?: boolean): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1) {
      const on0 = value ?? !selectionHasFormat(key, blocks[0]);
      const multiOp =
        key === "superscript" || key === "subscript"
          ? { vertAlign: on0 ? key : "" }
          : { [key]: on0 };
      if (this.applyInlineOpAcrossBlocks(blocks, multiOp)) return;
    }
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;

    const span = selectionSpanIn(block);
    const on = value ?? !selectionHasFormat(key, block);
    // Super/subscript map to the single-valued w:vertAlign; the rest are boolean toggles.
    const op =
      key === "superscript" || key === "subscript"
        ? { vertAlign: on ? key : "" }
        : { [key]: on };
    fullId = this.syncBlock(block, fullId); // don't clobber uncommitted typing
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify(op),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /**
   * Set the font size (in points) of the current selection in the active block; with no
   * selection, applies to the whole paragraph. `pts <= 0` clears the explicit size. Routes
   * through DocxSession `ApplyFormat` (`w:sz`), so it is lossless and survives save.
   */
  setFontSize(pts: number): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1 && this.applyInlineOpAcrossBlocks(blocks, { fontSizePts: pts })) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    // Use the live selection; if the font-size combobox stole focus and collapsed it, fall back to
    // the last real selection cached for this block so a sub-range still sizes (finding 3).
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify({ fontSizePts: pts }),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /**
   * Set the font family of the current selection in the active block; with no selection,
   * applies to the whole paragraph. `""` clears the explicit font (inherits the style/default).
   * Routes through DocxSession `ApplyFormat` (`w:rFonts`), so it is lossless and survives save.
   * Multi-block + last-selection plumbing matches {@link setFontSize} (a focus-stealing font
   * dropdown still applies to the real sub-range).
   */
  setFontFamily(name: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (blocks.length > 1 && this.applyInlineOpAcrossBlocks(blocks, { fontFamily: name })) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    let span = selectionSpanIn(block);
    if (!span && this.lastSelection && this.lastSelection.unid === unid) span = this.lastSelection.span;
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle,
        fullId,
        span ? JSON.stringify(span) : "",
        JSON.stringify({ fontFamily: name }),
      ),
    );
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, this.blockIndex(block), false); return; }
    const fresh = this.swapBlock(block, unid, res.modified?.[0]);
    if (fresh && span) selectRange(fresh, span.start, span.length);
    else fresh?.focus();
  }

  /** Set paragraph alignment (left/center/right/justify) on the active block. */
  setAlignment(alignment: EditorAlignment): void {
    this.applyParagraphFormat({ alignment });
  }

  /**
   * Insert an S-1-style horizontal rule (an empty paragraph with a bottom border) after the
   * active block. `weight` is the rule thickness in eighths of a point (default 12 ≈ 1.5pt).
   * Re-renders fully (a new block needs whole-document context to lay out).
   */
  insertHorizontalRule(weight = 12, style = "single", position: "above" | "below" = "below"): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.InsertHorizontalRule(
        this.handle,
        fullId,
        position === "above" ? "before" : "after",
        JSON.stringify({ style, size: weight, color: "auto" }),
      ),
    );
    if (!res.success) return;
    // remount from the active block's index re-renders the new rule whether it landed just
    // above (at idx) or just below (at idx+1) the active block.
    this.refreshAfter(block, idx, false);
  }

  /**
   * Insert a `rows`×`cols` table after the active block. `options.cellContents` (row-major
   * markdown) seeds the cells, `options.borderless` makes an invisible layout table, and
   * `options.cellAlignment` aligns every cell. Re-renders fully (tables need document context).
   */
  insertTable(
    rows: number,
    cols: number,
    options?: {
      borderless?: boolean;
      cellContents?: string[];
      cellAlignment?: EditorAlignment;
      columnWidths?: number[];
    },
  ): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    // If the caret is on an empty paragraph (not a table cell), insert the table BEFORE it so the
    // empty paragraph becomes the editable line BELOW the table — no stray blank line above it, and
    // a reachable paragraph below (S-1 smoke-test findings 2 + 4). Otherwise insert after.
    const emptyHere =
      !block.closest("table") && blockContentText(block).replace(/[\s ]+/g, "").length === 0;
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.InsertTable(
        this.handle,
        fullId,
        emptyHere ? "before" : "after",
        rows,
        cols,
        options ? JSON.stringify(options) : "",
      ),
    );
    if (!res.success) return;
    this.refreshAfter(block, idx, false);
  }

  // ─── Table row / column editing (active block must be inside a table cell) ──────────

  /** Run a table-structure op on the active cell (a cell-paragraph block) and re-render. */
  private tableEdit(run: (cellAnchor: string) => string): void {
    const block = this.activeBlock;
    if (this.closed || !block || !block.closest("table")) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId); // flush uncommitted cell text first
    const res = this.parseEdit(run(fullId));
    if (!res.success) return;
    this.refreshAfter(block, idx, false);
  }

  /** Insert a row above/below the active cell's row. No-op outside a table. */
  insertTableRow(where: "above" | "below"): void {
    this.tableEdit((a) =>
      this.exports.DocxSessionBridge.InsertTableRow(this.handle, a, where === "above" ? "before" : "after"),
    );
  }

  /** Insert a column left/right of the active cell's column. No-op outside a table. */
  insertTableColumn(where: "left" | "right"): void {
    this.tableEdit((a) =>
      this.exports.DocxSessionBridge.InsertTableColumn(this.handle, a, where === "left" ? "before" : "after"),
    );
  }

  /** Delete the active cell's row (deleting the last row removes the table). No-op outside a table. */
  deleteTableRow(): void {
    this.tableEdit((a) => this.exports.DocxSessionBridge.DeleteTableRow(this.handle, a));
  }

  /** Delete the active cell's column (deleting the last column removes the table). No-op outside a table. */
  deleteTableColumn(): void {
    this.tableEdit((a) => this.exports.DocxSessionBridge.DeleteTableColumn(this.handle, a));
  }

  /**
   * Indent/outdent the active block. On a LIST item this changes the list NESTING LEVEL
   * (`SetListLevel`) so numbering nests (e.g. 1, 2 → a sub-level) rather than the item just
   * shifting sideways with flat numbering. On a plain paragraph it adjusts the left indent by
   * `deltaTwips` (default ±720 = 0.5"), clamped at 0.
   */
  indent(deltaTwips = 720): void {
    if (this.activeBlock && isListBlock(this.activeBlock)) {
      this.setListLevel(deltaTwips >= 0 ? 1 : -1);
      return;
    }
    this.applyParagraphFormat({ indentDelta: deltaTwips });
  }

  /** Change the active list item's nesting level by `delta` (+1 deeper, −1 shallower). */
  private setListLevel(delta: number): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetListLevel(this.handle, fullId, delta));
    if (!res.success) return;
    // A level change ripples through the whole list's numbering — re-render with full document
    // context (a single-block render can't compute nested numbering), keeping the caret in place.
    this.refreshAfter(block, idx, false);
  }

  /** Toggle (or set) page-break-before on the active block. */
  pageBreakBefore(value = true): void {
    this.applyParagraphFormat({ pageBreakBefore: value });
  }

  /**
   * Toggle the active block between a bullet/numbered list item and a plain paragraph.
   * Clicking the same kind it already is removes the list; any other state applies the kind.
   */
  toggleList(kind: "bullet" | "decimal"): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;

    let membership: { format?: string } | null = null;
    try {
      membership = JSON.parse(this.exports.DocxSessionBridge.GetListMembership(this.handle, fullId));
    } catch { /* treat as not-a-list */ }
    const isThisKind =
      !!membership && typeof membership.format === "string" &&
      membership.format.toLowerCase().startsWith(kind === "bullet" ? "bullet" : "decimal");

    const idx = this.blockIndex(block); // capture before the op
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ApplyListFormat(this.handle, fullId, isThisKind ? "none" : kind),
    );
    if (!res.success) return;
    // Numbering continuation across the list needs whole-document context — re-render fully
    // (a single-block render would show every numbered item as "1.").
    this.refreshAfter(block, idx, false);
  }

  /** Clear all paragraph borders (e.g. remove an inserted horizontal rule) on the active block —
   *  or every block in a multi-block selection. The engine/wire already accept `clearBorders`;
   *  this surfaces it on the editor so an HR border is removable (S-1 smoke-test finding 1b). */
  clearParagraphBorders(): void {
    this.applyParagraphFormat({ clearBorders: true });
  }

  /**
   * Delete the active block (e.g. a stray empty paragraph left above/below a table). Routes
   * through DocxSession `DeleteBlock` + re-render, focusing the previous block. No-op when the
   * caret is inside a table (remove cells via the table toolbar's delete row/column instead) and
   * no-op when it is the only editable block (don't empty the document). Closes the S-1
   * smoke-test "no block-delete affordance" gap.
   */
  deleteBlock(): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    if (block.closest("table")) return; // cells are removed via the table toolbar, not here
    if (this.editableList().length <= 1) return; // never delete the last editable block
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    const fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    const res = this.parseEdit(this.exports.DocxSessionBridge.DeleteBlock(this.handle, fullId));
    if (!res.success) return;
    this.refreshAfter(block, Math.max(0, idx - 1), true);
  }

  private applyParagraphFormat(op: {
    alignment?: EditorAlignment;
    indentDelta?: number;
    pageBreakBefore?: boolean;
    clearBorders?: boolean;
  }): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (
      blocks.length > 1 &&
      this.applyParagraphOpAcrossBlocks(
        blocks,
        (id) => this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, id, JSON.stringify(op)),
        // A border change regroups the wrapping border <div>s — whole-document context.
        !!op.clearBorders,
      )
    )
      return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, fullId, JSON.stringify(op)),
    );
    if (!res.success) return;
    // A border change adds/removes the wrapping border <div>, so a single-block swap can't restructure
    // it correctly — re-render fully (like list edits) so the wrapper appears/disappears cleanly.
    if (this.affectsList(res) || op.clearBorders) { this.refreshAfter(block, idx, false); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /** Set the paragraph style of the active block — or of every block in a multi-block selection
   *  (e.g. "Heading1", "Heading2", "Normal"). */
  setParagraphStyle(styleId: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    const blocks = this.selectedBlocks();
    if (
      blocks.length > 1 &&
      this.applyParagraphOpAcrossBlocks(blocks, (id) =>
        this.exports.DocxSessionBridge.SetParagraphStyle(this.handle, id, styleId),
      )
    )
      return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetParagraphStyle(this.handle, fullId, styleId));
    if (!res.success) return;
    if (this.affectsList(res)) { this.refreshAfter(block, idx, false); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /** Undo the last edit (re-renders the document). */
  undo(): void {
    if (this.closed) return;
    if (this.exports.DocxSessionBridge.Undo(this.handle)) this.remount();
  }

  /** Redo the last undone edit (re-renders the document). */
  redo(): void {
    if (this.closed) return;
    if (this.exports.DocxSessionBridge.Redo(this.handle)) this.remount();
  }

  // ─── Header/footer region commands (no-ops unless `headerFooter` is on) ───────────────

  /**
   * Select which story kind a band edits (`"default"` / `"first"` / `"even"`). A kind with no
   * existing part is created empty, so the band always presents something editable. `"first"`
   * sets the section's `w:titlePg`; `"even"` sets the document-global `w:evenAndOddHeaders`
   * (which also governs footers — the band surfaces that caveat inline).
   */
  setHeaderFooterKind(which: BandWhich, kind: HeaderFooterKind): void {
    this.assertOpen();
    this.region?.setKind(which, kind);
  }

  /** The story kind a band is currently editing, or null when the region is off. */
  headerFooterKind(which: BandWhich): HeaderFooterKind | null {
    return this.region?.kindOf(which) ?? null;
  }

  /**
   * Append a page-number field to the focused header/footer story paragraph (falling back to the
   * band's last paragraph — Word's convention). No-op outside a band.
   */
  insertPageNumber(field: "currentPage" | "totalPages" = "currentPage"): void {
    this.assertOpen();
    if (!this.region) return;
    const block = this.activeBlock;
    const band = block ? this.region.bandOf(block) : null;
    const anchorId = block?.getAttribute("data-hf-anchor");
    if (band && anchorId) {
      this.region.insertPageNumber(this.region.whichOf(band), anchorId, field);
      return;
    }
    // No band block focused — target the footer, where page numbers overwhelmingly live.
    this.region.insertPageNumberInBand("footer", field);
  }

  /** Which inline formats the current selection carries — for ribbon button highlighting. */
  queryFormatState(): Record<FormatKey, boolean> {
    const block = this.activeBlock ?? this.editRoot;
    return {
      bold: selectionHasFormat("bold", block),
      italic: selectionHasFormat("italic", block),
      underline: selectionHasFormat("underline", block),
      strike: selectionHasFormat("strike", block),
      code: selectionHasFormat("code", block),
      superscript: selectionHasFormat("superscript", block),
      subscript: selectionHasFormat("subscript", block),
    };
  }

  /** Re-render one block from the live session by EditResult ref, swapping it in place. */
  private swapBlock(oldEl: HTMLElement, oldUnid: string, ref?: AnchorRef): HTMLElement | null {
    const inBand = this.isBandBlock(oldEl);
    const anchorId = ref?.id ?? this.anchorIdOf(oldEl);
    const newUnid = ref?.unid ?? oldUnid;
    if (!anchorId) return null;
    const fresh = this.renderInto(anchorId);
    if (!fresh || !this.replaceNode(oldEl, fresh)) return null;
    // A band block resolves through its stamped `data-hf-anchor`, never the unid map — and its
    // unid can collide with another part's, so writing it here would corrupt that entry.
    if (!inBand) {
      this.unidToFullId.delete(oldUnid);
      this.unidToFullId.set(newUnid, anchorId);
    }
    // Stamp BEFORE wiring so wireBlock's own resolution sees the authoritative id.
    if (inBand) this.region!.adoptBlock(fresh, anchorId);
    this.wireBlock(fresh);
    this.activeBlock = fresh;
    this.options.onEdit?.({ anchorId, unid: newUnid });
    return fresh;
  }

  /**
   * Full-document HTML from the live session. Prefers the session-attached
   * `RenderHtml` bridge — the saved bytes never cross the JS/WASM boundary
   * (two multi-MB copies per remount on a large doc) — and falls back to
   * Save + ConvertDocxToHtmlComplete for older WASM bundles. Both paths use
   * the same option profile, so the rendered HTML is identical.
   */
  private renderFullHtml(): string {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderHtml === "function") {
      const html = bridge.RenderHtml(
        this.handle,
        this.options.cssPrefix,
        this.options.fabricateClasses,
        this.options.paginated,
        this.options.scale,
      );
      if (html.charCodeAt(0) !== 0x7b /* not an error object */) return html;
    }
    const bytes = bridge.Save(this.handle);
    return this.exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(bytes, this.options.cssPrefix, this.options.fabricateClasses, this.options.paginated, this.options.scale),
    );
  }

  /** Editable BODY blocks in document order (band blocks are enumerated by `ownerRoot`). */
  private editableList(): HTMLElement[] {
    return Array.from(this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
  }

  private blockIndex(el: HTMLElement): number {
    return Array.from(
      this.ownerRoot(el).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    ).indexOf(el);
  }

  /**
   * True when an edit produced or touched a list item (kind "li"). List markers and
   * numbering CONTINUATION need whole-document context, which a single-block render lacks
   * (every item would render as "1."), so such edits re-render the whole document.
   */
  private affectsList(res: EditResultLite): boolean {
    return [...(res.modified ?? []), ...(res.created ?? [])].some((r) => r.kind === "li");
  }

  /**
   * Full re-render from current session state (after undo/redo, and after list edits where
   * single-block rendering can't compute numbering). Optionally focus the editable block at
   * `focusIndex` (caret at start, or end if `caretAtEnd`) — addressed by index because a
   * block's content-hashed unid changes across the save/reproject a remount performs.
   */
  private remount(focusIndex = -1, caretAtEnd = false): void {
    this.refreshAnchorMap();
    const fullHtml = this.renderFullHtml();
    this.activeBlock = null;
    if (this.options.paginated) this.mountPaginated(fullHtml);
    else this.mountHtml(fullHtml);
    if (focusIndex >= 0) {
      const blocks = this.editableList();
      const target = blocks[Math.min(focusIndex, blocks.length - 1)];
      if (target) {
        this.activeBlock = target;
        placeCaretAtOffset(target, caretAtEnd ? (target.textContent ?? "").length : 0);
      }
    }
    // A remount rebuilds the body from the live session; re-resolve the section so undo/redo of a
    // section-affecting edit (or a pagination toggle) leaves the bands describing the right one.
    this.syncRegionToBody(this.activeBlock ?? undefined);
  }
}

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
    Save: (handle: number) => Uint8Array;
    Undo: (handle: number) => boolean;
    Redo: (handle: number) => boolean;
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
    // Skip generated list-marker spans — they aren't part of the paragraph's content.
    if (child.nodeType === 1 && (child as HTMLElement).hasAttribute?.("data-list-marker")) return;
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

/** True if `node` is, or is inside, a generated list-marker span (not editable content). */
function isInMarker(node: Node | null): boolean {
  let el: HTMLElement | null = node && node.nodeType === 1 ? (node as HTMLElement) : node?.parentElement ?? null;
  while (el) {
    if (el.hasAttribute && el.hasAttribute("data-list-marker")) return true;
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
  const content = blockContentText(block);
  const leading = content.length - content.replace(/^\s+/, "").length;
  const trimmedLen = content.trim().length;
  return Math.max(0, Math.min(domOffset - Math.min(domOffset, leading), trimmedLen));
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
  return { start: Math.min(start, end), length: Math.abs(end - start) };
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
    /* renderFootnotesAndEndnotes */ false, /* renderHeadersAndFooters */ paginated,
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

  /** The editable block (contenteditable [data-anchor]) containing `node`, if any, within this editor. */
  private editableBlockOf(node: Node | null): HTMLElement | null {
    if (!node) return null;
    const start = node.nodeType === 1 ? (node as HTMLElement) : node.parentElement;
    const block = start?.closest<HTMLElement>('[data-anchor][contenteditable="true"]') ?? null;
    return block && this.editRoot.contains(block) ? block : null;
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
      onEdit: options.onEdit,
    };
    // persistAnchorIds=true keeps PtOpenXml:Unid attributes in Save() output, so a remount's
    // full re-render keeps the SAME unids the live session uses (a content change like becoming
    // a list otherwise re-derives a fresh unid, leaving the block unwired). The cost is that
    // saved bytes carry the Unid attributes (Word ignores them).
    const handle = exports.DocxSessionBridge.OpenSession(bytes, '{"persistAnchorIds":true}');
    const editor = new DocxEditor(container, exports, handle, opts);
    editor.refreshAnchorMap();
    const fullHtml = exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(bytes, opts.cssPrefix, opts.fabricateClasses, opts.paginated, opts.scale),
    );
    if (opts.paginated) editor.mountPaginated(fullHtml);
    else editor.mountHtml(fullHtml);
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

  /** Rebuild unid → full-anchor-id from the live session projection. */
  private refreshAnchorMap(): void {
    const proj = JSON.parse(this.exports.DocxSessionBridge.Project(this.handle)) as {
      anchorIndex: Record<string, AnchorTargetLite>;
    };
    this.unidToFullId.clear();
    for (const [fullId, target] of Object.entries(proj.anchorIndex)) {
      this.unidToFullId.set(target.unid, fullId);
    }
  }

  /** Continuous (non-paginated) mount: inject the converter's styles + body, wire blocks. */
  private mountHtml(fullHtml: string): void {
    const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    this.container.innerHTML = styles + parsed.body.innerHTML;
    this.editRoot = this.container;
    if (this.options.editable) this.wireBlocks(this.container);
  }

  /** Paginated mount: flow blocks into page boxes via pagination.ts, wire the page clones. */
  private mountPaginated(fullHtml: string): void {
    paginateHtml(fullHtml, this.container, { scale: this.options.scale, cssPrefix: "page-" });
    // pagination.ts measures the hidden #pagination-staging subtree ONCE, then flows CLONES of its
    // blocks into the visible page boxes. Leaving staging in the live DOM is a trap: every
    // data-anchor exists twice (staging + page-box copy), so document.querySelector('[data-anchor]')
    // is ambiguous (hits the hidden copy first), and the staging copy goes stale because edits land
    // only on the page-box copy — a future reflow-from-staging would silently revert them. Staging
    // is a transient measurement scaffold; drop it so the page-box copies are the single source of
    // truth. A remount (setPaginated, list/undo edits) rebuilds staging fresh from the live session.
    this.container.querySelector("#pagination-staging, .page-staging")?.remove();
    const pageRoot =
      this.container.querySelector<HTMLElement>("#pagination-container") ?? this.container;
    this.editRoot = pageRoot;
    if (this.options.editable) this.wireBlocks(pageRoot);
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
    if (!unid || !this.unidToFullId.has(unid)) return;
    el.setAttribute("contenteditable", "true");
    // Generated list markers (number/bullet + suffix) are not editable content — keep the
    // caret out of them so offsets stay aligned with the paragraph's run text.
    el.querySelectorAll<HTMLElement>("[data-list-marker]").forEach((m) => m.setAttribute("contenteditable", "false"));
    // Baseline for the commit diff: CONTENT text (list markers + injected bidi marks excluded),
    // matching the session's flat run-text offset space.
    el.dataset.committedText = blockContentText(el);
    el.addEventListener("focus", () => { this.activeBlock = el; });
    el.addEventListener("blur", () => this.commitBlock(el));
    el.addEventListener("keydown", (ev) => this.onKeydown(el, ev as KeyboardEvent));
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
    const fullId = this.unidToFullId.get(unid);
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
      if (fresh && this.replaceNode(el, fresh)) {
        this.unidToFullId.delete(unid);
        this.unidToFullId.set(newUnid, newAnchor);
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
    let fullId = this.unidToFullId.get(unid);
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
      this.remount(idx + 1, false);
      this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
      return;
    }

    const firstEl = this.renderInto(first.id);
    const secondEl = this.renderInto(second.id);
    if (!firstEl || !secondEl) return;

    // el is the focused block — replaceNode guards the re-entrant blur→commit and tolerates a
    // node detached mid-focus-transfer; replacing with both new blocks at once keeps them adjacent.
    if (!this.replaceNode(el, firstEl, secondEl)) return;
    this.unidToFullId.delete(unid);
    this.unidToFullId.set(first.unid, first.id);
    this.unidToFullId.set(second.unid, second.id);
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
    let prevId = this.unidToFullId.get(prevUnid);
    let thisId = this.unidToFullId.get(thisUnid);
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
      this.remount(prevIdx, true);
      this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
      return;
    }

    const mergedEl = this.renderInto(merged.id);
    if (!mergedEl) return;

    // prev may be focused — replaceNode guards re-entrancy and tolerates a detached node.
    if (!this.replaceNode(prev, mergedEl)) return;
    el.remove();
    this.unidToFullId.delete(prevUnid);
    this.unidToFullId.delete(thisUnid);
    this.unidToFullId.set(merged.unid, merged.id);
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

  /** The editable block immediately before `el` in document order, or null. */
  private previousEditable(el: HTMLElement): HTMLElement | null {
    const all = Array.from(
      this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
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
    const all = Array.from(
      this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
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
    const contentLen = blockContentText(block).length;
    if (hasStart) {
      const start = contentOffsetOf(block, range.startContainer, range.startOffset);
      return { start, length: Math.max(0, contentLen - start) };
    }
    if (hasEnd) {
      const end = contentOffsetOf(block, range.endContainer, range.endOffset);
      return { start: 0, length: end };
    }
    return { start: 0, length: contentLen }; // fully-spanned middle block
  }

  /** Apply an inline ApplyFormat op to each block's slice of the selection, then re-render.
   *  Returns false (caller falls back to the single-block path) for a 1-block selection. */
  private applyInlineOpAcrossBlocks(blocks: HTMLElement[], op: object): boolean {
    if (blocks.length <= 1) return false;
    const targets = blocks.map((b) => {
      const unid = b.getAttribute("data-anchor");
      return { block: b, fullId: unid ? this.unidToFullId.get(unid) : undefined, span: this.blockSpanForSelection(b) };
    });
    for (const t of targets) {
      if (!t.fullId || (t.span && t.span.length === 0)) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      this.exports.DocxSessionBridge.ApplyFormat(
        this.handle, synced, t.span ? JSON.stringify(t.span) : "", JSON.stringify(op),
      );
    }
    this.remount();
    return true;
  }

  /** Apply a whole-block (paragraph-level) op to each selected block, then re-render.
   *  Returns false for a 1-block selection (caller uses the single-block path). */
  private applyParagraphOpAcrossBlocks(blocks: HTMLElement[], run: (fullId: string) => string): boolean {
    if (blocks.length <= 1) return false;
    const targets = blocks.map((b) => {
      const unid = b.getAttribute("data-anchor");
      return { block: b, fullId: unid ? this.unidToFullId.get(unid) : undefined };
    });
    for (const t of targets) {
      if (!t.fullId) continue;
      const synced = this.syncBlock(t.block, t.fullId);
      run(synced);
    }
    this.remount();
    return true;
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
    let fullId = this.unidToFullId.get(unid);
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
    if (this.affectsList(res)) { this.remount(this.blockIndex(block), false); return; }
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
    let fullId = this.unidToFullId.get(unid);
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
    if (this.affectsList(res)) { this.remount(this.blockIndex(block), false); return; }
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
    let fullId = this.unidToFullId.get(unid);
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
    if (this.affectsList(res)) { this.remount(this.blockIndex(block), false); return; }
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
    let fullId = this.unidToFullId.get(unid);
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
    this.remount(idx, false);
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
    let fullId = this.unidToFullId.get(unid);
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
    this.remount(idx, false);
  }

  // ─── Table row / column editing (active block must be inside a table cell) ──────────

  /** Run a table-structure op on the active cell (a cell-paragraph block) and re-render. */
  private tableEdit(run: (cellAnchor: string) => string): void {
    const block = this.activeBlock;
    if (this.closed || !block || !block.closest("table")) return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId); // flush uncommitted cell text first
    const res = this.parseEdit(run(fullId));
    if (!res.success) return;
    this.remount(idx, false);
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
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetListLevel(this.handle, fullId, delta));
    if (!res.success) return;
    // A level change ripples through the whole list's numbering — re-render with full document
    // context (a single-block render can't compute nested numbering), keeping the caret in place.
    this.remount(idx, false);
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
    let fullId = this.unidToFullId.get(unid);
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
    this.remount(idx, false);
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
    const fullId = this.unidToFullId.get(unid);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    const res = this.parseEdit(this.exports.DocxSessionBridge.DeleteBlock(this.handle, fullId));
    if (!res.success) return;
    this.remount(Math.max(0, idx - 1), true);
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
      this.applyParagraphOpAcrossBlocks(blocks, (id) =>
        this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, id, JSON.stringify(op)),
      )
    )
      return;
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.SetParagraphFormat(this.handle, fullId, JSON.stringify(op)),
    );
    if (!res.success) return;
    // A border change adds/removes the wrapping border <div>, so a single-block swap can't restructure
    // it correctly — re-render fully (like list edits) so the wrapper appears/disappears cleanly.
    if (this.affectsList(res) || op.clearBorders) { this.remount(idx, false); return; }
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
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    fullId = this.syncBlock(block, fullId);
    const res = this.parseEdit(this.exports.DocxSessionBridge.SetParagraphStyle(this.handle, fullId, styleId));
    if (!res.success) return;
    if (this.affectsList(res)) { this.remount(idx, false); return; }
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
    const anchorId = ref?.id ?? this.unidToFullId.get(oldUnid);
    const newUnid = ref?.unid ?? oldUnid;
    if (!anchorId) return null;
    const fresh = this.renderInto(anchorId);
    if (!fresh || !this.replaceNode(oldEl, fresh)) return null;
    this.unidToFullId.delete(oldUnid);
    this.unidToFullId.set(newUnid, anchorId);
    this.wireBlock(fresh);
    this.activeBlock = fresh;
    this.options.onEdit?.({ anchorId, unid: newUnid });
    return fresh;
  }

  /** Editable blocks in document order. */
  private editableList(): HTMLElement[] {
    return Array.from(this.editRoot.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
  }

  private blockIndex(el: HTMLElement): number {
    return this.editableList().indexOf(el);
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
    const bytes = this.exports.DocxSessionBridge.Save(this.handle);
    const fullHtml = this.exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(bytes, this.options.cssPrefix, this.options.fabricateClasses, this.options.paginated, this.options.scale),
    );
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
  }
}

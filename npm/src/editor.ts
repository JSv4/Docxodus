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
    Project: (handle: number) => string;
    ReplaceText: (handle: number, anchor: string, md: string) => string;
    SplitParagraph: (handle: number, anchor: string, offset: number) => string;
    MergeParagraphs: (handle: number, first: string, second: string) => string;
    ApplyFormat: (handle: number, anchor: string, spanJson: string, opJson: string) => string;
    SetParagraphStyle: (handle: number, anchor: string, styleId: string) => string;
    SetParagraphFormat: (handle: number, anchor: string, opJson: string) => string;
    ApplyListFormat: (handle: number, anchor: string, kind: string) => string;
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
  if (seg.text === "\n") return "\n";
  let md = escapeInlineMarkdown(seg.text);
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
 * Content-text offset of (container, offset) within `block`, EXCLUDING generated list-marker
 * text. This is the offset DocxSession ops expect (the paragraph's run text, not the rendered
 * number/bullet the converter injects).
 */
function contentOffsetOf(block: HTMLElement, container: Node, offset: number): number {
  let count = 0;
  let done = false;
  const walk = (node: Node): void => {
    if (done) return;
    if (node.nodeType === 3 /* TEXT_NODE */) {
      if (node === container) { if (!isInMarker(node)) count += offset; done = true; return; }
      if (!isInMarker(node)) count += node.textContent?.length ?? 0;
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
      const len = node.textContent?.length ?? 0;
      if (remaining <= len) {
        range.setStart(node, remaining);
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
  if (!sel) return;
  el.focus();
  const range = document.createRange();
  const end = start + length;
  let pos = 0;
  let startSet = false;
  const walk = (node: Node): boolean => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 3 /* TEXT_NODE */) {
        if (isInMarker(child)) continue; // marker text isn't part of the content offset space
        const len = child.textContent?.length ?? 0;
        if (!startSet && pos + len >= start) {
          range.setStart(child, start - pos);
          startSet = true;
        }
        if (startSet && pos + len >= end) {
          range.setEnd(child, end - pos);
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

  /** Lossless DOCX bytes reflecting all edits. */
  save(): Uint8Array {
    this.assertOpen();
    return this.exports.DocxSessionBridge.Save(this.handle);
  }

  /** Release the underlying WASM session. The editor is unusable afterward. */
  close(): void {
    if (this.closed) return;
    this.closed = true;
    this.exports.DocxSessionBridge.CloseSession(this.handle);
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
    // pagination clones blocks into pages (hidden originals stay in #pagination-staging),
    // so wire ONLY the visible page container — never the staging copies.
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
    // Only blocks the markdown projection addresses (top-level body paragraphs /
    // headings) are editable via the text path. Paragraphs inside opaque tables are
    // stamped in the HTML but not individually indexed, so they stay read-only in v1.
    if (!unid || !this.unidToFullId.has(unid)) return;
    el.setAttribute("contenteditable", "true");
    // Generated list markers (number/bullet + suffix) are not editable content — keep the
    // caret out of them so offsets stay aligned with the paragraph's run text.
    el.querySelectorAll<HTMLElement>("[data-list-marker]").forEach((m) => m.setAttribute("contenteditable", "false"));
    el.dataset.committedText = el.textContent ?? "";
    el.addEventListener("focus", () => { this.activeBlock = el; });
    el.addEventListener("blur", () => this.commitBlock(el));
    el.addEventListener("keydown", (ev) => this.onKeydown(el, ev as KeyboardEvent));
  }

  /** Commit a block edit on blur: DocxSession op → re-render only this block → patch DOM. */
  private commitBlock(el: HTMLElement): void {
    if (this.closed) return;
    const unid = el.getAttribute("data-anchor");
    if (!unid) return;
    const newText = (el.textContent ?? "").trim();
    if (newText === (el.dataset.committedText ?? "").trim()) return; // no change

    const fullId = this.unidToFullId.get(unid);
    if (!fullId) return;

    // M1: serialize the block's inline content to markdown so bold/italic/links survive
    // the edit, instead of flattening to plain text.
    const markdown = serializeInlineMarkdown(el);
    const result = this.parseEdit(this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, markdown));
    if (!result.success) {
      // Revert the DOM to the last committed text so the view stays in sync with truth.
      el.textContent = el.dataset.committedText ?? "";
      return;
    }

    const newAnchor = result.modified?.[0]?.id ?? fullId;
    const newUnid = result.modified?.[0]?.unid ?? unid;

    // List items: do NOT re-render on a text commit. A text edit never changes the item's number,
    // and commit fires on blur — re-rendering replaces this block's DOM node *during* the blur,
    // which cancels the browser's in-flight focus transfer when the user clicks straight to
    // another bullet (focus falls to <body>, so typing into the next bullet does nothing). The
    // DOM already shows exactly what the user typed, with the correct marker, so we leave it in
    // place and only sync the session (done above) + bookkeeping. The marker would otherwise need
    // whole-document context to renumber, which a single-block render lacks anyway.
    if (el.querySelector(":scope > [data-list-marker]")) {
      el.dataset.committedText = el.textContent ?? "";
      this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
      return;
    }

    // Plain block: re-render ONLY this block from the live session and patch it in place for
    // canonical HTML. Swapping the just-blurred node here is safe (verified — focus stays on the
    // newly-clicked block).
    const html = this.exports.DocxSessionBridge.RenderBlockHtml(
      this.handle,
      newAnchor,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) !== 0x7b /* not an error object */) {
      const fresh = new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
      if (fresh) {
        el.replaceWith(fresh);
        // Keep the anchor map current and re-wire the replacement.
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
    if (ev.key === "Enter" && !ev.shiftKey && !ev.isComposing) {
      ev.preventDefault();
      this.splitAtCaret(el);
    } else if (ev.key === "Backspace") {
      const sel = typeof window !== "undefined" ? window.getSelection() : null;
      if (sel && sel.isCollapsed && caretOffsetIn(el) === 0) {
        const prev = this.previousEditable(el);
        if (prev) {
          ev.preventDefault();
          this.mergeWithPrevious(prev, el);
        }
      }
    }
  }

  /** Enter: split the block at the caret into two paragraphs. */
  private splitAtCaret(el: HTMLElement): void {
    const offset = caretOffsetIn(el);
    const unid = el.getAttribute("data-anchor");
    if (offset == null || !unid) return;
    let fullId = this.unidToFullId.get(unid);
    if (!fullId) return;

    const idx = this.blockIndex(el); // capture before the op (for list remount focus)
    fullId = this.syncBlock(el, fullId); // flush any uncommitted typing first
    const res = this.parseEdit(this.exports.DocxSessionBridge.SplitParagraph(this.handle, fullId, offset));
    if (!res.success) return;
    const first = res.modified?.[0];
    const second = res.created?.[0];
    if (!first || !second) return;

    // Splitting a list item makes a continuing list item — re-render the whole document so
    // numbering continues and the new item shows its marker; put the caret in the new item.
    if (this.affectsList(res)) {
      this.remount(idx + 1, false);
      this.options.onEdit?.({ anchorId: second.id, unid: second.unid });
      return;
    }

    const firstEl = this.renderInto(first.id);
    const secondEl = this.renderInto(second.id);
    if (!firstEl || !secondEl) return;

    el.replaceWith(firstEl);
    firstEl.after(secondEl);
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
    prevId = this.syncBlock(prev, prevId);
    thisId = this.syncBlock(el, thisId);
    const caret = (prev.textContent ?? "").length; // merge boundary

    const res = this.parseEdit(this.exports.DocxSessionBridge.MergeParagraphs(this.handle, prevId, thisId));
    if (!res.success) return;
    const merged = res.modified?.[0];
    if (!merged) return;

    // Merging list items renumbers the list — re-render fully, caret at the merge boundary.
    if (this.affectsList(res)) {
      this.remount(prevIdx, true);
      this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
      return;
    }

    const mergedEl = this.renderInto(merged.id);
    if (!mergedEl) return;

    prev.replaceWith(mergedEl);
    el.remove();
    this.unidToFullId.delete(prevUnid);
    this.unidToFullId.delete(thisUnid);
    this.unidToFullId.set(merged.unid, merged.id);
    this.wireBlock(mergedEl);
    placeCaretAtOffset(mergedEl, caret);
    this.options.onEdit?.({ anchorId: merged.id, unid: merged.unid });
  }

  /** Flush a block's current (uncommitted) text to the session; returns the live full id. */
  private syncBlock(el: HTMLElement, fullId: string): string {
    const cur = (el.textContent ?? "").trim();
    if (cur === (el.dataset.committedText ?? "").trim()) return fullId; // unchanged
    const res = this.parseEdit(
      this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, serializeInlineMarkdown(el)),
    );
    return res.success ? res.modified?.[0]?.id ?? fullId : fullId;
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

  /**
   * Toggle (or set) an inline format on the current selection in the active block.
   * With no selection, applies to the whole paragraph. Routes through DocxSession
   * (`ApplyFormat`) so it is lossless and supports underline/strike, not just markdown.
   */
  format(key: FormatKey, value?: boolean): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
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

  /** Set paragraph alignment (left/center/right/justify) on the active block. */
  setAlignment(alignment: EditorAlignment): void {
    this.applyParagraphFormat({ alignment });
  }

  /** Adjust the active block's left indent by `deltaTwips` (default ±720 = 0.5"), clamped at 0. */
  indent(deltaTwips = 720): void {
    this.applyParagraphFormat({ indentDelta: deltaTwips });
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

  private applyParagraphFormat(op: {
    alignment?: EditorAlignment;
    indentDelta?: number;
    pageBreakBefore?: boolean;
  }): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
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
    if (this.affectsList(res)) { this.remount(idx, false); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /** Set the paragraph style of the active block (e.g. "Heading1", "Heading2", "Normal"). */
  setParagraphStyle(styleId: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
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
    if (!fresh) return null;
    oldEl.replaceWith(fresh);
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

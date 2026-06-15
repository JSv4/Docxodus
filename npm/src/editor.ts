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
    RenderBlockHtml: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    Save: (handle: number) => Uint8Array;
    Undo: (handle: number) => string;
    Redo: (handle: number) => string;
  };
  DocumentConverter: {
    ConvertDocxToHtmlComplete: (...args: any[]) => string;
  };
}

export interface DocxEditorOptions {
  /** CSS class prefix for rendered HTML. Default "docx-". */
  cssPrefix?: string;
  /** Fabricate CSS classes (vs inline styles). Default true. */
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
      fabricateClasses: options.fabricateClasses ?? true,
      editable: options.editable ?? true,
      paginated: options.paginated ?? false,
      scale: options.scale ?? 1,
      onEdit: options.onEdit,
    };
    const handle = exports.DocxSessionBridge.OpenSession(bytes, "");
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
    if (this.options.editable) this.wireBlocks(this.container);
  }

  /** Paginated mount: flow blocks into page boxes via pagination.ts, wire the page clones. */
  private mountPaginated(fullHtml: string): void {
    paginateHtml(fullHtml, this.container, { scale: this.options.scale, cssPrefix: "page-" });
    // pagination clones blocks into pages (hidden originals stay in #pagination-staging),
    // so wire ONLY the visible page container — never the staging copies.
    const pageRoot =
      this.container.querySelector<HTMLElement>("#pagination-container") ?? this.container;
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
    el.dataset.committedText = el.textContent ?? "";
    el.addEventListener("blur", () => this.commitBlock(el));
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

    const result = JSON.parse(this.exports.DocxSessionBridge.ReplaceText(this.handle, fullId, newText)) as {
      success: boolean;
      modified?: Array<{ id: string; unid: string }>;
    };
    if (!result.success) {
      // Revert the DOM to the last committed text so the view stays in sync with truth.
      el.textContent = el.dataset.committedText ?? "";
      return;
    }

    const newAnchor = result.modified?.[0]?.id ?? fullId;
    const newUnid = result.modified?.[0]?.unid ?? unid;

    // Re-render ONLY this block from the live session and patch it in place.
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
      }
    }

    this.options.onEdit?.({ anchorId: newAnchor, unid: newUnid });
  }
}

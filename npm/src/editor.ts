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
import { DocumentViewport } from "./viewport.js";
import type { ColumnWidth } from "./viewport.js";
import { HeaderFooterRegion } from "./editor-headerfooter.js";
import type { BandWhich } from "./editor-headerfooter.js";
import {
  draggable,
  dropTargetForElements,
  monitorForElements,
} from "@atlaskit/pragmatic-drag-and-drop/element/adapter";
import { setCustomNativeDragPreview } from "@atlaskit/pragmatic-drag-and-drop/element/set-custom-native-drag-preview";
import { pointerOutsideOfPreview } from "@atlaskit/pragmatic-drag-and-drop/element/pointer-outside-of-preview";
import {
  autoScrollForElements,
  autoScrollWindowForElements,
} from "@atlaskit/pragmatic-drag-and-drop-auto-scroll/element";
import { TrackedChangeMode } from "./types.js";
import type { HeaderFooterKind, NumberFormat } from "./types.js";
import { diffUnits, needsRemount, tokenOf, unidOf } from "./editor-reconcile.js";
import type { RenderPlan, RenderUnit, UnitDiff } from "./editor-reconcile.js";

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
    MoveBlock?: (handle: number, sourceAnchor: string, targetAnchor: string, pos: string) => string;
    /** JSON `{anchorId, before, after}[]`: where a block may legally move, per side. Optional —
     *  without it the drag UI offers every block and lets the engine refuse, as it did before. */
    ValidMoveTargets?: (handle: number, sourceAnchor: string) => string;
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
    RenderBlockHtmlForReview?: (
      handle: number,
      anchorId: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean,
    ) => string;
    /** Session-attached full-document render (optional: older WASM bundles lack it). */
    RenderHtml?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
    ) => string;
    RenderHtmlForReview?: (
      handle: number,
      cssPrefix: string,
      fabricateClasses: boolean,
      paginated: boolean,
      scale: number,
      renderTrackedChanges: boolean,
    ) => string;
    Save: (handle: number) => Uint8Array;
    /** Save keeping the projector's Unid bookkeeping — remount only, never a user download.
     *  Optional so a bridge predating it still satisfies this type. */
    SaveWithAnchorIds?: (handle: number) => Uint8Array;
    Undo: (handle: number) => boolean;
    Redo: (handle: number) => boolean;
    /** Header/footer region: the section a body anchor belongs to (kind → part mapping). */
    GetSectionInfo: (handle: number, anchorId: string) => string;
    SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
    InsertPageNumberField: (handle: number, anchor: string, field: string, format: string) => string;
    EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
    SetPageNumbering: (handle: number, anchor: string, opJson: string) => string;
    ClearPageNumbering: (handle: number, anchor: string) => string;
    /** Note authoring (optional: older WASM bundles predate it). */
    InsertFootnote?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      markdown: string,
    ) => string;
    InsertEndnote?: (
      handle: number,
      anchor: string,
      characterOffset: number,
      markdown: string,
    ) => string;
    /** Incremental-reconcile endpoints (optional: older WASM bundles predate them;
     *  the editor falls back to full remounts / full projections without them). */
    ListBlocks?: (handle: number) => string;
    ListRenderedBlocks?: (handle: number, renderTrackedChanges: boolean) => string;
    ListNotes?: (handle: number, endnotes: boolean) => string;
    ListAnchors?: (handle: number) => string;
    RenderBlocksHtml?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean,
    ) => string;
    RenderBlocksHtmlForReview?: (
      handle: number,
      anchorIdsJson: string,
      cssPrefix: string,
      fabricateClasses: boolean,
      renderTrackedChanges: boolean,
    ) => string;
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
  /** Author-pinned page render scale (1.0 = 100%). Default 1. */
  scale?: number;
  /**
   * How the continuous view sizes its text column. `"section"` (default) uses the width the
   * document's `w:sectPr` defines, so line breaking matches Word on every screen; `"fluid"`
   * lets the column follow the host, the pre-geometry behavior.
   */
  columnWidth?: ColumnWidth;
  /**
   * Zoom the page down to fit a host narrower than it, the way a word processor's
   * fit-to-width does, instead of letting it overflow. Default true.
   */
  fitToWidth?: boolean;
  /**
   * Render docked Header/Footer editing bands around the body flow. Default FALSE — with it off
   * the editor's DOM is unchanged, so existing consumers are unaffected. When on, the body flow
   * is wrapped in a `.docx-body-flow` element that becomes the edit root, keeping band blocks out
   * of the body's block list (which indexes remount focus).
   */
  headerFooter?: boolean;
  /** Enable the editor-owned block drag handle. Default false (the ribbon defaults it to true). */
  blockDrag?: boolean;
  /** How editor mutations are recorded. RenderInline enables native Word revisions. */
  trackedChanges?: TrackedChangeMode;
  /** Author stamped on native Word revisions. Default "docxodus". */
  revisionAuthor?: string;
  /** Called after a block edit commits (with the affected anchor). */
  onEdit?: (info: { anchorId: string; unid: string }) => void;
  /** Called after a successful block move. */
  onMove?: (info: { sourceAnchorId: string; destinationAnchorId: string }) => void;
}

interface AnchorTargetLite {
  unid: string;
  kind: string;
  scope: string;
  textPreview?: string;
}

const EDITABLE_TAGS = new Set(["P", "H1", "H2", "H3", "H4", "H5", "H6"]);

const BLOCK_DRAG_TYPE = "docxodus-block";
const blockDragStyledDocuments = new WeakSet<Document>();

/** One-line description of a block, for the handle's label and the drag preview chip. */
function blockPreviewText(unit: HTMLElement): string {
  if (unit.tagName === "TABLE") return "Table";
  const text = (unit.textContent ?? "").trim().replace(/\s+/g, " ");
  return text.length > 48 ? `${text.slice(0, 48)}…` : text;
}

/**
 * A capture-time snapshot of one movable block's box, in viewport coordinates.
 *
 * Taken once per drag: nothing reflows while a drag is in flight, so re-measuring every block on
 * every `dragover` would buy nothing. Scrolling (including drag autoscroll) moves the whole flow
 * rigidly, which `dropZoneShift` corrects for with one subtraction.
 */
interface BlockDropZone {
  unit: HTMLElement;
  anchorId: string;
  /** Position in `dropZones`, so a neighbour is one index away — see `dropEdgeY`. */
  index: number;
  top: number;
  bottom: number;
  left: number;
  width: number;
  /** Own outer margin, read only for the first/last zone — see `dropEdgeY`. */
  marginBefore: number;
  marginAfter: number;
}

function ensureBlockDragStyles(doc: Document): void {
  if (blockDragStyledDocuments.has(doc)) return;
  blockDragStyledDocuments.add(doc);
  const style = doc.createElement("style");
  style.dataset.docxodusBlockDrag = "true";
  style.textContent = `
.docx-block-handle {
  position: fixed; z-index: 2147483000; display: none; width: 26px; height: 28px;
  align-items: center; justify-content: center; padding: 0; border: 1px solid #d0d5dd;
  border-radius: 6px; background: #fff; color: #667085; box-shadow: 0 1px 3px rgba(16,24,40,.12);
  cursor: grab; font: 700 17px/1 system-ui, sans-serif; user-select: none; touch-action: none;
}
.docx-block-handle:hover, .docx-block-handle:focus-visible { color: #344054; border-color: #98a2b3; outline: none; }
.docx-block-handle[aria-pressed="true"] { color: #175cd3; border-color: #84adff; background: #eff8ff; }
.docx-block-handle.docx-block-dragging { cursor: grabbing; opacity: .35; }
/* The block being carried. Dimming it is the "a drag is happening" signal that survives the
   pointer being anywhere on screen — the drop line only says where, not what. */
.docx-block-drag-source { opacity: .38; transition: opacity 120ms ease-out; }
/* Positioned by transform so tracking the pointer costs no layout. Flipping display none→block
   restarts the fade — one cheap entry animation per appearance, none while it tracks. */
.docx-block-drop-indicator {
  position: fixed; top: 0; left: 0; z-index: 2147482999; display: none; height: 0;
  pointer-events: none; border-top: 2px solid #2e90fa;
  filter: drop-shadow(0 1px 2px rgba(46,144,250,.5));
  animation: docx-block-drop-in 110ms ease-out;
}
.docx-block-drop-indicator::before {
  content: ""; position: absolute; top: -5px; left: -2px; width: 8px; height: 8px;
  border-radius: 50%; background: #2e90fa;
}
@keyframes docx-block-drop-in { from { opacity: 0; } to { opacity: 1; } }
.docx-block-drag-preview {
  max-width: 320px; padding: 6px 10px; border: 1px solid #b2ddff; border-radius: 6px;
  background: #eff8ff; color: #175cd3; box-shadow: 0 6px 16px rgba(16,24,40,.18);
  font: 500 13px/1.35 system-ui, sans-serif; white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis;
}
.docx-block-move-menu {
  position: fixed; z-index: 2147483001; display: none; min-width: 150px; padding: 5px;
  border: 1px solid #d0d5dd; border-radius: 8px; background: #fff;
  box-shadow: 0 8px 24px rgba(16,24,40,.18); font: 13px/1.35 system-ui, sans-serif;
}
.docx-block-move-menu button { display: block; width: 100%; padding: 7px 9px; border: 0; border-radius: 5px; background: transparent; color: #344054; text-align: left; }
.docx-block-move-menu button:hover, .docx-block-move-menu button:focus-visible { background: #f2f4f7; outline: none; }
.docx-block-move-menu button:disabled { color: #98a2b3; background: transparent; }
.docx-block-move-flash { animation: docx-block-move-flash 800ms ease-out; }
@keyframes docx-block-move-flash { from { box-shadow: 0 0 0 3px rgba(46,144,250,.55); } to { box-shadow: none; } }
.docx-block-live-region { position: fixed; width: 1px; height: 1px; overflow: hidden; clip-path: inset(50%); white-space: nowrap; }
`;
  (doc.head ?? doc.documentElement).appendChild(style);
}

function trackedChangeWireName(mode: TrackedChangeMode): "accept" | "render_inline" | "strip_deletions" {
  if (mode === TrackedChangeMode.RenderInline) return "render_inline";
  if (mode === TrackedChangeMode.StripDeletions) return "strip_deletions";
  return "accept";
}

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

interface DomSelectionPoint {
  node: Node;
  offset: number;
}

interface CrossBlockSelectionBookmark {
  startUnid: string;
  startOffset: number;
  endUnid: string;
  endOffset: number;
}

interface CrossBlockDragSelection {
  anchor: DomSelectionPoint;
  origin: HTMLElement;
  crossedBlockBoundary: boolean;
  focus: DomSelectionPoint | null;
  frame: number | null;
}

/** True if `block` renders as a list item (has a generated marker as its first child). */
function isListBlock(block: HTMLElement): boolean {
  return !!block.querySelector(":scope > [data-list-marker]");
}

/**
 * The converter's tab box must stay text-empty because any child changes its inline baseline.
 * Editable list items still need a textContent separator between generated marker chrome and an
 * empty body so typing produces "2. BB", not "2.BB". Keep that editor-only separator hidden and
 * generated, outside the converter-owned marker wrapper.
 */
function ensureListMarkerSeparator(block: HTMLElement): void {
  const marker = block.querySelector<HTMLElement>(":scope > [data-list-marker]");
  if (!marker) return;
  const next = marker.nextElementSibling as HTMLElement | null;
  if (next?.hasAttribute("data-editor-list-separator")) return;

  const separator = block.ownerDocument.createElement("span");
  separator.setAttribute("data-editor-list-separator", "");
  separator.setAttribute("data-list-marker", "true");
  separator.setAttribute("aria-hidden", "true");
  separator.style.display = "none";
  separator.textContent = " ";
  marker.after(separator);
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

/**
 * Resolve a viewport coordinate to a DOM caret. Firefox exposes the standards-track
 * caretPositionFromPoint; Chromium/WebKit expose caretRangeFromPoint. Keeping the compatibility
 * shim here lets the editor bridge native mouse selection across its intentionally-independent
 * per-paragraph editing hosts without changing their commit boundaries.
 */
function caretPointFromClient(
  doc: Document,
  clientX: number,
  clientY: number,
): DomSelectionPoint | null {
  const caretDoc = doc as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  const position = caretDoc.caretPositionFromPoint?.(clientX, clientY);
  if (position) return { node: position.offsetNode, offset: position.offset };
  const range = caretDoc.caretRangeFromPoint?.(clientX, clientY);
  return range ? { node: range.startContainer, offset: range.startOffset } : null;
}

/**
 * Set a possibly-backward pair of points as one normalized DOM Range. Chromium's
 * Selection.setBaseAndExtent() exposes cross-contenteditable endpoints but clips toString() and
 * the effective contents to the originating host; addRange() carries the complete cross-block
 * contents. Direction is immaterial to editor commands, which consume normalized start/end.
 */
function setSelectionBetween(anchor: DomSelectionPoint, focus: DomSelectionPoint): boolean {
  const sel = typeof window !== "undefined" ? window.getSelection() : null;
  if (!sel || !anchor.node.isConnected || !focus.node.isConnected) return false;
  try {
    const probe = document.createRange();
    probe.setStart(anchor.node, anchor.offset);
    probe.collapse(true);
    const focusIsBefore = probe.comparePoint(focus.node, focus.offset) < 0;
    const range = document.createRange();
    const start = focusIsBefore ? focus : anchor;
    const end = focusIsBefore ? anchor : focus;
    range.setStart(start.node, start.offset);
    range.setEnd(end.node, end.offset);
    sel.removeAllRanges();
    sel.addRange(range);
    return true;
  } catch {
    return false;
  }
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
  renderTrackedChanges: boolean,
): any[] {
  return [
    bytes, "Document", cssPrefix, fabricate, "", -1, "comment-",
    /* paginationMode */ paginated ? 1 : 0, /* paginationScale */ scale, "page-",
    false, 0, "annot-",
    // Footnotes/endnotes ON: they are document content, and the editor makes the rendered note
    // paragraphs editable. Must stay in step with DocxSessionOps.RenderHtml (the remount path),
    // whose output has to match this first paint byte-for-byte.
    /* renderFootnotesAndEndnotes */ true, /* renderHeadersAndFooters */ paginated,
    renderTrackedChanges, true, true, false, null, /* stampAnchors */ true,
  ];
}

export class DocxEditor {
  private readonly exports: DocxEditorExports;
  private readonly container: HTMLElement;
  private readonly handle: number;
  private readonly options: Required<Omit<DocxEditorOptions, "onEdit" | "onMove">> &
    Pick<DocxEditorOptions, "onEdit" | "onMove">;
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

  /** Editor-owned block-move chrome. It is deliberately outside the rendered unit tree. */
  private blockDragHandle: HTMLElement | null = null;
  private blockDropIndicator: HTMLElement | null = null;
  private blockMoveMenu: HTMLElement | null = null;
  private blockMoveLive: HTMLElement | null = null;
  private blockDragSource: HTMLElement | null = null;
  private blockDragCleanup: Array<() => void> = [];
  /** Block boxes measured at drag start — see `BlockDropZone`. Empty when no drag is in flight. */
  private dropZones: BlockDropZone[] = [];
  /** Combined scroll offset when `dropZones` was measured, and the scroller measured against. */
  private dropZoneOrigin = 0;
  private dropZoneScroller: HTMLElement | null = null;
  private blockDragging = false;
  private blockDragPointerDown = false;
  /** Anchors the current drag source may legally move next to, per the engine's own rules.
   *  Null when the bridge predates ValidMoveTargets — then every block is offered, as before. */
  private blockMoveTargets: Map<string, { before: boolean; after: boolean }> | null = null;
  /** Memoized `ValidMoveTargets` answers, keyed by source anchor and dropped whenever an edit
   *  lands. The legal-target set is a property of the document, so hovering back and forth over
   *  the same blocks between edits must not re-ask the engine. */
  private blockMoveTargetCache = new Map<string, Map<string, { before: boolean; after: boolean }> | null>();
  /** Cancels the pending idle prefetch of the hovered block's targets — see `showBlockHandle`. */
  private blockMoveTargetPrefetch: (() => void) | null = null;
  /** Why the last move was refused, verbatim from the engine — diagnostics, not announcement copy. */
  lastMoveError: string | null = null;

  /**
   * The last real (non-collapsed) text selection inside an editable block. A toolbar control that
   * must take focus to be used — the font-size combobox — blurs the block and collapses the live
   * selection, so without this an operation triggered from such a control could only target the
   * whole paragraph (S-1 smoke-test finding 3). Refreshed whenever a non-empty selection sits in a
   * block, and cleared when a caret is collapsed inside a block (so it never goes stale).
   */
  private lastSelection: { unid: string; span: { start: number; length: number } } | null = null;

  /**
   * Stable bookmark for a selection spanning independent editable blocks. Native controls such as
   * a font-size combobox can take focus and collapse the live DOM selection; block ids + content
   * offsets let the command restore that same range before applying. This is the multi-block
   * counterpart to lastSelection above.
   */
  private lastCrossBlockSelection: CrossBlockSelectionBookmark | null = null;

  /** State for the mouse-selection bridge between independent contenteditable block hosts. */
  private dragSelection: CrossBlockDragSelection | null = null;

  /** The docked header/footer bands, when `options.headerFooter` is on. */
  private region: HeaderFooterRegion | null = null;

  /** Page geometry + fit-to-width zoom for the mounted document. Re-attached on every mount. */
  private readonly viewport: DocumentViewport;

  /** Why the last reconcile() fell back to a full remount (null = it patched). For
   *  diagnostics/specs; not part of the public API. */
  private lastReconcileFallback: string | null = null;

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
    this.viewport = new DocumentViewport(container, {
      columnWidth: options.columnWidth,
      fitToWidth: options.fitToWidth,
      scale: options.scale,
    });
    if (typeof document !== "undefined") {
      document.addEventListener("selectionchange", this.onSelectionChange);
      document.addEventListener("mousedown", this.onMouseDown, true);
      document.addEventListener("mousemove", this.onMouseMove, true);
      document.addEventListener("mouseup", this.onMouseUp, true);
    }
  }

  /** Track the last meaningful selection so focus-stealing toolbar controls can still target it. */
  private readonly onSelectionChange = (): void => {
    if (this.closed) return;
    const sel = typeof window !== "undefined" ? window.getSelection() : null;
    if (!sel || sel.rangeCount === 0) return; // no selection info — keep the cache as-is
    const range = sel.getRangeAt(0);
    const startBlock = this.editableBlockOf(range.startContainer);
    const endBlock = this.editableBlockOf(range.endContainer);
    const block = startBlock ?? endBlock;
    if (!block) return; // selection is outside the editor (e.g. a toolbar field) — keep the cache
    if (range.collapsed) {
      this.lastSelection = null; // an explicit caret in a block — drop any stale selection
      this.lastCrossBlockSelection = null;
      return;
    }
    if (
      startBlock && endBlock && startBlock !== endBlock &&
      this.ownerRoot(startBlock) === this.ownerRoot(endBlock)
    ) {
      const startUnid = startBlock.getAttribute("data-anchor");
      const endUnid = endBlock.getAttribute("data-anchor");
      if (startUnid && endUnid) {
        this.lastSelection = null;
        this.lastCrossBlockSelection = {
          startUnid,
          startOffset: contentOffsetOf(startBlock, range.startContainer, range.startOffset),
          endUnid,
          endOffset: contentOffsetOf(endBlock, range.endContainer, range.endOffset),
        };
      }
      return;
    }
    const unid = block.getAttribute("data-anchor");
    if (!unid) return;
    this.lastCrossBlockSelection = null;
    const span = selectionSpanIn(block);
    if (span) this.lastSelection = { unid, span };
  };

  /** Start tracking a normal primary-button text drag inside one editable block. */
  private readonly onMouseDown = (event: MouseEvent): void => {
    if (this.closed || !this.options.editable || event.button !== 0) return;
    const target = event.target instanceof Node ? event.target : null;
    const origin = this.editableBlockOf(target);
    if (!origin || isInMarker(target)) return;
    const anchor = caretPointFromClient(document, event.clientX, event.clientY);
    if (!anchor || this.editableBlockOf(anchor.node) !== origin) return;
    this.clearDragSelection();
    this.dragSelection = {
      anchor,
      origin,
      crossedBlockBoundary: false,
      focus: null,
      frame: null,
    };
  };

  /**
   * Apply the latest cross-block endpoint after the browser finishes its native mousemove
   * selection update. Firefox rewrites Selection back into the originating contenteditable after
   * event dispatch even when mousemove is cancelled; writing in requestAnimationFrame wins that
   * race and happens before paint. Coalescing also avoids rebuilding a Range for every raw pointer
   * event when the mouse is moving faster than the display can refresh.
   */
  private queueDragSelection(drag: CrossBlockDragSelection, focus: DomSelectionPoint): void {
    drag.focus = focus;
    if (drag.frame !== null) return;
    const view = this.container.ownerDocument.defaultView;
    if (!view) {
      setSelectionBetween(drag.anchor, focus);
      return;
    }
    drag.frame = view.requestAnimationFrame(() => {
      drag.frame = null;
      if (this.dragSelection !== drag || !drag.focus) return;
      setSelectionBetween(drag.anchor, drag.focus);
    });
  }

  /** Cancel a queued repaint and discard the current gesture. */
  private clearDragSelection(): void {
    const drag = this.dragSelection;
    if (drag && drag.frame !== null) {
      this.container.ownerDocument.defaultView?.cancelAnimationFrame(drag.frame);
    }
    this.dragSelection = null;
  }

  /**
   * Browsers fence native mouse selection at a contenteditable host boundary. Once a drag reaches
   * another block in the same OOXML story, take over just that gesture and create the cross-block
   * Selection the editor's existing multi-block command path consumes. Intra-block selection stays
   * entirely native, and separate hosts remain intact for safe per-anchor commits.
   */
  private readonly onMouseMove = (event: MouseEvent): void => {
    const drag = this.dragSelection;
    if (!drag) return;
    if ((event.buttons & 1) === 0) {
      this.clearDragSelection();
      return;
    }
    const focus = caretPointFromClient(document, event.clientX, event.clientY);
    if (!focus || isInMarker(focus.node)) return;
    const focusBlock = this.editableBlockOf(focus.node);
    if (!focusBlock || this.ownerRoot(focusBlock) !== this.ownerRoot(drag.origin)) return;
    if (focusBlock !== drag.origin) drag.crossedBlockBoundary = true;
    if (!drag.crossedBlockBoundary) return;
    event.preventDefault();
    this.queueDragSelection(drag, focus);
  };

  /** Commit the final cross-block endpoint before releasing the gesture state. */
  private readonly onMouseUp = (event: MouseEvent): void => {
    const drag = this.dragSelection;
    if (!drag) return;
    if (drag.crossedBlockBoundary) {
      const focus = caretPointFromClient(document, event.clientX, event.clientY);
      const focusBlock = focus ? this.editableBlockOf(focus.node) : null;
      if (focus && focusBlock && this.ownerRoot(focusBlock) === this.ownerRoot(drag.origin)) {
        event.preventDefault();
        setSelectionBetween(drag.anchor, focus);
      }
    }
    this.clearDragSelection();
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
   * user's place in it — untouched; a body edit reconciles incrementally. `forceRemount` is
   * for ops whose repaint provably needs whole-document context the reconciler cannot see:
   * list membership/level changes (sibling numbering shifts without sibling XML changing)
   * and border-div regrouping (HR insert, clearBorders).
   */
  private refreshAfter(block: HTMLElement, focusIndex: number, caretAtEnd = false, forceRemount = false): void {
    const band = this.region?.bandOf(block);
    if (band) {
      this.region!.refresh(this.region!.whichOf(band));
      return;
    }
    if (forceRemount) this.remount(focusIndex, caretAtEnd);
    else this.reconcile(focusIndex, caretAtEnd);
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
      columnWidth: options.columnWidth ?? "section",
      fitToWidth: options.fitToWidth ?? true,
      headerFooter: options.headerFooter ?? false,
      blockDrag: options.blockDrag ?? false,
      trackedChanges: options.trackedChanges ?? TrackedChangeMode.Accept,
      revisionAuthor: options.revisionAuthor ?? "docxodus",
      onEdit: options.onEdit,
      onMove: options.onMove,
    };
    // NOT persistAnchorIds: that setting applies to every Save on the session, so it put the
    // projector's Unid bookkeeping into the bytes the USER downloads — ~6x the file size for
    // attributes no renderer reads. Only the remount's re-render needs id stability across a
    // save/re-render hop, and it asks for that per call via SaveWithAnchorIds.
    // emitMarkdownPatch off: the editor re-renders from HTML, never from markdown patches,
    // so paying a whole-document re-projection per op would be dead weight.
    const handle = exports.DocxSessionBridge.OpenSession(bytes, JSON.stringify({
      emitMarkdownPatch: false,
      trackedChanges: trackedChangeWireName(opts.trackedChanges),
      revisionAuthor: opts.revisionAuthor,
    }));
    const editor = new DocxEditor(container, exports, handle, opts);
    editor.refreshAnchorMap();
    if (opts.headerFooter) editor.createRegion();
    const fullHtml = exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(
        bytes, opts.cssPrefix, opts.fabricateClasses, opts.paginated, opts.scale,
        opts.trackedChanges === TrackedChangeMode.RenderInline,
      ),
    );
    if (opts.paginated) editor.mountPaginated(fullHtml);
    else editor.mountHtml(fullHtml);
    editor.syncRegionToBody();
    editor.setupBlockDrag();
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
    if (typeof document !== "undefined") {
      document.removeEventListener("selectionchange", this.onSelectionChange);
      document.removeEventListener("mousedown", this.onMouseDown, true);
      document.removeEventListener("mousemove", this.onMouseMove, true);
      document.removeEventListener("mouseup", this.onMouseUp, true);
    }
    this.clearDragSelection();
    this.teardownBlockDrag();
    this.viewport.dispose();
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

  /**
   * The zoom the viewport is currently applying (1 = 100%). Below 1 the page is wider than the
   * host and has been scaled to fit rather than reflowed — the honest thing to show a user who
   * is wondering why a phone shows the whole page.
   */
  get zoom(): number {
    return this.viewport.scale;
  }

  /**
   * The live `DocxSession` handle backing this editor — the model of record.
   *
   * Surfaced because chrome around the editor (the anchor rail) reports it as engine
   * state, and reaching into the private field from a host page only worked because
   * the bundle erases TypeScript's visibility.
   */
  get sessionHandle(): number {
    return this.handle;
  }

  /**
   * Repaint from the live session after it was mutated OUTSIDE the editor's own
   * commands — a host driving {@link sessionHandle} directly (an agent pipeline,
   * `raw.replaceXml`, a batch import). Continuous mode patches incrementally from
   * the render plan (a Unid-preserving single-block mutation repaints just that
   * block); paginated mode — or anything the reconciler cannot prove — remounts.
   * The editor cannot observe external mutations, so the host owns calling this
   * once per mutation batch.
   */
  refresh(): void {
    this.assertOpen();
    this.reconcile();
  }

  /** Move one top-level body block relative to another and repaint from the live session. */
  moveBlock(
    sourceAnchorId: string,
    targetAnchorId: string,
    position: "before" | "after",
  ): boolean {
    this.assertOpen();
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.MoveBlock !== "function") return false;
    const res = this.parseEdit(
      bridge.MoveBlock(this.handle, sourceAnchorId, targetAnchorId, position),
    );
    if (!res.success) {
      // The engine's message names an OOXML construct ("a section-break paragraph"); the live
      // region is read aloud, so announce the outcome and keep the raw reason for debugging.
      this.lastMoveError = res.error?.message ?? null;
      this.announceBlockMove("This block cannot be moved there.");
      return false;
    }
    if (
      (res.created?.length ?? 0) === 0 &&
      (res.modified?.length ?? 0) === 0 &&
      (res.removed?.length ?? 0) === 0
    ) {
      this.announceBlockMove("The block is already in that position.");
      return true;
    }

    const destination = res.created?.[0] ?? res.modified?.[0];
    // Both modes reconcile. A direct move relocates the existing element, so the reconciler moves
    // the exact DOM node it already has. Review rendering additionally rewrites the SOURCE in
    // place — it becomes the move-from half — and that is visible to the diff because the plan
    // signs every unit with a content hash, so the source diffs as an in-place substitution.
    // This used to remount, which on a real charter is seconds for a one-block change.
    this.reconcile();

    const moved = destination
      ? this.bodyUnitNodes().find((el) => el.getAttribute("data-anchor") === destination.unid)
      : null;
    if (moved) {
      moved.classList.add("docx-block-move-flash");
      moved.addEventListener("animationend", () => moved.classList.remove("docx-block-move-flash"), {
        once: true,
      });
      const focusTarget = EDITABLE_TAGS.has(moved.tagName)
        ? moved
        : moved.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]');
      if (focusTarget) {
        this.activeBlock = focusTarget;
        focusTarget.focus({ preventScroll: true });
      }
    }
    this.options.onMove?.({
      sourceAnchorId,
      destinationAnchorId: destination?.id ?? sourceAnchorId,
    });
    this.announceBlockMove("Block moved.");
    return true;
  }

  /** Resolve any rendered descendant to its top-level body move unit (tables stay whole). */
  private blockUnitOf(target: EventTarget | Node | null): HTMLElement | null {
    const node = target instanceof Node ? target : null;
    const el = node?.nodeType === Node.ELEMENT_NODE
      ? node as HTMLElement
      : node?.parentElement;
    if (!el || !this.editRoot.contains(el)) return null;
    // Climb rather than scan: this runs on every pointer move over the document, and listing
    // every anchored node to find the one containing the pointer is linear in the document.
    let unit: HTMLElement | null = null;
    for (
      let candidate = el.closest<HTMLElement>("[data-anchor]");
      candidate && this.editRoot.contains(candidate);
      candidate = candidate.parentElement?.closest<HTMLElement>("[data-anchor]") ?? null
    ) {
      unit = candidate;
    }
    if (!unit || unit.closest("section.footnotes, section.endnotes")) return null;
    return unit;
  }

  private isMovableBlockUnit(unit: HTMLElement | null): unit is HTMLElement {
    if (!unit || !this.anchorIdOf(unit)) return false;
    // A source representation of an existing tracked move/deletion is evidence, not a new
    // editable block. Its accepted counterpart remains draggable.
    if (unit.matches("[class$='row-del'], [class*='row-del ']") ||
        unit.querySelector("tr[class$='row-del'], tr[class*='row-del '], del[class$='move-from'], del[class*='move-from ']")
    ) return false;
    return unit.tagName === "TABLE" || EDITABLE_TAGS.has(unit.tagName);
  }

  private currentBlockDragSource(): HTMLElement | null {
    if (this.isMovableBlockUnit(this.blockDragSource) && this.blockDragSource.isConnected)
      return this.blockDragSource;
    const fromActive = this.blockUnitOf(this.activeBlock);
    if (this.isMovableBlockUnit(fromActive)) {
      this.blockDragSource = fromActive;
      return fromActive;
    }
    return null;
  }

  private showBlockHandle(unit: HTMLElement): void {
    const handle = this.blockDragHandle;
    if (!handle || !this.isMovableBlockUnit(unit)) return;
    const changed = unit !== this.blockDragSource;
    this.blockDragSource = unit;
    // A block the engine will not move anywhere — one owning a section break, or already carrying
    // revision markup a tracked move would have to re-wrap — gets no handle rather than a handle
    // that always fails. Asking costs an engine round trip, so hovering only ever CONSUMES a
    // cached answer and schedules the ask for idle time; the drag start and the menu ask for
    // real. Merely moving the pointer across the document must not run document-scale work.
    const known = this.blockMoveTargetsFor(unit, { cachedOnly: true });
    if (known?.size === 0) {
      handle.style.display = "none";
      return;
    }
    if (changed && known === undefined) this.prefetchBlockMoveTargets(unit);
    const rect = unit.getBoundingClientRect();
    handle.style.display = "flex";
    handle.style.left = `${Math.max(4, rect.left - 32)}px`;
    handle.style.top = `${Math.max(4, rect.top + (unit.tagName === "TABLE" ? 6 : Math.max(0, (rect.height - 28) / 2)))}px`;
    const preview = blockPreviewText(unit);
    handle.setAttribute("aria-label", preview ? `Move block: ${preview}` : "Move block");
  }

  private hideBlockHandle(): void {
    if (this.blockDragging || this.blockMoveMenu?.style.display === "block") return;
    if (this.blockDragHandle) this.blockDragHandle.style.display = "none";
  }

  private positionBlockHandle(): void {
    const source = this.currentBlockDragSource();
    if (source && this.blockDragHandle?.style.display !== "none") this.showBlockHandle(source);
  }

  /** Draw the drop line on `zone`'s requested edge, or take it away when there is no target. */
  private paintDropIndicator(data: Record<string | symbol, unknown>): void {
    const indicator = this.blockDropIndicator;
    const zone = data.zone as BlockDropZone | undefined;
    if (!indicator) return;
    if (!zone) {
      this.hideDropIndicator();
      return;
    }
    const y = Math.round(this.dropEdgeY(zone, data.position === "after" ? "after" : "before")
      + this.dropZoneShift()) - 1;
    // Position before revealing, so the entry fade never plays at a stale spot.
    indicator.style.transform = `translate3d(${Math.round(zone.left)}px, ${y}px, 0)`;
    indicator.style.width = `${Math.max(24, Math.round(zone.width))}px`;
    indicator.style.display = "block";
  }

  /**
   * Where to draw the line for an insertion on `position` of `zone` — the MIDDLE of the gap to
   * the neighbour on that side, not the zone's own border-box edge. A paragraph's `w:spacing`
   * becomes a CSS margin, which sits outside the box, so drawing on the edge underlines the
   * block's last line instead of reading as a gap between two blocks. Falls back to the raw edge
   * at the ends of the flow, and degrades to the same value when blocks are contiguous.
   */
  private dropEdgeY(zone: BlockDropZone, position: "before" | "after"): number {
    const neighbour = this.dropZones[zone.index + (position === "after" ? 1 : -1)];
    // At the ends of the flow there is nothing to bisect against, so half the block's own
    // margin stands in for half the gap — the same position, one contributor instead of two.
    if (!neighbour) {
      return position === "after"
        ? zone.bottom + zone.marginAfter / 2
        : zone.top - zone.marginBefore / 2;
    }
    return position === "after"
      ? (zone.bottom + neighbour.top) / 2
      : (neighbour.bottom + zone.top) / 2;
  }

  private hideDropIndicator(): void {
    if (this.blockDropIndicator) this.blockDropIndicator.style.display = "none";
  }

  private announceBlockMove(message: string): void {
    const live = this.blockMoveLive;
    if (!live) return;
    live.textContent = "";
    this.container.ownerDocument.defaultView?.setTimeout(() => { live.textContent = message; }, 0);
  }

  /**
   * Ask the engine which anchors the current source may move next to. One call per drag source,
   * so the UI offers only drops MoveBlock will accept — a document with section breaks is
   * partitioned into regions, and drawing an indicator across one only to fail the drop was the
   * behaviour this replaces. Null (no bridge support) keeps the previous offer-everything path.
   */
  private refreshBlockMoveTargets(source: HTMLElement | null): void {
    this.blockMoveTargets = source ? this.blockMoveTargetsFor(source) ?? null : null;
  }

  /**
   * This block's legal destinations, from the memo when it is there. Returns `undefined` — not
   * `null` — for "not asked yet", so a caller can tell an unknown answer from the engine's
   * "no bridge support, offer everything" one.
   */
  private blockMoveTargetsFor(
    source: HTMLElement,
    options: { cachedOnly?: boolean } = {},
  ): Map<string, { before: boolean; after: boolean }> | null | undefined {
    const bridge = this.exports.DocxSessionBridge;
    const sourceId = this.anchorIdOf(source);
    if (!sourceId || typeof bridge.ValidMoveTargets !== "function") return null;
    if (this.blockMoveTargetCache.has(sourceId)) return this.blockMoveTargetCache.get(sourceId);
    if (options.cachedOnly) return undefined;

    let targets: Map<string, { before: boolean; after: boolean }> | null;
    try {
      const parsed = JSON.parse(bridge.ValidMoveTargets(this.handle, sourceId)) as Array<
        { anchorId: string; before: boolean; after: boolean }
      >;
      targets = new Map(parsed.map((t) => [t.anchorId, { before: t.before, after: t.after }]));
    } catch {
      targets = null;
    }
    this.blockMoveTargetCache.set(sourceId, targets);
    return targets;
  }

  /**
   * Ask for the hovered block's destinations off the interaction path, and hide the handle if
   * the answer comes back empty and that block is still the one under the pointer. The handle
   * therefore appears immediately on hover and withdraws a beat later on the rare immovable
   * block, instead of every hover paying for the query up front.
   */
  private prefetchBlockMoveTargets(source: HTMLElement): void {
    const view = this.container.ownerDocument.defaultView;
    if (!view) return;
    this.blockMoveTargetPrefetch?.();
    const run = (): void => {
      this.blockMoveTargetPrefetch = null;
      // The pointer has moved on (or the DOM was repainted) — the answer is no longer wanted.
      if (this.closed || !source.isConnected || this.blockDragSource !== source) return;
      if (this.blockMoveTargetsFor(source)?.size === 0 && this.blockDragHandle)
        this.blockDragHandle.style.display = "none";
    };
    const idle = (view as unknown as {
      requestIdleCallback?: (cb: () => void, opts?: { timeout: number }) => number;
      cancelIdleCallback?: (id: number) => void;
    });
    if (idle.requestIdleCallback && idle.cancelIdleCallback) {
      const id = idle.requestIdleCallback(run, { timeout: 500 });
      this.blockMoveTargetPrefetch = () => idle.cancelIdleCallback!(id);
    } else {
      const id = view.setTimeout(run, 0);
      this.blockMoveTargetPrefetch = () => view.clearTimeout(id);
    }
  }

  /**
   * Whether `unit` is a legal destination for the current drag source — for a SPECIFIC side when
   * one is given. A cross-block range or a section break between the blocks can make one side
   * legal and the other not, so "this target is reachable" is not enough to pick a position.
   */
  private isValidMoveTarget(unit: HTMLElement, position?: "before" | "after"): boolean {
    if (!this.blockMoveTargets) return true; // no bridge support: engine stays the only gate
    const id = this.anchorIdOf(unit);
    const sides = id ? this.blockMoveTargets.get(id) : undefined;
    if (!sides) return false;
    return position ? sides[position] : sides.before || sides.after;
  }

  /** Measure every movable block once, at drag start. See `BlockDropZone`. */
  private captureDropZones(): void {
    const view = this.container.ownerDocument.defaultView;
    this.dropZoneScroller = this.scrollContainer();
    this.dropZoneOrigin = this.scrollOffsetSum();
    this.dropZones = [];
    const boxes: HTMLElement[] = [];
    for (const unit of this.bodyUnitNodes()) {
      const anchorId = this.isMovableBlockUnit(unit) ? this.anchorIdOf(unit) : null;
      if (!anchorId) continue;
      const box = this.unitWrapperOf(unit);
      const rect = box.getBoundingClientRect();
      boxes.push(box);
      this.dropZones.push({
        unit, anchorId, index: this.dropZones.length,
        top: rect.top, bottom: rect.bottom, left: rect.left, width: rect.width,
        marginBefore: 0, marginAfter: 0,
      });
    }
    // Only the flow's two ends ever consult a margin, so only they are worth a style read.
    const ends = new Set([0, this.dropZones.length - 1].filter((i) => i >= 0 && i < boxes.length));
    for (const i of ends) {
      const style = view?.getComputedStyle(boxes[i]);
      this.dropZones[i].marginBefore = parseFloat(style?.marginTop ?? "0") || 0;
      this.dropZones[i].marginAfter = parseFloat(style?.marginBottom ?? "0") || 0;
    }
  }

  private scrollOffsetSum(): number {
    const view = this.container.ownerDocument.defaultView;
    return (view?.scrollY ?? 0) + (this.dropZoneScroller?.scrollTop ?? 0);
  }

  /** How far the measured boxes have travelled since capture, from scrolling (drag autoscroll). */
  private dropZoneShift(): number {
    return this.dropZoneOrigin - this.scrollOffsetSum();
  }

  /**
   * Where a drop at `clientY` lands, or null when nothing there is legal.
   *
   * Resolution is by VERTICAL GEOMETRY over the measured blocks, not by which element the pointer
   * is over: the drag handle floats in the page margin, so a drag straight down the gutter — the
   * natural gesture — never crosses a paragraph box, and element hit testing gave those drags no
   * indicator and no drop at all. The nearest block by vertical distance is the target; the half
   * the pointer is in picks the side, snapped to the other side when only that one is legal
   * (a section break or a cross-block range usually makes exactly one side illegal). When neither
   * side is legal — the pointer is in a region this block cannot reach — there is no drop, and
   * nothing is drawn.
   */
  private resolveDropAt(clientY: number): { zone: BlockDropZone; position: "before" | "after" } | null {
    const y = clientY - this.dropZoneShift();
    let best: BlockDropZone | null = null;
    let bestGap = Infinity;
    for (const zone of this.dropZones) {
      const gap = y < zone.top ? zone.top - y : y > zone.bottom ? y - zone.bottom : 0;
      if (gap < bestGap) {
        best = zone;
        bestGap = gap;
      }
      if (gap === 0) break; // inside this block; zones are in document order and never overlap
    }
    if (!best || best.unit === this.blockDragSource) return null;
    const preferred = y < (best.top + best.bottom) / 2 ? "before" : "after";
    if (this.isValidMoveTarget(best.unit, preferred)) return { zone: best, position: preferred };
    const other = preferred === "before" ? "after" : "before";
    return this.isValidMoveTarget(best.unit, other) ? { zone: best, position: other } : null;
  }

  private closeBlockMoveMenu(restoreFocus = false): void {
    if (!this.blockMoveMenu || !this.blockDragHandle) return;
    this.blockMoveMenu.style.display = "none";
    this.blockDragHandle.setAttribute("aria-expanded", "false");
    this.blockDragHandle.setAttribute("aria-pressed", "false");
    if (restoreFocus) this.blockDragHandle.focus({ preventScroll: true });
  }

  private openBlockMoveMenu(): void {
    const menu = this.blockMoveMenu;
    const handle = this.blockDragHandle;
    const source = this.currentBlockDragSource();
    if (!menu || !handle || !source) return;
    this.refreshBlockMoveTargets(source);
    menu.querySelectorAll<HTMLButtonElement>("button[data-move]").forEach((button) => {
      button.disabled = !this.blockMoveDestination(source, button.dataset.move ?? "");
    });
    const rect = handle.getBoundingClientRect();
    menu.style.display = "block";
    menu.style.left = `${Math.max(4, rect.right + 6)}px`;
    menu.style.top = `${Math.max(4, rect.top)}px`;
    handle.setAttribute("aria-expanded", "true");
    handle.setAttribute("aria-pressed", "true");
    menu.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus({ preventScroll: true });
  }

  /**
   * Resolve a move-menu action to the block it should land against, considering only destinations
   * the engine accepts. "Top"/"bottom" therefore mean the ends of the source's own movable region
   * — on a document with section breaks that is the section, not the document, which is what makes
   * the commands work at all there instead of always failing.
   */
  private blockMoveDestination(
    source: HTMLElement,
    action: string,
  ): { target: HTMLElement; position: "before" | "after" } | null {
    const units = this.bodyUnitNodes().filter((el) => this.isMovableBlockUnit(el));
    const index = units.indexOf(source);
    if (index < 0) return null;

    // Only pairs the engine accepts: a target can be reachable on one side and refused on the
    // other, so the SIDE is part of what makes a candidate valid.
    const position: "before" | "after" = action === "up" || action === "top" ? "before" : "after";
    const side = position === "before" ? units.slice(0, index) : units.slice(index + 1);
    const candidates = side.filter((el) => this.isValidMoveTarget(el, position));
    if (candidates.length === 0) return null;

    if (action === "up") return { target: candidates[candidates.length - 1], position };
    if (action === "down") return { target: candidates[0], position };
    if (action === "top") return { target: candidates[0], position };
    if (action === "bottom") return { target: candidates[candidates.length - 1], position };
    return null;
  }

  private runBlockMoveMenuAction(action: string): void {
    const source = this.currentBlockDragSource();
    if (!source) return;
    const destination = this.blockMoveDestination(source, action);
    this.closeBlockMoveMenu();
    if (!destination) return;
    const sourceId = this.anchorIdOf(source);
    const targetId = this.anchorIdOf(destination.target);
    if (sourceId && targetId) this.moveBlock(sourceId, targetId, destination.position);
  }

  /** The nearest ancestor that actually scrolls the document flow, for drag autoscroll. */
  private scrollContainer(): HTMLElement | null {
    const view = this.container.ownerDocument.defaultView;
    for (let el: HTMLElement | null = this.editRoot; el; el = el.parentElement) {
      const overflowY = view?.getComputedStyle(el).overflowY ?? "visible";
      if ((overflowY === "auto" || overflowY === "scroll") && el.scrollHeight > el.clientHeight)
        return el;
    }
    return null;
  }

  /** Mount one floating handle; PDD owns pointer dragging while the menu owns keyboard moves. */
  private setupBlockDrag(): void {
    if (!this.options.blockDrag || this.options.paginated || this.closed || this.blockDragHandle)
      return;
    const doc = this.container.ownerDocument;
    ensureBlockDragStyles(doc);

    const handle = doc.createElement("div");
    handle.className = "docx-block-handle";
    handle.textContent = "⠿";
    handle.tabIndex = 0;
    handle.setAttribute("role", "button");
    handle.setAttribute("contenteditable", "false");
    handle.setAttribute("aria-haspopup", "menu");
    handle.setAttribute("aria-expanded", "false");
    handle.setAttribute("aria-pressed", "false");

    const indicator = doc.createElement("div");
    indicator.className = "docx-block-drop-indicator";
    const menu = doc.createElement("div");
    menu.className = "docx-block-move-menu";
    menu.setAttribute("role", "menu");
    menu.innerHTML = `
      <button type="button" role="menuitem" data-move="up">Move up</button>
      <button type="button" role="menuitem" data-move="down">Move down</button>
      <button type="button" role="menuitem" data-move="top">Move to top</button>
      <button type="button" role="menuitem" data-move="bottom">Move to bottom</button>`;
    const live = doc.createElement("div");
    live.className = "docx-block-live-region";
    live.setAttribute("role", "status");
    live.setAttribute("aria-live", "polite");
    this.container.append(handle, indicator, menu, live);
    this.blockDragHandle = handle;
    this.blockDropIndicator = indicator;
    this.blockMoveMenu = menu;
    this.blockMoveLive = live;

    const onPointerMove = (event: PointerEvent): void => {
      // Keep the floating draggable stationary between pointerdown and native dragstart.
      // Repositioning it under the pointer during the browser's drag threshold cancels
      // drag initiation before PDD can receive dragstart.
      if (this.blockDragging || this.blockDragPointerDown) return;
      const unit = this.blockUnitOf(event.target);
      if (this.isMovableBlockUnit(unit)) this.showBlockHandle(unit);
    };
    const onFocusIn = (event: FocusEvent): void => {
      const unit = this.blockUnitOf(event.target);
      if (this.isMovableBlockUnit(unit)) this.showBlockHandle(unit);
    };
    const onPointerLeave = (event: PointerEvent): void => {
      const next = event.relatedTarget instanceof Node ? event.relatedTarget : null;
      if (next && (handle.contains(next) || menu.contains(next))) return;
      this.hideBlockHandle();
    };
    const onViewportChange = (): void => this.positionBlockHandle();
    this.container.addEventListener("pointermove", onPointerMove);
    this.container.addEventListener("focusin", onFocusIn);
    this.container.addEventListener("pointerleave", onPointerLeave);
    doc.addEventListener("scroll", onViewportChange, true);
    doc.defaultView?.addEventListener("resize", onViewportChange);
    this.blockDragCleanup.push(
      () => this.container.removeEventListener("pointermove", onPointerMove),
      () => this.container.removeEventListener("focusin", onFocusIn),
      () => this.container.removeEventListener("pointerleave", onPointerLeave),
      () => doc.removeEventListener("scroll", onViewportChange, true),
      () => doc.defaultView?.removeEventListener("resize", onViewportChange),
    );

    const onHandleClick = (): void => {
      if (menu.style.display === "block") this.closeBlockMoveMenu();
      else this.openBlockMoveMenu();
    };
    const onHandlePointerDown = (): void => { this.blockDragPointerDown = true; };
    const onPointerUp = (): void => { this.blockDragPointerDown = false; };
    const onHandleKeydown = (event: KeyboardEvent): void => {
      if (event.key !== "Enter" && event.key !== " ") return;
      event.preventDefault();
      onHandleClick();
    };
    const onMenuClick = (event: MouseEvent): void => {
      const button = (event.target as Element | null)?.closest<HTMLButtonElement>("button[data-move]");
      if (button && !button.disabled) this.runBlockMoveMenuAction(button.dataset.move ?? "");
    };
    const onMenuKeydown = (event: KeyboardEvent): void => {
      if (event.key === "Escape") { event.preventDefault(); this.closeBlockMoveMenu(true); }
    };
    handle.addEventListener("click", onHandleClick);
    handle.addEventListener("pointerdown", onHandlePointerDown);
    handle.addEventListener("keydown", onHandleKeydown);
    menu.addEventListener("click", onMenuClick);
    menu.addEventListener("keydown", onMenuKeydown);
    doc.addEventListener("pointerup", onPointerUp, true);
    doc.addEventListener("pointercancel", onPointerUp, true);
    this.blockDragCleanup.push(
      () => handle.removeEventListener("click", onHandleClick),
      () => handle.removeEventListener("pointerdown", onHandlePointerDown),
      () => handle.removeEventListener("keydown", onHandleKeydown),
      () => menu.removeEventListener("click", onMenuClick),
      () => menu.removeEventListener("keydown", onMenuKeydown),
      () => doc.removeEventListener("pointerup", onPointerUp, true),
      () => doc.removeEventListener("pointercancel", onPointerUp, true),
    );

    this.blockDragCleanup.push(draggable({
      element: handle,
      canDrag: () => !!this.currentBlockDragSource(),
      getInitialData: () => {
        const source = this.currentBlockDragSource();
        return { type: BLOCK_DRAG_TYPE, sourceAnchorId: source ? this.anchorIdOf(source) : undefined };
      },
      // The browser would otherwise ghost the 26px grip, which says nothing about what is moving.
      onGenerateDragPreview: ({ nativeSetDragImage }) => {
        const source = this.currentBlockDragSource();
        setCustomNativeDragPreview({
          nativeSetDragImage,
          getOffset: pointerOutsideOfPreview({ x: "14px", y: "10px" }),
          render: ({ container }) => {
            const chip = doc.createElement("div");
            chip.className = "docx-block-drag-preview";
            chip.textContent = (source && blockPreviewText(source)) || "Move block";
            container.appendChild(chip);
            return () => chip.remove();
          },
        });
      },
      onDragStart: () => {
        const source = this.currentBlockDragSource();
        this.blockDragging = true;
        handle.classList.add("docx-block-dragging");
        // After the preview snapshot, so the chip is not dimmed too.
        source?.classList.add("docx-block-drag-source");
        this.closeBlockMoveMenu();
        this.refreshBlockMoveTargets(source);
        this.captureDropZones();
      },
      onDrop: () => {
        this.blockDragging = false;
        this.blockDragPointerDown = false;
        this.dropZones = [];
        handle.classList.remove("docx-block-dragging");
        // Cleared by selector rather than from a captured reference: a completed move may have
        // re-rendered the block, and the node that carries the class outlives either handle on it.
        doc.querySelectorAll(".docx-block-drag-source")
          .forEach((el) => el.classList.remove("docx-block-drag-source"));
        this.hideDropIndicator();
      },
    }));
    // ONE drop target for the whole flow. Which block a drop belongs to is decided by
    // `resolveDropAt` from the pointer's vertical position, so the gutter the handle is dragged
    // down counts as being over the document — see that method.
    this.blockDragCleanup.push(dropTargetForElements({
      element: this.editRoot,
      canDrop: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      getData: ({ input }) => {
        const hit = this.resolveDropAt(input.clientY);
        return hit
          ? { type: BLOCK_DRAG_TYPE, targetAnchorId: hit.zone.anchorId, position: hit.position, zone: hit.zone }
          : { type: BLOCK_DRAG_TYPE };
      },
      onDragEnter: ({ self }) => this.paintDropIndicator(self.data),
      onDrag: ({ self }) => this.paintDropIndicator(self.data),
      onDragLeave: () => this.hideDropIndicator(),
    }));
    this.blockDragCleanup.push(monitorForElements({
      canMonitor: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      onDrop: ({ source, location }) => {
        this.hideDropIndicator();
        const drop = location.current.dropTargets[0]?.data;
        const sourceId = source.data.sourceAnchorId;
        const targetId = drop?.targetAnchorId;
        const position = drop?.position === "after" ? "after" : "before";
        if (typeof sourceId === "string" && typeof targetId === "string")
          this.moveBlock(sourceId, targetId, position);
        else
          // Released outside any valid target. Say so — a silent no-op reads as a broken drag.
          this.announceBlockMove("Move cancelled — the block was not dropped on a valid position.");
      },
    }));
    // Register on the element that actually scrolls. editRoot is the document flow — in the
    // ribbon surface the scroller is an ancestor of it, and registering the flow made Pragmatic
    // log "attached to an element that appears not to be scrollable" on every document open
    // while doing nothing.
    const scroller = this.scrollContainer();
    if (scroller) {
      this.blockDragCleanup.push(autoScrollForElements({
        element: scroller,
        canScroll: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
        getAllowedAxis: () => "vertical",
      }));
    }
    this.blockDragCleanup.push(autoScrollWindowForElements({
      canScroll: ({ source }) => source.data.type === BLOCK_DRAG_TYPE,
      getAllowedAxis: () => "vertical",
    }));
  }

  private teardownBlockDrag(): void {
    this.blockMoveTargetPrefetch?.();
    this.blockMoveTargetPrefetch = null;
    this.blockMoveTargetCache.clear();
    this.dropZones = [];
    this.dropZoneScroller = null;
    for (const cleanup of this.blockDragCleanup.splice(0)) cleanup();
    this.blockDragHandle?.remove();
    this.blockDropIndicator?.remove();
    this.blockMoveMenu?.remove();
    this.blockMoveLive?.remove();
    this.blockDragHandle = null;
    this.blockDropIndicator = null;
    this.blockMoveMenu = null;
    this.blockMoveLive = null;
    this.blockDragSource = null;
    this.blockDragging = false;
    this.blockDragPointerDown = false;
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
    const bridge = this.exports.DocxSessionBridge;
    // ListAnchors returns the same {anchorIndex} object WITHOUT the markdown payload —
    // marshaling the full projection (a couple hundred KB on a real document) made
    // this refresh the single biggest term of every incremental repaint.
    const raw =
      typeof bridge.ListAnchors === "function"
        ? bridge.ListAnchors(this.handle)
        : bridge.Project(this.handle);
    const proj = JSON.parse(raw) as {
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

  /**
   * Continuous (non-paginated) mount: inject the converter's styles + body, wire blocks.
   *
   * The body always gets its own `.docx-body-flow` wrapper — not only when bands are docked.
   * It is the sheet: the element the viewport gives page geometry to and zooms, and the one
   * the bands dock around. Without it the container would have to be both the scrolling host
   * and the scaled page, which are different boxes.
   */
  private mountHtml(fullHtml: string): void {
    const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    this.container.innerHTML = styles;
    const flow = document.createElement("div");
    flow.className = "docx-body-flow";
    flow.innerHTML = parsed.body.innerHTML;
    this.container.appendChild(flow);
    this.editRoot = flow;
    if (this.options.editable) this.wireBlocks(flow);
    this.stampPlanState();
    if (this.region) this.dockBands(flow);
    this.viewport.attach(flow, true);
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
    // Page boxes already carry the section's dimensions, so the viewport contributes only the
    // fit zoom — a full page is far wider than a phone, and clipping it is not an option.
    this.viewport.attach(pageRoot, false);
  }

  private wireBlocks(root: HTMLElement): void {
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((el) => this.wireBlock(el));
  }

  private wireBlock(el: HTMLElement): void {
    if (!EDITABLE_TAGS.has(el.tagName)) return;
    if (
      el.closest("tr[class$='row-del'], tr[class*='row-del ']") ||
      el.querySelector(":scope > del[class$='move-from'], :scope > del[class*='move-from ']")
    ) {
      el.setAttribute("contenteditable", "false");
      return;
    }
    const unid = el.getAttribute("data-anchor");
    // Only blocks the markdown projection addresses are editable via the text path. This INCLUDES
    // table-cell paragraphs (the projection indexes them), so cell text IS editable. Table-aware
    // key handling keeps structural edits safe while still providing Word-style cell navigation
    // (see onKeydown). Anything the projection does not index (unstamped content) stays read-only.
    // A band block is authoritative via its stamped `data-hf-anchor` even when the unid map
    // resolves that unid to a different part (content-addressed unids collide across parts).
    if (!unid || !this.anchorIdOf(el)) return;
    el.setAttribute("contenteditable", "true");
    ensureListMarkerSeparator(el);
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
    const html = this.renderBlockHtml(newAnchor);
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
        // The throwaway render numbers citation markers from 1 — repair in place.
        this.maybeRenumberNotes(fresh);
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
    // Inside a table cell, Backspace at the cell's start does not merge across the cell
    // boundary (mid-text Backspace still deletes normally). Enter splits the cell paragraph
    // WITHIN the same cell — the engine keeps the new w:p in the w:tc, so the grid is unchanged.
    // Tab navigation is handled explicitly below; at the final cell it uses the existing safe
    // whole-table row insertion path to match Word's add-a-row behavior.
    const inTableCell = !!el.closest("table");

    // Plain Tab / Shift+Tab moves between table cells. Outside tables it retains list
    // nest/un-nest behavior. Modified Tab chords remain available to the browser/platform.
    if (ev.key === "Tab" && !ev.ctrlKey && !ev.metaKey && !ev.altKey && !ev.isComposing) {
      if (inTableCell) {
        ev.preventDefault();
        this.navigateTableCell(el, ev.shiftKey);
        return;
      }
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

  /** Editable cells in visual document order, excluding cells from any nested table. */
  private tableCellsFor(el: HTMLElement): { cells: HTMLElement[]; index: number } | null {
    const cell = el.closest<HTMLElement>("td, th");
    const table = cell?.closest<HTMLTableElement>("table");
    if (!cell || !table) return null;
    const cells = Array.from(table.querySelectorAll<HTMLElement>("td, th"))
      .filter((candidate) => candidate.closest("table") === table);
    const index = cells.indexOf(cell);
    return index >= 0 ? { cells, index } : null;
  }

  /** First addressable paragraph in a cell, fenced against nested-table descendants. */
  private editableInCell(cell: HTMLElement): HTMLElement | null {
    return Array.from(
      cell.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    ).find((block) => block.closest("td, th") === cell) ?? null;
  }

  /** Word-style cell navigation: Shift+Tab goes back, Tab goes forward, and Tab at the
   *  final cell appends a row before entering its first cell. The first cell is a hard
   *  boundary for Shift+Tab so focus never leaks out of the editor table. */
  private navigateTableCell(el: HTMLElement, backwards: boolean): void {
    const state = this.tableCellsFor(el);
    if (!state) return;
    const adjacent = state.cells[state.index + (backwards ? -1 : 1)];
    if (adjacent) {
      const target = this.editableInCell(adjacent);
      if (target) placeCaretAtOffset(target, 0);
      return;
    }
    if (backwards) return;

    // The current block may contain uncommitted typing. insertTableRow flushes it before the
    // structural edit, then reconcile/remount restores focus by block index. Resolve the new
    // adjacent cell from that fresh DOM rather than retaining a node from the replaced table.
    this.activeBlock = el;
    this.insertTableRow("below");
    const fresh = this.activeBlock ? this.tableCellsFor(this.activeBlock) : null;
    if (!fresh) return;
    const target = fresh.cells[fresh.index + 1];
    const block = target ? this.editableInCell(target) : null;
    if (block) placeCaretAtOffset(block, 0);
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

    const [firstEl, secondEl] = this.renderTwo(first.id, second.id);
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
    const html = this.renderBlockHtml(anchorId);
    if (html.charCodeAt(0) === 0x7b /* error object */) return null;
    return new DOMParser().parseFromString(html, "text/html").body.firstElementChild as HTMLElement | null;
  }

  private get renderTrackedChanges(): boolean {
    return this.options.trackedChanges === TrackedChangeMode.RenderInline;
  }

  private renderBlockHtml(anchorId: string): string {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderBlockHtmlForReview === "function") {
      return bridge.RenderBlockHtmlForReview(
        this.handle, anchorId, this.options.cssPrefix, this.options.fabricateClasses,
        this.renderTrackedChanges,
      );
    }
    return bridge.RenderBlockHtml(
      this.handle, anchorId, this.options.cssPrefix, this.options.fabricateClasses,
    );
  }

  private renderPlanJson(): string {
    const bridge = this.exports.DocxSessionBridge;
    return typeof bridge.ListRenderedBlocks === "function"
      ? bridge.ListRenderedBlocks(this.handle, this.renderTrackedChanges)
      : bridge.ListBlocks!(this.handle);
  }

  /** Render two blocks in ONE batched bridge call when the bundle carries RenderBlocksHtml —
   *  the per-render shell/converter setup is paid once instead of twice, which matters on the
   *  Enter path (split renders both halves synchronously under the keystroke). Falls back to
   *  two per-block renders on older bundles. Output is renderInto-identical per block (both
   *  routes share the same extraction). */
  private renderTwo(a: string, b: string): [HTMLElement | null, HTMLElement | null] {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.RenderBlocksHtml === "function") {
      try {
        const idsJson = JSON.stringify([a, b]);
        const json = typeof bridge.RenderBlocksHtmlForReview === "function"
          ? bridge.RenderBlocksHtmlForReview(
              this.handle, idsJson, this.options.cssPrefix, this.options.fabricateClasses,
              this.renderTrackedChanges,
            )
          : bridge.RenderBlocksHtml(
            this.handle,
            idsJson,
            this.options.cssPrefix,
            this.options.fabricateClasses,
          );
        const map = JSON.parse(json) as Record<string, string | null> & { error?: string };
        if (!map.error) {
          const parse = (h: string | null | undefined): HTMLElement | null =>
            h
              ? (new DOMParser().parseFromString(h, "text/html").body
                  .firstElementChild as HTMLElement | null)
              : null;
          const fa = parse(map[a]);
          const fb = parse(map[b]);
          if (fa && fb) return [fa, fb];
        }
      } catch {
        /* fall through to per-block renders */
      }
    }
    return [this.renderInto(a), this.renderInto(b)];
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
      const result = JSON.parse(json) as EditResultLite;
      if (result.success) this.invalidateBlockMoveTargets();
      return result;
    } catch {
      return { success: false };
    }
  }

  /**
   * Drop the memoized `ValidMoveTargets` answers. Which blocks a block may move next to is a
   * fact about the DOCUMENT, so it survives hovering but not editing — and the two places a
   * document changes are `parseEdit` (every mutation that returns an `EditResult`) and
   * undo/redo, which return a bare boolean and so cannot go through it.
   */
  private invalidateBlockMoveTargets(): void {
    this.blockMoveTargetCache.clear();
  }

  // ─── M5: formatting commands (ribbon) ────────────────────────────────

  // ─── Multi-block selection helpers (format a whole stack of paragraphs at once) ──────

  /**
   * Restore a cross-block selection after a native toolbar control took focus. The bookmark uses
   * stable anchor ids and content offsets, so it also survives incremental block swaps.
   */
  private restoreCrossBlockSelection(): boolean {
    const bookmark = this.lastCrossBlockSelection;
    const active = this.activeBlock;
    if (!bookmark || !active) return false;
    const all = Array.from(
      this.ownerRoot(active).querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
    );
    const first = all.find((block) => block.getAttribute("data-anchor") === bookmark.startUnid);
    const last = all.find((block) => block.getAttribute("data-anchor") === bookmark.endUnid);
    if (!first || !last || first === last || all.indexOf(first) > all.indexOf(last)) return false;
    const start = contentPositionIn(first, bookmark.startOffset);
    const end = contentPositionIn(last, bookmark.endOffset);
    return setSelectionBetween(start, end);
  }

  /** Editable blocks the current selection covers, in document order. Uses Range.comparePoint
   *  (robust to a selection boundary that normalized onto a wrapper element rather than a block
   *  or text node — Range.intersectsNode misses the end block at a `(block, childCount)` boundary).
   *  A collapsed or single-block selection yields just the active block. */
  private selectedBlocks(): HTMLElement[] {
    let sel = typeof window !== "undefined" ? window.getSelection() : null;
    if ((!sel || sel.rangeCount === 0 || sel.isCollapsed) && this.lastCrossBlockSelection) {
      this.restoreCrossBlockSelection();
      sel = typeof window !== "undefined" ? window.getSelection() : null;
    }
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
    // Remount from the active block's index re-renders the new rule whether it landed just
    // above (at idx) or just below (at idx+1) the active block. Forced: a rule is a bordered
    // paragraph, and border-div grouping is whole-document render context.
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
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
    // context (sibling numbers shift without sibling XML changing), keeping the caret in place.
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
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
    // (sibling numbers shift without sibling XML changing).
    this.refreshAfter(block, idx, false, /* forceRemount */ true);
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

  /**
   * Cite a new footnote from the caret position in the active body block. The note definition is
   * created (writing the whole Word scaffold — part, reserved separator notes, settings
   * declaration, styles — on a document that has none yet) and its body renders as ordinary
   * editable `data-anchor` blocks in the notes section, so editing it afterwards needs no new op.
   *
   * Body blocks only: Word disallows a note reference inside a header/footer story or inside
   * another note, and the session rejects those with `AnchorWrongKind`. Reconciliation creates
   * the first notes section when needed and renumbers every affected citation in place.
   */
  insertFootnote(markdown = "New footnote."): void {
    this.insertNote("footnote", markdown);
  }

  /** Cite a new endnote from the caret — see {@link insertFootnote}; writes the endnotes part. */
  insertEndnote(markdown = "New endnote."): void {
    this.insertNote("endnote", markdown);
  }

  private insertNote(kind: "footnote" | "endnote", markdown: string): void {
    const block = this.activeBlock;
    if (this.closed || !block) return;
    // A note reference is legal only in the main story: not in a header/footer band, and not
    // inside an existing note's body (both render as editable blocks here).
    if (this.isBandBlock(block) || block.closest(".footnotes, .endnotes")) return;
    let fullId = this.anchorIdOf(block);
    if (!fullId) return;
    const idx = this.blockIndex(block);
    // Offset first: syncBlock re-renders the block and would drop the live selection.
    const raw = caretOffsetIn(block);
    fullId = this.syncBlock(block, fullId);
    const offset = trimmedSplitOffset(block, raw ?? (block.textContent ?? "").length);
    const bridge = this.exports.DocxSessionBridge;
    const call = kind === "footnote" ? bridge.InsertFootnote : bridge.InsertEndnote;
    if (!call) return; // bridge predates note authoring
    const res = this.parseEdit(call.call(bridge, this.handle, fullId, offset, markdown));
    if (!res.success) return;
    this.reconcile(idx, false);
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
    if (this.affectsList(res) || op.clearBorders) { this.refreshAfter(block, idx, false, true); return; }
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
    if (this.affectsList(res)) { this.refreshAfter(block, idx, false, true); return; }
    this.swapBlock(block, unid, res.modified?.[0])?.focus();
  }

  /** Undo the last edit (incremental repaint; falls back to a full re-render). */
  undo(): void {
    if (this.closed) return;
    if (!this.exports.DocxSessionBridge.Undo(this.handle)) return;
    this.invalidateBlockMoveTargets();
    this.reconcile();
  }

  /** Redo the last undone edit (incremental repaint; falls back to a full re-render). */
  redo(): void {
    if (this.closed) return;
    if (!this.exports.DocxSessionBridge.Redo(this.handle)) return;
    this.invalidateBlockMoveTargets();
    this.reconcile();
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

  /**
   * Set the page numbering of the section the bands describe (`w:pgNumType`) — Word's *Format Page
   * Numbers…*: `start` restarts numbering at that number, `format` chooses `1, 2, 3` vs
   * `i, ii, iii` etc. Omitted fields are left unchanged. Requires the header/footer region
   * (`{ headerFooter: true }`); a no-op otherwise.
   *
   * Inserted page-number fields are plain, so they render through this. The editor's own view still
   * shows each field's cached result — Word recomputes on open — but `{ paginated: true }`
   * substitutes the real per-page number and so reflects the change immediately.
   */
  setPageNumbering(op: { start?: number; format?: NumberFormat }): void {
    this.assertOpen();
    this.region?.setPageNumbering(op);
  }

  /** Remove the section's page-numbering start/format: it reverts to continuing the previous
   *  section's numbering in Word's default `1, 2, 3`. */
  clearPageNumbering(): void {
    this.assertOpen();
    this.region?.clearPageNumbering();
  }

  /** This section's page numbering as the document currently states it — `{}` when the section
   *  sets neither (it continues the previous section in the default format). */
  pageNumbering(): { start?: number; format?: NumberFormat } {
    return this.region?.pageNumbering() ?? {};
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
    // The throwaway render numbers citation markers from 1 — repair in place.
    this.maybeRenumberNotes(fresh);
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
    if (typeof bridge.RenderHtmlForReview === "function") {
      const html = bridge.RenderHtmlForReview(
        this.handle,
        this.options.cssPrefix,
        this.options.fabricateClasses,
        this.options.paginated,
        this.options.scale,
        this.renderTrackedChanges,
      );
      if (html.charCodeAt(0) !== 0x7b) return html;
    }
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
    // Fallback only (no RenderHtml on this bridge, or it errored). These bytes are re-rendered and
    // discarded, and the re-render has to resolve to the SAME anchors the live session holds — a
    // content change re-derives a block's content-hashed unid, which would leave it unwired. So ask
    // for the Unid-bearing save here, and here only; DocxEditor.save() stays clean.
    const bytes =
      typeof bridge.SaveWithAnchorIds === "function"
        ? bridge.SaveWithAnchorIds(this.handle)
        : bridge.Save(this.handle);
    return this.exports.DocumentConverter.ConvertDocxToHtmlComplete(
      ...completeArgs(
        bytes, this.options.cssPrefix, this.options.fabricateClasses,
        this.options.paginated, this.options.scale, this.renderTrackedChanges,
      ),
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

  // ─── Incremental structural reconcile ─────────────────────────────────
  //
  // After a structural op (insert table/row/col, footnote, delete block, undo/redo)
  // the DOM is patched from a unit-sequence diff against the session's render plan
  // instead of remounting the whole document (~3 s of full-document conversion on a
  // 350-block file). Full remount remains the universal FALLBACK: any ambiguity,
  // unsupported bridge, paginated mode, list-membership change, or thrown error
  // lands there — correctness never depends on the diff being right.

  /** True when the bridge carries the reconcile trio and the mode allows patching. */
  private canReconcile(): boolean {
    const b = this.exports.DocxSessionBridge;
    return (
      !this.options.paginated &&
      typeof b.ListBlocks === "function" &&
      typeof b.RenderBlocksHtml === "function" &&
      typeof b.ListNotes === "function"
    );
  }

  /** The body's top-level unit nodes in document order: `[data-anchor]` elements not
   *  nested in another unit (cell paragraphs collapse into their table) and not in the
   *  notes sections. */
  private bodyUnitNodes(): HTMLElement[] {
    const all = Array.from(this.editRoot.querySelectorAll<HTMLElement>("[data-anchor]"));
    return all.filter((el) => {
      if (el.closest("section.footnotes, section.endnotes")) return false;
      const ancestor = el.parentElement?.closest("[data-anchor]");
      return !(ancestor && this.editRoot.contains(ancestor));
    });
  }

  /** The DOM diff token for a body unit node (see editor-reconcile.tokenOf). */
  private static domTokenOf(el: HTMLElement): string {
    const unid = el.getAttribute("data-anchor") ?? "";
    const sig = el.getAttribute("data-render-sig");
    return sig ? `${unid}|${sig}` : unid;
  }

  /** The kind a body unit node would have in the plan (only 'li'/'tbl' matter to the
   *  remount guard). */
  private static domKindOf(el: HTMLElement): string {
    if (el.tagName === "TABLE") return "tbl";
    return el.querySelector(":scope > [data-list-marker]") ? "li" : "p";
  }

  private static listMarkerText(el: Element | null): string | null {
    const m = el?.querySelector(":scope > [data-list-marker]");
    return m ? m.textContent : null;
  }

  /**
   * Incrementally patch the DOM from the session's render plan; falls back to
   * {@link remount} whenever it cannot prove the patch correct. Same focus contract
   * as remount.
   */
  private reconcile(focusIndex = -1, caretAtEnd = false): void {
    if (!this.canReconcile()) {
      this.remount(focusIndex, caretAtEnd);
      return;
    }
    try {
      if (!this.reconcileCore()) {
        this.remount(focusIndex, caretAtEnd);
        return;
      }
    } catch (err) {
      this.lastReconcileFallback = `threw: ${err instanceof Error ? err.message : String(err)}`;
      this.remount(focusIndex, caretAtEnd);
      return;
    }
    if (focusIndex >= 0) {
      const blocks = this.editableList();
      const target = blocks[Math.min(focusIndex, blocks.length - 1)];
      if (target) {
        this.activeBlock = target;
        placeCaretAtOffset(target, caretAtEnd ? (target.textContent ?? "").length : 0);
      }
    }
    this.syncRegionToBody(this.activeBlock ?? undefined);
  }

  /** The patch itself. Returns false to request the remount fallback. */
  private reconcileCore(): boolean {
    const bridge = this.exports.DocxSessionBridge;
    // Refresh the unid → anchor map FIRST: wiring freshly rendered nodes (wireBlock)
    // resolves through it, and the map must reflect the post-op session.
    this.refreshAnchorMap();
    const plan = JSON.parse(this.renderPlanJson()) as RenderPlan & { error?: string };
    if (plan.error) return this.bail(`plan error: ${plan.error}`);

    const oldNodes = this.bodyUnitNodes();
    const oldTokens = oldNodes.map(DocxEditor.domTokenOf);
    const oldKinds = oldNodes.map(DocxEditor.domKindOf);
    const bodyDiff = diffUnits(oldTokens, plan.body);
    if (needsRemount(bodyDiff, plan.body, oldKinds)) {
      // Name the shape, not just the verdict: "churn" and "a list item moved in or out" are very
      // different findings when a repaint unexpectedly costs a whole-document render.
      return this.bail(
        `needsRemount (li change or churn): +${bodyDiff.added.length} -${bodyDiff.removed.length} ` +
          `~${bodyDiff.substituted.length} moved=${bodyDiff.moved.length} of ${plan.body.length}`,
      );
    }

    const fnState = this.notesDiff("footnotes", plan.footnotes);
    const enState = this.notesDiff("endnotes", plan.endnotes);
    if (fnState === null || enState === null) return this.bail("notes container unstampable/missing");

    // One batch render for everything that needs fresh HTML.
    const movedBodyNew = new Set(bodyDiff.moved.map((m) => m.newIndex));
    const addedBodyIds = bodyDiff.added
      .filter((j) => !movedBodyNew.has(j))
      .map((j) => plan.body[j].id);
    const addedNoteIds = fnState.diff.added
      .map((j) => plan.footnotes[j].id)
      .concat(enState.diff.added.map((j) => plan.endnotes[j].id));
    const allIds = addedBodyIds.concat(addedNoteIds);
    let rendered: Record<string, string | null> = {};
    if (allIds.length > 0) {
      const idsJson = JSON.stringify(allIds);
      const renderedJson = typeof bridge.RenderBlocksHtmlForReview === "function"
        ? bridge.RenderBlocksHtmlForReview(
            this.handle, idsJson, this.options.cssPrefix, this.options.fabricateClasses,
            this.renderTrackedChanges,
          )
        : bridge.RenderBlocksHtml!(
            this.handle, idsJson, this.options.cssPrefix, this.options.fabricateClasses,
          );
      rendered = JSON.parse(renderedJson);
      if ((rendered as { error?: string }).error) return this.bail(`render error: ${(rendered as { error?: string }).error}`);
      for (const id of allIds) if (!rendered[id]) return this.bail(`unrenderable: ${id}`);
    }

    // A substituted list item may only swap in place if its rendered marker matches
    // the old node's — a marker change (level/membership/numbering) means sibling
    // numbers shifted too, which only a remount repaints.
    const parse = (html: string): HTMLElement | null => {
      const el = new DOMParser().parseFromString(html, "text/html").body
        .firstElementChild as HTMLElement | null;
      // Per-block converter output carries the XHTML xmlns; a full render only has it
      // on the document root, so drop it to keep reconciled DOM ≡ remounted DOM.
      el?.removeAttribute("xmlns");
      return el;
    };
    const freshBody = new Map<number, HTMLElement>();
    for (const j of bodyDiff.added.filter((j) => !movedBodyNew.has(j))) {
      const el = parse(rendered[plan.body[j].id]!);
      if (!el) return this.bail(`unparseable render: ${plan.body[j].id}`);
      freshBody.set(j, el);
    }
    for (const { oldIndex, newIndex } of bodyDiff.substituted) {
      const freshRoot = freshBody.get(newIndex);
      const oldMarker = DocxEditor.listMarkerText(oldNodes[oldIndex]);
      const newMarker = DocxEditor.listMarkerText(
        freshRoot ? DocxEditor.anchorElOf(freshRoot) : null,
      );
      if (oldMarker !== newMarker) return this.bail("substituted li marker drift");
    }

    const bodyApply = this.applyBodyDiff(oldNodes, plan.body, bodyDiff, freshBody);
    if (bodyApply !== true) return this.bail(`applyBodyDiff bail: ${bodyApply}`);
    this.applyNotesDiff("footnotes", plan.footnotes, fnState, rendered);
    this.applyNotesDiff("endnotes", plan.endnotes, enState, rendered);

    // Note chrome (marker sup text / hrefs / li values) is position-derived and NOT
    // covered by the unit diff — renumber whenever notes changed or any fresh body
    // node carries a citation marker.
    const freshHasMarker = [...freshBody.values()].some((el) =>
      el.querySelector("a.footnote-ref, a.endnote-ref"),
    );
    if (fnState.diff.added.length + fnState.diff.removed.length > 0 || freshHasMarker)
      this.renumberNoteChrome("footnote");
    if (enState.diff.added.length + enState.diff.removed.length > 0 || freshHasMarker)
      this.renumberNoteChrome("endnote");
    this.lastReconcileFallback = null;
    return true;
  }

  private bail(reason: string): false {
    this.lastReconcileFallback = reason;
    return false;
  }

  /** The generated single-child wrapper chain around a unit node (a table's alignment
   *  <div>). Climbs while the parent is an anchor-less DIV whose ONLY element child is
   *  the current node — never a section div (multi-child) or the edit root. */
  private unitWrapperOf(el: HTMLElement): HTMLElement {
    let n: HTMLElement = el;
    while (
      n.parentElement &&
      n.parentElement !== this.editRoot &&
      n.parentElement.tagName === "DIV" &&
      !n.parentElement.hasAttribute("data-section-index") &&
      !n.parentElement.hasAttribute("data-anchor") &&
      n.parentElement.childElementCount === 1
    ) {
      n = n.parentElement;
    }
    return n;
  }

  /** The `[data-anchor]` element of a fresh render root (the root itself for a leaf
   *  block, its descendant for a wrapper-shaped render like a table's align div). */
  private static anchorElOf(root: HTMLElement): HTMLElement | null {
    return root.hasAttribute("data-anchor")
      ? root
      : root.querySelector<HTMLElement>("[data-anchor]");
  }

  /** Insert/remove/swap body unit nodes per the diff. Returns `true` on success or a
   *  bail-reason string (parent ambiguity, order violation, wrapper semantics) — the
   *  session is already correct, so bailing just means a full repaint. */
  private applyBodyDiff(
    oldNodes: HTMLElement[],
    units: RenderUnit[],
    diff: UnitDiff,
    fresh: Map<number, HTMLElement>,
  ): true | string {
    // In-place substitutions first: replace at WRAPPER level so a table swaps with its
    // alignment div. A LEAF render replacing a wrapped node would break the wrapper's
    // semantics (border <div> grouping) — that is remount territory.
    const subOldByNew = new Map(diff.substituted.map((s) => [s.newIndex, s.oldIndex]));
    for (const [nj, oi] of subOldByNew) {
      const freshRoot = fresh.get(nj)!;
      const oldWrapper = this.unitWrapperOf(oldNodes[oi]);
      if (!freshRoot.hasAttribute("data-anchor")) {
        oldWrapper.replaceWith(freshRoot); // wrapper-shaped render (table) ⇄ wrapper
      } else if (oldWrapper === oldNodes[oi]) {
        oldNodes[oi].replaceWith(freshRoot);
      } else if (
        !oldNodes[oi].hasAttribute("data-render-sig") &&
        unidOf(units[nj].id) === oldNodes[oi].getAttribute("data-anchor")
      ) {
        // Same block, leaf render, wrapped slot, and the old node carries NO signature:
        // a sig-less node is one a SINGLE-BLOCK swap placed (mount and reconcile both
        // stamp signatures), and those swaps never change border grouping — every
        // border-changing op forces a full remount. The wrapper is therefore still
        // valid; swap the leaf inside it, exactly as the single-block path itself does
        // for a bordered paragraph. A STAMPED wrapped slot whose content changed stays
        // remount territory — the change could be the border itself.
        oldNodes[oi].replaceWith(freshRoot);
      } else {
        // border-div semantics, remount
        return `leaf render into wrapped slot (unit ${units[nj].id.slice(0, 24)})`;
      }
      this.wireUnit(freshRoot, units[nj]);
    }

    const movedOldByNew = new Map(diff.moved.map((m) => [m.newIndex, m.oldIndex]));
    const movedNew = new Set(diff.moved.map((m) => m.newIndex));
    const movedOld = new Set(diff.moved.map((m) => m.oldIndex));
    const pureAdded = diff.added.filter((j) => !subOldByNew.has(j) && !movedNew.has(j));
    const pureRemoved = diff.removed.filter(
      (i) => !diff.substituted.some((s) => s.oldIndex === i) && !movedOld.has(i),
    );
    // DOMParser roots report isConnected=true because they belong to their detached
    // HTMLDocument. Track only fresh additions actually inserted into the live editor;
    // otherwise a later not-yet-applied sibling looks like a neighbor in another parent.
    const appliedFresh = new Set<number>();
    const nodeAt = (j: number): HTMLElement | null => {
      const oi = diff.keep.get(j);
      if (oi !== undefined) return oldNodes[oi];
      const movedOi = movedOldByNew.get(j);
      if (movedOi !== undefined) return oldNodes[movedOi];
      if (subOldByNew.has(j)) return fresh.get(j)!;
      const f = fresh.get(j);
      return f && appliedFresh.has(j) ? f : null;
    };

    // Preserve the established incremental insertion path: body render output may
    // legitimately use several wrapper parents (for example border groups), so it is
    // not valid to normalize the whole body against one common parent.
    for (const j of pureAdded) {
      const el = fresh.get(j)!;
      let prev: HTMLElement | null = null;
      for (let k = j - 1; k >= 0 && !prev; k--) prev = nodeAt(k);
      let next: HTMLElement | null = null;
      for (let k = j + 1; k < units.length && !next; k++) next = nodeAt(k);
      const prevW = prev ? this.unitWrapperOf(prev) : null;
      const nextW = next ? this.unitWrapperOf(next) : null;
      if (prevW && nextW && prevW.parentElement !== nextW.parentElement)
        return "insert neighbors in different parents";
      if (prevW) prevW.after(el);
      else if (nextW) nextW.before(el);
      else return "insert with no anchored neighbor";
      appliedFresh.add(j);
      this.wireUnit(el, units[j]);
    }

    // Pure removals take their now-empty generated wrappers with them. Exact-token
    // move sources remain live even though the LCS also reports them as removed.
    for (const i of pureRemoved) {
      this.unitWrapperOf(oldNodes[i]).remove();
    }

    // A block move changes one exact-token unit's slot. Move only that existing
    // wrapper beside its final neighbor; this preserves identity without disturbing
    // unrelated nodes or assuming every rendered unit shares one DOM parent.
    for (const { newIndex } of [...diff.moved].sort((a, b) => a.newIndex - b.newIndex)) {
      const moving = nodeAt(newIndex);
      if (!moving?.isConnected) return "move source is not connected";
      const movingW = this.unitWrapperOf(moving);
      let prevW: HTMLElement | null = null;
      for (let k = newIndex - 1; k >= 0 && !prevW; k--) {
        const n = nodeAt(k);
        if (n?.isConnected) prevW = this.unitWrapperOf(n);
      }
      let nextW: HTMLElement | null = null;
      for (let k = newIndex + 1; k < units.length && !nextW; k++) {
        const n = nodeAt(k);
        if (n?.isConnected) nextW = this.unitWrapperOf(n);
      }
      if (prevW && prevW !== movingW && prevW.parentElement === movingW.parentElement) {
        prevW.after(movingW);
      } else if (nextW && nextW !== movingW && nextW.parentElement === movingW.parentElement) {
        nextW.before(movingW);
      } else {
        return "move neighbors in different parents";
      }
    }
    return true;
  }

  /** Wire a freshly rendered unit root (and its nested blocks) and stamp the unit's
   *  content signature on its `[data-anchor]` element — the element the next
   *  reconcile's DOM walk reads tokens from. */
  private wireUnit(root: HTMLElement, unit: RenderUnit): void {
    const anchorEl = DocxEditor.anchorElOf(root);
    if (anchorEl && unit.sig) anchorEl.setAttribute("data-render-sig", unit.sig);
    if (anchorEl) this.wireBlock(anchorEl);
    root.querySelectorAll<HTMLElement>("[data-anchor]").forEach((b) => this.wireBlock(b));
  }

  /** Old-sequence diff state for one notes section. `null` requests remount when an
   *  existing list is not stampable; an absent list is the valid first-note case. */
  private notesDiff(
    sectionClass: "footnotes" | "endnotes",
    units: RenderUnit[],
  ): { lis: HTMLElement[]; diff: UnitDiff } | null {
    const ol = this.editRoot.querySelector<HTMLElement>(`section.${sectionClass} > ol`);
    const lis = ol ? (Array.from(ol.children).filter((c) => c.tagName === "LI") as HTMLElement[]) : [];
    if (units.length === 0 && lis.length === 0) return { lis, diff: diffUnits([], []) };
    // Full render omits empty note sections. Treat the first note as an insertion into
    // an empty list; applyNotesDiff creates the converter-equivalent section + list.
    if (!ol) return { lis, diff: diffUnits([], units) };
    const tokens: string[] = [];
    for (const li of lis) {
      const unid = li.getAttribute("data-note-anchor");
      if (!unid) return null; // unstamped DOM (older mount) — remount restamps
      const sig = li.getAttribute("data-render-sig");
      tokens.push(sig ? `${unid}|${sig}` : unid);
    }
    return { lis, diff: diffUnits(tokens, units) };
  }

  /** Apply a notes-section diff: rebuild the `<ol>`'s li list, preserving kept nodes. */
  private applyNotesDiff(
    sectionClass: "footnotes" | "endnotes",
    units: RenderUnit[],
    state: { lis: HTMLElement[]; diff: UnitDiff },
    rendered: Record<string, string | null>,
  ): void {
    if (state.diff.added.length === 0 && state.diff.removed.length === 0) return;
    // Removing the LAST note removes the whole section — a full render emits no
    // section for a document without notes, and equivalence with remount is the pin.
    if (units.length === 0) {
      this.editRoot.querySelector(`section.${sectionClass}`)?.remove();
      return;
    }
    let section = this.editRoot.querySelector<HTMLElement>(`:scope > section.${sectionClass}`);
    let ol = section?.querySelector<HTMLOListElement>(`:scope > ol`) ?? null;
    if (!ol) {
      if (!section) {
        section = document.createElement("section");
        section.className = sectionClass;
        const following = sectionClass === "footnotes"
          ? this.editRoot.querySelector<HTMLElement>(`:scope > section.endnotes`)
          : null;
        if (following) following.before(section);
        else this.editRoot.append(section);
      }
      if (!section.querySelector(":scope > hr")) section.prepend(document.createElement("hr"));
      ol = document.createElement("ol");
      section.append(ol);
    }
    const prefix = sectionClass === "footnotes" ? "fn" : "en";
    const nodes: HTMLElement[] = [];
    for (let j = 0; j < units.length; j++) {
      const oi = state.diff.keep.get(j);
      if (oi !== undefined) {
        nodes.push(state.lis[oi]);
        continue;
      }
      nodes.push(this.buildNoteLi(prefix, units[j], rendered[units[j].id]!));
    }
    ol.replaceChildren(...nodes);
  }

  /** Build a notes-section `<li>` for a freshly rendered note — replicating the
   *  converter's chrome (id/value are re-stamped by the renumber pass; the backref
   *  goes inside the last paragraph, matching RenderFootnoteItem). */
  private buildNoteLi(prefix: "fn" | "en", unit: RenderUnit, html: string): HTMLElement {
    const li = document.createElement("li");
    li.setAttribute("data-note-anchor", unidOf(unit.id));
    if (unit.sig) li.setAttribute("data-render-sig", unit.sig);
    li.innerHTML = html;
    const paras = li.querySelectorAll<HTMLElement>(":scope > p");
    const last = paras[paras.length - 1];
    if (last) {
      const backref = document.createElement("a");
      backref.setAttribute("class", `${prefix}-backref`);
      backref.setAttribute("contenteditable", "false");
      backref.textContent = "↩";
      last.append(" ", backref);
    }
    li.querySelectorAll<HTMLElement>("[data-anchor]").forEach((b) => this.wireBlock(b));
    return li;
  }

  /**
   * Rewrite position-derived note chrome from the session's citation-ordered note
   * list: the k-th marker in document order IS note k (ids ascend in reference
   * order), so marker sup text, hrefs/ids, li ids/values and backref hrefs are all
   * re-derived positionally. Pure attribute/text patching of generated chrome.
   */
  private renumberNoteChrome(kind: "footnote" | "endnote"): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ListNotes !== "function") return;
    const prefix = kind === "footnote" ? "fn" : "en";
    let notes: Array<{ id: string; defAnchorId: string; ordinal: number }>;
    try {
      notes = JSON.parse(bridge.ListNotes(this.handle, kind === "endnote"));
    } catch {
      return;
    }
    if (!Array.isArray(notes)) return;
    const markers = Array.from(
      this.editRoot.querySelectorAll<HTMLElement>(`a.${kind}-ref`),
    ).filter((a) => !a.closest("section.footnotes, section.endnotes"));
    markers.forEach((a, k) => {
      const n = notes[k];
      if (!n) return;
      a.setAttribute("href", `#${prefix}-${n.id}`);
      a.id = `${prefix}-ref-${n.id}`;
      if (kind === "footnote") a.setAttribute("data-footnote-id", n.id);
      const sup = a.querySelector("sup");
      if (sup) sup.textContent = String(n.ordinal);
    });
    // Match list items by their stamped note anchor, NOT position: the section can
    // hold rendered-but-never-cited notes (Word's continuationNotice) that ListNotes
    // — a citation walk — does not list; positional pairing would relabel them.
    const byUnid = new Map(notes.map((n) => [unidOf(n.defAnchorId), n]));
    const lis = this.editRoot.querySelectorAll<HTMLElement>(`section.${kind}s > ol > li`);
    lis.forEach((li) => {
      const unid = li.getAttribute("data-note-anchor");
      const n = unid ? byUnid.get(unid) : undefined;
      if (!n) return;
      li.id = `${prefix}-${n.id}`;
      li.setAttribute("value", String(n.ordinal));
      li.querySelectorAll<HTMLElement>(`a.${prefix}-backref`).forEach((b) =>
        b.setAttribute("href", `#${prefix}-ref-${n.id}`),
      );
    });
  }

  /** After an incremental block swap, stale marker chrome in the swapped node (the
   *  throwaway render numbers citations from 1) is repaired in place. */
  private maybeRenumberNotes(fresh: HTMLElement): void {
    if (fresh.querySelector("a.footnote-ref")) this.renumberNoteChrome("footnote");
    if (fresh.querySelector("a.endnote-ref")) this.renumberNoteChrome("endnote");
  }

  /** Stamp the DOM state the reconciler diffs against: container signatures on body
   *  tables and `data-note-anchor` + signature on notes-section items. Called after
   *  every full mount; reconcile stamps its own insertions. */
  private stampPlanState(): void {
    const bridge = this.exports.DocxSessionBridge;
    if (typeof bridge.ListBlocks !== "function") return;
    try {
      const plan = JSON.parse(this.renderPlanJson()) as RenderPlan & { error?: string };
      if (plan.error) return;
      // Positional pairing: a fresh full mount renders exactly the plan's units in
      // order (verified invariant). On any mismatch, leave unstamped — an unstamped
      // unit just diffs as changed and re-renders once.
      const nodes = this.bodyUnitNodes();
      if (nodes.length === plan.body.length) {
        nodes.forEach((el, k) => {
          if (plan.body[k].sig && el.getAttribute("data-anchor") === unidOf(plan.body[k].id))
            el.setAttribute("data-render-sig", plan.body[k].sig!);
        });
      }
      const stampNotes = (sectionClass: string, units: RenderUnit[]): void => {
        const lis = this.editRoot.querySelectorAll<HTMLElement>(`section.${sectionClass} > ol > li`);
        if (lis.length !== units.length) return; // inconsistent — leave unstamped (reconcile will remount)
        lis.forEach((li, k) => {
          li.setAttribute("data-note-anchor", unidOf(units[k].id));
          if (units[k].sig) li.setAttribute("data-render-sig", units[k].sig!);
        });
      };
      stampNotes("footnotes", plan.footnotes);
      stampNotes("endnotes", plan.endnotes);
    } catch {
      /* stamping is best-effort; unstamped DOM just falls back to remount */
    }
  }

  /**
   * Full re-render from current session state (after undo/redo, and after list edits where
   * single-block rendering can't compute numbering). Optionally focus the editable block at
   * `focusIndex` (caret at start, or end if `caretAtEnd`) — addressed by index because a
   * block's content-hashed unid changes across the save/reproject a remount performs.
   */
  private remount(focusIndex = -1, caretAtEnd = false): void {
    this.teardownBlockDrag();
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
    this.setupBlockDrag();
  }
}

/**
 * The ribbon surface — the full editor UI as a mountable module.
 *
 * `DocxEditor` (editor.ts) is the document engine: it owns the live `DocxSession`,
 * the block wiring and every command. It has deliberately no chrome. This module is
 * the chrome — the tabbed ribbon, the status bar (with the anchor rail), the table
 * picker, the find bar and the loading overlay — wired onto exactly that command
 * surface and nothing else.
 *
 * It exists because the same UI was hand-written three times (the standalone demo,
 * the GitHub Pages landing page, the compact iframe player) and drifted. Now the
 * surface has one owner, and the three hosts differ only in how they obtain the
 * WASM exports and how much of the chrome they turn on:
 *
 * ```ts
 * // Host that already booted the .NET runtime itself:
 * const ribbon = mountRibbon(document.querySelector("#app")!, { exports });
 * ribbon.open(bytes, "contract.docx");
 *
 * // Host that wants one call and a CDN (see embed.ts):
 * await createRibbonEditor("#app", "./contract.docx", { chrome: "auto" });
 * ```
 *
 * Chrome density is measured from the ROOT ELEMENT, not the viewport: a narrow
 * embed on a wide desktop page is narrow. See ribbon-chrome.ts for the layout.
 */

import { DocxEditor } from "./editor.js";
import type {
  DocxEditorExports,
  DocxEditorOptions,
  EditorAlignment,
  EditorMatch,
  FormatKey,
} from "./editor.js";
import { threadMembers } from "./editor-comments.js";
import type { BandWhich } from "./editor-headerfooter.js";
import { TrackedChangeMode } from "./types.js";
import type { ListFormat, NumberFormat } from "./types.js";
import {
  COMMON_FONTS,
  HIGHLIGHT_COLORS,
  RIBBON_CSS,
  RIBBON_HINT_HTML,
  RIBBON_HTML,
  RIBBON_STYLE_ATTR,
  RIBBON_STYLE_VERSION,
} from "./ribbon-chrome.js";

// Re-exported so the IIFE bundle built from this module (window.DocxodusEditor)
// carries both the shell and the engine it wraps.
export { DocxEditor } from "./editor.js";
export type { DocxEditorExports, DocxEditorOptions } from "./editor.js";

/** Which layout the chrome uses. "auto" measures the root and switches at `compactBreakpoint`. */
export type RibbonChromeMode = "full" | "compact" | "auto";

/** The lifecycle the surface reports on its root as `data-state`. */
export type RibbonState = "idle" | "loading" | "ready" | "error";

/** One step of the loading narrative: what is happening, and how far along it is. */
export interface RibbonLoaderStage {
  title: string;
  copy: string;
  /** 0–100. Drives the progress bar. */
  progress: number;
  /** The short machine-side label under the bar. */
  label: string;
}

/** A rotating capability card shown while the engine streams. */
export interface RibbonLoaderFeature {
  number: string;
  title: string;
  copy: string;
}

export interface RibbonLoaderOptions {
  /** Replace the four default stages. */
  stages?: RibbonLoaderStage[];
  /** Replace the rotating cards. Pass `[]` to hide the card entirely. */
  features?: RibbonLoaderFeature[];
  /** Micro-caps line above the title. */
  eyebrow?: string;
  /** Right-hand label under the progress bar. */
  meta?: string;
  /** Milliseconds between card rotations. Default 1750. */
  rotateMs?: number;
  /** What the Retry button does. Default: reload the page. */
  onRetry?: () => void;
}

export interface RibbonOptions extends DocxEditorOptions {
  /** WASM exports. Omit to mount chrome first (showing the loader) and call `setExports` later. */
  exports?: DocxEditorExports;
  /** Layout density. Default "auto". */
  chrome?: RibbonChromeMode;
  /**
   * Boundary treatment. "flush" (default) is edge-to-edge for full-bleed hosts
   * and hosts that frame the surface themselves; "card" makes the surface carry
   * its own boundary — rounded corners, hairline border, house shadow — for a
   * drop-in embed in the middle of a page. `createRibbonEditor` defaults to
   * "card", since that is the drop-in path.
   */
  frame?: "card" | "flush";
  /** Root width, in px, below which "auto" picks compact. Default 720. */
  compactBreakpoint?: number;
  /** Show the status bar with the anchor rail (full chrome only). Default true. */
  rail?: boolean;
  /** Show New / Open / Save. Default true. */
  fileActions?: boolean;
  /**
   * Editing hint above the document: false to hide, or a string to replace it.
   * Rendered as HTML (the default copy uses `<kbd>`), so pass author-controlled
   * markup only. Default true.
   */
  hint?: boolean | string;
  /** Loading overlay: false to suppress it, or an object to reword it. Default true. */
  loader?: boolean | RibbonLoaderOptions;
  /** Name shown in the title bar and used for the downloaded file. Default "untitled.docx". */
  documentName?: string;
  /**
   * Element-id prefix. Controls are addressable as `data-dxr="<name>"` always, and
   * additionally get `id = idPrefix + name`. Omit and the surface uses bare ids when
   * they are free on the page, falling back to a generated prefix when they are not —
   * so a second ribbon never collides with the first.
   */
  idPrefix?: string;
  /** Replace the default Save behaviour (download the bytes). */
  onSave?: (bytes: Uint8Array, name: string) => void;
  /** Called after any document opens, with the live editor. */
  onOpen?: (editor: DocxEditor) => void;
  /** Called whenever the status line changes. */
  onStatus?: (text: string) => void;
  /** Called after every ribbon command, with its label and measured duration. */
  onCommand?: (label: string, ms: number) => void;
}

/** Drives the loading overlay. Safe to call when the loader is disabled — every method no-ops. */
export interface RibbonLoader {
  /** Re-show the overlay (e.g. before loading a second document). */
  show(): void;
  /** Jump to a numbered stage, or set one inline. */
  stage(step: number | Partial<RibbonLoaderStage>): void;
  /** Move the bar without changing the copy. */
  progress(percent: number, label?: string): void;
  /** Fade the overlay out and stop the rotation. */
  done(): void;
  /** Show the failure state with a Retry button. */
  fail(error: unknown): void;
}

export interface RibbonEditor {
  /** The mounted root element (carries `data-state` and `data-chrome`). */
  readonly element: HTMLElement;
  /** The element the document renders into. */
  readonly surface: HTMLElement;
  /** The live editor, or null before a document is open. */
  readonly editor: DocxEditor | null;
  /** The density actually in effect right now. */
  readonly chrome: "full" | "compact";
  /** The loading overlay controller. */
  readonly loader: RibbonLoader;
  /** Supply (or replace) the WASM exports after mounting. */
  setExports(exports: DocxEditorExports): void;
  /** Open a document, replacing any open one. Throws if exports are not set yet. */
  open(bytes: Uint8Array, name?: string): DocxEditor;
  /** Open a fresh blank document. */
  openBlank(name?: string): DocxEditor;
  /** Lossless DOCX bytes, or null when nothing is open. */
  save(): Uint8Array | null;
  /** Save and hand the bytes to `onSave` (default: browser download). */
  download(name?: string): void;
  /** Set the status line. */
  setStatus(text: string): void;
  /** Activate a ribbon tab by name ("home" | "insert" | "layout" | "references" | "review" | "view" | "table" | "headerfooter"). */
  selectTab(name: string): void;
  /** Force a density, or hand control back to measurement with "auto". */
  setChrome(mode: RibbonChromeMode): void;
  /** Look up a control by its `data-dxr` name. */
  control<T extends HTMLElement = HTMLElement>(name: string): T | null;
  /** Close the session, drop listeners and empty the container. */
  destroy(): void;
}

const DEFAULT_STAGES: RibbonLoaderStage[] = [
  {
    title: "Booting .NET inside your browser",
    copy: "Streaming the trimmed WebAssembly runtime. No document bytes are sent to a server.",
    progress: 16,
    label: "Loading engine",
  },
  {
    title: "Opening the document",
    copy: "Parsing OOXML parts, styles, numbering, tables, notes, and tracked revisions locally.",
    progress: 54,
    label: "Reading document",
  },
  {
    title: "Wiring lossless editing",
    copy: "Connecting every editable block to its native WordprocessingML anchor.",
    progress: 84,
    label: "Mounting editor",
  },
  {
    title: "Your local editor is ready",
    copy: "Select text, use the ribbon, switch page view, and save a real DOCX.",
    progress: 100,
    label: "Ready",
  },
];

const DEFAULT_FEATURES: RibbonLoaderFeature[] = [
  { number: "01", title: "Zero-upload architecture", copy: "Your source file and edits never leave this browser session." },
  { number: "02", title: "Word-grade OOXML fidelity", copy: "Tables, numbering, footnotes, redlines, comments, and styles stay native." },
  { number: "03", title: "Surgical editing", copy: "Formatting and text edits target real document anchors instead of flattening the file." },
  { number: "04", title: "Lossless DOCX out", copy: "Undo, redo, edit, and save a Word document that remains a Word document." },
];

/** Every `data-dxr` name the template defines — the id-collision probe set. */
const CONTROL_NAMES = [
  "docname", "new", "file", "save", "undo", "redo", "status", "ribbon",
  "fontsize", "fontsizes", "fontfamily", "fontgrow", "fontshrink", "clearformat", "smallcaps",
  "fontcolorbtn", "fontcolorbar", "fontcolor", "highlight", "listmenu", "linespacing",
  "style", "delblock", "findtoggle", "replacetoggle",
  "table", "picturefile", "link", "unlink", "insertcomment", "gotoheader", "gotofooter", "pagenummenu",
  "hr", "hrThick", "hrDouble", "rulepos", "hrClear",
  "margins", "orientation", "pagesize", "pagesetupnote", "spacebefore", "spaceafter",
  "specialindent", "specialindentval", "pgfmt", "pgstart", "pgclear", "paginated", "headerfooter",
  "toc", "footnote", "endnote", "pagenum", "totalpages",
  "comment", "commentdelete", "commentprev", "commentnext", "commentresolve", "showcomments",
  "commentcount", "trackchanges", "markup", "author", "revcount",
  "viewpage", "viewweb", "zoom", "zoomfit", "showrail", "showhint",
  "shadebtn", "shadebar", "shade", "repeatheader",
  "hfFirst", "hfEven", "hfstory",
  "findbar", "findtext", "findprev", "findnext", "findcount", "findcase", "replacegroup",
  "replacetext", "replaceone", "replaceall", "findclose",
  "railAnchor", "railBlocks", "railSession", "railOp", "pageinfo", "wordcount",
  "zoomout", "zoomlevel", "zoomin",
  "gridpicker", "gridcells", "gridlabel", "gridalign", "gridborderless",
  "linkpop", "linkform", "linkurl", "linkcancel",
  "editor",
  "loader", "loaderEyebrow", "loaderTitle", "loaderCopy", "loaderAd", "loaderNumber",
  "loaderAdTitle", "loaderAdCopy", "loaderBar", "loaderLabel", "loaderMeta", "loaderRetry",
];

const GRID_ROWS = 8;
const GRID_COLS = 10;
const ZOOM_STEPS = [0.5, 0.75, 0.9, 1, 1.25, 1.5, 2];

let nextIdPrefixSeed = 0;

/** Inject the ribbon stylesheet once per document (replacing an older version). */
function ensureStyles(doc: Document): void {
  const existing = doc.querySelector<HTMLStyleElement>(`style[${RIBBON_STYLE_ATTR}]`);
  if (existing?.getAttribute(RIBBON_STYLE_ATTR) === RIBBON_STYLE_VERSION) return;
  existing?.remove();
  const style = doc.createElement("style");
  style.setAttribute(RIBBON_STYLE_ATTR, RIBBON_STYLE_VERSION);
  style.textContent = RIBBON_CSS;
  (doc.head ?? doc.documentElement).appendChild(style);
}

/**
 * Pick an id prefix. An explicit one always wins. Otherwise bare ids are used when
 * every name is free on the page — which keeps the historical ids of the standalone
 * demo intact — and a generated prefix is used the moment any of them is taken.
 */
function resolveIdPrefix(explicit: string | undefined, doc: Document): string {
  if (explicit !== undefined) return explicit;
  if (!CONTROL_NAMES.some((name) => doc.getElementById(name))) return "";
  for (;;) {
    const candidate = `dxr${++nextIdPrefixSeed}-`;
    if (!CONTROL_NAMES.some((name) => doc.getElementById(candidate + name))) return candidate;
  }
}

function resolveContainer(container: string | HTMLElement): HTMLElement {
  if (typeof container !== "string") return container;
  const el = document.querySelector<HTMLElement>(container);
  if (!el) throw new Error(`Docxodus ribbon: no element matches "${container}"`);
  return el;
}

/** CSS colour → "#rrggbb" (for the colour inputs), or null when it cannot be expressed. */
function toHex(color: string): string | null {
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color);
  if (m) {
    return "#" + [m[1], m[2], m[3]].map((n) => parseInt(n, 10).toString(16).padStart(2, "0")).join("");
  }
  if (/^#[0-9a-f]{6}$/i.test(color)) return color.toLowerCase();
  return null;
}

/**
 * Mount the ribbon surface into `container`.
 *
 * The chrome paints synchronously — including the loading overlay — so a host can
 * mount before its WASM runtime exists, narrate the boot through `loader`, then call
 * `setExports` and `open`.
 */
export function mountRibbon(
  container: string | HTMLElement,
  options: RibbonOptions = {},
): RibbonEditor {
  return new RibbonSurface(resolveContainer(container), options);
}

class RibbonSurface implements RibbonEditor {
  readonly element: HTMLElement;
  readonly surface: HTMLElement;
  readonly loader: RibbonLoader;

  private readonly options: RibbonOptions;
  private readonly idPrefix: string;
  private readonly loaderOptions: RibbonLoaderOptions | null;
  private readonly stages: RibbonLoaderStage[];
  private readonly features: RibbonLoaderFeature[];

  private exports: DocxEditorExports | null;
  private live: DocxEditor | null = null;
  private documentName: string;
  private chromeMode: RibbonChromeMode;
  private density: "full" | "compact" = "full";
  private headerFooter: boolean;
  private author: string;
  private destroyed = false;

  private featureTimer: ReturnType<typeof setInterval> | null = null;
  private featureIndex = 0;
  private resizeObserver: ResizeObserver | null = null;
  private strips: HTMLElement[] = [];
  private stripObserver: ResizeObserver | null = null;
  private lastAnchorText = "";
  private selectionFrame: number | null = null;
  private statsTimer: ReturnType<typeof setTimeout> | null = null;
  private findMatches: EditorMatch[] = [];
  private findIndex = -1;
  private readonly onSelectionChange: () => void;
  private readonly onDocumentMouseDown: (event: MouseEvent) => void;
  private readonly onKeydown: (event: KeyboardEvent) => void;

  constructor(container: HTMLElement, options: RibbonOptions) {
    this.options = options;
    this.exports = options.exports ?? null;
    this.documentName = options.documentName ?? "untitled.docx";
    this.chromeMode = options.chrome ?? "auto";
    // Headers and footers are document content, so the surface edits them by default; the
    // engine's own default stays off for consumers that mount `DocxEditor` bare.
    this.headerFooter = options.headerFooter ?? true;
    this.author = options.commentAuthor ?? options.revisionAuthor ?? "Reviewer";
    // false suppresses the overlay; true/omitted takes the defaults; an object rewords it.
    this.loaderOptions =
      options.loader === false
        ? null
        : typeof options.loader === "object"
          ? options.loader
          : {};
    this.stages = this.loaderOptions?.stages ?? DEFAULT_STAGES;
    this.features = this.loaderOptions?.features ?? DEFAULT_FEATURES;

    const doc = container.ownerDocument ?? document;
    ensureStyles(doc);
    this.idPrefix = resolveIdPrefix(options.idPrefix, doc);

    const root = doc.createElement("div");
    root.className = "dxr";
    root.dataset.state = "idle";
    root.dataset.frame = this.options.frame ?? "flush";
    root.innerHTML = RIBBON_HTML;
    for (const el of Array.from(root.querySelectorAll<HTMLElement>("[data-dxr]"))) {
      el.id = this.idPrefix + el.dataset.dxr;
    }
    for (const el of Array.from(root.querySelectorAll<HTMLElement>("[data-dxr-list]"))) {
      el.setAttribute("list", this.idPrefix + el.dataset.dxrList);
    }
    container.replaceChildren(root);
    this.element = root;
    this.surface = this.require("editor");

    this.applyStaticOptions();
    this.buildGrid();
    this.populateStaticMenus();
    this.wire();
    this.wireScrollFades();
    this.applyChrome();

    this.loader = this.createLoaderController();
    if (this.loaderOptions) this.loader.show();

    // Coalesced to one frame: selectionchange fires per keystroke and continuously
    // during a drag, and the sync reads computed styles.
    this.onSelectionChange = () => {
      if (this.selectionFrame != null) cancelAnimationFrame(this.selectionFrame);
      this.selectionFrame = requestAnimationFrame(() => {
        this.selectionFrame = null;
        this.syncSelection();
      });
    };
    doc.addEventListener("selectionchange", this.onSelectionChange);
    this.onDocumentMouseDown = (event) => this.maybeClosePopovers(event);
    doc.addEventListener("mousedown", this.onDocumentMouseDown);
    this.onKeydown = (event) => this.handleShortcut(event);
    root.addEventListener("keydown", this.onKeydown);

    if (this.chromeMode === "auto" && typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.applyChrome());
      this.resizeObserver.observe(root);
    }
  }

  // ── element lookup ──────────────────────────────────────────────────────────

  control<T extends HTMLElement = HTMLElement>(name: string): T | null {
    return this.element.querySelector<T>(`[data-dxr="${name}"]`);
  }

  private require<T extends HTMLElement = HTMLElement>(name: string): T {
    const el = this.control<T>(name);
    if (!el) throw new Error(`Docxodus ribbon: template is missing "${name}"`);
    return el;
  }

  // ── mount-time configuration ────────────────────────────────────────────────

  private applyStaticOptions(): void {
    const hintEl = this.element.querySelector<HTMLElement>("[data-dxr-hint]");
    if (hintEl) {
      if (this.options.hint === false) hintEl.remove();
      else hintEl.innerHTML = typeof this.options.hint === "string" ? this.options.hint : RIBBON_HINT_HTML;
    }
    if (this.options.rail === false) {
      this.element.querySelector("[data-dxr-rail]")?.remove();
    }
    if (this.options.fileActions === false) {
      this.element.querySelector("[data-dxr-files]")?.remove();
    }
    if (!this.loaderOptions) this.control("loader")?.remove();

    this.require("docname").textContent = this.documentName;
    this.require<HTMLInputElement>("paginated").checked = this.options.paginated ?? false;
    this.require<HTMLInputElement>("headerfooter").checked = this.headerFooter;
    this.require<HTMLInputElement>("trackchanges").checked =
      this.options.trackedChanges === TrackedChangeMode.RenderInline;
    this.require<HTMLInputElement>("author").value = this.author;
    this.surface.dataset.view = this.options.paginated ? "paginated" : "continuous";
    this.surface.dataset.comments = this.options.comments === false ? "off" : "on";
  }

  private buildGrid(): void {
    const cells = this.require("gridcells");
    const fragment = document.createDocumentFragment();
    for (let r = 0; r < GRID_ROWS; r++) {
      for (let c = 0; c < GRID_COLS; c++) {
        const cell = document.createElement("div");
        cell.dataset.r = String(r);
        cell.dataset.c = String(c);
        fragment.appendChild(cell);
      }
    }
    cells.replaceChildren(fragment);
  }

  /** Menus whose entries do not depend on the document. */
  private populateStaticMenus(): void {
    const highlight = this.require<HTMLSelectElement>("highlight");
    for (const color of HIGHLIGHT_COLORS) {
      const opt = document.createElement("option");
      opt.value = color.value;
      opt.textContent = `■ ${color.label}`;
      opt.style.color = color.css;
      highlight.appendChild(opt);
    }
    this.populateFonts([]);
  }

  /** The font menu: the document's own fonts first, then the common list, no duplicates. */
  private populateFonts(documentFonts: string[]): void {
    const select = this.require<HTMLSelectElement>("fontfamily");
    const seen = new Set<string>();
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Font…";
    select.appendChild(placeholder);
    const add = (name: string, group: HTMLElement | HTMLSelectElement) => {
      const key = name.toLowerCase();
      if (seen.has(key)) return;
      seen.add(key);
      const opt = document.createElement("option");
      opt.value = name;
      opt.textContent = name;
      opt.style.fontFamily = `"${name}", sans-serif`;
      group.appendChild(opt);
    };
    if (documentFonts.length > 0) {
      const group = document.createElement("optgroup");
      group.label = "In this document";
      for (const name of documentFonts) add(name, group);
      select.appendChild(group);
    }
    const common = document.createElement("optgroup");
    common.label = "Common fonts";
    for (const name of COMMON_FONTS) add(name, common);
    select.appendChild(common);
  }

  /** The style gallery: every paragraph style the document defines, quick styles first, plus
   *  the built-ins Word always offers (the engine creates a missing built-in on first use). */
  private populateStyles(): void {
    const select = this.require<HTMLSelectElement>("style");
    const styles = this.live?.styles() ?? [];
    const paragraphStyles = styles.filter((s) => s.type === "paragraph" && !s.semiHidden);
    if (paragraphStyles.length === 0) return; // older bundle: keep the static five
    const have = new Set(paragraphStyles.map((s) => s.id.toLowerCase()));
    const builtIns: Array<[string, string, number]> = [
      ["Normal", "Normal", 0], ["Heading1", "Heading 1", 9], ["Heading2", "Heading 2", 9],
      ["Heading3", "Heading 3", 9], ["Title", "Title", 10], ["Subtitle", "Subtitle", 11],
      ["Quote", "Quote", 29], ["ListParagraph", "List Paragraph", 34],
    ];
    for (const [id, name, priority] of builtIns) {
      if (have.has(id.toLowerCase())) continue;
      paragraphStyles.push({
        id, name, type: "paragraph", isDefault: false, isCustom: false, hasLatentException: false,
        uiPriority: priority, quickFormat: true,
      });
    }
    paragraphStyles.sort((a, b) => {
      const qa = a.quickFormat ? 0 : 1;
      const qb = b.quickFormat ? 0 : 1;
      if (qa !== qb) return qa - qb;
      const pa = a.uiPriority ?? 99;
      const pb = b.uiPriority ?? 99;
      if (pa !== pb) return pa - pb;
      return a.name.localeCompare(b.name);
    });
    select.replaceChildren();
    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "Style…";
    select.appendChild(placeholder);
    for (const style of paragraphStyles) {
      const opt = document.createElement("option");
      opt.value = style.id;
      opt.textContent = style.name;
      select.appendChild(opt);
    }
  }

  // ── chrome density ──────────────────────────────────────────────────────────

  get chrome(): "full" | "compact" {
    return this.density;
  }

  setChrome(mode: RibbonChromeMode): void {
    this.chromeMode = mode;
    if (mode === "auto" && !this.resizeObserver && typeof ResizeObserver !== "undefined") {
      this.resizeObserver = new ResizeObserver(() => this.applyChrome());
      this.resizeObserver.observe(this.element);
    }
    this.applyChrome();
  }

  private applyChrome(): void {
    const breakpoint = this.options.compactBreakpoint ?? 720;
    const width = this.element.clientWidth || this.element.getBoundingClientRect().width;
    const next: "full" | "compact" =
      this.chromeMode === "auto"
        // Width 0 means the root is not laid out yet (display:none, or measured before
        // first paint). Assume the roomier layout and let the observer correct it.
        ? (width > 0 && width < breakpoint ? "compact" : "full")
        : this.chromeMode;
    if (next === this.density && this.element.dataset.chrome) return;
    this.density = next;
    this.element.dataset.chrome = next;
    // A popover's absolute position is meaningless after a density flip.
    this.closePopovers();
    // A density flip relays out every strip; re-measure without waiting for
    // the observer so the fades never lag a frame behind the new layout.
    this.syncAllStripFades();
  }

  // ── scroll-edge affordances ─────────────────────────────────────────────────

  private syncStripFade(strip: HTMLElement): void {
    const max = strip.scrollWidth - strip.clientWidth;
    const tokens =
      max <= 1 ? "" : `${strip.scrollLeft > 1 ? "l " : ""}${strip.scrollLeft < max - 1 ? "r" : ""}`.trim();
    if (tokens) strip.setAttribute("data-fade", tokens);
    else strip.removeAttribute("data-fade");
  }

  private syncAllStripFades(): void {
    for (const strip of this.strips) this.syncStripFade(strip);
  }

  private wireScrollFades(): void {
    this.strips = Array.from(
      this.element.querySelectorAll<HTMLElement>(".dxr-titlebar, .dxr-tabs, .dxr-panel, .dxr-rail, .dxr-findbar"),
    );
    if (typeof ResizeObserver !== "undefined") {
      // One observer for all strips. It also fires when a panel becomes the
      // active one (display flips from none), which is what keeps a freshly
      // selected tab's fades honest without per-tab wiring.
      this.stripObserver = new ResizeObserver((entries) => {
        for (const entry of entries) this.syncStripFade(entry.target as HTMLElement);
      });
    }
    for (const strip of this.strips) {
      strip.addEventListener("scroll", () => this.syncStripFade(strip), { passive: true });
      this.stripObserver?.observe(strip);
      this.syncStripFade(strip);
    }
  }

  // ── status ──────────────────────────────────────────────────────────────────

  /** Publish the lifecycle on the root, where CSS and host pages can key off it. */
  private setState(state: RibbonState): void {
    this.element.dataset.state = state;
  }

  setStatus(text: string): void {
    const el = this.control("status");
    if (el) el.textContent = text;
    this.options.onStatus?.(text);
  }

  // ── document lifecycle ──────────────────────────────────────────────────────

  get editor(): DocxEditor | null {
    return this.live;
  }

  setExports(exports: DocxEditorExports): void {
    this.exports = exports;
  }

  open(bytes: Uint8Array, name?: string): DocxEditor {
    if (!this.exports) throw new Error("Docxodus ribbon: WASM exports are not set yet");
    if (this.live) {
      try {
        this.live.close();
      } catch {
        /* already closed */
      }
      this.live = null;
    }
    if (name) this.documentName = name;
    this.require("docname").textContent = this.documentName;

    const paginated = this.require<HTMLInputElement>("paginated").checked;
    this.surface.dataset.view = paginated ? "paginated" : "continuous";
    this.surface.replaceChildren();
    this.closeFindBar();

    const started = performance.now();
    const tracked = this.require<HTMLInputElement>("trackchanges").checked
      ? TrackedChangeMode.RenderInline
      : (this.options.trackedChanges ?? TrackedChangeMode.Accept);
    this.live = DocxEditor.open(this.surface, bytes, this.exports, {
      cssPrefix: this.options.cssPrefix,
      fabricateClasses: this.options.fabricateClasses,
      editable: this.options.editable,
      scale: this.options.scale,
      columnWidth: this.options.columnWidth,
      fitToWidth: this.options.fitToWidth,
      onEdit: (info) => {
        this.options.onEdit?.(info);
        this.scheduleStats();
      },
      onMove: this.options.onMove,
      onStoryChange: (which) => this.onStoryChange(which),
      onCommentsChange: (info) => this.onCommentsChange(info),
      paginated,
      headerFooter: this.headerFooter,
      blockDrag: this.options.blockDrag ?? true,
      trackedChanges: tracked,
      revisionAuthor: this.author,
      comments: this.options.comments,
      commentAuthor: this.author,
    });
    this.require<HTMLButtonElement>("save").disabled = false;
    this.require("ribbon").setAttribute("aria-disabled", "false");
    this.setState("ready");
    this.setStatus(`Rendered in ${Math.round(performance.now() - started)} ms`);
    this.populateStyles();
    this.populateFonts(this.documentFonts());
    this.syncPageNumbering();
    this.syncLayoutState();
    this.syncReviewState();
    this.syncZoom();
    this.refreshRailCounts();
    this.refreshRailAnchor();
    this.scheduleStats();
    this.options.onOpen?.(this.live);
    return this.live;
  }

  openBlank(name = "untitled.docx"): DocxEditor {
    if (!this.exports) throw new Error("Docxodus ribbon: WASM exports are not set yet");
    return this.open(this.exports.DocxSessionBridge.CreateBlankDocx(), name);
  }

  save(): Uint8Array | null {
    return this.live ? this.live.save() : null;
  }

  download(name?: string): void {
    const bytes = this.save();
    if (!bytes) return;
    const filename = name ?? this.documentName ?? "edited.docx";
    if (this.options.onSave) {
      this.options.onSave(bytes, filename);
      return;
    }
    const url = URL.createObjectURL(
      new Blob([bytes as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      }),
    );
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    this.setStatus(`Saved ${filename}`);
  }

  destroy(): void {
    if (this.destroyed) return;
    this.destroyed = true;
    const doc = this.element.ownerDocument ?? document;
    doc.removeEventListener("selectionchange", this.onSelectionChange);
    doc.removeEventListener("mousedown", this.onDocumentMouseDown);
    this.element.removeEventListener("keydown", this.onKeydown);
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    this.stripObserver?.disconnect();
    this.stripObserver = null;
    if (this.selectionFrame != null) cancelAnimationFrame(this.selectionFrame);
    this.selectionFrame = null;
    if (this.statsTimer != null) clearTimeout(this.statsTimer);
    this.stopRotation();
    try {
      this.live?.close();
    } catch {
      /* already closed */
    }
    this.live = null;
    this.element.remove();
  }

  // ── command plumbing ────────────────────────────────────────────────────────

  /**
   * Run one ribbon command and report its real cost on the rail.
   *
   * Every control routes through here, which is what makes the rail's "last op"
   * an honest measurement rather than a label the surface makes up.
   */
  private run(label: string, fn: () => void): void {
    if (!this.live) return;
    const started = performance.now();
    try {
      fn();
    } finally {
      const ms = performance.now() - started;
      const el = this.control("railOp");
      if (el) el.textContent = `${label} ${ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${Math.round(ms)} ms`}`;
      this.options.onCommand?.(label, ms);
      this.refreshRailCounts();
      this.refreshRailAnchor();
      this.scheduleStats();
    }
  }

  /** Format controls must not steal the document selection they are about to act on. */
  private keepSelection(el: HTMLElement): void {
    el.addEventListener("mousedown", (event) => event.preventDefault());
  }

  private wire(): void {
    const ribbon = this.require("ribbon");

    for (const tab of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-tab"))) {
      this.keepSelection(tab);
      tab.addEventListener("click", () => this.selectTab(tab.dataset.tab ?? "home"));
    }

    const delegate = <T extends HTMLElement>(selector: string, label: (el: T) => string, fn: (el: T) => void) => {
      for (const el of Array.from(ribbon.querySelectorAll<T>(selector))) {
        this.keepSelection(el);
        el.addEventListener("click", () => this.run(label(el), () => fn(el)));
      }
    };

    delegate<HTMLElement>("button[data-cmd]", (b) => b.dataset.cmd!, (b) =>
      this.live!.format(b.dataset.cmd as FormatKey));
    delegate<HTMLElement>("button[data-align]", (b) => `align ${b.dataset.align}`, (b) =>
      this.live!.setAlignment(b.dataset.align as EditorAlignment));
    delegate<HTMLElement>("button[data-indent]", () => "indent", (b) =>
      this.live!.indent(parseInt(b.dataset.indent ?? "720", 10)));
    delegate<HTMLElement>("button[data-list]", (b) => `list ${b.dataset.list}`, (b) =>
      this.live!.toggleList(b.dataset.list as "bullet" | "decimal"));
    delegate<HTMLElement>("button[data-pagebreak]", () => "page break", () =>
      this.live!.pageBreakBefore(true));
    delegate<HTMLElement>('.dxr-panel[data-panel="table"] button[data-tt]', (b) => b.dataset.tt!, (b) => {
      const ops: Record<string, () => void> = {
        rowAbove: () => this.live!.insertTableRow("above"),
        rowBelow: () => this.live!.insertTableRow("below"),
        colLeft: () => this.live!.insertTableColumn("left"),
        colRight: () => this.live!.insertTableColumn("right"),
        delRow: () => this.live!.deleteTableRow(),
        delCol: () => this.live!.deleteTableColumn(),
        mergeRight: () => this.live!.mergeCells(1, 2),
        mergeDown: () => this.live!.mergeCells(2, 1),
        unmerge: () => this.live!.unmergeCells(),
        borders: () => this.live!.setTableBorders({ scope: "all", style: "single", size: 4 }),
        noborders: () => this.live!.setTableBorders({ scope: "all", style: "none" }),
        noshade: () => this.live!.setCellShading("", "cell"),
        delTable: () => this.live!.deleteTable(),
      };
      ops[b.dataset.tt ?? ""]?.();
    });
    delegate<HTMLElement>("button[data-rev]", (b) => `revision ${b.dataset.rev}`, (b) =>
      this.runRevisionCommand(b.dataset.rev ?? ""));
    delegate<HTMLElement>("button[data-hf]", (b) => `header/footer ${b.dataset.hf}`, (b) =>
      this.runHeaderFooterCommand(b.dataset.hf ?? ""));

    const rulepos = () => (this.require<HTMLSelectElement>("rulepos").value === "above" ? "above" : "below");
    const simple: Array<[string, string, () => void]> = [
      ["undo", "undo", () => this.live!.undo()],
      ["redo", "redo", () => this.live!.redo()],
      ["hr", "rule", () => this.live!.insertHorizontalRule(12, "single", rulepos())],
      ["hrThick", "thick rule", () => this.live!.insertHorizontalRule(24, "single", rulepos())],
      ["hrDouble", "double rule", () => this.live!.insertHorizontalRule(12, "double", rulepos())],
      ["hrClear", "clear border", () => this.live!.clearParagraphBorders()],
      ["delblock", "delete block", () => this.live!.deleteBlock()],
      ["footnote", "footnote", () => this.live!.insertFootnote()],
      ["endnote", "endnote", () => this.live!.insertEndnote()],
      ["pagenum", "page number", () => this.live!.insertPageNumber("currentPage")],
      ["totalpages", "total pages", () => this.live!.insertPageNumber("totalPages")],
      ["fontgrow", "grow font", () => this.live!.adjustFontSize(1)],
      ["fontshrink", "shrink font", () => this.live!.adjustFontSize(-1)],
      ["clearformat", "clear formatting", () => this.live!.clearFormatting()],
      ["toc", "table of contents", () => this.live!.insertTableOfContents()],
      ["unlink", "remove link", () => this.live!.removeHyperlink()],
      ["gotoheader", "go to header", () => this.live!.goToHeaderFooter("header")],
      ["gotofooter", "go to footer", () => this.live!.goToHeaderFooter("footer")],
      ["viewpage", "page view", () => this.setPaginated(true)],
      ["viewweb", "continuous view", () => this.setPaginated(false)],
      ["zoomfit", "fit width", () => this.setZoom(1)],
      ["zoomin", "zoom in", () => this.stepZoom(1)],
      ["zoomout", "zoom out", () => this.stepZoom(-1)],
    ];
    for (const [name, label, fn] of simple) {
      const el = this.control(name);
      if (!el) continue;
      this.keepSelection(el);
      el.addEventListener("click", () => this.run(label, fn));
    }

    const smallcaps = this.control("smallcaps");
    if (smallcaps) {
      this.keepSelection(smallcaps);
      smallcaps.addEventListener("click", () => {
        const on = smallcaps.getAttribute("aria-pressed") !== "true";
        this.run("small caps", () => this.live!.setSmallCaps(on));
        smallcaps.setAttribute("aria-pressed", String(on));
      });
    }

    // Font size: `change` only (it fires on Enter and on blur). Binding keydown as
    // well fired twice and stacked two undo snapshots for one size change.
    const fontsize = this.require<HTMLInputElement>("fontsize");
    fontsize.addEventListener("change", () => {
      const pts = parseFloat(fontsize.value);
      if (this.live && pts > 0) this.run("font size", () => this.live!.setFontSize(pts));
    });
    fontsize.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        fontsize.blur();
      }
    });

    this.wireSelect("fontfamily", "font family", (v) => this.live!.setFontFamily(v));
    this.wireSelect("style", "style", (v) => this.live!.setParagraphStyle(v));
    this.wireSelect("listmenu", "list style", (v) => this.live!.setListFormat(v as ListFormat));
    this.wireSelect("linespacing", "line spacing", (v) => this.live!.setLineSpacing(parseFloat(v)));
    this.wireSelect("highlight", "highlight", (v) => this.live!.setHighlight(v === "none" ? "" : v));
    this.wireSelect("pagenummenu", "page number", (v) =>
      this.live!.insertPageNumber(v === "pageOf" ? "pageOfTotal" : (v as "currentPage" | "totalPages")));

    // Colour pickers: the native input steals focus, so the engine's last-selection
    // bookmark is what the command applies to.
    const fontcolor = this.require<HTMLInputElement>("fontcolor");
    fontcolor.addEventListener("input", () => {
      this.require("fontcolorbar").style.setProperty("--dxr-swatch", fontcolor.value);
    });
    fontcolor.addEventListener("change", () => {
      this.require("fontcolorbar").style.setProperty("--dxr-swatch", fontcolor.value);
      this.run("font color", () => this.live!.setFontColor(fontcolor.value));
    });
    const fontcolorbtn = this.control("fontcolorbtn");
    if (fontcolorbtn) {
      this.keepSelection(fontcolorbtn);
      fontcolorbtn.addEventListener("click", () => this.run("font color", () => this.live!.setFontColor(fontcolor.value)));
    }
    const shade = this.require<HTMLInputElement>("shade");
    shade.addEventListener("input", () => this.require("shadebar").style.setProperty("--dxr-swatch", shade.value));
    shade.addEventListener("change", () => {
      this.require("shadebar").style.setProperty("--dxr-swatch", shade.value);
      this.run("cell shading", () => this.live!.setCellShading(shade.value, "cell"));
    });
    const shadebtn = this.control("shadebtn");
    if (shadebtn) {
      this.keepSelection(shadebtn);
      shadebtn.addEventListener("click", () => this.run("cell shading", () => this.live!.setCellShading(shade.value, "cell")));
    }
    const repeatheader = this.require<HTMLInputElement>("repeatheader");
    repeatheader.addEventListener("change", () =>
      this.run("repeat header row", () => this.live!.setRepeatHeaderRow(repeatheader.checked)));

    // Comments: New/Insert open a draft bubble in the gutter (Word's flow); the rest act on the
    // active thread.
    for (const name of ["comment", "insertcomment"]) {
      const button = this.control(name);
      if (!button) continue;
      this.keepSelection(button);
      button.addEventListener("click", () => this.beginComment());
    }
    this.control("commentdelete")?.addEventListener("click", () => {
      const active = this.live?.activeComment;
      if (!active) return;
      this.run("delete comment", () => {
        for (const id of threadMembers(this.live!.listComments(), active)) this.live!.removeComment(id);
        this.live!.removeComment(active);
      });
    });
    this.control("commentresolve")?.addEventListener("click", () => {
      const active = this.live?.activeComment;
      if (!active) return;
      const entry = this.live!.listComments().find((c) => c.anchorId === active);
      if (!entry) return;
      this.run(entry.resolved ? "reopen comment" : "resolve comment", () =>
        this.live!.setCommentResolved(active, !entry.resolved));
      this.live!.activateComment(active);
    });
    this.control("commentprev")?.addEventListener("click", () => this.live?.stepComment(-1));
    this.control("commentnext")?.addEventListener("click", () => this.live?.stepComment(1));
    const showcomments = this.require<HTMLInputElement>("showcomments");
    showcomments.addEventListener("change", () => {
      this.live?.showComments(showcomments.checked);
      this.surface.dataset.comments = showcomments.checked && this.options.comments !== false ? "on" : "off";
    });

    // Tracking.
    const trackchanges = this.require<HTMLInputElement>("trackchanges");
    trackchanges.addEventListener("change", () => {
      if (!this.live) return;
      this.run(trackchanges.checked ? "track changes on" : "track changes off", () =>
        this.live!.setTrackedChanges(trackchanges.checked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept));
      this.syncReviewState();
    });
    const markup = this.require<HTMLSelectElement>("markup");
    markup.addEventListener("change", () => {
      if (!this.live) return;
      // "No markup" shows the document as if every change were accepted; edits keep recording.
      this.run(`markup ${markup.value}`, () =>
        this.live!.setTrackedChanges(markup.value === "none" ? TrackedChangeMode.Accept : TrackedChangeMode.RenderInline));
      trackchanges.checked = markup.value !== "none";
    });
    const author = this.require<HTMLInputElement>("author");
    author.addEventListener("change", () => {
      this.author = author.value.trim() || "Reviewer";
      this.live?.setRevisionAuthor(this.author);
    });

    this.wireFileActions();
    this.wireLayout();
    this.wirePicker();
    this.wireLinkPopover();
    this.wireFindBar();
    this.wireHeaderFooterTab();
    this.wireZoom();
    this.wireViewToggles();
  }

  /** A `<select>` that acts on change and snaps back to its placeholder. */
  private wireSelect(name: string, label: string, apply: (value: string) => void): void {
    const select = this.require<HTMLSelectElement>(name);
    select.addEventListener("change", () => {
      const value = select.value;
      if (this.live && value) this.run(label, () => apply(value));
      select.value = "";
    });
  }

  private wireFileActions(): void {
    const file = this.control<HTMLInputElement>("file");
    file?.addEventListener("change", async () => {
      const chosen = file.files?.[0];
      if (!chosen) return;
      this.setStatus(`Loading ${chosen.name}…`);
      this.open(new Uint8Array(await chosen.arrayBuffer()), chosen.name);
      // Reset so re-picking the same file fires `change` again.
      file.value = "";
    });
    this.control("new")?.addEventListener("click", () => {
      if (this.exports) this.openBlank("untitled.docx");
    });
    this.control("save")?.addEventListener("click", () => this.download());

    const picture = this.control<HTMLInputElement>("picturefile");
    picture?.addEventListener("change", async () => {
      const chosen = picture.files?.[0];
      picture.value = "";
      if (!chosen || !this.live) return;
      const started = performance.now();
      const ok = await this.live.insertImageFile(chosen, { altText: chosen.name });
      const ms = performance.now() - started;
      const el = this.control("railOp");
      if (el) el.textContent = `picture ${Math.round(ms)} ms`;
      this.setStatus(ok ? `Inserted ${chosen.name}` : "Could not insert the picture here");
      this.refreshRailCounts();
    });
  }

  private wireLayout(): void {
    // Pagination toggles re-render the LIVE session, so edits survive the switch.
    const paginated = this.require<HTMLInputElement>("paginated");
    paginated.addEventListener("change", () => this.setPaginated(paginated.checked));

    // The region is chosen at open time, so toggling it re-opens the LIVE
    // session's bytes rather than the originally loaded file — edits so far survive.
    const headerFooter = this.require<HTMLInputElement>("headerfooter");
    headerFooter.addEventListener("change", () => {
      this.headerFooter = headerFooter.checked;
      if (!this.live) return;
      this.open(this.live.save(), this.documentName);
    });

    const pgfmt = this.require<HTMLSelectElement>("pgfmt");
    pgfmt.addEventListener("change", () => {
      if (!this.live || !pgfmt.value) return;
      this.run("page format", () => this.live!.setPageNumbering({ format: pgfmt.value as NumberFormat }));
      this.syncPageNumbering();
    });
    const pgstart = this.require<HTMLInputElement>("pgstart");
    pgstart.addEventListener("change", () => {
      const value = parseInt(pgstart.value, 10);
      if (!this.live || !(value > 0)) return;
      this.run("page start", () => this.live!.setPageNumbering({ start: value }));
      this.syncPageNumbering();
    });
    this.control("pgclear")?.addEventListener("click", () => {
      this.run("clear numbering", () => this.live!.clearPageNumbering());
      this.syncPageNumbering();
    });

    // Page setup (Word's Layout > Page Setup): margins presets, orientation, size.
    const pageSetup = (op: Parameters<DocxEditor["setPageSetup"]>[0]) => {
      if (!this.live!.setPageSetup(op)) this.setStatus("This engine build cannot change the page setup");
      this.syncLayoutState();
    };
    this.wireSelect("margins", "margins", (v) => {
      const [top, bottom, left, right] = v.split(",").map((n) => parseInt(n, 10));
      pageSetup({ marginTopTwips: top, marginBottomTwips: bottom, marginLeftTwips: left, marginRightTwips: right });
    });
    this.wireSelect("orientation", "orientation", (v) => pageSetup({ landscape: v === "landscape" }));
    this.wireSelect("pagesize", "page size", (v) => {
      const [w, h] = v.split(",").map((n) => parseInt(n, 10));
      const landscape = !!this.live!.sectionInfo()?.landscape;
      pageSetup(landscape
        ? { pageWidthTwips: h, pageHeightTwips: w, landscape: true }
        : { pageWidthTwips: w, pageHeightTwips: h });
    });

    // Paragraph spacing / special indent.
    const before = this.require<HTMLInputElement>("spacebefore");
    const after = this.require<HTMLInputElement>("spaceafter");
    const spacing = (which: "before" | "after", input: HTMLInputElement) => {
      const pts = parseFloat(input.value);
      if (!this.live || !(pts >= 0)) return;
      this.run(`space ${which}`, () =>
        this.live!.setParagraphSpacing(which === "before" ? { beforePt: pts } : { afterPt: pts }));
    };
    before.addEventListener("change", () => spacing("before", before));
    after.addEventListener("change", () => spacing("after", after));
    const special = this.require<HTMLSelectElement>("specialindent");
    const specialVal = this.require<HTMLInputElement>("specialindentval");
    const applySpecial = () => {
      if (!this.live || !special.value) return;
      const inches = parseFloat(specialVal.value);
      const twips = Math.round((inches >= 0 ? inches : 0.5) * 1440);
      this.run("special indent", () => {
        if (special.value === "firstLine") this.live!.setFirstLineIndent(twips);
        else if (special.value === "hanging") this.live!.setHangingIndent(twips);
        else this.live!.setFirstLineIndent(0);
      });
    };
    special.addEventListener("change", applySpecial);
    specialVal.addEventListener("change", applySpecial);
  }

  private setPaginated(on: boolean): void {
    const paginated = this.require<HTMLInputElement>("paginated");
    paginated.checked = on;
    this.surface.dataset.view = on ? "paginated" : "continuous";
    if (!this.live) return;
    this.run(on ? "page view" : "continuous view", () => this.live!.setPaginated(on));
    this.syncViewButtons();
  }

  private syncViewButtons(): void {
    const on = this.require<HTMLInputElement>("paginated").checked;
    this.control("viewpage")?.classList.toggle("dxr-on", on);
    this.control("viewweb")?.classList.toggle("dxr-on", !on);
  }

  private wireViewToggles(): void {
    const showrail = this.require<HTMLInputElement>("showrail");
    showrail.addEventListener("change", () => {
      const rail = this.element.querySelector<HTMLElement>("[data-dxr-rail]");
      if (rail) rail.hidden = !showrail.checked;
    });
    const showhint = this.require<HTMLInputElement>("showhint");
    showhint.addEventListener("change", () => {
      const hint = this.element.querySelector<HTMLElement>("[data-dxr-hint]");
      if (hint) hint.hidden = !showhint.checked;
    });
    this.syncViewButtons();
  }

  /** The bands own the same setting and read the live session, so both stay in step. */
  private syncPageNumbering(): void {
    if (!this.live) return;
    const numbering = this.live.pageNumbering() ?? {};
    this.require<HTMLSelectElement>("pgfmt").value = numbering.format ?? "";
    this.require<HTMLInputElement>("pgstart").value = numbering.start != null ? String(numbering.start) : "";
  }

  /** Layout tab: describe the section under the caret. */
  private syncLayoutState(): void {
    const note = this.control("pagesetupnote");
    const info = this.live?.sectionInfo() ?? null;
    if (!note) return;
    if (!info) {
      note.textContent = "";
      return;
    }
    const inches = (twips: number) => (twips / 1440).toFixed(2).replace(/\.?0+$/, "");
    note.textContent =
      `${inches(info.pageWidthTwips)}″ × ${inches(info.pageHeightTwips)}″ ${info.landscape ? "landscape" : "portrait"}, ` +
      `margins ${inches(info.marginTopTwips)}″ / ${inches(info.marginRightTwips)}″ / ${inches(info.marginBottomTwips)}″ / ${inches(info.marginLeftTwips)}″`;
  }

  // ── review ──────────────────────────────────────────────────────────────────

  private beginComment(): void {
    if (!this.live) return;
    if (this.density === "compact") {
      // No gutter on a phone: post directly on the selection with a prompt.
      const text = window.prompt("Comment");
      if (!text?.trim()) return;
      this.run("comment", () => this.live!.addComment(text.trim(), this.author));
      return;
    }
    if (!this.live.beginComment()) this.setStatus("Click into a paragraph (or select text) to comment on it");
  }

  private onCommentsChange(info: { threads: number; open: number; active: string | null }): void {
    const count = this.control("commentcount");
    if (count) count.textContent = info.threads === 0 ? "No comments" : `${info.open} open · ${info.threads} total`;
    const resolve = this.control<HTMLButtonElement>("commentresolve");
    const del = this.control<HTMLButtonElement>("commentdelete");
    const entry = info.active ? this.live?.listComments().find((c) => c.anchorId === info.active) : null;
    if (resolve) {
      resolve.disabled = !entry;
      resolve.textContent = entry?.resolved ? "Reopen" : "Resolve";
    }
    if (del) del.disabled = !entry;
  }

  private syncReviewState(): void {
    if (!this.live) return;
    const tracked = this.live.trackedChanges === TrackedChangeMode.RenderInline;
    this.require<HTMLInputElement>("trackchanges").checked = tracked;
    this.require<HTMLSelectElement>("markup").value = tracked ? "all" : "none";
    const revs = this.live.listRevisions();
    const count = this.control("revcount");
    if (count) count.textContent = revs.length === 0 ? "No tracked changes" : `${revs.length} change${revs.length === 1 ? "" : "s"}`;
    this.require<HTMLInputElement>("showcomments").checked = this.live.commentsVisible || this.options.comments === false;
  }

  private runRevisionCommand(op: string): void {
    if (!this.live) return;
    const marks = this.live.revisionElements();
    const current = this.currentRevisionMark(marks);
    switch (op) {
      case "accept":
      case "reject": {
        const mark = current ?? marks[0];
        const entry = mark ? this.live.revisionAt(mark) : null;
        if (!entry) { this.setStatus("Put the caret in a tracked change first"); return; }
        if (op === "accept") this.live.acceptRevision(entry.id);
        else this.live.rejectRevision(entry.id);
        break;
      }
      case "acceptall": this.live.acceptAllRevisions(); break;
      case "rejectall": this.live.rejectAllRevisions(); break;
      case "next":
      case "prev": {
        if (marks.length === 0) { this.setStatus("No tracked changes"); return; }
        const index = current ? marks.indexOf(current) : -1;
        const target = marks[(index + (op === "next" ? 1 : -1) + marks.length) % marks.length];
        this.selectRevisionMark(target);
        return;
      }
      default: return;
    }
    this.syncReviewState();
  }

  private currentRevisionMark(marks: HTMLElement[]): HTMLElement | null {
    const sel = (this.element.ownerDocument ?? document).getSelection();
    const node = sel?.anchorNode ?? null;
    if (!node) return null;
    const el = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    const mark = el?.closest<HTMLElement>('ins, del, tr[class*="row-"]') ?? null;
    return mark && marks.includes(mark) ? mark : null;
  }

  private selectRevisionMark(mark: HTMLElement): void {
    const doc = this.element.ownerDocument ?? document;
    const block = mark.closest<HTMLElement>('[data-anchor][contenteditable="true"]');
    block?.focus({ preventScroll: true });
    const range = doc.createRange();
    range.selectNodeContents(mark);
    const sel = doc.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    mark.scrollIntoView({ block: "center", behavior: "smooth" });
  }

  // ── header & footer tab ─────────────────────────────────────────────────────

  private wireHeaderFooterTab(): void {
    for (const kind of ["first", "even"] as const) {
      const box = this.require<HTMLInputElement>(kind === "first" ? "hfFirst" : "hfEven");
      box.addEventListener("change", () => {
        if (!this.live) return;
        const ok = this.live.setHeaderFooterKindEnabled(kind, box.checked);
        if (!ok) {
          box.checked = !box.checked;
          this.setStatus("This engine build cannot change that option");
        }
        this.run(`${kind === "first" ? "different first page" : "different odd & even"} ${box.checked ? "on" : "off"}`, () => {});
        this.syncHeaderFooterState();
      });
    }
  }

  private runHeaderFooterCommand(op: string): void {
    if (!this.live) return;
    switch (op) {
      case "pagenum": this.live.insertPageNumber("currentPage"); break;
      case "totalpages": this.live.insertPageNumber("totalPages"); break;
      case "pageof": this.live.insertPageNumber("pageOfTotal"); break;
      case "header": this.live.goToHeaderFooter("header"); break;
      case "footer": this.live.goToHeaderFooter("footer"); break;
      case "close": this.live.closeHeaderFooter(); break;
    }
    this.syncHeaderFooterState();
  }

  private onStoryChange(which: BandWhich | null): void {
    const tab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="headerfooter"]');
    if (!tab) return;
    tab.hidden = which === null;
    if (which) {
      this.selectTab("headerfooter");
    } else if (tab.getAttribute("aria-selected") === "true") {
      this.selectTab("home");
    }
    this.syncHeaderFooterState();
  }

  private syncHeaderFooterState(): void {
    if (!this.live) return;
    this.require<HTMLInputElement>("hfFirst").checked = this.live.headerFooterKindEnabled("first");
    this.require<HTMLInputElement>("hfEven").checked = this.live.headerFooterKindEnabled("even");
    const story = this.control("hfstory");
    if (story) {
      const which = this.live.activeStoryKind ?? "header";
      story.textContent = this.live.headerFooterStoryLabel(which);
    }
  }

  // ── zoom ────────────────────────────────────────────────────────────────────

  private wireZoom(): void {
    for (const name of ["zoom", "zoomlevel"]) {
      const select = this.control<HTMLSelectElement>(name);
      select?.addEventListener("change", () => this.setZoom(parseFloat(select.value)));
    }
  }

  private setZoom(scale: number): void {
    if (!this.live || !(scale > 0)) return;
    this.run(`zoom ${Math.round(scale * 100)}%`, () => this.live!.setZoom(scale));
    this.syncZoom();
  }

  private stepZoom(direction: 1 | -1): void {
    if (!this.live) return;
    const current = this.live.requestedZoom;
    const next = direction > 0
      ? ZOOM_STEPS.find((z) => z > current + 1e-6) ?? ZOOM_STEPS[ZOOM_STEPS.length - 1]
      : [...ZOOM_STEPS].reverse().find((z) => z < current - 1e-6) ?? ZOOM_STEPS[0];
    this.setZoom(next);
  }

  private syncZoom(): void {
    if (!this.live) return;
    const requested = this.live.requestedZoom;
    const value = String(ZOOM_STEPS.find((z) => Math.abs(z - requested) < 1e-6) ?? requested);
    for (const name of ["zoom", "zoomlevel"]) {
      const select = this.control<HTMLSelectElement>(name);
      if (!select) continue;
      if (!Array.from(select.options).some((o) => o.value === value)) {
        const opt = document.createElement("option");
        opt.value = value;
        opt.textContent = `${Math.round(requested * 100)}%`;
        select.appendChild(opt);
      }
      select.value = value;
    }
  }

  // ── find & replace ──────────────────────────────────────────────────────────

  private wireFindBar(): void {
    const bar = this.require("findbar");
    const text = this.require<HTMLInputElement>("findtext");
    const replace = this.require<HTMLInputElement>("replacetext");
    const matchCase = this.require<HTMLInputElement>("findcase");
    this.control("findtoggle")?.addEventListener("click", () => this.openFindBar(false));
    this.control("replacetoggle")?.addEventListener("click", () => this.openFindBar(true));
    this.control("findclose")?.addEventListener("click", () => this.closeFindBar());
    text.addEventListener("input", () => this.refreshFind(true));
    matchCase.addEventListener("change", () => this.refreshFind(true));
    text.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.stepFind(event.shiftKey ? -1 : 1);
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closeFindBar();
      }
    });
    replace.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        this.replaceCurrent();
      } else if (event.key === "Escape") {
        event.preventDefault();
        this.closeFindBar();
      }
    });
    this.control("findprev")?.addEventListener("click", () => this.stepFind(-1));
    this.control("findnext")?.addEventListener("click", () => this.stepFind(1));
    this.control("replaceone")?.addEventListener("click", () => this.replaceCurrent());
    this.control("replaceall")?.addEventListener("click", () => {
      if (!this.live || !text.value) return;
      let count = 0;
      this.run("replace all", () => {
        count = this.live!.replaceAll(text.value, replace.value, { matchCase: matchCase.checked });
      });
      this.setStatus(`Replaced ${count} occurrence${count === 1 ? "" : "s"}`);
      this.refreshFind(true);
    });
    bar.hidden = true;
  }

  private openFindBar(withReplace: boolean): void {
    const bar = this.require("findbar");
    bar.hidden = false;
    this.require("replacegroup").hidden = !withReplace;
    this.syncStripFade(bar);
    const text = this.require<HTMLInputElement>("findtext");
    text.focus();
    text.select();
    this.refreshFind(false);
  }

  private closeFindBar(): void {
    const bar = this.control("findbar");
    if (bar) bar.hidden = true;
    this.findMatches = [];
    this.findIndex = -1;
  }

  private refreshFind(select: boolean): void {
    const text = this.require<HTMLInputElement>("findtext");
    const matchCase = this.require<HTMLInputElement>("findcase").checked;
    this.findMatches = this.live && text.value ? this.live.find(text.value, { matchCase }) : [];
    this.findIndex = this.findMatches.length > 0 ? 0 : -1;
    this.updateFindCount();
    if (select && this.findIndex >= 0) this.live?.selectMatch(this.findMatches[this.findIndex]);
  }

  private stepFind(direction: 1 | -1): void {
    if (!this.live) return;
    if (this.findMatches.length === 0) this.refreshFind(false);
    if (this.findMatches.length === 0) return;
    this.findIndex = (this.findIndex + direction + this.findMatches.length) % this.findMatches.length;
    this.live.selectMatch(this.findMatches[this.findIndex]);
    this.updateFindCount();
  }

  private replaceCurrent(): void {
    if (!this.live) return;
    if (this.findMatches.length === 0 || this.findIndex < 0) this.refreshFind(true);
    const match = this.findMatches[this.findIndex];
    if (!match) return;
    const replacement = this.require<HTMLInputElement>("replacetext").value;
    this.run("replace", () => this.live!.replaceMatch(match, replacement));
    // Offsets after the replacement shifted; re-scan and continue from the same slot.
    const keep = this.findIndex;
    this.refreshFind(false);
    if (this.findMatches.length > 0) {
      this.findIndex = Math.min(keep, this.findMatches.length - 1);
      this.live.selectMatch(this.findMatches[this.findIndex]);
      this.updateFindCount();
    }
  }

  private updateFindCount(): void {
    const el = this.control("findcount");
    if (!el) return;
    el.textContent = this.findMatches.length === 0
      ? (this.require<HTMLInputElement>("findtext").value ? "0 of 0" : "")
      : `${this.findIndex + 1} of ${this.findMatches.length}`;
  }

  // ── keyboard shortcuts the engine does not own ──────────────────────────────

  private handleShortcut(event: KeyboardEvent): void {
    if (!this.live || !(event.ctrlKey || event.metaKey)) return;
    const key = event.key.toLowerCase();
    const target = event.target instanceof Element ? event.target : null;
    const inDocument = this.surface.contains(event.target as Node);
    // Ctrl+Alt+M is Word's New Comment; every other chord here is plain Ctrl.
    if (event.altKey) {
      if (key === "m" && inDocument) { event.preventDefault(); this.beginComment(); }
      return;
    }
    if (key === "f") { event.preventDefault(); this.openFindBar(false); return; }
    if (key === "h") { event.preventDefault(); this.openFindBar(true); return; }
    if (!inDocument) return;
    // Below here every chord edits the focused BLOCK, so a text box inside the surface — the
    // comment gutter's reply and edit boxes — owns its keys: Ctrl+Enter posts a reply there and
    // must not also drop a page break on the commented paragraph, nor Ctrl+E/L/R/J re-align it
    // mid-reply. Find and replace stay reachable from a comment box, as they are in Word.
    if (target?.closest(".docx-comment-gutter, textarea, input")) return;
    const align = (alignment: EditorAlignment) => {
      event.preventDefault();
      this.run(`align ${alignment}`, () => this.live!.setAlignment(alignment));
    };
    switch (key) {
      case "k": event.preventDefault(); this.openLinkPopover(); return;
      case "e": align("center"); return;
      case "l": align("left"); return;
      case "r": align("right"); return;
      case "j": align("justify"); return;
      case "]": event.preventDefault(); this.run("grow font", () => this.live!.adjustFontSize(1)); return;
      case "[": event.preventDefault(); this.run("shrink font", () => this.live!.adjustFontSize(-1)); return;
      case "enter": event.preventDefault(); this.run("page break", () => this.live!.pageBreakBefore(true)); return;
      default: return;
    }
  }

  // ── popovers ────────────────────────────────────────────────────────────────

  private positionPopover(popover: HTMLElement, anchor: HTMLElement): void {
    if (this.density !== "full") return;
    const rect = anchor.getBoundingClientRect();
    const host = this.element.getBoundingClientRect();
    popover.style.left = `${rect.left - host.left}px`;
    popover.style.top = `${rect.bottom - host.top + 5}px`;
  }

  private wirePicker(): void {
    const button = this.require("table");
    const picker = this.require("gridpicker");
    const cells = this.require("gridcells");

    const highlight = (rows: number, cols: number) => {
      for (const cell of Array.from(cells.children) as HTMLElement[]) {
        const on = Number(cell.dataset.r) < rows && Number(cell.dataset.c) < cols;
        if (on) cell.setAttribute("data-on", "");
        else cell.removeAttribute("data-on");
      }
      this.require("gridlabel").textContent = `${rows} × ${cols}`;
    };

    // `pointerover` covers mouse hover; touch users get the same feedback from the
    // press itself, since a tap fires pointerover immediately before pointerdown.
    cells.addEventListener("pointerover", (event) => {
      const cell = (event.target as HTMLElement).closest<HTMLElement>("[data-r]");
      if (cell) highlight(Number(cell.dataset.r) + 1, Number(cell.dataset.c) + 1);
    });
    cells.addEventListener("mousedown", (event) => {
      const cell = (event.target as HTMLElement).closest<HTMLElement>("[data-r]");
      if (!cell || !this.live) return;
      event.preventDefault();
      const rows = Number(cell.dataset.r) + 1;
      const cols = Number(cell.dataset.c) + 1;
      this.closePopovers();
      this.run(`table ${rows}×${cols}`, () =>
        this.live!.insertTable(rows, cols, {
          borderless: this.require<HTMLInputElement>("gridborderless").checked,
          cellAlignment: this.require<HTMLSelectElement>("gridalign").value as EditorAlignment,
        }));
    });

    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (!this.live) return;
      if (picker.hasAttribute("data-open")) {
        this.closePopovers();
        return;
      }
      this.closePopovers();
      highlight(0, 0);
      picker.setAttribute("data-open", "");
      this.positionPopover(picker, button);
    });
  }

  private wireLinkPopover(): void {
    const button = this.require("link");
    const pop = this.require("linkpop");
    const form = this.require<HTMLFormElement>("linkform");
    const url = this.require<HTMLInputElement>("linkurl");
    this.keepSelection(button);
    button.addEventListener("click", () => {
      if (pop.hasAttribute("data-open")) this.closePopovers();
      else this.openLinkPopover();
    });
    this.control("linkcancel")?.addEventListener("click", () => this.closePopovers());
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      const target = url.value.trim();
      if (!target || !this.live) return;
      this.closePopovers();
      let ok = false;
      this.run("link", () => { ok = this.live!.insertHyperlink(target); });
      if (!ok) this.setStatus("Select the text to link first");
    });
    url.addEventListener("keydown", (event) => {
      if (event.key === "Escape") { event.preventDefault(); this.closePopovers(); }
    });
  }

  private openLinkPopover(): void {
    if (!this.live) return;
    this.closePopovers();
    const pop = this.require("linkpop");
    const url = this.require<HTMLInputElement>("linkurl");
    const existing = this.live.hyperlinkAtCaret();
    url.value = existing?.target ?? "";
    pop.setAttribute("data-open", "");
    this.positionPopover(pop, this.require("link"));
    url.focus();
    url.select();
  }

  private closePopovers(): void {
    this.control("gridpicker")?.removeAttribute("data-open");
    this.control("linkpop")?.removeAttribute("data-open");
  }

  private maybeClosePopovers(event: MouseEvent): void {
    const target = event.target as Node;
    for (const [popName, buttonName] of [["gridpicker", "table"], ["linkpop", "link"]] as const) {
      const pop = this.control(popName);
      if (!pop?.hasAttribute("data-open")) continue;
      if (pop.contains(target) || this.control(buttonName)?.contains(target)) continue;
      pop.removeAttribute("data-open");
    }
  }

  // ── tabs ────────────────────────────────────────────────────────────────────

  selectTab(name: string): void {
    for (const tab of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-tab"))) {
      tab.setAttribute("aria-selected", String(tab.dataset.tab === name));
    }
    for (const panel of Array.from(this.element.querySelectorAll<HTMLElement>(".dxr-panel"))) {
      if (panel.dataset.panel === name) panel.setAttribute("data-active", "");
      else panel.removeAttribute("data-active");
    }
    if (name === "layout") { this.syncPageNumbering(); this.syncLayoutState(); }
    if (name === "review") this.syncReviewState();
    if (name === "headerfooter") this.syncHeaderFooterState();
    if (name === "view") this.syncZoom();
  }

  // ── selection-driven state ──────────────────────────────────────────────────

  private selectionElement(): HTMLElement | null {
    const selection = (this.element.ownerDocument ?? document).getSelection();
    const node = selection?.anchorNode ?? null;
    if (!node) return null;
    const el = node.nodeType === Node.ELEMENT_NODE ? (node as HTMLElement) : node.parentElement;
    // Ignore selections belonging to another editor (or another ribbon) on the page.
    return el && this.surface.contains(el) ? el : null;
  }

  private syncSelection(): void {
    if (!this.live || this.destroyed) return;
    const el = this.selectionElement();

    const state = this.live.queryFormatState();
    for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-cmd]"))) {
      button.classList.toggle("dxr-on", !!state[button.dataset.cmd as FormatKey]);
    }

    const doc = this.element.ownerDocument ?? document;
    if (el && typeof getComputedStyle === "function") {
      const cs = getComputedStyle(el);
      // Reflect the caret's font size, unless the field itself is being edited.
      const fontsize = this.control<HTMLInputElement>("fontsize");
      if (fontsize && doc.activeElement !== fontsize) {
        const px = parseFloat(cs.fontSize);
        if (px) fontsize.value = String(Math.round(px * 0.75 * 2) / 2);
      }
      const hex = toHex(cs.color);
      if (hex) {
        this.control("fontcolorbar")?.style.setProperty("--dxr-swatch", hex);
        const input = this.control<HTMLInputElement>("fontcolor");
        if (input && doc.activeElement !== input) input.value = hex;
      }
      this.control("smallcaps")?.setAttribute("aria-pressed", String(cs.fontVariant.includes("small-caps")));
      const align = cs.textAlign === "start" ? "left" : cs.textAlign === "end" ? "right" : cs.textAlign;
      for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-align]"))) {
        button.classList.toggle("dxr-on", button.dataset.align === align);
      }
    }
    const listFormat = this.live.listFormatAtCaret();
    for (const button of Array.from(this.element.querySelectorAll<HTMLElement>("button[data-list]"))) {
      const kind = button.dataset.list;
      button.classList.toggle(
        "dxr-on",
        !!listFormat && (kind === "bullet" ? listFormat.startsWith("bullet") : !listFormat.startsWith("bullet")),
      );
    }

    // Reveal the contextual Table tab only while the caret is inside a table.
    const inTable = !!el?.closest("table");
    const tableTab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="table"]');
    if (tableTab) {
      tableTab.hidden = !inTable;
      if (!inTable && tableTab.getAttribute("aria-selected") === "true") this.selectTab("home");
      if (inTable) {
        const row = el?.closest("tr");
        this.require<HTMLInputElement>("repeatheader").checked = !!row?.classList.toString().includes("header");
      }
    }
    this.refreshRailAnchor();
    this.scheduleStats();
  }

  /** Scope is read from where the block lives, which is exactly what the anchor encodes. */
  private scopeOf(el: HTMLElement): string {
    const band = el.closest<HTMLElement>("[data-hf-band]");
    if (band) return band.getAttribute("data-hf-band") === "header" ? "hdr" : "ftr";
    if (el.closest(".footnotes")) return "fn";
    if (el.closest(".endnotes")) return "en";
    return "body";
  }

  private kindOf(el: HTMLElement): string {
    if (/^H[1-6]$/.test(el.tagName)) return "h";
    if (el.tagName === "LI" || el.hasAttribute("data-list-marker") || el.querySelector("[data-list-marker]")) {
      return "li";
    }
    return "p";
  }

  /**
   * Block count and session handle — an O(blocks) query, so it runs only when
   * something could have changed them (a command, or a document opening), never on
   * the selectionchange path that fires per keystroke.
   */
  private refreshRailCounts(): void {
    const blocks = this.control("railBlocks");
    if (!blocks) return; // rail disabled
    blocks.textContent = this.live ? String(this.surface.querySelectorAll("[data-anchor]").length) : "—";
    const session = this.control("railSession");
    if (session) session.textContent = this.live ? `#${this.live.sessionHandle}` : "—";
  }

  /** The focused anchor — cheap, and the only part the caret can change. */
  private refreshRailAnchor(): void {
    const anchorEl = this.control("railAnchor");
    if (!anchorEl) return;
    const block = this.selectionElement()?.closest<HTMLElement>("[data-anchor]") ?? null;
    if (!block) {
      anchorEl.textContent = this.live ? "none" : "—";
      this.lastAnchorText = "";
      return;
    }
    const unid = block.getAttribute("data-anchor") ?? "";
    const text = `${this.kindOf(block)}:${this.scopeOf(block)}:${unid.slice(0, 10)}…`;
    if (text === this.lastAnchorText) return;
    anchorEl.textContent = text;
    this.lastAnchorText = text;
    anchorEl.classList.remove("dxr-flash");
    void anchorEl.offsetWidth;
    anchorEl.classList.add("dxr-flash");
  }

  /** Page position and word count are O(document); debounced off the hot paths. */
  private scheduleStats(): void {
    if (this.statsTimer != null) clearTimeout(this.statsTimer);
    this.statsTimer = setTimeout(() => {
      this.statsTimer = null;
      this.refreshStats();
    }, 150);
  }

  private refreshStats(): void {
    if (!this.live || this.destroyed) return;
    const pageinfo = this.control("pageinfo");
    if (pageinfo) {
      const info = this.live.pageInfo();
      pageinfo.textContent = info ? `Page ${info.page} of ${info.total}` : "Continuous";
    }
    const words = this.control("wordcount");
    if (words) {
      const n = this.live.wordCount();
      words.textContent = `${n.toLocaleString()} word${n === 1 ? "" : "s"}`;
    }
  }

  /** Font families the document's rendered text actually uses, most common first. */
  private documentFonts(): string[] {
    const counts = new Map<string, number>();
    for (const el of Array.from(this.surface.querySelectorAll<HTMLElement>("[data-anchor] [style*='font-family'], [data-anchor][style*='font-family']"))) {
      const family = el.style.fontFamily.split(",")[0]?.trim().replace(/^['"]|['"]$/g, "");
      if (!family || /^(serif|sans-serif|monospace)$/i.test(family)) continue;
      counts.set(family, (counts.get(family) ?? 0) + 1);
    }
    return Array.from(counts.entries()).sort((a, b) => b[1] - a[1]).map(([name]) => name).slice(0, 12);
  }

  // ── loading overlay ─────────────────────────────────────────────────────────

  private createLoaderController(): RibbonLoader {
    const overlay = this.control("loader");
    if (!overlay) {
      const noop = () => {};
      return { show: noop, stage: noop, progress: noop, done: noop, fail: noop };
    }

    const options = this.loaderOptions ?? {};
    const eyebrow = this.control("loaderEyebrow");
    if (eyebrow && options.eyebrow) eyebrow.textContent = options.eyebrow;
    const meta = this.control("loaderMeta");
    if (meta && options.meta) meta.textContent = options.meta;
    this.control("loaderRetry")?.addEventListener("click", () => {
      if (options.onRetry) options.onRetry();
      else location.reload();
    });
    if (this.features.length === 0) this.control("loaderAd")?.remove();

    const setProgress = (percent: number, label?: string) => {
      const bar = this.control("loaderBar");
      if (bar) bar.style.width = `${Math.max(0, Math.min(100, percent))}%`;
      if (label) {
        const el = this.control("loaderLabel");
        if (el) el.textContent = label;
      }
    };

    return {
      show: () => {
        overlay.hidden = false;
        overlay.removeAttribute("data-done");
        overlay.removeAttribute("data-error");
        this.setState("loading");
        this.showFeature(0);
        this.startRotation();
        this.loaderStage(0);
      },
      stage: (step) => this.loaderStage(step),
      progress: setProgress,
      done: () => {
        this.stopRotation();
        this.setState("ready");
        // Two-step so the fade actually plays before the overlay leaves the layout.
        setTimeout(() => overlay.setAttribute("data-done", ""), 420);
        setTimeout(() => {
          overlay.hidden = true;
        }, 1050);
      },
      fail: (error) => {
        this.stopRotation();
        this.setState("error");
        overlay.hidden = false;
        overlay.setAttribute("data-error", "");
        overlay.removeAttribute("data-done");
        const message = error instanceof Error ? error.message : String(error);
        const title = this.control("loaderTitle");
        if (title) title.textContent = "The local engine did not start";
        const copy = this.control("loaderCopy");
        if (copy) copy.textContent = message.slice(0, 220);
        setProgress(100, "Load failed");
        this.setStatus(message.slice(0, 180));
      },
    };
  }

  private loaderStage(step: number | Partial<RibbonLoaderStage>): void {
    const stage = typeof step === "number" ? this.stages[Math.max(0, Math.min(this.stages.length - 1, step))] : step;
    if (!stage) return;
    const title = this.control("loaderTitle");
    if (title && stage.title) title.textContent = stage.title;
    const copy = this.control("loaderCopy");
    if (copy && stage.copy) copy.textContent = stage.copy;
    const bar = this.control("loaderBar");
    if (bar && stage.progress != null) bar.style.width = `${stage.progress}%`;
    const label = this.control("loaderLabel");
    if (label && stage.label) label.textContent = stage.label;
  }

  private showFeature(index: number): void {
    const feature = this.features[index];
    if (!feature) return;
    const card = this.control("loaderAd");
    const number = this.control("loaderNumber");
    const title = this.control("loaderAdTitle");
    const copy = this.control("loaderAdCopy");
    if (number) number.textContent = feature.number;
    if (title) title.textContent = feature.title;
    if (copy) copy.textContent = feature.copy;
    if (card) {
      card.removeAttribute("data-swap");
      void card.offsetWidth;
      card.setAttribute("data-swap", "");
    }
  }

  private startRotation(): void {
    this.stopRotation();
    if (this.features.length < 2) return;
    const every = this.loaderOptions?.rotateMs ?? 1750;
    this.featureTimer = setInterval(() => {
      this.featureIndex = (this.featureIndex + 1) % this.features.length;
      this.showFeature(this.featureIndex);
    }, every);
  }

  private stopRotation(): void {
    if (this.featureTimer == null) return;
    clearInterval(this.featureTimer);
    this.featureTimer = null;
  }
}

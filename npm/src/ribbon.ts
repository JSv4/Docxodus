/**
 * The ribbon surface — the full editor UI as a mountable module.
 *
 * `DocxEditor` (editor.ts) is the document engine: it owns the live `DocxSession`,
 * the block wiring and every command. It has deliberately no chrome. This module is
 * the chrome — the tabbed ribbon, the anchor rail, the table picker and the loading
 * overlay — wired onto exactly that command surface and nothing else.
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
  FormatKey,
} from "./editor.js";
import type { NumberFormat } from "./types.js";
import {
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
  /** Show the anchor rail (full chrome only). Default true. */
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
  /** Activate a ribbon tab by name ("home" | "insert" | "layout" | "table"). */
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
  "fontsize", "fontsizes", "fontfamily", "style", "delblock",
  "table", "hr", "hrThick", "hrDouble", "rulepos", "hrClear", "footnote", "endnote",
  "paginated", "headerfooter", "pgfmt", "pgstart", "pgclear", "pagenum", "totalpages",
  "railAnchor", "railBlocks", "railSession", "railOp",
  "gridpicker", "gridcells", "gridlabel", "gridalign", "gridborderless",
  "editor",
  "loader", "loaderEyebrow", "loaderTitle", "loaderCopy", "loaderAd", "loaderNumber",
  "loaderAdTitle", "loaderAdCopy", "loaderBar", "loaderLabel", "loaderMeta", "loaderRetry",
];

const GRID_ROWS = 8;
const GRID_COLS = 10;

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
  private destroyed = false;

  private featureTimer: ReturnType<typeof setInterval> | null = null;
  private featureIndex = 0;
  private resizeObserver: ResizeObserver | null = null;
  private lastAnchorText = "";
  private selectionFrame: number | null = null;
  private readonly onSelectionChange: () => void;
  private readonly onDocumentMouseDown: (event: MouseEvent) => void;

  constructor(container: HTMLElement, options: RibbonOptions) {
    this.options = options;
    this.exports = options.exports ?? null;
    this.documentName = options.documentName ?? "untitled.docx";
    this.chromeMode = options.chrome ?? "auto";
    this.headerFooter = options.headerFooter ?? false;
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
    this.wire();
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
    this.onDocumentMouseDown = (event) => this.maybeClosePicker(event);
    doc.addEventListener("mousedown", this.onDocumentMouseDown);

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
    (this.require<HTMLInputElement>("paginated")).checked = this.options.paginated ?? false;
    (this.require<HTMLInputElement>("headerfooter")).checked = this.headerFooter;
    this.surface.dataset.view = this.options.paginated ? "paginated" : "continuous";
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
    // The picker's absolute position is meaningless after a density flip.
    this.closePicker();
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

    const started = performance.now();
    this.live = DocxEditor.open(this.surface, bytes, this.exports, {
      cssPrefix: this.options.cssPrefix,
      fabricateClasses: this.options.fabricateClasses,
      editable: this.options.editable,
      scale: this.options.scale,
      columnWidth: this.options.columnWidth,
      fitToWidth: this.options.fitToWidth,
      onEdit: this.options.onEdit,
      onMove: this.options.onMove,
      paginated,
      headerFooter: this.headerFooter,
      blockDrag: this.options.blockDrag ?? true,
      trackedChanges: this.options.trackedChanges,
      revisionAuthor: this.options.revisionAuthor,
    });
    this.require<HTMLButtonElement>("save").disabled = false;
    this.require("ribbon").setAttribute("aria-disabled", "false");
    this.setState("ready");
    this.setStatus(`Rendered in ${Math.round(performance.now() - started)} ms`);
    this.syncPageNumbering();
    this.refreshRailCounts();
    this.refreshRailAnchor();
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
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    if (this.selectionFrame != null) cancelAnimationFrame(this.selectionFrame);
    this.selectionFrame = null;
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
      };
      ops[b.dataset.tt ?? ""]?.();
    });

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
    ];
    for (const [name, label, fn] of simple) {
      const el = this.control(name);
      if (!el) continue;
      this.keepSelection(el);
      el.addEventListener("click", () => this.run(label, fn));
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

    const fontfamily = this.require<HTMLSelectElement>("fontfamily");
    fontfamily.addEventListener("change", () => {
      if (this.live && fontfamily.value) this.run("font family", () => this.live!.setFontFamily(fontfamily.value));
      fontfamily.value = "";
    });

    const style = this.require<HTMLSelectElement>("style");
    style.addEventListener("change", () => {
      if (this.live && style.value) this.run("style", () => this.live!.setParagraphStyle(style.value));
      style.value = "";
    });

    // Comment authoring (issue #580): the button comments the live selection (whole block
    // when collapsed) with the typed text, attributed to the configured revision author.
    const commentButton = this.control("comment");
    if (commentButton) {
      this.keepSelection(commentButton);
      commentButton.addEventListener("click", () => {
        if (!this.live) return;
        const input = this.control<HTMLInputElement>("commenttext");
        const text = input?.value.trim() || "New comment.";
        this.run("comment", () =>
          this.live!.addComment(text, this.options.revisionAuthor ?? "Reviewer"));
        if (input) input.value = "";
        this.renderThreads();
      });
    }

    this.wireFileActions();
    this.wireLayout();
    this.wirePicker();
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
  }

  private wireLayout(): void {
    // Pagination toggles re-render the LIVE session, so edits survive the switch.
    const paginated = this.require<HTMLInputElement>("paginated");
    paginated.addEventListener("change", () => {
      this.surface.dataset.view = paginated.checked ? "paginated" : "continuous";
      if (!this.live) return;
      this.run(paginated.checked ? "page view" : "continuous view", () =>
        this.live!.setPaginated(paginated.checked));
    });

    // The band region is chosen at open time, so toggling it re-opens the LIVE
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
  }

  /** The bands own the same setting and read the live session, so both stay in step. */
  private syncPageNumbering(): void {
    if (!this.live) return;
    const numbering = this.live.pageNumbering() ?? {};
    this.require<HTMLSelectElement>("pgfmt").value = numbering.format ?? "";
    this.require<HTMLInputElement>("pgstart").value = numbering.start != null ? String(numbering.start) : "";
  }

  // ── table size picker ───────────────────────────────────────────────────────

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
      this.closePicker();
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
        this.closePicker();
        return;
      }
      highlight(0, 0);
      picker.setAttribute("data-open", "");
      // Compact chrome docks the picker to the viewport bottom (see the stylesheet),
      // so only the roomy layout needs it positioned under its button.
      if (this.density === "full") {
        const rect = button.getBoundingClientRect();
        const host = this.element.getBoundingClientRect();
        picker.style.left = `${rect.left - host.left}px`;
        picker.style.top = `${rect.bottom - host.top + 5}px`;
      }
    });
  }

  private closePicker(): void {
    this.control("gridpicker")?.removeAttribute("data-open");
  }

  private maybeClosePicker(event: MouseEvent): void {
    const picker = this.control("gridpicker");
    if (!picker?.hasAttribute("data-open")) return;
    const target = event.target as Node;
    if (picker.contains(target) || this.control("table")?.contains(target)) return;
    this.closePicker();
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
    if (name === "layout") this.syncPageNumbering();
    if (name === "review") this.renderThreads();
  }

  /** Re-render the Review tab's thread list from session truth ({@link DocxEditor.listComments}):
   *  one row per comment, replies indented under their thread root, resolve/reopen per root. */
  private renderThreads(): void {
    const host = this.control("threads");
    if (!host || !this.live) return;
    const doc = this.element.ownerDocument ?? document;
    host.textContent = "";
    const comments = this.live.listComments();
    if (comments.length === 0) {
      const empty = doc.createElement("span");
      empty.className = "dxr-note";
      empty.textContent = "No comments yet.";
      host.appendChild(empty);
      return;
    }
    for (const entry of comments) {
      const row = doc.createElement("div");
      row.className = "dxr-thread";
      row.setAttribute("data-thread", entry.anchorId);
      if (entry.resolved) row.setAttribute("data-resolved", "");
      if (entry.parentAnchorId) row.setAttribute("data-reply", "");

      const author = doc.createElement("span");
      author.className = "dxr-tauthor";
      author.textContent = entry.author || "—";
      row.appendChild(author);

      const text = doc.createElement("span");
      text.className = "dxr-ttext";
      text.textContent = entry.text;
      text.title = entry.text;
      row.appendChild(text);

      // Resolution is a THREAD state: Word keys it on the root, so replies get no toggle.
      if (!entry.parentAnchorId) {
        const toggle = doc.createElement("button");
        toggle.type = "button";
        toggle.textContent = entry.resolved ? "Reopen" : "Resolve";
        toggle.title = entry.resolved
          ? "Reopen this comment thread"
          : "Mark this comment thread resolved";
        toggle.addEventListener("click", () => {
          if (!this.live) return;
          this.run(entry.resolved ? "reopen comment" : "resolve comment", () =>
            this.live!.setCommentResolved(entry.anchorId, !entry.resolved));
          this.renderThreads();
        });
        row.appendChild(toggle);
      }
      host.appendChild(row);
    }
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

    // Reflect the caret's font size, unless the field itself is being edited.
    const fontsize = this.control<HTMLInputElement>("fontsize");
    if (el && fontsize && (this.element.ownerDocument ?? document).activeElement !== fontsize) {
      const px = parseFloat(getComputedStyle(el).fontSize);
      if (px) fontsize.value = String(Math.round(px * 0.75 * 2) / 2);
    }

    // Reveal the contextual Table tab only while the caret is inside a table.
    const inTable = !!el?.closest("table");
    const tableTab = this.element.querySelector<HTMLElement>('.dxr-tab[data-tab="table"]');
    if (tableTab) {
      tableTab.hidden = !inTable;
      if (!inTable && tableTab.getAttribute("aria-selected") === "true") this.selectTab("home");
    }
    this.refreshRailAnchor();
  }

  /** Scope is read from where the block lives, which is exactly what the anchor encodes. */
  private scopeOf(el: HTMLElement): string {
    const band = el.closest<HTMLElement>(".docx-hf-band");
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

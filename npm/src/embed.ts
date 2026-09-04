/**
 * Embeddable viewer/editor entry point — the CDN story.
 *
 * This module packages the whole stack behind two one-call factories so a page
 * can embed a DOCX viewer or editor with a single script tag and no build step:
 *
 * ```html
 * <div id="doc"></div>
 * <script type="module">
 *   import { createViewer } from "https://cdn.jsdelivr.net/npm/docxodus@12.0.0/dist/embed.bundle.js";
 *   await createViewer("#doc", "./contract.docx");
 * </script>
 * ```
 *
 * It ships in three shapes (see package.json build scripts):
 *  - `dist/embed.js`         — plain ESM with relative imports, for bundler users
 *                              (`import { createEditor } from "docxodus/embed"`).
 *                              Shares module state with the main `docxodus` entry.
 *  - `dist/embed.bundle.js`  — self-contained ESM bundle for CDN `<script type="module">`.
 *  - `dist/embed.iife.js`    — classic-script bundle exposing `window.Docxodus`.
 *
 * WASM asset resolution: the .NET runtime files live in `wasm/` next to this
 * file (`dist/wasm/` in the published package), which auto-resolves through
 * `import.meta.url` in the ESM shapes. The IIFE shape has no `import.meta`, so
 * it falls back to `document.currentScript.src` captured at load time. An
 * explicit `wasmBasePath` option always wins.
 */

import {
  initialize,
  getWasmExports,
  convertDocxToHtml,
} from "./index.js";
import { DocxEditor } from "./editor.js";
import type { DocxEditorExports, DocxEditorOptions } from "./editor.js";
import { mountRibbon } from "./ribbon.js";
import type { RibbonEditor, RibbonOptions } from "./ribbon.js";
import type { ConversionOptions } from "./types.js";

// Everything the main entry exports is re-exported so the CDN bundle is a
// one-stop surface (convert, compare, diff, sessions, annotations, ...).
export * from "./index.js";
export { DocxEditor } from "./editor.js";
export type { DocxEditorExports, DocxEditorOptions } from "./editor.js";

/**
 * A document to load: raw bytes, a Blob/File (e.g. from an <input type="file">),
 * or a URL string to fetch (must be same-origin or CORS-readable).
 */
export type DocumentSource = Uint8Array | ArrayBuffer | Blob | string;

/**
 * `document.currentScript.src`, captured synchronously at module evaluation.
 * Only meaningful for the classic-script (IIFE) shape — for module scripts
 * `currentScript` is null and `import.meta.url` is used instead.
 */
const scriptBase: string = (() => {
  try {
    const el =
      typeof document !== "undefined"
        ? (document.currentScript as HTMLScriptElement | null)
        : null;
    if (el?.src) return el.src.substring(0, el.src.lastIndexOf("/") + 1);
  } catch {
    /* non-browser environment */
  }
  return "";
})();

/**
 * The directory this bundle was loaded from: `import.meta.url` for the ESM
 * shapes (esbuild shims it to `{}` in the IIFE build, so the probe fails
 * there), falling back to the captured `currentScript` base for the IIFE.
 */
function moduleBaseDir(): string {
  try {
    const url = import.meta.url;
    if (typeof url === "string" && url.length > 0) {
      return url.substring(0, url.lastIndexOf("/") + 1);
    }
  } catch {
    /* shimmed or unavailable */
  }
  return scriptBase;
}

/**
 * Initialize the WASM runtime for an embed call.
 *
 * An explicit `wasmBasePath` wins. Otherwise probe the two layouts consumers
 * actually have: `wasm/` next to this bundle (the published package layout —
 * `dist/embed.bundle.js` + `dist/wasm/`), then the assets directly next to the
 * bundle (a webroot that serves the wasm directory itself). A failed probe is
 * retryable because `initialize()` clears its cached promise on rejection.
 */
async function ensureWasm(wasmBasePath?: string): Promise<void> {
  if (wasmBasePath) return initialize(wasmBasePath);
  const dir = moduleBaseDir();
  if (!dir) return initialize(); // last resort: index.ts's own auto-detection
  try {
    await initialize(dir + "wasm/");
  } catch {
    await initialize(dir);
  }
}

function resolveContainer(container: string | HTMLElement): HTMLElement {
  if (typeof container !== "string") return container;
  const el = document.querySelector<HTMLElement>(container);
  if (!el) throw new Error(`Docxodus embed: no element matches "${container}"`);
  return el;
}

// A style element nested in a normal DOM container is still document-global.
// The converter emits selectors such as `body` and `span`, so inserting its
// stylesheet verbatim would restyle the host page. Embed factories mount into a
// private inner root and prefix every ordinary selector with that root. This
// keeps regular DOM/querySelector/contenteditable behavior (unlike Shadow DOM,
// whose selection boundary is awkward for DocxEditor) without leaking CSS.
const EMBED_ROOT_ATTR = "data-docxodus-embed-root";
const SCOPED_STYLE_ATTR = "data-docxodus-scoped-style";
let nextEmbedRootId = 0;

interface ScopedMount {
  root: HTMLElement;
  selector: string;
}

function createScopedMount(container: HTMLElement): ScopedMount {
  let id: string;
  do {
    id = `d${++nextEmbedRootId}`;
  } while (document.querySelector(`[${EMBED_ROOT_ATTR}="${id}"]`));

  const root = document.createElement("div");
  root.setAttribute(EMBED_ROOT_ATTR, id);
  container.replaceChildren(root);
  return { root, selector: `[${EMBED_ROOT_ATTR}="${id}"]` };
}

/** Split a selector list only at top-level commas (not commas inside :is(), attributes, etc.). */
function splitSelectorList(selectorText: string): string[] {
  const selectors: string[] = [];
  let start = 0;
  let parentheses = 0;
  let brackets = 0;
  let quote = "";
  let escaped = false;

  for (let i = 0; i < selectorText.length; i++) {
    const ch = selectorText[i];
    if (escaped) {
      escaped = false;
      continue;
    }
    if (ch === "\\") {
      escaped = true;
      continue;
    }
    if (quote) {
      if (ch === quote) quote = "";
      continue;
    }
    if (ch === '"' || ch === "'") {
      quote = ch;
      continue;
    }
    if (ch === "(") parentheses++;
    else if (ch === ")") parentheses--;
    else if (ch === "[") brackets++;
    else if (ch === "]") brackets--;
    else if (ch === "," && parentheses === 0 && brackets === 0) {
      selectors.push(selectorText.slice(start, i).trim());
      start = i + 1;
    }
  }
  selectors.push(selectorText.slice(start).trim());
  return selectors.filter(Boolean);
}

function scopeSelector(selector: string, rootSelector: string): string {
  let remaining = selector.trim();
  let consumedRoot = false;
  let separated = false;

  // The converter's full-document CSS starts at body/:root. In an embed, the
  // private mount root is the equivalent document root. Consume repeated roots
  // too (`html body ...`) before attaching the remainder.
  for (;;) {
    const match = /^(?:html|body|:root)(?=$|[\s>+~.#[:])/.exec(remaining);
    if (!match) break;
    consumedRoot = true;
    remaining = remaining.slice(match[0].length);
    const whitespace = /^\s+/.exec(remaining);
    if (whitespace) {
      separated = true;
      remaining = remaining.slice(whitespace[0].length);
    } else {
      separated = false;
      break;
    }
  }

  if (!consumedRoot) return `${rootSelector} ${remaining}`;
  if (!remaining) return rootSelector;
  if (separated || /^[>+~]/.test(remaining)) return `${rootSelector} ${remaining}`;
  return `${rootSelector}${remaining}`;
}

type MutableCssRules = {
  readonly cssRules: CSSRuleList;
  deleteRule(index: number): void;
};

function scopeCssRules(container: MutableCssRules, rootSelector: string): void {
  // Walk backwards because document-level @page and @import rules are removed.
  // Neither can be selector-scoped; keeping them would let print settings or an
  // imported stylesheet escape into the host page.
  for (let i = container.cssRules.length - 1; i >= 0; i--) {
    const rule = container.cssRules[i];
    if (rule.type === CSSRule.PAGE_RULE || rule.type === CSSRule.IMPORT_RULE) {
      container.deleteRule(i);
      continue;
    }
    if (rule.type === CSSRule.STYLE_RULE) {
      const styleRule = rule as CSSStyleRule;
      styleRule.selectorText = splitSelectorList(styleRule.selectorText)
        .map((selector) => scopeSelector(selector, rootSelector))
        .join(", ");
      continue;
    }
    const grouping = rule as CSSRule & Partial<MutableCssRules>;
    if (grouping.cssRules && typeof grouping.deleteRule === "function") {
      scopeCssRules(grouping as MutableCssRules, rootSelector);
    }
  }
}

function scopeStyleElement(style: HTMLStyleElement, rootSelector: string): void {
  if (style.hasAttribute(SCOPED_STYLE_ATTR)) return;

  // Parse through the browser's CSSOM rather than regex-rewriting CSS. `media`
  // keeps this temporary parser stylesheet inert while it is attached, so the
  // unscoped rules never affect layout even for a single frame.
  const parser = document.createElement("style");
  parser.media = "not all";
  parser.textContent = style.textContent ?? "";
  (document.head ?? document.documentElement).appendChild(parser);
  try {
    const sheet = parser.sheet as CSSStyleSheet | null;
    if (!sheet) throw new Error("browser did not expose a CSS stylesheet");
    scopeCssRules(sheet, rootSelector);
    style.setAttribute(SCOPED_STYLE_ATTR, "");
    style.textContent = Array.from(sheet.cssRules, (rule) => rule.cssText).join("\n");
  } catch (error) {
    throw new Error(`Docxodus embed: failed to scope document CSS: ${String(error)}`);
  } finally {
    parser.remove();
  }
}

function scopeDocumentStyles(fullHtml: string, rootSelector: string): string {
  const parsed = new DOMParser().parseFromString(fullHtml, "text/html");
  parsed.querySelectorAll<HTMLStyleElement>("style")
    .forEach((style) => scopeStyleElement(style, rootSelector));
  return `<!doctype html>\n${parsed.documentElement.outerHTML}`;
}

/** Scope initial and remount renders synchronously, before DocxEditor inserts their HTML. */
function createScopedEditorExports(
  exports: DocxEditorExports,
  rootSelector: string,
): DocxEditorExports {
  const convert = exports.DocumentConverter.ConvertDocxToHtmlComplete;
  const bridge = { ...exports.DocxSessionBridge };
  // EVERY whole-document render has to be scoped, not just the first one the editor happened to
  // use: the converter's stylesheet contains `body`/`span` rules, so one unwrapped path is enough
  // to restyle the host page. `RenderHtmlForReview` is the revision-view twin of `RenderHtml` and
  // is what remount prefers when it exists.
  const renderHtml = bridge.RenderHtml;
  if (renderHtml) {
    bridge.RenderHtml = (...args) => scopeDocumentStyles(renderHtml(...args), rootSelector);
  }
  const renderHtmlForReview = bridge.RenderHtmlForReview;
  if (renderHtmlForReview) {
    bridge.RenderHtmlForReview = (...args) =>
      scopeDocumentStyles(renderHtmlForReview(...args), rootSelector);
  }
  // The editor's first paint prefers `RenderEditorHtml` when the bundle has it (it carries the
  // comment markup the plain `RenderHtml` omits). It is a whole-document render too, so its
  // stylesheet's `body`/`span` rules restyle the host page unless it is scoped like the others.
  // The per-block editor renders (`RenderEditorBlockHtml` / `RenderEditorBlocksHtml`) return body
  // fragments without a stylesheet, so they need no scoping.
  const renderEditorHtml = bridge.RenderEditorHtml;
  if (renderEditorHtml) {
    bridge.RenderEditorHtml = (...args) =>
      scopeDocumentStyles(renderEditorHtml(...args), rootSelector);
  }
  return {
    DocxSessionBridge: bridge,
    DocumentConverter: {
      ConvertDocxToHtmlComplete: (...args) =>
        scopeDocumentStyles(convert(...args), rootSelector),
    },
  };
}

/** Normalize any DocumentSource to bytes. */
export async function toDocumentBytes(source: DocumentSource): Promise<Uint8Array> {
  if (source instanceof Uint8Array) return source;
  if (source instanceof ArrayBuffer) return new Uint8Array(source);
  if (typeof Blob !== "undefined" && source instanceof Blob) {
    return new Uint8Array(await source.arrayBuffer());
  }
  if (typeof source === "string") {
    const response = await fetch(source);
    if (!response.ok) {
      throw new Error(
        `Docxodus embed: failed to fetch document "${source}" (HTTP ${response.status})`,
      );
    }
    return new Uint8Array(await response.arrayBuffer());
  }
  throw new Error("Docxodus embed: unsupported document source");
}

export interface DocxViewerOptions extends ConversionOptions {
  /** Explicit URL of the wasm assets directory; omit to auto-detect. */
  wasmBasePath?: string;
}

export interface DocxViewer {
  /** The container the document was rendered into. */
  readonly element: HTMLElement;
  /** The full HTML document string the converter produced. */
  readonly html: string;
  /** Re-render a (new) document into the same container. */
  reload(source: DocumentSource, options?: ConversionOptions): Promise<void>;
  /** Empty the container. */
  destroy(): void;
}

/**
 * Render a read-only DOCX viewer into `container`.
 *
 * Injects the converter's stylesheet plus the document body into the container
 * (the same mounting the editor uses), so the page's own styles are untouched.
 * Footnotes/endnotes render by default — they are document content; override
 * any conversion option via `options`.
 */
export async function createViewer(
  container: string | HTMLElement,
  source: DocumentSource,
  options: DocxViewerOptions = {},
): Promise<DocxViewer> {
  const el = resolveContainer(container);
  const { wasmBasePath, ...conversion } = options;
  await ensureWasm(wasmBasePath);
  const mount = createScopedMount(el);

  let lastHtml = "";
  const render = async (src: DocumentSource, opts: ConversionOptions) => {
    const bytes = await toDocumentBytes(src);
    lastHtml = await convertDocxToHtml(bytes, {
      renderFootnotesAndEndnotes: true,
      ...opts,
    });
    const parsed = new DOMParser().parseFromString(lastHtml, "text/html");
    const fragment = document.createDocumentFragment();
    parsed.querySelectorAll("style").forEach((sourceStyle) => {
      const style = document.createElement("style");
      style.textContent = sourceStyle.textContent;
      scopeStyleElement(style, mount.selector);
      fragment.appendChild(style);
    });
    const body = document.createElement("template");
    body.innerHTML = parsed.body.innerHTML;
    fragment.appendChild(body.content);
    mount.root.replaceChildren(fragment);
  };

  await render(source, conversion);
  return {
    element: el,
    get html() {
      return lastHtml;
    },
    reload: (src, opts) => render(src, { ...conversion, ...opts }),
    destroy: () => {
      el.innerHTML = "";
    },
  };
}

export interface CreateEditorOptions extends DocxEditorOptions {
  /** Explicit URL of the wasm assets directory; omit to auto-detect. */
  wasmBasePath?: string;
}

/**
 * Open an editable DOCX editor in `container`.
 *
 * With a `source`, opens that document; without one, opens a blank "New
 * document". Returns the `DocxEditor` instance — `save()` for the edited
 * bytes, `close()` to release the WASM session, plus the full command surface
 * (formatting, tables, notes, headers/footers, undo/redo).
 */
export async function createEditor(
  container: string | HTMLElement,
  source?: DocumentSource | null,
  options: CreateEditorOptions = {},
): Promise<DocxEditor> {
  const el = resolveContainer(container);
  const { wasmBasePath, ...editorOptions } = options;
  await ensureWasm(wasmBasePath);
  const mount = createScopedMount(el);
  // The runtime object is the real bridge; DocxEditorExports is the editor's
  // narrower view of it, so the cast is safe by construction.
  const exports = createScopedEditorExports(
    getWasmExports() as unknown as DocxEditorExports,
    mount.selector,
  );
  try {
    return source == null
      ? DocxEditor.openBlank(mount.root, exports, editorOptions)
      : DocxEditor.open(mount.root, await toDocumentBytes(source), exports, editorOptions);
  } catch (error) {
    el.replaceChildren();
    throw error;
  }
}

export interface CreateRibbonEditorOptions extends RibbonOptions {
  /** Explicit URL of the wasm assets directory; omit to auto-detect. */
  wasmBasePath?: string;
}

/** Best-effort file name from a URL source, so the title bar says something useful. */
function nameFromSource(source: DocumentSource | null | undefined): string | undefined {
  if (typeof source !== "string") return undefined;
  try {
    const path = new URL(source, typeof location === "undefined" ? undefined : location.href).pathname;
    return decodeURIComponent(path.split("/").pop() ?? "") || undefined;
  } catch {
    return source.split("/").pop() || undefined;
  }
}

/**
 * Open the FULL editor surface — ribbon chrome, anchor rail and all — in one call.
 *
 * `createEditor` gives you a bare editable document and leaves the UI to you.
 * This gives you the whole instrument, which is what the shipped demo pages use:
 *
 * ```html
 * <div id="app" style="height:100dvh"></div>
 * <script type="module">
 *   import { createRibbonEditor } from "https://cdn.jsdelivr.net/npm/docxodus/dist/embed.bundle.js";
 *   await createRibbonEditor("#app", "./contract.docx");
 * </script>
 * ```
 *
 * The chrome (and its loading overlay) paint immediately; the .NET runtime streams
 * behind them, and the overlay narrates that wait instead of hiding it. Density
 * follows the CONTAINER's width, so the same call serves a full-page editor and a
 * narrow embedded panel.
 */
export async function createRibbonEditor(
  container: string | HTMLElement,
  source?: DocumentSource | null,
  options: CreateRibbonEditorOptions = {},
): Promise<RibbonEditor> {
  const el = resolveContainer(container);
  const { wasmBasePath, ...ribbonOptions } = options;
  const mount = createScopedMount(el);
  // The CSS-scoping wrapper sits between the caller's element and the surface, so
  // it has to pass the container's height through — otherwise a full-height mount
  // (`height:100dvh`) collapses to content height and the surface never scrolls.
  mount.root.style.height = "100%";
  mount.root.style.minHeight = "0";
  const ribbon = mountRibbon(mount.root, {
    documentName: ribbonOptions.documentName ?? nameFromSource(source),
    // A drop-in embed sits in the middle of someone's page, so it carries its own
    // boundary (rounded card + shadow). Full-bleed hosts pass frame: "flush".
    frame: "card",
    ...ribbonOptions,
    // Exports arrive after the runtime boots; the loader covers that gap.
    exports: undefined,
  });

  try {
    ribbon.loader.stage(0);
    await ensureWasm(wasmBasePath);
    // The document renders inside the ribbon's surface, so that — not the mount
    // root — is the scope the converter's document CSS must be confined to.
    ribbon.setExports(
      createScopedEditorExports(
        getWasmExports() as unknown as DocxEditorExports,
        `${mount.selector} [data-dxr-surface]`,
      ),
    );

    ribbon.loader.stage(1);
    const bytes = source == null ? null : await toDocumentBytes(source);

    ribbon.loader.stage(2);
    if (bytes == null) ribbon.openBlank(ribbonOptions.documentName);
    else ribbon.open(bytes, ribbonOptions.documentName ?? nameFromSource(source));

    ribbon.loader.stage(3);
    ribbon.loader.done();
    return ribbon;
  } catch (error) {
    ribbon.loader.fail(error);
    throw error;
  }
}

/**
 * Embeddable viewer/editor entry point — the CDN story.
 *
 * This module packages the whole stack behind two one-call factories so a page
 * can embed a DOCX viewer or editor with a single script tag and no build step:
 *
 * ```html
 * <div id="doc"></div>
 * <script type="module">
 *   import { createViewer } from "https://cdn.jsdelivr.net/npm/docxodus@9/dist/embed.bundle.js";
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

  let lastHtml = "";
  const render = async (src: DocumentSource, opts: ConversionOptions) => {
    const bytes = await toDocumentBytes(src);
    lastHtml = await convertDocxToHtml(bytes, {
      renderFootnotesAndEndnotes: true,
      ...opts,
    });
    const parsed = new DOMParser().parseFromString(lastHtml, "text/html");
    const styles = Array.from(parsed.querySelectorAll("style"))
      .map((s) => s.outerHTML)
      .join("");
    el.innerHTML = styles + parsed.body.innerHTML;
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
  // The runtime object is the real bridge; DocxEditorExports is the editor's
  // narrower view of it, so the cast is safe by construction.
  const exports = getWasmExports() as unknown as DocxEditorExports;
  if (source == null) return DocxEditor.openBlank(el, exports, editorOptions);
  const bytes = await toDocumentBytes(source);
  return DocxEditor.open(el, bytes, exports, editorOptions);
}

/**
 * Deterministic browser readiness primitives shared by standalone HTML
 * materialization and the final document reopened for Chromium printing.
 */

import { documentFontReadinessCandidates } from "./font-runtime.js";

export type PrintReadinessPhase =
  | "font_loading"
  | "image_decoding"
  | "chart_svg_materialization"
  | "page_tree_stability";

export interface PrintReadinessTask<T> {
  pending(): string[];
  wait(signal: AbortSignal): Promise<T>;
}

export interface FontReadinessProbe {
  requestedFamily: string;
  available: boolean;
}

export interface VisualResourceProbe {
  kind: "image" | "svg" | "chart";
  resource: string;
  anchorId?: string;
  status: "complete" | "failed";
  message?: string;
}

export interface PageTreeStabilityProbe {
  signature: string;
  pageCount: number;
  quietIntervalMs: number;
  animationFrames: number;
  mutations: number;
  resizes: number;
}

export interface FinalPrintReadinessResult {
  fonts: FontReadinessProbe[];
  images: VisualResourceProbe[];
  graphics: VisualResourceProbe[];
  pageTree: PageTreeStabilityProbe;
}

export class PrintReadinessError extends Error {
  readonly phase: PrintReadinessPhase;
  readonly pending: readonly string[];

  constructor(phase: PrintReadinessPhase, message: string, pending: readonly string[] = []) {
    super(message);
    this.name = "PrintReadinessError";
    this.phase = phase;
    this.pending = Object.freeze([...pending]);
  }
}

const MATERIALIZATION_KIND = "data-docxodus-materialization";
const MATERIALIZATION_STATE = "data-docxodus-materialization-state";
const MATERIALIZATION_ID = "data-docxodus-materialization-id";
const GRAPHIC_SELECTOR = `svg, [${MATERIALIZATION_KIND}]`;

function abortError(): DOMException {
  return new DOMException("Print readiness was aborted", "AbortError");
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw abortError();
}

async function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  throwIfAborted(signal);
  let rejectAbort: ((reason?: unknown) => void) | undefined;
  const aborted = new Promise<never>((_, reject) => {
    rejectAbort = reject;
  });
  const onAbort = (): void => rejectAbort?.(abortError());
  signal.addEventListener("abort", onAbort, { once: true });
  try {
    return await Promise.race([promise, aborted]);
  } finally {
    signal.removeEventListener("abort", onAbort);
  }
}

function resourceAnchor(element: Element): string | undefined {
  return element.closest<HTMLElement>("[data-source-anchor-id]")?.dataset.sourceAnchorId;
}

function imageResource(image: HTMLImageElement, index: number): string {
  return image.alt.trim()
    || resourceAnchor(image)
    || `image-${index + 1}`;
}

function uniqueResourceLabels(labels: string[]): string[] {
  const totals = new Map<string, number>();
  const seen = new Map<string, number>();
  labels.forEach((label) => totals.set(label, (totals.get(label) ?? 0) + 1));
  return labels.map((label) => {
    if ((totals.get(label) ?? 0) === 1) return label;
    const occurrence = (seen.get(label) ?? 0) + 1;
    seen.set(label, occurrence);
    return `${label}#${occurrence}`;
  });
}

function graphicKind(element: Element): "svg" | "chart" {
  const declared = element.getAttribute(MATERIALIZATION_KIND);
  return declared === "chart"
    || element.classList.contains("chart")
    || element.closest("[class*='chart']") !== null
    ? "chart"
    : "svg";
}

function graphicResource(element: Element, index: number): string {
  return element.getAttribute(MATERIALIZATION_ID)?.trim()
    || resourceAnchor(element)
    || `${graphicKind(element)}-${index + 1}`;
}

export function documentFontReadiness(document: Document): PrintReadinessTask<FontReadinessProbe[]> {
  const candidates = documentFontReadinessCandidates(document);
  const labels = uniqueResourceLabels(candidates.map(({ family }) => family));
  const pending = new Set(labels);
  let fontSetReady = document.fonts !== undefined;
  return {
    pending: () => [
      ...Array.from(pending, (family) => `font:${family}`).sort(),
      ...(fontSetReady ? ["document.fonts.ready"] : []),
    ],
    async wait(signal) {
      if (!document.fonts) return [];
      const probes = await Promise.all(candidates.map(async ({ family, specification, sample }, index) => {
        const label = labels[index];
        try {
          await abortable(document.fonts.load(specification, sample), signal);
          return {
            requestedFamily: family,
            available: document.fonts.check(specification, sample),
          };
        } catch (error) {
          if (signal.aborted) throw error;
          return { requestedFamily: family, available: false };
        } finally {
          pending.delete(label);
        }
      }));
      await abortable(document.fonts.ready.then(() => undefined), signal);
      fontSetReady = false;
      return probes;
    },
  };
}

export function documentImageReadiness(document: Document): PrintReadinessTask<VisualResourceProbe[]> {
  const images = Array.from(document.images);
  const labels = uniqueResourceLabels(images.map(imageResource));
  const pending = new Set(labels);
  return {
    pending: () => Array.from(pending, (label) => `image:${label}`).sort(),
    wait: (signal) => Promise.all(images.map(async (image, index) => {
      const resource = labels[index];
      const anchorId = resourceAnchor(image);
      try {
        if (typeof image.decode === "function") await abortable(image.decode(), signal);
        throwIfAborted(signal);
        if (!image.complete || image.naturalWidth <= 0 || image.naturalHeight <= 0) {
          throw new Error("the browser reported no decoded pixels");
        }
        return {
          kind: "image" as const,
          resource,
          ...(anchorId ? { anchorId } : {}),
          status: "complete" as const,
        };
      } catch (error) {
        if (signal.aborted) throw error;
        return {
          kind: "image" as const,
          resource,
          ...(anchorId ? { anchorId } : {}),
          status: "failed" as const,
          message: error instanceof Error ? error.message : String(error),
        };
      } finally {
        pending.delete(resource);
      }
    })),
  };
}

function validateGraphic(element: Element): string | undefined {
  const svg = element.localName === "svg"
    ? element as SVGSVGElement
    : element.querySelector<SVGSVGElement>("svg");
  if (!svg) return "the materializer did not produce an SVG root";
  const width = Number.parseFloat(svg.getAttribute("width") ?? "");
  const height = Number.parseFloat(svg.getAttribute("height") ?? "");
  if (!svg.hasAttribute("viewBox") && !(width > 0 && height > 0)) {
    return "the SVG has neither a viewBox nor explicit dimensions";
  }
  return undefined;
}

export function documentGraphicReadiness(document: Document): PrintReadinessTask<VisualResourceProbe[]> {
  const graphics = Array.from(new Set(Array.from(
    document.querySelectorAll<Element>(GRAPHIC_SELECTOR),
  )))
    .filter((element) => element.closest(`[${MATERIALIZATION_KIND}]`) === element
      || !element.parentElement?.closest(`[${MATERIALIZATION_KIND}]`));
  const labels = uniqueResourceLabels(graphics.map(graphicResource));
  const pending = new Set<string>();
  const refreshPending = (): void => {
    pending.clear();
    graphics.forEach((element, index) => {
      if (element.getAttribute(MATERIALIZATION_STATE) === "pending") pending.add(labels[index]);
    });
  };
  refreshPending();
  return {
    pending: () => Array.from(pending, (label) => `materialization:${label}`).sort(),
    async wait(signal) {
      if (pending.size > 0) {
        const view = document.defaultView;
        if (!view) throw new Error("render document has no defaultView");
        await abortable(new Promise<void>((resolve) => {
          const observer = new view.MutationObserver(() => {
            refreshPending();
            if (pending.size === 0) {
              observer.disconnect();
              resolve();
            }
          });
          observer.observe(document.documentElement, {
            attributes: true,
            attributeFilter: [MATERIALIZATION_STATE],
            childList: true,
            subtree: true,
          });
          signal.addEventListener("abort", () => observer.disconnect(), { once: true });
        }), signal);
      }
      throwIfAborted(signal);
      return graphics.map((element, index) => {
        const resource = labels[index];
        const anchorId = resourceAnchor(element);
        const declaredState = element.getAttribute(MATERIALIZATION_STATE);
        const validation = validateGraphic(element);
        const message = declaredState === "failed"
          ? element.getAttribute("data-docxodus-materialization-error")
            || "the materializer reported failure"
          : validation;
        return {
          kind: graphicKind(element),
          resource,
          ...(anchorId ? { anchorId } : {}),
          status: message ? "failed" as const : "complete" as const,
          ...(message ? { message } : {}),
        };
      });
    },
  };
}

function treeSignature(document: Document, pages: HTMLElement[]): string {
  const geometry = pages.map((page) => {
    const rect = page.getBoundingClientRect();
    return [
      page.dataset.pageNumber,
      page.dataset.sectionIndex,
      rect.width.toFixed(3),
      rect.height.toFixed(3),
      page.scrollWidth,
      page.scrollHeight,
    ];
  });
  const fragments = Array.from(
    document.querySelectorAll<HTMLElement>(".page-box [data-source-anchor-id]"),
    (element) => {
      const rect = element.getBoundingClientRect();
      const style = document.defaultView!.getComputedStyle(element);
      return [
        element.dataset.sourceAnchorId,
        element.dataset.pageNumber,
        element.dataset.fragmentIndex,
        rect.left.toFixed(3),
        rect.top.toFixed(3),
        rect.width.toFixed(3),
        rect.height.toFixed(3),
        style.display,
        style.visibility,
      ];
    },
  );
  return JSON.stringify({
    fragments,
    geometry,
    nodes: document.querySelectorAll("*").length,
    textLength: document.body.textContent?.length ?? 0,
  });
}

async function animationFrame(document: Document, signal: AbortSignal): Promise<void> {
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  throwIfAborted(signal);
  await new Promise<void>((resolve, reject) => {
    let settled = false;
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      view.cancelAnimationFrame(requestId);
      reject(abortError());
    };
    const requestId = view.requestAnimationFrame(() => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      resolve();
    });
    signal.addEventListener("abort", onAbort, { once: true });
    if (signal.aborted) onAbort();
  });
}

async function delay(milliseconds: number, signal: AbortSignal): Promise<void> {
  throwIfAborted(signal);
  await new Promise<void>((resolve, reject) => {
    let settled = false;
    const onAbort = (): void => {
      if (settled) return;
      settled = true;
      globalThis.clearTimeout(timer);
      reject(abortError());
    };
    const timer = globalThis.setTimeout(() => {
      if (settled) return;
      settled = true;
      signal.removeEventListener("abort", onAbort);
      resolve();
    }, milliseconds);
    signal.addEventListener("abort", onAbort, { once: true });
    if (signal.aborted) onAbort();
  });
}

export function pageTreeReadiness(
  document: Document,
  pages: HTMLElement[],
  quietIntervalMs = 100,
): PrintReadinessTask<PageTreeStabilityProbe> {
  return {
    pending: () => [
      `page-tree:${pages.length}-pages`,
      `quiet-interval:${quietIntervalMs}ms`,
    ],
    async wait(signal) {
      const view = document.defaultView as (Window & typeof globalThis) | null;
      if (!view) throw new Error("render document has no defaultView");
      let mutations = 0;
      let resizes = 0;
      const mutationObserver = new view.MutationObserver((records) => {
        mutations += records.length;
      });
      const resizeObserver = typeof view.ResizeObserver === "function"
        ? new view.ResizeObserver((records) => {
          resizes += records.length;
        })
        : undefined;
      mutationObserver.observe(document.documentElement, {
        attributes: true,
        characterData: true,
        childList: true,
        subtree: true,
      });
      for (const page of pages) resizeObserver?.observe(page);
      try {
        await animationFrame(document, signal);
        await animationFrame(document, signal);
        const first = treeSignature(document, pages);
        mutations = 0;
        resizes = 0;
        await delay(quietIntervalMs, signal);
        await animationFrame(document, signal);
        await animationFrame(document, signal);
        const second = treeSignature(document, pages);
        if (first !== second || mutations !== 0 || resizes !== 0) {
          throw new PrintReadinessError(
            "page_tree_stability",
            `Final page tree changed during the quiet interval (mutations=${mutations}, resizes=${resizes})`,
            [`page-tree:${pages.length}-pages`],
          );
        }
        return {
          signature: second,
          pageCount: pages.length,
          quietIntervalMs,
          animationFrames: 4,
          mutations,
          resizes,
        };
      } finally {
        mutationObserver.disconnect();
        resizeObserver?.disconnect();
      }
    },
  };
}

async function boundedTask<T>(
  phase: PrintReadinessPhase,
  task: PrintReadinessTask<T>,
  deadline: number,
): Promise<T> {
  const remaining = deadline - Date.now();
  if (remaining <= 0) {
    throw new PrintReadinessError(phase, `Print readiness timed out during ${phase}.`, task.pending());
  }
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      task.wait(controller.signal),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => {
          const pending = task.pending();
          controller.abort();
          reject(new PrintReadinessError(
            phase,
            `Print readiness timed out during ${phase}.`,
            pending,
          ));
        }, remaining);
      }),
    ]);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    controller.abort();
  }
}

/**
 * Re-check the serialized standalone page tree in the exact document Chromium
 * will print. This closes the serialization/reopen race without re-pagination.
 */
export async function awaitFinalPrintReadiness(
  document: Document,
  options: { timeoutMs: number; quietIntervalMs?: number },
): Promise<FinalPrintReadinessResult> {
  const deadline = Date.now() + options.timeoutMs;
  const fonts = await boundedTask("font_loading", documentFontReadiness(document), deadline);
  const images = await boundedTask("image_decoding", documentImageReadiness(document), deadline);
  const failedImage = images.find(({ status }) => status === "failed");
  if (failedImage) {
    throw new PrintReadinessError(
      "image_decoding",
      `Image failed to decode: ${failedImage.resource}${failedImage.message ? ` (${failedImage.message})` : ""}`,
      [`image:${failedImage.resource}`],
    );
  }
  const graphics = await boundedTask(
    "chart_svg_materialization",
    documentGraphicReadiness(document),
    deadline,
  );
  const failedGraphic = graphics.find(({ status }) => status === "failed");
  if (failedGraphic) {
    throw new PrintReadinessError(
      "chart_svg_materialization",
      `${failedGraphic.kind} failed to materialize: ${failedGraphic.resource}${failedGraphic.message ? ` (${failedGraphic.message})` : ""}`,
      [`materialization:${failedGraphic.resource}`],
    );
  }
  const pages = Array.from(document.querySelectorAll<HTMLElement>(".page-box"));
  if (pages.length === 0) {
    throw new PrintReadinessError(
      "page_tree_stability",
      "The final print document contains no page boxes.",
      ["page-tree:missing"],
    );
  }
  const pageTree = await boundedTask(
    "page_tree_stability",
    pageTreeReadiness(document, pages, options.quietIntervalMs ?? 100),
    deadline,
  );
  return { fonts, images, graphics, pageTree };
}

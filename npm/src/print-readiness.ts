/**
 * Deterministic browser readiness primitives shared by standalone HTML
 * materialization and the final document reopened for Chromium printing.
 */

import {
  automaticUrlAllowed,
  cssSecurityTokens,
  dataUrlInfo,
} from "./standalone-resource-policy.js";

export type PrintReadinessPhase =
  | "font_loading"
  | "image_decoding"
  | "chart_svg_materialization"
  | "page_tree_stability";

export interface PrintReadinessTask<T> {
  pending(): string[];
  wait(signal: AbortSignal): Promise<T>;
}

export interface PrintReadinessLimits {
  fontRequests: number;
  fontSampleCodePoints: number;
  visualResources: number;
  domNodes: number;
  automaticResourceBytes: number;
}

export interface FinalPrintReadinessOptions {
  timeoutMs: number;
  quietIntervalMs?: number;
  signal?: AbortSignal;
  limits?: Partial<PrintReadinessLimits>;
}

export interface FontReadinessProbe {
  /** SHA-256 commitment to the exact FontFaceSet.load/check specification and sample. */
  requestKey: string;
  requestedFamily: string;
  available: boolean;
}

export interface VisualResourceProbe {
  kind: "image" | "svg" | "chart";
  source: "html-image" | "css-background" | "svg-image" | "graphic" | "svg-use";
  resource: string;
  anchorId?: string;
  status: "complete" | "failed";
  /** SHA-256 commitment to the exact bounded source dependency. */
  contentKey: string;
  message?: string;
  mediaType?: string;
  byteLength?: number;
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
  readonly reason: "readiness_failure" | "resource_limit";

  constructor(
    phase: PrintReadinessPhase,
    message: string,
    pending: readonly string[] = [],
    reason: "readiness_failure" | "resource_limit" = "readiness_failure",
  ) {
    super(message);
    this.name = "PrintReadinessError";
    this.phase = phase;
    this.pending = Object.freeze(boundedPending(pending));
    this.reason = reason;
  }
}

const DEFAULT_LIMITS: Readonly<PrintReadinessLimits> = Object.freeze({
  fontRequests: 4_096,
  fontSampleCodePoints: 65_536,
  visualResources: 10_000,
  domNodes: 1_000_000,
  automaticResourceBytes: 268_435_456,
});
const PENDING_DETAILS_MAX = 64;
const RESOURCE_LABEL_MAX = 256;
const RESOURCE_SETTLE_PASSES_MAX = 8;
const IMAGE_DECODE_CONCURRENCY = 16;
const CSS_BACKGROUND_PSEUDOS = [null, "::before", "::after"] as const;

const MATERIALIZATION_KIND = "data-docxodus-materialization";
const MATERIALIZATION_STATE = "data-docxodus-materialization-state";
const MATERIALIZATION_ID = "data-docxodus-materialization-id";
const GRAPHIC_SELECTOR = `svg, [${MATERIALIZATION_KIND}]`;

function monotonicNow(document?: Document): number {
  return document?.defaultView?.performance?.now()
    ?? globalThis.performance?.now()
    ?? Date.now();
}

function boundedLabel(value: string, fallback: string): string {
  const normalized = value.replace(/[\u0000-\u001f\u007f]+/g, " ").trim() || fallback;
  return normalized.length <= RESOURCE_LABEL_MAX
    ? normalized
    : `${normalized.slice(0, RESOURCE_LABEL_MAX - 3)}...`;
}

function boundedPending(values: readonly string[]): string[] {
  const bounded = values.slice(0, PENDING_DETAILS_MAX).map((value, index) =>
    boundedLabel(value, `resource-${index + 1}`));
  if (values.length > PENDING_DETAILS_MAX) {
    bounded.push(`... ${values.length - PENDING_DETAILS_MAX} more`);
  }
  return bounded;
}

function normalizeLimit(value: number | undefined, fallback: number): number {
  return Number.isSafeInteger(value) && value! > 0 ? value! : fallback;
}

function readinessLimits(overrides?: Partial<PrintReadinessLimits>): PrintReadinessLimits {
  return {
    fontRequests: normalizeLimit(overrides?.fontRequests, DEFAULT_LIMITS.fontRequests),
    fontSampleCodePoints: normalizeLimit(
      overrides?.fontSampleCodePoints,
      DEFAULT_LIMITS.fontSampleCodePoints,
    ),
    visualResources: normalizeLimit(overrides?.visualResources, DEFAULT_LIMITS.visualResources),
    domNodes: normalizeLimit(overrides?.domNodes, DEFAULT_LIMITS.domNodes),
    automaticResourceBytes: normalizeLimit(
      overrides?.automaticResourceBytes,
      DEFAULT_LIMITS.automaticResourceBytes,
    ),
  };
}

function abortError(): DOMException {
  return new DOMException("Print readiness was aborted", "AbortError");
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw abortError();
}

async function sha256Hex(
  document: Document,
  phase: PrintReadinessPhase,
  domain: string,
  value: string,
): Promise<string> {
  const crypto = document.defaultView?.crypto ?? globalThis.crypto;
  if (!crypto?.subtle) {
    throw new PrintReadinessError(
      phase,
      "Web Crypto SHA-256 is unavailable for print-readiness evidence.",
      ["readiness-evidence-digest"],
    );
  }
  const material = new TextEncoder().encode(`${domain}\u0000${value}`);
  const digest = await crypto.subtle.digest("SHA-256", material);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
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

async function boundedMap<T, U>(
  values: readonly T[],
  concurrency: number,
  operation: (value: T, index: number) => Promise<U>,
): Promise<U[]> {
  const results = new Array<U>(values.length);
  let cursor = 0;
  const workers = Array.from(
    { length: Math.min(concurrency, values.length) },
    async () => {
      while (cursor < values.length) {
        const index = cursor++;
        results[index] = await operation(values[index], index);
      }
    },
  );
  await Promise.all(workers);
  return results;
}

function resourceAnchor(element: Element): string | undefined {
  return element.closest<HTMLElement>("[data-source-anchor-id]")?.dataset.sourceAnchorId;
}

function imageResource(image: HTMLImageElement, index: number): string {
  return boundedLabel(image.alt.trim()
    || resourceAnchor(image)
    || `image-${index + 1}`, `image-${index + 1}`);
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
  return boundedLabel(element.getAttribute(MATERIALIZATION_ID)?.trim()
    || resourceAnchor(element)
    || `${graphicKind(element)}-${index + 1}`, `${graphicKind(element)}-${index + 1}`);
}

function graphicInventory(document: Document): Element[] {
  return Array.from(new Set(Array.from(
    document.querySelectorAll<Element>(GRAPHIC_SELECTOR),
  )))
    .filter((element) => element.closest(`[${MATERIALIZATION_KIND}]`) === element
      || !element.parentElement?.closest(`[${MATERIALIZATION_KIND}]`));
}

async function fontCandidates(
  document: Document,
  limits: PrintReadinessLimits,
  signal: AbortSignal,
): Promise<Array<{ family: string; specification: string; sample: string }>> {
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  const candidates = new Map<string, { family: string; specification: string; sample: string }>();
  let sampledCodePoints = 0;
  let visitedTextNodes = 0;
  const walker = document.createTreeWalker(document.body, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (++visitedTextNodes % 1_024 === 0) {
      throwIfAborted(signal);
      await delay(0, signal);
    }
    const element = node.parentElement;
    const sample = node.textContent?.trim();
    if (!element || !sample || element.closest("script, style, template, noscript")) continue;
    const style = view.getComputedStyle(element);
    if (style.display === "none" || style.visibility === "hidden"
      || style.contentVisibility === "hidden") continue;
    const familySpecification = style.fontFamily.trim();
    if (!familySpecification) continue;
    const family = boundedLabel(familySpecification, "sans-serif");
    const specification = [
      style.fontStyle || "normal",
      style.fontVariant || "normal",
      style.fontWeight || "400",
      style.fontStretch || "normal",
      style.fontSize || "12px",
      familySpecification,
    ].join(" ");
    const existing = candidates.get(specification);
    const remainingSamples = Math.max(0, limits.fontSampleCodePoints - sampledCodePoints);
    const addition = Array.from(sample).slice(0, Math.min(256, remainingSamples)).join("");
    if (existing && addition) {
      const remainingForCandidate = Math.max(0, 256 - Array.from(existing.sample).length);
      const appended = Array.from(addition).slice(0, remainingForCandidate).join("");
      existing.sample += appended;
      sampledCodePoints += Array.from(appended).length;
    } else if (!existing) {
      if (candidates.size >= limits.fontRequests) {
        throw new PrintReadinessError(
          "font_loading",
          `Font readiness exceeded its ${limits.fontRequests}-request limit.`,
          [`font-request-limit:${limits.fontRequests}`],
          "resource_limit",
        );
      }
      const boundedSample = addition || " ";
      candidates.set(specification, { family, specification, sample: boundedSample });
      sampledCodePoints += Array.from(addition).length;
    }
  }
  return Array.from(candidates.values()).sort((left, right) =>
    left.specification < right.specification ? -1 : left.specification > right.specification ? 1 : 0);
}

export function documentFontReadiness(
  document: Document,
  configuredLimits?: Partial<PrintReadinessLimits>,
): PrintReadinessTask<FontReadinessProbe[]> {
  const limits = readinessLimits(configuredLimits);
  const pending = new Set<string>();
  let fontSetReady = document.fonts !== undefined;
  let settled = false;
  let currentLabels: string[] = [];
  const inventory = async (signal: AbortSignal) => {
    const candidates = await fontCandidates(document, limits, signal);
    const labels = uniqueResourceLabels(candidates.map(({ family }) => family));
    const signature = candidates.map(({ specification, sample }) =>
      `${specification}\u0000${sample}`).join("\u0001");
    return { candidates, labels, signature };
  };
  return {
    pending: () => [
      ...(settled ? [] : pending.size > 0
        ? Array.from(pending, (family) => `font:${family}`).sort()
        : currentLabels.length > 0
          ? currentLabels.map((family) => `font:${family}`).sort()
          : ["font-inventory"]),
      ...(settled ? [] : fontSetReady ? ["document.fonts.ready"] : []),
    ],
    async wait(signal) {
      if (!document.fonts) {
        settled = true;
        return [];
      }
      for (let pass = 1; pass <= RESOURCE_SETTLE_PASSES_MAX; pass++) {
        const before = await inventory(signal);
        currentLabels = before.labels;
        pending.clear();
        before.labels.forEach((label) => pending.add(label));
        fontSetReady = true;
        const probes = await Promise.all(before.candidates.map(
          async ({ family, specification, sample }, index) => {
            const label = before.labels[index];
            const requestKey = await sha256Hex(
              document,
              "font_loading",
              "docxodus:font-request:v1",
              JSON.stringify({ specification, sample }),
            );
            try {
              await abortable(document.fonts.load(specification, sample), signal);
              return {
                requestKey,
                requestedFamily: family,
                available: document.fonts.check(specification, sample),
              };
            } catch (error) {
              if (signal.aborted) throw error;
              return { requestKey, requestedFamily: family, available: false };
            } finally {
              pending.delete(label);
            }
          },
        ));
        await abortable(document.fonts.ready.then(() => undefined), signal);
        fontSetReady = false;
        await animationFrame(document, signal);
        if ((await inventory(signal)).signature === before.signature) {
          settled = true;
          return probes;
        }
      }
      throw new PrintReadinessError(
        "font_loading",
        "The computed font inventory did not settle.",
        currentLabels.map((label) => `font:${label}`),
      );
    },
  };
}

interface ImageDependency {
  source: "html-image" | "css-background" | "svg-image";
  element: Element;
  url: string;
  rawSignature: string;
  contentMaterial: string;
  fallbackLabel: string;
  image?: HTMLImageElement;
}

interface ImageDependencyInventory {
  dependencies: ImageDependency[];
  labels: string[];
  signature: string;
}

function localSvgReference(document: Document, value: string): Element | undefined {
  if (!value.startsWith("#") || value.length === 1) return undefined;
  let id = value.slice(1);
  try { id = decodeURIComponent(id); } catch { return undefined; }
  return document.getElementById(id) ?? undefined;
}

function drawableSvgTarget(target: Element): boolean {
  return target.matches(
    "path, rect, circle, ellipse, line, polyline, polygon, text, image, use, foreignObject",
  ) || target.querySelector(
    "path, rect, circle, ellipse, line, polyline, polygon, text, image, use, foreignObject",
  ) !== null;
}

async function imageDependencyInventory(
  document: Document,
  limits: PrintReadinessLimits,
  signal: AbortSignal,
  identities: WeakMap<Element, number>,
  nextIdentity: { value: number },
): Promise<ImageDependencyInventory> {
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  const dependencies: ImageDependency[] = [];
  const identity = (element: Element): number => {
    const existing = identities.get(element);
    if (existing !== undefined) return existing;
    const assigned = nextIdentity.value++;
    identities.set(element, assigned);
    return assigned;
  };
  const admit = (dependency: ImageDependency): void => {
    if (dependencies.length >= limits.visualResources) {
      throw new PrintReadinessError(
        "image_decoding",
        `Image readiness exceeded its ${limits.visualResources}-resource limit.`,
        [`image-resource-limit:${limits.visualResources}`],
        "resource_limit",
      );
    }
    dependencies.push(dependency);
  };

  Array.from(document.images).forEach((image, index) => {
    admit({
      source: "html-image",
      element: image,
      image,
      url: image.currentSrc || image.getAttribute("src") || "",
      rawSignature: `${identity(image)}\u0000${image.getAttribute("src") ?? ""}\u0000${image.getAttribute("srcset") ?? ""}\u0000${image.currentSrc}`,
      contentMaterial: JSON.stringify({
        src: image.getAttribute("src") ?? "",
        srcset: image.getAttribute("srcset") ?? "",
      }),
      fallbackLabel: boundedLabel(`html-image:${imageResource(image, index)}`, `html-image-${index + 1}`),
    });
  });

  const elements = Array.from(document.querySelectorAll<Element>("body, body *"));
  if (elements.length > limits.domNodes) {
    throw new PrintReadinessError(
      "image_decoding",
      `CSS background readiness exceeded its ${limits.domNodes}-element work limit.`,
      [`css-background-element-limit:${limits.domNodes}`],
      "resource_limit",
    );
  }
  const backgroundTokens = new Map<string, {
    tokens: ReturnType<typeof cssSecurityTokens>;
    digest: string;
  }>();
  let backgroundCodeUnits = 0;
  for (const [elementIndex, element] of elements.entries()) {
    if ((elementIndex + 1) % 256 === 0) await delay(0, signal);
    throwIfAborted(signal);
    for (const pseudo of CSS_BACKGROUND_PSEUDOS) {
      let style: CSSStyleDeclaration;
      try { style = view.getComputedStyle(element, pseudo); } catch { continue; }
      if (style.display === "none" || style.visibility === "hidden"
        || style.contentVisibility === "hidden"
        || (pseudo !== null && (style.content === "none" || style.content === "normal"))) continue;
      const background = style.backgroundImage.trim();
      if (!background || background === "none") continue;
      backgroundCodeUnits += background.length;
      if (backgroundCodeUnits > limits.automaticResourceBytes) {
        throw new PrintReadinessError(
          "image_decoding",
          `CSS background readiness exceeded its ${limits.automaticResourceBytes}-code-unit work limit.`,
          [`css-background-code-unit-limit:${limits.automaticResourceBytes}`],
          "resource_limit",
        );
      }
      let backgroundEvidence = backgroundTokens.get(background);
      if (!backgroundEvidence) {
        backgroundEvidence = {
          tokens: cssSecurityTokens(background),
          digest: await sha256Hex(
            document,
            "image_decoding",
            "docxodus:computed-background:v1",
            background,
          ),
        };
        backgroundTokens.set(background, backgroundEvidence);
      }
      for (const [urlIndex, token] of backgroundEvidence.tokens.entries()) {
        if (token.kind !== "url") continue;
        const url = token.value.trim();
        // `data:,` is the sanitizer's inert omitted-resource sentinel, not pixels.
        if (url === "data:,") continue;
        const anchor = resourceAnchor(element);
        const pseudoLabel = pseudo === null ? "element" : pseudo.slice(2);
        admit({
          source: "css-background",
          element,
          url,
          rawSignature: `${identity(element)}\u0000${pseudoLabel}\u0000${backgroundEvidence.digest}\u0000${urlIndex}`,
          contentMaterial: JSON.stringify({
            backgroundDigest: backgroundEvidence.digest,
            pseudo: pseudoLabel,
            urlIndex,
          }),
          fallbackLabel: boundedLabel(
            `css-background:${anchor || `${elementIndex + 1}-${pseudoLabel}-${urlIndex + 1}`}`,
            `css-background-${elementIndex + 1}-${pseudoLabel}-${urlIndex + 1}`,
          ),
        });
      }
    }
  }

  const svgImages = Array.from(document.querySelectorAll<SVGImageElement>("svg image"));
  svgImages.forEach((element, index) => {
    const url = element.getAttribute("href") ?? element.getAttribute("xlink:href") ?? "";
    admit({
      source: "svg-image",
      element,
      url: url.trim(),
      rawSignature: `${identity(element)}\u0000${url}\u0000${element.outerHTML}`,
      contentMaterial: JSON.stringify({ url: url.trim(), markup: element.outerHTML }),
      fallbackLabel: boundedLabel(
        `svg-image:${resourceAnchor(element) || index + 1}`,
        `svg-image-${index + 1}`,
      ),
    });
  });

  const labels = uniqueResourceLabels(dependencies.map(({ fallbackLabel }) => fallbackLabel));
  return {
    dependencies,
    labels,
    signature: dependencies.map((dependency, index) =>
      `${dependency.source}\u0000${labels[index]}\u0000${dependency.rawSignature}`).join("\u0001"),
  };
}

async function decodeImageDependency(
  document: Document,
  dependency: ImageDependency,
  signal: AbortSignal,
): Promise<void> {
  if (dependency.image) {
    if (typeof dependency.image.decode === "function") {
      await abortable(dependency.image.decode(), signal);
    }
    throwIfAborted(signal);
    if (!dependency.image.complete || dependency.image.naturalWidth <= 0
      || dependency.image.naturalHeight <= 0) {
      throw new Error("the browser reported no decoded pixels");
    }
    return;
  }
  const info = dataUrlInfo(dependency.url);
  if (!automaticUrlAllowed(dependency.url) || !info || !info.mediaType.startsWith("image/")) {
    throw new Error("the image reference is not an allowed embedded image data URL");
  }
  if (dependency.source === "svg-image") {
    const svgImage = dependency.element as SVGImageElement & { decode?: () => Promise<void> };
    if (typeof svgImage.decode === "function") {
      await abortable(svgImage.decode(), signal);
      throwIfAborted(signal);
      const bounds = svgImage.getBBox();
      if (!(bounds.width > 0 && bounds.height > 0)) {
        throw new Error("the SVG image produced no drawable bounds");
      }
      return;
    }
    throw new Error("SVG image decoding is unavailable");
  }
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  const image = new view.Image();
  try {
    image.src = dependency.url;
    if (typeof image.decode === "function") await abortable(image.decode(), signal);
    throwIfAborted(signal);
    if (!image.complete || image.naturalWidth <= 0 || image.naturalHeight <= 0) {
      throw new Error("the browser reported no decoded pixels");
    }
  } finally {
    image.removeAttribute("src");
  }
}

export function documentImageReadiness(
  document: Document,
  configuredLimits?: Partial<PrintReadinessLimits>,
): PrintReadinessTask<VisualResourceProbe[]> {
  const limits = readinessLimits(configuredLimits);
  const pending = new Set<string>();
  let settled = false;
  let started = false;
  let currentLabels: string[] = [];
  const identities = new WeakMap<Element, number>();
  const nextIdentity = { value: 1 };
  const inventory = (signal: AbortSignal) =>
    imageDependencyInventory(document, limits, signal, identities, nextIdentity);
  return {
    pending: () => {
      if (settled) return [];
      if (!started) return ["image-inventory"];
      if (pending.size > 0) return Array.from(pending, (label) => `image:${label}`).sort();
      return currentLabels.map((label) => `image:${label}`).sort();
    },
    async wait(signal) {
      started = true;
      for (let pass = 1; pass <= RESOURCE_SETTLE_PASSES_MAX; pass++) {
        throwIfAborted(signal);
        const before = await inventory(signal);
        currentLabels = before.labels;
        pending.clear();
        before.labels.forEach((label) => pending.add(label));
        const cssDecodeCache = new Map<string, Promise<void>>();
        const probes = await boundedMap(
          before.dependencies,
          IMAGE_DECODE_CONCURRENCY,
          async (dependency, index) => {
          const resource = before.labels[index];
          const anchorId = resourceAnchor(dependency.element);
          const info = dataUrlInfo(dependency.url);
          const contentKey = await sha256Hex(
            document,
            "image_decoding",
            "docxodus:visual-resource:v1",
            JSON.stringify({ source: dependency.source, material: dependency.contentMaterial }),
          );
          try {
            let decoding: Promise<void>;
            if (dependency.source === "css-background") {
              decoding = cssDecodeCache.get(dependency.url)
                ?? decodeImageDependency(document, dependency, signal);
              cssDecodeCache.set(dependency.url, decoding);
            } else {
              decoding = decodeImageDependency(document, dependency, signal);
            }
            await decoding;
            return {
              kind: "image" as const,
              source: dependency.source,
              resource,
              ...(anchorId ? { anchorId } : {}),
              status: "complete" as const,
              contentKey,
              ...(info?.mediaType ? { mediaType: info.mediaType } : {}),
              ...(info ? { byteLength: info.byteLength } : {}),
            };
          } catch (error) {
            if (signal.aborted) throw error;
            return {
              kind: "image" as const,
              source: dependency.source,
              resource,
              ...(anchorId ? { anchorId } : {}),
              status: "failed" as const,
              contentKey,
              ...(info?.mediaType ? { mediaType: info.mediaType } : {}),
              ...(info ? { byteLength: info.byteLength } : {}),
              message: boundedLabel(
                error instanceof Error ? error.message : String(error),
                "image readiness failure",
              ),
            };
          } finally {
            pending.delete(resource);
          }
          },
        );
        await animationFrame(document, signal);
        if ((await inventory(signal)).signature === before.signature) {
          settled = true;
          return probes;
        }
      }
      throw new PrintReadinessError(
        "image_decoding",
        "The image resource inventory did not settle.",
        currentLabels.map((label) => `image:${label}`),
      );
    },
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
  if (!svg.querySelector(
    "path, rect, circle, ellipse, line, polyline, polygon, text, image, use, foreignObject",
  )) {
    return "the SVG contains no drawable content";
  }
  return undefined;
}

function validateSvgUse(document: Document, element: SVGUseElement): string | undefined {
  const href = (element.getAttribute("href") ?? element.getAttribute("xlink:href") ?? "").trim();
  if (!href.startsWith("#") || href.length === 1) {
    return "the SVG use reference is external or empty";
  }
  const target = localSvgReference(document, href);
  if (!target || target === element || target.contains(element) || !drawableSvgTarget(target)) {
    return "the SVG use reference is missing, cyclic, or has no drawable target";
  }
  try {
    const bounds = element.getBBox();
    if (!(bounds.width > 0 || bounds.height > 0)) {
      return "the SVG use reference produced no drawable bounds";
    }
  } catch {
    return "the SVG use reference could not be materialized";
  }
  return undefined;
}

export function documentGraphicReadiness(
  document: Document,
  configuredLimits?: Partial<PrintReadinessLimits>,
): PrintReadinessTask<VisualResourceProbe[]> {
  const limits = readinessLimits(configuredLimits);
  const identities = new WeakMap<Element, number>();
  let nextIdentity = 1;
  let currentPending: string[] = [];
  let settled = false;
  let started = false;
  const inventory = (): {
    graphics: Element[];
    graphicLabels: string[];
    uses: SVGUseElement[];
    useLabels: string[];
    labels: string[];
    signature: string;
  } => {
    const graphics = graphicInventory(document);
    const uses = Array.from(document.querySelectorAll<SVGUseElement>("svg use"));
    if (graphics.length + uses.length > limits.visualResources) {
      throw new PrintReadinessError(
        "chart_svg_materialization",
        `Graphic readiness exceeded its ${limits.visualResources}-resource limit.`,
        [`graphic-resource-limit:${limits.visualResources}`],
        "resource_limit",
      );
    }
    const graphicLabels = uniqueResourceLabels(graphics.map((element, index) =>
      boundedLabel(
        `graphic:${graphicKind(element)}:${graphicResource(element, index)}`,
        `graphic-${index + 1}`,
      )));
    const useLabels = uniqueResourceLabels(uses.map((element, index) => boundedLabel(
      `svg-use:${resourceAnchor(element) || index + 1}`,
      `svg-use-${index + 1}`,
    )));
    const labels = [...graphicLabels, ...useLabels];
    currentPending = graphics.flatMap((element, index) =>
      element.getAttribute(MATERIALIZATION_STATE) === "pending" ? [graphicLabels[index]] : []);
    const signature = [...graphics, ...uses].map((element) => {
      let identity = identities.get(element);
      if (identity === undefined) {
        identity = nextIdentity++;
        identities.set(element, identity);
      }
      return `${identity}\u0000${element.outerHTML}`;
    }).join("\u0001");
    return { graphics, graphicLabels, uses, useLabels, labels, signature };
  };
  return {
    pending: () => {
      if (settled) return [];
      if (!started) return ["graphic-inventory"];
      inventory();
      return currentPending.map((label) => `materialization:${label}`).sort();
    },
    async wait(signal) {
      started = true;
      let settlePasses = 0;
      while (true) {
        throwIfAborted(signal);
        const before = inventory();
        if (currentPending.length > 0) {
          settlePasses = 0;
          await animationFrame(document, signal);
          continue;
        }
        await animationFrame(document, signal);
        const after = inventory();
        if (after.signature !== before.signature) {
          settlePasses++;
          if (settlePasses >= RESOURCE_SETTLE_PASSES_MAX) {
            throw new PrintReadinessError(
              "chart_svg_materialization",
              "The graphic resource inventory did not settle.",
              after.labels.map((label) => `materialization:${label}`),
            );
          }
          continue;
        }
        const probes = await boundedMap(
          after.graphics,
          IMAGE_DECODE_CONCURRENCY,
          async (element, index): Promise<VisualResourceProbe> => {
          const resource = after.graphicLabels[index];
          const anchorId = resourceAnchor(element);
          const declaredState = element.getAttribute(MATERIALIZATION_STATE);
          const validation = validateGraphic(element);
          const hasMaterializer = element.hasAttribute(MATERIALIZATION_KIND);
          const message = hasMaterializer && declaredState === null
            ? "the materializer did not publish a completion state"
            : declaredState !== null && !["complete", "failed"].includes(declaredState)
              ? `the materializer published an unknown state: ${declaredState}`
              : declaredState === "failed"
                ? element.getAttribute("data-docxodus-materialization-error")
                  || "the materializer reported failure"
                : validation;
          return {
            kind: graphicKind(element),
            source: "graphic" as const,
            resource,
            ...(anchorId ? { anchorId } : {}),
            status: message ? "failed" as const : "complete" as const,
            contentKey: await sha256Hex(
              document,
              "chart_svg_materialization",
              "docxodus:visual-resource:v1",
              JSON.stringify({ source: "graphic", markup: element.outerHTML }),
            ),
            ...(message ? {
              message: boundedLabel(message, "graphic readiness failure"),
            } : {}),
          };
          },
        );
        const targetDigests = new Map<Element, string>();
        let targetMarkupCodeUnits = 0;
        for (const [index, element] of after.uses.entries()) {
          const resource = after.useLabels[index];
          const anchorId = resourceAnchor(element);
          const message = validateSvgUse(document, element);
          const href = (element.getAttribute("href")
            ?? element.getAttribute("xlink:href") ?? "").trim();
          const target = localSvgReference(document, href);
          let targetDigest = "";
          if (target) {
            const cached = targetDigests.get(target);
            if (cached) {
              targetDigest = cached;
            } else {
              const markup = target.outerHTML;
              targetMarkupCodeUnits += markup.length;
              if (targetMarkupCodeUnits > limits.automaticResourceBytes) {
                throw new PrintReadinessError(
                  "chart_svg_materialization",
                  `SVG use readiness exceeded its ${limits.automaticResourceBytes}-code-unit work limit.`,
                  [`svg-use-target-code-unit-limit:${limits.automaticResourceBytes}`],
                  "resource_limit",
                );
              }
              targetDigest = await sha256Hex(
                document,
                "chart_svg_materialization",
                "docxodus:svg-use-target:v1",
                markup,
              );
              targetDigests.set(target, targetDigest);
            }
          }
          probes.push({
            kind: "svg",
            source: "svg-use",
            resource,
            ...(anchorId ? { anchorId } : {}),
            status: message ? "failed" : "complete",
            contentKey: await sha256Hex(
              document,
              "chart_svg_materialization",
              "docxodus:visual-resource:v1",
              JSON.stringify({ source: "svg-use", href, targetDigest }),
            ),
            ...(message ? { message: boundedLabel(message, "SVG use readiness failure") } : {}),
          });
        }
        settled = true;
        return probes;
      }
    },
  };
}

async function treeSignature(document: Document, pages: HTMLElement[]): Promise<string> {
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
  const serialized = JSON.stringify({
    fragments,
    geometry,
    nodes: document.querySelectorAll("*").length,
    pageTree: pages.map((page) => page.outerHTML),
  });
  const crypto = document.defaultView?.crypto ?? globalThis.crypto;
  if (!crypto?.subtle) {
    throw new PrintReadinessError(
      "page_tree_stability",
      "Web Crypto SHA-256 is unavailable for the final page-tree signature.",
      ["page-tree-signature"],
    );
  }
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(serialized));
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0")).join("");
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
        const first = await treeSignature(document, pages);
        mutations = 0;
        resizes = 0;
        await delay(quietIntervalMs, signal);
        await animationFrame(document, signal);
        await animationFrame(document, signal);
        const second = await treeSignature(document, pages);
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
  document: Document,
  phase: PrintReadinessPhase,
  task: PrintReadinessTask<T>,
  deadline: number,
  externalSignal?: AbortSignal,
): Promise<T> {
  if (externalSignal?.aborted) throw abortError();
  const remaining = deadline - monotonicNow(document);
  if (remaining <= 0) {
    throw new PrintReadinessError(phase, `Print readiness timed out during ${phase}.`, task.pending());
  }
  const controller = new AbortController();
  let timer: ReturnType<typeof setTimeout> | undefined;
  let abortListener: (() => void) | undefined;
  try {
    const contenders: Array<Promise<T>> = [
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
    ];
    if (externalSignal) {
      contenders.push(new Promise<never>((_, reject) => {
        abortListener = () => {
          controller.abort();
          reject(abortError());
        };
        externalSignal.addEventListener("abort", abortListener, { once: true });
        if (externalSignal.aborted) abortListener();
      }));
    }
    return await Promise.race(contenders);
  } finally {
    if (timer !== undefined) clearTimeout(timer);
    if (externalSignal && abortListener) {
      externalSignal.removeEventListener("abort", abortListener);
    }
    controller.abort();
  }
}

interface ResourceReadinessFixedPoint {
  fonts: FontReadinessProbe[];
  images: VisualResourceProbe[];
  graphics: VisualResourceProbe[];
}

async function admitCombinedVisualInventory(
  document: Document,
  limits: PrintReadinessLimits,
  signal: AbortSignal,
): Promise<number> {
  const imageInventory = await imageDependencyInventory(
    document,
    limits,
    signal,
    new WeakMap<Element, number>(),
    { value: 1 },
  );
  const count = imageInventory.dependencies.length
    + graphicInventory(document).length
    + document.querySelectorAll("svg use").length;
  if (count > limits.visualResources) {
    throw new PrintReadinessError(
      "chart_svg_materialization",
      `Combined visual readiness exceeded its ${limits.visualResources}-resource limit.`,
      [`visual-resource-limit:${limits.visualResources}`],
      "resource_limit",
    );
  }
  return count;
}

/** Admit the entire image/graphic/use inventory before any visual decode work begins. */
export async function admitPrintVisualResources(
  document: Document,
  configuredLimits: Partial<PrintReadinessLimits>,
  signal: AbortSignal,
): Promise<number> {
  return admitCombinedVisualInventory(document, readinessLimits(configuredLimits), signal);
}

async function resourceFixedPointSignature(
  document: Document,
  result: ResourceReadinessFixedPoint,
  identities: WeakMap<Element, number>,
  nextIdentity: { value: number },
): Promise<string> {
  const identity = (element: Element): number => {
    const existing = identities.get(element);
    if (existing !== undefined) return existing;
    const assigned = nextIdentity.value++;
    identities.set(element, assigned);
    return assigned;
  };
  const images = Array.from(document.images, (image) => [
    identity(image),
    image.getAttribute("src") ?? "",
    image.getAttribute("srcset") ?? "",
    image.currentSrc,
    image.complete,
    image.naturalWidth,
    image.naturalHeight,
  ]);
  const graphics = graphicInventory(document).map((element) => [
    identity(element),
    element.outerHTML,
  ]);
  const serialized = JSON.stringify({
    body: document.body.outerHTML,
    fonts: result.fonts,
    images,
    imageOutcomes: result.images,
    graphics,
    graphicOutcomes: result.graphics,
  });
  const crypto = document.defaultView?.crypto ?? globalThis.crypto;
  if (!crypto?.subtle) {
    throw new PrintReadinessError(
      "page_tree_stability",
      "Web Crypto SHA-256 is unavailable for the resource fixed-point signature.",
      ["resource-fixed-point-signature"],
    );
  }
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(serialized));
  return Array.from(new Uint8Array(digest), (value) =>
    value.toString(16).padStart(2, "0")).join("");
}

async function awaitResourceReadinessFixedPoint(
  document: Document,
  limits: PrintReadinessLimits,
  deadline: number,
  signal?: AbortSignal,
): Promise<ResourceReadinessFixedPoint> {
  const fixedPointSignal = signal ?? new AbortController().signal;
  const identities = new WeakMap<Element, number>();
  const nextIdentity = { value: 1 };
  const maximumReservedWork = BigInt(RESOURCE_SETTLE_PASSES_MAX)
    * (BigInt(limits.fontRequests) + BigInt(limits.visualResources));
  let reservedWork = 0n;
  let previousSignature: string | undefined;
  let lastPending: string[] = ["resource-fixed-point"];
  for (let pass = 1; pass <= RESOURCE_SETTLE_PASSES_MAX; pass++) {
    throwIfAborted(fixedPointSignal);
    const visualCount = await admitCombinedVisualInventory(document, limits, fixedPointSignal);
    const passReservation = BigInt(limits.fontRequests) + BigInt(visualCount);
    if (reservedWork + passReservation > maximumReservedWork) {
      throw new PrintReadinessError(
        "page_tree_stability",
        "The resource fixed point exceeded its bounded work budget.",
        [`resource-fixed-point-work:${reservedWork.toString()}`],
        "resource_limit",
      );
    }
    reservedWork += passReservation;
    const fontTask = documentFontReadiness(document, limits);
    const fonts = await boundedTask(document, "font_loading", fontTask, deadline, signal);
    await admitCombinedVisualInventory(document, limits, fixedPointSignal);
    const imageTask = documentImageReadiness(document, limits);
    const images = await boundedTask(document, "image_decoding", imageTask, deadline, signal);
    const failedImage = images.find(({ status }) => status === "failed");
    if (failedImage) {
      throw new PrintReadinessError(
        "image_decoding",
        `Image failed to decode: ${failedImage.resource}${failedImage.message ? ` (${failedImage.message})` : ""}`,
        [`image:${failedImage.resource}`],
      );
    }
    await admitCombinedVisualInventory(document, limits, fixedPointSignal);
    const graphicTask = documentGraphicReadiness(document, limits);
    const graphics = await boundedTask(
      document,
      "chart_svg_materialization",
      graphicTask,
      deadline,
      signal,
    );
    const failedGraphic = graphics.find(({ status }) => status === "failed");
    if (failedGraphic) {
      throw new PrintReadinessError(
        "chart_svg_materialization",
        `${failedGraphic.kind} failed to materialize: ${failedGraphic.resource}${failedGraphic.message ? ` (${failedGraphic.message})` : ""}`,
        [`materialization:${failedGraphic.resource}`],
      );
    }
    await admitCombinedVisualInventory(document, limits, fixedPointSignal);
    const result = { fonts, images, graphics };
    const signature = await resourceFixedPointSignature(
      document,
      result,
      identities,
      nextIdentity,
    );
    if (signature === previousSignature) return result;
    previousSignature = signature;
    lastPending = [
      `resource-fixed-point:${pass}/${RESOURCE_SETTLE_PASSES_MAX}`,
      ...fontTask.pending(),
      ...imageTask.pending(),
      ...graphicTask.pending(),
    ];
    await animationFrame(document, fixedPointSignal);
  }
  throw new PrintReadinessError(
    "page_tree_stability",
    `The combined print-resource inventory did not settle within ${RESOURCE_SETTLE_PASSES_MAX} passes.`,
    lastPending,
  );
}

/**
 * Re-check the serialized standalone page tree in the exact document Chromium
 * will print. This closes the serialization/reopen race without re-pagination.
 */
export async function awaitFinalPrintReadiness(
  document: Document,
  options: FinalPrintReadinessOptions,
): Promise<FinalPrintReadinessResult> {
  if (!Number.isSafeInteger(options.timeoutMs) || options.timeoutMs <= 0) {
    throw new TypeError("timeoutMs must be a positive safe integer");
  }
  const quietIntervalMs = options.quietIntervalMs ?? 100;
  if (!Number.isSafeInteger(quietIntervalMs) || quietIntervalMs < 100 || quietIntervalMs > 10_000) {
    throw new TypeError("quietIntervalMs must be an integer from 100 through 10000");
  }
  const limits = readinessLimits(options.limits);
  const deadline = monotonicNow(document) + options.timeoutMs;
  const view = document.defaultView;
  if (!view) throw new Error("render document has no defaultView");
  let resourceMutationVersion = 0;
  const recordResourceMutations = (records: MutationRecord[]): void => {
    if (records.some((record) => record.type === "childList"
      || record.attributeName === "src"
      || record.attributeName === "srcset"
      || record.attributeName === "href"
      || record.attributeName === "xlink:href"
      || record.attributeName === "style"
      || record.attributeName === "class"
      || record.attributeName === MATERIALIZATION_KIND
      || record.attributeName === MATERIALIZATION_STATE
      || record.attributeName === MATERIALIZATION_ID
      || record.type === "characterData")) {
      resourceMutationVersion++;
    }
  };
  const resourceObserver = new view.MutationObserver(recordResourceMutations);
  const drainResourceMutations = (): number => {
    recordResourceMutations(resourceObserver.takeRecords());
    return resourceMutationVersion;
  };
  resourceObserver.observe(document.documentElement, {
    attributes: true,
    attributeFilter: [
      "src", "srcset", "href", "xlink:href", "style", "class",
      MATERIALIZATION_KIND, MATERIALIZATION_STATE, MATERIALIZATION_ID,
    ],
    childList: true,
    characterData: true,
    subtree: true,
  });
  try {
    const { fonts, images, graphics } = await awaitResourceReadinessFixedPoint(
      document,
      limits,
      deadline,
      options.signal,
    );
    const resourceVersionBeforeStability = drainResourceMutations();
    const pages = Array.from(document.querySelectorAll<HTMLElement>(".page-box"));
    if (pages.length === 0) {
      throw new PrintReadinessError(
        "page_tree_stability",
        "The final print document contains no page boxes.",
        ["page-tree:missing"],
      );
    }
    const pageTree = await boundedTask(
      document,
      "page_tree_stability",
      pageTreeReadiness(document, pages, quietIntervalMs),
      deadline,
      options.signal,
    );
    if (drainResourceMutations() !== resourceVersionBeforeStability) {
      throw new PrintReadinessError(
        "page_tree_stability",
        "The final resource inventory changed after resource readiness completed.",
        [`page-tree:${pages.length}-pages`],
      );
    }
    return { fonts, images, graphics, pageTree };
  } finally {
    resourceObserver.disconnect();
  }
}

import {
  FONT_RESOLVER_CONTRACT_ID,
  FONT_RESOLVER_SCHEMA_VERSION,
  FONT_SUBSTITUTION_CONTRACT_MATERIAL,
  FONT_SUBSTITUTION_CONTRACT_VERSION,
  fontFamilyKey,
  normalizeFontFamilyName,
  type FontConfigurationIdentity,
  type FontFamilyKind,
  type FontFaceStyle,
  type FontFileFormat,
  type FontRequest,
  type FontResolution,
  type FontResolver,
  type FontResolverFace,
  type FontResolverOutcome,
  type FontResolverResponse,
} from "./font-contract.js";

export interface BrowserFontLimits {
  fontFiles: number;
  fontFileBytes: number;
  fontTotalBytes: number;
  fontRequests: number;
  fontSampleCodePoints: number;
}

export interface BrowserFontResult {
  identity: FontConfigurationIdentity;
  resolutions: FontResolution[];
  renderedTextNodeCount: number;
}

export interface BrowserFontTask {
  pending(): string[];
  wait(signal: AbortSignal): Promise<BrowserFontResult>;
}

export class BrowserFontError extends Error {
  readonly kind: "invalid_response" | "resource_limit";
  readonly detail?: string;

  constructor(
    kind: BrowserFontError["kind"],
    message: string,
    detail?: string,
  ) {
    super(message);
    this.name = "BrowserFontError";
    this.kind = kind;
    this.detail = detail;
  }
}

interface TextFaceUse {
  element: HTMLElement;
  requestKey: string;
  originalStyle: string | null;
}

interface MutableRequest {
  familyStack: string[];
  familyKinds: FontFamilyKind[];
  style: FontFaceStyle;
  weight: number;
  stretch: number;
  sampleCodePoints: Set<number>;
}

interface FontInventory {
  requests: readonly FontRequest[];
  uses: TextFaceUse[];
  renderedTextNodeCount: number;
}

interface ValidatedFace {
  record: FontResolverFace;
  bytes: Uint8Array;
}

interface ValidatedResponse {
  response: Omit<FontResolverResponse, "faces" | "outcomes"> & {
    faces: readonly FontResolverFace[];
    outcomes: readonly FontResolverOutcome[];
  };
  faces: Map<string, ValidatedFace>;
}

const TEXT_ENCODER = new TextEncoder();
const MAX_FAMILY_COUNT = 64;
const MAX_FAMILY_CHARACTERS = 256;
const MAX_FAMILY_STACK_CHARACTERS = 4096;
const FONT_LOAD_CONCURRENCY = 16;
const GENERIC_FAMILIES = new Set([
  "cursive", "emoji", "fangsong", "fantasy", "math", "monospace", "sans-serif",
  "serif", "system-ui", "ui-monospace", "ui-rounded", "ui-sans-serif", "ui-serif",
]);
const FONT_STRETCH_PERCENT = new Map<string, number>([
  ["ultra-condensed", 50],
  ["extra-condensed", 62.5],
  ["condensed", 75],
  ["semi-condensed", 87.5],
  ["normal", 100],
  ["semi-expanded", 112.5],
  ["expanded", 125],
  ["extra-expanded", 150],
  ["ultra-expanded", 200],
]);
const FORMAT_MEDIA = new Map<FontFileFormat, string>([
  ["ttf", "font/ttf"],
  ["otf", "font/otf"],
  ["woff", "font/woff"],
  ["woff2", "font/woff2"],
]);
const FORMAT_HINT = new Map<FontFileFormat, string>([
  ["ttf", "truetype"],
  ["otf", "opentype"],
  ["woff", "woff"],
  ["woff2", "woff2"],
]);

function compareText(left: string, right: string): number {
  return left < right ? -1 : left > right ? 1 : 0;
}

function isWellFormedUnicode(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) {
      return false;
    }
  }
  return true;
}

function canonicalValue(value: unknown): unknown {
  if (value === null || typeof value === "boolean") return value;
  if (typeof value === "string") {
    if (!isWellFormedUnicode(value)) {
      throw new TypeError("Canonical JSON does not support unpaired UTF-16 surrogates");
    }
    return value;
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) throw new TypeError("Canonical JSON does not support non-finite numbers");
    return Object.is(value, -0) ? 0 : value;
  }
  if (Array.isArray(value)) return value.map(canonicalValue);
  if (typeof value === "object") {
    const result: Record<string, unknown> = {};
    for (const key of Object.keys(value as Record<string, unknown>).sort(compareText)) {
      const member = (value as Record<string, unknown>)[key];
      if (member !== undefined) result[key] = canonicalValue(member);
    }
    return result;
  }
  throw new TypeError(`Canonical JSON does not support ${typeof value}`);
}

function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalValue(value));
}

async function sha256(bytes: Uint8Array): Promise<string> {
  if (!globalThis.crypto?.subtle) {
    throw new BrowserFontError("invalid_response", "Web Crypto SHA-256 is unavailable.");
  }
  const owned = new Uint8Array(bytes);
  const digest = await globalThis.crypto.subtle.digest("SHA-256", owned.buffer);
  return Array.from(new Uint8Array(digest), (value) => value.toString(16).padStart(2, "0")).join("");
}

async function digestJson(value: unknown): Promise<string> {
  return sha256(TEXT_ENCODER.encode(canonicalJson(value)));
}

async function abortable<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  if (signal.aborted) throw new DOMException("Font loading was aborted", "AbortError");
  let rejectAbort: ((reason?: unknown) => void) | undefined;
  const aborted = new Promise<never>((_, reject) => { rejectAbort = reject; });
  const onAbort = (): void => rejectAbort?.(new DOMException("Font loading was aborted", "AbortError"));
  signal.addEventListener("abort", onAbort, { once: true });
  try {
    return await Promise.race([promise, aborted]);
  } finally {
    signal.removeEventListener("abort", onAbort);
  }
}

async function forEachBounded<T>(
  values: readonly T[],
  operation: (value: T, index: number) => Promise<void>,
): Promise<void> {
  let cursor = 0;
  const worker = async (): Promise<void> => {
    while (cursor < values.length) {
      const index = cursor++;
      await operation(values[index], index);
    }
  };
  await Promise.all(Array.from(
    { length: Math.min(FONT_LOAD_CONCURRENCY, values.length) },
    () => worker(),
  ));
}

function cssEscape(source: string, start: number): { value: string; end: number } {
  let cursor = start + 1;
  if (cursor >= source.length) return { value: "\ufffd", end: cursor };
  if (source[cursor] === "\r" && source[cursor + 1] === "\n") return { value: "", end: cursor + 2 };
  if (source[cursor] === "\n" || source[cursor] === "\r" || source[cursor] === "\f") {
    return { value: "", end: cursor + 1 };
  }
  const hexStart = cursor;
  while (cursor < source.length && cursor - hexStart < 6 && /[0-9a-f]/i.test(source[cursor])) cursor++;
  if (cursor > hexStart) {
    const point = Number.parseInt(source.slice(hexStart, cursor), 16);
    if (/\s/.test(source[cursor] ?? "")) {
      if (source[cursor] === "\r" && source[cursor + 1] === "\n") cursor += 2;
      else cursor++;
    }
    return {
      value: point === 0 || point > 0x10ffff || (point >= 0xd800 && point <= 0xdfff)
        ? "\ufffd"
        : String.fromCodePoint(point),
      end: cursor,
    };
  }
  return { value: source[cursor], end: cursor + 1 };
}

interface ParsedFontFamily {
  name: string;
  kind: FontFamilyKind;
}

function parseCssFontFamilyTokens(value: string): ParsedFontFamily[] {
  if (!isWellFormedUnicode(value)) {
    throw new BrowserFontError(
      "invalid_response",
      "A computed font-family value contains an unpaired UTF-16 surrogate.",
    );
  }
  const families: ParsedFontFamily[] = [];
  let familyCharacters = 0;
  let family = "";
  let quote = "";
  let quoted = false;
  let cursor = 0;
  const finish = (): void => {
    const normalized = normalizeFontFamilyName(family);
    if (normalized) {
      if (!isWellFormedUnicode(normalized)
        || normalized.length > MAX_FAMILY_CHARACTERS
        || /[\u0000-\u001f\u007f]/u.test(normalized)) {
        throw new BrowserFontError(
          "resource_limit",
          `A computed font family must contain at most ${MAX_FAMILY_CHARACTERS} bounded characters.`,
          "fontFamilyCharacters",
        );
      }
      if (families.length >= MAX_FAMILY_COUNT) {
        throw new BrowserFontError(
          "resource_limit",
          `A computed font-family stack may contain at most ${MAX_FAMILY_COUNT} families.`,
          "fontFamilyCount",
        );
      }
      familyCharacters += normalized.length;
      if (familyCharacters > MAX_FAMILY_STACK_CHARACTERS) {
        throw new BrowserFontError(
          "resource_limit",
          `A computed font-family stack may contain at most ${MAX_FAMILY_STACK_CHARACTERS} characters.`,
          "fontFamilyCharacters",
        );
      }
      families.push({
        name: normalized,
        kind: quoted || !GENERIC_FAMILIES.has(fontFamilyKey(normalized)) ? "named" : "generic",
      });
    }
    family = "";
    quoted = false;
  };
  while (cursor < value.length) {
    const character = value[cursor];
    if (character === "\\") {
      const escape = cssEscape(value, cursor);
      family += escape.value;
      cursor = escape.end;
      continue;
    }
    if (quote) {
      if (character === quote) quote = "";
      else family += character;
      cursor++;
      continue;
    }
    if (character === "\"" || character === "'") {
      quote = character;
      quoted = true;
      cursor++;
      continue;
    }
    if (character === ",") {
      finish();
      cursor++;
      continue;
    }
    family += character;
    cursor++;
  }
  finish();
  return families;
}

/** Parse a computed CSS font-family value without splitting quoted or escaped commas. */
export function parseCssFontFamily(value: string): string[] {
  return parseCssFontFamilyTokens(value).map(({ name }) => name);
}

function faceStyle(value: string): FontFaceStyle {
  const key = value.trim().toLowerCase();
  if (key.startsWith("italic")) return "italic";
  if (key.startsWith("oblique")) return "oblique";
  return "normal";
}

function faceWeight(value: string): number {
  if (value === "bold") return 700;
  if (value === "normal") return 400;
  const weight = Number.parseFloat(value);
  return Number.isFinite(weight) ? Math.max(1, Math.min(1000, Math.round(weight))) : 400;
}

function faceStretch(value: string): number {
  const key = value.trim().toLowerCase();
  const named = FONT_STRETCH_PERCENT.get(key);
  if (named !== undefined) return named;
  const percentage = /^([0-9]+(?:\.[0-9]+)?)%$/.exec(key);
  if (!percentage) return 100;
  const stretch = Number.parseFloat(percentage[1]);
  return Number.isFinite(stretch) && stretch > 0 ? stretch : 100;
}

function requestKey(request: Omit<MutableRequest, "sampleCodePoints">): string {
  return canonicalJson({
    familyStack: request.familyStack,
    familyKinds: request.familyKinds,
    stretch: request.stretch,
    style: request.style,
    weight: request.weight,
  });
}

/**
 * Whether the element is *being rendered* in the HTML sense: it generates a painted box and
 * therefore resolves a font. Content the author hid contributes no glyphs to the output and
 * is deliberately left out of the inventory.
 *
 * The converter's measurement staging area is itself `visibility: hidden`, and at
 * `font_loading` every piece of document content still lives inside it. The export pipeline
 * lifts that one container's own visibility for the duration of the phase (see
 * `revealMeasurementStaging` in export-browser.ts) so this predicate keeps its meaning:
 * a descendant that the document itself hid still computes to hidden and stays out.
 */
function participatesInRendering(element: HTMLElement, view: Window): boolean {
  const leaf = view.getComputedStyle(element);
  if (leaf.visibility === "hidden" || leaf.visibility === "collapse") return false;
  for (let current: HTMLElement | null = element; current; current = current.parentElement) {
    const computed = current === element ? leaf : view.getComputedStyle(current);
    if (computed.display === "none" || computed.contentVisibility === "hidden") return false;
  }
  return true;
}

function collectFontInventory(document: Document, limits: BrowserFontLimits): FontInventory {
  const view = document.defaultView;
  if (!view || !document.body) throw new BrowserFontError("invalid_response", "The render document has no font realm.");
  const requests = new Map<string, MutableRequest>();
  const uses: TextFaceUse[] = [];
  let sampledCodePoints = 0;
  let renderedTextNodeCount = 0;
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? "";
    const element = node.parentElement;
    if (!text || !element || /^(?:script|style|template|noscript)$/i.test(element.localName)) continue;
    if (!participatesInRendering(element, view)) continue;
    renderedTextNodeCount++;
    const computed = view.getComputedStyle(element);
    const parsedFamilies = parseCssFontFamilyTokens(computed.fontFamily);
    if (parsedFamilies.length === 0) continue;
    const descriptor = {
      familyStack: parsedFamilies.map(({ name }) => name),
      familyKinds: parsedFamilies.map(({ kind }) => kind),
      style: faceStyle(computed.fontStyle),
      weight: faceWeight(computed.fontWeight),
      stretch: faceStretch(computed.fontStretch),
    };
    const key = requestKey(descriptor);
    let request = requests.get(key);
    if (!request) {
      if (requests.size >= limits.fontRequests) {
        throw new BrowserFontError(
          "resource_limit",
          `fontRequests limit exceeded (${requests.size + 1} > ${limits.fontRequests}).`,
          "fontRequests",
        );
      }
      request = { ...descriptor, sampleCodePoints: new Set<number>() };
      requests.set(key, request);
    }
    for (const character of text) {
      const scalar = character.codePointAt(0)!;
      const point = scalar >= 0xd800 && scalar <= 0xdfff ? 0xfffd : scalar;
      if (request.sampleCodePoints.has(point)) continue;
      sampledCodePoints++;
      if (sampledCodePoints > limits.fontSampleCodePoints) {
        throw new BrowserFontError(
          "resource_limit",
          `fontSampleCodePoints limit exceeded (${sampledCodePoints} > ${limits.fontSampleCodePoints}).`,
          "fontSampleCodePoints",
        );
      }
      request.sampleCodePoints.add(point);
    }
    uses.push({ element, requestKey: key, originalStyle: element.getAttribute("style") });
  }
  const ordered = Array.from(requests, ([key, request]) => ({ key, request }))
    .sort((left, right) => compareText(left.key, right.key));
  const idByKey = new Map<string, string>();
  const result = Object.freeze(ordered.map(({ key, request }, index) => {
    const id = `font-${String(index + 1).padStart(4, "0")}`;
    idByKey.set(key, id);
    return Object.freeze({
      id,
      familyStack: Object.freeze([...request.familyStack]),
      familyKinds: Object.freeze([...request.familyKinds]),
      style: request.style,
      weight: request.weight,
      stretch: request.stretch,
      sampleCodePoints: Object.freeze(Array.from(request.sampleCodePoints).sort((left, right) => left - right)),
    });
  }));
  return {
    requests: result,
    uses: uses.map((use) => ({ ...use, requestKey: idByKey.get(use.requestKey)! })),
    renderedTextNodeCount,
  };
}

/** Browser-readable request inventory; element associations remain internal. */
export function inventoryDocumentFontRequests(
  document: Document,
  limits: Pick<BrowserFontLimits, "fontRequests" | "fontSampleCodePoints">,
): FontRequest[] {
  return [...collectFontInventory(document, {
    ...limits,
    fontFiles: Number.MAX_SAFE_INTEGER,
    fontFileBytes: Number.MAX_SAFE_INTEGER,
    fontTotalBytes: Number.MAX_SAFE_INTEGER,
  }).requests];
}

function requireObject(value: unknown, label: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new BrowserFontError("invalid_response", `${label} must be an object.`);
  }
  return value as Record<string, unknown>;
}

function exactKeys(record: Record<string, unknown>, allowed: readonly string[], label: string): void {
  const unknown = Object.keys(record).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) {
    throw new BrowserFontError("invalid_response", `${label} contains unknown fields: ${unknown.sort(compareText).join(", ")}.`);
  }
}

function shortString(value: unknown, label: string, maximum = 512): string {
  if (typeof value !== "string" || value.trim() === "" || value.length > maximum
    || !isWellFormedUnicode(value) || /[\u0000-\u001f\u007f]/u.test(value)) {
    throw new BrowserFontError("invalid_response", `${label} must be a non-empty string of at most ${maximum} characters.`);
  }
  return value;
}

function digestString(value: unknown, label: string): string {
  if (typeof value !== "string" || !/^[0-9a-f]{64}$/.test(value)) {
    throw new BrowserFontError("invalid_response", `${label} must be a lowercase SHA-256 digest.`);
  }
  return value;
}

function boundedNumber(value: unknown, label: string, minimum: number, maximum: number): number {
  if (typeof value !== "number" || !Number.isFinite(value)
    || value < minimum || value > maximum) {
    throw new BrowserFontError(
      "invalid_response",
      `${label} must be a finite number from ${minimum} through ${maximum}.`,
    );
  }
  return value;
}

function integer(value: unknown, label: string, minimum: number, maximum: number): number {
  if (!Number.isSafeInteger(value) || (value as number) < minimum || (value as number) > maximum) {
    throw new BrowserFontError("invalid_response", `${label} must be an integer from ${minimum} through ${maximum}.`);
  }
  return value as number;
}

function canonicalBase64(value: unknown, expectedBytes: number, label: string): Uint8Array {
  const expectedCharacters = 4 * Math.ceil(expectedBytes / 3);
  if (typeof value !== "string" || value.length !== expectedCharacters || value.length % 4 !== 0
    || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) {
    throw new BrowserFontError("invalid_response", `${label} must be canonical padded base64 without whitespace.`);
  }
  let decoded: string;
  try {
    decoded = globalThis.atob(value);
  } catch (cause) {
    throw new BrowserFontError("invalid_response", `${label} is not valid base64.`, String(cause));
  }
  if (decoded.length !== expectedBytes) {
    throw new BrowserFontError("invalid_response", `${label} byte length does not match its metadata.`);
  }
  const bytes = new Uint8Array(decoded.length);
  for (let index = 0; index < decoded.length; index++) bytes[index] = decoded.charCodeAt(index);
  let binary = "";
  for (let index = 0; index < bytes.length; index += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(index, index + 0x8000));
  }
  if (globalThis.btoa(binary) !== value) {
    throw new BrowserFontError("invalid_response", `${label} is not canonical base64.`);
  }
  return bytes;
}

function readU32(bytes: Uint8Array, offset: number): number {
  if (bytes.byteLength < offset + 4) {
    throw new BrowserFontError("invalid_response", "A configured webfont has a truncated format header.");
  }
  return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(offset, false);
}

function expandedFontByteLength(
  bytes: Uint8Array,
  format: FontFileFormat,
  label: string,
  maximum: number,
): number {
  if (format !== "woff" && format !== "woff2") return bytes.byteLength;
  const declaredLength = readU32(bytes, 8);
  if (declaredLength !== bytes.byteLength) {
    throw new BrowserFontError("invalid_response", `${label} declared length does not match its decoded bytes.`);
  }
  const expanded = readU32(bytes, 16);
  if (expanded === 0) {
    throw new BrowserFontError("invalid_response", `${label} declares an invalid expanded size.`);
  }
  if (expanded > maximum) {
    throw new BrowserFontError(
      "resource_limit",
      `fontFileBytes limit exceeded by ${label} expanded bytes (${expanded} > ${maximum}).`,
      "fontFileBytes",
    );
  }
  return expanded;
}

function fontSignatureMatches(bytes: Uint8Array, format: FontFileFormat): boolean {
  if (bytes.length < 4) return false;
  const signature = String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]);
  if (format === "otf") return signature === "OTTO";
  if (format === "woff") return signature === "wOFF";
  if (format === "woff2") return signature === "wOF2";
  return (bytes[0] === 0 && bytes[1] === 1 && bytes[2] === 0 && bytes[3] === 0)
    || signature === "true";
}

function scalarArray(value: unknown, label: string, request: FontRequest): number[] {
  if (!Array.isArray(value)) throw new BrowserFontError("invalid_response", `${label} must be an array.`);
  const requestPoints = new Set(request.sampleCodePoints);
  const result = value.map((point, index) => integer(point, `${label}[${index}]`, 0, 0x10ffff));
  if (result.some((point) => point >= 0xd800 && point <= 0xdfff)
    || result.some((point, index) => index > 0 && point <= result[index - 1])
    || result.some((point) => !requestPoints.has(point))) {
    throw new BrowserFontError("invalid_response", `${label} must be sorted, distinct scalar values from the request sample.`);
  }
  return result;
}

function snapshotFields(
  value: unknown,
  allowed: readonly string[],
  label: string,
): Readonly<Record<string, unknown>> {
  const record = requireObject(value, label);
  exactKeys(record, allowed, label);
  const snapshot: Record<string, unknown> = {};
  for (const key of allowed) {
    if (Object.prototype.hasOwnProperty.call(record, key)) snapshot[key] = record[key];
  }
  return Object.freeze(snapshot);
}

function snapshotResolverResponse(
  value: unknown,
  requestCount: number,
  limits: BrowserFontLimits,
): Readonly<Record<string, unknown>> {
  const response = snapshotFields(value, [
    "schemaVersion", "resolverContract", "substitutionContractVersion",
    "substitutionContractDigest", "outcomes", "faces",
  ], "font resolver response");
  const faceValues = response.faces;
  if (!Array.isArray(faceValues)) {
    throw new BrowserFontError("invalid_response", "font resolver response faces must be an array.");
  }
  if (faceValues.length > limits.fontFiles) {
    throw new BrowserFontError(
      "resource_limit",
      `fontFiles limit exceeded (${faceValues.length} > ${limits.fontFiles}).`,
      "fontFiles",
    );
  }
  const faceSnapshots: Readonly<Record<string, unknown>>[] = [];
  for (let index = 0; index < faceValues.length; index++) {
    const value = faceValues[index];
    const face = snapshotFields(value, [
      "id", "resolvedFamily", "postscriptName", "version", "style", "weight", "stretch",
      "format", "mediaType", "byteLength", "sha256", "bytesBase64", "licenseEvidence",
    ], `faces[${index}]`);
    const evidence = snapshotFields(
      face.licenseEvidence,
      ["kind", "identity", "noSubsetting"],
      `faces[${index}].licenseEvidence`,
    );
    faceSnapshots.push(Object.freeze({ ...face, licenseEvidence: evidence }));
  }
  const faces = Object.freeze(faceSnapshots);
  const outcomeValues = response.outcomes;
  if (!Array.isArray(outcomeValues) || outcomeValues.length !== requestCount) {
    throw new BrowserFontError("invalid_response", "The font resolver must return exactly one outcome per request.");
  }
  const outcomeSnapshots: Readonly<Record<string, unknown>>[] = [];
  for (let index = 0; index < outcomeValues.length; index++) {
    const value = outcomeValues[index];
    const outcome = snapshotFields(value, [
      "requestId", "status", "faceId", "requestedFamily", "resolvedFamily", "metricCompatible",
      "faceMatch", "glyphCoverage", "missingCodePoints",
    ], `outcomes[${index}]`);
    const missing = outcome.missingCodePoints;
    if (missing !== undefined && !Array.isArray(missing)) {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].missingCodePoints must be an array.`);
    }
    if (Array.isArray(missing) && missing.length > limits.fontSampleCodePoints) {
      throw new BrowserFontError(
        "resource_limit",
        `fontSampleCodePoints limit exceeded by outcomes[${index}].missingCodePoints.`,
        "fontSampleCodePoints",
      );
    }
    let missingSnapshot: readonly unknown[] | undefined;
    if (Array.isArray(missing)) {
      const copied: unknown[] = [];
      for (let missingIndex = 0; missingIndex < missing.length; missingIndex++) {
        copied.push(missing[missingIndex]);
      }
      missingSnapshot = Object.freeze(copied);
    }
    outcomeSnapshots.push(Object.freeze({
      ...outcome,
      ...(missingSnapshot ? { missingCodePoints: missingSnapshot } : {}),
    }));
  }
  const outcomes = Object.freeze(outcomeSnapshots);
  return Object.freeze({ ...response, faces, outcomes });
}

async function validateResolverResponse(
  value: unknown,
  requests: readonly FontRequest[],
  limits: BrowserFontLimits,
  expectedContractDigest: string,
): Promise<ValidatedResponse> {
  const record = snapshotResolverResponse(value, requests.length, limits);
  if (record.schemaVersion !== FONT_RESOLVER_SCHEMA_VERSION
    || record.resolverContract !== FONT_RESOLVER_CONTRACT_ID
    || record.substitutionContractVersion !== FONT_SUBSTITUTION_CONTRACT_VERSION
    || record.substitutionContractDigest !== expectedContractDigest) {
    throw new BrowserFontError("invalid_response", "The font resolver response uses a different contract identity.");
  }
  const faceValues = record.faces as readonly Readonly<Record<string, unknown>>[];
  const faces = new Map<string, ValidatedFace>();
  let totalBytes = 0;
  let totalExpandedBytes = 0;
  for (const [index, face] of faceValues.entries()) {
    const id = shortString(face.id, `faces[${index}].id`, 128);
    if (faces.has(id)) throw new BrowserFontError("invalid_response", `Duplicate configured face id: ${id}.`);
    const format = face.format;
    if (format !== "ttf" && format !== "otf" && format !== "woff" && format !== "woff2") {
      throw new BrowserFontError("invalid_response", `faces[${index}].format is unsupported.`);
    }
    if (face.mediaType !== FORMAT_MEDIA.get(format)) {
      throw new BrowserFontError("invalid_response", `faces[${index}].mediaType does not match its format.`);
    }
    const byteLength = integer(face.byteLength, `faces[${index}].byteLength`, 1, limits.fontFileBytes);
    totalBytes += byteLength;
    if (totalBytes > limits.fontTotalBytes) {
      throw new BrowserFontError("resource_limit", `fontTotalBytes limit exceeded (${totalBytes} > ${limits.fontTotalBytes}).`, "fontTotalBytes");
    }
    const bytesBase64 = face.bytesBase64;
    const bytes = canonicalBase64(bytesBase64, byteLength, `faces[${index}].bytesBase64`);
    if (!fontSignatureMatches(bytes, format)) {
      throw new BrowserFontError("invalid_response", `faces[${index}] decoded signature does not match its format.`);
    }
    const expandedBytes = expandedFontByteLength(
      bytes,
      format,
      `faces[${index}]`,
      limits.fontFileBytes,
    );
    totalExpandedBytes += expandedBytes;
    if (totalExpandedBytes > limits.fontTotalBytes) {
      throw new BrowserFontError(
        "resource_limit",
        `fontTotalBytes limit exceeded by expanded font bytes (${totalExpandedBytes} > ${limits.fontTotalBytes}).`,
        "fontTotalBytes",
      );
    }
    const expectedDigest = digestString(face.sha256, `faces[${index}].sha256`);
    if (await sha256(bytes) !== expectedDigest) {
      throw new BrowserFontError("invalid_response", `faces[${index}] bytes do not match sha256.`);
    }
    const evidence = face.licenseEvidence as Readonly<Record<string, unknown>>;
    if (evidence.kind !== "installable" && evidence.kind !== "previewPrint"
      && evidence.kind !== "editable" && evidence.kind !== "attested") {
      throw new BrowserFontError("invalid_response", `faces[${index}].licenseEvidence.kind is invalid.`);
    }
    if (evidence.noSubsetting !== true && evidence.noSubsetting !== false) {
      throw new BrowserFontError("invalid_response", `faces[${index}].licenseEvidence.noSubsetting must be boolean.`);
    }
    const style = face.style;
    if (style !== "normal" && style !== "italic" && style !== "oblique") {
      throw new BrowserFontError("invalid_response", `faces[${index}].style is invalid.`);
    }
    const licenseEvidence = Object.freeze({
      kind: evidence.kind,
      identity: digestString(evidence.identity, `faces[${index}].licenseEvidence.identity`),
      noSubsetting: evidence.noSubsetting,
    }) as FontResolverFace["licenseEvidence"];
    const resolved: FontResolverFace = Object.freeze({
      id,
      resolvedFamily: normalizeFontFamilyName(shortString(
        face.resolvedFamily,
        `faces[${index}].resolvedFamily`,
        MAX_FAMILY_CHARACTERS,
      )),
      ...(face.postscriptName === undefined
        ? {}
        : { postscriptName: shortString(face.postscriptName, `faces[${index}].postscriptName`) }),
      version: shortString(face.version, `faces[${index}].version`),
      style,
      weight: integer(face.weight, `faces[${index}].weight`, 1, 1000),
      stretch: boundedNumber(face.stretch, `faces[${index}].stretch`, 50, 200),
      format,
      mediaType: FORMAT_MEDIA.get(format)! as FontResolverFace["mediaType"],
      byteLength,
      sha256: expectedDigest,
      bytesBase64: bytesBase64 as string,
      licenseEvidence,
    });
    faces.set(id, { record: resolved, bytes });
  }
  const outcomeValues = record.outcomes as readonly Readonly<Record<string, unknown>>[];
  const requestById = new Map(requests.map((request) => [request.id, request]));
  const seen = new Set<string>();
  const outcomes: FontResolverOutcome[] = [];
  for (const [index, outcome] of outcomeValues.entries()) {
    const requestId = shortString(outcome.requestId, `outcomes[${index}].requestId`, 128);
    const request = requestById.get(requestId);
    if (!request || seen.has(requestId)) {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].requestId is unknown or duplicated.`);
    }
    seen.add(requestId);
    const status = outcome.status;
    if (status !== "resolved" && status !== "substituted" && status !== "missing" && status !== "unverified") {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].status is invalid.`);
    }
    const faceId = outcome.faceId === undefined
      ? undefined
      : shortString(outcome.faceId, `outcomes[${index}].faceId`, 128);
    if ((status === "resolved" || status === "substituted") !== (faceId !== undefined)) {
      throw new BrowserFontError("invalid_response", `outcomes[${index}] has an inconsistent face selection.`);
    }
    const selectedFace = faceId === undefined ? undefined : faces.get(faceId)?.record;
    if (faceId && !selectedFace) throw new BrowserFontError("invalid_response", `outcomes[${index}] references an unknown face.`);
    const requestedFamily = outcome.requestedFamily === undefined
      ? request.familyStack[0]
      : normalizeFontFamilyName(shortString(
        outcome.requestedFamily,
        `outcomes[${index}].requestedFamily`,
        MAX_FAMILY_CHARACTERS,
      ));
    if (!request.familyStack.some((family) => fontFamilyKey(family) === fontFamilyKey(requestedFamily))) {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].requestedFamily is not in the request stack.`);
    }
    const resolvedFamily = outcome.resolvedFamily === undefined
      ? selectedFace?.resolvedFamily
      : normalizeFontFamilyName(shortString(
        outcome.resolvedFamily,
        `outcomes[${index}].resolvedFamily`,
        MAX_FAMILY_CHARACTERS,
      ));
    if (selectedFace && (!resolvedFamily || fontFamilyKey(resolvedFamily) !== fontFamilyKey(selectedFace.resolvedFamily))) {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].resolvedFamily does not match its face.`);
    }
    if (status === "resolved"
      && (request.familyKinds[0] !== "named"
        || fontFamilyKey(requestedFamily) !== fontFamilyKey(request.familyStack[0])
        || !resolvedFamily
        || fontFamilyKey(resolvedFamily) !== fontFamilyKey(request.familyStack[0]))) {
      throw new BrowserFontError(
        "invalid_response",
        `outcomes[${index}] resolved status must match the request's primary family.`,
      );
    }
    if (outcome.metricCompatible !== undefined && typeof outcome.metricCompatible !== "boolean") {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].metricCompatible must be boolean.`);
    }
    if (outcome.faceMatch !== undefined && outcome.faceMatch !== "exact" && outcome.faceMatch !== "synthesized") {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].faceMatch is invalid.`);
    }
    if (outcome.glyphCoverage !== undefined
      && outcome.glyphCoverage !== "complete" && outcome.glyphCoverage !== "partial"
      && outcome.glyphCoverage !== "unverified") {
      throw new BrowserFontError("invalid_response", `outcomes[${index}].glyphCoverage is invalid.`);
    }
    const missingCodePoints = outcome.missingCodePoints === undefined
      ? undefined
      : scalarArray(outcome.missingCodePoints, `outcomes[${index}].missingCodePoints`, request);
    const missingCodePointCount = missingCodePoints?.length ?? 0;
    if (outcome.glyphCoverage === "complete" && missingCodePointCount > 0) {
      throw new BrowserFontError(
        "invalid_response",
        `outcomes[${index}] complete coverage cannot name missing code points.`,
      );
    }
    if ((outcome.glyphCoverage === "partial" && missingCodePointCount === 0)
      || (missingCodePointCount > 0 && outcome.glyphCoverage !== "partial")) {
      throw new BrowserFontError(
        "invalid_response",
        `outcomes[${index}] partial coverage must exactly match a non-empty missing-code-point list.`,
      );
    }
    if (selectedFace && missingCodePointCount === 0 && outcome.glyphCoverage !== "complete") {
      throw new BrowserFontError(
        "invalid_response",
        `outcomes[${index}] must declare complete coverage when no code points are missing.`,
      );
    }
    if (outcome.faceMatch === "exact" && selectedFace
      && (selectedFace.style !== request.style
        || selectedFace.weight !== request.weight
        || selectedFace.stretch !== request.stretch)) {
      throw new BrowserFontError(
        "invalid_response",
        `outcomes[${index}] claims an exact face with different style, weight, or stretch.`,
      );
    }
    if (status === "missing" || status === "unverified") {
      const hasSelectionMetadata = outcome.resolvedFamily !== undefined
        || outcome.metricCompatible !== undefined
        || outcome.faceMatch !== undefined
        || outcome.missingCodePoints !== undefined
        || (outcome.glyphCoverage !== undefined && outcome.glyphCoverage !== "unverified");
      if (hasSelectionMetadata) {
        throw new BrowserFontError(
          "invalid_response",
          `outcomes[${index}] cannot attach selection metadata to ${status} status.`,
        );
      }
    }
    outcomes.push(Object.freeze({
      requestId,
      status,
      ...(faceId ? { faceId } : {}),
      requestedFamily,
      ...(resolvedFamily ? { resolvedFamily } : {}),
      ...(outcome.metricCompatible === undefined ? {} : { metricCompatible: outcome.metricCompatible }),
      ...(outcome.faceMatch === undefined ? {} : { faceMatch: outcome.faceMatch }),
      ...(outcome.glyphCoverage === undefined ? {} : { glyphCoverage: outcome.glyphCoverage }),
      ...(missingCodePoints === undefined ? {} : { missingCodePoints: Object.freeze(missingCodePoints) }),
    }));
  }
  const referencedFaces = new Set(outcomes.flatMap((outcome) => outcome.faceId ? [outcome.faceId] : []));
  if (Array.from(faces.keys()).some((id) => !referencedFaces.has(id))) {
    throw new BrowserFontError("invalid_response", "The font resolver returned an unreferenced face.");
  }
  outcomes.sort((left, right) => compareText(left.requestId, right.requestId));
  const faceRecords = Array.from(faces.values(), ({ record }) => record)
    .sort((left, right) => compareText(left.id, right.id));
  return {
    response: Object.freeze({
      schemaVersion: FONT_RESOLVER_SCHEMA_VERSION,
      resolverContract: FONT_RESOLVER_CONTRACT_ID,
      substitutionContractVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
      substitutionContractDigest: expectedContractDigest,
      outcomes: Object.freeze(outcomes),
      faces: Object.freeze(faceRecords),
    }),
    faces,
  };
}

function cssString(value: string): string {
  return `"${value.replace(/\\/g, "\\\\").replace(/"/g, "\\\"")
    .replace(/[\n\r\f]/g, (character) => `\\${character.codePointAt(0)!.toString(16)} `)}"`;
}

function serializeFamilyStack(
  families: readonly string[],
  kinds: readonly FontFamilyKind[] = families.map(() => "named" as const),
): string {
  return families.map((family, index) => kinds[index] === "generic"
    ? fontFamilyKey(family)
    : cssString(family)).join(", ");
}

function fontSpecification(
  request: FontRequest,
  families: readonly string[],
  kinds?: readonly FontFamilyKind[],
): string {
  // Chromium's FontFace descriptor accepts percentage stretches, but
  // FontFaceSet.load() still parses the legacy font shorthand and rejects
  // percentage tokens. Preserve exact standard widths through their keyword;
  // an uncommon percentage remains enforced by the installed @font-face rule.
  const stretch = Array.from(FONT_STRETCH_PERCENT).find(([, value]) => value === request.stretch)?.[0];
  return `${request.style} ${request.weight}${stretch ? ` ${stretch}` : ""} 12px ${serializeFamilyStack(families, kinds)}`;
}

function responseIdentityMaterial(response: ValidatedResponse["response"]): unknown {
  return {
    schemaVersion: response.schemaVersion,
    resolverContract: response.resolverContract,
    substitutionContractVersion: response.substitutionContractVersion,
    substitutionContractDigest: response.substitutionContractDigest,
    outcomes: response.outcomes,
    faces: response.faces.map(({ bytesBase64: _bytes, ...face }) => face),
  };
}

async function syntheticFamily(
  resolverDigest: string,
  faceId: string,
): Promise<string> {
  const mapping = await digestJson({
    faceId,
    resolverDigest,
  });
  return `__DocxodusConfigured_${resolverDigest.slice(0, 16)}_${mapping.slice(0, 16)}`;
}

function restoreUse(use: TextFaceUse): void {
  if (use.originalStyle === null) use.element.removeAttribute("style");
  else use.element.setAttribute("style", use.originalStyle);
}

function faceRule(family: string, face: FontResolverFace): string {
  return `@font-face{font-family:${cssString(family)};src:url("data:${face.mediaType};base64,${face.bytesBase64}") format("${FORMAT_HINT.get(face.format)}");font-style:${face.style};font-weight:${face.weight};font-stretch:${face.stretch}%;font-display:block}`;
}

function baseResolution(request: FontRequest): Pick<FontResolution,
  "requestId" | "requestedFamily" | "requestedFamilies" | "requestedFamilyKinds"
  | "requestedStyle" | "requestedWeight"
  | "requestedStretch" | "sampleCodePointCount"> {
  return {
    requestId: request.id,
    requestedFamily: request.familyStack[0],
    requestedFamilies: [...request.familyStack],
    requestedFamilyKinds: [...request.familyKinds],
    requestedStyle: request.style,
    requestedWeight: request.weight,
    requestedStretch: request.stretch,
    sampleCodePointCount: request.sampleCodePoints.length,
  };
}

// Availability probing needs glyphs whose advance widths differ between typefaces;
// digits and mixed-case letters separate metric-compatible families far better than
// a single repeated character does.
const FONT_AVAILABILITY_SAMPLE = "MWmwilAaGg0189";
const FONT_AVAILABILITY_FALLBACKS = ["monospace", "serif", "sans-serif"] as const;

/**
 * Whether the render environment can actually paint `family`.
 *
 * `FontFaceSet.check()` cannot answer this. It reports whether pending downloads have
 * settled, so Chromium returns true for a family it has never heard of — every missing
 * font would be recorded as merely unverified. Measuring instead: a family that moves the
 * advance width away from every generic fallback is present, and one that matches all
 * three is being silently substituted for.
 */
function familyAvailable(document: Document, family: string, kind: FontFamilyKind): boolean {
  // A generic family always resolves to whatever the environment maps it to.
  if (kind === "generic") return true;
  const context = document.createElement("canvas").getContext("2d");
  if (!context) return true;
  const token = serializeFamilyStack([family], ["named"]);
  for (const fallback of FONT_AVAILABILITY_FALLBACKS) {
    context.font = `72px ${fallback}`;
    const baseline = context.measureText(FONT_AVAILABILITY_SAMPLE).width;
    context.font = `72px ${token}, ${fallback}`;
    if (context.measureText(FONT_AVAILABILITY_SAMPLE).width !== baseline) return true;
  }
  return false;
}

/**
 * The first family a request names is the one the document asked for; anything after it is
 * already a substitution the author accepted, and a trailing generic always resolves.
 */
function requestedFamilyAvailable(document: Document, request: FontRequest): boolean {
  const index = request.familyStack.findIndex((_, position) =>
    request.familyKinds[position] === "named");
  if (index < 0) return true;
  return familyAvailable(document, request.familyStack[index], "named");
}

async function observedFonts(
  document: Document,
  inventory: FontInventory,
  pending: Set<string>,
  signal: AbortSignal,
  contractDigest: string,
): Promise<BrowserFontResult> {
  const probes = new Map<string, boolean>();
  if (document.fonts) {
    await forEachBounded(inventory.requests, async (request) => {
      const sample = String.fromCodePoint(...request.sampleCodePoints.slice(0, 4096)) || " ";
      const specification = fontSpecification(request, request.familyStack, request.familyKinds);
      try {
        // Loading first settles any webface the document declared; availability of the
        // requested family itself is then measured, not asked of check().
        await abortable(document.fonts.load(specification, sample), signal);
        probes.set(request.id, requestedFamilyAvailable(document, request));
      } catch (error) {
        if (signal.aborted) throw error;
        probes.set(request.id, false);
      } finally {
        pending.delete(`request:${request.id}`);
      }
    });
    await abortable(document.fonts.ready.then(() => undefined), signal);
  } else {
    inventory.requests.forEach((request) => pending.delete(`request:${request.id}`));
  }
  pending.delete("document.fonts.ready");
  const resolutions: FontResolution[] = [];
  for (const request of inventory.requests) {
    resolutions.push({
      ...baseResolution(request),
      sampleDigest: await digestJson(request.sampleCodePoints),
      status: probes.get(request.id) === false ? "missing" : "unverified",
      source: "browser",
      glyphCoverage: "unverified",
    });
  }
  const resolutionDigest = await digestJson({ contractDigest, requests: inventory.requests, resolutions });
  return {
    identity: {
      resolverContract: FONT_RESOLVER_CONTRACT_ID,
      substitutionContractVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
      substitutionContractDigest: contractDigest,
      resolutionDigest,
    },
    resolutions,
    renderedTextNodeCount: inventory.renderedTextNodeCount,
  };
}

async function configuredFonts(
  document: Document,
  inventory: FontInventory,
  resolver: FontResolver,
  limits: BrowserFontLimits,
  pending: Set<string>,
  signal: AbortSignal,
  contractDigest: string,
): Promise<BrowserFontResult> {
  pending.add("resolver");
  let responseValue: FontResolverResponse;
  const resolverRequests = Object.freeze(inventory.requests.map((request) => Object.freeze({
    id: request.id,
    familyStack: Object.freeze([...request.familyStack]),
    familyKinds: Object.freeze([...request.familyKinds]),
    style: request.style,
    weight: request.weight,
    stretch: request.stretch,
    sampleCodePoints: Object.freeze([...request.sampleCodePoints]),
  })));
  const resolverRequest = Object.freeze({
    schemaVersion: FONT_RESOLVER_SCHEMA_VERSION,
    requests: resolverRequests,
  });
  try {
    responseValue = await abortable(resolver(resolverRequest, signal), signal);
  } catch (cause) {
    if (signal.aborted) throw cause;
    throw new BrowserFontError(
      "invalid_response",
      `The configured font resolver failed: ${cause instanceof Error ? cause.message : String(cause)}`,
    );
  } finally {
    pending.delete("resolver");
  }
  const validated = await validateResolverResponse(responseValue, inventory.requests, limits, contractDigest);
  const resolverDigest = await digestJson({
    requests: inventory.requests,
    response: responseIdentityMaterial(validated.response),
  });
  const successfulFaces = new Set<string>();
  for (const id of validated.faces.keys()) pending.add(`face:${id}`);
  await forEachBounded(Array.from(validated.faces), async ([id, face]) => {
    try {
      const view = document.defaultView;
      if (!view) throw new Error("The render document has no font realm.");
      const candidate = new view.FontFace("__DocxodusValidation", new Uint8Array(face.bytes).buffer, {
        style: face.record.style,
        weight: String(face.record.weight),
        stretch: `${face.record.stretch}%`,
      });
      await abortable(candidate.load(), signal);
      successfulFaces.add(id);
    } catch (error) {
      if (signal.aborted) throw error;
    } finally {
      pending.delete(`face:${id}`);
    }
  });

  const outcomeById = new Map(validated.response.outcomes.map((outcome) => [outcome.requestId, outcome]));
  const successfulRequests = new Set(validated.response.outcomes
    .filter((outcome) => outcome.faceId && successfulFaces.has(outcome.faceId))
    .map((outcome) => outcome.requestId));
  const syntheticByRequest = new Map<string, string>();
  const syntheticByFaceId = new Map<string, string>();
  for (const outcome of validated.response.outcomes) {
    if (!outcome.faceId || !successfulFaces.has(outcome.faceId)) continue;
    let synthetic = syntheticByFaceId.get(outcome.faceId);
    if (!synthetic) {
      synthetic = await syntheticFamily(resolverDigest, outcome.faceId);
      syntheticByFaceId.set(outcome.faceId, synthetic);
    }
    syntheticByRequest.set(outcome.requestId, synthetic);
  }

  const installedMappings = new Set<string>();
  const style = document.createElement("style");
  style.id = "docxodus-configured-fonts";
  const rebuildStyle = (): void => {
    const rules: string[] = [];
    installedMappings.clear();
    for (const outcome of validated.response.outcomes) {
      const synthetic = syntheticByRequest.get(outcome.requestId);
      const face = outcome.faceId ? validated.faces.get(outcome.faceId)?.record : undefined;
      if (!synthetic || !face || !successfulRequests.has(outcome.requestId)) continue;
      const mapping = `${synthetic}\u0000${face.id}`;
      if (installedMappings.has(mapping)) continue;
      installedMappings.add(mapping);
      rules.push(faceRule(synthetic, face));
    }
    style.textContent = rules.sort(compareText).join("\n");
  };
  rebuildStyle();
  if (style.textContent) document.head.appendChild(style);
  const usesByRequest = new Map<string, TextFaceUse[]>();
  for (const use of inventory.uses) {
    const list = usesByRequest.get(use.requestKey) ?? [];
    list.push(use);
    usesByRequest.set(use.requestKey, list);
  }
  for (const request of inventory.requests) {
    const synthetic = syntheticByRequest.get(request.id);
    if (!synthetic) continue;
    for (const use of usesByRequest.get(request.id) ?? []) {
      use.element.style.setProperty(
        "font-family",
        `${cssString(synthetic)}, ${serializeFamilyStack(request.familyStack, request.familyKinds)}`,
        "important",
      );
    }
  }

  const fallbackAvailable = new Map<string, boolean>();
  if (document.fonts) {
    await forEachBounded(inventory.requests, async (request) => {
      const synthetic = syntheticByRequest.get(request.id);
      try {
        if (synthetic) {
          const sample = String.fromCodePoint(...request.sampleCodePoints.slice(0, 4096)) || " ";
          const loaded = await abortable(document.fonts.load(
            fontSpecification(request, [synthetic]),
            sample,
          ), signal);
          if (loaded.length === 0) {
            successfulRequests.delete(request.id);
            syntheticByRequest.delete(request.id);
            for (const use of usesByRequest.get(request.id) ?? []) restoreUse(use);
          }
        } else {
          const specification = fontSpecification(request, request.familyStack, request.familyKinds);
          const sample = String.fromCodePoint(...request.sampleCodePoints.slice(0, 4096)) || " ";
          await abortable(document.fonts.load(specification, sample), signal);
          fallbackAvailable.set(request.id, document.fonts.check(specification, sample));
        }
      } catch (error) {
        if (signal.aborted) throw error;
        fallbackAvailable.set(request.id, false);
        successfulRequests.delete(request.id);
        syntheticByRequest.delete(request.id);
        for (const use of usesByRequest.get(request.id) ?? []) restoreUse(use);
      } finally {
        pending.delete(`request:${request.id}`);
      }
    });
    rebuildStyle();
    if (!style.textContent) style.remove();
    await abortable(document.fonts.ready.then(() => undefined), signal);
  } else {
    inventory.requests.forEach((request) => pending.delete(`request:${request.id}`));
  }
  pending.delete("document.fonts.ready");

  const resolutions: FontResolution[] = [];
  for (const request of inventory.requests) {
    const outcome = outcomeById.get(request.id)!;
    const selected = outcome.faceId ? validated.faces.get(outcome.faceId)?.record : undefined;
    const loadFailed = selected !== undefined && !successfulRequests.has(request.id);
    const browserFallbackAvailable = !selected && outcome.status === "missing"
      ? fallbackAvailable.get(request.id) === true
      : undefined;
    const glyphCoverage = selected ? outcome.glyphCoverage : "unverified";
    resolutions.push({
      ...baseResolution(request),
      sampleDigest: await digestJson(request.sampleCodePoints),
      ...(outcome.resolvedFamily ? { resolvedFamily: outcome.resolvedFamily } : {}),
      ...(selected?.postscriptName ? { resolvedFace: selected.postscriptName } : selected ? { resolvedFace: selected.id } : {}),
      // A browser-generic fallback can make document.fonts.check() succeed, but
      // it cannot turn an explicit resolver miss into a verified resolution.
      status: loadFailed ? "load_failed" : outcome.status,
      source: selected
        ? selected.licenseEvidence.kind === "attested" ? "attested" : "configured"
        : "browser",
      ...(selected ? {
        format: selected.format,
        fileSha256: selected.sha256,
        version: selected.version,
        licenseEvidence: { ...selected.licenseEvidence },
      } : {}),
      ...(outcome.faceMatch ? { faceMatch: outcome.faceMatch } : {}),
      ...(outcome.metricCompatible === undefined ? {} : { metricCompatible: outcome.metricCompatible }),
      ...(glyphCoverage ? { glyphCoverage } : {}),
      ...(outcome.missingCodePoints ? { missingCodePointCount: outcome.missingCodePoints.length } : {}),
      ...(browserFallbackAvailable === undefined ? {} : { browserFallbackAvailable }),
    });
  }
  const resolutionDigest = await digestJson({ resolverDigest, resolutions });
  return {
    identity: {
      resolverContract: FONT_RESOLVER_CONTRACT_ID,
      substitutionContractVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
      substitutionContractDigest: contractDigest,
      resolutionDigest,
    },
    resolutions,
    renderedTextNodeCount: inventory.renderedTextNodeCount,
  };
}

export function createBrowserFontTask(
  document: Document,
  resolver: FontResolver | undefined,
  limits: BrowserFontLimits,
): BrowserFontTask {
  const inventory = collectFontInventory(document, limits);
  const pending = new Set(inventory.requests.map((request) => `request:${request.id}`));
  if (document.fonts) pending.add("document.fonts.ready");
  return {
    pending: () => Array.from(pending, (item) => `font:${item}`).sort(compareText),
    async wait(signal) {
      const contractDigest = await digestJson(FONT_SUBSTITUTION_CONTRACT_MATERIAL);
      return resolver
        ? configuredFonts(document, inventory, resolver, limits, pending, signal, contractDigest)
        : observedFonts(document, inventory, pending, signal, contractDigest);
    },
  };
}

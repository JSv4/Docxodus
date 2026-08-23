import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  existsSync,
  lstatSync,
  readFileSync,
  readdirSync,
} from 'node:fs';
import { basename, dirname, isAbsolute, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDFDocument, PDFName, PDFNumber } from 'pdf-lib';
import type { PDFObject, PDFPage } from 'pdf-lib';
import { getDocument, OPS, type PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { MAXIMUM_PNG_PIXELS } from './png.js';

/** Directory pdfjs reads its standard-14 font data from; resolved against the installed package. */
const PDFJS_STANDARD_FONTS = fileURLToPath(
  new URL('../../node_modules/pdfjs-dist/standard_fonts/', import.meta.url));

/**
 * Resource ceilings for the fixed release corpus. These are deliberately far above its current
 * largest artifacts while preventing a broken renderer/parser from turning CI into an unbounded
 * PDF, raster, text, annotation, or operator-list expansion.
 */
export const PDF_PARITY_LIMITS = Object.freeze({
  maximumPdfBytes: 32 * 1024 * 1024,
  maximumPages: 32,
  maximumRasterBytesPerPage: 16 * 1024 * 1024,
  maximumTotalRasterBytes: 128 * 1024 * 1024,
  // Single owner: the PNG decoder enforces the same ceiling, and two literals drift.
  maximumRasterPixelsPerPage: MAXIMUM_PNG_PIXELS,
  maximumTextCharacters: 5_000_000,
  maximumAnnotationsPerPage: 4_096,
  maximumOperatorEntriesPerPage: 1_000_000,
  maximumPhysicalPagePoints: 14_400,
  maximumHyperlinkTargetCharacters: 8_192,
  /**
   * Floor a chart case's vector path operations must clear.
   *
   * `> 0` was not a contract: measured on this corpus, an ordinary text page (HC023) already
   * emits 4 construct-path operations from text decoration alone, so a chart that failed to
   * render entirely would still have passed. The supported clustered chart (HC043) emits 27.
   * Twelve sits well above the incidental baseline and well below a rendered chart, so a chart
   * collapsing to a blank frame or a raster fallback fails instead of passing.
   */
  minimumChartVectorPathOperations: 12,
});

/**
 * The one raster contract used for both the Docxodus PDF and its reference PDF.
 *
 * RGB is selected by deliberately omitting pdftoppm's `-mono` and `-gray`
 * switches. MediaBox is selected by deliberately omitting `-cropbox`. Keeping
 * those choices in this frozen, hashable record makes defaults that otherwise
 * look accidental part of the reviewable comparison contract.
 */
export const PDF_RASTER_CONTRACT = Object.freeze({
  schemaVersion: 1 as const,
  tool: 'pdftoppm' as const,
  dpi: 96 as const,
  format: 'png' as const,
  colorModel: 'rgb' as const,
  pageBox: 'media' as const,
  fontAntialiasing: true as const,
  vectorAntialiasing: true as const,
  freeType: true as const,
  thinLineMode: 'none' as const,
  forcePageNumber: true as const,
  maximumPages: PDF_PARITY_LIMITS.maximumPages,
  maximumPdfBytes: PDF_PARITY_LIMITS.maximumPdfBytes,
  maximumRasterBytesPerPage: PDF_PARITY_LIMITS.maximumRasterBytesPerPage,
  maximumTotalRasterBytes: PDF_PARITY_LIMITS.maximumTotalRasterBytes,
  maximumRasterPixelsPerPage: PDF_PARITY_LIMITS.maximumRasterPixelsPerPage,
});

function sha256(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex');
}

export const PDF_RASTER_CONTRACT_SHA256 = sha256(JSON.stringify(PDF_RASTER_CONTRACT));

export interface RasterArtifact {
  pageNumber: number;
  path: string;
  bytes: number;
  sha256: string;
}

export interface RasterizedPdf {
  pdfPath: string;
  pdfSha256: string;
  contractSha256: string;
  pages: RasterArtifact[];
}

export interface RasterizePdfOptions {
  env?: NodeJS.ProcessEnv;
  timeoutMs?: number;
  /** Absolute executable already resolved and hashed by the benchmark preflight. */
  executablePath?: string;
}

function escapedRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function numberedPngs(outputPrefix: string): Array<{ pageNumber: number; path: string }> {
  const directory = dirname(outputPrefix);
  const prefix = basename(outputPrefix);
  const pattern = new RegExp(`^${escapedRegExp(prefix)}-(\\d+)\\.png$`);
  return readdirSync(directory)
    .map((name) => ({ name, match: name.match(pattern) }))
    .filter((entry): entry is { name: string; match: RegExpMatchArray } => entry.match !== null)
    .map((entry) => {
      const pageNumber = Number(entry.match[1]);
      const path = join(directory, entry.name);
      if (!Number.isSafeInteger(pageNumber) || pageNumber < 1) {
        throw new Error(`pdftoppm produced an invalid page number at ${path}`);
      }
      const metadata = lstatSync(path);
      if (!metadata.isFile() || metadata.isSymbolicLink()) {
        throw new Error(`pdftoppm raster output is not a regular non-symlink file: ${path}`);
      }
      return { pageNumber, path };
    })
    .sort((left, right) => left.pageNumber - right.pageNumber);
}

/** Builds the exact argument vector used for every PDF in the comparison. */
export function popplerRasterArguments(pdfPath: string, outputPrefix: string): readonly string[] {
  const input = resolve(pdfPath);
  const output = resolve(outputPrefix);
  return Object.freeze([
    '-r', String(PDF_RASTER_CONTRACT.dpi),
    '-png',
    '-freetype', PDF_RASTER_CONTRACT.freeType ? 'yes' : 'no',
    '-thinlinemode', PDF_RASTER_CONTRACT.thinLineMode,
    '-aa', PDF_RASTER_CONTRACT.fontAntialiasing ? 'yes' : 'no',
    '-aaVector', PDF_RASTER_CONTRACT.vectorAntialiasing ? 'yes' : 'no',
    '-f', '1',
    '-l', String(PDF_RASTER_CONTRACT.maximumPages + 1),
    '-forcenum',
    '-q',
    input,
    output,
  ]);
}

/**
 * Rasterizes one PDF without accepting pre-existing numbered outputs. The
 * caller supplies a distinct prefix for each renderer, but this function owns
 * the command contract and artifact hashing for both.
 */
export function rasterizePdf(
  pdfPath: string,
  outputPrefix: string,
  options: RasterizePdfOptions = {},
): RasterizedPdf {
  if (!isAbsolute(pdfPath) || !isAbsolute(outputPrefix)) {
    throw new Error('PDF raster inputs and output prefixes must be absolute paths.');
  }
  const resolvedPdf = resolve(pdfPath);
  const resolvedPrefix = resolve(outputPrefix);
  const pdfMetadata = existsSync(resolvedPdf) ? lstatSync(resolvedPdf) : undefined;
  if (!pdfMetadata?.isFile() || pdfMetadata.isSymbolicLink()) {
    throw new Error(`PDF raster input is not a regular file: ${resolvedPdf}`);
  }
  const pdfBytes = pdfMetadata.size;
  if (pdfBytes > PDF_RASTER_CONTRACT.maximumPdfBytes) {
    throw new Error(`PDF raster input exceeds ${PDF_RASTER_CONTRACT.maximumPdfBytes} bytes: ${pdfBytes}`);
  }
  const outputDirectory = dirname(resolvedPrefix);
  const outputMetadata = existsSync(outputDirectory) ? lstatSync(outputDirectory) : undefined;
  if (!outputMetadata?.isDirectory() || outputMetadata.isSymbolicLink()) {
    throw new Error(`PDF raster output directory does not exist: ${outputDirectory}`);
  }
  if (numberedPngs(resolvedPrefix).length > 0) {
    throw new Error(`Refusing to mix stale raster pages for prefix: ${resolvedPrefix}`);
  }

  const executable = options.executablePath ?? PDF_RASTER_CONTRACT.tool;
  if (options.executablePath !== undefined && !isAbsolute(options.executablePath)) {
    throw new Error('The pinned pdftoppm executable path must be absolute.');
  }
  execFileSync(executable, [...popplerRasterArguments(resolvedPdf, resolvedPrefix)], {
    env: options.env,
    stdio: 'pipe',
    timeout: options.timeoutMs ?? 120_000,
  });
  const pages = numberedPngs(resolvedPrefix);
  if (pages.length === 0) {
    throw new Error(`pdftoppm produced no PNG pages for: ${resolvedPdf}`);
  }
  if (pages.length > PDF_RASTER_CONTRACT.maximumPages) {
    throw new Error(`pdftoppm produced more than ${PDF_RASTER_CONTRACT.maximumPages} pages.`);
  }
  // One read per page: the size check, the dimension header, and the digest all come from the
  // same bytes, so re-stat-ing and re-reading every raster three times bought nothing.
  let totalRasterBytes = 0;
  const artifacts = pages.map((page, index): RasterArtifact => {
    if (page.pageNumber !== index + 1) {
      throw new Error(`pdftoppm produced a non-contiguous page sequence at ${page.path}`);
    }
    const bytes = readFileSync(page.path);
    if (bytes.byteLength > PDF_RASTER_CONTRACT.maximumRasterBytesPerPage) {
      throw new Error(`Raster page exceeds ${PDF_RASTER_CONTRACT.maximumRasterBytesPerPage} bytes: ${page.path}`);
    }
    totalRasterBytes += bytes.byteLength;
    if (totalRasterBytes > PDF_RASTER_CONTRACT.maximumTotalRasterBytes) {
      throw new Error(`Raster output exceeds ${PDF_RASTER_CONTRACT.maximumTotalRasterBytes} total bytes.`);
    }
    const header = bytes.subarray(0, 24);
    const isPng = header.length === 24
      && header.subarray(0, 8).equals(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    const width = isPng ? header.readUInt32BE(16) : 0;
    const height = isPng ? header.readUInt32BE(20) : 0;
    if (!width || !height || width * height > PDF_RASTER_CONTRACT.maximumRasterPixelsPerPage) {
      throw new Error(`Raster page has invalid or excessive dimensions at ${page.path}: ${width}x${height}`);
    }
    return { ...page, bytes: bytes.byteLength, sha256: sha256(bytes) };
  });
  return {
    pdfPath: resolvedPdf,
    pdfSha256: sha256(readFileSync(resolvedPdf)),
    contractSha256: PDF_RASTER_CONTRACT_SHA256,
    pages: artifacts,
  };
}

export interface PdfBox {
  x: number;
  y: number;
  width: number;
  height: number;
}

export type PdfHyperlinkTarget =
  | { kind: 'url'; value: string; unsafeValue?: string }
  | { kind: 'destination'; value: string; namedDestination?: string; pageNumber: number }
  | { kind: 'unsupported'; value?: string; reason: string };

export interface PdfHyperlinkAnnotation {
  target: PdfHyperlinkTarget;
  rectangle?: readonly number[];
}

function annotationRectangle(value: unknown): readonly number[] | undefined {
  if (value === undefined) return undefined;
  if (!Array.isArray(value) || value.length !== 4
    || !value.every((coordinate) => typeof coordinate === 'number'
      && Number.isFinite(coordinate) && Math.abs(coordinate) <= 1_000_000)) {
    throw new Error('PDF link annotation has an invalid or excessive rectangle.');
  }
  return Object.freeze([...value] as number[]);
}

export interface PdfPageInspection {
  pageNumber: number;
  userUnit: number;
  rotation: number;
  mediaBox: PdfBox;
  cropBox: PdfBox;
  text: string;
  textSha256: string;
  hyperlinks: PdfHyperlinkAnnotation[];
  constructPathOperations: number;
}

export interface PdfSemanticExpectations {
  requireSelectableText?: boolean;
  requiredText?: readonly string[];
  forbiddenText?: readonly string[];
  requiredHyperlinkTargets?: readonly PdfHyperlinkExpectation[];
}

export type PdfHyperlinkExpectation =
  | { kind: 'url'; value: string; representation?: 'normalized' | 'source' }
  | { kind: 'destination'; value?: string; pageNumber?: number };

export type PdfSemanticStatus = 'observed' | 'verified' | 'failed';

export interface PdfSemanticResult {
  passed: boolean;
  selectableText: {
    status: PdfSemanticStatus;
    characterCount: number;
    sha256: string;
    required: boolean;
    requiredFragments: readonly string[];
    missingFragments: readonly string[];
    forbiddenFragments: readonly string[];
    presentForbiddenFragments: readonly string[];
  };
  hyperlinks: {
    status: PdfSemanticStatus;
    annotationCount: number;
    supportedCount: number;
    unsupportedCount: number;
    targets: readonly PdfHyperlinkTarget[];
    requiredTargets: readonly PdfHyperlinkExpectation[];
    missingTargets: readonly PdfHyperlinkExpectation[];
  };
}

export interface PdfInspection {
  pdfSha256: string;
  pageCount: number;
  searchableText: string;
  linkAnnotations: number;
  vectorPathOperations: number;
  marked: boolean;
  pages: PdfPageInspection[];
  semantics: PdfSemanticResult;
}

function pdfBox(box: { x: number; y: number; width: number; height: number }): PdfBox {
  return { x: box.x, y: box.y, width: box.width, height: box.height };
}

function inheritedPageNumber(page: PDFPage, name: string): number | undefined {
  const value = page.node.getInheritableAttribute(PDFName.of(name));
  if (value === undefined) return undefined;
  const resolved = page.node.context.lookup(value) as PDFObject | undefined;
  if (!(resolved instanceof PDFNumber)) {
    throw new Error(`PDF page attribute /${name} is not a number.`);
  }
  return resolved.asNumber();
}

/**
 * Scale a page-space box and orient its dimensions. x/y deliberately remain the scaled raw PDF
 * box origin: /Rotate changes the displayed axes but not the page dictionary's coordinate-space
 * origin. Keeping rotation as a separate field avoids inventing a viewer-specific translation.
 */
function scaledOrientedPdfBox(
  box: { x: number; y: number; width: number; height: number },
  userUnit: number,
  rotation: number,
): PdfBox {
  const swapsAxes = rotation === 90 || rotation === 270;
  const physical = pdfBox({
    x: box.x * userUnit,
    y: box.y * userUnit,
    width: (swapsAxes ? box.height : box.width) * userUnit,
    height: (swapsAxes ? box.width : box.height) * userUnit,
  });
  if (!Object.values(physical).every(Number.isFinite)
    || physical.width <= 0 || physical.height <= 0
    || Math.abs(physical.x) > PDF_PARITY_LIMITS.maximumPhysicalPagePoints
    || Math.abs(physical.y) > PDF_PARITY_LIMITS.maximumPhysicalPagePoints
    || physical.width > PDF_PARITY_LIMITS.maximumPhysicalPagePoints
    || physical.height > PDF_PARITY_LIMITS.maximumPhysicalPagePoints) {
    throw new Error(`PDF page has invalid or excessive physical geometry: ${JSON.stringify(physical)}`);
  }
  return physical;
}

function destinationValue(value: unknown): string {
  if (typeof value === 'string') return value;
  try {
    const serialized = JSON.stringify(value) ?? String(value);
    return serialized.slice(0, PDF_PARITY_LIMITS.maximumHyperlinkTargetCharacters);
  } catch {
    return String(value).slice(0, PDF_PARITY_LIMITS.maximumHyperlinkTargetCharacters);
  }
}

async function resolvedDestination(
  pdf: PDFDocumentProxy,
  destination: unknown,
): Promise<PdfHyperlinkTarget> {
  const namedDestination = typeof destination === 'string' ? destination : undefined;
  const explicit = namedDestination
    ? await pdf.getDestination(namedDestination)
    : Array.isArray(destination) ? destination : null;
  const value = destinationValue(destination);
  if (!explicit || explicit.length === 0) {
    return { kind: 'unsupported', value, reason: 'destination did not resolve' };
  }
  const pageReference = explicit[0];
  if (!pageReference || typeof pageReference !== 'object'
    || !('num' in pageReference) || !('gen' in pageReference)) {
    return { kind: 'unsupported', value, reason: 'destination has no page reference' };
  }
  try {
    const pageNumber = await pdf.getPageIndex(pageReference as { num: number; gen: number }) + 1;
    return {
      kind: 'destination',
      value,
      ...(namedDestination ? { namedDestination } : {}),
      pageNumber,
    };
  } catch {
    return { kind: 'unsupported', value, reason: 'destination page is outside the document' };
  }
}

async function hyperlinkTarget(
  pdf: PDFDocumentProxy,
  annotation: Record<string, unknown>,
): Promise<PdfHyperlinkTarget> {
  if (typeof annotation.url === 'string') {
    if (annotation.url.length > PDF_PARITY_LIMITS.maximumHyperlinkTargetCharacters
      || (typeof annotation.unsafeUrl === 'string'
        && annotation.unsafeUrl.length > PDF_PARITY_LIMITS.maximumHyperlinkTargetCharacters)) {
      return { kind: 'unsupported', reason: 'link target exceeds the inspection limit' };
    }
    let protocol: string;
    try {
      protocol = new URL(annotation.url).protocol.toLowerCase();
    } catch {
      return { kind: 'unsupported', reason: 'link target is not an absolute URL' };
    }
    if (!['http:', 'https:', 'mailto:'].includes(protocol)) {
      return { kind: 'unsupported', reason: `link target uses disallowed protocol ${protocol}` };
    }
    return {
      kind: 'url',
      value: annotation.url,
      ...(typeof annotation.unsafeUrl === 'string' ? { unsafeValue: annotation.unsafeUrl } : {}),
    };
  }
  if (annotation.dest !== undefined && annotation.dest !== null) {
    return resolvedDestination(pdf, annotation.dest);
  }
  return { kind: 'unsupported', reason: 'link annotation has no URI or destination' };
}

function targetMatches(target: PdfHyperlinkTarget, expected: PdfHyperlinkExpectation): boolean {
  if (target.kind !== expected.kind) return false;
  if (expected.kind === 'url') {
    return target.kind === 'url'
      && (expected.representation === 'source' ? target.unsafeValue : target.value) === expected.value;
  }
  return target.kind === 'destination'
    && (expected.value === undefined || target.value === expected.value)
    && (expected.pageNumber === undefined || target.pageNumber === expected.pageNumber);
}

function semanticResult(
  searchableText: string,
  hyperlinks: readonly PdfHyperlinkAnnotation[],
  expectations: PdfSemanticExpectations,
): PdfSemanticResult {
  const requiredText = Object.freeze([...(expectations.requiredText ?? [])]);
  const forbiddenText = Object.freeze([...(expectations.forbiddenText ?? [])]);
  const requiredTargets = Object.freeze((expectations.requiredHyperlinkTargets ?? [])
    .map((target) => Object.freeze({ ...target })));
  const missingFragments = Object.freeze(requiredText.filter((fragment) => !searchableText.includes(fragment)));
  const presentForbiddenFragments = Object.freeze(
    forbiddenText.filter((fragment) => searchableText.includes(fragment)),
  );
  const targets = Object.freeze(hyperlinks.map((annotation) => annotation.target));
  const missingTargets = Object.freeze(requiredTargets.filter((expected) =>
    !targets.some((target) => targetMatches(target, expected))));
  const textRequired = expectations.requireSelectableText === true
    || requiredText.length > 0
    || forbiddenText.length > 0;
  const textFailed = (expectations.requireSelectableText === true && searchableText.length === 0)
    || missingFragments.length > 0
    || presentForbiddenFragments.length > 0;
  const hyperlinksRequired = requiredTargets.length > 0;
  const hyperlinksFailed = missingTargets.length > 0
    || (hyperlinksRequired && targets.some((target) => target.kind === 'unsupported'));
  return {
    passed: !textFailed && !hyperlinksFailed,
    selectableText: {
      status: textRequired ? textFailed ? 'failed' : 'verified' : 'observed',
      characterCount: searchableText.length,
      sha256: sha256(searchableText),
      required: textRequired,
      requiredFragments: requiredText,
      missingFragments,
      forbiddenFragments: forbiddenText,
      presentForbiddenFragments,
    },
    hyperlinks: {
      status: hyperlinksRequired ? hyperlinksFailed ? 'failed' : 'verified' : 'observed',
      annotationCount: hyperlinks.length,
      supportedCount: targets.filter((target) => target.kind !== 'unsupported').length,
      unsupportedCount: targets.filter((target) => target.kind === 'unsupported').length,
      targets,
      requiredTargets,
      missingTargets,
    },
  };
}

/**
 * Inspects actual PDF bytes. pdf-lib owns physical page-box parsing, while the
 * exact pdfjs-dist seam used by the Node export tests owns selectable text and
 * hyperlink annotation semantics.
 */
export async function inspectPdf(
  bytes: Uint8Array,
  expectations: PdfSemanticExpectations = {},
): Promise<PdfInspection> {
  if (bytes.byteLength < 16 || bytes.byteLength > PDF_PARITY_LIMITS.maximumPdfBytes) {
    throw new Error(`PDF inspection input must be 16..${PDF_PARITY_LIMITS.maximumPdfBytes} bytes.`);
  }
  const owned = new Uint8Array(bytes);
  const geometryDocument = await PDFDocument.load(new Uint8Array(owned), {
    ignoreEncryption: false,
    updateMetadata: false,
    throwOnInvalidObject: true,
  });
  const geometryPages = geometryDocument.getPages();
  if (geometryPages.length < 1 || geometryPages.length > PDF_PARITY_LIMITS.maximumPages) {
    throw new Error(`PDF inspection page count must be 1..${PDF_PARITY_LIMITS.maximumPages}.`);
  }
  const loading = getDocument({
    data: new Uint8Array(owned),
    useSystemFonts: false,
    disableFontFace: true,
    // Without this, pdfjs warns and degrades on any PDF that leaves a base-14 face unembedded —
    // silently weakening the text extraction this contract gates on. The data ships in the same
    // pdfjs-dist the Node export tests use, so it stays version-locked with the parser.
    standardFontDataUrl: PDFJS_STANDARD_FONTS,
    stopAtErrors: true,
    maxImageSize: PDF_PARITY_LIMITS.maximumRasterPixelsPerPage,
    canvasMaxAreaInBytes: PDF_PARITY_LIMITS.maximumRasterBytesPerPage,
  });
  try {
    const pdf = await loading.promise;
    if (pdf.numPages !== geometryPages.length) {
      throw new Error(
        `PDF parsers disagree on page count (${pdf.numPages} != ${geometryPages.length}).`,
      );
    }
    const pages: PdfPageInspection[] = [];
    const allHyperlinks: PdfHyperlinkAnnotation[] = [];
    let vectorPathOperations = 0;
    let totalTextCharacters = 0;
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const text = content.items.map((item) => 'str' in item ? item.str : '').join(' ');
      totalTextCharacters += text.length;
      if (totalTextCharacters > PDF_PARITY_LIMITS.maximumTextCharacters) {
        throw new Error(`PDF selectable text exceeds ${PDF_PARITY_LIMITS.maximumTextCharacters} characters.`);
      }
      const annotations = await page.getAnnotations();
      if (annotations.length > PDF_PARITY_LIMITS.maximumAnnotationsPerPage) {
        throw new Error(`PDF page ${pageNumber} exceeds the annotation limit.`);
      }
      const linkAnnotations = annotations
        .filter((annotation) => annotation.subtype === 'Link');
      const hyperlinks = await Promise.all(linkAnnotations
        .map(async (annotation): Promise<PdfHyperlinkAnnotation> => {
          const rectangle = annotationRectangle(annotation.rect);
          return {
            target: await hyperlinkTarget(pdf, annotation as Record<string, unknown>),
            ...(rectangle === undefined ? {} : { rectangle }),
          };
        }));
      const operations = await page.getOperatorList();
      if (operations.fnArray.length > PDF_PARITY_LIMITS.maximumOperatorEntriesPerPage) {
        throw new Error(`PDF page ${pageNumber} exceeds the operator-list limit.`);
      }
      const pagePaths = operations.fnArray.filter((operation) => operation === OPS.constructPath).length;
      const geometryPage = geometryPages[pageNumber - 1];
      const userUnit = inheritedPageNumber(geometryPage, 'UserUnit') ?? 1;
      if (!Number.isFinite(userUnit) || userUnit < 1 || userUnit > 75_000) {
        throw new Error(`PDF page ${pageNumber} has invalid /UserUnit ${String(userUnit)}.`);
      }
      const rawRotation = inheritedPageNumber(geometryPage, 'Rotate') ?? 0;
      if (!Number.isFinite(rawRotation) || !Number.isInteger(rawRotation)
        || rawRotation % 90 !== 0) {
        throw new Error(`PDF page ${pageNumber} has invalid /Rotate ${String(rawRotation)}.`);
      }
      const rotation = ((rawRotation % 360) + 360) % 360;
      vectorPathOperations += pagePaths;
      allHyperlinks.push(...hyperlinks);
      pages.push({
        pageNumber,
        userUnit,
        rotation,
        mediaBox: scaledOrientedPdfBox(geometryPage.getMediaBox(), userUnit, rotation),
        cropBox: scaledOrientedPdfBox(geometryPage.getCropBox(), userUnit, rotation),
        text,
        textSha256: sha256(text),
        hyperlinks,
        constructPathOperations: pagePaths,
      });
    }
    const searchableText = pages.map((page) => page.text).join('\n').trim();
    const markInfo = await pdf.getMarkInfo();
    return {
      pdfSha256: sha256(owned),
      pageCount: pages.length,
      searchableText,
      linkAnnotations: allHyperlinks.length,
      vectorPathOperations,
      marked: markInfo?.Marked === true,
      pages,
      semantics: semanticResult(searchableText, allHyperlinks, expectations),
    };
  } finally {
    await loading.destroy();
  }
}

export function sha256File(path: string): string {
  return sha256(readFileSync(path));
}

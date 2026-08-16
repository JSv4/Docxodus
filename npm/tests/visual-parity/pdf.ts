import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  existsSync,
  readFileSync,
  readdirSync,
  statSync,
} from 'node:fs';
import { basename, dirname, isAbsolute, join, resolve } from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { getDocument, OPS, type PDFDocumentProxy } from 'pdfjs-dist/legacy/build/pdf.mjs';

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
});

function sha256(bytes: Uint8Array | string): string {
  return createHash('sha256').update(bytes).digest('hex');
}

export const PDF_RASTER_CONTRACT_SHA256 = sha256(JSON.stringify(PDF_RASTER_CONTRACT));

export interface RasterArtifact {
  pageNumber: number;
  path: string;
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
    .map((entry) => ({ pageNumber: Number(entry.match[1]), path: join(directory, entry.name) }))
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
  if (!existsSync(resolvedPdf) || !statSync(resolvedPdf).isFile()) {
    throw new Error(`PDF raster input is not a regular file: ${resolvedPdf}`);
  }
  if (!existsSync(dirname(resolvedPrefix)) || !statSync(dirname(resolvedPrefix)).isDirectory()) {
    throw new Error(`PDF raster output directory does not exist: ${dirname(resolvedPrefix)}`);
  }
  if (numberedPngs(resolvedPrefix).length > 0) {
    throw new Error(`Refusing to mix stale raster pages for prefix: ${resolvedPrefix}`);
  }

  execFileSync(PDF_RASTER_CONTRACT.tool, [...popplerRasterArguments(resolvedPdf, resolvedPrefix)], {
    env: options.env,
    stdio: 'pipe',
    timeout: options.timeoutMs ?? 120_000,
  });
  const pages = numberedPngs(resolvedPrefix);
  if (pages.length === 0) {
    throw new Error(`pdftoppm produced no PNG pages for: ${resolvedPdf}`);
  }
  pages.forEach((page, index) => {
    if (page.pageNumber !== index + 1) {
      throw new Error(`pdftoppm produced a non-contiguous page sequence at ${page.path}`);
    }
  });
  return {
    pdfPath: resolvedPdf,
    pdfSha256: sha256(readFileSync(resolvedPdf)),
    contractSha256: PDF_RASTER_CONTRACT_SHA256,
    pages: pages.map((page) => ({
      ...page,
      sha256: sha256(readFileSync(page.path)),
    })),
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

export interface PdfPageInspection {
  pageNumber: number;
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

function destinationValue(value: unknown): string {
  if (typeof value === 'string') return value;
  try {
    return JSON.stringify(value) ?? String(value);
  } catch {
    return String(value);
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
  const hyperlinksFailed = missingTargets.length > 0;
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
  const owned = new Uint8Array(bytes);
  const geometryDocument = await PDFDocument.load(new Uint8Array(owned), {
    ignoreEncryption: false,
    updateMetadata: false,
    throwOnInvalidObject: true,
  });
  const geometryPages = geometryDocument.getPages();
  const loading = getDocument({
    data: new Uint8Array(owned),
    useSystemFonts: true,
  });
  const pdf = await loading.promise;
  try {
    if (pdf.numPages !== geometryPages.length) {
      throw new Error(
        `PDF parsers disagree on page count (${pdf.numPages} != ${geometryPages.length}).`,
      );
    }
    const pages: PdfPageInspection[] = [];
    const allHyperlinks: PdfHyperlinkAnnotation[] = [];
    let vectorPathOperations = 0;
    for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber++) {
      const page = await pdf.getPage(pageNumber);
      const content = await page.getTextContent();
      const text = content.items.map((item) => 'str' in item ? item.str : '').join(' ');
      const annotations = await page.getAnnotations();
      const linkAnnotations = annotations
        .filter((annotation) => annotation.subtype === 'Link')
      const hyperlinks = await Promise.all(linkAnnotations
        .map(async (annotation): Promise<PdfHyperlinkAnnotation> => ({
          target: await hyperlinkTarget(pdf, annotation as Record<string, unknown>),
          ...(Array.isArray(annotation.rect)
            ? { rectangle: Object.freeze(annotation.rect.map(Number)) }
            : {}),
        })));
      const operations = await page.getOperatorList();
      const pagePaths = operations.fnArray.filter((operation) => operation === OPS.constructPath).length;
      vectorPathOperations += pagePaths;
      allHyperlinks.push(...hyperlinks);
      pages.push({
        pageNumber,
        mediaBox: pdfBox(geometryPages[pageNumber - 1].getMediaBox()),
        cropBox: pdfBox(geometryPages[pageNumber - 1].getCropBox()),
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

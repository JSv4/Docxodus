import { createHash } from 'node:crypto';
import { canonicalJson } from './canonical-json.js';
import { PDF_PARITY_LIMITS } from './pdf.js';

const DIGEST = /^[0-9a-f]{64}$/;
const MAXIMUM_FRAGMENTS = 100_000;
const MAXIMUM_FRAGMENT_STRING_CHARACTERS = 8_192;
const MAXIMUM_TOTAL_FRAGMENT_STRING_CHARACTERS = 16 * 1024 * 1024;
const PAGE_TOLERANCE = 0.1;

function record(value: unknown): Record<string, unknown> | undefined {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
    ? value as Record<string, unknown>
    : undefined;
}

function exactKeys(value: Record<string, unknown>, allowed: readonly string[]): boolean {
  return Object.keys(value).every((key) => allowed.includes(key));
}

function sha256(value: Uint8Array | string): string {
  return createHash('sha256').update(value).digest('hex');
}

export interface SupportedPdfResultExpectation {
  sourceSha256: string;
  reviewProfile: string;
  commentProfile: string;
  documentVersion?: number;
}

export interface VerifiedPdfResult {
  pageCount: number;
  rendererFingerprint: string;
  pages: Array<{ pageNumber: number; width: number; height: number }>;
}

/** Independently validate the supported API envelope before any PDF parser overwrites its claims. */
export function assertSupportedPdfResult(
  candidate: unknown,
  expected: SupportedPdfResultExpectation,
): VerifiedPdfResult {
  const result = record(candidate);
  if (!result || !(result.pdf instanceof Uint8Array)
    || result.pdf.byteLength < 16 || result.pdf.byteLength > PDF_PARITY_LIMITS.maximumPdfBytes
    || !Number.isSafeInteger(result.pageCount) || (result.pageCount as number) < 1
    || (result.pageCount as number) > PDF_PARITY_LIMITS.maximumPages
    || typeof result.rendererFingerprint !== 'string' || !DIGEST.test(result.rendererFingerprint)
    || !Array.isArray(result.warnings)) {
    throw new Error('Supported PDF result has an invalid top-level envelope.');
  }

  const map = record(result.pageMap);
  if (!map || !exactKeys(map, [
    'schemaVersion', 'mode', 'availability', 'documentVersion', 'rendererFingerprint',
    'pages', 'fragments',
  ]) || map.schemaVersion !== 1 || map.mode !== 'paginated' || map.availability !== 'available'
    || map.documentVersion !== (expected.documentVersion ?? 0)
    || map.rendererFingerprint !== result.rendererFingerprint
    || !Array.isArray(map.pages) || map.pages.length !== result.pageCount
    || !Array.isArray(map.fragments) || map.fragments.length > MAXIMUM_FRAGMENTS) {
    throw new Error('Supported PDF result has an invalid or inconsistent PageMap envelope.');
  }

  const pages: Array<{ pageNumber: number; width: number; height: number }> = [];
  const pageSizes = new Map<number, { width: number; height: number }>();
  for (let index = 0; index < map.pages.length; index++) {
    const page = record(map.pages[index]);
    if (!page || !exactKeys(page, [
      'pageNumber', 'pageInSection', 'width', 'height', 'sectionIndex', 'pageName',
    ]) || page.pageNumber !== index + 1
      || !Number.isSafeInteger(page.pageInSection) || (page.pageInSection as number) < 1
      || typeof page.width !== 'number' || !Number.isFinite(page.width) || page.width <= 0
      || page.width > PDF_PARITY_LIMITS.maximumPhysicalPagePoints
      || typeof page.height !== 'number' || !Number.isFinite(page.height) || page.height <= 0
      || page.height > PDF_PARITY_LIMITS.maximumPhysicalPagePoints
      || (page.sectionIndex !== undefined
        && (!Number.isSafeInteger(page.sectionIndex) || (page.sectionIndex as number) < 0))
      || typeof page.pageName !== 'string' || page.pageName.length < 1 || page.pageName.length > 512) {
      throw new Error(`Supported PDF result has an invalid PageMap page at index ${index}.`);
    }
    const verified = {
      pageNumber: page.pageNumber as number,
      width: page.width,
      height: page.height,
    };
    pages.push(verified);
    pageSizes.set(verified.pageNumber, verified);
  }

  const fragmentIds = new Set<string>();
  const nextFragmentIndex = new Map<string, number>();
  let fragmentStringCharacters = 0;
  for (const candidateFragment of map.fragments) {
    const fragment = record(candidateFragment);
    const geometry = fragment ? record(fragment.geometry) : undefined;
    const page = fragment && Number.isSafeInteger(fragment.pageNumber)
      ? pageSizes.get(fragment.pageNumber as number)
      : undefined;
    fragmentStringCharacters += typeof fragment?.fragmentId === 'string' ? fragment.fragmentId.length : 0;
    fragmentStringCharacters += typeof fragment?.anchorId === 'string' ? fragment.anchorId.length : 0;
    if (!fragment || !exactKeys(fragment, [
      'fragmentId', 'anchorId', 'fragmentIndex', 'pageNumber', 'geometry', 'story', 'inTableCell',
    ]) || typeof fragment.fragmentId !== 'string' || fragment.fragmentId.length < 1
      || fragment.fragmentId.length > MAXIMUM_FRAGMENT_STRING_CHARACTERS
      || fragmentIds.has(fragment.fragmentId)
      || typeof fragment.anchorId !== 'string' || fragment.anchorId.length < 1
      || fragment.anchorId.length > MAXIMUM_FRAGMENT_STRING_CHARACTERS
      || fragmentStringCharacters > MAXIMUM_TOTAL_FRAGMENT_STRING_CHARACTERS
      || !Number.isSafeInteger(fragment.fragmentIndex) || (fragment.fragmentIndex as number) < 0
      || !page
      || fragment.fragmentId !== `p${fragment.pageNumber}-f${fragment.fragmentIndex}-${fragment.anchorId}`
      || !geometry || !exactKeys(geometry, ['x', 'y', 'width', 'height'])
      || ![geometry.x, geometry.y, geometry.width, geometry.height].every((value) =>
        typeof value === 'number' && Number.isFinite(value) && value >= 0)
      || (geometry.x as number) + (geometry.width as number) > page.width + PAGE_TOLERANCE
      || (geometry.y as number) + (geometry.height as number) > page.height + PAGE_TOLERANCE
      || !['body', 'header', 'footer', 'footnote', 'endnote', 'comment']
        .includes(String(fragment.story))
      || typeof fragment.inTableCell !== 'boolean') {
      throw new Error('Supported PDF result has an invalid PageMap fragment.');
    }
    const expectedIndex = nextFragmentIndex.get(fragment.anchorId) ?? 0;
    if (fragment.fragmentIndex !== expectedIndex) {
      throw new Error('Supported PDF result has a discontinuous PageMap fragment sequence.');
    }
    nextFragmentIndex.set(fragment.anchorId, expectedIndex + 1);
    fragmentIds.add(fragment.fragmentId);
  }

  const report = record(result.renderReport);
  const reportEnvironment = report ? record(report.environment) : undefined;
  const reportBindings = report ? record(report.bindings) : undefined;
  const reportSource = report ? record(report.source) : undefined;
  const reportOptions = report ? record(report.options) : undefined;
  const reportFonts: Array<Record<string, unknown>> = Array.isArray(report?.fonts)
    ? (report.fonts as unknown[]).flatMap((font) => {
      const entry = record(font);
      return entry === undefined ? [] : [entry];
    })
    : [];
  // Named, individually reported checks. A single boolean chain collapsed eleven distinct
  // failures into one sentence, so a benchmark that rejected every case said only that something
  // was unbound — with no way to tell a renderer-fingerprint mismatch from a profile mismatch
  // without re-running against a patched harness.
  const bindingChecks: ReadonlyArray<readonly [string, () => boolean]> = [
    ['renderReport is an object', () => report !== undefined],
    ['renderReport.status === "complete"', () => report?.status === 'complete'],
    ['renderReport.environment is an object', () => reportEnvironment !== undefined],
    ['environment.rendererFingerprint matches the result',
      () => reportEnvironment?.rendererFingerprint === result.rendererFingerprint],
    // NOT `verification === "nodeVerified"`. That label requires every face to be `resolved`,
    // and this benchmark's whole premise is metric-compatible SUBSTITUTION — fonts.conf maps
    // Calibri to Carlito — so any document asking for Calibri reports `substituted` and the label
    // is unreachable by construction. Gate on the properties that actually carry the guarantee:
    // the faces came from the configured directories rather than whatever the browser had, and
    // substitution did not degrade layout. An injected browser is still rejected, by the
    // fidelityTier check above, which `releaseBaselined` grants only to a `pinned` launch.
    ['every face came from the configured font directories',
      () => reportFonts.length > 0 && reportFonts.every((font) => font.source === 'configured')],
    ['every face resolved or substituted, none missing',
      () => reportFonts.every((font) => font.status === 'resolved' || font.status === 'substituted')],
    ['every substituted face is metric-compatible with complete glyph coverage',
      () => reportFonts.every((font) => font.status !== 'substituted'
        || (font.metricCompatible === true && font.glyphCoverage === 'complete'))],
    ['environment.fidelityTier === "releaseBaselined"',
      () => reportEnvironment?.fidelityTier === 'releaseBaselined'],
    ['renderReport.bindings is an object', () => reportBindings !== undefined],
    ['bindings.pdfDigest binds the returned PDF',
      () => reportBindings?.pdfDigest === sha256(result.pdf as Uint8Array)],
    ['bindings.pageMapDigest binds the returned PageMap',
      () => reportBindings?.pageMapDigest === sha256(canonicalJson(map))],
    ['renderReport.source is an object', () => reportSource !== undefined],
    ['source.rawPackageBytesDigest binds the pinned fixture',
      () => reportSource?.rawPackageBytesDigest === expected.sourceSha256],
    ['renderReport.options is an object', () => reportOptions !== undefined],
    ['options.reviewProfile matches the corpus contract',
      () => reportOptions?.reviewProfile === expected.reviewProfile],
    ['options.commentProfile matches the corpus contract',
      () => reportOptions?.commentProfile === expected.commentProfile],
    ['renderReport.pages agrees with pageMap.pages',
      () => canonicalJson(report?.pages) === canonicalJson(map.pages)],
    ['renderReport.warnings agrees with result.warnings',
      () => canonicalJson(report?.warnings) === canonicalJson(result.warnings)],
  ];
  const unmet = bindingChecks.filter(([, holds]) => !holds()).map(([label]) => label);
  if (unmet.length > 0) {
    const observed = {
      verification: reportEnvironment?.verification,
      fonts: reportFonts.map((font) => ({
        requested: font.requestedFamily,
        resolved: font.resolvedFamily,
        status: font.status,
        source: font.source,
        metricCompatible: font.metricCompatible,
        glyphCoverage: font.glyphCoverage,
      })),
      fidelityTier: reportEnvironment?.fidelityTier,
      reviewProfile: reportOptions?.reviewProfile,
      commentProfile: reportOptions?.commentProfile,
      status: report?.status,
    };
    throw new Error('Supported PDF result is not bound to its report, PageMap, source, and '
      + `profiles. Unmet: ${unmet.join('; ')}. Observed: ${JSON.stringify(observed)}`);
  }
  return { pageCount: result.pageCount as number, rendererFingerprint: result.rendererFingerprint, pages };
}

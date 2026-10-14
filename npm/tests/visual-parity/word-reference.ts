import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { VISUAL_THRESHOLDS, background } from './metrics.js';
import type { RgbaImage } from './png.js';
import { RATCHET_PRECISION } from './ratchet.js';
import type { VisualCorpusEntry } from './corpus.js';

/**
 * The Word-reference evidence store (issue #402).
 *
 * LibreOffice is a comparison implementation, not the correctness oracle — but without recorded
 * Word evidence, every LibreOffice-vs-Docxodus disagreement ends in an undecidable "unless Word
 * says otherwise". This module owns the committed, numbers-only record of what Microsoft Word
 * actually renders for each corpus fixture, so `reference-deviation` dispositions can cite Word
 * data instead of inference.
 *
 * The boundary mirrors the ratchet's: `word-reference.json` holds measurements (page counts,
 * page geometry, ink extents, named per-case coordinates) and the Word/OS versions they were
 * taken under — never proprietary binaries, never a committed image corpus. The only manual step
 * a Word license holder performs is exporting each fixture to PDF; everything downstream
 * (rasterization at the benchmark's 96-DPI Poppler contract, measurement, record maintenance)
 * is automated by the capture spec so the procedure is reproducible rather than artisanal.
 *
 * One honesty note the numbers must carry: Word renders with the GENUINE Office fonts, while
 * the benchmark's two engines share license-safe substitutes under the issue-#379 contract.
 * Word evidence therefore decides STRUCTURAL questions — page counts, whether declared spacing
 * is suppressed, where a block starts relative to the margin — and glyph-level pixel scores
 * against Word are advisory, never gating.
 */

export const WORD_REFERENCE_SCHEMA_VERSION = 1;

export const WORD_REFERENCE_FILE =
  resolve(dirname(fileURLToPath(import.meta.url)), 'word-reference.json');

/** Environment variable naming the directory of Word-exported `<case-id>.pdf` files. */
export const WORD_PDFS_ENV = 'DOCXODUS_WORD_REFERENCE_PDFS';
/** Optional: a prior benchmark run's output root, enabling the three-way comparison. */
export const WORD_RUN_ENV = 'DOCXODUS_WORD_REFERENCE_RUN';
/** Operator-declared Word version (e.g. "Word for Microsoft 365 MSO 16.0.18129"), required. */
export const WORD_VERSION_ENV = 'DOCXODUS_WORD_VERSION';
/** Operator-declared OS the export was made on (e.g. "Windows 11 24H2"), required. */
export const WORD_OS_ENV = 'DOCXODUS_WORD_OS';

export interface WordInkBounds {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface WordReferencePage {
  page: number;
  /** Pixels at the benchmark's 96-DPI raster contract. */
  width: number;
  height: number;
  /** Extent of non-background ink under the shared ink model; null for a blank page. */
  inkBounds: WordInkBounds | null;
  /** Ink pixels / page pixels, rounded to the ratchet precision. */
  inkPixelRatio: number;
}

/** Advisory pixel scores against Word (real Office fonts vs contract substitutes — see above). */
export interface WordComparisonSide {
  meanSsim: number;
  worstInkF1: number;
}

export interface WordReferenceComparison {
  /** Commit of the benchmark run whose engine renders were compared against the Word pages. */
  runCommit: string;
  docxodusVsWord: WordComparisonSide;
  libreofficeVsWord: WordComparisonSide;
}

export interface WordReferenceCase {
  id: string;
  /** `pending` rows are the coverage ledger: a corpus case that has never been measured. */
  status: 'pending' | 'measured';
  /** SHA-256 of the tracked fixture bytes the operator opened in Word. */
  fixtureSha256?: string;
  pageCount?: number;
  pages?: WordReferencePage[];
  /**
   * Named, case-specific coordinates the automatic extraction cannot know the meaning of —
   * e.g. `"outerTableTopPx": 96` or `"spacingBeforeNestedTablePx": 0`. Filled by the operator
   * per the procedure doc; these are what disposition rationales cite.
   */
  keyMeasurements?: Record<string, number | string>;
  /** Free-form observations (view settings used, anything anomalous). */
  notes?: string;
  capturedAt?: string;
}

export interface WordReferenceEnvironment {
  word: string;
  os: string;
}

export interface WordReferenceRecord {
  schemaVersion: number;
  description: string;
  procedure: string;
  /** Null until the first capture; a capture must declare what it measured with. */
  environment: WordReferenceEnvironment | null;
  cases: WordReferenceCase[];
  comparisons?: Record<string, WordReferenceComparison>;
}

export function roundTo(value: number): number {
  const factor = 10 ** RATCHET_PRECISION;
  return Math.round(value * factor) / factor;
}

/**
 * Ink extent of one rasterized Word page, using the SAME background detection and ink threshold
 * as the pairwise metrics (`metrics.ts`), so "ink" means one thing across the whole benchmark.
 */
export function measureWordPage(image: RgbaImage, pageNumber: number): WordReferencePage {
  const bg = background(image);
  let left = image.width;
  let top = image.height;
  let right = -1;
  let bottom = -1;
  let ink = 0;
  for (let y = 0; y < image.height; y++) {
    for (let x = 0; x < image.width; x++) {
      const i = (y * image.width + x) * 4;
      const distance = Math.max(
        Math.abs(image.data[i] - bg[0]),
        Math.abs(image.data[i + 1] - bg[1]),
        Math.abs(image.data[i + 2] - bg[2]),
      );
      if (distance > VISUAL_THRESHOLDS.inkBackgroundDistance) {
        ink++;
        if (x < left) left = x;
        if (x > right) right = x;
        if (y < top) top = y;
        if (y > bottom) bottom = y;
      }
    }
  }
  return {
    page: pageNumber,
    width: image.width,
    height: image.height,
    inkBounds: ink ? { left, top, right, bottom } : null,
    inkPixelRatio: roundTo(ink / Math.max(1, image.width * image.height)),
  };
}

export function emptyRecord(corpus: readonly VisualCorpusEntry[]): WordReferenceRecord {
  return {
    schemaVersion: WORD_REFERENCE_SCHEMA_VERSION,
    description:
      'Word-reference evidence store (issue #402). Numbers only: page counts, page geometry, ' +
      'ink extents, and named per-case measurements taken from Microsoft Word PDF exports of ' +
      'the tracked corpus fixtures, rasterized under the same 96-DPI Poppler contract as the ' +
      'benchmark. No binaries and no image corpus are ever committed. Word renders with the ' +
      'genuine Office fonts (not the issue-#379 substitutes), so these measurements decide ' +
      'STRUCTURAL questions; pixel scores against Word are advisory. See WORD_REFERENCE.md ' +
      'for the capture procedure.',
    procedure: 'npm/tests/visual-parity/WORD_REFERENCE.md',
    environment: null,
    cases: [...corpus]
      .map(entry => ({ id: entry.id, status: 'pending' as const }))
      .sort((a, b) => a.id.localeCompare(b.id)),
  };
}

/** The committed record, or null before one exists. */
export function readWordReference(file: string = WORD_REFERENCE_FILE): WordReferenceRecord | null {
  if (!existsSync(file)) return null;
  return JSON.parse(readFileSync(file, 'utf8')) as WordReferenceRecord;
}

/** Stable serialization — the record is reviewed as a diff, like the ratchet's. */
export function serializeWordReference(record: WordReferenceRecord): string {
  return `${JSON.stringify(record, null, 2)}\n`;
}

/** Replaces (or appends) one case's measurement, keeping id order stable. */
export function upsertCase(
  record: WordReferenceRecord,
  measured: WordReferenceCase,
): WordReferenceRecord {
  const cases = record.cases.filter(entry => entry.id !== measured.id);
  cases.push(measured);
  cases.sort((a, b) => a.id.localeCompare(b.id));
  return { ...record, cases };
}

/**
 * Structural validation of the committed record against the corpus. Returns problems rather
 * than throwing so the spec can report them all at once.
 */
export function validateWordReference(
  record: WordReferenceRecord,
  corpus: readonly VisualCorpusEntry[],
): string[] {
  const problems: string[] = [];
  if (record.schemaVersion !== WORD_REFERENCE_SCHEMA_VERSION) {
    problems.push(`schemaVersion ${record.schemaVersion} != ${WORD_REFERENCE_SCHEMA_VERSION}`);
  }

  const recordIds = record.cases.map(entry => entry.id);
  const sorted = [...recordIds].sort((a, b) => a.localeCompare(b));
  if (recordIds.join(',') !== sorted.join(',')) {
    problems.push('cases are not sorted by id — the record must diff stably');
  }
  const corpusIds = new Set(corpus.map(entry => entry.id));
  for (const id of corpusIds) {
    if (!recordIds.includes(id)) {
      problems.push(`corpus case ${id} has no word-reference row (add it as pending)`);
    }
  }
  for (const id of recordIds) {
    if (!corpusIds.has(id)) problems.push(`word-reference row ${id} matches no corpus case`);
  }
  if (new Set(recordIds).size !== recordIds.length) problems.push('duplicate case ids');

  for (const entry of record.cases) {
    if (entry.status === 'pending') {
      const extras = ['fixtureSha256', 'pageCount', 'pages', 'keyMeasurements', 'capturedAt']
        .filter(key => (entry as unknown as Record<string, unknown>)[key] !== undefined);
      if (extras.length) {
        problems.push(`${entry.id} is pending but carries measurement fields: ${extras.join(', ')}`);
      }
      continue;
    }
    if (!entry.pages?.length || entry.pageCount !== entry.pages.length) {
      problems.push(`${entry.id} is measured but pageCount/pages disagree`);
      continue;
    }
    if (!entry.fixtureSha256 || !/^[0-9a-f]{64}$/.test(entry.fixtureSha256)) {
      problems.push(`${entry.id} is measured without a valid fixtureSha256`);
    }
    entry.pages.forEach((page, index) => {
      if (page.page !== index + 1) problems.push(`${entry.id} page numbering is not 1..N`);
      if (page.width <= 0 || page.height <= 0) problems.push(`${entry.id} page ${page.page} has non-positive dimensions`);
      if (page.inkBounds) {
        const { left, top, right, bottom } = page.inkBounds;
        if (left > right || top > bottom || left < 0 || top < 0 ||
            right >= page.width || bottom >= page.height) {
          problems.push(`${entry.id} page ${page.page} ink bounds are inconsistent`);
        }
      }
    });
  }

  if (record.cases.some(entry => entry.status === 'measured') && !record.environment) {
    problems.push('measured cases exist but no environment (Word/OS versions) is recorded');
  }

  for (const entry of corpus) {
    if (entry.disposition.wordEvidence === undefined) continue;
    const row = record.cases.find(candidate => candidate.id === entry.id);
    if (!row || row.status !== 'measured') {
      problems.push(`${entry.id} disposition cites wordEvidence but the word-reference row is ` +
        `${row ? row.status : 'missing'} — a citation needs a measurement`);
    }
  }

  for (const id of Object.keys(record.comparisons ?? {})) {
    const row = record.cases.find(candidate => candidate.id === id);
    if (!row || row.status !== 'measured') {
      problems.push(`comparison recorded for ${id}, which is not a measured case`);
    }
  }

  return problems;
}

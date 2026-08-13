import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { VISUAL_PARITY_CORPUS, type VisualCorpusEntry } from './visual-parity/corpus.js';
import type { RgbaImage } from './visual-parity/png.js';
import {
  WORD_REFERENCE_FILE,
  WORD_REFERENCE_SCHEMA_VERSION,
  emptyRecord,
  measureWordPage,
  readWordReference,
  serializeWordReference,
  upsertCase,
  validateWordReference,
  type WordReferenceCase,
  type WordReferenceRecord,
} from './visual-parity/word-reference.js';

/**
 * The Word-reference store's own regression suite (issue #402). Like the ratchet spec, it is
 * deliberately NOT gated behind `DOCXODUS_VISUAL_PARITY`: the validation and measurement layers
 * are pure, so every pull request proves — without Word, LibreOffice, or Poppler — that the
 * committed evidence store stays consistent with the corpus and that a disposition cannot cite
 * Word evidence that was never measured.
 */

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');

function image(width: number, height: number, draw?: (data: Uint8Array) => void): RgbaImage {
  const data = new Uint8Array(width * height * 4).fill(255);
  draw?.(data);
  return { width, height, data };
}

function measuredCase(overrides: Partial<WordReferenceCase> = {}): WordReferenceCase {
  return {
    id: 'shape',
    status: 'measured',
    fixtureSha256: 'a'.repeat(64),
    pageCount: 1,
    pages: [{
      page: 1,
      width: 816,
      height: 1056,
      inkBounds: { left: 96, top: 96, right: 719, bottom: 959 },
      inkPixelRatio: 0.05,
    }],
    capturedAt: '2026-08-13',
    ...overrides,
  };
}

function recordWith(cases: WordReferenceCase[]): WordReferenceRecord {
  let record: WordReferenceRecord = {
    ...emptyRecord(VISUAL_PARITY_CORPUS),
    environment: { word: 'Word 365', os: 'Windows 11' },
  };
  for (const entry of cases) record = upsertCase(record, entry);
  return record;
}

test.describe('word-reference measurement', () => {
  test('a blank page has null ink bounds and zero ink ratio', () => {
    const page = measureWordPage(image(100, 50), 1);
    expect(page).toEqual({ page: 1, width: 100, height: 50, inkBounds: null, inkPixelRatio: 0 });
  });

  test('ink bounds are the exact extent of non-background pixels', () => {
    const page = measureWordPage(image(100, 50, data => {
      for (const [x, y] of [[10, 5], [40, 30]]) {
        const i = (y * 100 + x) * 4;
        data[i] = data[i + 1] = data[i + 2] = 0;
      }
    }), 3);
    expect(page.page).toBe(3);
    expect(page.inkBounds).toEqual({ left: 10, top: 5, right: 40, bottom: 30 });
    expect(page.inkPixelRatio).toBeCloseTo(2 / (100 * 50), 6);
  });

  test('near-background antialiasing below the shared ink threshold is not ink', () => {
    const page = measureWordPage(image(10, 10, data => {
      data[0] = data[1] = data[2] = 240; // distance 15 < inkBackgroundDistance 24
    }), 1);
    expect(page.inkBounds).toBeNull();
  });
});

test.describe('word-reference validation', () => {
  test('the seeded record validates against the corpus', () => {
    expect(validateWordReference(emptyRecord(VISUAL_PARITY_CORPUS), VISUAL_PARITY_CORPUS))
      .toEqual([]);
  });

  test('a measured case validates and coexists with pending rows', () => {
    expect(validateWordReference(recordWith([measuredCase()]), VISUAL_PARITY_CORPUS)).toEqual([]);
  });

  test('a missing corpus row is a problem — coverage cannot silently shrink', () => {
    const record = emptyRecord(VISUAL_PARITY_CORPUS);
    record.cases = record.cases.filter(entry => entry.id !== 'nested-table');
    expect(validateWordReference(record, VISUAL_PARITY_CORPUS).join(' '))
      .toContain('nested-table has no word-reference row');
  });

  test('a row matching no corpus case is a problem — evidence cannot outlive its case', () => {
    const record = recordWith([measuredCase({ id: 'retired-case' } as WordReferenceCase)]);
    expect(validateWordReference(record, VISUAL_PARITY_CORPUS).join(' '))
      .toContain('retired-case matches no corpus case');
  });

  test('a pending row carrying measurement fields is a problem', () => {
    const record = recordWith([
      { id: 'shape', status: 'pending', pageCount: 1 } as WordReferenceCase,
    ]);
    expect(validateWordReference(record, VISUAL_PARITY_CORPUS).join(' '))
      .toContain('shape is pending but carries measurement fields');
  });

  test('measured rows need a fixture hash and internally consistent pages', () => {
    expect(validateWordReference(
      recordWith([measuredCase({ fixtureSha256: undefined })]), VISUAL_PARITY_CORPUS).join(' '))
      .toContain('without a valid fixtureSha256');
    expect(validateWordReference(
      recordWith([measuredCase({ pageCount: 2 })]), VISUAL_PARITY_CORPUS).join(' '))
      .toContain('pageCount/pages disagree');
    expect(validateWordReference(
      recordWith([measuredCase({
        pages: [{ page: 1, width: 816, height: 1056, inkBounds: { left: 500, top: 0, right: 10, bottom: 5 }, inkPixelRatio: 0.01 }],
      })]), VISUAL_PARITY_CORPUS).join(' '))
      .toContain('ink bounds are inconsistent');
  });

  test('a wordEvidence citation without a measurement is a problem — the #402 contract', () => {
    const corpusWithCitation: VisualCorpusEntry[] = VISUAL_PARITY_CORPUS.map(entry =>
      entry.id === 'nested-table'
        ? { ...entry, disposition: { ...entry.disposition, wordEvidence: 'Word suppresses the spacing' } }
        : entry);
    expect(validateWordReference(emptyRecord(VISUAL_PARITY_CORPUS), corpusWithCitation).join(' '))
      .toContain('nested-table disposition cites wordEvidence but the word-reference row is pending');
    // The same citation is fine once the row is measured.
    expect(validateWordReference(
      recordWith([measuredCase({ id: 'nested-table' })]), corpusWithCitation)).toEqual([]);
  });

  test('measured cases without a recorded Word/OS environment are a problem', () => {
    const record = { ...recordWith([measuredCase()]), environment: null };
    expect(validateWordReference(record, VISUAL_PARITY_CORPUS).join(' '))
      .toContain('no environment');
  });

  test('upsertCase keeps the record sorted and replaces in place', () => {
    const record = recordWith([measuredCase()]);
    const again = upsertCase(record, measuredCase({ pageCount: 1 }));
    expect(again.cases.map(entry => entry.id)).toEqual(record.cases.map(entry => entry.id));
    expect(again.cases.filter(entry => entry.id === 'shape')).toHaveLength(1);
  });

  test('serialization is stable so a re-capture produces a reviewable diff', () => {
    const record = recordWith([measuredCase()]);
    expect(serializeWordReference(record)).toBe(serializeWordReference({ ...record }));
    expect(serializeWordReference(record).endsWith('\n')).toBe(true);
  });
});

test.describe('the committed word-reference record', () => {
  test('exists, is current-schema, and validates against the corpus', () => {
    const record = readWordReference();
    expect(record, `${WORD_REFERENCE_FILE} must be committed`).not.toBeNull();
    expect(record!.schemaVersion).toBe(WORD_REFERENCE_SCHEMA_VERSION);
    expect(validateWordReference(record!, VISUAL_PARITY_CORPUS)).toEqual([]);
  });

  test('is numbers-only: no image, path, or artifact leakage', () => {
    const serialized = readFileSync(WORD_REFERENCE_FILE, 'utf8');
    expect(serialized).not.toContain('.png');
    expect(serialized).not.toContain('.pdf');
    expect(serialized).not.toContain('artifact');
  });

  test('is tracked by Git and lives outside any ignored path', () => {
    execFileSync('git', ['ls-files', '--error-unmatch', relative(repoRoot, WORD_REFERENCE_FILE)], {
      cwd: repoRoot,
      stdio: 'pipe',
    });
  });
});

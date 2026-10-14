import { test } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { VISUAL_PARITY_CORPUS } from './visual-parity/corpus.js';
import { compareImages } from './visual-parity/metrics.js';
import { decodePng } from './visual-parity/png.js';
import {
  WORD_OS_ENV,
  WORD_PDFS_ENV,
  WORD_REFERENCE_FILE,
  WORD_RUN_ENV,
  WORD_VERSION_ENV,
  emptyRecord,
  measureWordPage,
  readWordReference,
  roundTo,
  serializeWordReference,
  upsertCase,
  validateWordReference,
  type WordReferenceCase,
  type WordReferenceComparison,
} from './visual-parity/word-reference.js';

/**
 * The Word-reference capture (issue #402): turns a directory of Word-exported `<case-id>.pdf`
 * files into committed, numbers-only measurements in `word-reference.json`.
 *
 * The manual part of the procedure is ONLY the export — a Word license holder opens each
 * tracked fixture and saves a PDF (see WORD_REFERENCE.md for the exact steps, including the
 * No Markup view for `revisionMode: 'accepted'` cases). This spec does everything downstream
 * with the benchmark's own contract: Poppler rasterization at exactly 96 DPI, the shared ink
 * model from metrics.ts, and stable record serialization. That is what makes the procedure
 * reproducible instead of artisanal.
 *
 * Optionally, pointing DOCXODUS_WORD_REFERENCE_RUN at a completed benchmark run's output root
 * also records the three-way comparison (Docxodus-vs-Word and LibreOffice-vs-Word) per case —
 * advisory pixel scores, since Word renders with genuine Office fonts while both engines use
 * the issue-#379 substitutes; the decisive artifacts are the structural measurements.
 *
 * Usage:
 *   DOCXODUS_WORD_REFERENCE_PDFS=/path/to/word-pdfs \
 *   DOCXODUS_WORD_VERSION="Word for Microsoft 365 MSO (16.0.18129)" \
 *   DOCXODUS_WORD_OS="Windows 11 24H2" \
 *   [DOCXODUS_WORD_REFERENCE_RUN=/tmp/docxodus-visual-parity] \
 *   npm run capture:word-reference
 */

test.skip(!process.env[WORD_PDFS_ENV],
  `set ${WORD_PDFS_ENV} to a directory of Word-exported <case-id>.pdf files`);

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');

function rasterize(pdfPath: string, work: string): string[] {
  execFileSync('pdftoppm', ['-r', '96', '-png', pdfPath, join(work, 'word-page')], {
    stdio: 'pipe',
    timeout: 120000,
    env: { ...process.env, LANG: 'C.UTF-8', LC_ALL: 'C.UTF-8', TZ: 'UTC' },
  });
  const pattern = /^word-page-(\d+)\.png$/;
  return readdirSync(work)
    .map(name => ({ name, match: name.match(pattern) }))
    .filter((item): item is { name: string; match: RegExpMatchArray } => item.match !== null)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]))
    .map(item => join(work, item.name));
}

function enginePages(runCaseDir: string, prefix: string): string[] {
  const pattern = new RegExp(`^${prefix}-(\\d+)\\.png$`);
  return readdirSync(runCaseDir)
    .map(name => ({ name, match: name.match(pattern) }))
    .filter((item): item is { name: string; match: RegExpMatchArray } => item.match !== null)
    .sort((a, b) => Number(a.match[1]) - Number(b.match[1]))
    .map(item => join(runCaseDir, item.name));
}

function compareEngineToWord(engine: string[], word: string[]) {
  const paired = Math.min(engine.length, word.length);
  let ssimSum = 0;
  let worstInkF1 = Number.POSITIVE_INFINITY;
  for (let index = 0; index < paired; index++) {
    const { metrics } = compareImages(
      decodePng(readFileSync(engine[index])),
      decodePng(readFileSync(word[index])),
    );
    ssimSum += metrics.ssim;
    worstInkF1 = Math.min(worstInkF1, metrics.tolerantInkF1);
  }
  // A page-count mismatch against Word is itself a finding; score the unpaired side as absent.
  if (engine.length !== word.length) worstInkF1 = 0;
  return {
    meanSsim: roundTo(paired ? ssimSum / paired : 0),
    worstInkF1: roundTo(paired ? worstInkF1 : 0),
  };
}

test('capture Word-reference measurements from exported PDFs', () => {
  test.setTimeout(10 * 60 * 1000);

  const pdfDir = resolve(process.env[WORD_PDFS_ENV]!);
  const wordVersion = process.env[WORD_VERSION_ENV];
  const os = process.env[WORD_OS_ENV];
  if (!wordVersion || !os) {
    throw new Error(`A capture must declare what it measured with: set ${WORD_VERSION_ENV} ` +
      `(e.g. "Word for Microsoft 365 MSO (16.0.18129)") and ${WORD_OS_ENV} (e.g. "Windows 11 24H2").`);
  }

  const pdfs = readdirSync(pdfDir).filter(name => name.toLowerCase().endsWith('.pdf'));
  if (!pdfs.length) throw new Error(`${pdfDir} contains no PDF files`);
  const unknown = pdfs.filter(name =>
    !VISUAL_PARITY_CORPUS.some(entry => `${entry.id}.pdf` === name));
  if (unknown.length) {
    throw new Error(`PDFs that match no corpus case id: ${unknown.join(', ')}. ` +
      `Expected names: ${VISUAL_PARITY_CORPUS.map(entry => `${entry.id}.pdf`).join(', ')}`);
  }

  const runRoot = process.env[WORD_RUN_ENV] ? resolve(process.env[WORD_RUN_ENV]!) : null;
  const runCommit = runRoot && existsSync(join(runRoot, 'summary.json'))
    ? (JSON.parse(readFileSync(join(runRoot, 'summary.json'), 'utf8')).gitCommit ?? 'unknown')
    : null;

  let record = readWordReference() ?? emptyRecord(VISUAL_PARITY_CORPUS);
  record = { ...record, environment: { word: wordVersion, os } };
  const comparisons: Record<string, WordReferenceComparison> = { ...record.comparisons };
  const capturedAt = new Date().toISOString().slice(0, 10);

  for (const name of pdfs) {
    const id = name.replace(/\.pdf$/i, '');
    const entry = VISUAL_PARITY_CORPUS.find(candidate => candidate.id === id)!;
    const fixturePath = resolve(repoRoot, entry.path);
    execFileSync('git', ['ls-files', '--error-unmatch', entry.path], { cwd: repoRoot, stdio: 'pipe' });

    const work = mkdtempSync(join(tmpdir(), `docxodus-word-ref-${id}-`));
    try {
      const pagePngs = rasterize(join(pdfDir, name), work);
      if (!pagePngs.length) throw new Error(`pdftoppm produced no pages for ${name}`);
      const pages = pagePngs.map((path, index) =>
        measureWordPage(decodePng(readFileSync(path)), index + 1));

      const existing = record.cases.find(candidate => candidate.id === id);
      const measured: WordReferenceCase = {
        id,
        status: 'measured',
        fixtureSha256: createHash('sha256').update(readFileSync(fixturePath)).digest('hex'),
        pageCount: pages.length,
        pages,
        // Operator-authored fields survive a re-capture; only measurements are regenerated.
        ...(existing?.keyMeasurements ? { keyMeasurements: existing.keyMeasurements } : {}),
        ...(existing?.notes ? { notes: existing.notes } : {}),
        capturedAt,
      };
      record = upsertCase(record, measured);

      if (runRoot) {
        const caseDir = join(runRoot, id);
        if (existsSync(caseDir)) {
          comparisons[id] = {
            runCommit: runCommit ?? 'unknown',
            docxodusVsWord: compareEngineToWord(enginePages(caseDir, 'docxodus'), pagePngs),
            libreofficeVsWord: compareEngineToWord(enginePages(caseDir, 'libreoffice'), pagePngs),
          };
        }
      }
      console.log(`${id}: ${pages.length} page(s) measured` +
        (comparisons[id] ? ' + three-way comparison' : ''));
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  }

  if (Object.keys(comparisons).length) record = { ...record, comparisons };

  const problems = validateWordReference(record, VISUAL_PARITY_CORPUS);
  if (problems.length) {
    throw new Error(`Refusing to write an inconsistent word-reference record:\n  ${problems.join('\n  ')}`);
  }
  writeFileSync(WORD_REFERENCE_FILE, serializeWordReference(record));
  const measuredCount = record.cases.filter(entry => entry.status === 'measured').length;
  console.log(`Word-reference record updated: ${WORD_REFERENCE_FILE} ` +
    `(${measuredCount}/${record.cases.length} cases measured). Review and commit the diff.`);
});

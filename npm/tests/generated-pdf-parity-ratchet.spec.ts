import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDF_PARITY_CORPUS } from './visual-parity/pdf-corpus.js';
import {
  RATCHET_SCHEMA_VERSION,
  RATCHET_TOLERANCE,
  buildRecord,
  compareToRecord,
  readRecord,
  type RatchetRecord,
  type RatchetSummary,
} from './visual-parity/ratchet.js';

/**
 * Pure checks for #443's committed generated-PDF record. These run on ordinary PRs without
 * LibreOffice, Poppler, or Chromium so a corpus/disposition change cannot silently escape the
 * scheduled release ratchet.
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const recordFile = resolve(__dirname, 'visual-parity/generated-pdf-ratchet.json');

/**
 * A synthetic already-severe baseline. Deriving this from the committed record instead would make
 * the two tests below crash the moment a refresh clears the last severe case — which is the
 * project's declared direction (#444) — and they are the ONLY coverage of the physicalGeometry and
 * semantics branches in compareCase, so the crash would land exactly where coverage is unique.
 */
const SEVERE_BASELINE: RatchetSummary = {
  gitCommit: 'c'.repeat(40),
  workingTreeDirty: false,
  environment: {
    chromium: '143.0.7499.4',
    libreoffice: 'LibreOffice 25.8.7.3 580(Build:3)',
    pdftoppm: 'pdftoppm version 24.02.0',
    fontContract: { sha256: 'abc123' },
  },
  cases: [{
    id: 'pdf-synthetic-severe',
    disposition: { kind: 'renderer-bug' },
    docxodusPages: 1,
    libreofficePages: 1,
    severity: 'severe',
    pages: [{ ssim: 0.5, tolerantInkF1: 0.5 }],
    physicalGeometryPassed: true,
    semanticChecksPassed: true,
  }],
};

function measuredSummary(record: RatchetRecord): RatchetSummary {
  return {
    gitCommit: record.sourceCommit,
    workingTreeDirty: false,
    environment: {
      chromium: record.environment.chromium,
      libreoffice: record.environment.libreoffice,
      pdftoppm: record.environment.poppler,
      fontContract: { sha256: record.environment.fontContractSha256 },
    },
    cases: record.cases.map((entry) => ({
      id: entry.id,
      disposition: { kind: entry.disposition },
      docxodusPages: entry.pages.docxodus,
      libreofficePages: entry.pages.libreoffice,
      severity: entry.severity,
      pages: Array.from({ length: entry.pages.docxodus }, () => ({
        ssim: entry.ssim,
        tolerantInkF1: entry.worstInkF1,
      })),
      physicalGeometryPassed: true,
      semanticChecksPassed: true,
    })),
  };
}

test.describe('the committed generated-PDF parity ratchet', () => {
  test('is current-schema, complete, successful, and numbers-only', () => {
    const record = readRecord(recordFile);
    expect(record, `${recordFile} must exist`).not.toBeNull();
    expect(record!.schemaVersion).toBe(RATCHET_SCHEMA_VERSION);
    expect(record!.tolerance).toEqual(RATCHET_TOLERANCE);
    expect(record!.sourceCommit).toMatch(/^[0-9a-f]{40}$/);
    expect(record!.cases.map((entry) => entry.id).sort())
      .toEqual(PDF_PARITY_CORPUS.cases.map((entry) => entry.id).sort());

    for (const entry of record!.cases) {
      expect(Object.keys(entry).sort()).toEqual(
        ['disposition', 'id', 'pages', 'severity', 'ssim', 'worstInkF1']);
      expect(entry.pages.docxodus).toBeGreaterThan(0);
      expect(entry.pages.libreoffice).toBe(entry.pages.docxodus);
      expect(entry.ssim).toBeGreaterThanOrEqual(0);
      expect(entry.ssim).toBeLessThanOrEqual(1);
      expect(entry.worstInkF1).toBeGreaterThanOrEqual(0);
      expect(entry.worstInkF1).toBeLessThanOrEqual(1);
    }

    const serialized = readFileSync(recordFile, 'utf8');
    expect(serialized).not.toContain('.png');
    expect(serialized).not.toContain('.pdf');
    expect(serialized.match(/[0-9a-f]{64}/g) ?? [])
      .toEqual([record!.environment.fontContractSha256]);
  });

  test('the public generated-PDF command owns both builds before it launches Playwright', () => {
    const packageJson = JSON.parse(readFileSync(resolve(__dirname, '../package.json'), 'utf8')) as {
      scripts: Record<string, string>;
    };
    const command = packageJson.scripts['test:generated-pdf-parity'];
    expect(command).toContain('npm run build');
    expect(command).toContain('npm --prefix ../npm-export run build');
    expect(command).toContain('npm run pretest');
    expect(command).toContain('playwright test generated-pdf-parity.spec.ts');
    expect(command.indexOf('npm run build')).toBeLessThan(
      command.indexOf('npm --prefix ../npm-export run build'),
    );
    expect(command.indexOf('npm --prefix ../npm-export run build')).toBeLessThan(
      command.indexOf('playwright test generated-pdf-parity.spec.ts'),
    );
  });

  test('agrees with the corpus about every reviewable disposition', () => {
    const record = readRecord(recordFile)!;
    for (const entry of record.cases) {
      const corpusEntry = PDF_PARITY_CORPUS.cases.find((candidate) => candidate.id === entry.id)!;
      expect(entry.disposition, `${entry.id} disposition drifted from pdf-corpus.ts`)
        .toBe(corpusEntry.disposition.kind);
    }
  });

  test('an unchanged already-severe baseline holds', () => {
    const record = buildRecord(SEVERE_BASELINE, '2026-08-16');
    expect(compareToRecord(record, SEVERE_BASELINE).status).toBe('ok');
  });

  test('rejects physical-geometry failure even for an already-severe raster baseline', () => {
    const record = buildRecord(SEVERE_BASELINE, '2026-08-16');
    const summary = structuredClone(SEVERE_BASELINE);
    summary.cases[0].physicalGeometryPassed = false;

    const comparison = compareToRecord(record, summary);
    expect(comparison.status).toBe('regressed');
    expect(comparison.findings).toContainEqual(expect.objectContaining({
      caseId: 'pdf-synthetic-severe',
      signal: 'physicalGeometry',
    }));
  });

  test('rejects semantic failure even for an already-severe raster baseline', () => {
    const record = buildRecord(SEVERE_BASELINE, '2026-08-16');
    const summary = structuredClone(SEVERE_BASELINE);
    summary.cases[0].semanticChecksPassed = false;

    const comparison = compareToRecord(record, summary);
    expect(comparison.status).toBe('regressed');
    expect(comparison.findings).toContainEqual(expect.objectContaining({
      caseId: 'pdf-synthetic-severe',
      signal: 'semantics',
    }));
  });

  test('remediation messages name THIS benchmark\'s refresh switch, not the other one', () => {
    const record = buildRecord(SEVERE_BASELINE, '2026-08-16');
    const drifted = structuredClone(SEVERE_BASELINE);
    drifted.environment.libreoffice = 'LibreOffice 26.2.0.1 100(Build:1)';
    const comparison = compareToRecord(record, drifted, {
      expectComplete: true,
      updateRecordEnv: 'DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD',
    });

    expect(comparison.status).toBe('environment-changed');
    expect(comparison.message).toContain('DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD=1');
    // Applying the browser-page switch to this command refreshes nothing here and DOES
    // overwrite the other benchmark's record, so it must not appear in this guidance.
    expect(comparison.message).not.toContain('DOCXODUS_VISUAL_PARITY_UPDATE_RECORD');
  });

  test('the generated-PDF benchmark passes that switch name to the comparator', () => {
    const spec = readFileSync(resolve(__dirname, 'generated-pdf-parity.spec.ts'), 'utf8');
    expect(spec).toContain("updateRecordEnv: 'DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD'");
  });

  test('a provisional record refuses numeric comparison instead of blaming the renderer', () => {
    const record = readRecord(recordFile)!;
    // The committed record was measured off this lineage. Feeding the comparator numbers that
    // match it EXACTLY must still refuse: the environment fingerprint covers LibreOffice,
    // Chromium, Poppler and fonts, so a source-lineage difference shows no drift and would
    // otherwise be compared numerically and attributed to Docxodus.
    expect(record!.provisional).toBe(true);
    const comparison = compareToRecord(record, measuredSummary(record), {
      expectComplete: true,
      updateRecordEnv: 'DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD',
    });
    expect(comparison.status).toBe('record-mismatch');
    expect(comparison.message).toContain('provisional');
    expect(comparison.message).toContain('DOCXODUS_GENERATED_PDF_PARITY_UPDATE_RECORD=1');
  });

  test('clearing the flag restores ordinary numeric comparison', () => {
    const record = readRecord(recordFile)!;
    const promoted = { ...record, provisional: undefined };
    expect(compareToRecord(promoted, measuredSummary(record)).status).toBe('ok');
  });

  test('a refresh never re-emits the provisional flag', () => {
    expect(buildRecord(SEVERE_BASELINE, '2026-08-16')).not.toHaveProperty('provisional');
  });
});

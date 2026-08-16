import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { PDF_PARITY_CORPUS } from './visual-parity/pdf-corpus.js';
import {
  RATCHET_SCHEMA_VERSION,
  RATCHET_TOLERANCE,
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

  test('agrees with the corpus about every reviewable disposition', () => {
    const record = readRecord(recordFile)!;
    for (const entry of record.cases) {
      const corpusEntry = PDF_PARITY_CORPUS.cases.find((candidate) => candidate.id === entry.id)!;
      expect(entry.disposition, `${entry.id} disposition drifted from pdf-corpus.ts`)
        .toBe(corpusEntry.disposition.kind);
    }
  });

  test('rejects physical-geometry failure even for an already-severe raster baseline', () => {
    const record = readRecord(recordFile)!;
    const summary = measuredSummary(record);
    const severe = summary.cases.find((entry) => entry.severity === 'severe')!;
    severe.physicalGeometryPassed = false;

    const comparison = compareToRecord(record, summary);
    expect(comparison.status).toBe('regressed');
    expect(comparison.findings).toContainEqual(expect.objectContaining({
      caseId: severe.id,
      signal: 'physicalGeometry',
    }));
  });

  test('rejects semantic failure even for an already-severe raster baseline', () => {
    const record = readRecord(recordFile)!;
    const summary = measuredSummary(record);
    const severe = summary.cases.find((entry) => entry.severity === 'severe')!;
    severe.semanticChecksPassed = false;

    const comparison = compareToRecord(record, summary);
    expect(comparison.status).toBe('regressed');
    expect(comparison.findings).toContainEqual(expect.objectContaining({
      caseId: severe.id,
      signal: 'semantics',
    }));
  });
});

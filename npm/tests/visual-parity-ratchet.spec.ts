import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { VISUAL_PARITY_CORPUS } from './visual-parity/corpus.js';
import {
  KNOWN_LIBREOFFICE_VERSION_DIFFERENCES,
  LIBREOFFICE_CONTRACT,
  LIBREOFFICE_DRIFT_ENV,
  checkLibreOfficeContract,
} from './visual-parity/environment-contract.js';
import {
  RATCHET_PRECISION,
  RATCHET_RECORD_FILE,
  RATCHET_SCHEMA_VERSION,
  RATCHET_TOLERANCE,
  assertRecordUpdateProvenance,
  buildRecord,
  chromiumFingerprint,
  compareToRecord,
  libreofficeFingerprint,
  popplerFingerprint,
  readRecord,
  serializeRecord,
  type RatchetSummary,
  type RatchetSummaryCase,
} from './visual-parity/ratchet.js';

/**
 * The ratchet's own regression suite (issue #395).
 *
 * Deliberately NOT gated behind `DOCXODUS_VISUAL_PARITY`: the comparison layer is pure, so it runs
 * on every pull request without LibreOffice, Poppler, or a renderer. That is what makes the first
 * acceptance criterion — "a deliberately introduced regression fails the run naming the case and
 * signal" — a continuously proven property rather than a claim demonstrated once by hand. Breaking
 * a renderer on purpose to test the alarm is neither repeatable nor safe; feeding the comparator a
 * deliberately worsened summary is both.
 */

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, '../..');

const ENVIRONMENT = {
  chromium: '143.0.7499.4',
  libreoffice: 'LibreOffice 25.8.7.3 580(Build:3)',
  pdftoppm: 'pdftoppm version 24.02.0',
  fontContract: { sha256: 'abc123' },
};

function summaryCase(overrides: Partial<RatchetSummaryCase> = {}): RatchetSummaryCase {
  return {
    id: 'shape',
    disposition: { kind: 'renderer-bug' },
    docxodusPages: 1,
    libreofficePages: 1,
    severity: 'severe',
    pages: [{ ssim: 0.97428, tolerantInkF1: 0.63049 }],
    ...overrides,
  };
}

function summaryOf(cases: RatchetSummaryCase[]): RatchetSummary {
  return { gitCommit: 'deadbeef', environment: ENVIRONMENT, cases };
}

const BASELINE_SUMMARY = summaryOf([
  summaryCase(),
  summaryCase({
    id: 'multi-section',
    disposition: { kind: 'environment' },
    docxodusPages: 6,
    libreofficePages: 6,
    severity: 'close',
    // A worst-page aggregation must not average a bad page away.
    pages: [
      { ssim: 0.99934, tolerantInkF1: 0.99797 },
      { ssim: 0.99940, tolerantInkF1: 0.99900 },
    ],
  }),
]);

const BASELINE_RECORD = buildRecord(BASELINE_SUMMARY, '2026-08-11');

test.describe('visual parity ratchet', () => {
  test('an unchanged run holds against its own record', () => {
    const comparison = compareToRecord(BASELINE_RECORD, BASELINE_SUMMARY);
    expect(comparison.status).toBe('ok');
    expect(comparison.findings).toEqual([]);
    expect(comparison.improvements).toEqual([]);
  });

  test('records mean SSIM and WORST ink F1, not the mean of both', () => {
    const multiSection = BASELINE_RECORD.cases.find(entry => entry.id === 'multi-section')!;
    expect(multiSection.ssim).toBeCloseTo((0.99934 + 0.99940) / 2, RATCHET_PRECISION);
    expect(multiSection.worstInkF1).toBe(0.99797);
  });

  test('a deliberately worsened SSIM fails, naming the case and the signal', () => {
    const regressed = summaryOf([
      summaryCase({ pages: [{ ssim: 0.95000, tolerantInkF1: 0.63049 }] }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, regressed);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings).toHaveLength(1);
    expect(comparison.findings[0].caseId).toBe('shape');
    expect(comparison.findings[0].signal).toBe('ssim');
    expect(comparison.message).toContain('shape');
    expect(comparison.message).toContain('mean SSIM');
    expect(comparison.message).toContain('0.97428');
    expect(comparison.message).toContain('0.95000');
  });

  test('a deliberately worsened ink F1 fails, naming the case and the signal', () => {
    const regressed = summaryOf([
      summaryCase({ pages: [{ ssim: 0.97428, tolerantInkF1: 0.40000 }] }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, regressed);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings.map(finding => finding.signal)).toEqual(['inkF1']);
    expect(comparison.findings[0].caseId).toBe('shape');
    expect(comparison.message).toContain('worst ink F1');
  });

  test('a worsened severity fails even when SSIM and ink F1 hold', () => {
    const regressed = summaryOf([
      summaryCase(),
      { ...BASELINE_SUMMARY.cases[1], severity: 'minor' as const },
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, regressed);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings.map(finding => finding.signal)).toEqual(['severity']);
    expect(comparison.message).toContain('multi-section');
    expect(comparison.message).toContain('close to minor');
  });

  test('a page-count change fails — the loudest regression signal', () => {
    const regressed = summaryOf([
      summaryCase({ docxodusPages: 2 }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, regressed);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings.map(finding => finding.signal)).toEqual(['pages']);
    expect(comparison.message).toContain('page count changed from 1/1');
  });

  test('a new conversion error fails and suppresses its now-meaningless numbers', () => {
    const regressed = summaryOf([
      summaryCase({ error: 'LibreOffice did not produce a PDF', pages: [] }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, regressed);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings.map(finding => finding.signal)).toEqual(['error']);
    expect(comparison.message).toContain('LibreOffice did not produce a PDF');
  });

  test('an improved case does NOT fail, and is reported so the refresh is informed', () => {
    const improved = summaryOf([
      summaryCase({ severity: 'close', pages: [{ ssim: 0.99000, tolerantInkF1: 0.96000 }] }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, improved);

    expect(comparison.status).toBe('ok');
    expect(comparison.findings).toEqual([]);
    expect(comparison.improvements.map(finding => finding.signal).sort())
      .toEqual(['inkF1', 'severity', 'ssim']);
    expect(comparison.message).toContain('refresh the record to bank them');
  });

  test('movement within tolerance is neither a regression nor an improvement', () => {
    const jittered = summaryOf([
      summaryCase({
        pages: [{
          ssim: 0.97428 - RATCHET_TOLERANCE.ssim / 2,
          tolerantInkF1: 0.63049 - RATCHET_TOLERANCE.inkF1 / 2,
        }],
      }),
      BASELINE_SUMMARY.cases[1],
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, jittered);

    expect(comparison.status).toBe('ok');
    expect(comparison.improvements).toEqual([]);
  });

  test('movement just beyond tolerance is a regression', () => {
    const regressed = summaryOf([
      summaryCase({
        pages: [{ ssim: 0.97428 - RATCHET_TOLERANCE.ssim * 2, tolerantInkF1: 0.63049 }],
      }),
      BASELINE_SUMMARY.cases[1],
    ]);
    expect(compareToRecord(BASELINE_RECORD, regressed).status).toBe('regressed');
  });

  /**
   * The false-accusation guard. CI installs LibreOffice from unpinned `ubuntu-latest` apt (pinning
   * it is issue #403), and BASELINE.md documents how far a reference change moves the numbers —
   * LibreOffice 24.2 draws a different footnote separator than 25.8. Comparing across that would
   * blame Docxodus for someone else's release.
   */
  test('a reference-environment change is reported as such, never as a renderer regression', () => {
    const upgraded: RatchetSummary = {
      ...BASELINE_SUMMARY,
      environment: { ...ENVIRONMENT, libreoffice: 'LibreOffice 26.2.0.1 100(Build:1)' },
      // Numbers that WOULD read as a severe regression if compared naively.
      cases: [
        summaryCase({ severity: 'severe', pages: [{ ssim: 0.10000, tolerantInkF1: 0.10000 }] }),
        BASELINE_SUMMARY.cases[1],
      ],
    };
    const comparison = compareToRecord(BASELINE_RECORD, upgraded);

    expect(comparison.status).toBe('environment-changed');
    expect(comparison.findings).toEqual([]);
    expect(comparison.message).toContain('NOT a renderer regression');
    expect(comparison.message).toContain('record refresh required');
    expect(comparison.environmentDrift.join(' ')).toContain('recorded 25.8, measured 26.2');
  });

  test('a font-contract edit is an environment change, since it moves both renderers', () => {
    const comparison = compareToRecord(BASELINE_RECORD, {
      ...BASELINE_SUMMARY,
      environment: { ...ENVIRONMENT, fontContract: { sha256: 'different' } },
    });
    expect(comparison.status).toBe('environment-changed');
    expect(comparison.environmentDrift.join(' ')).toContain('fontContractSha256');
  });

  test('a LibreOffice patch release inside one minor stays comparable', () => {
    const comparison = compareToRecord(BASELINE_RECORD, {
      ...BASELINE_SUMMARY,
      environment: { ...ENVIRONMENT, libreoffice: 'LibreOffice 25.8.9.1 580(Build:9)' },
    });
    expect(comparison.status).toBe('ok');
  });

  test('fingerprints reduce versions to their layout-relevant boundary', () => {
    expect(libreofficeFingerprint('LibreOffice 25.8.7.3 580(Build:3)')).toBe('25.8');
    expect(libreofficeFingerprint(undefined)).toBe('unknown');
    expect(chromiumFingerprint('143.0.7499.4')).toBe('143');
    expect(chromiumFingerprint(undefined)).toBe('unknown');
    expect(popplerFingerprint('pdftoppm version 24.02.0')).toBe('24.02');
    expect(popplerFingerprint(undefined)).toBe('unknown');
  });

  test('a Poppler bump is an environment change: the rasterizer is part of the measurement', () => {
    const comparison = compareToRecord(BASELINE_RECORD, {
      ...BASELINE_SUMMARY,
      environment: { ...ENVIRONMENT, pdftoppm: 'pdftoppm version 25.03.0' },
    });
    expect(comparison.status).toBe('environment-changed');
    expect(comparison.environmentDrift.join(' ')).toContain('poppler: recorded 24.02, measured 25.03');
  });

  test('a case dropped from an unfiltered run fails, but not from a filtered one', () => {
    const partial = summaryOf([summaryCase()]);

    expect(compareToRecord(BASELINE_RECORD, partial, { expectComplete: true }).status)
      .toBe('regressed');
    expect(compareToRecord(BASELINE_RECORD, partial, { expectComplete: true })
      .findings.map(finding => finding.signal)).toEqual(['missing']);
    expect(compareToRecord(BASELINE_RECORD, partial, { expectComplete: false }).status).toBe('ok');
  });

  test('a corpus case with no record entry fails until the record is refreshed', () => {
    const extended = summaryOf([
      ...BASELINE_SUMMARY.cases,
      summaryCase({ id: 'nested-table', disposition: { kind: 'unattributed' } }),
    ]);
    const comparison = compareToRecord(BASELINE_RECORD, extended);

    expect(comparison.status).toBe('regressed');
    expect(comparison.findings.map(finding => finding.signal)).toEqual(['unrecorded']);
    expect(comparison.message).toContain('nested-table');
  });

  test('an unreadable schema version demands regeneration instead of comparing', () => {
    const comparison = compareToRecord(
      { ...BASELINE_RECORD, schemaVersion: RATCHET_SCHEMA_VERSION + 1 }, BASELINE_SUMMARY);
    expect(comparison.status).toBe('record-mismatch');
    expect(comparison.message).toContain('DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1');
  });

  test('the record serializes stably, so a rerun produces no diff churn', () => {
    const again = buildRecord(BASELINE_SUMMARY, '2026-08-11');
    expect(serializeRecord(again)).toBe(serializeRecord(BASELINE_RECORD));
    expect(serializeRecord(BASELINE_RECORD).endsWith('\n')).toBe(true);
    // Sorted by id so corpus reordering cannot churn the committed diff.
    expect(BASELINE_RECORD.cases.map(entry => entry.id)).toEqual(['multi-section', 'shape']);
  });

  test('record refresh requires an exact clean source commit', () => {
    const clean = {
      ...BASELINE_SUMMARY,
      gitCommit: 'a'.repeat(40),
      workingTreeDirty: false,
    };
    expect(() => assertRecordUpdateProvenance(clean)).not.toThrow();
    expect(() => assertRecordUpdateProvenance({ ...clean, workingTreeDirty: true }))
      .toThrow(/dirty or unverified worktree/);
    expect(() => assertRecordUpdateProvenance({ ...clean, gitCommit: 'deadbeef' }))
      .toThrow(/40-character source commit/);
  });
});

/**
 * The reference-version contract (issue #403). Pure decision layer, so — like the ratchet alarm
 * itself — the failure mode is proven on every pull request without installing an out-of-contract
 * LibreOffice on purpose.
 */
test.describe('LibreOffice reference-version contract', () => {
  test('accepts any build of the contract minor', () => {
    expect(checkLibreOfficeContract('LibreOffice 25.8.7.3 580(Build:3)').ok).toBe(true);
    expect(checkLibreOfficeContract('LibreOffice 25.8.9.1 580(Build:9)').ok).toBe(true);
  });

  test('rejects an out-of-contract version with install guidance, never a silent number shift', () => {
    const check = checkLibreOfficeContract('LibreOffice 24.2.7.2 420(Build:2)');
    expect(check.ok).toBe(false);
    expect(check.fingerprint).toBe('24.2');
    expect(check.message).toContain(`contract requires ${LIBREOFFICE_CONTRACT.version}`);
    expect(check.message).toContain(LIBREOFFICE_CONTRACT.archiveUrl);
    // The TDF bundled-font trap (issue #400 finding) travels with the install guidance.
    expect(check.message).toContain('REMOVE its bundled Carlito/Caladea/Liberation fonts');
    expect(check.message).toContain(LIBREOFFICE_DRIFT_ENV);
  });

  test('cites the known cross-version rendering differences in the failure', () => {
    const check = checkLibreOfficeContract('LibreOffice 24.2.7.2 420(Build:2)');
    for (const difference of KNOWN_LIBREOFFICE_VERSION_DIFFERENCES) {
      expect(check.message).toContain(difference.behavior);
    }
    // The catalogue must stay non-empty: the footnote-separator finding is the reason the
    // contract exists, and each future finding lands here so the failure message teaches it.
    expect(KNOWN_LIBREOFFICE_VERSION_DIFFERENCES.length).toBeGreaterThan(0);
  });

  test('a missing LibreOffice reads as unknown and still fails with guidance', () => {
    const check = checkLibreOfficeContract('');
    expect(check.ok).toBe(false);
    expect(check.fingerprint).toBe('unknown');
    expect(check.message).toContain('no version output');
  });

  test('agrees with the committed record: one declared version, two enforcement points', () => {
    // The contract module owns the version; the record's fingerprint must be the SAME version,
    // or the run-start assertion and the ratchet's environment check would disagree about which
    // LibreOffice the numbers mean anything against.
    expect(readRecord()!.environment.libreoffice).toBe(LIBREOFFICE_CONTRACT.version);
  });

  test('the declared build belongs to the declared minor', () => {
    expect(LIBREOFFICE_CONTRACT.build.startsWith(`${LIBREOFFICE_CONTRACT.version}.`)).toBe(true);
    expect(LIBREOFFICE_CONTRACT.archiveUrl).toContain(LIBREOFFICE_CONTRACT.build);
    expect(LIBREOFFICE_CONTRACT.archiveSignatureUrl).toBe(`${LIBREOFFICE_CONTRACT.archiveUrl}.asc`);
    expect(LIBREOFFICE_CONTRACT.signingKeyFingerprint).toMatch(/^[0-9A-F]{40}$/);
  });

  test('CI verifies the detached signature and exact signing-key fingerprint before extraction', () => {
    const workflow = readFileSync(resolve(__dirname, '../../.github/workflows/visual-parity.yml'), 'utf8');
    expect(workflow).toContain(LIBREOFFICE_CONTRACT.archiveSignatureUrl);
    expect(workflow).toContain(LIBREOFFICE_CONTRACT.signingKeyFingerprint);
    expect(workflow).toContain('--verify /tmp/libreoffice.tar.gz.asc /tmp/libreoffice.tar.gz');
    expect(workflow.indexOf('--verify /tmp/libreoffice.tar.gz.asc'))
      .toBeLessThan(workflow.indexOf('tar -xzf /tmp/libreoffice.tar.gz'));
  });

  test('CI never cancels an in-flight benchmark, and budgets for configured retries', () => {
    const workflow = readFileSync(resolve(__dirname, '../../.github/workflows/visual-parity.yml'), 'utf8');
    // The group collapses to github.ref, so cancelling would let a manual dispatch kill the
    // scheduled traversal that the ratchet depends on.
    expect(workflow).toContain('cancel-in-progress: false');
    const budget = Number(workflow.match(/timeout-minutes:\s*(\d+)/)?.[1]);
    const config = readFileSync(resolve(__dirname, '../playwright.config.ts'), 'utf8');
    const retries = Number(config.match(/retries:\s*process\.env\.CI\s*\?\s*(\d+)/)?.[1]);
    // A job timeout is a cancellation, so an under-budgeted job SKIPS the generated-PDF step
    // and uploads only its bootstrap page.
    expect(budget).toBeGreaterThanOrEqual((retries + 1) * 50);
  });
});

test.describe('the committed ratchet record', () => {
  test('exists, is current-schema, and is a numbers-only file', () => {
    const record = readRecord();
    expect(record, `${RATCHET_RECORD_FILE} must be committed`).not.toBeNull();
    expect(record!.schemaVersion).toBe(RATCHET_SCHEMA_VERSION);
    expect(record!.tolerance).toEqual({
      ssim: RATCHET_TOLERANCE.ssim,
      inkF1: RATCHET_TOLERANCE.inkF1,
    });

    // No image, path, or per-artifact hash may leak into the record: it is reviewed as a diff, and
    // artifact naming must never churn it. The only digest is the font contract's.
    const serialized = readFileSync(RATCHET_RECORD_FILE, 'utf8');
    expect(serialized).not.toContain('.png');
    expect(serialized).not.toContain('artifact');
    expect(serialized.match(/[0-9a-f]{64}/g) ?? []).toEqual([record!.environment.fontContractSha256]);
    for (const entry of record!.cases) {
      expect(Object.keys(entry).sort()).toEqual(
        entry.error === undefined
          ? ['disposition', 'id', 'pages', 'severity', 'ssim', 'worstInkF1']
          : ['disposition', 'error', 'id', 'pages', 'severity', 'ssim', 'worstInkF1']);
    }
  });

  test('covers exactly the corpus, so adding a case cannot skip the ratchet', () => {
    const record = readRecord()!;
    expect(record.cases.map(entry => entry.id).sort())
      .toEqual(VISUAL_PARITY_CORPUS.map(entry => entry.id).sort());
  });

  test('agrees with corpus.ts about every disposition', () => {
    const record = readRecord()!;
    for (const entry of record.cases) {
      const corpusEntry = VISUAL_PARITY_CORPUS.find(candidate => candidate.id === entry.id)!;
      expect(entry.disposition, `${entry.id} disposition drifted from corpus.ts`)
        .toBe(corpusEntry.disposition.kind);
    }
  });

  /**
   * The record's environment fingerprint must describe the environment it was measured in.
   * `fonts.conf` is the one component that lives in the repository, so it is the one a PR can
   * silently desynchronize — editing the contract without rerunning the benchmark would leave
   * numbers attributed to a substitution set that no longer exists.
   */
  test('pins the font contract actually committed alongside it', () => {
    const record = readRecord()!;
    const contract = resolve(__dirname, 'visual-parity/fonts.conf');
    expect(record.environment.fontContractSha256)
      .toBe(createHash('sha256').update(readFileSync(contract)).digest('hex'));
  });

  test('is tracked by Git and lives outside any ignored path', () => {
    execFileSync('git', ['ls-files', '--error-unmatch', relative(repoRoot, RATCHET_RECORD_FILE)], {
      cwd: repoRoot,
      stdio: 'pipe',
    });
  });
});

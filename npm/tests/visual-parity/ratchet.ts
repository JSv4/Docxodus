import { existsSync, readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import type { VisualSeverity } from './metrics.js';

/**
 * The regression ratchet (issue #395).
 *
 * The scheduled run used to upload an artifact nobody compared against anything: a renderer
 * regression was caught only if a human downloaded and eyeballed the report inside the 14-day
 * retention window. Full-strict mode stays unreachable while renderer-attributable severe cases
 * remain, but "no case may get worse than recorded" is enforceable today.
 *
 * This module is the pure comparison layer: it turns a run's `summary.json` into a numbers-only
 * per-case record, and compares a fresh summary against a committed record. It deliberately holds
 * no I/O so the spec can exercise it against synthetic summaries on every CI run, without
 * LibreOffice, without a renderer, and without actually breaking one.
 *
 * Two properties make a tight tolerance honest:
 *
 *  - Within one environment the benchmark is deterministic. BASELINE.md records two clean full
 *    passes producing identical normalized metrics and identical PNG SHA-256s for all 60 images,
 *    so in-environment noise is zero rather than merely small.
 *  - Across environments it is not: LibreOffice, Chromium, and the substituted fonts each move
 *    the numbers materially. The record therefore carries an environment fingerprint, and a
 *    mismatch is reported as `environment-changed` — never as a renderer regression.
 *
 * The ratchet is deliberately BROADER than strict mode. Strict mode gates only severe cases whose
 * disposition attributes them to the renderer; the ratchet covers every case at every severity,
 * because a `close` case silently sliding to `minor` is exactly the drift the weekly run existed
 * to catch and never did.
 */

// Schema 2 (issue #403): the environment fingerprint gained `poppler` — pdftoppm's rasterizer
// sits between LibreOffice's PDF and every recorded number, so a distro bump must surface as
// environment drift, not as a silent shift attributed to the renderer.
export const RATCHET_SCHEMA_VERSION = 2;

/**
 * The committed record. It lives beside the corpus it describes so a rendering PR touches the
 * corpus, the record, and the baseline narrative in one reviewable diff.
 */
export const RATCHET_RECORD_FILE =
  resolve(dirname(fileURLToPath(import.meta.url)), 'ratchet.json');

/**
 * Movement a run may show against the record without being called a regression.
 *
 * Zero would be defensible on the determinism evidence above, but the fingerprint is coarser than
 * the environment: it pins LibreOffice to major.minor and Chromium to its major, so a patch-level
 * update inside one fingerprint can still perturb rasterization slightly. These values absorb that
 * and nothing more — every renderer movement BASELINE.md records for a real fix is at least an
 * order of magnitude larger, the smallest being the two-inch footnote separator at +0.000128 SSIM
 * and +0.003537 ink F1.
 *
 * Loosening these needs the same evidence bar as any other threshold change: a reproduced
 * false positive, not a run that was inconvenient to explain.
 */
export const RATCHET_TOLERANCE = {
  ssim: 0.0005,
  inkF1: 0.001,
} as const;

/** Decimal places the record stores, so float summation order cannot churn the committed diff. */
export const RATCHET_PRECISION = 5;

const SEVERITY_ORDER: VisualSeverity[] = ['close', 'minor', 'major', 'severe'];

export interface RatchetEnvironment {
  /** LibreOffice major.minor — pinned by the reference-version contract (issue #403). */
  libreoffice: string;
  /** Chromium major. */
  chromium: string;
  /** Poppler (pdftoppm) major.minor — the rasterizer between the reference PDF and every number. */
  poppler: string;
  /** SHA-256 of `fonts.conf`; changing the contract is a deliberate, reviewable act. */
  fontContractSha256: string;
}

export interface RatchetCase {
  id: string;
  /** Recorded so a disposition edit shows up in the record's diff, not only in `corpus.ts`. */
  disposition: string;
  pages: { docxodus: number; libreoffice: number };
  severity: VisualSeverity;
  /** Mean SSIM over paired pages — the same aggregation BASELINE.md's per-case table reports. */
  ssim: number;
  /** WORST paired-page ink F1, so one blank or disjoint page cannot be averaged away. */
  worstInkF1: number;
  /** Present only when the case failed to convert; a run may never introduce one. */
  error?: true;
}

export interface RatchetRecord {
  schemaVersion: number;
  description: string;
  recordedAt: string;
  sourceCommit: string;
  environment: RatchetEnvironment;
  tolerance: { ssim: number; inkF1: number };
  cases: RatchetCase[];
}

/** The subset of `summary.json` the ratchet reads. */
export interface RatchetSummaryCase {
  id: string;
  disposition: { kind: string };
  docxodusPages: number;
  libreofficePages: number;
  severity: VisualSeverity;
  pages: { ssim: number; tolerantInkF1: number }[];
  error?: string;
}

export interface RatchetSummary {
  gitCommit?: string;
  environment: {
    chromium?: string;
    libreoffice?: string;
    pdftoppm?: string;
    fontContract?: { sha256?: string };
  };
  cases: RatchetSummaryCase[];
}

export type RatchetStatus = 'ok' | 'regressed' | 'environment-changed' | 'record-mismatch';

export interface RatchetOptions {
  /**
   * Whether the run covered the whole corpus. A `DOCXODUS_VISUAL_PARITY_FILTER` run legitimately
   * measures a subset, so its absent cases are not findings; an unfiltered run's are.
   */
  expectComplete: boolean;
}

export interface RatchetFinding {
  caseId: string;
  /** The signal that moved, so the failure names both the case and what got worse. */
  signal: 'severity' | 'ssim' | 'inkF1' | 'pages' | 'error' | 'missing' | 'unrecorded';
  recorded: string;
  measured: string;
  detail: string;
}

export interface RatchetComparison {
  status: RatchetStatus;
  findings: RatchetFinding[];
  /** Cases that measurably IMPROVED — reported so the record refresh is an informed act. */
  improvements: RatchetFinding[];
  environmentDrift: string[];
  message: string;
}

export function round(value: number): number {
  const factor = 10 ** RATCHET_PRECISION;
  return Math.round(value * factor) / factor;
}

const format = (value: number): string => value.toFixed(RATCHET_PRECISION);

/**
 * LibreOffice reports e.g. `LibreOffice 25.8.7.3 580(Build:3)`. Only major.minor participates in
 * the fingerprint: a patch bump inside one minor is a bugfix release the tolerance absorbs, while
 * a minor bump is the kind of change that redrew the footnote separator between 24.2 and 25.8.
 */
export function libreofficeFingerprint(version: string | undefined): string {
  const match = (version ?? '').match(/(\d+)\.(\d+)/);
  return match ? `${match[1]}.${match[2]}` : 'unknown';
}

/** Chromium reports e.g. `143.0.7499.4`; the major is the meaningful layout boundary. */
export function chromiumFingerprint(version: string | undefined): string {
  const match = (version ?? '').match(/(\d+)/);
  return match ? match[1] : 'unknown';
}

/** pdftoppm reports e.g. `pdftoppm version 24.02.0`; major.minor is the release boundary. */
export function popplerFingerprint(version: string | undefined): string {
  const match = (version ?? '').match(/(\d+)\.(\d+)/);
  return match ? `${match[1]}.${match[2]}` : 'unknown';
}

export function environmentOf(summary: RatchetSummary): RatchetEnvironment {
  return {
    libreoffice: libreofficeFingerprint(summary.environment?.libreoffice),
    chromium: chromiumFingerprint(summary.environment?.chromium),
    poppler: popplerFingerprint(summary.environment?.pdftoppm),
    fontContractSha256: summary.environment?.fontContract?.sha256 ?? 'unknown',
  };
}

/** Mean SSIM over paired pages; 1 for a case with no comparable page (a conversion error). */
export function meanSsim(entry: RatchetSummaryCase): number {
  if (!entry.pages.length) return 0;
  return entry.pages.reduce((sum, page) => sum + page.ssim, 0) / entry.pages.length;
}

/** Worst (minimum) paired-page ink F1 — the aggregation BASELINE.md's table uses. */
export function worstInkF1(entry: RatchetSummaryCase): number {
  if (!entry.pages.length) return 0;
  return entry.pages.reduce((min, page) => Math.min(min, page.tolerantInkF1), Number.POSITIVE_INFINITY);
}

export function recordCase(entry: RatchetSummaryCase): RatchetCase {
  return {
    id: entry.id,
    disposition: entry.disposition.kind,
    pages: { docxodus: entry.docxodusPages, libreoffice: entry.libreofficePages },
    severity: entry.severity,
    ssim: round(meanSsim(entry)),
    worstInkF1: round(worstInkF1(entry)),
    ...(entry.error !== undefined ? { error: true as const } : {}),
  };
}

/**
 * Builds the committed record from a completed run. Numbers only: no image, no path, no hash, so
 * the file stays reviewable as a diff and cannot drift with artifact naming.
 */
export function buildRecord(summary: RatchetSummary, recordedAt: string): RatchetRecord {
  return {
    schemaVersion: RATCHET_SCHEMA_VERSION,
    description:
      'Per-case visual-parity regression ratchet (issue #395). Numbers only. The scheduled run ' +
      'fails when any case gets worse than recorded beyond `tolerance`. Refresh this file ' +
      'deliberately in the PR that changes rendering, so improvements and accepted regressions ' +
      'are reviewed in the diff. See npm/tests/visual-parity/README.md.',
    recordedAt,
    sourceCommit: summary.gitCommit ?? 'unknown',
    environment: environmentOf(summary),
    tolerance: { ssim: RATCHET_TOLERANCE.ssim, inkF1: RATCHET_TOLERANCE.inkF1 },
    cases: [...summary.cases].sort((a, b) => a.id.localeCompare(b.id)).map(recordCase),
  };
}

/** Stable, human-readable serialization; the record is reviewed as a diff. */
export function serializeRecord(record: RatchetRecord): string {
  return `${JSON.stringify(record, null, 2)}\n`;
}

/** The committed record, or null before one exists. The only I/O in this module. */
export function readRecord(file: string = RATCHET_RECORD_FILE): RatchetRecord | null {
  if (!existsSync(file)) return null;
  return JSON.parse(readFileSync(file, 'utf8')) as RatchetRecord;
}

function environmentDrift(record: RatchetRecord, measured: RatchetEnvironment): string[] {
  const drift: string[] = [];
  for (const key of ['libreoffice', 'chromium', 'poppler', 'fontContractSha256'] as const) {
    if (record.environment[key] !== measured[key]) {
      drift.push(`${key}: recorded ${record.environment[key]}, measured ${measured[key]}`);
    }
  }
  return drift;
}

/**
 * Compares a run against the committed record.
 *
 * Ordering matters. An environment change is reported BEFORE any numeric comparison, because
 * comparing numbers measured under a different LibreOffice against numbers measured under this
 * one would attribute a reference-renderer change to Docxodus — the exact false accusation the
 * fingerprint exists to prevent. That outcome still fails the run: a stale record silently
 * comparing across environments is worse than an explicit demand to refresh it.
 */
export function compareToRecord(
  record: RatchetRecord,
  summary: RatchetSummary,
  options: RatchetOptions = { expectComplete: true },
): RatchetComparison {
  if (record.schemaVersion !== RATCHET_SCHEMA_VERSION) {
    return {
      status: 'record-mismatch',
      findings: [],
      improvements: [],
      environmentDrift: [],
      message: `Ratchet record schema ${record.schemaVersion} is not the expected ` +
        `${RATCHET_SCHEMA_VERSION}; regenerate it with DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1.`,
    };
  }

  const measuredEnvironment = environmentOf(summary);
  const drift = environmentDrift(record, measuredEnvironment);
  if (drift.length) {
    return {
      status: 'environment-changed',
      findings: [],
      improvements: [],
      environmentDrift: drift,
      message:
        'Reference environment changed, record refresh required — NOT a renderer regression.\n  ' +
        drift.join('\n  ') +
        '\nThe recorded numbers were measured under a different reference environment, so ' +
        'comparing against them would attribute that change to Docxodus. Rerun the benchmark ' +
        'in the new environment with DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1 and commit the ' +
        'refreshed record.',
    };
  }

  const findings: RatchetFinding[] = [];
  const improvements: RatchetFinding[] = [];
  const measuredById = new Map(summary.cases.map(entry => [entry.id, entry]));

  for (const recorded of record.cases) {
    const measured = measuredById.get(recorded.id);
    if (!measured) {
      if (!options.expectComplete) continue;
      findings.push({
        caseId: recorded.id,
        signal: 'missing',
        recorded: recorded.severity,
        measured: '(absent)',
        detail: `${recorded.id} is in the record but was not measured by this run`,
      });
      continue;
    }
    findings.push(...compareCase(recorded, measured, record.tolerance, improvements));
  }

  for (const measured of summary.cases) {
    if (record.cases.some(recorded => recorded.id === measured.id)) continue;
    findings.push({
      caseId: measured.id,
      signal: 'unrecorded',
      recorded: '(absent)',
      measured: measured.severity,
      detail: `${measured.id} was measured but is not in the record; add it with ` +
        'DOCXODUS_VISUAL_PARITY_UPDATE_RECORD=1',
    });
  }

  if (findings.length) {
    return {
      status: 'regressed',
      findings,
      improvements,
      environmentDrift: [],
      message: `Visual parity regressed against the record in ${findings.length} signal(s):\n  ` +
        findings.map(finding => finding.detail).join('\n  '),
    };
  }

  const improved = improvements.length
    ? ` ${improvements.length} signal(s) improved; refresh the record to bank them:\n  ` +
      improvements.map(finding => finding.detail).join('\n  ')
    : '';
  const compared = record.cases.filter(recorded => measuredById.has(recorded.id)).length;
  return {
    status: 'ok',
    findings: [],
    improvements,
    environmentDrift: [],
    message: `Visual parity holds against the record for ${compared} case(s).${improved}`,
  };
}

function compareCase(
  recorded: RatchetCase,
  measured: RatchetSummaryCase,
  tolerance: { ssim: number; inkF1: number },
  improvements: RatchetFinding[],
): RatchetFinding[] {
  const findings: RatchetFinding[] = [];

  if (measured.error !== undefined && !recorded.error) {
    findings.push({
      caseId: recorded.id,
      signal: 'error',
      recorded: 'converted',
      measured: 'conversion error',
      detail: `${recorded.id}: conversion error where the record has none — ${measured.error}`,
    });
    // Every other signal is meaningless once the case failed to render at all.
    return findings;
  }

  if (measured.docxodusPages !== recorded.pages.docxodus ||
      measured.libreofficePages !== recorded.pages.libreoffice) {
    findings.push({
      caseId: recorded.id,
      signal: 'pages',
      recorded: `${recorded.pages.docxodus}/${recorded.pages.libreoffice}`,
      measured: `${measured.docxodusPages}/${measured.libreofficePages}`,
      detail: `${recorded.id}: page count changed from ` +
        `${recorded.pages.docxodus}/${recorded.pages.libreoffice} (Docxodus/LibreOffice) to ` +
        `${measured.docxodusPages}/${measured.libreofficePages}`,
    });
  }

  const severityDelta = SEVERITY_ORDER.indexOf(measured.severity) -
    SEVERITY_ORDER.indexOf(recorded.severity);
  if (severityDelta > 0) {
    findings.push({
      caseId: recorded.id,
      signal: 'severity',
      recorded: recorded.severity,
      measured: measured.severity,
      detail: `${recorded.id}: severity worsened from ${recorded.severity} to ${measured.severity}`,
    });
  } else if (severityDelta < 0) {
    improvements.push({
      caseId: recorded.id,
      signal: 'severity',
      recorded: recorded.severity,
      measured: measured.severity,
      detail: `${recorded.id}: severity improved from ${recorded.severity} to ${measured.severity}`,
    });
  }

  const signals = [
    { signal: 'ssim' as const, recorded: recorded.ssim, measured: meanSsim(measured), tolerance: tolerance.ssim, label: 'mean SSIM' },
    { signal: 'inkF1' as const, recorded: recorded.worstInkF1, measured: worstInkF1(measured), tolerance: tolerance.inkF1, label: 'worst ink F1' },
  ];
  for (const entry of signals) {
    const delta = entry.measured - entry.recorded;
    if (delta < -entry.tolerance) {
      findings.push({
        caseId: recorded.id,
        signal: entry.signal,
        recorded: format(entry.recorded),
        measured: format(entry.measured),
        detail: `${recorded.id}: ${entry.label} fell from ${format(entry.recorded)} to ` +
          `${format(entry.measured)} (${format(delta)}, tolerance ${entry.tolerance})`,
      });
    } else if (delta > entry.tolerance) {
      improvements.push({
        caseId: recorded.id,
        signal: entry.signal,
        recorded: format(entry.recorded),
        measured: format(entry.measured),
        detail: `${recorded.id}: ${entry.label} rose from ${format(entry.recorded)} to ` +
          `${format(entry.measured)} (+${format(delta)})`,
      });
    }
  }

  return findings;
}

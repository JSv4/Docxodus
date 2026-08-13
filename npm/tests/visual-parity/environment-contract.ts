import { spawnSync } from 'node:child_process';
import { libreofficeFingerprint, popplerFingerprint } from './ratchet.js';

/** Version banner of a command, tolerant of tools that print it on stderr (pdftoppm does). */
function versionBanner(command: string, args: string[]): string {
  const result = spawnSync(command, args, { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'] });
  if (result.error) return '';
  return `${result.stdout ?? ''}\n${result.stderr ?? ''}`.trim();
}

/**
 * The reference-renderer version contract (issue #403).
 *
 * With fonts pinned by the issue-#379 contract, the LibreOffice version was the last
 * uncontracted variable in the comparison: 24.2 and 25.8 render the same document differently
 * (24.2 draws its legacy 25%-of-column footnote separator where 25.8 draws the two-inch
 * default), so a runner-image bump could silently shift weekly numbers and masquerade as a
 * renderer change. This module declares the version the benchmark's numbers mean anything
 * against, and asserts it BEFORE any rendering starts — mirroring the font contract's failure
 * mode: fail fast, name what to install.
 *
 * The declaration is exact at major.minor, not a minimum. The ratchet record is measured under
 * one reference version; a "minimum" would let a newer LibreOffice run for twenty minutes and
 * then trip the ratchet's environment-changed check anyway. Exact fails in the first second
 * with instructions. Upgrading the reference is a deliberate act: bump `version` here, rerun
 * the corpus, refresh `ratchet.json`, and review all three in one diff.
 */
export const LIBREOFFICE_CONTRACT = {
  /** Exact major.minor the benchmark is contracted against (the ratchet fingerprint boundary). */
  version: '25.8',
  /** A known-good full build of that minor, used by CI and the install guidance. */
  build: '25.8.7.3',
  /** Deterministic download for hosts whose distro does not package the contract minor. */
  archiveUrl:
    'https://downloadarchive.documentfoundation.org/libreoffice/old/25.8.7.3/deb/x86_64/' +
    'LibreOffice_25.8.7.3_Linux_x86-64_deb.tar.gz',
} as const;

/**
 * Cross-version rendering differences the corpus is known to be sensitive to. Data, not prose,
 * so the failure message can cite them and the README table is generated from one source.
 */
export const KNOWN_LIBREOFFICE_VERSION_DIFFERENCES = [
  {
    versions: '24.2 vs ≥ 25.8',
    behavior: 'Footnote separator width: 24.2 draws its legacy 25%-of-column separator where ' +
      '25.8 draws the two-inch Word default.',
    sensitiveCases: ['footnote'],
    discovered: 'issue #378 verification (BASELINE.md, 2026-08-11)',
  },
] as const;

/**
 * Escape hatch for deliberate cross-version exploration (e.g. reproducing a case on a second
 * environment to corroborate an `environment` disposition, as the issue-#378 verification did).
 * The run proceeds, the summary records the real version, and the ratchet still reports
 * `environment-changed` against the committed record — the escape hatch skips the fast failure,
 * never the fingerprint guard.
 */
export const LIBREOFFICE_DRIFT_ENV = 'DOCXODUS_VISUAL_PARITY_ALLOW_VERSION_DRIFT';

export interface LibreOfficeContractCheck {
  ok: boolean;
  /** Full `libreoffice --version` output, recorded in `summary.json` either way. */
  version: string;
  /** major.minor actually present, `unknown` when unparsable. */
  fingerprint: string;
  /** Failure guidance when out of contract; empty when in contract. */
  message: string;
}

/**
 * Pure contract decision so the ratchet spec can prove the failure mode on every pull request
 * without a LibreOffice install — the same device that keeps the ratchet alarm continuously
 * verified instead of demonstrated once.
 */
export function checkLibreOfficeContract(versionOutput: string): LibreOfficeContractCheck {
  const fingerprint = libreofficeFingerprint(versionOutput);
  if (fingerprint === LIBREOFFICE_CONTRACT.version) {
    return { ok: true, version: versionOutput, fingerprint, message: '' };
  }
  const differences = KNOWN_LIBREOFFICE_VERSION_DIFFERENCES
    .map(entry => `  - ${entry.versions}: ${entry.behavior} (sensitive: ${entry.sensitiveCases.join(', ')})`)
    .join('\n');
  return {
    ok: false,
    version: versionOutput,
    fingerprint,
    message:
      `LibreOffice reference-version contract not satisfied: found ${fingerprint} ` +
      `(${versionOutput.trim() || 'no version output'}), contract requires ${LIBREOFFICE_CONTRACT.version}.\n` +
      `The benchmark's committed record (ratchet.json) is only meaningful against ` +
      `LibreOffice ${LIBREOFFICE_CONTRACT.version}; known cross-version differences:\n${differences}\n` +
      `Install LibreOffice ${LIBREOFFICE_CONTRACT.build}, e.g. the TDF build:\n` +
      `  ${LIBREOFFICE_CONTRACT.archiveUrl}\n` +
      `After a TDF install, REMOVE its bundled Carlito/Caladea/Liberation fonts ` +
      `(share/fonts/truetype/) — they silently override the font-substitution contract inside ` +
      `LibreOffice only (see README).\n` +
      `To deliberately measure a different version anyway (exploratory, non-record runs), set ` +
      `${LIBREOFFICE_DRIFT_ENV}=1; the ratchet will still report environment-changed rather ` +
      `than comparing across versions.`,
  };
}

/** `libreoffice --version` from the host, empty string when the binary is missing entirely. */
export function libreofficeVersionOutput(): string {
  return versionBanner('libreoffice', ['--version']);
}

/**
 * Run-start assertion, mirroring `assertFontContract()`: throws with install guidance when the
 * host's LibreOffice is out of contract, returns the full version string for the report when in
 * contract (or when drift is deliberately allowed).
 */
export function assertLibreOfficeContract(): string {
  const check = checkLibreOfficeContract(libreofficeVersionOutput());
  if (check.ok) return check.version;
  if (process.env[LIBREOFFICE_DRIFT_ENV] === '1') {
    console.warn(`[visual-parity] ${LIBREOFFICE_DRIFT_ENV}=1: measuring OUT-OF-CONTRACT ` +
      `LibreOffice ${check.fingerprint} (contract ${LIBREOFFICE_CONTRACT.version}); the ratchet ` +
      'will report environment-changed and this run cannot refresh the record.');
    return check.version;
  }
  throw new Error(check.message);
}

/** Poppler is fingerprinted (not install-asserted): pdftoppm's rasterizer is part of what the
 * recorded numbers were measured through, so a distro bump must surface as environment drift
 * rather than as a silent shift attributed to the renderer. */
export function popplerVersionOutput(): string {
  // pdftoppm prints its version banner on stderr with exit code 0.
  return versionBanner('pdftoppm', ['-v']).split('\n')[0].trim();
}

export { popplerFingerprint };

import { spawnSync } from 'node:child_process';
import { mkdirSync, writeFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

/**
 * The font-substitution contract shared by both renderers.
 *
 * Neither Chromium nor LibreOffice ships Microsoft's Office fonts, so every Office family a
 * fixture names is SUBSTITUTED. When the two engines substitute differently, their line breaking
 * and baselines differ for reasons that have nothing to do with Docxodus — and a benchmark that
 * cannot tell those apart from renderer regressions is not measuring the renderer.
 *
 * The fix has to be shared, not renderer-side. A Chromium-only `Calibri Light -> Carlito`
 * fallback was tried and rejected: it made ONE engine change its mind, so the two disagreed more
 * than before (accepted-revision SSIM 0.93177 -> 0.92905, ink F1 0.46817 -> 0.42477). Both
 * engines read fontconfig on Linux, so ONE fontconfig fragment governs both, by construction.
 *
 * Nothing here bundles a font. Carlito, Caladea, and the Liberation family are the metric-
 * compatible, freely redistributable substitutes the distributions already package; the contract
 * only pins WHICH of them each Office family must resolve to.
 */

const __dirname = dirname(fileURLToPath(import.meta.url));

export interface FontSubstitution {
  /** The family a document names. */
  family: string;
  /** The family both renderers must resolve it to. */
  substitute: string;
  /** Where the substitute comes from, for the setup instructions and the failure message. */
  package: string;
  /**
   * Whether the substitute is metrically compatible with the original. `false` means no
   * license-safe metric clone exists, so the contract's job is agreement between the engines
   * rather than fidelity to Word — worth stating, because such a family's wrap points are
   * expected to differ from Word's and must not be "fixed" in one renderer.
   */
  metricCompatible: boolean;
}

export const FONT_SUBSTITUTION_CONTRACT: readonly FontSubstitution[] = [
  { family: 'Calibri', substitute: 'Carlito', package: 'fonts-crosextra-carlito', metricCompatible: true },
  { family: 'Cambria', substitute: 'Caladea', package: 'fonts-crosextra-caladea', metricCompatible: true },
  { family: 'Times New Roman', substitute: 'Liberation Serif', package: 'fonts-liberation2', metricCompatible: true },
  { family: 'Arial', substitute: 'Liberation Sans', package: 'fonts-liberation2', metricCompatible: true },
  { family: 'Courier New', substitute: 'Liberation Mono', package: 'fonts-liberation2', metricCompatible: true },
  // Carlito has no Light weight, and no freely redistributable font is metrically compatible with
  // Calibri Light. Left unbound it falls through to whatever generic each engine prefers — here
  // Chromium and LibreOffice both landed on Noto Sans, but that is an accident of this host, not
  // a contract. Binding it makes the choice explicit and identical in both engines.
  { family: 'Calibri Light', substitute: 'Carlito', package: 'fonts-crosextra-carlito', metricCompatible: false },
];

/** The apt packages a Debian/Ubuntu host needs for the whole contract. */
export const FONT_CONTRACT_PACKAGES = [...new Set(
  FONT_SUBSTITUTION_CONTRACT.map(entry => entry.package))].sort();

/** The shipped fontconfig fragment that binds the contract for every application on the host. */
export const FONTCONFIG_FRAGMENT = resolve(
  __dirname, 'fontconfig', '60-docxodus-office-substitutes.conf');

export interface ResolvedFont extends FontSubstitution {
  /** The family `fc-match` actually returned. */
  resolvedFamily: string;
  /** The font file backing it, so a report records the exact bytes that were rendered. */
  resolvedFile: string;
  satisfied: boolean;
}

export interface FontContractStatus {
  satisfied: boolean;
  entries: ResolvedFont[];
  /** Human-readable reason, empty when satisfied. */
  problem: string;
  /** The fontconfig configuration in force, so a report can be reproduced. */
  fontconfigFile: string;
}

function fcMatch(family: string, env: NodeJS.ProcessEnv): { family: string; file: string } {
  const result = spawnSync('fc-match', [family, '--format=%{family}\\t%{file}'],
    { encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], env });
  if (result.error || result.status !== 0) return { family: '', file: '' };
  const [matched = '', file = ''] = result.stdout.trim().split('\t');
  return { family: matched, file };
}

/**
 * A fontconfig root that layers {@link FONTCONFIG_FRAGMENT} over the host's own configuration.
 *
 * Written outside the repository and pointed at through `FONTCONFIG_FILE`, so the contract binds
 * the benchmark's two subprocesses WITHOUT installing anything into the developer's home
 * directory or `/etc`. CI can install the fragment permanently instead; the effect is the same.
 */
export function writeFontconfigRoot(directory: string): string {
  mkdirSync(directory, { recursive: true });
  const path = join(directory, 'fonts.conf');
  writeFileSync(path, `<?xml version="1.0"?>
<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">
<!-- Generated by npm/tests/visual-parity/fonts.ts. Host configuration first, contract last. -->
<fontconfig>
  <include ignore_missing="yes">/etc/fonts/fonts.conf</include>
  <include ignore_missing="yes">${FONTCONFIG_FRAGMENT}</include>
</fontconfig>
`);
  return path;
}

/** Resolves every declared family through `fc-match` under `env` and reports the contract state. */
export function resolveFontContract(env: NodeJS.ProcessEnv = process.env): FontContractStatus {
  const entries = FONT_SUBSTITUTION_CONTRACT.map<ResolvedFont>(entry => {
    const matched = fcMatch(entry.family, env);
    return {
      ...entry,
      resolvedFamily: matched.family,
      resolvedFile: matched.file,
      satisfied: matched.family === entry.substitute,
    };
  });

  const broken = entries.filter(entry => !entry.satisfied);
  const missingPackages = [...new Set(broken.map(entry => entry.package))].sort();
  return {
    satisfied: broken.length === 0,
    entries,
    problem: broken.length === 0 ? '' :
      `font substitution contract unsatisfied: ` +
      broken.map(e => `${e.family} -> ${e.resolvedFamily || '(no match)'} (want ${e.substitute})`).join('; ') +
      `. Install ${missingPackages.join(' ')} and apply ${FONTCONFIG_FRAGMENT} ` +
      `(see npm/tests/visual-parity/README.md).`,
    fontconfigFile: env.FONTCONFIG_FILE ?? '(host default)',
  };
}

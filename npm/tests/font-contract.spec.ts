import { test, expect } from '@playwright/test';
import { spawnSync } from 'node:child_process';
import { readFileSync } from 'node:fs';
import {
  FONT_CONTRACT_PACKAGES,
  FONT_SUBSTITUTION_CONTRACT,
  FONTCONFIG_FRAGMENT,
  resolveFontContract,
} from './visual-parity/fonts.js';
import {
  compareProbeLines,
  PROBE_ADVANCE_TOLERANCE_PX,
  PROBE_FAMILIES,
  type InkLine,
} from './visual-parity/font-probe.js';

/**
 * Issue #379 — the font-substitution contract itself, checked without needing LibreOffice.
 *
 * The visual-parity benchmark is opt-in and host-dependent; this spec is not. It pins the two
 * things that would silently rot between benchmark runs: the shipped fontconfig fragment agreeing
 * with the TypeScript declaration of the same contract, and the drift detector actually
 * detecting drift.
 */

/** One ink line per entry, `right - left` wide, at an arbitrary but consistent y. */
function lines(widths: number[]): InkLine[] {
  return widths.map((width, index) => ({
    top: index * 30,
    bottom: index * 30 + 12,
    left: 96,
    right: 96 + width,
  }));
}

/** The advance lines the probe compares, followed by some wrapped ones it only counts. */
function probePage(advances: number[], wrapped = [600, 600, 400]): InkLine[] {
  return lines([...advances, ...wrapped]);
}

const BASELINE_ADVANCES = [210, 213, 208, 226, 304, 210];

test.describe('Font substitution contract', () => {
  /** Every family→substitute binding the fragment actually declares, in either fontconfig form. */
  function fragmentBindings(fragment: string): Map<string, string> {
    const bindings = new Map<string, string>();
    const aliases = fragment.matchAll(
      /<alias[^>]*>\s*<family>([^<]+)<\/family>\s*<accept>\s*<family>([^<]+)<\/family>/g);
    for (const [, family, substitute] of aliases) bindings.set(family, substitute);
    const matches = fragment.matchAll(
      /<test[^>]*name="family"[^>]*>\s*<string>([^<]+)<\/string>[\s\S]*?<edit name="family"[^>]*>\s*<string>([^<]+)<\/string>/g);
    for (const [, family, substitute] of matches) bindings.set(family, substitute);
    return bindings;
  }

  test('the shipped fontconfig fragment binds exactly the declared families', () => {
    // Both engines read the fragment; the TypeScript list is only how the run REPORTS the
    // contract. If they drift apart, every report describes a policy the renderers do not follow.
    const bindings = fragmentBindings(readFileSync(FONTCONFIG_FRAGMENT, 'utf8'));
    const declared = new Map(FONT_SUBSTITUTION_CONTRACT.map(e => [e.family, e.substitute]));

    expect(Object.fromEntries([...bindings].sort()))
      .toEqual(Object.fromEntries([...declared].sort()));
    expect(FONT_CONTRACT_PACKAGES.length).toBeGreaterThan(0);
  });

  test('the drift detector flags a substituted face', () => {
    // Calibri Light resolved to Noto Sans in one engine and Carlito in the other: a 38 px
    // difference on the probe's short line, which is what this tolerance exists to catch.
    const drifted = [...BASELINE_ADVANCES];
    drifted[5] += 38;

    const result = compareProbeLines(probePage(BASELINE_ADVANCES), probePage(drifted));
    expect(result.agreed).toBe(false);
    expect(result.problem).toContain(PROBE_FAMILIES[5]);
    expect(result.maxAdvanceDeltaPx).toBe(38);
  });

  test('the drift detector flags a different wrapped line count', () => {
    const result = compareProbeLines(
      probePage(BASELINE_ADVANCES, [600, 600, 400]),
      probePage(BASELINE_ADVANCES, [600, 600, 600, 200]),
    );
    expect(result.agreed).toBe(false);
    expect(result.problem).toContain('different line counts');
  });

  test('the drift detector tolerates rasterisation noise', () => {
    // Hinting moves a short line's end by a pixel or two between engines; that is not drift.
    const noisy = BASELINE_ADVANCES.map((width, index) =>
      width + (index % 2 === 0 ? PROBE_ADVANCE_TOLERANCE_PX : -PROBE_ADVANCE_TOLERANCE_PX));

    const result = compareProbeLines(probePage(BASELINE_ADVANCES), probePage(noisy));
    expect(result.agreed, result.problem).toBe(true);
    expect(result.maxAdvanceDeltaPx).toBe(PROBE_ADVANCE_TOLERANCE_PX);
    expect(result.advances.map(a => a.family)).toEqual(PROBE_FAMILIES);
  });

  test('an unsatisfied contract reports what to install', () => {
    test.skip(spawnSync('fc-match', ['--version']).error !== undefined,
      'fontconfig is not installed on this host');

    // Resolve against the HOST configuration, deliberately without the repository fragment. The
    // result is whatever this machine happens to do — the assertion is about the report, not the
    // machine: either the contract holds, or it says exactly which family and package are wrong.
    const status = resolveFontContract({ ...process.env, FONTCONFIG_FILE: undefined });
    expect(status.entries.map(entry => entry.family))
      .toEqual(FONT_SUBSTITUTION_CONTRACT.map(entry => entry.family));

    if (status.satisfied) {
      expect(status.problem).toBe('');
      for (const entry of status.entries) expect(entry.resolvedFile).not.toBe('');
    } else {
      const broken = status.entries.filter(entry => !entry.satisfied);
      expect(broken.length).toBeGreaterThan(0);
      for (const entry of broken) {
        expect(status.problem).toContain(entry.family);
        expect(status.problem).toContain(entry.package);
      }
      expect(status.problem).toContain('README.md');
    }
  });
});

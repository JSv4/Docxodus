import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { canonicalJson } from './visual-parity/canonical-json.js';

/**
 * `canonical-json.ts` is a copy of the exporter's `npm-export/src/canonical.ts`. The benchmark
 * cannot import across package roots, but a copy that silently drifts is worse than no copy at
 * all: `pdf-result.ts` compares the exporter's `bindings.pageMapDigest` against a digest IT
 * computes, so a divergence would read as an exporter defect rather than as harness drift.
 *
 * Comparing the two implementations on values that actually distinguish canonicalizations — key
 * order, dropped `undefined`, negative zero, non-BMP text — is what makes the copy a checked
 * mirror instead of a fork waiting to happen.
 */
const __dirname = dirname(fileURLToPath(import.meta.url));
const EXPORTER_SOURCE = resolve(__dirname, '../../npm-export/src/canonical.ts');

/** The exporter module is TypeScript-free apart from its annotations; strip them and its Node
 *  import so the same source can be evaluated here without a build step. */
function exporterCanonicalJson(): (value: unknown) => string {
  const source = readFileSync(EXPORTER_SOURCE, 'utf8');
  const body = source
    .slice(source.indexOf('function assertWellFormedUnicode'))
    .replace(/export function canonicalJsonBytes[\s\S]*$/, '')
    .replace(/export function sha256[\s\S]*$/, '')
    .replace(/: string(?=\))/g, '')
    .replace(/: unknown(?=[),])/g, '')
    .replace(/: void(?= \{)/g, '')
    .replace(/\)\: (string|unknown) \{/g, ') {')
    .replace(/ as Record<string, unknown>/g, '')
    .replace(/const result: Record<string, unknown> =/, 'const result =')
    .replace(/export function/g, 'function');
  // eslint-disable-next-line no-new-func
  return new Function(`${body}\nreturn canonicalJson;`)() as (value: unknown) => string;
}

test.describe('canonical JSON mirrors the exporter', () => {
  test('agrees with the exporter on values that distinguish canonicalizations', () => {
    const reference = exporterCanonicalJson();
    const cases: unknown[] = [
      { b: 1, a: 2 },
      { nested: { z: [3, { y: 1, x: 2 }], a: null } },
      { dropped: undefined, kept: false },
      { minusZero: -0, plusZero: 0 },
      { unicode: 'é \u{1F600}' },
      { '': 'empty key', ' ': 'space key' },
      [1, [2, [3]]],
      [],
      {},
    ];
    for (const value of cases) {
      expect(canonicalJson(value), JSON.stringify(value) ?? 'undefined')
        .toBe(reference(value));
    }
    // A sanity check on the harness itself: the reference really is canonicalizing, not just
    // calling JSON.stringify, or every assertion above would be trivially true.
    expect(reference({ b: 1, a: 2 })).toBe('{"a":2,"b":1}');
    expect(JSON.stringify({ b: 1, a: 2 })).toBe('{"b":1,"a":2}');
  });

  test('rejects what the exporter rejects', () => {
    expect(() => canonicalJson({ n: Number.NaN })).toThrow(/non-finite/);
    expect(() => canonicalJson({ s: '\ud800' })).toThrow(/surrogate/);
    expect(() => canonicalJson({ d: new Date(0) })).toThrow(/plain objects/);
    expect(() => canonicalJson(() => undefined)).toThrow(/function/);
  });
});

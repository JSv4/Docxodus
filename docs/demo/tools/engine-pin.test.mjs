// The demo pages load the library from jsDelivr at an exact version, and that
// pin is the whole reason the arcade's DOOM cartridge broke in public while
// every Playwright spec stayed green: the specs override `?engine=` to the
// locally built bundle, so the pin is the ONE thing under test here that no
// browser test can see. Up to and including 10.0.0 the single-block renderer
// refreshed an image-bearing paragraph blank, so a site pinned there boots
// Doom, plays it, writes real frames into the .docx — and shows nothing.
//
// So this test asks three offline questions of the checked-in files:
//
//   1. Do all the pages agree on one version? A half-moved pin is how you get
//      a landing page on one engine and the arcade it frames on another.
//   2. Is that version at or above IMAGE_ENGINE_MINIMUM — the oldest engine
//      whose incremental render carries an inline image?
//   3. Does npm/tests/social-demo.spec.ts's RELEASE_ENGINE — the guard that
//      proves the deployed pages load the pin rather than a 404 — still name
//      the same version the pages carry?
//
// The fourth question, whether that version actually exists on the CDN, needs
// the network: it runs under DOCXODUS_CHECK_CDN=1 and is skipped otherwise.
import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync, readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';

import { IMAGE_ENGINE_MINIMUM } from '../ascii-arcade.js';

const demoDir = new URL('../', import.meta.url);
const repoRoot = new URL('../../../', import.meta.url);

/** Every file that can carry a pin: the demo pages plus the docs that quote them. */
const PINNED_FILES = [
  ...readdirSync(fileURLToPath(demoDir))
    .filter((name) => name.endsWith('.html') || name === 'README.md')
    .map((name) => `docs/demo/${name}`),
  'docs/npm-package.md',
  'npm/README.md',
  'npm/examples/embed.html',
  // Copy-pasteable CDN examples shipped in the package's own doc comments: a
  // reader who pastes one gets whatever engine they name, so they move too.
  'npm/src/embed.ts',
  'npm/src/index.ts',
];

// `docxodus@12.0.0`, `docxodus@12.0.0/dist/embed.bundle.js`, the bare npm URL —
// all of them, but never `docxodus@latest`, which the docs use deliberately in
// one "unpinned" example.
const PIN = /docxodus@(\d+\.\d+\.\d+)/g;

// Prose that recounts a version is written WITHOUT the `docxodus@` prefix, so
// this scan never sees it. The one grandfathered exception is the README's
// account of the observatory being pinned ahead of the 9.6.0 release, which is
// deliberately preserved wording.
const HISTORICAL = [/pinned to `docxodus@[\d.]+` ahead of that release/];

const read = (relative) => readFileSync(new URL(relative, repoRoot), 'utf8');

/** Semver compare restricted to the release shape these pins use. */
function compare(a, b) {
  const pa = a.split('.').map(Number);
  const pb = b.split('.').map(Number);
  for (let i = 0; i < 3; i++) {
    if (pa[i] !== pb[i]) return pa[i] - pb[i];
  }
  return 0;
}

/** Every (file, version) pin in the demo/doc set, in file order. */
function pins() {
  const found = [];
  for (const file of PINNED_FILES) {
    const text = read(file);
    for (const match of text.matchAll(PIN)) {
      if (HISTORICAL.some((h) => h.test(text.slice(Math.max(0, match.index - 60),
        match.index + 120)))) continue;
      found.push({ file, version: match[1] });
    }
  }
  return found;
}

test('every demo page and doc pins the same engine version', () => {
  const found = pins();
  assert.ok(found.length > 0, 'no docxodus@X.Y.Z pins found — did the demo pages move?');
  const versions = [...new Set(found.map((p) => p.version))];
  assert.deepEqual(
    versions, [found[0].version],
    'demo pins disagree: ' + found
      .filter((p) => p.version !== found[0].version)
      .map((p) => `${p.file} → ${p.version}`).join(', '));
});

test('the pinned engine can render an image-bearing cartridge', () => {
  const [{ version }] = pins();
  assert.ok(
    compare(version, IMAGE_ENGINE_MINIMUM) >= 0,
    `the demos pin docxodus@${version}, but the arcade's DOOM cartridge needs `
    + `${IMAGE_ENGINE_MINIMUM} or newer to put its inline image on screen — on an `
    + 'older engine the frame lands in the .docx and never in the page.');
});

test("the release-pin spec guard names the pages' version", () => {
  const spec = read('npm/tests/social-demo.spec.ts');
  const declared = /const RELEASE_ENGINE = 'docxodus@(\d+\.\d+\.\d+)\//.exec(spec);
  assert.ok(declared, 'RELEASE_ENGINE is no longer a pinned docxodus@X.Y.Z literal');
  assert.equal(declared[1], pins()[0].version,
    'npm/tests/social-demo.spec.ts RELEASE_ENGINE drifted from the demo pages — '
    + 'that spec is what proves the deployed pages load the pin rather than a 404.');
});

// Isolated integration check: the pin has to be a version jsDelivr actually
// serves, which is the failure mode a re-pin ahead of a publish produces (the
// pages 404 on the bundle and boot nothing). Network, so opt-in.
test('the pinned bundle is published on the CDN', {
  skip: process.env.DOCXODUS_CHECK_CDN === '1'
    ? false
    : 'set DOCXODUS_CHECK_CDN=1 to check jsDelivr',
}, async () => {
  const { version } = pins()[0];
  const url = `https://cdn.jsdelivr.net/npm/docxodus@${version}/dist/embed.bundle.js`;
  const response = await fetch(url, { method: 'GET', headers: { range: 'bytes=0-0' } });
  assert.ok(response.ok, `${url} returned ${response.status}`);
  assert.match(response.headers.get('content-type') ?? '', /javascript/);
});

// The canvas is a character grid, and the grid only holds if every cell
// advances the same width. docs/demo/fonts/docxodus-canvas-mono.woff2 is what
// guarantees that (one advance for all 538 codepoints in it, on every device),
// and createCanvasPin pins the canvas paragraph to it — so a character the
// subset does NOT cover falls through to whatever the platform has, which is
// exactly the failure the font exists to prevent: on Android the block
// elements land in a proportional fallback and every cell after one is
// displaced, differently on each row, tilting the art.
//
// So this test asks the only question that keeps that true as the demos grow:
// is every character the scenes, the attract screen and the cartridges can
// draw inside the shipped subset? It answers it by DRIVING them — all four
// phenomena over a long timeline, the whole title-card sweep, and all three
// cartridges under input scripts that reach the banners, deaths and win
// screens — rather than by scanning the source for literals.
//
// The Doom cartridge is driven differently from the other two, because its
// picture comes from a WebAssembly Doom that will not boot in a Node test: its
// framebuffer painter is exported and fed synthetic frames instead, which is
// enough because the glyph a picture cell can hold is decided entirely by
// whether its two half-pixels match. Its chrome, loading and error screens
// come from render() as usual.
//
// If this fails, either widen the subset (edit UNICODES in
// tools/build-canvas-font.sh and rebuild) or pick a character already in it.
import assert from 'node:assert/strict';
import test from 'node:test';
import { readFileSync } from 'node:fs';

import { SCENES } from '../ascii-scenes.js';
import { platformerCart, dungeonCart, introFrame } from '../ascii-arcade.js';
import { doomCart, paintFramebuffer, paintFramebuffer8Bit } from '../doom-cart.js';

const manifest = JSON.parse(
  readFileSync(new URL('../fonts/docxodus-canvas-mono.json', import.meta.url), 'utf8'));

const covers = (cp) => manifest.ranges.some(([lo, hi]) => cp >= lo && cp <= hi);
const hex = (cp) => 'U+' + cp.toString(16).toUpperCase().padStart(4, '0');

/** A deterministic stand-in for the browser input the cartridges read. */
class ScriptedInput {
  constructor(script) { this.script = script; this.down = new Set(); this.edges = new Set(); }
  step(i) { this.down = new Set(this.script(i)); this.edges = new Set(this.down); }
  held(...codes) { return codes.some((c) => this.down.has(c)); }
  took(code) { const hit = this.edges.has(code); this.edges.delete(code); return hit; }
  endTick() { this.edges.clear(); }
}

// Run, jump, turn, shoot, die, restart — enough traffic through each cartridge
// to surface its HUD, its banners and its end states.
const SCRIPTS = [
  (i) => (i % 40 < 26 ? ['KeyD', 'ArrowRight'] : ['KeyW', 'Space', 'ArrowUp']),
  (i) => (i % 30 < 20 ? ['KeyW', 'ArrowUp', 'KeyD'] : ['KeyA', 'ArrowLeft']),
  (i) => (i % 17 < 9 ? ['KeyD', 'ArrowRight'] : i % 17 < 13 ? ['Space'] : ['KeyR']),
  (i) => (i % 23 < 12 ? ['ArrowLeft', 'KeyA'] : ['ArrowRight', 'Space', 'KeyE']),
];

/** Every distinct character the canvas can hold, mapped to where it came from. */
function drawnCharacters() {
  const seen = new Map();
  const collect = (grid, source) => {
    for (const row of grid.chars) for (const ch of row) if (!seen.has(ch)) seen.set(ch, source);
  };

  for (const scene of SCENES) {
    scene.reset?.();
    for (let i = 0; i < 400; i++) collect(scene.gen(i * 0.08), `scene:${scene.name}`);
  }
  for (let i = 0; i < 900; i++) collect(introFrame(i * 0.02).grid, 'attract screen');
  for (const make of [platformerCart, dungeonCart]) {
    for (const script of SCRIPTS) {
      const cart = make();
      const input = new ScriptedInput(script);
      for (let i = 0; i < 700; i++) {
        input.step(i);
        cart.tick(0.05, input);
        input.endTick();
        collect(cart.render().grid, `cart:${cart.name}`);
      }
    }
  }

  // Doom: the chrome and the pre-game screens, straight from render().
  const doom = doomCart({ engineUrl: 'about:blank', wadUrl: 'about:blank' });
  for (let i = 0; i < 40; i++) collect(doom.render().grid, 'cart:doom');

  // Doom: the picture itself. A framebuffer whose halves sometimes match and
  // sometimes do not exercises both cell glyphs; the gradient makes sure the
  // painter is not accidentally emitting anything else.
  const fb = new Uint8Array(320 * 200 * 4);
  for (let y = 0; y < 200; y++) {
    for (let x = 0; x < 320; x++) {
      const i = (y * 320 + x) * 4;
      fb[i] = (x * 7) & 0xff;                       // B
      fb[i + 1] = (y * 5) & 0xff;                   // G
      fb[i + 2] = y < 100 ? 0x40 : (x ^ y) & 0xff;  // R — flat above, noisy below
    }
  }
  const doomGrid = doom.render().grid;
  paintFramebuffer(doomGrid, fb);
  collect(doomGrid, 'cart:doom framebuffer');

  // The playable projection can emit all sixteen quadrant patterns, so it
  // needs its own pass — the pinned subset has to cover every glyph either
  // projection can emit.
  const eightBitGrid = doom.render().grid;
  paintFramebuffer8Bit(eightBitGrid, fb);
  collect(eightBitGrid, 'cart:doom 8-bit projection');

  // A second frame with a hard horizontal split in every cell, to force the
  // edge glyphs the gradient above may never trigger.
  const split = new Uint8Array(320 * 200 * 4);
  for (let y = 0; y < 200; y++) {
    for (let x = 0; x < 320; x++) {
      const i = (y * 320 + x) * 4;
      const bright = (y % 9) < 4;
      split[i] = bright ? 0xF0 : 0x08;
      split[i + 1] = bright ? 0x30 : 0x08;
      split[i + 2] = bright ? 0xC0 : 0x08;
    }
  }
  const edgeGrid = doom.render().grid;
  paintFramebuffer8Bit(edgeGrid, split);
  collect(edgeGrid, 'cart:doom 8-bit edges');

  return seen;
}

test('the bitmap run budget preserves near-colour texture through free glyphs', () => {
  // This is the failure mode from the shipped GIF in executable form. Every
  // adjacent sample differs by only 30 grey levels: the old 13% pair-snap
  // tolerance swallowed the entire row into one flat bar. A colour run does
  // not break on glyphs, so the budgeted painter can keep one ink/bg pair and
  // alternate solid/empty cells to retain nearly every boundary for free.
  const fb = new Uint8Array(320 * 200 * 4);
  for (let y = 0; y < 200; y++) {
    for (let x = 0; x < 320; x++) {
      const projectedX = Math.min(93, Math.floor(x * 94 / 320));
      const grey = projectedX % 2 ? 110 : 80;
      const i = (y * 320 + x) * 4;
      fb[i] = grey; fb[i + 1] = grey; fb[i + 2] = grey;
    }
  }

  const doom = doomCart({ engineUrl: 'about:blank', wadUrl: 'about:blank' });
  const grid = doom.render().grid;
  paintFramebuffer(grid, fb);
  const row = grid.chars[2].slice(1, -1);
  const expectedBoundaries = row.length - 1;
  const glyphTransitions = row.slice(1).reduce(
    (n, ch, i) => n + Number(ch !== row[i]), 0);
  assert.ok(glyphTransitions >= expectedBoundaries - 6,
    `near-colour texture smeared into bands: only ${glyphTransitions}/${expectedBoundaries} glyph boundaries survived`);
  assert.ok(row.includes(' ') && row.includes('█'),
    'both bitmap endpoints must be selectable through the free glyph channel');

  let propertyRuns = 0;
  for (let y = 2; y < grid.chars.length - 1; y++) {
    let prior = null;
    for (let x = 1; x < grid.chars[y].length - 1; x++) {
      const pair = `${grid.colors[y][x]}/${grid.bgs[y][x]}`;
      if (pair !== prior) { propertyRuns++; prior = pair; }
    }
  }
  assert.ok(propertyRuns <= 900,
    `bitmap exceeded its 900-run picture budget: ${propertyRuns}`);
});

test('the pinned canvas font covers every character the demos can draw', () => {
  const drawn = drawnCharacters();
  assert.ok(drawn.size > 80, `expected a rich repertoire, got ${drawn.size} characters`);

  const missing = [...drawn]
    .filter(([ch]) => !covers(ch.codePointAt(0)))
    .map(([ch, source]) => `${hex(ch.codePointAt(0))} ${JSON.stringify(ch)} drawn by ${source}`);
  assert.deepEqual(missing, [],
    'characters outside docxodus-canvas-mono.woff2 fall back to the platform font and break the '
    + 'grid on any device that renders them at a different advance:\n  ' + missing.join('\n  '));
});

test('the shipped subset is single-advance, which is the whole guarantee', () => {
  assert.equal(typeof manifest.advanceUnits, 'number');
  assert.ok(Math.abs(manifest.advanceUnits / manifest.unitsPerEm - manifest.advanceEm) < 1e-6,
    'manifest advance is inconsistent — rebuild with tools/build-canvas-font.sh');
  // Cross-check the file the manifest describes is the file that ships.
  const woff2 = readFileSync(new URL(`../fonts/${manifest.file}`, import.meta.url));
  assert.equal(woff2.subarray(0, 4).toString('latin1'), 'wOF2', 'not a woff2 file');
  assert.ok(woff2.length < 64 * 1024, `canvas font grew to ${woff2.length} bytes — keep it small`);
});

test('the non-ASCII characters the art depends on are all in the subset', () => {
  // The ones the bug was actually about, named so a regression reads clearly.
  const loadBearing = ['█', '▀', '▄', '·', '░', '▒', '▓', '─', '│', '┌', '┐', '└', '┘', '═', '▶', '◀', '►', '◄',
    '▲', '▼', '§', '¶', '·', '→', '←'];
  for (const ch of loadBearing) {
    assert.ok(covers(ch.codePointAt(0)),
      `${hex(ch.codePointAt(0))} ${ch} is not in the pinned font`);
  }
});

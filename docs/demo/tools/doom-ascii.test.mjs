import assert from 'node:assert/strict';
import test from 'node:test';
import { asciiFramebuffer, ASCII_COLS, ASCII_ROWS, ASCII_METRICS } from '../doom-ascii.js';
import { frameXml } from '../ascii-scenes.js';
import { rowsFromXml } from '../ascii-arcade.js';

test('every framebuffer pixel has a printable ASCII cell, including single-pixel HUD strokes', () => {
  const fb = new Uint8Array(320 * 200 * 4);
  for (const [x, y] of [[0, 0], [319, 0], [0, 199], [319, 199], [37, 183], [295, 194]]) {
    const o = (y * 320 + x) * 4;
    fb[o + 2] = 255;
  }
  const grid = asciiFramebuffer(fb);
  assert.equal(ASCII_COLS, 320); assert.equal(ASCII_ROWS, 200);
  assert.equal(grid.chars.length, 200);
  assert.equal(grid.bgs, undefined, 'the image must live in glyphs, not shaded cells');
  for (let y = 0; y < 200; y++) {
    assert.equal(grid.chars[y].length, 321); // printable guard + full frame
    assert.match(grid.chars[y].join(''), /^[\x20-\x7e]+$/);
    for (let x = 0; x < 320; x++)
      assert.equal(grid.chars[y][x + 1] !== ' ', fb[(y * 320 + x) * 4 + 2] !== 0);
  }
  assert.deepEqual(grid, asciiFramebuffer(fb), 'unchanged frames must not shimmer');
  const { xml } = frameXml('<w:p>', grid, '000000', ASCII_METRICS);
  assert.deepEqual(rowsFromXml(xml), grid.chars.map(r => r.join('')));
  assert.equal((xml.match(/<w:shd/g) ?? []).length, 1, 'only the paragraph has a background');
  assert.doesNotMatch(xml, /w:drawing|w:pict/);
});

test('a sharp red/grey boundary is preserved in ink at the exact source column', () => {
  const fb = new Uint8Array(320 * 200 * 4);
  for (let y = 0; y < 200; y++) for (let x = 0; x < 320; x++) {
    const o = (y * 320 + x) * 4;
    fb[o + 2] = 220;
    fb[o + 1] = fb[o] = x < 101 ? 220 : 0;
  }
  const g = asciiFramebuffer(fb);
  assert.equal(g.colors[180][101], 'FFFFFF');
  assert.notEqual(g.colors[180][102], 'FFFFFF');
  assert.equal(g.chars[180][101], g.chars[180][102], 'hue changes do not erase pixel tone');
});

test('invalid framebuffer dimensions are rejected', () => {
  assert.throws(() => asciiFramebuffer(new Uint8Array(4)), RangeError);
});

test('a run budget cannot repaint small blue strokes with the surrounding white ink', () => {
  const fb = new Uint8Array(320 * 200 * 4);
  for (let y = 0; y < 200; y++) for (let x = 0; x < 320; x++) {
    const o = (y * 320 + x) * 4;
    // More than 700 boundaries, including dark saturated single-pixel strokes.
    fb.set(x % 16 === 8 ? [64, 0, 0, 255] : [220, 220, 220, 255], o);
  }
  const grid = asciiFramebuffer(fb);
  for (let y = 0; y < 200; y++) for (let x = 8; x < 320; x += 16) {
    assert.notEqual(grid.chars[y][x + 1], ' ');
    assert.equal(grid.colors[y][x + 1], '0000FF', 'dark blue must not become pale grey');
    assert.equal(grid.colors[y][x], 'FFFFFF', 'the neighboring white stroke must survive too');
  }
});

test('cached hues do not depend on previously rendered frames', async () => {
  const cold = await import('../doom-ascii.js?cold');
  const warm = await import('../doom-ascii.js?warm');
  const frame = new Uint8Array(320 * 200 * 4);
  const preceding = new Uint8Array(frame.length);
  for (let i = 0; i < frame.length; i += 4) {
    const n = i / 4;
    frame.set([n % 256, (n * 17) % 256, (n * 31) % 256, 255], i);
    preceding.set([(n * 13) % 256, (n * 7) % 256, n % 256, 255], i);
  }
  warm.asciiFramebuffer(preceding);
  assert.deepEqual(warm.asciiFramebuffer(frame), cold.asciiFramebuffer(frame));
});

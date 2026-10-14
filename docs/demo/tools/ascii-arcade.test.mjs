import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import {
  ARCADE_KEY_CODES, rowsFromXml,
} from '../ascii-arcade.js';
import { DOOM_KEY_MAP, DOOM_TOUCH, doomCart } from '../doom-cart.js';
import { frameXml } from '../ascii-scenes.js';

const DEMO_DIR = dirname(dirname(fileURLToPath(import.meta.url)));


test('frame XML coalesces matching formatting across line breaks', () => {
  const grid = {
    chars: [['A', 'A'], ['B', 'B']],
    colors: [['FFFFFF', 'FFFFFF'], ['FFFFFF', 'FFFFFF']],
    bgs: [['000000', '000000'], ['000000', '000000']],
  };
  const frame = frameXml('<w:p xmlns:w="urn:test">', grid, '000000');
  assert.equal(frame.runs, 1, 'matching row properties must cross the break in one run');
  assert.equal((frame.xml.match(/<w:r>/g) ?? []).length, 1,
    'line breaks must not create standalone OOXML runs');
  assert.doesNotMatch(frame.xml, /<w:r><w:br\s*\/><\/w:r>/);
  assert.deepEqual(rowsFromXml(frame.xml), ['AA', 'BB']);
});

test('the checked-in Doom GIFs keep their native, tightly framed embed size', () => {
  const readGif = (name) => {
    const bytes = readFileSync(join(DEMO_DIR, '..', 'images', name));
    assert.match(bytes.subarray(0, 6).toString('ascii'), /^GIF8[79]a$/);
    const size = [bytes.readUInt16LE(6), bytes.readUInt16LE(8)];
    let offset = 13;
    const globalTable = bytes[10];
    if (globalTable & 0x80) offset += 3 * (2 ** ((globalTable & 0x07) + 1));
    let frames = 0;
    let durationMs = 0;
    const skipBlocks = () => {
      while (offset < bytes.length) {
        const length = bytes[offset++];
        if (length === 0) return;
        offset += length;
      }
      assert.fail(`unterminated GIF block in ${name}`);
    };
    while (offset < bytes.length) {
      const marker = bytes[offset++];
      if (marker === 0x3b) break;
      if (marker === 0x21) {
        const label = bytes[offset++];
        if (label === 0xf9) {
          assert.equal(bytes[offset++], 4, `bad graphic-control block in ${name}`);
          durationMs += bytes.readUInt16LE(offset + 1) * 10;
          offset += 4;
          assert.equal(bytes[offset++], 0, `unterminated graphic-control block in ${name}`);
        } else {
          skipBlocks();
        }
        continue;
      }
      assert.equal(marker, 0x2c, `unknown GIF block 0x${marker.toString(16)} in ${name}`);
      const localTable = bytes[offset + 8];
      offset += 9;
      if (localTable & 0x80) offset += 3 * (2 ** ((localTable & 0x07) + 1));
      offset++; // LZW minimum code size
      skipBlocks();
      frames++;
    }
    return { size, frames, durationMs, bytes: bytes.length };
  };

  const walkthrough = readGif('arcade-doom.gif');
  assert.deepEqual(walkthrough.size, [656, 716]);
  assert.ok(walkthrough.frames >= 50, `walkthrough has only ${walkthrough.frames} frames`);
  assert.ok(walkthrough.durationMs >= 6500,
    `walkthrough lasts only ${walkthrough.durationMs}ms`);
  assert.ok(walkthrough.bytes < 2 * 1024 * 1024,
    `walkthrough GIF grew to ${walkthrough.bytes} bytes`);
  assert.deepEqual(readGif('arcade-doom-bitmap.gif').size, [656, 660]);
  const readme = readFileSync(join(DEMO_DIR, 'README.md'), 'utf8');
  assert.match(readme, /arcade-doom\.gif\)/, 'the earlier native-image recording remains linked');
  assert.match(readme, /arcade-doom-ascii\.gif[^>]+width="720"/);
  assert.doesNotMatch(readme, /arcade-doom-bitmap\.gif/,
    'the slow bitmap inspection mode must not be showcased as playable');
  assert.doesNotMatch(readme, /arcade-doom\.gif[^>]+width="(?:100%|60%)"/);
});

/** Every button a profile puts on the pad, tray included. */
function touchButtons(profile) {
  const { extras = [], ...slots } = profile;
  return [...Object.values(slots), ...extras];
}

test('every touch button sends a key the arcade actually claims', () => {
  for (const cart of [doomCart()]) {
    assert.ok(cart.touch, `${cart.name} declares no touch profile`);
    for (const button of touchButtons(cart.touch)) {
      assert.ok(ARCADE_KEY_CODES.has(button.code),
        `${cart.name}: the pad's "${button.glyph}" sends ${button.code}, which the arcade `
        + 'does not claim while playing — the press would reach the document, not the game');
      assert.ok(button.glyph && button.label, `${cart.name}: ${button.code} needs a glyph and a label`);
    }
  }
});

test('a thumb can reach every function Doom gives a keyboard', () => {
  // Compared as Doom's OWN key bytes, not as browser codes: the keyboard map
  // has aliases (ArrowUp for KeyW, ShiftRight for ShiftLeft) that a pad has no
  // reason to draw twice. What may not go missing is a Doom function.
  const byTouch = new Set(touchButtons(DOOM_TOUCH).map((button) => {
    const key = DOOM_KEY_MAP[button.code];
    assert.ok(key !== undefined, `the pad's ${button.code} means nothing to Doom`);
    return key;
  }));
  const byKeyboard = new Set(Object.values(DOOM_KEY_MAP));
  const missing = [...byKeyboard].filter((key) => !byTouch.has(key));
  assert.deepEqual(missing, [],
    'Doom functions a keyboard can reach and a thumb cannot: '
    + missing.map((key) => `0x${key.toString(16)}`).join(', '));
});

test('turning and strafing never wear the same arrow', () => {
  // Two pairs of side buttons that do different things; told apart by shape
  // (rotate vs translate), because on a phone they sit one row apart.
  for (const cart of [doomCart()]) {
    const { left, right, strafeLeft, strafeRight } = cart.touch;
    assert.ok(strafeLeft && strafeRight, `${cart.name} moves in a Doom-format level and must strafe`);
    assert.notEqual(left.glyph, strafeLeft.glyph, `${cart.name}: turn and strafe look identical`);
    assert.notEqual(right.glyph, strafeRight.glyph, `${cart.name}: turn and strafe look identical`);
    const labels = touchButtons(cart.touch).map((button) => button.label);
    assert.equal(new Set(labels).size, labels.length, `${cart.name} has two buttons with one name`);
  }
});

test('a modifier a steering thumb cannot hold is offered as a latch', () => {
  for (const cart of [doomCart()]) {
    assert.equal(cart.touch.run.code, 'ShiftLeft');
    assert.equal(cart.touch.run.toggle, true, `${cart.name}: RUN must latch, not ask for a held thumb`);
  }

});

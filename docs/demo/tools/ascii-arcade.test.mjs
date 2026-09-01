import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { dungeonCart, freedoomCart, rowsFromXml } from '../ascii-arcade.js';
import { frameXml } from '../ascii-scenes.js';

const DEMO_DIR = dirname(dirname(fileURLToPath(import.meta.url)));

const MAP_TOP = 4;

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
  assert.match(readme, /arcade-doom\.gif[^>]+width="656"/);
  assert.match(readme, /opener[\s\S]+movement and fire[\s\S]+caption strip/i);
  assert.doesNotMatch(readme, /arcade-doom-bitmap\.gif/,
    'the slow bitmap inspection mode must not be showcased as playable');
  assert.doesNotMatch(readme, /arcade-doom\.gif[^>]+width="(?:100%|60%)"/);
});

function renderedRows(cart) {
  return cart.render().grid.chars.map((row) => row.join(''));
}

function editMapCell(cart, rows, worldX, worldY, glyph) {
  const [wx, wy] = cart.state().window;
  const rowIndex = MAP_TOP + worldY - wy;
  const divider = rows[rowIndex].indexOf('│', 1);
  assert.notEqual(divider, -1, 'rendered map row must contain the view/map divider');
  const column = divider + 2 + worldX - wx;
  rows[rowIndex] = rows[rowIndex].slice(0, column) + glyph + rows[rowIndex].slice(column + 1);
}

class TestInput {
  down = new Set();
  edges = new Set();

  held(...codes) { return codes.some((code) => this.down.has(code)); }
  took(code) {
    const hit = this.edges.has(code);
    this.edges.delete(code);
    return hit;
  }
  set(code, value) {
    if (value && !this.down.has(code)) this.edges.add(code);
    if (value) this.down.add(code);
    else this.down.delete(code);
  }
  clear() { this.down.clear(); this.edges.clear(); }
  endTick() { this.edges.clear(); }
}

test('an unchanged dungeon document keeps D.O.C.X as walls and spawns no monsters', () => {
  const cart = dungeonCart();
  const before = cart.state().mapRow(6);
  cart.syncFromRows(renderedRows(cart));
  const after = cart.state();

  assert.equal(after.mapRow(6), before);
  assert.equal(after.mapRow(6).slice(6, 14), '.D.O.C.X');
  assert.deepEqual(after.enemies, []);
  assert.equal(after.killsTotal, 0);
});

test("only '&' is enemy authoring syntax; z/Z/D remain ordinary letter walls", () => {
  const cart = dungeonCart();
  const rows = renderedRows(cart);
  for (const [x, glyph] of [[9, 'z'], [10, 'Z'], [11, 'D'], [12, '&']]) {
    editMapCell(cart, rows, x, 8, glyph);
  }
  cart.syncFromRows(rows);
  const state = cart.state();

  assert.equal(state.mapRow(8).slice(9, 13), 'zZD.');
  assert.deepEqual(state.enemies.map((enemy) => enemy.kind), ['imp']);
  assert.equal(state.killsTotal, 1);
});

test('a no-edit pause/resume preserves live enemy HP, awareness, position, and terrain', () => {
  const cart = dungeonCart();
  const edited = renderedRows(cart);
  editMapCell(cart, edited, 5, 8, '&');
  cart.syncFromRows(edited);

  const input = new TestInput();
  input.set('Space', true);
  cart.tick(0.01, input);
  input.endTick();
  const before = cart.state();
  assert.equal(before.enemies[0].hp, 1, 'the imp must be wounded before round-trip');
  assert.equal(before.enemies[0].awake, true);

  cart.syncFromRows(renderedRows(cart));
  const after = cart.state();
  assert.deepEqual(after.enemies, before.enemies);
  assert.equal(after.mapRow(8), before.mapRow(8));
  assert.equal(after.killsTotal, before.killsTotal);
});

function bfsNext(state) {
  const rows = [];
  for (let y = 0; ; y++) {
    const row = state.mapRow(y);
    if (!row) break;
    rows.push(row);
  }
  const seekSigil = state.sigilsLeft > 0;
  const target = (ch) => (seekSigil ? ch === '§' : ch === '*');
  const walkable = (ch) => ch === '.' || ch === '§' || (ch === '*' && !seekSigil);
  const sx = Math.floor(state.player.x), sy = Math.floor(state.player.y);
  const previous = new Map([[`${sx},${sy}`, null]]);
  const queue = [[sx, sy]];
  for (let head = 0; head < queue.length; head++) {
    const [x, y] = queue[head];
    if (target(rows[y][x])) {
      let key = `${x},${y}`;
      let prior = previous.get(key);
      while (prior && previous.get(prior) !== null) {
        key = prior;
        prior = previous.get(key);
      }
      const [fx, fy] = key.split(',').map(Number);
      return { fx, fy };
    }
    for (const [nx, ny] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
      const key = `${nx},${ny}`;
      if (nx < 0 || nx >= rows[0].length || ny < 0 || ny >= rows.length || previous.has(key)) continue;
      if (!walkable(rows[ny][nx]) && !target(rows[ny][nx])) continue;
      previous.set(key, `${x},${y}`);
      queue.push([nx, ny]);
    }
  }
  return null;
}

// Both raycaster level packs, cleared end to end by the same autopilot: the
// hand-drawn dungeon and Freedoom's rasterized E1M1. The real Doom engine
// (doom-cart.js) deliberately has no equivalent — its world lives in a
// WebAssembly heap and cannot be walked by a headless script, which is
// exactly why the E1M1 pack stays: its world is readable document text.
for (const [packName, makeCart] of [['dungeon', dungeonCart], ['e1m1', freedoomCart]]) {
test(`the fast headless ${packName} run reaches every objective and the exit`, () => {
  const cart = makeCart();
  const input = new TestInput();
  let goal = null;
  let goalTicks = 0;
  let engaged = null;
  const ignoredUntil = new Map();

  for (let tick = 0; tick < 120_000 && cart.state().mode !== 'won'; tick++) {
    const state = cart.state();
    input.clear();
    if (state.mode === 'dead') {
      cart.tick(0.05, input);
      continue;
    }

    for (const [key, until] of ignoredUntil) if (until < tick) ignoredUntil.delete(key);
    const enemyKey = (enemy) => `${enemy.kind}:${enemy.x.toFixed(0)},${enemy.y.toFixed(0)}`;
    const foe = state.enemies
      .filter((enemy) => enemy.awake && !ignoredUntil.has(enemyKey(enemy)))
      .map((enemy) => ({ enemy, distance: Math.hypot(
        enemy.x - state.player.x, enemy.y - state.player.y,
      ) }))
      .filter(({ distance }) => distance < 6)
      .sort((a, b) => a.distance - b.distance)[0];

    if (foe) {
      const key = enemyKey(foe.enemy);
      if (engaged?.key === key && engaged.hp === foe.enemy.hp) {
        engaged.ticks++;
        if (engaged.ticks > 60) {
          ignoredUntil.set(key, tick + 300);
          engaged = null;
        }
      } else {
        engaged = { key, hp: foe.enemy.hp, ticks: 0 };
      }
      if (engaged) {
        const vx = foe.enemy.x - state.player.x, vy = foe.enemy.y - state.player.y;
        const cross = state.player.dx * vy - state.player.dy * vx;
        const dot = state.player.dx * vx + state.player.dy * vy;
        const aligned = dot > 0 && Math.abs(cross) < 0.25 * foe.distance;
        if (!aligned) input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        if (aligned) input.set('Space', true);
      }
    } else {
      engaged = null;
      if (!goal || (Math.floor(state.player.x) === goal.fx && Math.floor(state.player.y) === goal.fy)
        || ++goalTicks > 30) {
        goal = bfsNext(state);
        goalTicks = 0;
      }
      assert.ok(goal, 'every remaining objective must have a path');
      const tx = goal.fx + 0.5 - state.player.x, ty = goal.fy + 0.5 - state.player.y;
      const cross = state.player.dx * ty - state.player.dy * tx;
      const dot = state.player.dx * tx + state.player.dy * ty;
      if (dot < 0 || Math.abs(cross) > 0.35 * Math.hypot(tx, ty)) {
        input.set(cross < 0 ? 'ArrowLeft' : 'ArrowRight', true);
        if (dot > 0) input.set('KeyW', true);
      } else {
        input.set('KeyW', true);
        input.set('ShiftLeft', true);
      }
    }
    cart.tick(0.05, input);
    input.endTick();
  }

  const final = cart.state();
  assert.equal(final.sigilsLeft, 0);
  assert.equal(final.mode, 'won');
});
}

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { dungeonCart } from '../ascii-arcade.js';

const DEMO_DIR = dirname(dirname(fileURLToPath(import.meta.url)));

const MAP_TOP = 4;

test('the checked-in Doom GIFs keep their native, tightly framed embed size', () => {
  const readGifSize = (name) => {
    const bytes = readFileSync(join(DEMO_DIR, '..', 'images', name));
    assert.match(bytes.subarray(0, 6).toString('ascii'), /^GIF8[79]a$/);
    return [bytes.readUInt16LE(6), bytes.readUInt16LE(8)];
  };

  assert.deepEqual(readGifSize('arcade-doom.gif'), [656, 699]);
  assert.deepEqual(readGifSize('arcade-doom-bitmap.gif'), [656, 699]);
  const readme = readFileSync(join(DEMO_DIR, 'README.md'), 'utf8');
  assert.match(readme, /arcade-doom\.gif[^>]+width="656"/);
  assert.match(readme, /arcade-doom-bitmap\.gif[^>]+width="656"/);
  assert.doesNotMatch(readme, /arcade-doom(?:-bitmap)?\.gif[^>]+width="(?:100%|60%)"/);
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

// This used to drive the Freedoom E1M1 cartridge, which was the same
// raycaster fed a bigger level pack. That cartridge is real Doom now (see
// doom-cart.js), whose world lives in a WebAssembly heap and cannot be walked
// by a headless script — so the raycaster's own "clear the level" logic is
// proved on the pack that still ships.
test('the fast headless cartridge run reaches every objective and the exit', () => {
  const cart = dungeonCart();
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

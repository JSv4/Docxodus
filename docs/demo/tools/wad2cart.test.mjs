import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';

import { FREEDOOM_LEVEL } from '../freedoom-e1m1.js';
import {
  FREEDOOM_BSD_3_CLAUSE,
  formatLevelModule,
  proveReachable,
  rasterize,
} from './wad2cart.mjs';

const SOURCE = {
  kind: 'freedoom',
  ref: 'd14dbbee3b6fbfb2c11cdb65eb61216e86d4ee85',
  path: 'levels/e1m1.wad',
  blob: '98d3dac0a2aaff8d1647fe6f8948742647a46de3',
  sha256: 'd8acd8cf992f1f3af907634587f8a40765d6c43ac15adef1bfa29a2ff1b9e83d',
};

test('generated Freedoom data carries a complete, narrowly scoped BSD notice', async () => {
  assert.match(FREEDOOM_BSD_3_CLAUSE, /Redistributions of source code must retain/);
  assert.match(FREEDOOM_BSD_3_CLAUSE, /Neither the name of the Freedoom project/);
  assert.match(FREEDOOM_BSD_3_CLAUSE, /THIS SOFTWARE IS PROVIDED.*“AS\nIS”/s);

  const formatted = formatLevelModule(
    { ...FREEDOOM_LEVEL, stats: { cell: 32 } },
    FREEDOOM_LEVEL.name,
    { source: SOURCE },
  );
  assert.match(formatted, /SCOPE IS LIMITED TO FREEDOOM_LEVEL/);
  assert.match(formatted, /Docxodus engine, converter, and repository remain MIT-licensed/);
  assert.match(formatted, new RegExp(SOURCE.ref));
  assert.match(formatted, new RegExp(SOURCE.sha256));

  const checkedIn = await readFile(new URL('../freedoom-e1m1.js', import.meta.url), 'utf8');
  assert.equal(checkedIn, formatted, 'checked-in level must be canonical converter output');
});

test('the converter refuses to invent licensing or provenance', () => {
  const out = {
    w: 1, h: 1, rows: ['#'], spawn: { x: 0.5, y: 0.5, dx: 1, dy: 0 },
    monsters: [], stats: { cell: 32 },
  };
  assert.throws(() => formatLevelModule(out, 'E1M1'), /refusing to emit unattributed level data/);
  assert.throws(
    () => formatLevelModule(out, 'E1M1', { source: { ...SOURCE, ref: 'master' } }),
    /--source-ref must be exactly 40 hexadecimal characters/,
  );
  assert.throws(
    () => formatLevelModule(out, 'E1M1', { source: { ...SOURCE, path: '../doom.wad' } }),
    /--source-path must be a relative repository path/,
  );
});

test('rasterization produces a closed, reachable cartridge with objectives and monsters', () => {
  const noSide = 0xffff;
  const level = {
    vertexes: [
      { x: 0, y: 0 }, { x: 256, y: 0 }, { x: 256, y: 256 }, { x: 0, y: 256 },
    ],
    linedefs: [
      { v1: 0, v2: 1, flags: 0, special: 0, tag: 0, right: noSide, left: noSide },
      // Oriented top→bottom so the exit switch's right normal points inside.
      { v1: 2, v2: 1, flags: 0, special: 11, tag: 0, right: noSide, left: noSide },
      { v1: 3, v2: 2, flags: 0, special: 0, tag: 0, right: noSide, left: noSide },
      { v1: 0, v2: 3, flags: 0, special: 0, tag: 0, right: noSide, left: noSide },
    ],
    sidedefs: [],
    sectors: [],
    things: [
      { x: 64, y: 64, angle: 0, type: 1, flags: 7 },
      { x: 96, y: 96, angle: 0, type: 2001, flags: 7 },
      { x: 128, y: 96, angle: 0, type: 3001, flags: 7 },
    ],
  };

  const out = rasterize(level, { cell: 32 });
  assert.equal(out.sigils, 1);
  assert.deepEqual(out.monsters.map(({ kind }) => kind), ['imp']);
  assert.deepEqual(proveReachable(out.rows, out.spawn), []);
  assert.ok(out.rows.some((row) => row.includes('*')));
  assert.ok(out.rows.some((row) => row.includes('§')));
  assert.ok(out.rows.every((row) => row.length === out.w));
  assert.match(out.rows[0], /^#+$/);
  assert.match(out.rows.at(-1), /^#+$/);
});

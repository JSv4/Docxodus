#!/usr/bin/env node
// wad2cart — rasterize a Doom-format level WAD into an Arcade raycaster grid.
//
// The Docx Dungeon cartridge (docs/demo/ascii-arcade.js) walks a cell grid:
// '#' wall, '.' floor, '§' sigil pickup, '*' exit gate. A real Doom level is
// vector geometry — VERTEXES + LINEDEFS bounding SECTORS, with THINGS placed
// in the open space. This tool bridges the two so a real, freely-licensed
// level (Freedoom — BSD licensed) can be PLAYED inside a Word document:
//
//   1. parse the classic binary lumps (THINGS/LINEDEFS/SIDEDEFS/VERTEXES/SECTORS);
//   2. rasterize every *blocking* linedef onto a grid at CELL map units per
//      cell — one-sided lines are solid walls; two-sided lines block when
//      flagged impassable, when the opening is too short to pass through, or
//      when the step up is taller than Doom's 24-unit climb (unless a door /
//      lift / stair special on some linedef animates that sector, in which
//      case it is treated in its player-friendly open/lowered state — the
//      grid world has no door mechanic, so doors stand open);
//   3. flood-fill from the player-1 start: every cell the fill cannot reach
//      is solid, which turns the decorative outside of the map into wall;
//   4. drop '§' sigils on the level's own pickups (keys, weapons, big
//      bonuses — most valuable first, deduped to one per cell) and '*' on
//      the floor in front of the exit switch linedef (special 11/51/52),
//      keeping only spots the BFS proves reachable;
//   5. emit a JS module exporting the grid + spawn, with the source's
//      license notice carried along.
//
// Usage: node wad2cart.mjs <level.wad> <MAPxx|ExMy> <out.js> [--cell=64]
//        node wad2cart.mjs <level.wad> --inspect

import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';

// ─── Binary lump parsing (classic Doom map format) ─────────────────────
export function parseWad(buf) {
  const magic = buf.toString('ascii', 0, 4);
  if (magic !== 'IWAD' && magic !== 'PWAD') throw new Error(`not a WAD: ${magic}`);
  const numLumps = buf.readInt32LE(4);
  const dirOfs = buf.readInt32LE(8);
  const lumps = [];
  for (let i = 0; i < numLumps; i++) {
    const o = dirOfs + i * 16;
    lumps.push({
      pos: buf.readInt32LE(o),
      size: buf.readInt32LE(o + 4),
      name: buf.toString('ascii', o + 8, o + 16).replace(/\0+$/, ''),
    });
  }
  return { magic, lumps };
}

export function mapLumps(wad, mapName) {
  const start = wad.lumps.findIndex((l) => l.name === mapName);
  if (start < 0) throw new Error(`map ${mapName} not found`);
  const out = {};
  const MAP_LUMPS = new Set(['THINGS', 'LINEDEFS', 'SIDEDEFS', 'VERTEXES', 'SEGS',
    'SSECTORS', 'NODES', 'SECTORS', 'REJECT', 'BLOCKMAP', 'BEHAVIOR']);
  for (let i = start + 1; i < wad.lumps.length && MAP_LUMPS.has(wad.lumps[i].name); i++) {
    out[wad.lumps[i].name] = wad.lumps[i];
  }
  return out;
}

export function parseLevel(buf, lumps) {
  const slice = (l) => buf.subarray(l.pos, l.pos + l.size);
  const vertexes = [];
  {
    const b = slice(lumps.VERTEXES);
    for (let o = 0; o + 4 <= b.length; o += 4) {
      vertexes.push({ x: b.readInt16LE(o), y: b.readInt16LE(o + 2) });
    }
  }
  const sidedefs = [];
  {
    const b = slice(lumps.SIDEDEFS);
    for (let o = 0; o + 30 <= b.length; o += 30) {
      sidedefs.push({ sector: b.readInt16LE(o + 28) });
    }
  }
  const sectors = [];
  {
    const b = slice(lumps.SECTORS);
    for (let o = 0; o + 26 <= b.length; o += 26) {
      sectors.push({
        floor: b.readInt16LE(o),
        ceil: b.readInt16LE(o + 2),
        special: b.readInt16LE(o + 22),
        tag: b.readInt16LE(o + 24),
      });
    }
  }
  const linedefs = [];
  {
    const b = slice(lumps.LINEDEFS);
    for (let o = 0; o + 14 <= b.length; o += 14) {
      linedefs.push({
        v1: b.readUint16LE(o), v2: b.readUint16LE(o + 2),
        flags: b.readUint16LE(o + 4),
        special: b.readUint16LE(o + 6),
        tag: b.readUint16LE(o + 8),
        right: b.readUint16LE(o + 10), left: b.readUint16LE(o + 12),
      });
    }
  }
  const things = [];
  {
    const b = slice(lumps.THINGS);
    for (let o = 0; o + 10 <= b.length; o += 10) {
      things.push({
        x: b.readInt16LE(o), y: b.readInt16LE(o + 2),
        angle: b.readInt16LE(o + 4),
        type: b.readUint16LE(o + 6),
        flags: b.readUint16LE(o + 8),
      });
    }
  }
  return { vertexes, linedefs, sidedefs, sectors, things };
}

// ─── Doom semantics ────────────────────────────────────────────────────
const NO_SIDE = 0xffff;
const FLAG_IMPASSABLE = 0x0001;

// Linedef specials that animate a tagged sector into a passable state.
// Doors raise the ceiling; lifts/floors lower or raise the floor to a
// neighbor's level; stair builders raise a run of steps. For a flat grid
// world we treat all of them as "this sector ends up walkable".
const DOOR_SPECIALS = new Set([1, 2, 3, 4, 16, 26, 27, 28, 29, 31, 32, 33, 34,
  42, 46, 50, 61, 63, 75, 76, 86, 90, 99, 103, 105, 106, 107, 108, 109, 110,
  111, 112, 113, 114, 115, 116, 117, 118, 133, 134, 135, 136, 137]);
const FLOOR_MOVER_SPECIALS = new Set([5, 9, 10, 14, 15, 18, 19, 20, 21, 22, 23,
  24, 30, 36, 37, 38, 47, 53, 55, 56, 58, 59, 60, 62, 64, 65, 66, 67, 68, 69,
  70, 71, 88, 91, 92, 93, 94, 95, 96, 98, 100, 101, 102, 119, 120, 121, 122,
  123, 127, 128, 129, 130, 131, 132, 140, 7, 8]);
const EXIT_SPECIALS = new Set([11, 51, 52, 124]);

// THINGS worth a '§' sigil, most iconic first (keys, then big powerups, then
// weapons). Deduped one per grid cell; capped so the HUD count stays readable.
const SIGIL_THINGS = [
  5, 6, 13, 38, 39, 40,          // keycards + skull keys
  2013, 2019, 2023, 2022, 2024,  // soulsphere, blue armor, berserk, invuln, invis
  2001, 2002, 2003, 2004, 2005, 2006, // shotgun, chaingun, RL, plasma, chainsaw, BFG
  8, 2018,                       // backpack, green armor
];
const MTF_NOT_SINGLE = 0x0010;

// ─── Rasterization ─────────────────────────────────────────────────────
export function rasterize(level, { cell = 64, sigilCap = 10, onGrid = null } = {}) {
  const { vertexes, linedefs, sidedefs, sectors, things } = level;

  // Sectors animated by some special become passable: doors open, lifts low.
  const animated = new Set();
  for (const ld of linedefs) {
    if (DOOR_SPECIALS.has(ld.special) || FLOOR_MOVER_SPECIALS.has(ld.special)) {
      if (ld.tag > 0) sectors.forEach((s, i) => { if (s.tag === ld.tag) animated.add(i); });
      // Manual doors (DR/D1: specials 1,26,27,28,31..34,117,118) act on the
      // sector BEHIND the line — no tag needed.
      if ([1, 26, 27, 28, 31, 32, 33, 34, 117, 118].includes(ld.special) && ld.left !== NO_SIDE) {
        animated.add(sidedefs[ld.left].sector);
      }
    }
  }

  const secOf = (sideIx) => (sideIx === NO_SIDE ? null : sectors[sidedefs[sideIx].sector]);
  const secIx = (sideIx) => (sideIx === NO_SIDE ? null : sidedefs[sideIx].sector);

  /** Does this linedef block a walking player, with animated sectors resolved
   *  to their friendly state? */
  function blocks(ld) {
    if (ld.left === NO_SIDE || ld.right === NO_SIDE) return true; // one-sided
    if (ld.flags & FLAG_IMPASSABLE) return true;
    const a = secOf(ld.right), b = secOf(ld.left);
    if (!a || !b) return true;
    const aAnim = animated.has(secIx(ld.right)), bAnim = animated.has(secIx(ld.left));
    // Opening height: door sectors count as open (ceiling out of the way).
    const ceil = Math.min(aAnim ? Infinity : a.ceil, bAnim ? Infinity : b.ceil);
    const floorHi = Math.max(a.floor, b.floor);
    const floorLo = Math.min(a.floor, b.floor);
    if (ceil !== Infinity && ceil - floorHi < 56) return true; // can't fit through
    // Step: >24 units is unclimbable — unless either side animates (lift /
    // stairs / mover), which we resolve to "you can get up there".
    if (floorHi - floorLo > 24 && !aAnim && !bAnim) return true;
    return false;
  }

  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const v of vertexes) {
    minX = Math.min(minX, v.x); maxX = Math.max(maxX, v.x);
    minY = Math.min(minY, v.y); maxY = Math.max(maxY, v.y);
  }
  // One-cell margin so the boundary walls always have a cell to land in.
  minX -= cell; minY -= cell; maxX += cell; maxY += cell;
  const W = Math.ceil((maxX - minX) / cell);
  const H = Math.ceil((maxY - minY) / cell);

  // Doom Y grows north; grid rows grow south — flip Y so the map reads like
  // the automap.
  const toCell = (x, y) => [
    Math.min(W - 1, Math.max(0, Math.floor((x - minX) / cell))),
    Math.min(H - 1, Math.max(0, Math.floor((maxY - y) / cell))),
  ];

  const grid = Array.from({ length: H }, () => new Array(W).fill('.'));

  // Stamp blocking linedefs: sample each segment at quarter-cell steps.
  for (const ld of linedefs) {
    if (!blocks(ld)) continue;
    const a = vertexes[ld.v1], b = vertexes[ld.v2];
    const len = Math.hypot(b.x - a.x, b.y - a.y);
    const steps = Math.max(1, Math.ceil((len / cell) * 4));
    for (let s = 0; s <= steps; s++) {
      const t = s / steps;
      const [cx, cy] = toCell(a.x + (b.x - a.x) * t, a.y + (b.y - a.y) * t);
      grid[cy][cx] = '#';
    }
  }

  // Player-1 start.
  const p1 = things.find((t) => t.type === 1);
  if (!p1) throw new Error('no player 1 start');
  const [px, py] = toCell(p1.x, p1.y);
  if (grid[py][px] === '#') throw new Error('player start rasterized into a wall — try a different --cell');

  // Flood fill: reachable floor. Everything else becomes solid.
  const reach = Array.from({ length: H }, () => new Array(W).fill(false));
  {
    const q = [[px, py]];
    reach[py][px] = true;
    while (q.length) {
      const [x, y] = q.pop();
      for (const [nx, ny] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
        if (nx < 0 || nx >= W || ny < 0 || ny >= H) continue;
        if (reach[ny][nx] || grid[ny][nx] === '#') continue;
        reach[ny][nx] = true;
        q.push([nx, ny]);
      }
    }
  }
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if (!reach[y][x]) grid[y][x] = '#';
  }
  if (onGrid) onGrid(grid, toCell);

  // Exit gate: the cell in front of the exit switch (its right side).
  let exitCell = null;
  for (const ld of linedefs) {
    if (!EXIT_SPECIALS.has(ld.special)) continue;
    const a = vertexes[ld.v1], b = vertexes[ld.v2];
    const mx = (a.x + b.x) / 2, my = (a.y + b.y) / 2;
    // Right-side normal of (v1→v2) in Doom coords is (dy, -dx).
    const dx = b.x - a.x, dy = b.y - a.y;
    const n = Math.hypot(dx, dy) || 1;
    for (const d of [0.6, 1.1, 1.8]) {
      const [cx, cy] = toCell(mx + (dy / n) * cell * d, my - (dx / n) * cell * d);
      if (grid[cy][cx] === '.') { exitCell = [cx, cy]; break; }
    }
    if (exitCell) break;
  }
  if (!exitCell) throw new Error('no reachable exit cell found');
  grid[exitCell[1]][exitCell[0]] = '*';

  // Sigils on the level's own pickups, in SIGIL_THINGS priority order. A
  // pickup whose alcove sealed at raster resolution slides to the nearest
  // reachable floor cell within a few cells — the spot stays the level's own.
  const sigils = [];
  const taken = new Set([exitCell.join(','), [px, py].join(',')]);
  const nearestFloor = (cx, cy, radius) => {
    for (let r = 0; r <= radius; r++) {
      for (let dy = -r; dy <= r; dy++) for (let dx = -r; dx <= r; dx++) {
        if (Math.max(Math.abs(dx), Math.abs(dy)) !== r) continue;
        const x = cx + dx, y = cy + dy;
        if (x < 0 || x >= W || y < 0 || y >= H) continue;
        if (grid[y][x] === '.' && !taken.has(x + ',' + y)) return [x, y];
      }
    }
    return null;
  };
  for (const type of SIGIL_THINGS) {
    for (const t of things) {
      if (t.type !== type || (t.flags & MTF_NOT_SINGLE)) continue;
      const spot = nearestFloor(...toCell(t.x, t.y), 3);
      if (!spot) continue;
      taken.add(spot.join(','));
      sigils.push([...spot, type]);
      if (sigils.length >= sigilCap) break;
    }
    if (sigils.length >= sigilCap) break;
  }
  if (sigils.length === 0) throw new Error('no reachable sigil spots');
  for (const [cx, cy] of sigils) grid[cy][cx] = '§';

  // Spawn facing: Doom angle is degrees CCW from east; grid Y is flipped, so
  // the y component negates.
  const rad = ((p1.angle || 0) * Math.PI) / 180;
  const spawn = {
    x: px + 0.5, y: py + 0.5,
    dx: Math.round(Math.cos(rad) * 1000) / 1000,
    dy: Math.round(-Math.sin(rad) * 1000) / 1000,
  };

  return {
    rows: grid.map((r) => r.join('')),
    w: W, h: H, spawn,
    sigils: sigils.length,
    exit: exitCell,
    stats: { cell, mapUnits: [maxX - minX, maxY - minY] },
  };
}

/** BFS proof: from spawn, every § and the * must be reachable. */
export function proveReachable(rows, spawn) {
  const H = rows.length, W = rows[0].length;
  const walk = (ch) => ch !== '#';
  const sx = Math.floor(spawn.x), sy = Math.floor(spawn.y);
  const seen = Array.from({ length: H }, () => new Array(W).fill(false));
  const q = [[sx, sy]];
  seen[sy][sx] = true;
  while (q.length) {
    const [x, y] = q.pop();
    for (const [nx, ny] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
      if (nx < 0 || nx >= W || ny < 0 || ny >= H) continue;
      if (seen[ny][nx] || !walk(rows[ny][nx])) continue;
      seen[ny][nx] = true;
      q.push([nx, ny]);
    }
  }
  const missing = [];
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    if ((rows[y][x] === '§' || rows[y][x] === '*') && !seen[y][x]) missing.push([x, y, rows[y][x]]);
  }
  return missing;
}

// ─── CLI ───────────────────────────────────────────────────────────────
if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  cli();
}

function cli() {
const [, , wadPath, mapNameArg, outPath, ...rest] = process.argv;
if (!wadPath) {
  console.error('usage: node wad2cart.mjs <level.wad> <MAPxx|ExMy> <out.js> [--cell=64]');
  process.exit(2);
}
const buf = readFileSync(wadPath);
const wad = parseWad(buf);
if (mapNameArg === '--inspect' || !mapNameArg) {
  console.log(wad.magic, wad.lumps.length, 'lumps');
  for (const l of wad.lumps) console.log(`  ${l.name.padEnd(8)} ${l.size}`);
  process.exit(0);
}

const cellArg = rest.find((a) => a.startsWith('--cell='));
const cell = cellArg ? Number(cellArg.split('=')[1]) : 64;
const level = parseLevel(buf, mapLumps(wad, mapNameArg));
const out = rasterize(level, { cell });
const missing = proveReachable(out.rows, out.spawn);
if (missing.length) {
  console.error('UNREACHABLE:', missing);
  process.exit(1);
}
console.log(`grid ${out.w}×${out.h} (${out.stats.mapUnits[0]}×${out.stats.mapUnits[1]} map units @ ${cell}/cell)`);
console.log(`spawn (${out.spawn.x},${out.spawn.y}) facing (${out.spawn.dx},${out.spawn.dy})`);
console.log(`sigils ${out.sigils} · exit at (${out.exit[0]},${out.exit[1]}) · all proven reachable`);
console.log(out.rows.join('\n'));

if (outPath) {
  const notice = `// Level geometry derived from the Freedoom project's ${mapNameArg} (freedoom.github.io),
// rasterized by docs/demo/tools/wad2cart.mjs. Freedoom is © 2001-2024
// Contributors to the Freedoom project, all rights reserved, and is
// distributed under a BSD-style license: redistribution and use in source
// and binary forms, with or without modification, are permitted provided
// the copyright notice and license conditions are retained — see
// https://github.com/freedoom/freedoom/blob/master/COPYING.adoc for the
// full text (conditions + no-endorsement clause + warranty disclaimer).
`;
  const body = `${notice}
// A real Doom-format level as an Arcade raycaster grid: '#' wall, '.' floor,
// '§' sigil (the level's own key/weapon/powerup spots), '*' the exit switch.
// ${out.w}×${out.h} cells at ${cell} map units per cell.
export const FREEDOOM_LEVEL = {
  name: ${JSON.stringify(mapNameArg)},
  w: ${out.w},
  h: ${out.h},
  spawn: ${JSON.stringify(out.spawn)},
  rows: [
${out.rows.map((r) => '    ' + JSON.stringify(r) + ',').join('\n')}
  ],
};
`;
  writeFileSync(outPath, body);
  console.log(`wrote ${outPath}`);
}
}

// THE DOCX ARCADE — playable games rendered INTO a live Word document.
//
// Sibling of ascii-scenes.js (the Observatory) and same contract: the game
// screen is ONE Word paragraph of colored monospaced runs, animated by a
// Unid-preserving `DocxSession.raw.replaceXml` + `DocxEditor.refresh()` per
// frame — the editor's public "the session changed behind your back" seam,
// which reconciles exactly one block in continuous mode. No canvas, no PNG,
// no separate machinery: pause (or click the screen) and the game is only a
// document — the caret lands in it, the whole ribbon works, Ctrl+Z rewinds
// frame by frame, and Save downloads the current frame as a real .docx.
//
// What the games add over the Observatory is INPUT, in both directions:
//   - keyboard in: a capture-phase listener owns WASD/arrows/Space while the
//     game runs, and releases them the moment it pauses (no mode switch — the
//     editor was live the whole time);
//   - document in: on resume, the driver re-parses the game world FROM the
//     paragraph the user just edited. The screen wears a box-drawing bezel so
//     no row can start a markdown block construct (`# ` heading, `- ` bullet)
//     or serialize as a blank line when the editor's blur-commit round-trips
//     the paragraph through markdown `ReplaceText` — which is what makes
//     "type bricks into the level, resume, and they are solid" safe rather
//     than lucky.
//
// This file's home is docs/demo/ for the same reason ascii-scenes.js lives
// there: GitHub Pages deploys docs/ verbatim with no build step, and the npm
// Playwright webroot gets a pretest copy. It is demo content, not library
// machinery, and is deliberately NOT shipped in the npm package.

import { COLS, ROWS, frameXml } from './ascii-scenes.js';
import { FREEDOOM_LEVEL } from './freedoom-e1m1.js';

// ─── Screen geometry ──────────────────────────────────────────────────
// Same 92×26 cell grid as the Observatory (proven to repaint incrementally at
// interactive rates). The bezel spends the outer ring; the HUD spends row 1;
// the playfield is rows FIELD_TOP..ROWS-2 inside columns 1..COLS-2.
const INNER_W = COLS - 2;                 // 90 playfield columns
const FIELD_TOP = 2;                      // 0 = bezel, 1 = HUD
const FIELD_ROWS = ROWS - FIELD_TOP - 1;  // 23 playfield rows

const BEZEL_INK = '33465B';
const HUD_INK = '9CB3C9';
const BANNER_INK = 'FFD166';

function makeGrid() {
  const chars = [], colors = [];
  for (let y = 0; y < ROWS; y++) {
    chars.push(new Array(COLS).fill(' '));
    colors.push(new Array(COLS).fill('FFFFFF'));
  }
  return { chars, colors };
}

function writeText(g, y, x, text, ink) {
  for (let k = 0; k < text.length && x + k < COLS - 1; k++) {
    g.chars[y][x + k] = text[k];
    g.colors[y][x + k] = ink;
  }
}

/** Bezel + HUD chrome shared by both cartridges. The bezel is load-bearing:
 *  every row begins with `│`/`┌`/`└`, so the editor's markdown blur-commit
 *  can never read a game row as a heading or bullet, and no row is ever
 *  whitespace-only (a blank line would split the screen paragraph in two). */
function drawChrome(g, hudText) {
  for (let x = 0; x < COLS; x++) {
    g.chars[0][x] = x === 0 ? '┌' : x === COLS - 1 ? '┐' : '─';
    g.chars[ROWS - 1][x] = x === 0 ? '└' : x === COLS - 1 ? '┘' : '─';
    g.colors[0][x] = BEZEL_INK;
    g.colors[ROWS - 1][x] = BEZEL_INK;
  }
  for (let y = 1; y < ROWS - 1; y++) {
    g.chars[y][0] = '│'; g.chars[y][COLS - 1] = '│';
    g.colors[y][0] = BEZEL_INK; g.colors[y][COLS - 1] = BEZEL_INK;
  }
  writeText(g, 1, 2, hudText.slice(0, INNER_W - 2), HUD_INK);
}

/** Centered banner for win/death moments. Drawn with characters no cartridge
 *  parses back as terrain, so pausing ON a banner cannot deposit tiles. */
function drawBanner(g, lines) {
  const w = Math.max(...lines.map((l) => l.length)) + 6;
  const x0 = Math.floor((COLS - w) / 2);
  const y0 = Math.floor(ROWS / 2) - Math.ceil(lines.length / 2) - 1;
  for (let y = y0; y < y0 + lines.length + 2; y++) {
    for (let x = x0; x < x0 + w; x++) {
      const edge = y === y0 || y === y0 + lines.length + 1;
      g.chars[y][x] = edge ? '~' : x === x0 || x === x0 + w - 1 ? ':' : ' ';
      g.colors[y][x] = BANNER_INK;
    }
  }
  lines.forEach((l, i) =>
    writeText(g, y0 + 1 + i, x0 + Math.floor((w - l.length) / 2), l, BANNER_INK));
}

// Deterministic noise, same recipe as the Observatory (no Math.random: a
// frame must be a pure function of its inputs so tests and repro stay honest).
function hash2(x, y) {
  let h = (x * 374761393 + y * 668265263) | 0;
  h = Math.imul(h ^ (h >>> 13), 1274126177);
  return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
}

// ─── Input: the game borrows the keyboard, the editor gets it back ────
// A capture-phase listener on window sees keys before the contenteditable
// blocks do. While the game plays it claims ONLY the game keys (never a
// chorded shortcut — Ctrl/Cmd/Alt pass through untouched); paused, it claims
// nothing, which is why typing into the document just works.
const GAME_CODES = new Set([
  'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown',
  'KeyW', 'KeyA', 'KeyS', 'KeyD', 'Space', 'KeyR',
  'ShiftLeft', 'ShiftRight',
]);

function createInput(isPlaying) {
  const down = new Set();
  const pressed = new Set(); // keydown edges, consumed once per tick
  const onKeyDown = (e) => {
    if (!isPlaying() || e.metaKey || e.ctrlKey || e.altKey) return;
    if (!GAME_CODES.has(e.code)) return;
    e.preventDefault();
    e.stopPropagation();
    if (!down.has(e.code)) pressed.add(e.code);
    down.add(e.code);
  };
  // keyup always clears, playing or not — a key released while paused must
  // not stay latched into the next resume.
  const onKeyUp = (e) => { down.delete(e.code); };
  const onBlur = () => { down.clear(); pressed.clear(); };
  window.addEventListener('keydown', onKeyDown, true);
  window.addEventListener('keyup', onKeyUp, true);
  window.addEventListener('blur', onBlur);
  return {
    held: (...codes) => codes.some((c) => down.has(c)),
    took: (...codes) => {
      const hit = codes.some((c) => pressed.has(c));
      codes.forEach((c) => pressed.delete(c));
      return hit;
    },
    endTick: () => pressed.clear(),
    /** Synthetic press/release — the on-screen touch pad and tests use this. */
    set: (code, isDown) => {
      if (isDown) { if (!down.has(code)) pressed.add(code); down.add(code); }
      else down.delete(code);
    },
    dispose: () => {
      window.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('keyup', onKeyUp, true);
      window.removeEventListener('blur', onBlur);
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Cartridge 1 — PILCROW'S QUEST (side-scrolling platformer)
// You are ¶, the pilcrow. Collect §, stomp typo gremlins, reach the flag.
// The level IS the document: pause and TYPE terrain into the screen —
// `#` bricks, `=` platforms, `$` coins, `^` spikes, `&` gremlins — resume,
// and the physics engine walks on what you wrote.
// ═══════════════════════════════════════════════════════════════════════

const LEVEL_W = 200;

/** Level tiles are canonical single chars; the renderer picks display glyphs.
 *  '#' solid  '=' platform  '§' coin  '^' spike  'F' flag  ' ' air */
function buildLevel() {
  const rows = [];
  for (let y = 0; y < FIELD_ROWS; y++) rows.push(new Array(LEVEL_W).fill(' '));
  const set = (x, y, ch) => {
    if (x >= 0 && x < LEVEL_W && y >= 0 && y < FIELD_ROWS) rows[y][x] = ch;
  };
  const ground = (x0, x1) => { for (let x = x0; x <= x1; x++) { set(x, 21, '#'); set(x, 22, '#'); } };
  const plat = (x0, x1, y) => { for (let x = x0; x <= x1; x++) set(x, y, '='); };
  const coins = (x0, x1, y) => { for (let x = x0; x <= x1; x += 2) set(x, y, '§'); };
  const spikes = (x0, x1) => { for (let x = x0; x <= x1; x++) set(x, 20, '^'); };
  const wall = (x, y0, y1) => { for (let y = y0; y <= y1; y++) set(x, y, '#'); };

  wall(0, 0, 22); wall(1, 12, 22);           // left boundary
  ground(0, 27);
  coins(10, 16, 18);
  ground(31, 58);                            // first pit at 28..30
  plat(34, 39, 16); coins(34, 38, 15);
  spikes(46, 49);
  plat(45, 50, 17);
  plat(59, 61, 18);                          // stepping stone over pit 59..62
  ground(63, 96);
  coins(66, 74, 19);
  plat(72, 77, 16); plat(80, 85, 12); coins(80, 84, 11);
  ground(101, 132);                          // third pit at 97..100
  spikes(110, 114);
  plat(108, 116, 17);
  coins(108, 116, 16);
  plat(120, 124, 14); plat(128, 132, 11);
  plat(136, 154, 9); coins(138, 152, 8);     // sky bridge
  ground(140, 168);
  ground(173, 199);                          // final pit at 169..172
  for (let y = 14; y <= 20; y++) set(186, y, 'F'); // flagpole: any overlap wins
  wall(198, 0, 22); wall(199, 0, 22);        // right boundary
  return rows;
}

const P_TILE_INK = { '#': '6B7A90', '=': 'C9A96A', '§': 'FFD166', '^': 'FF6B6B', F: '4ADE80' };
// What a typed character means when the document is parsed back into a level.
// '$' aliases '§' (typeable coin); '█'/'▓' alias '#' and '═' aliases '=' (what
// the renderer shows reads back as what it means); '¶' is the player, not a
// tile. Anything unrecognized is air.
const P_PARSE = {
  '#': '#', '█': '#', '▓': '#', '=': '=', '═': '=',
  '§': '§', $: '§', '^': '^', '&': '&', F: 'F', '>': 'F', '║': 'F',
};

/** Exported for the Playwright spec and headless logic checks. */
export function platformerCart() {
  const solid = (ch) => ch === '#' || ch === '=';
  let level, gremlins, player, camX, coinsGot, coinsTotal, score, deaths, state, msgT;

  function countCoins() {
    let n = 0;
    for (const row of level) for (const ch of row) if (ch === '§') n++;
    return n;
  }
  const spawnGremlin = (x, y) => gremlins.push({ x, y, dir: -1 });
  const spawnPlayer = () => ({ x: 4, y: 20, vx: 0, vy: 0, onGround: true });

  function reset() {
    level = buildLevel();
    gremlins = [];
    // Gremlin posts, standing on the ground/platforms they patrol.
    [[36, 20], [70, 20], [82, 11], [112, 16], [146, 8], [160, 20]].forEach(([x, y]) => spawnGremlin(x, y));
    player = spawnPlayer();
    camX = 0; coinsGot = 0; score = 0; deaths = 0; state = 'run'; msgT = 0;
    coinsTotal = countCoins();
  }
  reset();

  const at = (x, y) =>
    x < 0 || x >= LEVEL_W ? '#' : y >= 0 && y < FIELD_ROWS ? level[y][x] : ' ';

  function die() {
    deaths++; msgT = 1.0; state = 'dead';
  }

  /** Frame ticks arrive at whatever cadence the document render sustains
   *  (≈100–200 ms) — far too coarse to integrate against directly: at
   *  dt = 0.2 a fall covers 2.4 cells per step, which tunnels straight
   *  through the two-row ground. Physics therefore sub-steps at ≤ 60 ms;
   *  key EDGES (`took`) are consumed by the first sub-step only, which is
   *  exactly the semantics an edge should have. */
  function tickSim(dt, input) {
    const n = Math.max(1, Math.ceil(dt / 0.06));
    for (let i = 0; i < n; i++) stepSim(dt / n, input);
  }

  function stepSim(dt, input) {
    if (state === 'won') {
      if (input.took('KeyR')) reset();
      return;
    }
    if (state === 'dead') {
      msgT -= dt;
      if (msgT <= 0) { player = spawnPlayer(); state = 'run'; }
      return;
    }
    const p = player;
    const left = input.held('ArrowLeft', 'KeyA');
    const right = input.held('ArrowRight', 'KeyD');
    const ACCEL = 60, FRICTION = 40, VMAX = 16, GRAV = 60, JUMP = 26;
    if (left && !right) p.vx = Math.max(-VMAX, p.vx - ACCEL * dt);
    else if (right && !left) p.vx = Math.min(VMAX, p.vx + ACCEL * dt);
    else p.vx -= Math.sign(p.vx) * Math.min(Math.abs(p.vx), FRICTION * dt);
    if (input.took('ArrowUp', 'KeyW', 'Space') && p.onGround) { p.vy = -JUMP; p.onGround = false; }
    // Releasing jump early cuts the rise — the classic variable-height jump.
    if (p.vy < 0 && !input.held('ArrowUp', 'KeyW', 'Space')) p.vy += GRAV * 0.9 * dt;

    const nx = p.x + p.vx * dt;
    if (solid(at(Math.round(nx), Math.round(p.y)))) p.vx = 0;
    else p.x = Math.max(1, Math.min(LEVEL_W - 2, nx));

    // Grounded, the player is GLUED to the surface (no gravity accumulating
    // into a sub-pixel bounce); airborne, integrate and snap to the row
    // outside the first solid cell. Walking off a ledge takes one sub-step
    // to register — accidental coyote time, kept on purpose.
    if (p.onGround && !solid(at(Math.round(p.x), Math.round(p.y) + 1))) p.onGround = false;
    if (!p.onGround) {
      p.vy = Math.min(30, p.vy + GRAV * dt);
      const ny = p.y + p.vy * dt;
      const cy = Math.round(ny);
      if (solid(at(Math.round(p.x), cy))) {
        if (p.vy > 0) { p.y = cy - 1; p.onGround = true; }
        else p.y = cy + 1;
        p.vy = 0;
      } else p.y = ny;
    } else p.vy = 0;

    const cx = Math.round(p.x), pcy = Math.round(p.y);
    const here = at(cx, pcy);
    if (here === '§') { level[pcy][cx] = ' '; coinsGot++; score += 50; }
    if (here === '^') return die();
    if (here === 'F') { state = 'won'; score += 500; return; }
    if (p.y > FIELD_ROWS + 2) return die();

    for (const g of gremlins) {
      if (g.dead) continue;
      const gx = g.x + g.dir * 4 * dt;
      const front = at(Math.round(gx + g.dir * 0.5), Math.round(g.y));
      const footing = at(Math.round(gx + g.dir * 0.5), Math.round(g.y) + 1);
      if (solid(front) || !solid(footing)) g.dir = -g.dir;
      else g.x = gx;
      if (Math.abs(p.x - g.x) < 0.9) {
        // The stomp band reaches 2.4 cells above the gremlin — deeper than
        // one sub-step of terminal fall speed (~1.2 cells), so a fast fall
        // cannot tunnel past its head between collision checks.
        if (p.vy > 2 && p.y < g.y - 0.1 && p.y > g.y - 2.4) {
          g.dead = true; p.vy = -14; score += 100;
        } else if (Math.abs(p.y - g.y) < 0.9) return die();
      }
    }
    camX = Math.max(0, Math.min(LEVEL_W - INNER_W, Math.round(p.x) - Math.floor(INNER_W / 2)));
  }

  function render() {
    const g = makeGrid();
    drawChrome(g,
      `¶ PILCROW'S QUEST   § ${String(coinsGot).padStart(2, '0')}/${String(coinsTotal).padStart(2, '0')}` +
      `   score ${String(score).padStart(5, '0')}   splats ${deaths}   → reach the > flag`);
    for (let ry = 0; ry < FIELD_ROWS; ry++) {
      const gy = FIELD_TOP + ry;
      for (let rx = 0; rx < INNER_W; rx++) {
        const lx = camX + rx;
        const ch = at(lx, ry);
        if (ch === 'F') {
          // The pole tops out in a pennant.
          g.chars[gy][1 + rx] = at(lx, ry - 1) === 'F' ? '║' : '>';
          g.colors[gy][1 + rx] = P_TILE_INK.F;
        } else if (ch !== ' ') {
          g.chars[gy][1 + rx] = ch === '#' ? '█' : ch === '=' ? '═' : ch;
          g.colors[gy][1 + rx] = P_TILE_INK[ch] ?? HUD_INK;
        } else if (ry < 14 && hash2(lx, ry) < 0.02) {
          g.chars[gy][1 + rx] = '·';        // sparse stars in bezel ink, so a
          g.colors[gy][1 + rx] = BEZEL_INK; // sky row stays a single run
        }
      }
    }
    for (const gr of gremlins) {
      if (gr.dead) continue;
      const rx = Math.round(gr.x) - camX, ry = Math.round(gr.y);
      if (rx >= 0 && rx < INNER_W && ry >= 0 && ry < FIELD_ROWS) {
        g.chars[FIELD_TOP + ry][1 + rx] = '&';
        g.colors[FIELD_TOP + ry][1 + rx] = 'C084FC';
      }
    }
    const prx = Math.round(player.x) - camX, pry = Math.round(player.y);
    if (prx >= 0 && prx < INNER_W && pry >= 0 && pry < FIELD_ROWS) {
      g.chars[FIELD_TOP + pry][1 + prx] = '¶';
      g.colors[FIELD_TOP + pry][1 + prx] = '5EEAD4';
    }
    if (state === 'won') drawBanner(g, ['LEVEL CLEAR — the pilcrow prevails', 'press R to run it back']);
    else if (state === 'dead') drawBanner(g, ['SPLAT']);
    return { grid: g, bg: '0E1C30' };
  }

  /** Merge document-parsed playfield rows back into the level at the camera
   *  window. Rows the parse cannot validate (bezel gone) keep their old
   *  content — forgiving of any chaos typed while paused. */
  function syncFromRows(rows) {
    if (state !== 'run') return; // a banner is on screen, not the level
    for (let ry = 0; ry < FIELD_ROWS; ry++) {
      const line = rows[FIELD_TOP + ry];
      if (line == null) continue;
      const bezel = line.search(/[│|]/);
      if (bezel < 0) continue;
      const content = line.slice(bezel + 1);
      // Gremlins inside the window are re-seeded from what the text says.
      gremlins = gremlins.filter((gr) =>
        Math.round(gr.y) !== ry || gr.x < camX || gr.x >= camX + INNER_W);
      for (let rx = 0; rx < INNER_W; rx++) {
        const raw = content[rx];
        if (raw === undefined) break;
        const ch = P_PARSE[raw] ?? ' ';
        if (ch === '&') {
          // A gremlin is an entity, not a tile — keep whatever tile it was
          // standing over (it patrols across coins and spikes), so pausing
          // on top of one cannot erase the tile beneath it.
          spawnGremlin(camX + rx, ry);
          const prev = level[ry][camX + rx];
          level[ry][camX + rx] = prev === '§' || prev === '^' ? prev : ' ';
        } else level[ry][camX + rx] = ch;
      }
    }
    coinsTotal = countCoins() + coinsGot;
  }

  return {
    name: 'quest',
    label: '¶ Pilcrow’s Quest',
    caption:
      'Run **A/D** or **←/→** · jump **W/↑/Space** · stomp the & gremlins, collect every §, ' +
      'reach the > flag. Pause and TYPE terrain into the screen — # bricks, = platforms, ' +
      '$ coins, ^ spikes, & gremlins — then resume: you wrote the level.',
    hint: '<b>A/D</b> run · <b>W</b> jump · pause & type <b>#</b> bricks / <b>$</b> coins straight into the document — resume and they are real.',
    reset,
    tick: tickSim,
    render,
    syncFromRows,
    state: () => ({
      player: { x: player.x, y: player.y },
      camX, coinsGot, coinsTotal, score, deaths, mode: state,
      gremlins: gremlins.filter((gr) => !gr.dead).length,
      gremlinList: gremlins.filter((gr) => !gr.dead)
        .map((gr) => ({ x: gr.x, y: gr.y, dir: gr.dir })),
      levelRow: (y) => (level[y] ?? []).join(''),
    }),
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Cartridge 2 — THE DOCX DUNGEON (Doom-style raycast crawler)
// A first-person maze rendered by a textbook DDA raycaster — into Word runs.
// The MAP panel on the right is PART of the same paragraph: pause, type a
// wall (any letter — your name works) into the map, resume, and it stands
// in the corridor as 3-D towers of that letter. Collect every § sigil, then
// step through the * gate.
//
// The raycaster is a LEVEL PACK player: the same renderer, controls, and
// map-band round-trip run both the hand-drawn 24×16 dungeon and a real
// Doom-format level (Cartridge 3 — Freedoom's E1M1, rasterized to a grid by
// docs/demo/tools/wad2cart.mjs). Maps larger than the band scroll it as a
// 24×16 window that follows the player; typing into the window edits the
// world cells it shows, exactly as before.
// ═══════════════════════════════════════════════════════════════════════

const VIEW_W = 64;                 // 3-D viewport columns inside the bezel
const DIV_X = 1 + VIEW_W;          // grid column of the │ divider
const BAND_X = DIV_X + 1;          // first grid column of the map band
const MAP_TOP = FIELD_TOP + 2;     // first grid row of map cells in the band
const WIN_W = 24, WIN_H = 16;      // map-band window (the classic map's size)

// Enemy bestiary. Entities, not tiles — they billboard in the 3-D view and
// overlay the map band, and typing '&' into the MAP conjures one (the same
// glyph the platformer's gremlins answer to). Speeds are cells/second at
// wallScale 1 and scale with the pack, like the player's stride.
const ENEMY_KINDS = {
  zombie:   { glyph: 'z', ink: 'A8B78A', hp: 1, speed: 1.1, dps: 9 },
  sergeant: { glyph: 'Z', ink: '9AAFC4', hp: 1, speed: 1.2, dps: 13 },
  imp:      { glyph: '&', ink: 'C084FC', hp: 2, speed: 1.5, dps: 11 },
  demon:    { glyph: 'D', ink: 'FF6B6B', hp: 3, speed: 2.0, dps: 17 },
};
const ENEMY_BY_GLYPH = Object.fromEntries(
  Object.entries(ENEMY_KINDS).map(([kind, k]) => [k.glyph, kind]));

// ── Enemy stamps: original ASCII takes on the classic archetypes — a rifle
// grunt, its armored sergeant, a horned imp, a big-jawed demon. Drawn on a
// tall 10×13 grid (a text cell is ~2:1, so this reads as a standing figure),
// ' ' transparent, each char mapped to an ink; a hit-flash paints them all
// white. The 3-D renderer nearest-neighbor samples a stamp into the sprite's
// distance-scaled rectangle: at range a monster is a smudge, up close a face.
const GRUNT_ROWS = [
  '   ####   ',
  '  ######  ',
  '  #%%%%#  ',
  '  %o%%o%  ',
  '  #%~~%#  ',
  '  .####.  ',
  ' ######## ',
  '##########',
  '## #### ##',
  '== #### ==',
  '  ##  ##  ',
  '  ##  ##  ',
  ' ###  ### ',
];
const ENEMY_STAMPS = {
  zombie: {
    rows: GRUNT_ROWS,
    inks: { '#': 'A8B78A', '%': 'D6C39A', '~': '8A5B4A', o: 'FF5555', '=': '9AA5B1', '.': '6E7D5B' },
  },
  sergeant: {
    rows: GRUNT_ROWS,
    inks: { '#': '8B9DAF', '%': 'D6C39A', '~': '8A5B4A', o: 'FF5555', '=': 'E2E8F0', '.': '5C6B7E' },
  },
  imp: {
    rows: [
      ' \\      / ',
      '  \\####/  ',
      ' ######## ',
      ' #@####@# ',
      ' ######## ',
      '  #vvvv#  ',
      ' ######## ',
      '##  ##  ##',
      '#  ####  #',
      '   ####   ',
      '  ##  ##  ',
      ' ##    ## ',
      ' #      # ',
    ],
    inks: { '#': 'C084FC', '\\': 'E9D5FF', '/': 'E9D5FF', '@': 'FFD166', v: 'FFFFFF' },
  },
  demon: {
    rows: [
      '  ######  ',
      ' ######## ',
      '#o######o#',
      '##########',
      '#MMMMMMMM#',
      '#WWWWWWWW#',
      ' ######## ',
      ' ######## ',
      '##########',
      '###    ###',
      '##      ##',
      '###    ###',
      '####  ####',
    ],
    inks: { '#': 'FF6B6B', o: 'FFF7B0', M: 'FFFFFF', W: 'F1D1D1' },
  },
};

// 24×16, every row exactly MAPW chars (the headless harness re-checks this
// and walks the maze to prove every § and the * gate stay reachable). The
// D.O.C.X cells are free-standing letter pillars in the entry hall — the
// first thing the spawn view shows, and the template for "type your own".
const DUNGEON_MAP = [
  '########################',
  '#......................#',
  '#.####.###.##.#####.##.#',
  '#.#..§...#.....§....#..#',
  '#.#..##..######..##....#',
  '#......#.......#..#..#.#',
  '#..#...D.O.C.X....#....#',
  '#.##...........##.###..#',
  '#......................#',
  '#..###..#####..###..##.#',
  '#..#.§..#...#..§#....#.#',
  '#..#....#.*.#...#..#...#',
  '#..######...#####..#...#',
  '#....#...§.....#...#.§.#',
  '#......................#',
  '########################',
];

/** The two level packs the raycaster plays. Geometry is the only difference:
 *  the dungeon's cells are corridor-sized (≈64 map units of a Doom level),
 *  while the rasterized Freedoom level uses 32-unit cells so its 64-unit
 *  corridors survive the grid — hence its doubled stride and wall height. */
const DUNGEON_PACK = {
  name: 'dungeon',
  label: '▓ The Docx Dungeon',
  hudTitle: 'THE DOCX DUNGEON',
  bg: '0A0F1A',
  caption:
    'Move **W/S** · strafe **A/D** · turn **←/→** · hold **Shift** to sprint. Collect every ' +
    '§ sigil, then step through the * gate. The MAP panel is part of the document — pause, ' +
    'type walls into it (any letter: your name works), resume, and walk your word in 3-D.',
  hint: '<b>WASD</b> move · <b>←/→</b> turn · pause & type your name into the <b>MAP</b> — resume and walk through it in 3-D.',
  winBanner: ['THE DUNGEON IS CLEARED', 'every § recovered - press R to delve again'],
  rows: DUNGEON_MAP, w: 24, h: 16,
  spawn: { x: 3.5, y: 8.5, dx: 1, dy: 0 }, // the open hall, down its long axis
  moveSpeed: 3.4, // cells/second
  wallScale: 1,   // one map cell = one full-height wall
  coneRadius: 7,
};

const FREEDOOM_PACK = {
  name: 'e1m1',
  label: '☩ Freedoom E1M1',
  hudTitle: 'FREEDOOM E1M1',
  bg: '120D0A',
  caption:
    'A REAL Doom-format level — E1M1 from the Freedoom project (BSD-licensed), its 1,175 ' +
    'linedefs rasterized onto this grid by `wad2cart.mjs` — playable inside a Word paragraph, ' +
    'monsters included (its own zombies, imps and demons, right where the level put them). ' +
    'Move **W/S** · strafe **A/D** · turn **←/→** · **Shift** sprints · **Space** fires. ' +
    'Recover every § the level placed, then step through the * exit switch. ' +
    'The scrolling MAP window is still just document text: pause, type walls — or `&` baddies — resume.',
  hint: '<b>WASD</b> move · <b>←/→</b> turn · <b>Space</b> fire — a real Doom-format map with its own monsters; § heal and open the gate.',
  winBanner: ['E1M1 CLEARED', 'a real Doom-format level, beaten inside a Word document', 'press R to rip and tear again'],
  rows: FREEDOOM_LEVEL.rows, w: FREEDOOM_LEVEL.w, h: FREEDOOM_LEVEL.h,
  spawn: FREEDOOM_LEVEL.spawn,
  monsters: FREEDOOM_LEVEL.monsters,
  moveSpeed: 6.8, // 32-unit cells: double the dungeon's 64-unit stride
  wallScale: 2,   // Doom walls span two 32-unit cells
  coneRadius: 14,
};

/** The raycaster, as a level-pack player. Exported to the Playwright spec and
 *  headless logic checks through dungeonCart()/freedoomCart() below. */
function raycastCart(pack) {
  const W = pack.w, H = pack.h, S = pack.wallScale;
  let map, px, py, dx, dy, plx, ply, state;
  let enemies, health, kills, killsTotal, fireCooldown, muzzleT, hurtT, deathT;

  const cell = (x, y) => (x >= 0 && x < W && y >= 0 && y < H ? map[y][x] : '#');
  function sigilsLeft() {
    let n = 0;
    for (const row of map) for (const ch of row) if (ch === '§') n++;
    return n;
  }
  const isWall = (ch) => !(ch === '.' || ch === ' ' || ch === '§' || (ch === '*' && sigilsLeft() === 0));

  function normalizeMap(rows) {
    map = [];
    for (let y = 0; y < H; y++) {
      const src = rows[y] ?? '';
      const out = [];
      for (let x = 0; x < W; x++) {
        let ch = src[x] ?? '.';
        if (ch === ' ') ch = '.';
        if (ch === '$') ch = '§';
        if (ch === '│' || ch === '|') ch = '#';
        out.push(ch);
      }
      map.push(out);
    }
    // The outer ring is always wall — the maze may be rewritten at will, but
    // the world stays closed.
    for (let x = 0; x < W; x++) { map[0][x] = '#'; map[H - 1][x] = '#'; }
    for (let y = 0; y < H; y++) { map[y][0] = '#'; map[y][W - 1] = '#'; }
  }

  /** The map-band window: the classic 24×16 map in full, or — when the level
   *  outgrows the band — a 24×16 view that follows the player. Both render
   *  and the resume-parse derive it from the (unmoving-while-paused) player
   *  position, so what you typed lands exactly where you typed it. */
  const winPos = () => [
    Math.max(0, Math.min(Math.max(0, W - WIN_W), Math.floor(px) - (WIN_W >> 1))),
    Math.max(0, Math.min(Math.max(0, H - WIN_H), Math.floor(py) - (WIN_H >> 1))),
  ];

  /** If the player's cell became a wall (someone typed on them), step to the
   *  nearest floor cell so the world stays playable. */
  function unstickPlayer() {
    const cx = Math.floor(px), cy = Math.floor(py);
    if (!isWall(cell(cx, cy))) return;
    const q = [[cx, cy]], seen = new Set([cx + ',' + cy]);
    while (q.length) {
      const [x, y] = q.shift();
      if (!isWall(cell(x, y))) { px = x + 0.5; py = y + 0.5; return; }
      for (const [ax, ay] of [[x + 1, y], [x - 1, y], [x, y + 1], [x, y - 1]]) {
        const k = ax + ',' + ay;
        if (ax >= 0 && ax < W && ay >= 0 && ay < H && !seen.has(k)) {
          seen.add(k); q.push([ax, ay]);
        }
      }
    }
  }

  const spawnEnemy = (x, y, kind) => enemies.push({
    x, y, kind, hp: ENEMY_KINDS[kind].hp, awake: false, flashT: 0,
  });

  function reset() {
    normalizeMap(pack.rows);
    px = pack.spawn.x; py = pack.spawn.y;
    const n = Math.hypot(pack.spawn.dx, pack.spawn.dy) || 1;
    dx = pack.spawn.dx / n; dy = pack.spawn.dy / n;
    plx = -dy * 0.577; ply = dx * 0.577; // FOV ≈ 60°
    enemies = [];
    for (const m of pack.monsters ?? []) spawnEnemy(m.x + 0.5, m.y + 0.5, m.kind);
    health = 100; kills = 0; killsTotal = enemies.length;
    fireCooldown = 0; muzzleT = 0; hurtT = 0; deathT = 0;
    state = 'run';
  }
  reset();

  /** Straight-line sight check between two points, sampled sub-cell. */
  function lineOfSight(ax, ay, bx, by) {
    const d = Math.hypot(bx - ax, by - ay);
    const steps = Math.max(1, Math.ceil(d * 3));
    for (let s = 1; s < steps; s++) {
      const t = s / steps;
      if (isWall(cell(Math.floor(ax + (bx - ax) * t), Math.floor(ay + (by - ay) * t)))) return false;
    }
    return true;
  }

  /** Distance to the first wall straight ahead — the sidearm's reach. */
  function wallDistAhead() {
    let mx = Math.floor(px), my = Math.floor(py);
    const ddx = Math.abs(1 / (dx || 1e-9)), ddy = Math.abs(1 / (dy || 1e-9));
    const stx = dx < 0 ? -1 : 1, sty = dy < 0 ? -1 : 1;
    let sx = dx < 0 ? (px - mx) * ddx : (mx + 1 - px) * ddx;
    let sy = dy < 0 ? (py - my) * ddy : (my + 1 - py) * ddy;
    let side = 0, guard = 0;
    while (guard++ < W + H + 8) {
      if (sx < sy) { sx += ddx; mx += stx; side = 0; } else { sy += ddy; my += sty; side = 1; }
      if (isWall(cell(mx, my))) break;
    }
    return Math.max(0.05, side === 0 ? sx - ddx : sy - ddy);
  }

  /** The sidearm: hitscan along the view center. Hits the nearest live enemy
   *  inside a narrow cone, if no wall stands in front of it. */
  function fire() {
    fireCooldown = 0.45; muzzleT = 0.15;
    const reach = wallDistAhead();
    let best = null;
    for (const e of enemies) {
      if (e.hp <= 0) continue;
      const vx = e.x - px, vy = e.y - py;
      const d = Math.hypot(vx, vy);
      if (d < 0.2 || d > reach + 0.4) continue;
      const ahead = (vx * dx + vy * dy) / d;
      // Cone widens up close (a body fills more of the view): base 4° plus
      // the angular half-width of a ~0.45-cell-wide target.
      if (ahead < Math.cos(0.07 + Math.atan(0.45 / d))) continue;
      if (!best || d < best.d) best = { e, d };
    }
    if (best) {
      best.e.awake = true;
      best.e.hp -= 1;
      best.e.flashT = 0.2;
      if (best.e.hp <= 0) kills++;
    }
  }

  function rotate(a) {
    const c = Math.cos(a), s = Math.sin(a);
    [dx, dy] = [dx * c - dy * s, dx * s + dy * c];
    [plx, ply] = [plx * c - ply * s, plx * s + ply * c];
  }

  function tryMove(nx, ny) {
    const R = 0.22;
    if (!isWall(cell(Math.floor(nx + Math.sign(nx - px) * R), Math.floor(py)))) px = nx;
    if (!isWall(cell(Math.floor(px), Math.floor(ny + Math.sign(ny - py) * R)))) py = ny;
  }

  /** Same sub-stepping as the platformer: browser frame dt (up to 0.2 s)
   *  is too coarse for wall-slide collision and turn feel. */
  function tickSim(dt, input) {
    const n = Math.max(1, Math.ceil(dt / 0.06));
    for (let i = 0; i < n; i++) stepSim(dt / n, input);
  }

  function stepSim(dt, input) {
    if (state === 'won') {
      if (input.took('KeyR')) reset();
      return;
    }
    fireCooldown = Math.max(0, fireCooldown - dt);
    muzzleT = Math.max(0, muzzleT - dt);
    hurtT = Math.max(0, hurtT - dt);
    if (state === 'dead') {
      deathT -= dt;
      if (deathT <= 0) {
        // Doom's contract: back to the start, monsters keep their grudges,
        // your progress (sigils, kills) stands.
        px = pack.spawn.x; py = pack.spawn.y;
        health = 100; state = 'run';
      }
      return;
    }
    const sprint = input.held('ShiftLeft', 'ShiftRight') ? 1.7 : 1;
    const mv = pack.moveSpeed * sprint * dt, rot = 2.6 * dt;
    if (input.held('ArrowLeft')) rotate(-rot);
    if (input.held('ArrowRight')) rotate(rot);
    if (input.held('KeyW', 'ArrowUp')) tryMove(px + dx * mv, py + dy * mv);
    if (input.held('KeyS', 'ArrowDown')) tryMove(px - dx * mv, py - dy * mv);
    if (input.held('KeyA')) tryMove(px + dy * mv, py - dx * mv);
    if (input.held('KeyD')) tryMove(px - dy * mv, py + dx * mv);
    // Hold-to-fire, Doom-pistol style: the cooldown sets the rate, holding
    // Space keeps shooting. (Edge-triggered fire starved at low frame rates —
    // one shot per rendered frame at best.)
    if (input.held('Space') && fireCooldown <= 0) fire();
    else input.took('Space'); // consume stray edges so pause/resume can't bank one

    // Enemies: sleep until they see you (or get shot), then close in. A
    // chaser that loses sight of you for a few seconds loses interest and
    // stands down — otherwise one stuck behind a wall would besiege that
    // wall forever.
    for (const e of enemies) {
      if (e.hp <= 0) continue;
      e.flashT = Math.max(0, e.flashT - dt);
      const vx = px - e.x, vy = py - e.y;
      const d = Math.hypot(vx, vy);
      if (!e.awake) {
        if (d < 8 * S && lineOfSight(e.x, e.y, px, py)) { e.awake = true; e.boredT = 0; }
        else continue;
      } else {
        e.losCheckT = (e.losCheckT ?? 0) + dt;
        if (e.losCheckT >= 0.5) {
          e.losCheckT = 0;
          if (lineOfSight(e.x, e.y, px, py)) e.boredT = 0;
          else if ((e.boredT = (e.boredT ?? 0) + 0.5) >= 5) { e.awake = false; continue; }
        }
      }
      if (d > 0.8) {
        const step = ENEMY_KINDS[e.kind].speed * S * dt;
        const nx = e.x + (vx / d) * step, ny = e.y + (vy / d) * step;
        if (!isWall(cell(Math.floor(nx), Math.floor(e.y)))) e.x = nx;
        if (!isWall(cell(Math.floor(e.x), Math.floor(ny)))) e.y = ny;
      }
      if (d < 1.0) {
        health -= ENEMY_KINDS[e.kind].dps * dt;
        hurtT = 0.25;
      }
    }
    if (health <= 0) {
      health = 0; state = 'dead'; deathT = 1.2;
      return;
    }

    const cx = Math.floor(px), cy = Math.floor(py);
    if (cell(cx, cy) === '§') {
      map[cy][cx] = '.';
      health = Math.min(100, health + 15); // the level's supplies patch you up
    }
    if (cell(cx, cy) === '*' && sigilsLeft() === 0) state = 'won';
  }

  // Ink is the expensive channel (every color change inside a row is its own
  // w:r; the converter pays ~1 ms per run), so distance uses only THREE ink
  // bands while the five-step glyph ramp — and the E/W side-lighting, done as
  // a glyph-density shift — carry the depth for free.
  const SLATE = ['C8D4E2', '75879E', '3D4C60'];
  const TEAL = ['7CEFDC', '2FB5A0', '17604F'];
  const SHADE = ['█', '█', '▓', '▒', '░', '░'];

  function render() {
    const g = makeGrid();
    const left = sigilsLeft();
    const combat = killsTotal > 0 || kills > 0 || enemies.some((e) => e.hp > 0);
    drawChrome(g,
      `${pack.hudTitle}   ` +
      (combat ? `HP ${String(Math.ceil(health)).padStart(3)}   kills ${kills}/${killsTotal}   ` : '') +
      `§ left ${left}   gate ${left === 0 ? 'OPEN - step on *' : 'SEALED'}   ` +
      (combat ? 'WASD move - Space fire' : 'WASD move - arrows turn'));
    // The divider starts below the HUD row, which spans the full bezel width.
    for (let y = FIELD_TOP; y < ROWS - 1; y++) {
      g.chars[y][DIV_X] = '│';
      g.colors[y][DIV_X] = BEZEL_INK;
    }

    // Floor: one dim ink, denser toward the viewer — merges into ~1 run/row.
    const mid = FIELD_TOP + Math.floor(FIELD_ROWS / 2);
    for (let y = mid + 1; y < FIELD_TOP + FIELD_ROWS; y++) {
      const depth = y - mid;
      for (let x = 0; x < VIEW_W; x++) {
        if ((x * 7 + y * 13) % Math.max(3, 10 - depth) === 0) {
          g.chars[y][1 + x] = '.';
          g.colors[y][1 + x] = '3A4A5E';
        }
      }
    }

    const zbuf = new Array(VIEW_W).fill(Infinity);
    for (let col = 0; col < VIEW_W; col++) {
      const cam = (2 * col) / VIEW_W - 1;
      const rdx = dx + plx * cam, rdy = dy + ply * cam;
      let mx = Math.floor(px), my = Math.floor(py);
      const ddx = Math.abs(1 / (rdx || 1e-9)), ddy = Math.abs(1 / (rdy || 1e-9));
      const stx = rdx < 0 ? -1 : 1, sty = rdy < 0 ? -1 : 1;
      let sx = rdx < 0 ? (px - mx) * ddx : (mx + 1 - px) * ddx;
      let sy = rdy < 0 ? (py - my) * ddy : (my + 1 - py) * ddy;
      let side = 0, hit = '#', guard = 0;
      const maxSteps = W + H + 8; // a ray can cross at most W+H cell borders
      while (guard++ < maxSteps) {
        if (sx < sy) { sx += ddx; mx += stx; side = 0; } else { sy += ddy; my += sty; side = 1; }
        hit = cell(mx, my);
        if (isWall(hit)) break;
      }
      const dist = Math.max(0.05, side === 0 ? sx - ddx : sy - ddy);
      zbuf[col] = dist;
      const h = Math.min(FIELD_ROWS, Math.round((FIELD_ROWS * S) / dist));
      const y0 = FIELD_TOP + Math.floor((FIELD_ROWS - h) / 2);
      const band = dist < 2.4 * S ? 0 : dist < 5.5 * S ? 1 : 2;
      const density =
        (dist < 1.6 * S ? 0 : dist < 3 * S ? 1 : dist < 5 * S ? 2 : dist < 8 * S ? 3 : 4) + side;
      const letter = /[A-Za-z0-9+]/.test(hit);
      const gate = hit === '*';
      const glyph = letter ? hit : gate ? '▒' : SHADE[Math.min(SHADE.length - 1, density)];
      const ink = gate ? '4ADE80' : (letter ? TEAL : SLATE)[band];
      for (let y = y0; y < y0 + h; y++) {
        g.chars[y][1 + col] = glyph;
        g.colors[y][1 + col] = ink;
      }
    }

    // Billboard sprites: § pickups, the * gate once open, and every live
    // enemy (white for a beat when shot).
    const sprites = [];
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const ch = map[y][x];
      if (ch === '§' || (ch === '*' && left === 0)) {
        sprites.push({ x: x + 0.5, y: y + 0.5, ch, ink: ch === '§' ? 'FFD166' : '4ADE80' });
      }
    }
    for (const e of enemies) {
      if (e.hp <= 0) continue;
      const k = ENEMY_KINDS[e.kind];
      sprites.push({
        x: e.x, y: e.y, ch: k.glyph, ink: e.flashT > 0 ? 'FFFFFF' : k.ink,
        enemy: true, kind: e.kind, flash: e.flashT > 0,
      });
    }
    const inv = 1 / (plx * dy - dx * ply);
    sprites.sort((a, b) =>
      ((b.x - px) ** 2 + (b.y - py) ** 2) - ((a.x - px) ** 2 + (a.y - py) ** 2));
    for (const s of sprites) {
      const rx = s.x - px, ry = s.y - py;
      const tx = inv * (dy * rx - dx * ry);
      const ty = inv * (-ply * rx + plx * ry);
      if (ty <= 0.2) continue;
      const screen = Math.floor((VIEW_W / 2) * (1 + tx / ty));
      // Pickups stay small floating tokens; enemies LOOM — twice the width
      // cap, stood on the floor line at their distance (the bottom of the
      // wall slice), with a narrowed top row so a close one reads as a
      // head-and-shoulders silhouette instead of a square.
      const cap = s.enemy ? 12 : 6;
      const size = Math.max(1, Math.min(cap, Math.round((5 * S) / ty)));
      const hgt = s.enemy ? Math.max(1, Math.round(size * 1.3)) : Math.max(1, Math.floor(size * 0.7));
      const wallH = Math.min(FIELD_ROWS, Math.round((FIELD_ROWS * S) / ty));
      const floorLine = FIELD_TOP + Math.floor((FIELD_ROWS + wallH) / 2);
      const y1 = s.enemy
        ? floorLine - hgt
        : FIELD_TOP + Math.floor(FIELD_ROWS / 2) - Math.floor(size / 3);
      const half = Math.floor(size / 2);
      const stamp = s.enemy ? ENEMY_STAMPS[s.kind] : null;
      const sw = stamp ? stamp.rows[0].length : 0;
      const sh = stamp ? stamp.rows.length : 0;
      const wS = half * 2 + 1;
      for (let sxp = screen - half; sxp <= screen + half; sxp++) {
        if (sxp < 0 || sxp >= VIEW_W || ty >= zbuf[sxp]) continue;
        const sxSrc = stamp ? Math.min(sw - 1, Math.floor(((sxp - screen + half) * sw) / wS)) : 0;
        for (let syp = y1; syp < y1 + hgt; syp++) {
          if (syp < FIELD_TOP || syp >= FIELD_TOP + FIELD_ROWS) continue;
          if (stamp && hgt >= 4) {
            // Close enough to have a body: sample the stamp. Transparent
            // cells let the wall show through, so the silhouette is real.
            const ch = stamp.rows[Math.min(sh - 1, Math.floor(((syp - y1) * sh) / hgt))][sxSrc];
            if (ch === ' ') continue;
            g.chars[syp][1 + sxp] = ch;
            g.colors[syp][1 + sxp] = s.flash ? 'FFFFFF' : stamp.inks[ch] ?? s.ink;
          } else {
            g.chars[syp][1 + sxp] = s.ch;
            g.colors[syp][1 + sxp] = s.ink;
          }
        }
      }
    }

    // ── Weapon chrome, drawn in the 3-D view (which the resume-parse never
    // reads, so it can't leak into the level): crosshair, a sidearm wedge at
    // the bottom, a muzzle star while firing, and red view edges when hurt.
    if (combat) {
      const cxv = 1 + (VIEW_W >> 1), cyv = FIELD_TOP + (FIELD_ROWS >> 1);
      if (g.chars[cyv][cxv] === ' ' || muzzleT > 0) {
        g.chars[cyv][cxv] = '+';
        g.colors[cyv][cxv] = muzzleT > 0 ? 'FFF7B0' : '9CB3C9';
      }
      const gy = FIELD_TOP + FIELD_ROWS - 1;
      g.chars[gy][cxv - 1] = '/'; g.colors[gy][cxv - 1] = 'C8D4E2';
      g.chars[gy][cxv] = '█'; g.colors[gy][cxv] = '75879E';
      g.chars[gy][cxv + 1] = '\\'; g.colors[gy][cxv + 1] = 'C8D4E2';
      if (muzzleT > 0) {
        g.chars[gy - 1][cxv] = '*';
        g.colors[gy - 1][cxv] = 'FFF7B0';
      }
      if (hurtT > 0) {
        for (let y = FIELD_TOP; y < FIELD_TOP + FIELD_ROWS; y++) {
          g.chars[y][1] = '░'; g.colors[y][1] = 'FF6B6B';
          g.chars[y][VIEW_W] = '░'; g.colors[y][VIEW_W] = 'FF6B6B';
        }
      }
    }

    // ── The MAP band: raw, typeable characters — this IS the level source.
    // A level bigger than the band scrolls it as a window over the world.
    const [wx, wy] = winPos();
    const windowed = W > WIN_W || H > WIN_H;
    writeText(g, FIELD_TOP, BAND_X + 1,
      windowed ? `MAP @${wx},${wy} · edit!` : 'MAP · edit me!', '5EEAD4');
    // One ink for '.'/'#' (the glyphs already tell them apart) keeps a map
    // row at ~1 run; only the payload cells spend color.
    for (let y = 0; y < WIN_H; y++) {
      const gy = MAP_TOP + y;
      for (let x = 0; x < WIN_W; x++) {
        const ch = cell(wx + x, wy + y);
        g.chars[gy][BAND_X + 1 + x] = ch;
        g.colors[gy][BAND_X + 1 + x] =
          ch === '§' ? 'FFD166' : ch === '*' ? '4ADE80'
            : ch === '.' || ch === '#' ? '46556B' : '5EEAD4';
      }
    }
    // The camera's view cone, cast onto the map with line-of-sight: bright
    // dots on exactly the floor the 3-D view is showing. This is the seam
    // that makes the map and the corridor read as one world — as you turn,
    // the lit wedge sweeps with the render.
    const FCOS = 0.83; // cos(FOV/2)
    for (let y = 0; y < WIN_H; y++) for (let x = 0; x < WIN_W; x++) {
      const mx2 = wx + x, my2 = wy + y;
      if (cell(mx2, my2) !== '.') continue;
      const vx = mx2 + 0.5 - px, vy = my2 + 0.5 - py;
      const d = Math.hypot(vx, vy);
      if (d < 0.4 || d > pack.coneRadius) continue;
      if ((vx * dx + vy * dy) / d < FCOS) continue;
      let lit = true;
      const steps = Math.ceil(d * 3);
      for (let s = 1; s < steps; s++) {
        const t = s / steps;
        const cx2 = Math.floor(px + vx * t), cy2 = Math.floor(py + vy * t);
        if (cx2 === mx2 && cy2 === my2) break;
        if (isWall(cell(cx2, cy2))) { lit = false; break; }
      }
      if (!lit) continue;
      g.chars[MAP_TOP + y][BAND_X + 1 + x] = '·';
      g.colors[MAP_TOP + y][BAND_X + 1 + x] = 'C8D4E2';
    }
    // Live enemies overlay the band as their glyphs — entities on top of
    // tiles, exactly how the resume-parse reads them back.
    for (const e of enemies) {
      if (e.hp <= 0) continue;
      const ex = Math.floor(e.x) - wx, ey = Math.floor(e.y) - wy;
      if (ex >= 0 && ex < WIN_W && ey >= 0 && ey < WIN_H) {
        g.chars[MAP_TOP + ey][BAND_X + 1 + ex] = ENEMY_KINDS[e.kind].glyph;
        g.colors[MAP_TOP + ey][BAND_X + 1 + ex] = ENEMY_KINDS[e.kind].ink;
      }
    }
    // Directional player marker — the map's compass for the 3-D camera.
    const pmx = Math.floor(px) - wx, pmy = Math.floor(py) - wy;
    if (pmx >= 0 && pmx < WIN_W && pmy >= 0 && pmy < WIN_H) {
      g.chars[MAP_TOP + pmy][BAND_X + 1 + pmx] =
        Math.abs(dx) > Math.abs(dy) ? (dx > 0 ? '►' : '◄') : (dy > 0 ? '▼' : '▲');
      g.colors[MAP_TOP + pmy][BAND_X + 1 + pmx] = 'FF6B6B';
    }
    writeText(g, MAP_TOP + WIN_H + 1, BAND_X + 1, 'letters become walls', '46556B');
    writeText(g, MAP_TOP + WIN_H + 2, BAND_X + 1,
      killsTotal > 0 ? '$ heals  & = baddie' : '$ = treasure  @ = you', '46556B');

    if (state === 'won') drawBanner(g, pack.winBanner);
    else if (state === 'dead') drawBanner(g, ['YOU DIED', 'the document respawns you at the start']);
    return { grid: g, bg: pack.bg };
  }

  /** Rebuild the world from the MAP band of the parsed document rows. Typed
   *  letters become walls; a moved '@' teleports the player. The band shows
   *  the winPos() window, so edits land on exactly the world cells it shows —
   *  the rest of a big level is untouched. */
  function syncFromRows(rows) {
    if (state !== 'run') return; // a banner is on screen, not the maze
    const [wx, wy] = winPos();
    const parsed = map.map((r) => r.slice());
    let atX = null, atY = null;
    // Enemies inside the window are re-seeded from what the text says — same
    // entity contract as the platformer's gremlins. Anything outside the
    // window is beyond edit reach and survives untouched.
    const inWindow = (x, y) => x >= wx && x < wx + WIN_W && y >= wy && y < wy + WIN_H;
    enemies = enemies.filter((e) =>
      e.hp > 0 && !inWindow(Math.floor(e.x), Math.floor(e.y)));
    for (let y = 0; y < WIN_H; y++) {
      const line = rows[MAP_TOP + y];
      if (line == null) continue;
      const div = line.indexOf('│', 1);
      if (div < 0) continue;
      const band = line.slice(div + 2, div + 2 + WIN_W); // skip │ + pad col
      for (let x = 0; x < WIN_W; x++) {
        let ch = band[x] ?? parsed[wy + y][wx + x];
        // '@' is the typeable teleport; ►◄▲▼ is how the renderer draws the
        // player, and '·' is the rendered view cone — all parse back to
        // position-or-floor, never to walls. An enemy glyph is an entity:
        // it respawns that enemy and leaves floor beneath (typing '&' — or
        // any of the bestiary's letters — conjures one from the document).
        if (ch === '@' || ch === '►' || ch === '◄' || ch === '▲' || ch === '▼') {
          atX = wx + x; atY = wy + y; ch = '.';
        } else if (ch === '·') {
          ch = '.';
        } else if (ENEMY_BY_GLYPH[ch]) {
          spawnEnemy(wx + x + 0.5, wy + y + 0.5, ENEMY_BY_GLYPH[ch]);
          ch = '.';
        }
        parsed[wy + y][wx + x] = ch;
      }
    }
    killsTotal = Math.max(killsTotal, kills + enemies.filter((e) => e.hp > 0).length);
    normalizeMap(parsed.map((r) => r.join('')));
    if (atX != null && (atX !== Math.floor(px) || atY !== Math.floor(py))) {
      px = atX + 0.5; py = atY + 0.5;
    }
    unstickPlayer();
  }

  return {
    name: pack.name,
    label: pack.label,
    caption: pack.caption,
    hint: pack.hint,
    reset,
    tick: tickSim,
    render,
    syncFromRows,
    state: () => ({
      player: { x: px, y: py, dx, dy },
      sigilsLeft: sigilsLeft(), mode: state,
      window: winPos(),
      health, kills, killsTotal,
      enemies: enemies.filter((e) => e.hp > 0)
        .map((e) => ({ x: e.x, y: e.y, kind: e.kind, hp: e.hp, awake: e.awake })),
      mapRow: (y) => (map[y] ?? []).join(''),
    }),
  };
}

/** Exported for the Playwright spec and headless logic checks. */
export function dungeonCart() { return raycastCart(DUNGEON_PACK); }

/** Cartridge 3 — a REAL Doom-format level: Freedoom's E1M1 (BSD-licensed),
 *  rasterized to the raycaster's grid by docs/demo/tools/wad2cart.mjs. Same
 *  controls, same renderer, same document round-trip — only the level pack
 *  differs. */
export function freedoomCart() { return raycastCart(FREEDOOM_PACK); }

// ─── Frame → document plumbing ────────────────────────────────────────

/** Decode the screen paragraph's XML back into text rows: `w:t` text joined
 *  within a row, `w:br` starting the next. The inverse both of what frameXml
 *  emits and of what the editor's markdown blur-commit leaves behind. */
export function rowsFromXml(xml) {
  const body = xml.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '');
  const rows = [''];
  const re = /<w:br(?:\s[^>]*)?\/?>|<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(body)) !== null) {
    if (m[1] === undefined) rows.push('');
    else rows[rows.length - 1] += m[1]
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
      .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(Number(d)))
      .replace(/&amp;/g, '&');
  }
  return rows;
}

/** Seed a freshly opened blank session with the Arcade document — title, game
 *  screen, caption, and a real footnote — entirely through the agentic editing
 *  surface, then capture the screen paragraph's opening tag (it carries the
 *  Unid, THE thing that keeps the anchor stable across frames). */
export function seedArcade(session) {
  const check = (r, what) => {
    if (!r.success) throw new Error(`${what} failed: ${r.error?.code} ${r.error?.message}`);
    return r;
  };
  const firstP = session.findByKind('p', 'body')[0];
  if (!firstP) throw new Error('blank document has no body paragraph');
  const titleAnchor = firstP.id;
  check(session.replaceText(titleAnchor, 'THE DOCX ARCADE'), 'title replaceText');
  check(session.setParagraphFormat(titleAnchor, { alignment: 'center', spacingAfter: 160 }), 'title format');
  check(session.applyFormat(titleAnchor, null, { bold: true, fontFamily: 'Courier New', fontSizePts: 13, color: '1F2937' }), 'title run format');

  const canvasResult = check(session.insertParagraph(titleAnchor, 'after', '(inserting coin…)'), 'screen insert');
  const canvasAnchor = canvasResult.created[0].id;

  const captionResult = check(session.insertParagraph(canvasAnchor, 'after', 'loading cartridge…'), 'caption insert');
  const captionAnchor = captionResult.created[0].id;
  check(session.setParagraphFormat(captionAnchor, { alignment: 'center', spacingBefore: 160 }), 'caption format');
  check(session.applyFormat(captionAnchor, null, { fontFamily: 'Courier New', fontSizePts: 8, color: '6B7280' }), 'caption run format');

  // A real footnote, because the game screen is a real document.
  check(session.insertFootnote(captionAnchor, 7, // after "loading"
    'Every frame of the game is OOXML: colored runs and `w:br` breaks in one Word paragraph, ' +
    'swapped in by `DocxSession.raw.replaceXml` and repainted incrementally by `DocxEditor.refresh()`. ' +
    'Pause mid-jump and Save: the frame downloads as a real .docx.'), 'footnote');

  const seedXml = session.raw.getXml(canvasAnchor);
  const gt = seedXml.indexOf('>');
  let openTag = seedXml.slice(0, gt + 1);
  if (openTag.endsWith('/>')) openTag = openTag.slice(0, -2) + '>';
  return { titleAnchor, canvasAnchor, captionAnchor, openTag };
}

// ─── The editor-hosted driver ─────────────────────────────────────────

/**
 * Seed the Arcade into a ribbon-hosted editor's session and run the game loop
 * against it. Owns the dock (cartridge switch, pause/resume, restart, pace,
 * telemetry) and the keyboard: game keys are claimed only while playing.
 * Clicking the document pauses — the frame you clicked is now just a
 * paragraph with your caret in it. Resuming blurs the edit (the editor
 * commits on blur), re-parses the game world from the session's XML, and
 * hands the keyboard back to the game.
 *
 * `ui`: { carts, playpause, restart, pace, stats, hint, pad? } — dock DOM.
 * Returns the controller the host page publishes as `window.__arcade`.
 */
export function startArcade({ editor, session, ui, cart: startCart }) {
  if (typeof editor.refresh !== 'function') {
    throw new Error('This engine predates DocxEditor.refresh() — the Arcade needs docxodus ≥ 9.6.0.');
  }
  const seeded = seedArcade(session);
  let canvasAnchor = seeded.canvasAnchor;
  let openTag = seeded.openTag;

  const carts = [platformerCart(), dungeonCart(), freedoomCart()];
  let cart = carts.find((c) => c.name === startCart) ?? carts[0];

  let playing = false;
  let timer = 0;
  let lastWall = performance.now();
  let frames = 0;
  let fps = 0;
  let lastRuns = 0;
  let lastFrameEnd = performance.now();
  const timings = { mutate: 0, refresh: 0 };
  let interval = Number(ui.pace.value);

  const input = createInput(() => playing);

  const unidOf = (anchor) => anchor.split(':')[2];
  const canvasEl = () => editor.root.querySelector(`[data-anchor="${unidOf(canvasAnchor)}"]`);

  function setCaption() {
    session.replaceText(seeded.captionAnchor, cart.caption);
    ui.hint.innerHTML = cart.hint +
      ' · <b>Esc</b> pauses/resumes · <b>Undo</b> rewinds frames · <b>Save</b> ships the frame as .docx';
  }

  function drawFrame() {
    const { grid, bg } = cart.render();
    const { xml, runs } = frameXml(openTag, grid, bg);
    lastRuns = runs;
    const t0 = performance.now();
    const res = session.raw.replaceXml(canvasAnchor, xml);
    const t1 = performance.now();
    if (!res.success) throw new Error(`replaceXml: ${res.error?.code} ${res.error?.message}`);
    canvasAnchor = res.modified[0]?.id ?? res.created[0]?.id ?? canvasAnchor;
    editor.refresh();
    const t2 = performance.now();

    const mix = (a, b) => (a === 0 ? b : a * 0.9 + b * 0.1);
    timings.mutate = mix(timings.mutate, t1 - t0);
    timings.refresh = mix(timings.refresh, t2 - t1);
    fps = mix(fps, 1000 / Math.max(1, t2 - lastFrameEnd));
    lastFrameEnd = t2;
    frames++;

    const fb = editor.lastReconcileFallback;
    ui.stats.innerHTML =
      `<b>${cart.label}</b> · frame <b>${frames}</b> · <b>${fps.toFixed(1)}</b> fps · ` +
      `replaceXml <b>${timings.mutate.toFixed(1)}</b> ms · refresh <b>${timings.refresh.toFixed(1)}</b> ms · ` +
      `<b>${lastRuns}</b> runs · ` +
      (fb ? `remounted (${fb})` : `<span class="inc">incremental — one block repainted</span>`);
  }

  function loop() {
    if (!playing) return;
    const started = performance.now();
    const dt = Math.min(0.2, (started - lastWall) / 1000);
    lastWall = started;
    try {
      cart.tick(dt, input);
      input.endTick();
      drawFrame();
    } catch (e) {
      playing = false;
      ui.stats.textContent = 'halted: ' + e.message;
      throw e;
    }
    timer = setTimeout(loop, Math.max(0, interval - (performance.now() - started)));
  }

  /** Re-read the game world from the document. Called on every resume: the
   *  editor commits block edits on blur, so blur first, then parse the
   *  SESSION's XML (authoritative), then let the cartridge merge it in. */
  function syncFromDocument() {
    if (document.activeElement instanceof HTMLElement) document.activeElement.blur();
    let xml = '';
    try { xml = session.raw.getXml(canvasAnchor) ?? ''; } catch { xml = ''; }
    if (!xml) {
      // The screen paragraph itself was deleted while paused — rebuild it.
      const created = session.insertParagraph(seeded.titleAnchor, 'after', '(re-inserting coin…)');
      canvasAnchor = created.created[0].id;
      xml = session.raw.getXml(canvasAnchor);
    }
    const gt = xml.indexOf('>');
    let tag = xml.slice(0, gt + 1);
    if (tag.endsWith('/>')) tag = tag.slice(0, -2) + '>';
    openTag = tag;
    cart.syncFromRows(rowsFromXml(xml));
    // Sweep paragraphs an Enter-split stranded between screen and caption, so
    // pausing to edit can never slowly litter the document.
    const ids = session.findByKind('p', 'body').map((r) => r.id);
    const from = ids.indexOf(canvasAnchor), to = ids.indexOf(seeded.captionAnchor);
    if (from >= 0 && to > from + 1) {
      for (const id of ids.slice(from + 1, to)) session.deleteBlock(id);
    }
  }

  function setPlaying(next) {
    if (playing === next) return;
    if (next) {
      syncFromDocument();
      playing = true;
      ui.playpause.textContent = '⏸ Pause & edit';
      lastWall = performance.now();
      lastFrameEnd = performance.now();
      loop();
    } else {
      playing = false;
      clearTimeout(timer);
      ui.playpause.textContent = '▶ Resume';
      ui.stats.innerHTML = 'paused — the screen is an ordinary paragraph now. ' +
        '<b>Type into it</b>, then Resume to make it real.';
    }
  }

  const cartBtns = new Map();
  for (const c of carts) {
    const b = document.createElement('button');
    b.textContent = c.label;
    b.setAttribute('aria-pressed', String(c === cart));
    b.addEventListener('click', () => setCart(c.name));
    cartBtns.set(c.name, b);
    ui.carts.appendChild(b);
  }
  function setCart(name) {
    const next = carts.find((c) => c.name === name);
    if (!next) return;
    cart = next;
    cart.reset();
    cartBtns.forEach((b, n) => b.setAttribute('aria-pressed', String(n === name)));
    setCaption();
    if (!playing) {
      // Paint the new cartridge's own frame BEFORE resuming, so the resume
      // parse reads this cartridge's screen, never the other one's.
      drawFrame();
      setPlaying(true);
    }
  }

  ui.playpause.addEventListener('click', () => setPlaying(!playing));
  ui.restart.addEventListener('click', () => {
    cart.reset();
    if (!playing) { drawFrame(); setPlaying(true); }
  });
  ui.pace.addEventListener('change', () => { interval = Number(ui.pace.value); });

  // Esc toggles play/pause from anywhere — including from inside the document.
  window.addEventListener('keydown', (e) => {
    if (e.code !== 'Escape') return;
    e.preventDefault();
    setPlaying(!playing);
  }, true);

  // Click the game (or any block) while it plays: the frame freezes and the
  // caret lands in it. No mode switch — it was a document the whole time.
  editor.root.addEventListener('pointerdown', () => setPlaying(false), true);

  // On-screen pad (touch): buttons carry data-code="ArrowLeft" etc.
  if (ui.pad) {
    ui.pad.querySelectorAll('[data-code]').forEach((btn) => {
      const code = btn.getAttribute('data-code');
      const press = (e) => { e.preventDefault(); if (!playing) setPlaying(true); input.set(code, true); };
      const release = (e) => { e.preventDefault(); input.set(code, false); };
      btn.addEventListener('pointerdown', press);
      btn.addEventListener('pointerup', release);
      btn.addEventListener('pointercancel', release);
      btn.addEventListener('pointerleave', release);
    });
  }

  setCaption();
  setPlaying(true);

  return {
    canvasAnchor: () => canvasAnchor,
    canvasText: () => canvasEl()?.textContent ?? '',
    canvasElement: () => canvasEl(),
    frames: () => frames,
    fps: () => fps,
    timings: () => ({ ...timings, runs: lastRuns }),
    cart: () => cart.name,
    setCart,
    game: () => cart.state(),
    playing: () => playing,
    pause: () => setPlaying(false),
    resume: () => setPlaying(true),
    input,
    save: () => editor.save(),
  };
}

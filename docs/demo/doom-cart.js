// ═══════════════════════════════════════════════════════════════════════
// Cartridge 3 — DOOM. The actual game, not an impression of it.
//
// LICENSE NOTE — THIS FILE IS GPL-2.0-or-later, NOT MIT.
// Docxodus is MIT (see the root LICENSE) and every other file in this
// directory stays MIT. This one is different: it is written against, and at
// runtime combined with, doomgeneric — a build of id Software's Doom source,
// which id released under the GNU General Public License v2. So this glue is
// offered under GPL-2.0-or-later too. The engine itself is not in this
// repository: it is a pinned jsDelivr URL loaded through a dynamic `import()`,
// which is also why the 3 MB build never downloads for a visitor who plays the
// other two cartridges. See vendor/NOTICE.md.
//
// WHAT THIS REPLACES
// ------------------
// Cartridge 3 used to be a hand-written ASCII raycaster fed by Freedoom's
// E1M1 geometry, rasterized to a character grid offline. That was a real Doom
// *level* in a Word document. This is the real Doom *engine* in a Word
// document: id's own BSP renderer, its own 320×200 framebuffer, its own
// physics, monsters, doors, weapons, menus and status bar — every frame of it
// pushed through `DocxSession.raw.replaceXml` into a single Word paragraph.
//
// HOW A 320×200 FRAMEBUFFER FITS IN A PARAGRAPH
// ---------------------------------------------
// The screen is the same 92×26 character grid the rest of the arcade draws on
// (see ascii-scenes.js), so it inherits the pinned canvas font and the
// markdown-safe bezel unchanged. A character cell is far too coarse to hold a
// Doom frame at one color per cell, so each cell holds TWO pixels: the glyph
// is `▀` (upper half block), its ink is the top pixel and its run shading
// (`w:shd` in `w:rPr`) is the bottom one. That buys 64 × 46 pixels inside the
// bezel — still tiny, but enough that E1M1's courtyard, the shotgun guy in
// front of you and the face on the status bar all read.
//
// Where both pixels agree — most of a Doom frame, which is large flat spans
// of wall and floor — the glyph degrades to a plain `█` and the two colors
// collapse to one, so the whole span merges into a single run. That run
// merging, not the glyph, is what the per-frame cost is actually made of.
//
// The viewport is 64 columns wide because that is 4:3 at this cell's aspect:
// a cell advances 4.8pt and a half-cell is 5pt tall, so 64 × 46 half-pixels
// is 307.2pt × 230pt — Doom's own picture shape, not a stretched one.
// ═══════════════════════════════════════════════════════════════════════

import { COLS, ROWS } from './ascii-scenes.js';

// ─── Doom's framebuffer, and where it lands on the grid ───────────────
const DOOM_W = 320, DOOM_H = 200;

const FIELD_TOP = 2;                      // 0 = bezel, 1 = HUD
const FIELD_ROWS = ROWS - FIELD_TOP - 1;  // 23 rows of picture
const VIEW_W = 64;                        // 4:3 at this cell aspect
const DIV_X = 1 + VIEW_W;                 // the │ divider column
const PANEL_X = DIV_X + 1;                // first column of the side panel
const PANEL_W = COLS - 1 - PANEL_X;       // 25 columns

/** Two half-pixels per cell row. */
const PIX_H = FIELD_ROWS * 2;             // 46

const BEZEL_INK = '33465B';
const HUD_INK = '9CB3C9';
const PANEL_INK = '8FA3B8';
const PANEL_HI = 'FFD166';
const DEAD_INK = '5D6975';
const BG = '000000';

// Nearest-neighbour sample offsets, precomputed once: Doom's own palette
// survives the downsample intact this way, which matters more than it
// sounds. Averaging would invent thousands of in-between colors, and every
// distinct color is a run boundary — the picture would cost three times the
// XML to say the same thing, and look muddier saying it.
const SX = Array.from({ length: VIEW_W }, (_, x) => Math.floor((x + 0.5) * DOOM_W / VIEW_W));
const SY = Array.from({ length: PIX_H }, (_, y) => Math.floor((y + 0.5) * DOOM_H / PIX_H));

const HEX = Array.from({ length: 256 }, (_, i) => i.toString(16).toUpperCase().padStart(2, '0'));

// ─── Keyboard: arcade key codes → Doom key codes ──────────────────────
// From doomkeys.h. Doom wants a byte per key plus a pressed flag, delivered
// through DG_GetKey one transition at a time, so the cart keeps a queue and
// drains the arcade input's transition log into it.
const KEY = {
  RIGHTARROW: 0xae, LEFTARROW: 0xac, UPARROW: 0xad, DOWNARROW: 0xaf,
  STRAFE_L: 0xa0, STRAFE_R: 0xa1, USE: 0xa2, FIRE: 0xa3,
  ESCAPE: 27, ENTER: 13, TAB: 9, RSHIFT: 0x80 + 0x36,
};

// Esc is NOT in this table on purpose: the arcade owns Esc (it pauses the
// game and hands the paragraph back to the editor), so Doom's own menu is on
// Q instead. Everything else is where a Doom player's hands expect it.
const KEY_MAP = {
  KeyW: KEY.UPARROW, ArrowUp: KEY.UPARROW,
  KeyS: KEY.DOWNARROW, ArrowDown: KEY.DOWNARROW,
  KeyA: KEY.STRAFE_L, KeyD: KEY.STRAFE_R,
  ArrowLeft: KEY.LEFTARROW, ArrowRight: KEY.RIGHTARROW,
  Space: KEY.FIRE,            // the dock's round FIRE button sends Space
  KeyE: KEY.USE,              // doors and switches
  ShiftLeft: KEY.RSHIFT, ShiftRight: KEY.RSHIFT,
  Enter: KEY.ENTER,
  KeyQ: KEY.ESCAPE,           // Doom's own menu
  KeyM: KEY.TAB,              // the automap (Tab moves focus in a browser)
  Digit1: 0x31, Digit2: 0x32, Digit3: 0x33, Digit4: 0x34,
  Digit5: 0x35, Digit6: 0x36, Digit7: 0x37,
};

/** Switches the projection rather than reaching Doom — see PROJECTIONS. */
export const PROJECTION_KEY = 'KeyP';

/** The arcade key codes this cartridge wants to be handed. ascii-arcade.js
 *  unions this into the set its capture-phase listener claims while playing,
 *  so Doom gets Enter/E/Q/M/digits without the other cartridges caring.
 *  PROJECTION_KEY is claimed too, so the browser never sees it, but it is
 *  handled here and deliberately not in KEY_MAP: Doom must not receive it. */
export const DOOM_KEY_CODES = [...Object.keys(KEY_MAP), PROJECTION_KEY];

// ─── Grid helpers ─────────────────────────────────────────────────────
function makeGrid() {
  const chars = [], colors = [], bgs = [];
  for (let y = 0; y < ROWS; y++) {
    chars.push(new Array(COLS).fill(' '));
    colors.push(new Array(COLS).fill(HUD_INK));
    bgs.push(new Array(COLS).fill(null));
  }
  return { chars, colors, bgs };
}

function write(g, y, x, text, ink) {
  for (let k = 0; k < text.length && x + k < COLS - 1; k++) {
    if (y < 0 || y >= ROWS || x + k < 0) continue;
    g.chars[y][x + k] = text[k];
    g.colors[y][x + k] = ink;
  }
}

/** The same load-bearing bezel every cartridge wears: each row starts with a
 *  box-drawing character, so the editor's markdown blur-commit can never read
 *  a row as a heading or a bullet, and no row is ever whitespace-only (a
 *  blank line would split the screen paragraph in two). */
function drawChrome(g, hud) {
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
  write(g, 1, 2, hud.slice(0, COLS - 4), HUD_INK);
}

function drawDivider(g) {
  for (let y = FIELD_TOP; y < ROWS - 1; y++) {
    g.chars[y][DIV_X] = '│';
    g.colors[y][DIV_X] = BEZEL_INK;
  }
}

/** Wrap text into the side panel's narrow column. */
function wrap(text, width) {
  const out = [];
  let line = '';
  for (const word of text.split(' ')) {
    if (line && line.length + 1 + word.length > width) { out.push(line); line = word; }
    else line = line ? `${line} ${word}` : word;
  }
  if (line) out.push(line);
  return out;
}

// ─── Framebuffer → grid ───────────────────────────────────────────────
// A run breaks when the ink or the shading changes, and this projection
// changes both on almost every cell: a photographic downsample of Doom agrees
// with its neighbour about 1.02 cells in a row, so nearly every cell is its
// own run and a frame lands near 1,370 of them. That is the entire frame
// budget — the OOXML→HTML conversion of this paragraph is linear in runs.
//
// So neighbouring cells that are nearly the same colour should be exactly the
// same colour, and merge. An earlier attempt at that quantised each channel
// and snapped the top and bottom pixels INDEPENDENTLY, barely moved the run
// count, and led me to call ~1 fps a property of the medium. It is not. It
// was a property of that rule: a run needs BOTH halves to match, so deciding
// the halves separately makes the merge probability a product of two chances
// and almost guarantees it fails. Decide the PAIR jointly — adopt both of the
// previous cell's colours or neither — and the same tolerance that did
// nothing before takes a frame from ~1,370 runs to ~420.
//
// The comparison is deliberately not RGB distance. The eye resolves green far
// better than blue, so an even tolerance spends most of its budget where it
// is least visible; weighting by luma puts the error where it will not be
// seen. And the previous cell's EMITTED colour is what a candidate is
// compared against, not the previous cell's true colour, so the tolerance
// bounds accumulated drift along a row rather than each step of it — a long
// gradient still ends up the colour it should be.
const PAIR_SNAP_TOLERANCE = 40 * 100;   // luma-weighted, summed over both halves
const LUMA_R = 30, LUMA_G = 59, LUMA_B = 11;

/** Paint one 320×200 BGRA frame into the grid's viewport as half-block cells.
 *
 *  Exported (and pure) so the headless logic tests can drive it with a
 *  synthetic framebuffer: this is the one piece of the cartridge that has
 *  interesting behaviour and does not need a WebAssembly Doom to exercise. */
export function paintFramebuffer(g, fb) {
  for (let row = 0; row < FIELD_ROWS; row++) {
    const gy = FIELD_TOP + row;
    const topBase = SY[row * 2] * DOOM_W;
    const botBase = SY[row * 2 + 1] * DOOM_W;
    // Runs are built per row, so the snap chain starts fresh on each one.
    let ptr = -1, ptg = 0, ptb = 0, pbr = 0, pbg = 0, pbb = 0;
    for (let x = 0; x < VIEW_W; x++) {
      const sx = SX[x];
      // BGRA in memory: blue first, alpha always 0 — hence the explicit swap
      // rather than a straight copy.
      const t = (topBase + sx) * 4, b = (botBase + sx) * 4;
      let tr = fb[t + 2], tg = fb[t + 1], tb = fb[t];
      let br = fb[b + 2], bg2 = fb[b + 1], bb = fb[b];

      if (ptr >= 0) {
        const d = LUMA_R * Math.abs(ptr - tr) + LUMA_G * Math.abs(ptg - tg) + LUMA_B * Math.abs(ptb - tb)
                + LUMA_R * Math.abs(pbr - br) + LUMA_G * Math.abs(pbg - bg2) + LUMA_B * Math.abs(pbb - bb);
        if (d <= PAIR_SNAP_TOLERANCE) { tr = ptr; tg = ptg; tb = ptb; br = pbr; bg2 = pbg; bb = pbb; }
      }
      ptr = tr; ptg = tg; ptb = tb; pbr = br; pbg = bg2; pbb = bb;

      const gx = 1 + x;
      const top = HEX[tr] + HEX[tg] + HEX[tb];
      const bottom = (tr === br && tg === bg2 && tb === bb)
        ? top : HEX[br] + HEX[bg2] + HEX[bb];
      // Every picture cell carries shading, even where both halves agree and
      // the glyph is a plain full block. That is deliberate: a block GLYPH
      // only covers the font's content box, which is shorter than the exact
      // 10pt line the canvas pins, so ink alone leaves a hairline of
      // paragraph fill above and below every row — black scan lines straight
      // across the picture. The run's background is what actually fills the
      // line box (the canvas pin pads it to), so the background is the cell's
      // real color and the glyph only ever adds the top half on top of it.
      // Equal halves still merge into one run, so this costs run COUNT
      // nothing; it costs one w:shd element per run.
      g.chars[gy][gx] = bottom === top ? '█' : '▀';
      g.colors[gy][gx] = top;
      g.bgs[gy][gx] = bottom;
    }
  }
}

// ─── Framebuffer → grid, the ASCII way ────────────────────────────────
// The bitmap painter above treats the paragraph as a screen. That is the
// faithful reading and it is also the expensive one: a photographic
// downsample of Doom agrees with its neighbour about 1.07 cells in a row, so
// nearly every cell starts a new run, and a frame landed near 1,370 runs.
//
// The cost of a frame is the OOXML→HTML conversion of that paragraph, and it
// is LINEAR in runs (measured: ~35 ms fixed + ~0.70 ms per run across a 24×
// range). So runs are the dial. And the thing that turns it is this:
//
//   a run breaks only when INK or SHADING changes — the glyph does not
//   break it, because every character in a span shares one w:t.
//
// So a cell has two channels with completely different prices. The glyph is
// free and the ink is not, and the job is to route each part of the picture
// to the channel that carries it best.
//
// THE FIRST ATTEMPT AT THAT WAS WRONG, and it is worth writing down why.
// It sent luminance to the glyph and kept the ink at one grey, which drove a
// frame to 23 picture runs — and produced something unreadable. Measuring the
// framebuffer says exactly what went wrong:
//
//   * Doom is dark AND its exposure moves. Median luminance is 0.110 in a
//     corridor and 0.185 in an open lit area, with the 90th percentile at
//     0.19 and 0.36. One fixed tone curve cannot serve both: the curve that
//     was there clipped 14% of a corridor to black and 7.5% of a lit area to
//     white, and squeezed everything left into two of its six ramp steps.
//   * Luminance alone does not separate a Doom frame anyway. 39% of the lit
//     area is a single hue — E1M1's brick and wood are all the same brown —
//     so a greyscale projection of it is a grey rectangle with dithering on.
//
// Hence three changes, in order of what they bought:
//
//   1. EXPOSE PER FRAME. Fit the curve to this frame's own 6th and 96th
//      percentiles instead of to a constant. Free — no run costs anything —
//      and on its own it is the difference between mush and architecture.
//   2. PUT HUE BACK, COARSELY. Seven families, chosen so the ink says WHAT a
//      surface is (brick, sky, blood, a key) while the glyph says how bright
//      it is. Hue is spatially coherent in a way exact colour is not, so a
//      whole wall is one family and therefore one run.
//   3. SPEND THE INK ON THREE BRIGHTNESS TIERS. Five ramp glyphs is not
//      enough tonal range; three tiers × the ramp is a nine-rung ladder,
//      ordered by ink value × the glyph's fill fraction so that it is
//      actually monotone in brightness.
//
// (2) and (3) do cost runs — but most of what they were costing was texture
// noise flipping neighbouring cells between two palette entries nobody can
// tell apart at 6.4 pixels per cell. Snapping a cell back to its neighbour's
// ink when the two are that close took a frame from 431 runs to ~216 with no
// visible change; see SNAP_NEAR for the one direction of that substitution
// which turned out not to be free.
//
// Measured over three captured E1M1 frames the picture is ~216 runs, and in
// the live editor, walking and turning through E1M1, the whole paragraph is
// ~470 runs at 2.7 fps.
//
// That used to be 3.7× the bitmap projection, which is what justified this
// mode existing. It is now about 1.5×, because the bitmap painter learned the
// same trick — see PAIR_SNAP_TOLERANCE. On some frames this one is the more
// expensive of the two. What still separates them is the palette: this one is
// closed at 21 inks and degrades by flattening colour, the other is open and
// degrades by smearing horizontally.
const RAMP = [' ', '░', '▒', '▓', '█'];
/** How much of a cell each ramp glyph inks — the other half of "how bright
 *  does this cell look", the half the ink does not carry. */
const RAMP_FILL = [0, 0.25, 0.5, 0.75, 1];
const EDGE_SPLIT = 0.38;   // half-tone gap that earns a ▀ / ▄ instead of a ramp glyph
const EDGE_FLOOR = 0.12;   // …but not in near-black, where it is only noise

/** Ink brightness tiers. Three, because two is visibly banded and four costs
 *  runs for a rung the eye does not resolve at this cell size. */
const TIERS = [0.40, 0.68, 1];

/** The tone ladder: every (tier, glyph) pair, ordered by what it actually
 *  looks like — ink value times fill fraction — then thinned so each rung is
 *  visibly brighter than the one below. Where two rungs tie, the dimmer INK
 *  wins, because that is the one more likely to match the cell to its left. */
const LADDER_MIN_STEP = 0.055;
const LADDER = (() => {
  const all = [];
  for (let t = 0; t < TIERS.length; t++) {
    for (let r = 0; r < RAMP.length; r++) all.push({ tier: t, glyph: r, v: TIERS[t] * RAMP_FILL[r] });
  }
  all.sort((a, b) => a.v - b.v || a.tier - b.tier);
  const out = [];
  for (const rung of all) {
    const prev = out[out.length - 1];
    if (!prev) { out.push(rung); continue; }
    if (rung.v - prev.v < LADDER_MIN_STEP) {
      if (rung.tier < prev.tier) out[out.length - 1] = rung;
      continue;
    }
    out.push(rung);
  }
  return out;
})();

/** Hue families, at full tier. Named for what they are in Doom, because that
 *  is what decides whether a family earns its place: each one has to be a
 *  thing a player needs to tell apart from the others at a glance. */
const FAMILIES = [
  [0.82, 0.84, 0.88],   // grey    — concrete, tech panels, the status bar
  [1.00, 0.34, 0.30],   // red     — blood, damage, the health numerals
  [1.00, 0.66, 0.32],   // orange  — brick, wood, floor: 39% of an E1M1 frame
  [0.98, 0.94, 0.40],   // yellow  — lights, ammo, the yellow key
  [0.42, 0.92, 0.44],   // green   — armour, slime, the green key
  [0.46, 0.62, 1.00],   // blue    — sky, the blue key
  [0.94, 0.46, 0.94],   // magenta
];
/** Below this saturation a cell is grey. Low enough that the brown holds —
 *  brown is what the level is made of — and high enough that concrete and
 *  metal do not pick up a tint they do not have. */
const SAT_GATE = 0.26;
/** Index of the grey family in FAMILIES — the only one a cell may be pulled
 *  into from another, because losing a tint is a smaller lie than gaining one. */
const GREY = 0;

/** The whole palette: one ink per (family, tier). 21 of them, and a picture
 *  run boundary can only ever be a move between two of these. */
const INK = [];
const INK_RGB = [];
for (const fam of FAMILIES) {
  for (const tier of TIERS) {
    const rgb = fam.map((v) => Math.min(255, Math.round(v * tier * 255)));
    INK_RGB.push(rgb);
    INK.push(HEX[rgb[0]] + HEX[rgb[1]] + HEX[rgb[2]]);
  }
}

/** Which palette entries are close enough to substitute for each other.
 *  This is the hysteresis: texture noise makes neighbouring cells flip
 *  between two nearby inks, and every flip buys a run boundary for a
 *  difference invisible at 6.4 pixels per cell. Precomputed as a table so the
 *  per-cell test is one array lookup.
 *
 *  It is applied twice per cell, in this order: match the cell to the LEFT
 *  first, because that is the substitution that actually saves a run; failing
 *  that, match the cell ABOVE, which saves nothing but stops each row from
 *  settling on its own ink independently and striping the picture. */
const SNAP_TOLERANCE = 160;
const SNAP_NEAR = INK_RGB.map((a, i) => INK_RGB.map((b, j) => {
  if (Math.abs(a[0] - b[0]) + Math.abs(a[1] - b[1]) + Math.abs(a[2] - b[2]) > SNAP_TOLERANCE) return false;
  // Distance alone is not enough, and the asymmetry matters. At the dark end
  // a dim grey and a dim brown are only 94 apart, so a distance-only rule let
  // the chain carry brown UP a concrete ceiling and painted the whole level
  // one colour. Substituting a tier is fine — it is a shade of the same
  // thing. Substituting a family is a claim about what the surface is, and
  // only one direction of that claim is safe: a faint tint may collapse into
  // the grey beside it, but grey may never pick a colour up.
  const keep = Math.floor(i / TIERS.length), want = Math.floor(j / TIERS.length);
  return keep === want || keep === GREY;
}));

function familyOf(r, g, b) {
  const mx = Math.max(r, g, b), mn = Math.min(r, g, b), d = mx - mn;
  if (d === 0 || mx === 0 || d / mx < SAT_GATE) return 0;
  let h;
  if (mx === r) h = ((g - b) / d + 6) % 6;
  else if (mx === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  const deg = h * 60;
  if (deg < 15 || deg >= 330) return 1;
  if (deg < 52) return 2;
  if (deg < 80) return 3;
  if (deg < 160) return 4;
  if (deg < 265) return 5;
  return 6;
}

// ─── Auto-exposure ────────────────────────────────────────────────────
// Fit the tone curve to this frame. A 64-bucket histogram over every fifth
// pixel of the picture area is enough to place two percentiles and costs a
// fraction of the downsample that follows it; sorting was never needed.
// Doom's status bar is excluded because it is bright and constant, and
// letting it into the histogram would stop a dark room from opening up.
const EXPOSE_ROWS = 168;
const EXPOSE_STRIDE = 5;
const EXPOSE_LO = 0.06, EXPOSE_HI = 0.96;
const EXPOSE_MIN_RANGE = 0.08;   // a nearly uniform frame must not amplify its own noise
const EXPOSE_BUCKETS = 64;
const EXPOSE_GAMMA = 0.85;
const exposeHisto = new Uint32Array(EXPOSE_BUCKETS);

function exposure(fb) {
  exposeHisto.fill(0);
  let n = 0;
  const end = DOOM_W * EXPOSE_ROWS;
  for (let i = 0; i < end; i += EXPOSE_STRIDE) {
    const j = i * 4;   // BGRA
    const lum = 0.2126 * fb[j + 2] + 0.7152 * fb[j + 1] + 0.0722 * fb[j];
    exposeHisto[Math.min(EXPOSE_BUCKETS - 1, (lum * EXPOSE_BUCKETS / 256) | 0)]++;
    n++;
  }
  const percentile = (q) => {
    let seen = 0;
    const target = q * n;
    for (let b = 0; b < EXPOSE_BUCKETS; b++) {
      seen += exposeHisto[b];
      if (seen >= target) return (b + 0.5) / EXPOSE_BUCKETS;
    }
    return 1;
  };
  const lo = percentile(EXPOSE_LO);
  return [lo, Math.max(lo + EXPOSE_MIN_RANGE, percentile(EXPOSE_HI))];
}

// Cell → source-block bounds, precomputed once. Unlike the bitmap painter,
// which point-samples so that Doom's own palette survives, this one averages
// each cell's top and bottom halves: the glyph is chosen from a tone ladder,
// so smooth input is what it wants, and averaging is also what lets a cell
// notice a horizontal edge inside itself.
const CELL_X0 = Array.from({ length: VIEW_W }, (_, x) => Math.floor(x * DOOM_W / VIEW_W));
const CELL_X1 = Array.from({ length: VIEW_W }, (_, x) => Math.max(
  Math.floor(x * DOOM_W / VIEW_W) + 1, Math.floor((x + 1) * DOOM_W / VIEW_W)));
const CELL_Y0 = Array.from({ length: FIELD_ROWS }, (_, y) => Math.floor(y * DOOM_H / FIELD_ROWS));
const CELL_Y1 = Array.from({ length: FIELD_ROWS }, (_, y) => Math.max(
  Math.floor(y * DOOM_H / FIELD_ROWS) + 1, Math.floor((y + 1) * DOOM_H / FIELD_ROWS)));

/** Paint one frame as ASCII: brightness in the glyph, identity in the ink.
 *
 *  Exported and pure for the same reason the bitmap painter is — the headless
 *  tests drive it with synthetic frames, and the canvas-font guard needs to
 *  see every glyph this can emit. */
const aboveInk = new Int16Array(VIEW_W);

export function paintFramebufferAscii(g, fb) {
  const [lo, hi] = exposure(fb);
  const range = hi - lo;
  const tone = (lum) => {
    const t = (lum / 255 - lo) / range;
    return t <= 0 ? 0 : t >= 1 ? 1 : t ** EXPOSE_GAMMA;
  };

  aboveInk.fill(-1);
  for (let row = 0; row < FIELD_ROWS; row++) {
    const gy = FIELD_TOP + row;
    const y0 = CELL_Y0[row], y1 = CELL_Y1[row];
    const mid = Math.min(y1 - 1, Math.max(y0 + 1, (y0 + y1) >> 1));
    // Runs are built per row, so the left-neighbour hysteresis resets here.
    let prevInk = -1;
    for (let x = 0; x < VIEW_W; x++) {
      const x0 = CELL_X0[x], x1 = CELL_X1[x];
      let tr = 0, tg = 0, tb = 0, tn = 0;
      let br = 0, bg = 0, bb = 0, bn = 0;
      for (let y = y0; y < y1; y++) {
        const rowBase = y * DOOM_W;
        const lower = y >= mid;
        for (let sx = x0; sx < x1; sx++) {
          // BGRA in memory: blue first.
          const i = (rowBase + sx) * 4;
          if (lower) { br += fb[i + 2]; bg += fb[i + 1]; bb += fb[i]; bn++; }
          else { tr += fb[i + 2]; tg += fb[i + 1]; tb += fb[i]; tn++; }
        }
      }
      if (!tn) { tr = br; tg = bg; tb = bb; tn = bn || 1; }
      if (!bn) { br = tr; bg = tg; bb = tb; bn = tn || 1; }
      tr /= tn; tg /= tn; tb /= tn;
      br /= bn; bg /= bn; bb /= bn;

      const lt = tone(0.2126 * tr + 0.7152 * tg + 0.0722 * tb);
      const lb = tone(0.2126 * br + 0.7152 * bg + 0.0722 * bb);
      const l = (lt + lb) / 2;

      const rung = LADDER[Math.min(LADDER.length - 1, Math.round(l * (LADDER.length - 1)))];
      // A strong top/bottom split is an edge: draw it with a half block, which
      // is sub-cell vertical detail at no run cost because the ink is one value.
      const ch = (l > EDGE_FLOOR && Math.abs(lt - lb) > EDGE_SPLIT)
        ? (lt > lb ? '▀' : '▄')
        : RAMP[rung.glyph];

      let ink = familyOf((tr + br) / 2, (tg + bg) / 2, (tb + bb) / 2) * TIERS.length + rung.tier;
      if (prevInk >= 0 && SNAP_NEAR[prevInk][ink]) ink = prevInk;
      else if (aboveInk[x] >= 0 && SNAP_NEAR[aboveInk[x]][ink]) ink = aboveInk[x];
      prevInk = ink;
      aboveInk[x] = ink;

      g.chars[gy][1 + x] = ch;
      g.colors[gy][1 + x] = INK[ink];
      g.bgs[gy][1 + x] = null;      // no shading: the glyph carries the brightness
    }
  }
}

// ─── The engine, loaded once per page ─────────────────────────────────
// doomgeneric talks to its host through bare global functions (its C calls
// them with EM_ASM, which resolves free names against globalThis), so there
// can only ever be one Doom per page. That is fine — the arcade shows one
// cartridge at a time — but it does mean the module is a page-level
// singleton rather than per-cartridge state.
let enginePromise = null;

// ─── Where the engine and the game data come from ────────────────────
// Neither is in this repository. Together they are 13 MB of binary that would
// sit in a docs directory forever, show up in every clone, and never diff
// usefully — so both are pinned on jsDelivr instead, which serves them with
// `access-control-allow-origin: *` and an immutable cache.
//
// Both pins are IMMUTABLE by construction: jsDelivr resolves `gh/…@<40-hex>`
// to that exact commit, so the bytes behind these URLs cannot change under us.
// The recorded SHA-256 of each is in vendor/NOTICE.md, along with how to
// re-derive it.
//
// The engine is upstream's own published build, straight from the doomgenericjs
// tree — vendoring it only ever meant copying it. The IWAD is Freedoom's
// release asset, which lives in a small sibling repository because GitHub
// serves release assets without CORS and a browser therefore cannot fetch one.
const DEFAULT_ENGINE =
  'https://cdn.jsdelivr.net/gh/grubbyplaya/doomgenericjs'
  + '@99d7a55651b5f774e9b8911ef96e91a3652ef85f/doomgeneric/doomgeneric_module.js';
const DEFAULT_WAD =
  'https://cdn.jsdelivr.net/gh/JSv4/freedoom-iwad@70ee6ec942d090b4dd7ba04927f09ac79c8dc085/freedoom1.wad.gz';

/** Refuse any URL that is neither one of our own pinned constants nor
 *  same-origin.
 *
 *  `import()` EXECUTES what it fetches, with this page's privileges and on
 *  this page's origin, so an engine URL a link could choose would be remote
 *  code execution rather than a convenience — which is why `?doomEngine=` does
 *  not exist and why the engine is only ever the hardcoded pin above. The IWAD
 *  is only data and is magic-checked before use, but it goes through the same
 *  gate: `?wad=` is for pointing the cartridge at an IWAD you host yourself,
 *  and nothing is lost by requiring you to host it. The allowlist is therefore
 *  exactly {our pin} ∪ {same-origin} — a pin is trusted because it is written
 *  here, not because of where it points.
 *
 *  In Node (the headless logic tests) there is no location and nothing is ever
 *  loaded, so the guard is inert there. */
function allowedUrl(url, pinned, what) {
  if (url === pinned) return url;
  const here = globalThis.location;
  if (!here) return url;
  const resolved = new URL(url, here.href);
  if (resolved.origin !== here.origin) {
    throw new Error(`${what} must be same-origin (refusing ${resolved.origin})`);
  }
  return resolved.href;
}

/** Fetch the gzipped IWAD, reporting progress, and inflate it in the browser.
 *
 *  The WAD is stored gzipped (28.8 MB → 10.3 MB): a CDN will not
 *  content-encode an `application/octet-stream`, so pre-compressing it is the
 *  only way the visitor's download is the small number. DecompressionStream
 *  does the inflate natively — no library. */
async function fetchWad(url, onProgress) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`IWAD fetch failed: HTTP ${res.status}`);
  const total = Number(res.headers.get('content-length')) || 0;
  let received = 0;
  const counted = new TransformStream({
    transform(chunk, controller) {
      received += chunk.byteLength;
      onProgress(total ? received / total : 0, received, total);
      controller.enqueue(chunk);
    },
  });
  const stream = res.body.pipeThrough(counted).pipeThrough(new DecompressionStream('gzip'));
  const bytes = new Uint8Array(await new Response(stream).arrayBuffer());
  if (String.fromCharCode(...bytes.subarray(0, 4)) !== 'IWAD') {
    throw new Error('that file is not an IWAD');
  }
  return bytes;
}

/** Boot doomgeneric: install the DGJS_* host callbacks, instantiate the
 *  module, drop the IWAD into its in-memory filesystem under a name Doom's
 *  own IWAD table recognises, and create the game. Resolves to a handle the
 *  cartridge ticks. */
function loadEngine({ engineUrl, wadUrl, sound, onProgress }) {
  if (enginePromise) return enginePromise;

  enginePromise = (async () => {
    engineUrl = allowedUrl(engineUrl, DEFAULT_ENGINE, 'the Doom engine');
    wadUrl = allowedUrl(wadUrl, DEFAULT_WAD, 'the IWAD');
    onProgress({ phase: 'wad', ratio: 0 });
    const wad = await fetchWad(wadUrl, (ratio, got, total) =>
      onProgress({ phase: 'wad', ratio, got, total }));

    onProgress({ phase: 'engine', ratio: 0 });

    const handle = {
      frames: 0, title: null, module: null,
      keys: [],                      // queued [doomKey, pressed] transitions
      framebuffer: new Uint8Array(DOOM_W * DOOM_H * 4),
      framePtr: 0,
    };
    const audio = sound ? createAudio() : null;
    handle.audio = audio;

    installHostCallbacks(handle, audio);

    const createDoom = (await import(/* @vite-ignore */ engineUrl)).default;
    onProgress({ phase: 'engine', ratio: 1 });
    const module = await createDoom({});
    handle.module = module;

    // Doom's IWAD table (d_iwad.c) knows "freedoom1.wad" as Freedoom: Phase 1
    // and treats it as a retail Doom, so mounting it under its own name is
    // all the configuration this needs — no argv, no -iwad.
    // From the PATH, so a query string or fragment cannot end up in the
    // filename Doom's IWAD table is matched against.
    const path = new URL(wadUrl, globalThis.location?.href ?? 'https://localhost/').pathname;
    const name = '/' + path.split('/').pop().replace(/\.gz$/, '').toLowerCase();
    const stream = module.FS.open(name, 'w+');
    module.FS.write(stream, wad, 0, wad.length, 0);
    module.FS.close(stream);

    module.ccall('doomgeneric_Create', 'void', ['number', 'number'], [0, 0]);
    onProgress({ phase: 'ready', ratio: 1 });
    return handle;
  })().catch((error) => {
    enginePromise = null; // a failed boot must not poison a retry
    throw error;
  });

  return enginePromise;
}

/** The functions doomgeneric's C calls out to. They have to be globals: the
 *  engine invokes them from EM_ASM, whose free identifiers resolve against
 *  globalThis, not this module's scope. */
function installHostCallbacks(handle, audio) {
  const g = globalThis;
  const mem = () => handle.module.HEAPU8; // re-read: growth swaps the view

  g.DGJS_DrawFrame = (ptr, w, h) => {
    handle.framePtr = ptr;
    handle.frames++;
    handle.framebuffer.set(mem().subarray(ptr, ptr + w * h * 4));
  };
  g.DGJS_SetTitle = (ptr, len) => {
    handle.title = String.fromCharCode(...mem().subarray(ptr, ptr + len));
  };
  g.DGJS_GetKey = () => handle.keys.shift() ?? [0, 0];

  // Music would mean shipping a MIDI synthesizer and a soundfont next to a
  // demo that is already asking for 10 MB. Declining the driver is a
  // supported answer — i_jsmusic.c propagates the false and Doom runs mute.
  g.DGJS_MusicType = true;
  g.DGJS_InitMusic = () => false;
  g.DGJS_RegisterSong = () => 0;
  g.DGJS_UnRegisterSong = () => {};
  g.DGJS_PlaySong = () => {};
  g.DGJS_StopSong = () => {};
  g.DGJS_PauseSong = () => {};
  g.DGJS_ResumeSong = () => {};
  g.DGJS_SetMusicVolume = () => {};
  g.DGJS_PollMusic = () => {};

  if (!audio) {
    g.DGJS_InitSound = () => false;
    g.DGJS_ShutdownSound = () => {};
    g.DGJS_UpdateSound = () => {};
    g.DGJS_UpdateSoundParams = () => {};
    g.DGJS_StartSound = () => -1;
    g.DGJS_StopSound = () => {};
    g.DGJS_SoundIsPlaying = () => false;
    g.DGJS_CacheSFX_PCM = () => {};
    g.DGJS_CacheSFX_Buzzer = () => {};
    return;
  }
  audio.install(g, mem);
}

/** Doom's sound effects on WebAudio. Sixteen channels, each a gain + panner,
 *  fed from the DMX PCM lumps the engine hands over at cache time.
 *
 *  Every entry point is wrapped: a browser that refuses an AudioContext, or
 *  suspends one, must cost the player sound and nothing else — never the
 *  game. */
function createAudio() {
  let ctx = null;
  const buffers = [];
  const channels = new Array(16).fill(null).map(() => ({ src: null, gain: null, pan: null, playing: false }));

  const context = () => {
    if (ctx) return ctx;
    const Ctor = globalThis.AudioContext ?? globalThis.webkitAudioContext;
    if (!Ctor) return null;
    ctx = new Ctor();
    return ctx;
  };
  const guard = (fn, fallback) => (...args) => {
    try { return fn(...args); } catch { return fallback; }
  };

  return {
    /** The arcade calls this from the resume path, which is inside a user
     *  gesture — the only moment a browser will let an AudioContext start. */
    resume: () => { try { ctx?.resume?.(); } catch { /* not fatal */ } },
    install(g, mem) {
      g.DGJS_InitSound = guard(() => context() !== null, false);
      g.DGJS_ShutdownSound = guard(() => { ctx?.close?.(); ctx = null; });
      g.DGJS_UpdateSound = () => {};
      g.DGJS_UpdateSoundParams = guard((ch, volume, pan) => {
        const c = channels[ch];
        if (!c?.gain) return;
        c.gain.value = volume / 127;
        c.pan.value = (pan - 128) / 127;
      });
      g.DGJS_StartSound = guard((id, ch, volume, pan) => {
        const audioCtx = context();
        const buffer = buffers[id];
        if (!audioCtx || !buffer) return -1;
        const src = audioCtx.createBufferSource();
        const gain = audioCtx.createGain();
        const panner = audioCtx.createStereoPanner();
        src.buffer = buffer;
        gain.gain.value = volume / 127;
        panner.pan.value = (pan - 128) / 127;
        src.connect(panner); panner.connect(gain); gain.connect(audioCtx.destination);
        const c = channels[ch];
        c.src = src; c.gain = gain.gain; c.pan = panner.pan; c.playing = true;
        src.onended = () => { c.playing = false; };
        src.start();
        return ch;
      }, -1);
      g.DGJS_StopSound = guard((ch) => { channels[ch]?.src?.stop(); });
      g.DGJS_SoundIsPlaying = guard((ch) => channels[ch]?.playing === true, false);
      g.DGJS_CacheSFX_PCM = guard((ptr, len, id) => {
        const audioCtx = context();
        if (!audioCtx) return;
        const raw = mem().subarray(ptr, ptr + len);
        // DMX sound lump: 8-byte header (format, sample rate, sample count),
        // then unsigned 8-bit samples.
        const rate = (raw[3] << 8) | raw[2];
        const samples = raw.length - 8;
        if (samples <= 0 || !rate) return;
        const buffer = audioCtx.createBuffer(1, samples, rate);
        const out = buffer.getChannelData(0);
        for (let i = 0; i < samples; i++) out[i] = (raw[i + 8] - 0x80) / 128;
        buffers[id] = buffer;
      });
      g.DGJS_CacheSFX_Buzzer = () => {};
    },
  };
}

// ═══════════════════════════════════════════════════════════════════════
// The cartridge
// ═══════════════════════════════════════════════════════════════════════

const CONTROLS = [
  ['W / S', 'forward, back'],
  ['A / D', 'strafe'],
  ['← / →', 'turn'],
  ['Space', 'fire'],
  ['E', 'open, use'],
  ['Shift', 'run'],
  ['1…7', 'weapon'],
  ['M', 'automap'],
  ['Enter', 'confirm'],
  ['Q', "Doom's menu"],
  ['P', 'projection'],
  ['Esc', 'pause & edit'],
];

/** The two ways this cartridge can put Doom on a Word paragraph.
 *
 *  `bitmap` is the faithful reading of the framebuffer: two pixels per cell,
 *  ~550 runs a frame and ~1.8 through the converter. `ascii` spends the free
 *  axis (the glyph) instead of the expensive one (the ink) for ~470 runs and
 *  ~2.7. They were 3.7× apart before the bitmap painter learned to merge
 *  neighbouring cells; now they are close enough that the choice is about how
 *  each one FAILS, not how fast it is — the bitmap smears horizontally, the
 *  ASCII flattens to a 21-ink palette. Default is `bitmap`, because it is
 *  both the picture worth showing first and, now, a playable one. */
export const PROJECTIONS = ['bitmap', 'ascii'];

export function doomCart(options = {}) {
  const engineUrl = options.engineUrl ?? DEFAULT_ENGINE;
  const wadUrl = options.wadUrl ?? DEFAULT_WAD;
  const sound = options.sound !== false;
  let projection = PROJECTIONS.includes(options.projection) ? options.projection : 'bitmap';

  let handle = null;
  let status = 'idle';           // idle → loading → playing | error
  let progress = { phase: 'wad', ratio: 0 };
  let error = null;
  let paintedFrames = 0;
  let spinner = 0;
  let edited = false;

  function begin() {
    if (status !== 'idle') return;
    status = 'loading';
    loadEngine({
      engineUrl, wadUrl, sound,
      onProgress: (p) => { progress = p; },
    }).then((h) => {
      handle = h;
      status = 'playing';
    }).catch((e) => {
      error = e?.message ?? String(e);
      status = 'error';
    });
  }

  function tick(dt, input) {
    spinner += dt;
    if (status === 'idle') begin();
    if (status !== 'playing' || !handle) return;

    // Drain the arcade's key-transition log into Doom's queue. Doom reads one
    // transition per DG_GetKey call and expects both edges, which is why the
    // cartridge wants the log rather than the held/pressed sets the other two
    // cartridges use.
    for (const { code, down } of input.drain?.() ?? []) {
      if (code === PROJECTION_KEY) {
        // Local to the cartridge: Doom never sees it, and only the press edge
        // toggles, so holding the key does not flicker between projections.
        if (down) projection = projection === 'bitmap' ? 'ascii' : 'bitmap';
        continue;
      }
      const key = KEY_MAP[code];
      if (key !== undefined) handle.keys.push([key, down ? 1 : 0]);
    }

    // One Tick is one turn of Doom's own loop. It paces itself off wall-clock
    // time (DG_GetTicksMs), so calling it once per arcade frame is right at
    // any pace: fast, it re-renders without advancing a tic; slow, it runs
    // the tics it owes.
    handle.module._doomgeneric_Tick();
  }

  function drawLoading(g) {
    const bar = Math.round(Math.max(0, Math.min(1, progress.ratio ?? 0)) * 40);
    const dots = '.'.repeat(1 + (Math.floor(spinner * 3) % 3));
    const cy = Math.floor(ROWS / 2) - 2;
    const cx = Math.floor((COLS - 40) / 2);
    const phase = status === 'error' ? 'could not start'
      : progress.phase === 'engine' ? `starting the Doom engine${dots}`
        : progress.phase === 'ready' ? 'entering the level' + dots
          : `downloading the IWAD${dots}`;
    write(g, cy, cx, 'D O O M   I N   A   W O R D   D O C U M E N T'.slice(0, 40), PANEL_HI);
    write(g, cy + 2, cx, phase, HUD_INK);
    write(g, cy + 3, cx, '[' + '█'.repeat(bar) + '░'.repeat(40 - bar) + ']', PANEL_HI);
    if (progress.total) {
      const mb = (n) => (n / 1048576).toFixed(1);
      write(g, cy + 4, cx, `${mb(progress.got ?? 0)} / ${mb(progress.total)} MB`, DEAD_INK);
    }
    if (status === 'error') {
      for (const [i, line] of wrap(error ?? 'unknown', 60).slice(0, 3).entries()) {
        write(g, cy + 5 + i, cx, line, 'FF6B6B');
      }
    }
  }

  function drawPanel(g) {
    let y = FIELD_TOP;
    const x = PANEL_X + 1;
    write(g, y, x, 'DOOM', PANEL_HI);
    write(g, y + 1, x, handle?.title ? handle.title.trim().slice(0, PANEL_W - 2) : '', DEAD_INK);
    y += 3;
    for (const [key, what] of CONTROLS) {
      if (y >= ROWS - 4) break;
      write(g, y, x, key.padEnd(7), PANEL_HI);
      write(g, y, x + 7, what.slice(0, PANEL_W - 9), PANEL_INK);
      y++;
    }
    y = ROWS - 5;
    write(g, y, x, projection === 'ascii' ? 'ASCII · fast' : 'BITMAP · faithful', PANEL_HI);
    y += 1;
    write(g, y, x, `frame ${handle?.frames ?? 0}`, DEAD_INK);
    write(g, y + 1, x, 'id Software engine,', DEAD_INK);
    write(g, y + 2, x, 'GPL-2.0 · Freedoom data', DEAD_INK);
  }

  function render() {
    const g = makeGrid();
    if (status === 'playing' && handle) {
      if (projection === 'ascii') paintFramebufferAscii(g, handle.framebuffer);
      else paintFramebuffer(g, handle.framebuffer);
      paintedFrames++;
      drawDivider(g);
      drawPanel(g);
      drawChrome(g, projection === 'ascii'
        ? 'DOOM — id Software’s own engine, projected to ASCII in one Word paragraph'
        : 'DOOM — id Software’s own engine, drawing into one Word paragraph');
    } else {
      drawLoading(g);
      drawChrome(g, 'DOOM — the real engine, compiled to JavaScript');
    }
    return { grid: g, bg: BG };
  }

  return {
    name: 'doom',
    label: '☩ DOOM',
    caption:
      'The **actual** game: id Software’s Doom engine — GPL-2.0, compiled to JavaScript by ' +
      '[doomgeneric](https://github.com/grubbyplaya/doomgenericjs) — running on Freedoom’s ' +
      'BSD-licensed game data. Its 320×200 framebuffer is downsampled every frame into this Word ' +
      'paragraph as half-block characters: the ink of each `▀` is the top pixel and its `w:shd` ' +
      'run shading is the bottom one. **P** switches projection — that faithful bitmap costs ' +
      '~550 colored runs a frame, because neighbouring cells that are within a luma-weighted ' +
      'tolerance of each other are given the *same* color and merge into one run — without that ' +
      'a frame costs ~1,370 and repaints about once a second. The ASCII projection instead splits ' +
      'the picture between a cell’s two channels by what each one costs: brightness goes in the ' +
      'GLYPH, which is free because a run only breaks when the ink changes, and the ink is left ' +
      'to say what a surface *is* — brick, concrete, sky, blood, a key. Move **W/S** · strafe ' +
      '**A/D** · turn **←/→** · **Space** ' +
      'fires · **E** opens · **Q** is Doom’s own menu. **Esc** pauses — and then it is only a ' +
      'document again: put your caret in the frame, Undo rewinds it, Save downloads it as .docx.',
    hint: '<b>WASD</b> move · <b>←/→</b> turn · <b>Space</b> fire · <b>E</b> open · <b>Q</b> Doom’s menu · <b>P</b> switches projection — bitmap is the faithful one and smears when pushed, ASCII flattens to a fixed palette instead.',
    reset() {
      // Doom's own state lives inside the WebAssembly heap and the engine is
      // a page singleton, so a cartridge reset cannot restart the game. Q
      // opens Doom's menu, which is where a Doom player restarts a Doom.
      paintedFrames = 0;
      if (status === 'idle') begin();
    },
    tick,
    render,
    /** The other two cartridges re-read their world from the paragraph on
     *  every resume — type a wall into the map, resume, walk into it. Real
     *  Doom cannot: the level is BSP geometry in the WebAssembly heap, not
     *  text, and nothing typed into a downsampled framebuffer means anything
     *  to it. So the round-trip here is honest about what it is — the frame
     *  stays editable, undoable and saveable like any paragraph, and the next
     *  frame paints over whatever was typed. */
    syncFromRows() { edited = true; },
    /** Switch projection from outside the keyboard (the specs use this). */
    setProjection(next) {
      if (PROJECTIONS.includes(next)) projection = next;
      return projection;
    },
    state: () => ({
      status,
      error,
      progress,
      edited,
      projection,
      title: handle?.title ?? null,
      doomFrames: handle?.frames ?? 0,
      paintedFrames,
      /** A cheap digest of the live framebuffer, so a spec can prove the
       *  picture actually changes when keys are sent — the difference between
       *  "Doom booted" and "Doom is playing". */
      frameHash: () => {
        const fb = handle?.framebuffer;
        if (!fb) return 0;
        let h = 2166136261;
        for (let i = 0; i < fb.length; i += 997) h = Math.imul(h ^ fb[i], 16777619);
        return h >>> 0;
      },
      pixel: (x, y) => {
        const fb = handle?.framebuffer;
        if (!fb) return null;
        const i = (y * DOOM_W + x) * 4;
        return [fb[i + 2], fb[i + 1], fb[i]];
      },
      resumeAudio: () => handle?.audio?.resume?.(),
    }),
  };
}

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
// The screen is a character grid — the same kind the rest of the arcade draws
// on (see ascii-scenes.js), so it inherits the pinned canvas font and the
// markdown-safe bezel unchanged. A character cell is far too coarse to hold a
// Doom frame at one color per cell — but a cell is not one pixel. It has an
// ink, a `w:shd` shading, and a glyph, and only the two colors cost anything:
// a run breaks when the ink or the shading changes and NEVER when the glyph
// does, because every character in a span shares one `w:t`.
//
// Both projections are answers to that. The faithful `bitmap` one spends the
// expensive channel on every cell — `▀` with the top pixel as ink and the
// bottom as shading, one pixel pair per cell. The default `8bit` one spends
// the FREE channel instead: the sixteen quadrant block characters can draw
// every 2×2 arrangement of two colors, so one cell carries four sub-pixels and
// the picture is sampled at twice the cell grid in both axes. That extra
// resolution is genuinely free — it adds no runs and no colors — which is why
// the colors can then be rationed hard (see BUDGET) without the picture
// falling apart.
//
// THIS CARTRIDGE DRAWS ON ITS OWN GRID
// ------------------------------------
// The rest of the arcade shares one 92 × 26 grid of 8pt cells. Doom does not:
// it draws on 140 × 37 cells of 5.5pt text on a 7.05pt line, which occupies
// the very same 462pt × 261pt of page (the shared grid is 441pt × 260pt) and
// therefore keeps the pinned canvas font, the markdown-safe bezel and the
// saved .docx exactly as legible as before — every cell is still a real
// character with a real color in a real paragraph.
//
// The point of the denser grid is that CELLS are what carry resolution while
// RUNS are what cost. A run cannot cross a line break, so the run floor is the
// row count; everything above that floor is color the merge is free to spend.
// Going from 23 picture rows to 34 raises the floor by eleven runs and buys
// 2.2× the sub-pixels: 194 × 68 samples where there were 128 × 46.
//
// The viewport is 97 columns wide because that is 4:3 at this cell's aspect:
// a cell advances 3.3pt and is 7.05pt tall, so the 97 × 34 cells are
// 320pt × 240pt — Doom's own picture shape, not a stretched one. The
// sub-pixels inside them are 1.65pt × 3.53pt, finer across than down, which
// suits a game whose structure is mostly vertical edges.
// ═══════════════════════════════════════════════════════════════════════

// The cartridge's own grid, and the cell metrics that make it fit the page.
// `METRICS` travels with the frame (see the object this module returns) so the
// canvas paragraph is emitted at this cartridge's font size and line height
// rather than the arcade's shared ones.
const COLS = 140, ROWS = 37;
export const METRICS = { sz: 11, lineTwips: 141 };  // 5.5pt text, 7.05pt line

// ─── Doom's framebuffer, and where it lands on the grid ───────────────
const DOOM_W = 320, DOOM_H = 200;

const FIELD_TOP = 2;                      // 0 = bezel, 1 = HUD
const FIELD_ROWS = ROWS - FIELD_TOP - 1;  // 34 rows of picture
const VIEW_W = 97;                        // 4:3 at this cell aspect
const DIV_X = 1 + VIEW_W;                 // the │ divider column
const PANEL_X = DIV_X + 1;                // first column of the side panel
const PANEL_W = COLS - 1 - PANEL_X;       // 40 columns

/** Two half-pixels per cell row. */
const PIX_H = FIELD_ROWS * 2;             // 68

// Chrome colours. There is exactly one, and that is a frame-budget decision
// rather than a taste one: a run breaks when the ink changes, a null ink
// inherits the previous cell's, and every row of this paragraph is redrawn
// every frame. The bezel, the divider and the side panel were five colours,
// which cost five runs on every picture row — more than half the frame, for a
// column of static text. Sharing one ink makes the divider, the panel and the
// right bezel a single run per row.
//
// Rows are independent (each is its own sequence of runs, joined by w:br), so
// a row MAY use a second colour when it earns one — see PANEL_HI, spent on
// the two lines that change.
const CHROME_INK = '8FA3B8';
const BEZEL_INK = CHROME_INK;
const HUD_INK = CHROME_INK;
const PANEL_INK = CHROME_INK;
const PANEL_HI = CHROME_INK;
const DEAD_INK = CHROME_INK;
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

// ─── Framebuffer → grid, the 8-bit way ────────────────────────────────
// The bitmap painter above treats the paragraph as a screen. This one treats
// it as a console with a fixed palette and a run budget, and the difference
// between them is where each spends a cell's two channels.
//
// A cell has an INK and a w:shd SHADING, and a GLYPH. A run breaks when the
// ink or the shading changes; it never breaks on the glyph, because every
// character in a span shares one w:t. So colour is scarce and glyphs are free,
// and the measured frame cost is
//
//     frame_ms  ~=  63 + 0.67 * runs
//
// which makes five frames a second a budget of roughly 200 runs for the whole
// paragraph. Two ideas get a Doom frame inside it.
//
// ENDPOINTS, NOT PIXELS. Both earlier projections used the ink and shading as
// two independent samples — a top pixel and a bottom pixel. That wastes them.
// A glyph with fill fraction f renders as f*ink + (1-f)*bg, so those two
// colours are really the ENDPOINTS of a small ramp the whole run shares, and
// each cell picks its own weight along it for nothing. It is the same trade
// block texture compression makes, for the same reason: two endpoints plus
// cheap per-pixel weights beat two exact colours. Doom's shading is mostly
// light falling off along a surface, which is exactly what a ramp represents.
//
// A BUDGET, NOT A TOLERANCE. A tolerance spends whatever the scene happens to
// cost, so a busy frame blows the budget and a plain one wastes it. This
// allocates instead: quantise, then repeatedly merge the two adjacent runs
// whose merge adds the least error, until the frame is down to its allowance.
// Every run that survives is one the picture could least afford to lose, the
// allowance floats between rows — a blank ceiling keeps one run, the status
// bar keeps twenty — and the frame cost becomes a constant you choose rather
// than a property of the view. Measured: flat on every frame, whatever the view.
// Measured: with the screen fenced from the caption (see seedDocument in
// ascii-arcade.js) this lands the whole paragraph near 113 spans and the frame
// near 10.4 repaints a second.
const BUDGET = 64;

/** A console palette: few entries, well separated, deliberately more saturated
 *  than Doom's own. Every run's two endpoints are drawn from here, which is
 *  what keeps this reading as 8-bit art rather than as a blurred photograph. */
const PAL = [
  [0x00, 0x00, 0x00], [0x24, 0x1c, 0x18], [0x48, 0x41, 0x3c],   // shadow
  [0x7d, 0x7a, 0x80], [0xb9, 0xb6, 0xbd], [0xf2, 0xf0, 0xf5],   // concrete → highlight
  [0x3a, 0x24, 0x16], [0x6b, 0x44, 0x23], [0xa8, 0x6a, 0x30],   // brick and wood:
  [0xd9, 0xa0, 0x5b], [0xf0, 0xcb, 0x8a],                       //   most of E1M1
  [0x8c, 0x14, 0x14], [0xe0, 0x28, 0x28],                       // blood, the numerals
  [0x2e, 0x7a, 0x2e], [0x58, 0xd8, 0x58],                       // armour, slime
  [0x2a, 0x4a, 0x9c], [0x4a, 0x9a, 0xe0], [0xe8, 0xc8, 0x28],   // sky, keys, lights
];
const PAL_HEX = PAL.map((c) => HEX[c[0]] + HEX[c[1]] + HEX[c[2]]);

/** Glyphs, as which of a cell's four QUADRANTS the ink covers: top-left,
 *  top-right, bottom-left, bottom-right.
 *
 *  This is where the resolution comes from, and it is free. A run breaks on an
 *  ink or shading change and never on a glyph, so the number of picture
 *  samples is not what the frame costs — the run budget is, and the merge caps
 *  that regardless. A cell drawn with `▀` carries two sub-pixels; drawn with
 *  the quadrant blocks it carries FOUR, because all sixteen 2x2 patterns have
 *  a character. So the 97 x 34 cells sample the framebuffer at 194 x 68
 *  instead of 97 x 68, for no extra runs and no extra colours.
 *
 *  Still solid only — every quadrant is fully one endpoint or the other. The
 *  shade characters (░▒▓) are deliberately absent: they approximate TONE
 *  rather than carrying detail, and at the size this ships (a 3.3 x 7.05pt
 *  cell) they read as dots, not as a blend. `▚` and `▞` appear here for the
 *  opposite reason to why they were dropped before — as two of the sixteen
 *  real 2x2 patterns, not as a 50% grey. */
const GLYPHS = [
  [' ', 0, 0, 0, 0], ['▘', 1, 0, 0, 0], ['▝', 0, 1, 0, 0], ['▀', 1, 1, 0, 0],
  ['▖', 0, 0, 1, 0], ['▌', 1, 0, 1, 0], ['▞', 0, 1, 1, 0], ['▛', 1, 1, 1, 0],
  ['▗', 0, 0, 0, 1], ['▚', 1, 0, 0, 1], ['▐', 0, 1, 0, 1], ['▜', 1, 1, 0, 1],
  ['▄', 0, 0, 1, 1], ['▙', 1, 0, 1, 1], ['▟', 0, 1, 1, 1], ['█', 1, 1, 1, 1],
];
/** Fill fractions one HALF of a cell can render, which is what the endpoint
 *  fit below scores against: both quadrants background, one each, or both ink. */
const FILLS = [0, 0.5, 1];

// Luma-weighted distance: the eye resolves green far better than blue, so an
// even metric spends its budget where it will not be seen.
const LR = 0.30, LG = 0.59, LB = 0.11;
const dist3 = (ar, ag, ab, br, bg, bb) =>
  LR * Math.abs(ar - br) + LG * Math.abs(ag - bg) + LB * Math.abs(ab - bb);

/** Nearest palette entry, as a 16³ lookup built once — this is called for
 *  every endpoint of every trial merge, and a linear scan showed up. */
const PAL_LUT = new Uint8Array(4096);
for (let r = 0; r < 16; r++) {
  for (let g = 0; g < 16; g++) {
    for (let b = 0; b < 16; b++) {
      let best = 0, bd = Infinity;
      for (let i = 0; i < PAL.length; i++) {
        const d = dist3(PAL[i][0], PAL[i][1], PAL[i][2], r * 17, g * 17, b * 17);
        if (d < bd) { bd = d; best = i; }
      }
      PAL_LUT[(r << 8) | (g << 4) | b] = best;
    }
  }
}
const palOf = (r, g, b) => PAL_LUT[((r & 255) >> 4 << 8) | ((g & 255) >> 4 << 4) | ((b & 255) >> 4)];

// ─── Per-frame auto-exposure ──────────────────────────────────────────
// Doom is dark and its exposure moves — median luminance is 0.110 in a
// corridor and 0.185 in an open lit area — so one fixed curve serves neither.
// This was rejected in an earlier projection because it amplifies dark-region
// noise into extra runs; under a hard run budget that objection disappears,
// because the merge brings the frame back to its allowance whatever it is fed.
// Contrast is now free, and it is the single biggest legibility win available.
const EXP_ROWS = 168;          // Doom's picture; its status bar is fixed and bright
const EXP_BUCKETS = 64;
const EXP_LO = 0.04, EXP_HI = 0.98, EXP_MIN_RANGE = 0.08;
const expHisto = new Uint32Array(EXP_BUCKETS);

function exposure(fb) {
  expHisto.fill(0);
  let n = 0;
  for (let i = 0, end = DOOM_W * EXP_ROWS; i < end; i += 5) {
    const j = i * 4;   // BGRA
    const l = 0.2126 * fb[j + 2] + 0.7152 * fb[j + 1] + 0.0722 * fb[j];
    expHisto[Math.min(EXP_BUCKETS - 1, (l * EXP_BUCKETS / 256) | 0)]++;
    n++;
  }
  const at = (q) => {
    let seen = 0; const target = q * n;
    for (let b = 0; b < EXP_BUCKETS; b++) { seen += expHisto[b]; if (seen >= target) return (b + 0.5) / EXP_BUCKETS; }
    return 1;
  };
  const lo = at(EXP_LO);
  return [lo, Math.max(lo + EXP_MIN_RANGE, at(EXP_HI))];
}

// ─── Cell geometry, and the buffers the painter reuses ────────────────
const CELL_X0 = Array.from({ length: VIEW_W }, (_, x) => Math.floor(x * DOOM_W / VIEW_W));
const CELL_X1 = Array.from({ length: VIEW_W }, (_, x) => Math.max(
  Math.floor(x * DOOM_W / VIEW_W) + 1, Math.floor((x + 1) * DOOM_W / VIEW_W)));
const HALF_H = FIELD_ROWS * 2;
const CELL_Y0 = Array.from({ length: HALF_H }, (_, y) => Math.floor(y * DOOM_H / HALF_H));
const CELL_Y1 = Array.from({ length: HALF_H }, (_, y) => Math.max(
  Math.floor(y * DOOM_H / HALF_H) + 1, Math.floor((y + 1) * DOOM_H / HALF_H)));

/** Horizontal sub-sampling: two quadrant columns per cell. */
const QUAD_W = VIEW_W * 2;
const QUAD_X0 = Array.from({ length: QUAD_W }, (_, x) => Math.floor(x * DOOM_W / QUAD_W));
const QUAD_X1 = Array.from({ length: QUAD_W }, (_, x) => Math.max(
  Math.floor(x * DOOM_W / QUAD_W) + 1, Math.floor((x + 1) * DOOM_W / QUAD_W)));

/** Exposed mean colour of every QUADRANT, as flat RGB triples indexed
 *  [(row * 2 + half) * QUAD_W + qx] * 3. This is the 194 x 68 picture. */
const quadRGB = new Float32Array(HALF_H * QUAD_W * 3);
/** The same thing averaged to cell halves, [(row * 2 + half) * VIEW_W + x] * 3.
 *  The endpoint merge runs on this rather than on the quadrants: it only needs
 *  to know the colour RANGE a span covers, which the halves capture just as
 *  well, and running it at half the samples keeps the merge — the expensive
 *  part — exactly as cheap as it was before the resolution went up. */
const halfRGB = new Float32Array(HALF_H * VIEW_W * 3);
/** Each half's luminance, indexed the same way. The endpoint fit reads it
 *  three times per candidate span and the merge evaluates thousands of
 *  candidate spans a frame, so it is computed once with the halves. */
const halfL = new Float32Array(HALF_H * VIEW_W);

/** A binary heap of candidate merges, cheapest first. */
class MergeHeap {
  constructor() { this.a = []; }
  push(x) {
    const a = this.a; a.push(x);
    let i = a.length - 1;
    while (i > 0) { const p = (i - 1) >> 1; if (a[p].cost <= a[i].cost) break; const t = a[p]; a[p] = a[i]; a[i] = t; i = p; }
  }
  pop() {
    const a = this.a, top = a[0], last = a.pop();
    if (a.length) {
      a[0] = last;
      for (let i = 0; ;) {
        const l = 2 * i + 1, r = l + 1; let m = i;
        if (l < a.length && a[l].cost < a[m].cost) m = l;
        if (r < a.length && a[r].cost < a[m].cost) m = r;
        if (m === i) break;
        const t = a[m]; a[m] = a[i]; a[i] = t; i = m;
      }
    }
    return top;
  }
}

/** Paint one frame as 8-bit: colour allocated by budget, detail by glyph.
 *
 *  Exported and pure for the same reason the bitmap painter is — the headless
 *  tests drive it with synthetic frames, and the canvas-font guard needs to
 *  see every glyph this can emit. */
export function paintFramebuffer8Bit(g, fb) {
  const [lo, hi] = exposure(fb);
  const range = hi - lo;

  // Average the raw block, then expose ONCE per quadrant. Exposing every
  // source pixel instead cost seven times as much for no visible difference.
  for (let h = 0; h < HALF_H; h++) {
    const y0 = CELL_Y0[h], y1 = CELL_Y1[h];
    for (let qx = 0; qx < QUAD_W; qx++) {
      const x0 = QUAD_X0[qx], x1 = QUAD_X1[qx];
      let r = 0, gg = 0, b = 0, n = 0;
      for (let y = y0; y < y1; y++) {
        const base = y * DOOM_W;
        for (let sx = x0; sx < x1; sx++) { const i = (base + sx) * 4; r += fb[i + 2]; gg += fb[i + 1]; b += fb[i]; n++; }
      }
      r /= n; gg /= n; b /= n;
      const l = (0.2126 * r + 0.7152 * gg + 0.0722 * b) / 255;
      let er = 0, eg = 0, eb = 0;
      if (l > 0) {
        let t = (l - lo) / range;
        t = t <= 0 ? 0 : t >= 1 ? 1 : t;
        const gain = t / l;
        er = r * gain; eg = gg * gain; eb = b * gain;
        // Clamping each channel on its own walks a lit brown wall toward
        // white, because the channels clip at different points. Scale the
        // whole triple so the hue survives.
        const m = Math.max(er, eg, eb);
        if (m > 255) { const k = 255 / m; er *= k; eg *= k; eb *= k; }
      }
      const o = (h * QUAD_W + qx) * 3;
      quadRGB[o] = er; quadRGB[o + 1] = eg; quadRGB[o + 2] = eb;
    }
    // Cell halves, for the merge.
    for (let x = 0; x < VIEW_W; x++) {
      const a = (h * QUAD_W + x * 2) * 3, b2 = a + 3, o = (h * VIEW_W + x) * 3;
      const r = (quadRGB[a] + quadRGB[b2]) / 2;
      const gg = (quadRGB[a + 1] + quadRGB[b2 + 1]) / 2;
      const b = (quadRGB[a + 2] + quadRGB[b2 + 2]) / 2;
      halfRGB[o] = r; halfRGB[o + 1] = gg; halfRGB[o + 2] = b;
      halfL[h * VIEW_W + x] = 0.2126 * r + 0.7152 * gg + 0.0722 * b;
    }
  }

  /** Endpoints for one row's [x0,x1) span, plus the error of fitting the
   *  span's cell halves to the ramp between them. The ramp is what the glyph
   *  interpolates across, so the endpoints decide how much of the picture
   *  survives the run.
   *
   *  They are the two GROUP MEANS either side of the span's mean luminance —
   *  block truncation coding's rule, the same one a GPU texture format uses,
   *  and the reason this projection reads as shaded rather than as noise. The
   *  obvious rule, the span's darkest and brightest halves, is what it
   *  replaced: those are outliers, one specular highlight and one shadow set
   *  a ramp nothing else in the span lies on, and every mid-tone then snaps to
   *  whichever end is nearer. That is exactly the failure that shows up as a
   *  black-and-white checkerboard where a wall should be, and it gets worse
   *  the more samples a run has to cover — which is to say, worse at every
   *  resolution increase. Group means cost the same and hold. */
  const fit = (row, x0, x1) => {
    let sum = 0, n = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) { sum += halfL[rowBase + x]; n++; }
    }
    const mean = sum / n;
    let loR = 0, loG = 0, loB = 0, loN = 0, hiR = 0, hiG = 0, hiB = 0, hiN = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) {
        const o = (rowBase + x) * 3;
        const r = halfRGB[o], gg = halfRGB[o + 1], b = halfRGB[o + 2];
        if (halfL[rowBase + x] <= mean) { loR += r; loG += gg; loB += b; loN++; }
        else { hiR += r; hiG += gg; hiB += b; hiN++; }
      }
    }
    // A flat span puts every sample on one side; then both ends are that
    // colour and the glyph search below settles on a solid block.
    if (loN === 0) { loR = hiR; loG = hiG; loB = hiB; loN = hiN; }
    if (hiN === 0) { hiR = loR; hiG = loG; hiB = loB; hiN = loN; }
    loR /= loN; loG /= loN; loB /= loN;
    hiR /= hiN; hiG /= hiN; hiB /= hiN;
    const bgIdx = palOf(loR, loG, loB), inkIdx = palOf(hiR, hiG, hiB);
    const ink = PAL[inkIdx], bgc = PAL[bgIdx];
    let err = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) {
        const o = (rowBase + x) * 3;
        const r = halfRGB[o], gg = halfRGB[o + 1], b = halfRGB[o + 2];
        let bd = Infinity;
        for (let k = 0; k < FILLS.length; k++) {
          const f = FILLS[k], inv = 1 - f;
          const d = dist3(ink[0] * f + bgc[0] * inv, ink[1] * f + bgc[1] * inv, ink[2] * f + bgc[2] * inv, r, gg, b);
          if (d < bd) bd = d;
        }
        err += bd;
      }
    }
    return { ink: inkIdx, bg: bgIdx, err };
  };

  // Start from one run per cell, then merge down to the allowance.
  const rows = [];
  let total = 0;
  for (let row = 0; row < FIELD_ROWS; row++) {
    const seg = [];
    for (let x = 0; x < VIEW_W; x++) {
      const s = { row, x0: x, x1: x + 1, alive: true, ver: 0, prev: null, next: null };
      s.fit = fit(row, x, x + 1);
      seg.push(s);
    }
    for (let i = 0; i < seg.length; i++) { seg[i].prev = seg[i - 1] ?? null; seg[i].next = seg[i + 1] ?? null; }
    rows.push(seg); total += seg.length;
  }

  const heap = new MergeHeap();
  const offer = (a) => {
    if (!a || !a.next) return;
    const f = fit(a.row, a.x0, a.next.x1);
    heap.push({ cost: f.err - a.fit.err - a.next.fit.err, a, b: a.next, va: a.ver, vb: a.next.ver, fit: f });
  };
  for (const seg of rows) for (const s of seg) offer(s);

  while (total > BUDGET && heap.a.length) {
    const e = heap.pop();
    // Lazy invalidation: an entry is stale if either side has been merged
    // away or has since changed, which is cheaper than deleting from a heap.
    if (!e.a.alive || !e.b.alive || e.a.ver !== e.va || e.b.ver !== e.vb || e.a.next !== e.b) continue;
    const merged = { row: e.a.row, x0: e.a.x0, x1: e.b.x1, fit: e.fit, alive: true, ver: 0, prev: e.a.prev, next: e.b.next };
    if (merged.prev) merged.prev.next = merged;
    if (merged.next) merged.next.prev = merged;
    e.a.alive = false; e.b.alive = false;
    const seg = rows[merged.row];
    seg.splice(seg.indexOf(e.a), 2, merged);
    total--;
    if (merged.prev) { merged.prev.ver++; offer(merged.prev); }
    offer(merged);
  }

  // The free channel: inside its run, every cell picks its own weight.
  for (let row = 0; row < FIELD_ROWS; row++) {
    const gy = FIELD_TOP + row;
    for (const s of rows[row]) {
      const ink = PAL[s.fit.ink], bgc = PAL[s.fit.bg];
      const inkHex = PAL_HEX[s.fit.ink], bgHex = PAL_HEX[s.fit.bg];
      for (let x = s.x0; x < s.x1; x++) {
        // The four quadrant samples this cell has to represent, using only
        // its run's two colours. Sixteen patterns, so this is an exhaustive
        // search over every 2x2 arrangement the character set can draw.
        const q = [((row * 2) * QUAD_W + x * 2) * 3, ((row * 2) * QUAD_W + x * 2 + 1) * 3,
                   ((row * 2 + 1) * QUAD_W + x * 2) * 3, ((row * 2 + 1) * QUAD_W + x * 2 + 1) * 3];
        let best = 15, bd = Infinity;
        for (let k = 0; k < GLYPHS.length; k++) {
          const gl = GLYPHS[k];
          let d = 0;
          for (let i = 0; i < 4; i++) {
            const c = gl[i + 1] ? ink : bgc, o = q[i];
            d += dist3(c[0], c[1], c[2], quadRGB[o], quadRGB[o + 1], quadRGB[o + 2]);
          }
          if (d < bd) { bd = d; best = k; }
        }
        g.chars[gy][1 + x] = GLYPHS[best][0];
        g.colors[gy][1 + x] = inkHex;
        g.bgs[gy][1 + x] = bgHex;
      }
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
 *  `8bit` is the default because it is the one you can play: a fixed console
 *  palette, a hard run budget, and per-cell detail carried by the glyph, which
 *  is free. `bitmap` is the faithful reading — two pixels per cell, every
 *  colour the framebuffer had — and is three times the cost for it. The
 *  honest way to describe the pair is that one is the photograph and the other
 *  is the game. */
export const PROJECTIONS = ['8bit', 'bitmap'];

export function doomCart(options = {}) {
  const engineUrl = options.engineUrl ?? DEFAULT_ENGINE;
  const wadUrl = options.wadUrl ?? DEFAULT_WAD;
  const sound = options.sound !== false;
  let projection = PROJECTIONS.includes(options.projection) ? options.projection : '8bit';

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
        if (down) projection = projection === '8bit' ? 'bitmap' : '8bit';
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
    const title = 'D O O M   I N   A   W O R D   D O C U M E N T';
    const cx = Math.floor((COLS - title.length) / 2);
    const phase = status === 'error' ? 'could not start'
      : progress.phase === 'engine' ? `starting the Doom engine${dots}`
        : progress.phase === 'ready' ? 'entering the level' + dots
          : `downloading the IWAD${dots}`;
    write(g, cy, cx, title, PANEL_HI);
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
    write(g, y, x, projection === '8bit' ? '8-BIT · playable' : 'BITMAP · faithful', PANEL_HI);
    y += 1;
    write(g, y, x, `frame ${handle?.frames ?? 0}`, DEAD_INK);
    write(g, y + 1, x, 'id Software engine,', DEAD_INK);
    write(g, y + 2, x, 'GPL-2.0 · Freedoom data', DEAD_INK);
  }

/** Let the left bezel merge into the picture.
 *
 *  Every row of this paragraph starts with a box-drawing character so the
 *  editor's markdown blur-commit can never read a row as a heading or a
 *  bullet. That safety is a property of the CHARACTER, not of its colour — but
 *  giving the column its own ink cost it its own run on every picture row —
 *  one run per row of the frame, for a one-cell grey line. Painting the bezel in its
 *  neighbour's colours keeps the character exactly where it was and merges the
 *  run away. */
function mergeBezelIntoPicture(g) {
  for (let row = 0; row < FIELD_ROWS; row++) {
    const y = FIELD_TOP + row;
    g.colors[y][0] = g.colors[y][1];
    g.bgs[y][0] = g.bgs[y][1];
  }
}

  function render() {
    const g = makeGrid();
    if (status === 'playing' && handle) {
      if (projection === '8bit') paintFramebuffer8Bit(g, handle.framebuffer);
      else paintFramebuffer(g, handle.framebuffer);
      paintedFrames++;
      drawDivider(g);
      drawPanel(g);
      drawChrome(g, projection === '8bit'
        ? 'DOOM — id Software’s own engine, projected to 8-bit in one Word paragraph'
        : 'DOOM — id Software’s own engine, drawing into one Word paragraph');
      mergeBezelIntoPicture(g);
    } else {
      drawLoading(g);
      drawChrome(g, 'DOOM — the real engine, compiled to JavaScript');
    }
    return { grid: g, bg: BG, metrics: METRICS };
  }

  return {
    name: 'doom',
    label: '☩ DOOM',
    caption:
      'The **actual** game: id Software’s Doom engine — GPL-2.0, compiled to JavaScript by ' +
      '[doomgeneric](https://github.com/grubbyplaya/doomgenericjs) — running on Freedoom’s ' +
      'BSD-licensed game data. Its 320×200 framebuffer is redrawn into this Word paragraph every ' +
      'frame, and what it costs is *colored runs*: a run breaks when the ink or the `w:shd` ' +
      'shading changes, never when the glyph does, and the measured frame is `31 ms + 0.61 ms × ' +
      'runs`. So the default **8-bit** projection treats that as a budget and spends it — it ' +
      'squeezes the frame to a fixed ~64 runs by repeatedly merging whichever two neighbouring ' +
      'runs cost the least picture to lose, then gives each surviving run *two* colors to be the ' +
      'endpoints of a ramp, fitted the way a GPU texture format fits one. Every cell then picks ' +
      'its own arrangement of those two colors with a quadrant block — all sixteen 2×2 patterns ' +
      'have a character — so the picture is sampled at **194 × 68** even though only 97 × 34 ' +
      'cells are spent on it, and the extra resolution costs nothing at all: a run breaks on a ' +
      'color change, never on a glyph. Doom draws on its own denser grid for exactly that ' +
      'reason — 5.5pt cells in the same 462 × 261 points of page — because cells carry ' +
      'resolution while runs carry cost. Flat cost, whatever you are looking at, and about ten ' +
      'repaints a second. ' +
      '**P** switches to the faithful **bitmap**: every color the framebuffer had, two pixels per ' +
      'cell, six times the runs and a quarter of the rate. Move **W/S** · strafe ' +
      '**A/D** · turn **←/→** · **Space** ' +
      'fires · **E** opens · **Q** is Doom’s own menu. **Esc** pauses — and then it is only a ' +
      'document again: put your caret in the frame, Undo rewinds it, Save downloads it as .docx.',
    hint: '<b>WASD</b> move · <b>←/→</b> turn · <b>Space</b> fire · <b>E</b> open · <b>Q</b> Doom’s menu · <b>P</b> switches projection — 8-bit samples 194×68 on a fixed run budget and plays at ~10 fps, bitmap is faithful at ~2.4.',
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

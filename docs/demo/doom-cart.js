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
// Both projections are answers to that. The detailed `bitmap` one keeps an
// unrestricted framebuffer-derived palette, allocates 900 horizontal colour
// pairs where they preserve the most structure, and lets each cell's block
// glyph independently select those endpoints for its top and bottom samples.
// The default `8bit` one is the playable path: one high-contrast ink/shading
// pair across the picture and all sixteen quadrant block characters to draw
// every 2×2 arrangement of those endpoints. One cell carries four sub-pixels,
// and glyph changes cost no new Word runs.
//
// THIS CARTRIDGE DRAWS ON ITS OWN GRID
// ------------------------------------
// The rest of the arcade shares one 92 × 26 grid of 8pt cells. Doom does not:
// it draws on 96 × 32 cells of 8pt text on an 11.45pt line. The picture occupies
// 94 × 29 of them — almost the full 462pt document column, at Doom's 4:3
// display shape — while the controls live in a large 18pt document paragraph
// above it. That separation is load-bearing: text small enough to carry the
// framebuffer is not text a player can honestly read.
//
// The point of the denser grid is that CELLS carry resolution while rendered
// style boxes cost conversion time. One authored OOXML run can contain line
// breaks, but conversion still materializes one picture fragment per visual
// row. With quadrant blocks, those 94 × 29 cells carry 188 × 58 samples, enough
// to preserve room edges, sprites, the weapon and HUD while leaving the
// document conversion enough time for ten complete repaints a second.
//
// The viewport is 94 columns wide because an 8pt DejaVu Sans Mono cell advances
// about 4.8pt. On an 11.45pt line, 94 × 29 cells occupy about 451 × 332pt:
// Doom's own 4:3 display shape, not a stretched strip.
// ═══════════════════════════════════════════════════════════════════════

// The cartridge's own grid, and the cell metrics that make it fit the page.
// `METRICS` travels with the frame (see the object this module returns) so the
// canvas paragraph is emitted at this cartridge's font size and line height
// rather than the arcade's shared ones.
const COLS = 96, ROWS = 32;
export const METRICS = { sz: 16, lineTwips: 229 };  // 8pt text, 11.45pt line

// ─── Doom's framebuffer, and where it lands on the grid ───────────────
const DOOM_W = 320, DOOM_H = 200;

const FIELD_TOP = 2;                      // 0 = bezel, 1 = HUD
const FIELD_ROWS = ROWS - FIELD_TOP - 1;  // 29 rows of picture
const VIEW_W = COLS - 2;                  // full width between the two bezels

/** Two half-pixels per cell row. */
const PIX_H = FIELD_ROWS * 2;             // 58

// Chrome uses one colour. More importantly, the controls are not chrome at all
// any more: they are four static, large paragraphs outside this per-frame
// conversion. Both bezels borrow their adjacent picture cell's ink and shading
// below, so neither creates one extra run on every picture row.
const CHROME_INK = '8FA3B8';
const BEZEL_INK = CHROME_INK;
const HUD_INK = CHROME_INK;
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

/** Wrap loading errors without letting them widen the grid. */
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
// A run breaks when the ink or the shading changes, and an exact photographic
// downsample changes both on almost every cell: the 94 x 29 picture costs
// 6,000–7,000 runs in a normal E1M1 view. That is not remotely interactive —
// OOXML→HTML conversion is linear in runs — so this projection has always had
// to merge colour pairs.
//
// The old merge used a fixed tolerance: if a pixel pair was near the previous
// pair, both halves adopted the previous pair's colours. That met the existing
// 1,200-span guard, but it spent the loss in exactly the wrong place. Subtle
// wall and floor texture is made of near colours, so entire surfaces became
// horizontal bars; a real captured frame retained only 876 of 5,313 visible
// horizontal transitions. The result was cheaper but no longer legible.
//
// This painter uses the same two resources the high-contrast projection does:
//
//   * ALLOCATE expensive colour boundaries. Start with one exact segment per
//     cell, then merge the adjacent pair whose two-endpoint fit loses the least
//     picture until a hard budget is reached. Doors, monsters, HUD digits and
//     wall edges are expensive to merge, so the budget stays on them instead
//     of being consumed by harmless texture noise.
//   * SPEND the free glyph channel. A segment owns one open-palette ink/bg
//     pair, but each cell independently chooses ` `, `▀`, `▄` or `█`. Its top
//     and bottom samples therefore still select their nearer endpoint without
//     creating another run. On the frame that exposed the smear, this halves
//     perceptual error at the same run count and restores the room structure.
//
// 900 picture runs leaves room for the HUD/chrome inside the established
// 1,200 rendered-span ceiling. Unlike the high-contrast playable mode below,
// the endpoints remain means of Doom's own unrestricted framebuffer colours.
const BITMAP_BUDGET = 900;
const bitmapRGB = new Uint8Array(PIX_H * VIEW_W * 3);

/** Paint one 320×200 BGRA frame into the grid's viewport as half-block cells.
 *
 *  Exported (and pure) so the headless logic tests can drive it with a
 *  synthetic framebuffer: this is the one piece of the cartridge that has
 *  interesting behaviour and does not need a WebAssembly Doom to exercise. */
export function paintFramebuffer(g, fb) {
  // Sample once. `fit` evaluates thousands of candidate spans and must not
  // walk the BGRA framebuffer (or redo the channel swap) for each one.
  for (let h = 0; h < PIX_H; h++) {
    const base = SY[h] * DOOM_W;
    for (let x = 0; x < VIEW_W; x++) {
      const sx = SX[x];
      const from = (base + sx) * 4, to = (h * VIEW_W + x) * 3;
      bitmapRGB[to] = fb[from + 2];      // BGRA → RGB
      bitmapRGB[to + 1] = fb[from + 1];
      bitmapRGB[to + 2] = fb[from];
    }
  }

  /** Fit one unrestricted pair of colours to every half-cell in a horizontal
   *  span. Splitting at the span's mean luminance and taking the two group
   *  means is block truncation coding: robust against one bright outlier and
   *  cheap enough to evaluate for every candidate merge. */
  const fit = (row, x0, x1) => {
    let mean = 0, n = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) {
        const o = (rowBase + x) * 3;
        mean += 0.30 * bitmapRGB[o] + 0.59 * bitmapRGB[o + 1] + 0.11 * bitmapRGB[o + 2];
        n++;
      }
    }
    mean /= n;

    let lr = 0, lg = 0, lb = 0, ln = 0, hr = 0, hg = 0, hb = 0, hn = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) {
        const o = (rowBase + x) * 3;
        const r = bitmapRGB[o], gg = bitmapRGB[o + 1], b = bitmapRGB[o + 2];
        if (0.30 * r + 0.59 * gg + 0.11 * b <= mean) { lr += r; lg += gg; lb += b; ln++; }
        else { hr += r; hg += gg; hb += b; hn++; }
      }
    }
    if (!ln) { lr = hr; lg = hg; lb = hb; ln = hn; }
    if (!hn) { hr = lr; hg = lg; hb = lb; hn = ln; }
    lr = Math.round(lr / ln); lg = Math.round(lg / ln); lb = Math.round(lb / ln);
    hr = Math.round(hr / hn); hg = Math.round(hg / hn); hb = Math.round(hb / hn);

    let err = 0;
    for (let half = 0; half < 2; half++) {
      const rowBase = (row * 2 + half) * VIEW_W;
      for (let x = x0; x < x1; x++) {
        const o = (rowBase + x) * 3;
        const r = bitmapRGB[o], gg = bitmapRGB[o + 1], b = bitmapRGB[o + 2];
        err += Math.min(dist3(r, gg, b, lr, lg, lb), dist3(r, gg, b, hr, hg, hb));
      }
    }
    return { ink: [hr, hg, hb], bg: [lr, lg, lb], err };
  };

  // Begin exact — one segment per cell, whose two endpoints can reproduce
  // both halves — then make the globally least damaging adjacent merge.
  const rows = [];
  let total = 0;
  for (let row = 0; row < FIELD_ROWS; row++) {
    const seg = [];
    for (let x = 0; x < VIEW_W; x++) {
      const s = { row, x0: x, x1: x + 1, fit: fit(row, x, x + 1), alive: true,
        ver: 0, prev: null, next: null };
      seg.push(s);
    }
    for (let i = 0; i < seg.length; i++) { seg[i].prev = seg[i - 1] ?? null; seg[i].next = seg[i + 1] ?? null; }
    rows.push(seg); total += seg.length;
  }

  const heap = new MergeHeap();
  const offer = (a) => {
    if (!a || !a.next) return;
    const f = fit(a.row, a.x0, a.next.x1);
    heap.push({ cost: f.err - a.fit.err - a.next.fit.err, a, b: a.next,
      va: a.ver, vb: a.next.ver, fit: f });
  };
  for (const seg of rows) for (const s of seg) offer(s);

  while (total > BITMAP_BUDGET && heap.a.length) {
    const e = heap.pop();
    if (!e.a.alive || !e.b.alive || e.a.ver !== e.va || e.b.ver !== e.vb || e.a.next !== e.b) continue;
    const merged = { row: e.a.row, x0: e.a.x0, x1: e.b.x1, fit: e.fit,
      alive: true, ver: 0, prev: e.a.prev, next: e.b.next };
    if (merged.prev) merged.prev.next = merged;
    if (merged.next) merged.next.prev = merged;
    e.a.alive = false; e.b.alive = false;
    const seg = rows[merged.row];
    seg.splice(seg.indexOf(e.a), 2, merged);
    total--;
    if (merged.prev) { merged.prev.ver++; offer(merged.prev); }
    offer(merged);
  }

  for (let row = 0; row < FIELD_ROWS; row++) {
    const gy = FIELD_TOP + row;
    for (const s of rows[row]) {
      const ink = s.fit.ink, bgc = s.fit.bg;
      const inkHex = HEX[ink[0]] + HEX[ink[1]] + HEX[ink[2]];
      const bgHex = HEX[bgc[0]] + HEX[bgc[1]] + HEX[bgc[2]];
      for (let x = s.x0; x < s.x1; x++) {
        const top = ((row * 2) * VIEW_W + x) * 3;
        const bottom = ((row * 2 + 1) * VIEW_W + x) * 3;
        const topInk = dist3(bitmapRGB[top], bitmapRGB[top + 1], bitmapRGB[top + 2], ...ink)
          <= dist3(bitmapRGB[top], bitmapRGB[top + 1], bitmapRGB[top + 2], ...bgc);
        const bottomInk = dist3(bitmapRGB[bottom], bitmapRGB[bottom + 1], bitmapRGB[bottom + 2], ...ink)
          <= dist3(bitmapRGB[bottom], bitmapRGB[bottom + 1], bitmapRGB[bottom + 2], ...bgc);
        // Every cell still carries shading. It fills the exact line box; the
        // glyph only paints the selected halves over it, avoiding scan lines.
        g.chars[gy][1 + x] = topInk ? (bottomInk ? '█' : '▀') : (bottomInk ? '▄' : ' ');
        g.colors[gy][1 + x] = inkHex;
        g.bgs[gy][1 + x] = bgHex;
      }
    }
  }
}

// ─── Framebuffer → playable high-contrast grid ───────────────────────
// One foreground/background pair makes every picture row a single run. The
// glyph remains free inside that run, so all sixteen quadrant blocks preserve
// a literal 188 × 58 silhouette while the document stays cheap to repaint.

/** Glyphs, as which of a cell's four QUADRANTS the ink covers: top-left,
 *  top-right, bottom-left, bottom-right.
 *
 *  This is where the resolution comes from, and it is free. A run breaks on an
 *  ink or shading change and never on a glyph. A cell drawn with `▀` carries
 *  two sub-pixels; drawn with
 *  the quadrant blocks it carries FOUR, because all sixteen 2x2 patterns have
 *  a character. So the 94 x 29 cells sample the framebuffer at 188 x 58
 *  instead of 94 x 58, for no extra runs and no extra colours.
 *
 *  Still solid only — every quadrant is fully one endpoint or the other. The
 *  shade characters (░▒▓) are deliberately absent: they approximate TONE
 *  rather than carrying detail, and at the size this ships (a 4.8 x 11.45pt
 *  cell) they read as dots, not as a blend. `▚` and `▞` appear here for the
 *  opposite reason to why they were dropped before — as two of the sixteen
 *  real 2x2 patterns, not as a 50% grey. */
const GLYPHS = [
  [' ', 0, 0, 0, 0], ['▘', 1, 0, 0, 0], ['▝', 0, 1, 0, 0], ['▀', 1, 1, 0, 0],
  ['▖', 0, 0, 1, 0], ['▌', 1, 0, 1, 0], ['▞', 0, 1, 1, 0], ['▛', 1, 1, 1, 0],
  ['▗', 0, 0, 0, 1], ['▚', 1, 0, 0, 1], ['▐', 0, 1, 0, 1], ['▜', 1, 1, 0, 1],
  ['▄', 0, 0, 1, 1], ['▙', 1, 0, 1, 1], ['▟', 0, 1, 1, 1], ['█', 1, 1, 1, 1],
];
const FAST_INK = 'E8E2D8', FAST_BG = '181512';

// Luma-weighted distance: the eye resolves green far better than blue, so an
// even metric spends its budget where it will not be seen.
const LR = 0.30, LG = 0.59, LB = 0.11;
const dist3 = (ar, ag, ab, br, bg, bb) =>
  LR * Math.abs(ar - br) + LG * Math.abs(ag - bg) + LB * Math.abs(ab - bb);

// ─── Per-frame auto-exposure ──────────────────────────────────────────
// Doom is dark and its exposure moves — median luminance is 0.110 in a
// corridor and 0.185 in an open lit area — so one fixed curve serves neither.
// The binary projection needs the room and the weapon to keep their separation
// as lighting changes, so its threshold is relative to each frame.
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
 *  [(row * 2 + half) * QUAD_W + qx] * 3. This is the 188 x 58 picture. */
const quadRGB = new Float32Array(HALF_H * QUAD_W * 3);

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

/** Paint one frame as a high-contrast quadrant projection.
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
  }

  // One stable endpoint pair for the entire 188 x 58 picture. The four free
  // quadrant bits are a literal high-resolution monochrome projection of the
  // exposed Doom framebuffer: room edges, sprites, weapon and HUD stay crisp,
  // while every picture row is one compact shaded run. This is the playable
  // contract — image structure and a 32-span total, not a colour photograph
  // that costs a second to repaint.
  for (let row = 0; row < FIELD_ROWS; row++) {
    const gy = FIELD_TOP + row;
    for (let x = 0; x < VIEW_W; x++) {
      const q = [((row * 2) * QUAD_W + x * 2) * 3, ((row * 2) * QUAD_W + x * 2 + 1) * 3,
                 ((row * 2 + 1) * QUAD_W + x * 2) * 3, ((row * 2 + 1) * QUAD_W + x * 2 + 1) * 3];
      let mask = 0;
      for (let i = 0; i < 4; i++) {
        const o = q[i];
        const luma = 0.2126 * quadRGB[o] + 0.7152 * quadRGB[o + 1] + 0.0722 * quadRGB[o + 2];
        if (luma >= 127.5) mask |= 1 << i;
      }
      g.chars[gy][1 + x] = GLYPHS[mask][0];
      g.colors[gy][1 + x] = FAST_INK;
      g.bgs[gy][1 + x] = FAST_BG;
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

/** The two ways this cartridge can put Doom on a Word paragraph.
 *
 *  `8bit` is the legacy internal identifier for the playable high-contrast
 *  projection: one stable endpoint pair and a literal 188×58 quadrant
 *  silhouette, represented by three authored runs / 32 rendered spans.
 *  `bitmap` keeps open-palette detail in up to 900 picture segments and is a
 *  paused-frame inspection mode, not a playable renderer. */
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
    write(g, cy, cx, title, HUD_INK);
    write(g, cy + 2, cx, phase, HUD_INK);
    write(g, cy + 3, cx, '[' + '█'.repeat(bar) + '░'.repeat(40 - bar) + ']', HUD_INK);
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

/** Let both bezels merge into the picture.
 *
 *  Every row of this paragraph starts with a box-drawing character so the
 *  editor's markdown blur-commit can never read a row as a heading or a
 *  bullet. That safety is a property of the CHARACTER, not of its colour — but
 *  giving the column its own ink cost it its own run on every picture row —
 *  one run per row of the frame, for a one-cell grey line. Painting the bezel in its
 *  neighbour's colours keeps the character exactly where it was and merges the
 *  run away. The same rule on the right replaced the tiny control panel: those
 *  controls are now a legible paragraph above the screen, and the viewport gets
 *  the whole document width. */
function mergeBezelsIntoPicture(g) {
  for (let row = 0; row < FIELD_ROWS; row++) {
    const y = FIELD_TOP + row;
    g.colors[y][0] = g.colors[y][1];
    g.bgs[y][0] = g.bgs[y][1];
    g.colors[y][COLS - 1] = g.colors[y][COLS - 2];
    g.bgs[y][COLS - 1] = g.bgs[y][COLS - 2];
  }
}

  function render() {
    const g = makeGrid();
    if (status === 'playing' && handle) {
      if (projection === '8bit') paintFramebuffer8Bit(g, handle.framebuffer);
      else paintFramebuffer(g, handle.framebuffer);
      paintedFrames++;
      drawChrome(g, projection === '8bit'
        ? `DOOM — ${handle.title?.trim() ?? 'id Software engine'} — HIGH-CONTRAST · playable · frame ${handle.frames}`
        : `DOOM — ${handle.title?.trim() ?? 'id Software engine'} — BITMAP · detailed · frame ${handle.frames}`);
      mergeBezelsIntoPicture(g);
    } else {
      drawLoading(g);
      drawChrome(g, 'DOOM — the real engine, compiled to JavaScript');
    }
    return { grid: g, bg: BG, metrics: METRICS };
  }

  return {
    name: 'doom',
    label: '☩ DOOM',
    controls: [
      'CONTROLS · MOVE W/S · STRAFE A/D',
      'TURN ←/→ · FIRE SPACE · USE E',
      'RUN SHIFT · MENU Q · MAP M · WEAPON 1–7',
      'PROJECTION P · PAUSE/EDIT ESC',
    ],
    caption:
      'The **actual** game: id Software’s Doom engine — GPL-2.0, compiled to JavaScript by ' +
      '[doomgeneric](https://github.com/grubbyplaya/doomgenericjs) — running on Freedoom’s ' +
      'BSD-licensed game data. Its 320×200 framebuffer is redrawn into this Word paragraph every ' +
      'frame, and what it costs is *colored runs*: a run breaks when the ink or the `w:shd` ' +
      'shading changes, never when the glyph does. So the default **high-contrast** projection ' +
      'uses one stable pair for the entire picture, then lets every cell choose its own arrangement ' +
      'with a quadrant block — all sixteen 2×2 patterns have a character. The 94 × 29 picture ' +
      'cells therefore retain a **188 × 58** silhouette while the authored frame stays at three ' +
      'runs and the rendered paragraph at 32 spans. That is the playable contract: the room, ' +
      'sprites, weapon and HUD remain distinct while the document sustains about ten complete ' +
      'repaints a second. The controls are four separate 18pt document paragraphs, so they stay ' +
      'genuinely readable without being converted on every frame. ' +
      '**P** switches to the detailed **bitmap**: an unrestricted framebuffer-derived palette, ' +
      'two vertical samples per cell and far more rendered spans; it is a paused-frame inspection ' +
      'mode, not the playable renderer. Move **W/S** · strafe ' +
      '**A/D** · turn **←/→** · **Space** ' +
      'fires · **E** opens · **Q** is Doom’s own menu. **Esc** pauses — and then it is only a ' +
      'document again: put your caret in the frame, Undo rewinds it, Save downloads it as .docx.',
    hint: '<b>WASD</b> move · <b>←/→</b> turn · <b>Space</b> fire · <b>E</b> open · <b>Q</b> Doom’s menu · <b>P</b> switches projection — high contrast is the playable ~10 fps path; bitmap is for paused detail.',
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

// Printable ASCII projection of Doom's 320x200 BGRA framebuffer. No blocks,
// braille, images, cell shading or game-state HUD reconstruction: every mark
// is a character sampled from the same picture the native renderer displays.
export const ASCII_COLS = 320;
export const ASCII_ROWS = 200;
export const ASCII_METRICS = { sz: 5, lineTwips: 34, bold: true, spacingTwips: -2 };

// Tone lives in the glyph, hue in the ink. Separating them lets textured grey
// walls use one Word run while preserving a different intensity in every cell.
// Coverage is measured in repeated rows in the actual editor at the authored
// size, tracking and line pitch (Canvas Mono, synthetic bold, DPR 2), normalized
// to the densest glyph. Isolated, large-font measurements underestimate the
// brightness of small packed glyphs and wash out dark outlines. Keep source
// contrast linear; lifting shadows makes overlays look transparent.
export const ASCII_RAMP = ' .,:;ris235hSXGA&9HB#M@';
const coverage = [0, 130, 188, 258, 301, 475, 577, 642, 690, 692, 696,
  798, 704, 761, 770, 896, 829, 808, 869, 947, 972, 1000, 973];
const tones = Array.from({ length: 256 }, (_, value) => {
  const target = value / 255 * 1000;
  let best = 0;
  for (let i = 1; i < coverage.length; i++)
    if (Math.abs(coverage[i] - target) < Math.abs(coverage[best] - target)) best = i;
  return ASCII_RAMP[best];
});
export const ASCII_PALETTE = [
  'FFFFFF', 'FFD6AA', 'FFB878', 'FF944C', 'FF6438', 'FF3030',
  'FF9CB8', 'E8ACFF', 'B4B4FF', '80B8FF', '68DCFF', '80FFD0',
  'A0FF80', 'DCFF80', 'FFE878', 'FFECCB', 'FF0000', '00FF00', '0000FF',
];
const palette = ASCII_PALETTE.map(hex => [0, 2, 4].map(i => parseInt(hex.slice(i, i + 2), 16)));
const paletteNorm = palette.map(([r, g, b]) => .3 * r * r + .59 * g * g + .11 * b * b);
// Bound error for EVERY visible cell, not just the average of a merged run.
// Otherwise a small contrasting stroke can be sacrificed to a long backdrop.
const compatible = palette.map(a => palette.reduce((mask, b, i) =>
  .3 * (a[0] - b[0]) ** 2 + .59 * (a[1] - b[1]) ** 2 + .11 * (a[2] - b[2]) ** 2
    <= 80 ** 2 ? mask | (1 << i) : mask, 0));
const anyInk = (1 << palette.length) - 1;
const lookup = new Int8Array(32768).fill(-1);

function inkFor(r, g, b) {
  const peak = Math.max(r, g, b);
  if (peak < 8) return 0;
  r = r * 255 / peak; g = g * 255 / peak; b = b * 255 / peak;
  const key = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
  if (lookup[key] >= 0) return lookup[key];
  // Evaluate the bin centre so a cached hue never depends on which frame
  // happened to visit this bin first.
  r = ((key >> 10) & 31) * 8 + 4;
  g = ((key >> 5) & 31) * 8 + 4;
  b = (key & 31) * 8 + 4;
  let best = 0, error = Infinity;
  for (let i = 0; i < palette.length; i++) {
    const p = palette[i];
    const e = (r - p[0]) ** 2 * .3 + (g - p[1]) ** 2 * .59 + (b - p[2]) ** 2 * .11;
    if (e < error) { best = i; error = e; }
  }
  lookup[key] = best;
  return best;
}

/** One source pixel per glyph, including the original status bar, menus and
 * automap. No downsample or reconstructed game-state labels can lose a stroke. */
export function asciiFramebuffer(fb) {
  if (fb.length !== 320 * 200 * 4) throw new RangeError('Expected a 320x200 BGRA framebuffer');
  const chars = [], colors = [], segments = [];
  for (let y = 0; y < ASCII_ROWS; y++) {
    const row = [], inks = [];
    let last = null;
    const sy = Math.floor((y + .5) * 200 / ASCII_ROWS);
    for (let x = 0; x < ASCII_COLS; x++) {
      const sx = Math.floor((x + .5) * 320 / ASCII_COLS);
      const o = (sy * 320 + sx) * 4;
      const r = fb[o + 2], g = fb[o + 1], b = fb[o];
      const char = tones[Math.max(r, g, b)];
      row.push(char);
      const wanted = inkFor(r, g, b);
      inks.push(ASCII_PALETTE[wanted]);
      const weight = char === ' ' ? 0 : (Math.max(r, g, b) / 255) ** 2;
      const allowed = char === ' ' ? anyInk : compatible[wanted];
      const rgb = palette[wanted];
      if (!last || last.ink !== wanted) {
        const next = { y, x0: x, x1: x, ink: wanted, w: 0, r: 0, g: 0, b: 0, square: 0,
          error: 0, allowed, prev: last, next: null, alive: true };
        if (last) last.next = next;
        segments.push(next); last = next;
      }
      last.x1 = x + 1;
      last.allowed &= allowed;
      last.w += weight;
      last.r += weight * rgb[0]; last.g += weight * rgb[1]; last.b += weight * rgb[2];
      last.square += weight * paletteNorm[wanted];
    }
    // A printable guard prevents a paused row being interpreted as Markdown
    // syntax, and keeps completely black rows from splitting the paragraph.
    row.unshift('|'); inks.unshift(inks[0]);
    chars.push(row); colors.push(inks);
  }
  allocateColorRuns(segments, colors, 700);
  for (const row of colors) row[0] = row[1];
  return { chars, colors };
}

// Allocate the limited ink changes where they preserve the most visible hue.
// Every pixel keeps its own glyph/tone. The frame-wide run target avoids fixed
// tiles, but is soft: stop merging when no ink can satisfy every cell's color
// bound. High-contrast pictures may need more runs to keep their edges intact.
function allocateColorRuns(segments, colors, budget) {
  const heap = [];
  const push = e => {
    let i = heap.length; heap.push(e);
    while (i) { const p = (i - 1) >> 1; if (heap[p].cost <= e.cost) break;
      heap[i] = heap[p]; i = p; }
    heap[i] = e;
  };
  const pop = () => {
    const first = heap[0], last = heap.pop();
    if (heap.length) { let i = 0;
      while (i * 2 + 1 < heap.length) {
        let child = i * 2 + 1;
        if (child + 1 < heap.length && heap[child + 1].cost < heap[child].cost) child++;
        if (heap[child].cost >= last.cost) break;
        heap[i] = heap[child]; i = child;
      }
      heap[i] = last;
    }
    return first;
  };
  const offer = a => {
    const b = a?.next;
    if (!b) return;
    const allowed = a.allowed & b.allowed;
    if (!allowed) return;
    const w = a.w + b.w, r = a.r + b.r, g = a.g + b.g, blue = a.b + b.b;
    const square = a.square + b.square;
    let ink = 0, error = Infinity;
    for (let i = 0; i < palette.length; i++) {
      if (!(allowed & (1 << i))) continue;
      const p = palette[i];
      const e = Math.max(0, square + w * paletteNorm[i]
        - 2 * (.3 * p[0] * r + .59 * p[1] * g + .11 * p[2] * blue));
      if (e < error) { ink = i; error = e; }
    }
    push({ a, b, cost: error - a.error - b.error, w, r, g, blue, square, ink, error, allowed });
  };
  for (const s of segments) offer(s);
  let count = segments.length;
  const merged = [];
  while (count > budget && heap.length) {
    const e = pop(), { a, b } = e;
    if (!a.alive || !b.alive || a.next !== b) continue;
    const next = { w: e.w, r: e.r, g: e.g, b: e.blue, square: e.square, ink: e.ink, error: e.error, allowed: e.allowed,
      y: a.y, x0: a.x0, x1: b.x1, prev: a.prev, next: b.next, alive: true };
    a.alive = b.alive = false;
    if (next.prev) next.prev.next = next;
    if (next.next) next.next.prev = next;
    merged.push(next); count--;
    offer(next.prev); offer(next);
  }
  for (const s of [...segments, ...merged]) if (s.alive)
    colors[s.y].fill(ASCII_PALETTE[s.ink], s.x0 + 1, s.x1 + 1);
}

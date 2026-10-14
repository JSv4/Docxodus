// Printable ASCII projection of Doom's 320x200 BGRA framebuffer. No blocks,
// braille, images, cell shading or game-state HUD reconstruction: every mark
// is a character sampled from the same picture the native renderer displays.
export const ASCII_COLS = 320;
export const ASCII_ROWS = 200;
export const ASCII_METRICS = { sz: 4, lineTwips: 34, bold: true, spacingTwips: 4, asciiOnly: true };

// Tone lives in the glyph, hue in the ink. Separating them lets textured grey
// walls use one Word run while preserving intensity in every cell. Use glyphs
// whose ink stays between the baseline and the cap height: descenders and tall
// glyphs in an undersized line box bleed into neighboring source pixels.
// At 2pt the selected ink fits the 1.7pt row pitch, including synthetic bold;
// positive tracking keeps the original pixel width without enlarging the ink.
// Coverage is measured in the pinned Canvas Mono font, in packed native-editor
// rows at DPR 1 and 2. Exclude glyphs whose normalized coverage shifts by more
// than 5% between the two densities, then average. Keep contrast linear.
export const ASCII_RAMP = " .-':\"^*+?>z<iL71TnIhZVwP4ARWM";
const coverage = [0, 113, 140, 150, 239, 280, 315, 416, 434, 481, 492, 500, 501, 506, 508,
  565, 575, 578, 593, 623, 679, 688, 699, 700, 735, 745, 781, 847, 975, 999];
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
    <= 60 ** 2 ? mask | (1 << i) : mask, 0));
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
 * automap. No downsample or reconstructed game-state labels can lose a stroke.
 *
 * Extend each ink run to the longest prefix that satisfies EVERY cell's color
 * bound, then choose its least-squares ink. Taking the longest valid prefix
 * minimizes the number of runs: no valid first run can end later, and removing
 * that prefix cannot make the remaining suffix require more runs. This linear
 * pass replaces the allocation-heavy, frame-wide merge heap. There is no hard
 * budget that can erase a small contrasting stroke to save a formatting run. */
export function asciiFramebuffer(fb) {
  if (fb.length !== 320 * 200 * 4) throw new RangeError('Expected a 320x200 BGRA framebuffer');
  const chars = [], colors = [];
  for (let y = 0; y < ASCII_ROWS; y++) {
    const row = new Array(ASCII_COLS + 1), inks = new Array(ASCII_COLS + 1);
    let start = 0, allowed = anyInk, w = 0, red = 0, green = 0, blue = 0;
    const flush = end => {
      let best = 0, error = Infinity;
      for (let i = 0; i < palette.length; i++) {
        if (!(allowed & (1 << i))) continue;
        const p = palette[i];
        const e = w * paletteNorm[i] - 2 * (.3 * p[0] * red + .59 * p[1] * green + .11 * p[2] * blue);
        if (e < error) { best = i; error = e; }
      }
      inks.fill(ASCII_PALETTE[best], start + 1, end + 1);
    };
    for (let x = 0; x < ASCII_COLS; x++) {
      const o = (y * ASCII_COLS + x) * 4;
      const r = fb[o + 2], g = fb[o + 1], b = fb[o];
      const peak = Math.max(r, g, b), char = tones[peak];
      row[x + 1] = char;
      const wanted = inkFor(r, g, b);
      const next = char === ' ' ? anyInk : compatible[wanted];
      if (!(allowed & next)) {
        flush(x);
        start = x; allowed = anyInk; w = red = green = blue = 0;
      }
      allowed &= next;
      const weight = char === ' ' ? 0 : (peak / 255) ** 2, rgb = palette[wanted];
      w += weight; red += weight * rgb[0]; green += weight * rgb[1]; blue += weight * rgb[2];
    }
    flush(ASCII_COLS);
    // A printable guard prevents a paused row being interpreted as Markdown
    // syntax, and keeps completely black rows from splitting the paragraph.
    row[0] = '|'; inks[0] = inks[1];
    chars.push(row); colors.push(inks);
  }
  return { chars, colors };
}

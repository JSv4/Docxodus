// The DOCX Observatory's phenomena, shared by both host pages:
// ascii-animation.html (bespoke exhibit chrome, renderBlock + replaceWith) and
// ascii-animation-editor.html (the shipped ribbon editor, editor.refresh()).
// A scene fills a COLS×ROWS cell grid — chars[row][col] + colors[row][col]
// (hex RRGGBB, ignored for spaces) — and `bg` becomes the canvas paragraph's
// w:shd fill. frameXml() turns a grid into the canvas paragraph's OOXML;
// seedObservatory() builds the document itself through the agentic surface.

// ─── Canvas geometry ──────────────────────────────────────────────────
// 92 columns of 8pt Courier New ≈ 6.1in — fits the blank doc's 6.5in text
// column without wrapping. Line rule "exact" 200 twips (10pt) gives a
// terminal-ish cell aspect of ~2:1 (height:width).
export const COLS = 92, ROWS = 26;
export const FONT = 'Courier New';
export const SZ = 16;          // w:sz half-points → 8pt
export const LINE_TWIPS = 200; // 10pt exact line height

// ─── Tiny deterministic noise (no Math.random: frames must be a pure
//     function of t so tests and repro stay honest) ─────────────────────
function hash2(x, y) {
  let h = (x * 374761393 + y * 668265263) | 0;
  h = Math.imul(h ^ (h >>> 13), 1274126177);
  return ((h ^ (h >>> 16)) >>> 0) / 4294967296;
}

function makeGrid() {
  const chars = [], colors = [];
  for (let y = 0; y < ROWS; y++) {
    chars.push(new Array(COLS).fill(' '));
    colors.push(new Array(COLS).fill('FFFFFF'));
  }
  return { chars, colors };
}

const oceanScene = {
  name: 'ocean', label: 'Ocean swell', bg: '071426',
  gen(t) {
    const g = makeGrid();
    const surf = (x) =>
      ROWS * 0.42
      + 2.6 * Math.sin(x * 0.110 + t * 0.90)
      + 1.7 * Math.sin(x * 0.053 - t * 0.55)
      + 0.9 * Math.sin(x * 0.023 + t * 0.21);
    // GLYPHS follow the true wavy surface; INK is banded by flat row depth.
    // Every color change inside a row is its own w:r and the converter pays
    // ~1 ms per run, so letting color bands zigzag along the wave multiplies
    // runs ~2×. Flat bands keep most rows at ONE run — the silhouette still
    // reads because the glyphs carry it.
    const mean = ROWS * 0.42;
    const rowInk = (y) => {
      const d = y - mean;
      return d < 1.5 ? '7FD7F0' : d < 4 ? '4FB3DC' : d < 7 ? '2E86B8' : '1A5E8C';
    };
    for (let x = 0; x < COLS; x++) {
      const ys = surf(x);
      const slope = surf(x + 1) - surf(x - 1);
      for (let y = 0; y < ROWS; y++) {
        const d = y - ys;
        if (d < -1) {
          // Night sky: sparse twinkling stars, one shared color.
          const r = hash2(x, y);
          if (r < 0.012) {
            g.chars[y][x] = (r * 900 + t * 1.4) % 2 < 1 ? '.' : '+';
            g.colors[y][x] = 'D8D2B8';
          }
        } else if (d < 0.5) {
          // Surface: crest tips foam white, steep faces break into slashes.
          const crest = surf(x) < surf(x - 1) - 0.35 && surf(x) < surf(x + 1) - 0.35;
          g.chars[y][x] = crest ? '^' : slope > 0.9 ? '\\' : slope < -0.9 ? '/' : '~';
          g.colors[y][x] = crest ? 'F3FBFF' : rowInk(y);
        } else if (d < 2.2) {
          const n = Math.sin(x * 0.35 + t * 2.1 + y * 1.7);
          g.chars[y][x] = n > 0.25 ? '~' : '-';
          g.colors[y][x] = rowInk(y);
        } else if (d < 5.5) {
          const n = Math.sin(x * 0.21 - t * 1.2 + y * 2.3);
          g.chars[y][x] = n > 0.45 ? '=' : ':';
          g.colors[y][x] = rowInk(y);
        } else {
          const r = hash2(x, y * 7);
          g.chars[y][x] = r < 0.16 ? ':' : r < 0.4 ? '.' : ' ';
          g.colors[y][x] = rowInk(y);
        }
      }
    }
    // Two fish on their commutes, at fixed depths below the average surface.
    const fishRows = [Math.floor(ROWS * 0.62), Math.floor(ROWS * 0.80)];
    const sprites = ['><>', '<><'];
    fishRows.forEach((fy, i) => {
      const span = COLS + 8;
      let fx = Math.floor((t * (7 + i * 4) + i * 37) % span) - 4;
      if (i === 1) fx = COLS - 1 - fx; // second fish swims the other way
      const s = sprites[i];
      for (let k = 0; k < s.length; k++) {
        const x = fx + k;
        if (x >= 0 && x < COLS) { g.chars[fy][x] = s[k]; g.colors[fy][x] = 'F4B860'; }
      }
    });
    return g;
  },
};

const rippleScene = {
  name: 'ripples', label: 'Pond ripples', bg: '06131F',
  gen(t) {
    const g = makeGrid();
    // Deterministic drop schedule: one drop every 2.1s, position hashed
    // from the drop ordinal, each ringing for ~6s.
    const drops = [];
    const first = Math.max(0, Math.floor((t - 6) / 2.1));
    for (let k = first; k * 2.1 <= t; k++) {
      drops.push({
        x: 6 + hash2(k, 11) * (COLS - 12),
        y: 2 + hash2(k, 29) * (ROWS - 4),
        age: t - k * 2.1,
      });
    }
    // Five glyph intensities but only TWO ink colors (+foam at the very
    // peak): each color change inside a row is a fresh w:r, so the palette
    // is spent where it shows.
    const ramp = ['.', ':', 'o', 'O', '@'];
    const cols = ['2E86B8', '2E86B8', '9BE0F2', '9BE0F2', 'F3FBFF'];
    for (let y = 0; y < ROWS; y++) {
      for (let x = 0; x < COLS; x++) {
        let h = 0;
        for (const d of drops) {
          const dx = x - d.x, dy = (y - d.y) * 2; // ×2: cells are ~2:1 tall
          const r = Math.sqrt(dx * dx + dy * dy);
          if (r > d.age * 8 + 4) continue; // outside the wavefront
          h += Math.sin(r * 0.9 - d.age * 6) * Math.exp(-r * 0.085) * Math.exp(-d.age * 0.35);
        }
        const a = Math.abs(h);
        if (a < 0.055) {
          // Calm water carries a faint static weave so the rings read as
          // water. Same ink as the ring base so it merges into their runs.
          if ((x * 3 + y * 7) % 23 === 0) { g.chars[y][x] = '·'; g.colors[y][x] = '2E86B8'; }
        } else {
          const i = Math.min(ramp.length - 1, Math.floor(a * 7));
          g.chars[y][x] = ramp[i];
          g.colors[y][x] = cols[i];
        }
      }
    }
    return g;
  },
};

const rainScene = {
  name: 'rain', label: 'Squall', bg: '0A1220',
  gen(t) {
    const g = makeGrid();
    const groundY = ROWS - 1;
    // Ground line.
    for (let x = 0; x < COLS; x++) { g.chars[groundY][x] = '_'; g.colors[groundY][x] = '35567A'; }
    for (let x = 0; x < COLS; x++) {
      const density = hash2(x, 5);
      if (density < 0.28) continue; // dry lane
      const speed = 16 + hash2(x, 7) * 14;       // rows/second
      const cycle = ROWS + 6 + hash2(x, 13) * 20; // fall + regroup gap
      const p = (hash2(x, 3) * cycle + t * speed) % cycle;
      const head = Math.floor(p);
      // Streak glyphs share ONE ink (glyph shape carries the fade): with a
      // single color, a whole row of scattered drops merges into one w:r.
      if (head < groundY) {
        g.chars[head][x] = '|'; g.colors[head][x] = '7FA8CC';
        if (head - 1 >= 0) { g.chars[head - 1][x] = ':'; g.colors[head - 1][x] = '7FA8CC'; }
        if (head - 2 >= 0) { g.chars[head - 2][x] = '.'; g.colors[head - 2][x] = '7FA8CC'; }
      } else if (p < groundY + 1.2) {
        // Splash: a beat of white at the ground.
        const sy = groundY - 1;
        g.chars[sy][x] = 'o'; g.colors[sy][x] = 'D7E7F5';
        if (x > 0 && g.chars[sy][x - 1] === ' ') { g.chars[sy][x - 1] = '.'; g.colors[sy][x - 1] = '7FA8CC'; }
        if (x < COLS - 1 && g.chars[sy][x + 1] === ' ') { g.chars[sy][x + 1] = '.'; g.colors[sy][x + 1] = '7FA8CC'; }
      }
    }
    // A lightning bolt every ~6.3s, alive for 0.22s, zigzagging down.
    const k = Math.floor(t / 6.3);
    if (t - k * 6.3 < 0.22) {
      let bx = Math.floor(8 + hash2(k, 41) * (COLS - 16));
      for (let y = 0; y < groundY - 3; y++) {
        const j = hash2(k * 100 + y, 43);
        const step = j < 0.35 ? -1 : j < 0.7 ? 1 : 0;
        bx = Math.max(1, Math.min(COLS - 2, bx + step));
        g.chars[y][bx] = step < 0 ? '/' : step > 0 ? '\\' : '|';
        g.colors[y][bx] = 'FFF7B0';
      }
    }
    return g;
  },
};

const fireScene = {
  name: 'fire', label: 'Hearth fire', bg: '0D0603',
  heat: null,
  reset() { this.heat = null; },
  gen(t) {
    // Classic propagation fire: heat rises from a stoked bottom row and
    // cools on the way up. Stateful (the grid IS the phenomenon), but all
    // randomness is hashed off (t, x) so a given timeline replays identically.
    const H = ROWS, W = COLS;
    if (!this.heat) {
      this.heat = [];
      for (let y = 0; y < H; y++) this.heat.push(new Float32Array(W));
    }
    const heat = this.heat;
    const stoke = Math.floor(t * 24); // discrete stoking ticks
    for (let x = 0; x < W; x++) {
      // Fuel bed breathes: edges cooler, center roaring.
      const edge = Math.sin((x / W) * Math.PI);
      heat[H - 1][x] = 26 + 10 * edge * (0.7 + 0.3 * Math.sin(x * 0.5 + t * 3)) * (0.75 + 0.5 * hash2(x, stoke));
    }
    for (let y = 0; y < H - 1; y++) {
      for (let x = 0; x < W; x++) {
        const drift = Math.round((hash2(x * 7 + y, stoke) - 0.5) * 2);
        const sx = Math.max(0, Math.min(W - 1, x + drift));
        const cool = hash2(x + y * 31, stoke + 1) * 3.4;
        heat[y][x] = Math.max(0, heat[y + 1][sx] - cool);
      }
    }
    // Glyphs flicker on RAW per-cell heat; ink follows a horizontally
    // smoothed heat field. Raw heat dithers between palette bands cell to
    // cell, and every color change is its own w:r — unsmoothed this frame
    // costs ~1300 runs (≈550 ms renders); banded it sits nearer 150.
    const ramp = [' ', '.', ':', '+', '=', '*', '#', '%', '@'];
    const cols = ['000000', '7A1E06', 'C2410C', 'FB923C', 'FEF3C7'];
    const g = makeGrid();
    const smooth = new Float32Array(W);
    for (let y = 0; y < H; y++) {
      for (let x = 0; x < W; x++) {
        let acc = 0, n = 0;
        for (let k = -4; k <= 4; k++) {
          const sx = x + k;
          if (sx >= 0 && sx < W) { acc += heat[y][sx]; n++; }
        }
        smooth[x] = acc / n;
      }
      for (let x = 0; x < W; x++) {
        const i = Math.max(0, Math.min(8, Math.floor(heat[y][x] / 4)));
        const ci = Math.max(0, Math.min(4, Math.round(smooth[x] / 8)));
        g.chars[y][x] = ramp[i];
        g.colors[y][x] = cols[ci];
      }
    }
    return g;
  },
};

export const SCENES = [oceanScene, rippleScene, rainScene, fireScene];

// ─── Frame → OOXML ────────────────────────────────────────────────────
const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** Merge each row's cells into runs: a new run only where a visible glyph
 *  changes color (spaces piggyback on the current run — background ink
 *  is invisible, so they never force a split). */
function rowRuns(chars, colors) {
  const runs = [];
  let text = '', color = null;
  for (let x = 0; x < COLS; x++) {
    const ch = chars[x];
    if (ch !== ' ' && colors[x] !== color && color !== null) {
      runs.push([text, color]); text = '';
    }
    if (ch !== ' ') color = colors[x] ?? color;
    text += ch;
  }
  runs.push([text, color ?? '9CB3C9']);
  return runs;
}

/** The whole frame as one w:p: the captured opening tag (which carries the
 *  paragraph's Unid — THE thing that keeps the anchor stable across frames),
 *  a pPr with shading + exact line height, then per row: colored runs joined
 *  by w:br. */
export function frameXml(openTag, grid, bg) {
  const parts = [openTag,
    '<w:pPr>',
    `<w:spacing w:before="0" w:after="0" w:line="${LINE_TWIPS}" w:lineRule="exact"/>`,
    `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`,
    '</w:pPr>'];
  let runs = 0;
  for (let y = 0; y < ROWS; y++) {
    if (y > 0) parts.push('<w:r><w:br/></w:r>');
    for (const [text, color] of rowRuns(grid.chars[y], grid.colors[y])) {
      runs++;
      parts.push(
        `<w:r><w:rPr><w:rFonts w:ascii="${FONT}" w:hAnsi="${FONT}" w:cs="${FONT}"/>` +
        `<w:color w:val="${color}"/><w:sz w:val="${SZ}"/><w:szCs w:val="${SZ}"/></w:rPr>` +
        `<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`);
    }
  }
  parts.push('</w:p>');
  return { xml: parts.join(''), runs };
}

// ─── The document itself ──────────────────────────────────────────────

/** Seed a freshly opened blank session with the Observatory document — title,
 *  canvas paragraph, caption, and a real footnote — entirely through the
 *  agentic editing surface, then capture the canvas paragraph's opening tag
 *  (it carries the Unid attribute plus every namespace declaration, so frames
 *  built from it keep the anchor alive — replaceXml reports the block as
 *  Modified, not Removed+Created). */
export function seedObservatory(session) {
  const check = (r, what) => {
    if (!r.success) throw new Error(`${what} failed: ${r.error?.code} ${r.error?.message}`);
    return r;
  };

  const firstP = session.findByKind('p', 'body')[0];
  if (!firstP) throw new Error('blank document has no body paragraph');
  const titleAnchor = firstP.id;
  check(session.replaceText(titleAnchor, 'THE DOCX OBSERVATORY'), 'title replaceText');
  check(session.setParagraphFormat(titleAnchor, { alignment: 'center', spacingAfter: 160 }), 'title format');
  check(session.applyFormat(titleAnchor, null, { bold: true, fontFamily: FONT, fontSizePts: 13, color: '1F2937' }), 'title run format');

  const canvasResult = check(session.insertParagraph(titleAnchor, 'after', '(warming up the sea…)'), 'canvas insert');
  const canvasAnchor = canvasResult.created[0].id;

  const captionResult = check(session.insertParagraph(canvasAnchor, 'after',
    `Procedural phenomena drawn ${COLS} columns × ${ROWS} rows into a single monospaced paragraph.`), 'caption insert');
  const captionAnchor = captionResult.created[0].id;
  check(session.setParagraphFormat(captionAnchor, { alignment: 'center', spacingBefore: 160 }), 'caption format');
  check(session.applyFormat(captionAnchor, null, { fontFamily: FONT, fontSizePts: 8, color: '6B7280' }), 'caption run format');

  // A real footnote, because this is a real document.
  check(session.insertFootnote(captionAnchor, 20, // after "Procedural phenomena"
    'Every frame is OOXML: colored runs and `w:br` breaks in one Word paragraph, swapped in by ' +
    '`DocxSession.raw.replaceXml` and re-rendered by `renderBlock`. Save mid-wave and the sea freezes in Word.'),
    'footnote');

  const seedXml = session.raw.getXml(canvasAnchor);
  const gt = seedXml.indexOf('>');
  let openTag = seedXml.slice(0, gt + 1);
  if (openTag.endsWith('/>')) openTag = openTag.slice(0, -2) + '>';

  return { titleAnchor, canvasAnchor, captionAnchor, openTag };
}

// THE DOCX ARCADE — playable games rendered INTO a live Word document.
//
// Sibling of ascii-scenes.js (the Observatory) and same contract: the game
// screen is ONE Word paragraph, animated through the session API plus
// `DocxEditor.refresh()` — the editor's public "the session changed behind
// your back" seam, which reconciles exactly one block in continuous mode.
// DOOM can replace one
// native inline image or project every pixel to a colored ASCII character. Nothing
// is mounted over the document: pause (or click the screen) and the game is
// only a paragraph, Ctrl+Z rewinds frames, and Save stores the current frame
// in a real .docx.
//
// The game borrows the keyboard while running and gives it back when paused.
// Copy, paste, undo and DOCX save operate on the displayed document frame.
//
// This file's home is docs/demo/ for the same reason ascii-scenes.js lives
// there: the unified build stages it beside the editor package for Pages and tests. It is demo content, not library
// machinery, and is deliberately NOT shipped in the npm package.

import { COLS, ROWS, frameXml, createCanvasPin } from './ascii-scenes.js';
import { doomCart, DOOM_KEY_CODES } from './doom-cart.js';

// ─── Screen geometry ──────────────────────────────────────────────────
// The attract screen shares the Observatory’s 92×26 grid. DOOM uses its own
// full-resolution framebuffer metrics.
const INNER_W = COLS - 2;                 // 90 playfield columns

const BEZEL_INK = '33465B';
const HUD_INK = '9CB3C9';

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

/** Bezel and HUD for the attract screen. The bezel is load-bearing:
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
/** The keys the arcade claims from the document while a cartridge is playing.
 *  Exported so the headless logic checks can hold the touch pad to it: a pad
 *  button whose code is not in here sends a key the game never receives. */
export const ARCADE_KEY_CODES = new Set([
  'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown',
  'KeyW', 'KeyA', 'KeyS', 'KeyD', 'Space', 'KeyR',
  'ShiftLeft', 'ShiftRight',
  // Doom wants more of the keyboard than the ASCII cartridges do — a menu
  // key, a use key, weapon digits. Claiming Enter matters for a second
  // reason: unclaimed, it would split the screen paragraph in two.
  ...DOOM_KEY_CODES,
]);

function createInput(isPlaying) {
  const down = new Set();
  const pressed = new Set(); // keydown edges, consumed once per tick
  // Every press AND release, in order, for cartridges that need both edges.
  // The ASCII games read the held/pressed sets above and ignore this; Doom
  // reads only this, because its own input layer wants key-down and key-up
  // events one at a time. Capped so a paused tab cannot grow it forever.
  let transitions = [];
  const log = (code, isDown) => {
    if (transitions.length < 256) transitions.push({ code, down: isDown });
  };
  const onKeyDown = (e) => {
    if (!isPlaying() || e.metaKey || e.ctrlKey || e.altKey) return;
    if (!ARCADE_KEY_CODES.has(e.code)) return;
    e.preventDefault();
    e.stopPropagation();
    if (!down.has(e.code)) { pressed.add(e.code); log(e.code, true); }
    down.add(e.code);
  };
  // keyup always clears, playing or not — a key released while paused must
  // not stay latched into the next resume.
  const onKeyUp = (e) => { if (down.delete(e.code)) log(e.code, false); };
  const onBlur = () => {
    for (const code of down) log(code, false);
    down.clear(); pressed.clear();
  };
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
    /** Take the press/release log since the last call. */
    drain: () => { const out = transitions; transitions = []; return out; },
    /** Synthetic press/release — the on-screen touch pad and tests use this. */
    set: (code, isDown) => {
      if (isDown) {
        if (!down.has(code)) { pressed.add(code); log(code, true); }
        down.add(code);
      } else if (down.delete(code)) {
        log(code, false);
      }
    },
    dispose: () => {
      window.removeEventListener('keydown', onKeyDown, true);
      window.removeEventListener('keyup', onKeyUp, true);
      window.removeEventListener('blur', onBlur);
    },
  };
}

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

  // ─── Why the screen is fenced by two near-empty paragraphs ──────────
  // A single-block re-render does not render the block alone. The engine pads
  // each target with ONE REAL NEIGHBOUR on each side before converting it, so
  // that `w:contextualSpacing` resolves exactly as it would in a full render —
  // and those context clones are thrown away once the target's HTML is
  // extracted. That is correct, and it is cheap when the neighbours are small.
  //
  // The screen's neighbours were the title and the CAPTION, and the caption is
  // a long formatted prose paragraph (36 runs, with a footnote reference). So
  // every frame of every game converted it in full, purely as context, and
  // discarded the result. Fencing the screen with two one-character paragraphs
  // moves the caption out of that slot: measured 7.44 -> 9.25 fps on the Doom
  // cartridge, a 24% gain that costs the document two hairlines.
  const fence = (after, label) => {
    const res = check(session.insertParagraph(after, 'after', '\u00a0'), label);
    const id = res.created[0].id;
    check(session.setParagraphFormat(id, { spacingBefore: 0, spacingAfter: 0 }), `${label} format`);
    check(session.applyFormat(id, null, { fontFamily: 'Courier New', fontSizePts: 1 }), `${label} run format`);
    return id;
  };

  // Controls have to remain readable at the size the document is actually
  // shown. Putting Doom's key list inside its screen made the words technically
  // present but not honestly legible. Keep them as large 18pt document text,
  // outside the frame's hot paragraph and behind the tiny context fence, so
  // every repaint still converts only the screen and its one-character
  // neighbours.
  // Four deliberately short paragraphs make the complete map fit at 18pt in
  // fixed-layout renderers that do not reflow one long OOXML run. They stay
  // outside the screen's conversion fence, so this costs nothing per frame.
  const controlsAnchors = [];
  let controlsAfter = titleAnchor;
  for (let i = 0; i < 4; i++) {
    const controlsResult = check(
      session.insertParagraph(controlsAfter, 'after', i === 0 ? 'CONTROLS · loading cartridge…' : '\u00a0'),
      `controls line ${i + 1} insert`);
    const anchor = controlsResult.created[0].id;
    controlsAnchors.push(anchor);
    controlsAfter = anchor;
    check(session.setParagraphFormat(anchor, {
      alignment: 'center', spacingBefore: 0, spacingAfter: i === 3 ? 80 : 0,
    }), `controls line ${i + 1} format`);
    check(session.applyFormat(anchor, null,
      { bold: true, fontFamily: 'Courier New', fontSizePts: 18, color: '25324A' }),
    `controls line ${i + 1} run format`);
  }
  const controlsAnchor = controlsAnchors[0];

  const fenceAbove = fence(controlsAnchors[3], 'screen fence above');
  const canvasResult = check(session.insertParagraph(fenceAbove, 'after', '(inserting coin…)'), 'screen insert');
  const canvasAnchor = canvasResult.created[0].id;

  const fenceBelow = fence(canvasAnchor, 'screen fence below');

  const captionResult = check(session.insertParagraph(fenceBelow, 'after', 'loading cartridge…'), 'caption insert');
  const captionAnchor = captionResult.created[0].id;
  check(session.setParagraphFormat(captionAnchor, { alignment: 'center', spacingBefore: 160 }), 'caption format');
  check(session.applyFormat(captionAnchor, null, { fontFamily: 'Courier New', fontSizePts: 8, color: '6B7280' }), 'caption run format');

  // A real footnote, because the game screen is a real document.
  check(session.insertFootnote(captionAnchor, 7, // after "loading"
    'Every frame is OOXML document content: ASCII frames author colored runs and `w:br` ' +
    'breaks. Doom can use the same text path or replace one native inline image through the public ' +
    'session API. `DocxEditor.refresh()` repaints one block incrementally. Pause and Save: the frame ' +
    'downloads as a real .docx.'), 'footnote');

  const seedXml = session.raw.getXml(canvasAnchor);
  const gt = seedXml.indexOf('>');
  let openTag = seedXml.slice(0, gt + 1);
  if (openTag.endsWith('/>')) openTag = openTag.slice(0, -2) + '>';
  return { titleAnchor, controlsAnchor, controlsAnchors, canvasAnchor, captionAnchor, fenceBelow, openTag };
}

// ─── The attract screen ───────────────────────────────────────────────
// "OS LEGAL presents DOCXODUS" — the arcade's title card, drawn on the SAME
// canvas paragraph as the games (starfield, typewriter credit, a left-to-right
// sweep reveal of the block title, blinking coin prompt). Pure function of t,
// like every Observatory scene: replays identically, and pausing mid-reveal
// leaves an ordinary editable paragraph with half a title in it.

// Original 7×5 block font for the six letters the title needs.
const INTRO_FONT = {
  D: ['######.', '##...##', '##...##', '##...##', '######.'],
  O: ['.#####.', '##...##', '##...##', '##...##', '.#####.'],
  C: ['.######', '##.....', '##.....', '##.....', '.######'],
  X: ['##...##', '.##.##.', '..###..', '.##.##.', '##...##'],
  U: ['##...##', '##...##', '##...##', '##...##', '.#####.'],
  S: ['.######', '##.....', '.#####.', '.....##', '######.'],
};
const INTRO_TITLE = 'DOCXODUS';
const INTRO_LETTER_W = 7, INTRO_LETTER_GAP = 2;
const INTRO_TITLE_W =
  INTRO_TITLE.length * INTRO_LETTER_W + (INTRO_TITLE.length - 1) * INTRO_LETTER_GAP;

const centerX = (text) => Math.floor((COLS - text.length) / 2);

/** One attract frame at t seconds. Exported for the Playwright spec and the
 *  headless logic checks. */
export function introFrame(t) {
  const g = makeGrid();

  // Starfield: sparse, twinkling on a deterministic schedule.
  for (let y = 0; y < ROWS; y++) {
    for (let x = 0; x < COLS; x++) {
      const r = hash2(x, y * 3 + 7);
      if (r < 0.985) continue;
      const phase = (r * 900 + t * 0.9) % 3;
      g.chars[y][x] = phase < 1.6 ? '·' : '+';
      g.colors[y][x] = phase < 1.6 ? '33465B' : '9CB3C9';
    }
  }

  // Credit line, typed out one character at a time.
  const credit = 'OS LEGAL  PRESENTS';
  const typed = Math.max(0, Math.floor((t - 0.4) / 0.07));
  if (typed > 0) {
    writeText(g, 5, centerX(credit), credit.slice(0, typed), 'FFD166');
    if (typed <= credit.length) {
      // Typing cursor rides the leading edge, then vanishes.
      g.chars[5][centerX(credit) + Math.min(typed, credit.length - 1) + 1] = '_';
      g.colors[5][centerX(credit) + Math.min(typed, credit.length - 1) + 1] = 'FFD166';
    }
  }

  // The block title sweeps in left→right; the sweep front glows white for a
  // few columns before settling into teal.
  const x0 = Math.floor((COLS - INTRO_TITLE_W) / 2);
  const reveal = (t - 1.7) / 2.2; // 0..1 across the title's width
  if (reveal > 0) {
    const front = reveal * (INTRO_TITLE_W + 4);
    for (let li = 0; li < INTRO_TITLE.length; li++) {
      const glyph = INTRO_FONT[INTRO_TITLE[li]];
      const lx = x0 + li * (INTRO_LETTER_W + INTRO_LETTER_GAP);
      for (let row = 0; row < 5; row++) {
        for (let col = 0; col < INTRO_LETTER_W; col++) {
          if (glyph[row][col] !== '#') continue;
          const rel = lx + col - x0;
          if (rel > front) continue;
          const edge = front - rel;
          g.chars[9 + row][lx + col] = edge < 1.5 ? '░' : edge < 3 ? '▓' : '█';
          g.colors[9 + row][lx + col] = edge < 3 ? 'F3FBFF' : '5EEAD4';
        }
      }
    }
  }

  if (t > 4.1) {
    const sub = '·  T H E   D O C X   A R C A D E  ·';
    writeText(g, 16, centerX(sub), sub, '9CB3C9');
  }
  if (t > 4.5) {
    const foot = 'a video game running inside a live Word document';
    writeText(g, 18, centerX(foot), foot, '46556B');
    const foot2 = 'every frame is one paragraph · pause anytime and edit it';
    writeText(g, 19, centerX(foot2), foot2, '46556B');
  }
  if (t > 4.9 && (t % 1.1) < 0.75) {
    const prompt = '▶  PRESS  SPACE  TO  START  ◀';
    writeText(g, 22, centerX(prompt), prompt, 'FFD166');
  }

  return { grid: g, bg: '0A1020' };
}

// ─── The editor-hosted driver ─────────────────────────────────────────

/**
 * The oldest published engine whose SINGLE-BLOCK render carries a block's
 * inline image, and therefore the oldest one an image-bearing cartridge (Doom)
 * can be seen on. Up to and including 10.0.0 the incremental renderer cloned
 * the block's XML into a throwaway shell without copying the referenced media
 * part, and its converter settings carried no image handler, so
 * WmlToHtmlConverter omitted the w:drawing and the frame paragraph refreshed
 * blank. `docs/demo/tools/engine-pin.test.mjs` holds the demo pages' jsDelivr
 * pin at or above this, so the arcade cannot ship pointed at a blind engine.
 */
export const IMAGE_ENGINE_MINIMUM = '11.0.0';

/** The attract screen's one live control: Space is the coin drop, so the pad's
 *  round button says so rather than offering to fire at a title card. */
const INTRO_FIRE = { code: 'Space', glyph: 'START', label: 'Start DOOM' };

/**
 * Seed the Arcade into a ribbon-hosted editor's session and run the game loop
 * against it. Owns the dock (pause/resume, restart, pace,
 * telemetry) and the keyboard: game keys are claimed only while playing.
 * Clicking the document pauses — the frame you clicked is now just a
 * paragraph with your caret in it. Resuming blurs the edit (the editor
 * commits on blur), re-parses the game world from the session's XML, and
 * hands the keyboard back to the game.
 *
 * `ui`: { playpause, restart, pace, stats, hint, pad?, setPad? } — dock DOM,
 * plus the hook that re-points the touch pad at the selected cartridge's keys.
 * `intro` (default true) opens on the attract screen — the same canvas
 * paragraph running the title card until Space (or any dock action) drops
 * the coin. Returns the controller the host page publishes as
 * `window.__arcade`.
 */
export function startArcade({ editor, session, ui, intro = true, doom = {} }) {
  if (typeof editor.refresh !== 'function') {
    throw new Error('This engine predates DocxEditor.refresh() — the Arcade needs docxodus ≥ 9.6.0.');
  }
  const seeded = seedArcade(session);
  let canvasAnchor = seeded.canvasAnchor;
  let openTag = seeded.openTag;
  const pinCanvas = createCanvasPin();
  pinCanvas(canvasAnchor);

  const cart = doomCart(doom);

  let mode = intro ? 'intro' : 'game';
  let introT = 0;
  let playing = false;
  let timer = 0;
  let lastWall = performance.now();
  let frames = 0;
  let fps = 0;
  let lastRuns = 0;
  let canvasImageId = null;
  // A cartridge whose frame is a native inline image (Doom) depends on the
  // engine's SINGLE-BLOCK render carrying that image. Engines up to and
  // including 10.0.0 do not: the incremental renderer clones the block's XML
  // into a throwaway shell without copying the referenced media part, and its
  // converter settings carry no image handler, so WmlToHtmlConverter omits the
  // w:drawing and the paragraph refreshes to blank. The image is genuinely in
  // the package the whole time — Save and reopen shows it — which is exactly
  // what makes the failure invisible from the outside. The site's shared build
  // carries the fix; the probe survives for an ?engine= override aimed at an
  // older release. So prove the surface once, rather than painting
  // frames no one can see.
  let imageSurfaceProven = false;
  let imageFramesPainted = 0;
  let lastSurface = 'runs';
  let paintGeneration = 0;
  let lastImageOptions = null;
  let lastFrameEnd = performance.now();
  const timings = { mutate: 0, refresh: 0 };
  let interval = Number(ui.pace.value);

  const input = createInput(() => playing);

  // What the on-screen pad is currently holding down: pressed buttons keyed by
  // the pointer holding them, plus the latched modifiers (RUN) that stay down
  // after the thumb leaves. Declared here, beside the input they feed, because
  // pausing releases them and pausing can happen before the pad is wired.
  const padHeld = new Map();
  const padLatched = new Map();
  function releasePad() {
    for (const { code } of padHeld.values()) input.set(code, false);
    padHeld.clear();
    for (const [button, code] of padLatched) {
      input.set(code, false);
      button.setAttribute('aria-pressed', 'false');
    }
    padLatched.clear();
  }

  const unidOf = (anchor) => anchor.split(':')[2];
  const canvasEl = () => editor.root.querySelector(`[data-anchor="${unidOf(canvasAnchor)}"]`);
  const controlsEls = () => seeded.controlsAnchors.map((anchor) =>
    editor.root.querySelector(`[data-anchor="${unidOf(anchor)}"]`));
  const controlsEl = () => controlsEls()[0];

  // replaceText authors a fresh run, so re-assert the controls' presentation
  // after each cartridge/intro label swap. Formatting only the initial
  // "loading cartridge" run would leave the visible replacement at the
  // document default despite the paragraph having been seeded at 18pt.
  function setControls(lines) {
    for (let i = 0; i < seeded.controlsAnchors.length; i++) {
      const anchor = seeded.controlsAnchors[i];
      session.replaceText(anchor, lines[i] ?? '\u00a0');
      session.applyFormat(anchor, null,
        { bold: true, fontFamily: 'Courier New', fontSizePts: 18, color: '25324A' });
    }
  }

  function setCaption() {
    if (ui.rendering) {
      ui.rendering.hidden = cart.name !== 'doom';
      ui.rendering.value = cart.state().rendering ?? 'image';
    }
    // The pad follows the cartridge — and stands down on the attract screen,
    // where nothing steers anything and the only live control is the coin.
    ui.setPad?.(mode === 'intro' ? { fire: INTRO_FIRE } : cart.touch);
    if (mode === 'intro') {
      setControls([
        'CONTROLS · START SPACE',
        'DOOM · ASCII OR BITMAP',
        '\u00a0',
        'PAUSE/EDIT ESC',
      ]);
      session.replaceText(seeded.captionAnchor,
        'OS Legal presents **DOCXODUS** — press **Space** to start. ' +
        'This title card is a Word paragraph too: pause and put your caret in it.');
      ui.hint.innerHTML =
        '<b>Space</b> starts DOOM · choose ASCII or bitmap below · ' +
        '<b>Esc</b> pauses — even the title screen is just a document';
      return;
    }
    setControls(cart.controls);
    session.replaceText(seeded.captionAnchor, cart.caption);
    ui.hint.innerHTML = cart.hint +
      ' · <b>Esc</b> pauses/resumes · <b>Undo/Redo</b> scrubs frames · <b>Save</b> ships the frame as .docx';
  }

  function paintGrid(frame, label) {
    const { xml, runs } = frameXml(openTag, frame.grid, frame.bg, frame.metrics);
    lastRuns = runs;
    lastSurface = 'runs';
    const t0 = performance.now();
    const res = session.raw.replaceXml(canvasAnchor, xml);
    const t1 = performance.now();
    if (!res.success) throw new Error(`replaceXml: ${res.error?.code} ${res.error?.message}`);
    // replaceXml sweeps the replaced drawing's orphaned relationship itself.
    // One mutation keeps switching from an image to text a single undo step.
    canvasImageId = null;
    canvasAnchor = res.modified[0]?.id ?? res.created[0]?.id ?? canvasAnchor;
    pinCanvas(canvasAnchor);
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
      `<b>${label}</b> · frame <b>${frames}</b> · <b>${fps.toFixed(1)}</b> fps · ` +
      `replaceXml <b>${timings.mutate.toFixed(1)}</b> ms · refresh <b>${timings.refresh.toFixed(1)}</b> ms · ` +
      `<b>${lastRuns}</b> ${lastSurface} · ` +
      (fb ? `remounted (${fb})` : `<span class="inc">incremental — one block repainted</span>`);
    return null;
  }

  /** Paint a cartridge frame as a real inline DOCX image. Both insertion and
   *  replacement go through the public session surface; editor.refresh then
   *  asks the standard single-block converter for an <img> backed by that
   *  package media part. No out-of-document canvas is mounted. */
  function paintImage(frame, label) {
    lastRuns = 1;
    lastSurface = 'image';
    lastImageOptions = { ...frame.imageOptions };
    const t0 = performance.now();
    let res;
    if (canvasImageId) {
      res = session.replaceImage(canvasImageId, frame.imageBytes);
    } else {
      const steps = [
        { tool: 'raw', action: 'replaceXml', mutation: () =>
          session.raw.replaceXml(canvasAnchor, openTag + '</w:p>') },
        { tool: 'image', action: 'insert', mutation: () =>
          (res = session.insertImage(canvasAnchor, 0, frame.imageBytes, frame.imageOptions)) },
      ];
      if (typeof session.executeBatch === 'function') {
        const batch = session.executeBatch(steps);
        if (!batch.success) throw new Error('Could not switch the document frame to an image');
      } else {
        for (const step of steps) {
          const result = step.mutation();
          if (!result.success) throw new Error(`image frame: ${result.error?.message}`);
        }
      }
      canvasImageId = res.imageId ?? null;
    }
    const t1 = performance.now();
    if (!res.success) throw new Error(`image frame: ${res.error?.code} ${res.error?.message}`);
    canvasAnchor = res.modified?.[0]?.id ?? res.created?.[0]?.id ?? canvasAnchor;
    pinCanvas(canvasAnchor);
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
      `<b>${label}</b> · frame <b>${frames}</b> · <b>${fps.toFixed(1)}</b> fps · ` +
      `replaceImage <b>${timings.mutate.toFixed(1)}</b> ms · refresh <b>${timings.refresh.toFixed(1)}</b> ms · ` +
      `<b>1</b> inline image · ` +
      (fb ? `remounted (${fb})` : `<span class="inc">incremental — one block repainted</span>`);
    const img = canvasEl()?.querySelector('img') ?? null;
    // Two frames of slack before calling it: the check is a capability probe,
    // not a per-frame assertion, and one empty reconcile should not halt a game.
    imageFramesPainted++;
    if (img) imageSurfaceProven = true;
    else if (!imageSurfaceProven && imageFramesPainted >= 3) {
      const why =
        'this engine renders a single block without its inline image, so the frame '
        + 'is in the .docx but never on screen — the image-bearing cartridges need '
        + `docxodus ${IMAGE_ENGINE_MINIMUM} or newer, so check the ?engine= override `
        + `(this page pins ${IMAGE_ENGINE_MINIMUM}, which carries the fix). `
        + 'Save still downloads the frame, and '
        + 'Use the current build to play DOOM in ASCII or bitmap mode.';
      // Not every drawFrame call is inside the loop's try — the cartridge
      // switch and restart buttons repaint directly — so say it here rather
      // than relying on a catch that only one of the three call sites has.
      ui.stats.textContent = 'halted: ' + why;
      playing = false;
      throw new Error(why);
    }
    return img;
  }

  function drawFrame() {
    if (mode === 'intro') return paintGrid(introFrame(introT), 'attract mode');
    else {
      const frame = cart.render();
      if (frame.imageBytes) return paintImage(frame, cart.label);
      return paintGrid(frame, cart.label);
    }
  }

  /** Drop the coin: leave the attract screen and hand the canvas (and the
   *  keyboard) to the selected cartridge. */
  function startGame() {
    if (mode !== 'intro') return;
    mode = 'game';
    setCaption();
    if (playing) {
      lastWall = performance.now();
    } else {
      drawFrame();
      setPlaying(true);
    }
  }

  /** The data URI the editor will emit for a frame's media: the converter
   *  base64-encodes the package part verbatim, so this is the same string. */
  function dataUrlOf(bytes, mime = 'image/png') {
    let binary = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
      binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return `data:${mime};base64,${btoa(binary)}`;
  }

  /** Load and decode a frame's image, as the exact data URI the editor will
   *  render for it, BEFORE the document is touched.
   *
   *  The editor patches the new source onto the <img> already on screen rather
   *  than swapping the element, and Chromium and Firefox then keep the previous
   *  frame up until the new one is ready. WebKit does not: it paints the box
   *  empty for any frame in which a changed source is still decoding, which at
   *  game frame rate is a white strobe. Warming the same URI first puts it in
   *  the image cache, so the editor's element completes synchronously and the
   *  screen never shows a gap. A decode failure is not fatal — the element
   *  loads the source itself, exactly as before. */
  function prewarmImage(frame) {
    const url = frame.imageDataUrl ?? dataUrlOf(frame.imageBytes);
    const probe = new Image();
    probe.src = url;
    return typeof probe.decode === 'function' ? probe.decode().catch(() => {}) : Promise.resolve();
  }

  function loop() {
    if (!playing) return;
    const started = performance.now();
    const dt = Math.min(0.2, (started - lastWall) / 1000);
    lastWall = started;
    const halt = (e) => {
      playing = false;
      ui.stats.textContent = 'halted: ' + e.message;
      throw e;
    };
    let pendingPaint = null;
    try {
      if (mode === 'intro') {
        introT += dt;
        const start = input.took('Space');
        input.endTick();
        if (start) {
          startGame(); // startGame paints the cartridge's own first frame
        } else {
          drawFrame();
        }
      } else {
        cart.tick(dt, input);
        input.endTick();
        const frame = cart.render();
        // A native-image frame is decoded before it enters the document (see
        // prewarmImage); text frames paint synchronously as they always have.
        if (frame.imageBytes) {
          const generation = paintGeneration;
          pendingPaint = prewarmImage(frame).then(() =>
            playing && generation === paintGeneration ? paintImage(frame, cart.label) : null);
        }
        else paintGrid(frame, cart.label);
      }
    } catch (e) {
      halt(e);
    }
    const finish = (image) => {
      const delay = Math.max(0, interval - (performance.now() - started));
      const schedule = () => {
        if (!playing) return;
        timer = setTimeout(loop, delay);
      };
      // The refresh is complete but image decode/presentation is asynchronous,
      // and `complete` is false while a patched source is still pending. Never
      // write the next frame before the browser has decoded this one and
      // offered it to a paint frame; this makes the displayed FPS equal the
      // document FPS rather than flooding the element with intermediate
      // sources no player can see. (After prewarmImage this is normally an
      // immediate cache hit, and the wait costs one animation frame.)
      if (image instanceof HTMLImageElement) {
        const afterDecode = () => requestAnimationFrame(schedule);
        if (image.complete && image.naturalWidth > 0) afterDecode();
        else image.decode().catch(() => {}).then(afterDecode);
      } else {
        schedule();
      }
    };
    if (pendingPaint) pendingPaint.then(finish, halt);
    else finish(null);
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
    syncCanvasImageId();
    // The attract screen has no game world to parse back — whatever was typed
    // into the title card simply stays until the next frame repaints it.
    if (mode !== 'intro') cart.syncFromRows(rowsFromXml(xml));
    // Sweep paragraphs an Enter-split stranded below the screen, so pausing to
    // edit can never slowly litter the document. The boundary is the FENCE, not
    // the caption: the fence is a deliberate near-empty paragraph that keeps the
    // caption out of the screen's render context (see seedDocument), and sweeping
    // up to the caption would delete it on the first pause — which is exactly
    // what happened the first time this was tried.
    const ids = session.findByKind('p', 'body').map((r) => r.id);
    const fenceIdx = ids.indexOf(seeded.fenceBelow);
    const from = ids.indexOf(canvasAnchor);
    const to = fenceIdx >= 0 ? fenceIdx : ids.indexOf(seeded.captionAnchor);
    if (from >= 0 && to > from + 1) {
      for (const id of ids.slice(from + 1, to)) session.deleteBlock(id);
    }
  }

  function syncCanvasImageId() {
    // Undo, paste or deletion can replace a drawing while the game is paused.
    // Resolve its current identity at these transitions, outside the hot loop.
    canvasImageId = canvasEl()?.querySelector('img')
      ? session.listImages().find(image => image.anchorId === canvasAnchor)?.id ?? null
      : null;
  }

  function setPlaying(next) {
    if (playing === next) return;
    if (next) {
      syncFromDocument();
      // Resume is always reached from a gesture (a click, a key, the dock),
      // which is the only moment a browser will start an AudioContext. Only
      // the Doom cartridge has one; the others do not offer the hook.
      try { cart.state().resumeAudio?.(); } catch { /* sound is never fatal */ }
      playing = true;
      ui.playpause.textContent = '⏸ Pause & edit';
      lastWall = performance.now();
      lastFrameEnd = performance.now();
      loop();
    } else {
      playing = false;
      // A latched RUN (or a key still down when the frame froze) must not
      // survive into the paused document, nor into the next resume.
      releasePad();
      clearTimeout(timer);
      ui.playpause.textContent = '▶ Resume';
      ui.stats.innerHTML = lastSurface === 'image'
        ? 'paused — this frame is a real DOCX image. <b>Ctrl+C / Ctrl+V</b> duplicates it; ' +
          '<b>Undo / Redo</b> scrubs backward and forward through frames.'
        : cart.name === 'doom' && cart.state().rendering === 'ascii'
          ? 'paused — this frame is colored ASCII text. <b>Ctrl+C / Ctrl+V</b> duplicates it; ' +
            '<b>Undo / Redo</b> scrubs backward and forward through frames.'
        : 'paused — the screen is an ordinary paragraph now. ' +
          '<b>Type into it</b>, then Resume to make it real.';
    }
  }

  // Kept for embedding callers that restart DOOM programmatically.
  function setCart(name) {
    if (name !== 'doom') return;
    releasePad(); paintGeneration++; cart.reset();
    if (mode === 'intro') { startGame(); return; }
    setCaption();
    if (!playing) { drawFrame(); setPlaying(true); }
  }

  ui.playpause.addEventListener('click', () => setPlaying(!playing));
  ui.restart.addEventListener('click', () => {
    cart.reset();
    if (mode === 'intro') { startGame(); return; }
    if (!playing) { drawFrame(); setPlaying(true); }
  });
  ui.pace.addEventListener('change', () => { interval = Number(ui.pace.value); });
  function setRendering(value) {
    if (!cart.setRendering) return;
    const visibleImage = !!canvasEl()?.querySelector('img');
    if (value === cart.state().rendering && (mode !== 'game' ||
        (value === 'image' ? visibleImage : !visibleImage))) return;
    paintGeneration++;
    syncCanvasImageId();
    cart.setRendering(value);
    fps = 0;
    timings.mutate = 0; timings.refresh = 0;
    lastFrameEnd = performance.now();
    if (ui.rendering) ui.rendering.value = value;
    // Project the same engine frame immediately, even while paused. No tick,
    // reset, input transition or game-state change belongs to this control.
    if (mode === 'game') drawFrame();
  }
  ui.rendering?.addEventListener('change', () => setRendering(ui.rendering.value));

  // Esc toggles play/pause from anywhere — including from inside the document.
  window.addEventListener('keydown', (e) => {
    if (e.code !== 'Escape') return;
    e.preventDefault();
    setPlaying(!playing);
  }, true);

  // Click the game (or any block) while it plays: the frame freezes and the
  // caret lands in it. No mode switch — it was a document the whole time.
  editor.root.addEventListener('pointerdown', () => setPlaying(false), true);

  // Copy and paste the screen paragraph.
  //
  // Selecting the paused frame and pressing Ctrl+C already works — it is
  // ordinary document content, and the clipboard gets the picture. Pasting it
  // back is where the cabinet has to help: the editor's native paste surface is
  // text-only today. The copy therefore stashes either the paragraph OOXML or
  // the native image bytes, and paste inserts a real block through the public
  // image API. That also mints fresh drawing ids instead of duplicating the
  // source paragraph's non-visual ids.
  //
  // Both listeners are scoped to the screen paragraph. Any other copy or paste
  // in the document is the browser's, untouched.
  let copiedFrame = null;

  function dataImageBytes(element) {
    const source = element?.querySelector('img')?.getAttribute('src') ?? '';
    const comma = source.indexOf(',');
    if (!source.startsWith('data:image/') || comma < 0 || !source.slice(0, comma).includes(';base64')) {
      return null;
    }
    const binary = atob(source.slice(comma + 1));
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  editor.root.addEventListener('copy', (event) => {
    const sel = window.getSelection();
    const el = canvasEl();
    const overlaps = el && sel && !sel.isCollapsed && Array.from({ length: sel.rangeCount }, (_, i) =>
      sel.getRangeAt(i).intersectsNode(el)).some(Boolean);
    const visibleImage = overlaps && lastImageOptions ? dataImageBytes(el) : null;
    copiedFrame = overlaps ? {
      xml: session.raw.getXml(canvasAnchor),
      image: visibleImage && {
        // Read the rendered document image rather than the producer's last
        // frame cache. If Undo has scrubbed backward, copy must capture the
        // frame now visible in the editor, not the future frame it replaced.
        bytes: visibleImage,
        options: { ...lastImageOptions },
      },
    } : null;
    if (copiedFrame && !copiedFrame.image && cart.name === 'doom' && event.clipboardData) {
      // HTML uses nonbreaking spaces for layout; a text editor should receive
      // the literal printable ASCII and newlines authored in the document.
      event.preventDefault();
      event.clipboardData.setData('text/plain', rowsFromXml(copiedFrame.xml).join('\n'));
      event.clipboardData.setData('text/html', el.outerHTML);
    }
  });
  editor.root.addEventListener('paste', (event) => {
    if (!copiedFrame) return;
    const unid = event.target?.closest?.('[data-anchor]')?.getAttribute('data-anchor');
    if (!unid) return;
    event.preventDefault();

    // Where it lands. After the block holding the caret, unless that is the
    // screen itself or one of the fence paragraphs around it — everything in
    // that span is swept on resume (see the litter sweep), so a copy dropped
    // there would vanish the moment the game restarts. Those go to the end of
    // the body instead, where they keep.
    const ids = session.findByKind('p', 'body').map((t) => t.id);
    const fence = ids.indexOf(seeded.fenceBelow);
    const screen = ids.indexOf(canvasAnchor);
    let at = ids.findIndex((id) => id.endsWith(unid));
    if (at < 0 || (screen >= 0 && at >= screen - 1 && at <= fence)) at = ids.length - 1;
    if (at < 0) return;

    // A block-level insert, then the copied body under the NEW paragraph's
    // opening tag: the Unid in that tag is the block's identity, so reusing the
    // source's would give two blocks one anchor.
    const made = session.insertParagraph(ids[at], 'after', '\u00a0');
    const created = made.created?.[0]?.id;
    if (!created) return;
    const seed = session.raw.getXml(created);
    let open = seed.slice(0, seed.indexOf('>') + 1);
    if (open.endsWith('/>')) open = `${open.slice(0, -2)}>`;
    if (copiedFrame.image) {
      const inserted = session.insertImage(created, 0,
        copiedFrame.image.bytes, copiedFrame.image.options);
      if (!inserted.success) throw new Error(
        `paste image: ${inserted.error?.code} ${inserted.error?.message}`);
    } else {
      session.raw.replaceXml(created, open + copiedFrame.xml.slice(copiedFrame.xml.indexOf('>') + 1));
    }
    editor.refresh();
  });

  // On-screen pad (touch): buttons carry data-code="ArrowLeft" etc. Wired by
  // DELEGATION, and reading the code at event time, because the dock re-points
  // those buttons whenever the cartridge changes and mints the tray's keys on
  // the spot — a listener that captured a code at boot would go on sending the
  // previous cartridge's key.
  //
  // Presses are tracked BY POINTER ID: steering with one thumb while firing
  // with the other is two live pointers, and releasing one must not lift the
  // other's key. The release listener sits on the window so a thumb that
  // slides off the button it pressed still lifts that key rather than leaving
  // it stuck down for the rest of the game.
  if (ui.pad) {
    ui.pad.addEventListener('pointerdown', (e) => {
      const button = e.target.closest?.('[data-code]');
      if (!button) return;
      e.preventDefault();
      if (!playing) setPlaying(true);
      const code = button.getAttribute('data-code');
      // data-mode="toggle" LATCHES (RUN): holding a modifier is a keyboard
      // affordance, and on a phone that hand is busy steering.
      if (button.dataset.mode === 'toggle') {
        const latched = padLatched.has(button);
        if (latched) padLatched.delete(button);
        else padLatched.set(button, code);
        button.setAttribute('aria-pressed', String(!latched));
        input.set(code, !latched);
        return;
      }
      try { button.setPointerCapture(e.pointerId); } catch { /* mouse rigs */ }
      padHeld.set(e.pointerId, { button, code });
      input.set(code, true);
    });
    const release = (e) => {
      const entry = padHeld.get(e.pointerId);
      if (!entry) return;
      padHeld.delete(e.pointerId);
      input.set(entry.code, false);
    };
    window.addEventListener('pointerup', release);
    window.addEventListener('pointercancel', release);
  }

  setCaption();
  setPlaying(true);

  return {
    canvasAnchor: () => canvasAnchor,
    canvasText: () => canvasEl()?.textContent ?? '',
    canvasElement: () => canvasEl(),
    controlsText: () => controlsEls().map((line) => line?.textContent ?? '').join(' '),
    controlsElement: () => controlsEl(),
    controlsElements: () => controlsEls(),
    frames: () => frames,
    fps: () => fps,
    timings: () => ({ ...timings, runs: lastRuns }),
    cart: () => cart.name,
    setCart,
    setRendering,
    game: () => cart.state(),
    playing: () => playing,
    introActive: () => mode === 'intro',
    start: startGame,
    pause: () => setPlaying(false),
    resume: () => setPlaying(true),
    input,
    save: () => editor.save(),
  };
}

// ═══════════════════════════════════════════════════════════════════════
// Cartridge 3 — DOOM. The actual game, not an impression of it.
//
// LICENSE NOTE — THIS FILE IS GPL-2.0-or-later, NOT MIT.
// Docxodus is MIT (see the root LICENSE) and every other file in this
// directory stays MIT. This one is different: it is written against, and at
// runtime combined with, `vendor/doomgeneric/doomgeneric_module.js` — a build
// of id Software's Doom source, which id released under the GNU General
// Public License v2. So this glue is offered under GPL-2.0-or-later too, and
// the arcade reaches it only through a dynamic `import()` (see
// ascii-arcade.js), which is also why the 3 MB engine never loads for a
// visitor who plays the other two cartridges. See vendor/NOTICE.md.
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

/** The arcade key codes this cartridge wants to be handed. ascii-arcade.js
 *  unions this into the set its capture-phase listener claims while playing,
 *  so Doom gets Enter/E/Q/M/digits without the other cartridges caring. */
export const DOOM_KEY_CODES = Object.keys(KEY_MAP);

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
    for (let x = 0; x < VIEW_W; x++) {
      const sx = SX[x];
      // BGRA in memory: blue first, alpha always 0 — hence the explicit swap
      // rather than a straight copy.
      const t = (topBase + sx) * 4, b = (botBase + sx) * 4;
      const tr = fb[t + 2], tg = fb[t + 1], tb = fb[t];
      const br = fb[b + 2], bg2 = fb[b + 1], bb = fb[b];
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

// ─── The engine, loaded once per page ─────────────────────────────────
// doomgeneric talks to its host through bare global functions (its C calls
// them with EM_ASM, which resolves free names against globalThis), so there
// can only ever be one Doom per page. That is fine — the arcade shows one
// cartridge at a time — but it does mean the module is a page-level
// singleton rather than per-cartridge state.
let enginePromise = null;

const DEFAULT_ENGINE = new URL('./vendor/doomgeneric/doomgeneric_module.js', import.meta.url).href;
const DEFAULT_WAD = new URL('./vendor/freedoom/freedoom1.wad.gz', import.meta.url).href;

/** Resolve a URL and refuse it unless it is same-origin.
 *
 *  `import()` EXECUTES what it fetches, with this page's privileges and on
 *  this page's origin, so an engine URL that a link can choose is remote code
 *  execution rather than a convenience — which is why `?doomEngine=` no longer
 *  exists and why this guard stands behind the option that remains. The IWAD
 *  is only data and is magic-checked before use, but it goes through the same
 *  gate: `?wad=` is for pointing the cartridge at an IWAD you host yourself,
 *  and nothing is lost by requiring you to host it.
 *
 *  In Node (the headless logic tests) there is no location and nothing is ever
 *  loaded, so the guard is inert there. */
function sameOrigin(url, what) {
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
 *  The WAD ships gzipped (28.8 MB → 10.3 MB) because GitHub Pages will not
 *  content-encode an `application/octet-stream`, so compressing it in the
 *  repository is the only way the visitor's download is the small number.
 *  DecompressionStream does the inflate natively — no library. */
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
    engineUrl = sameOrigin(engineUrl, 'the Doom engine');
    wadUrl = sameOrigin(wadUrl, 'the IWAD');
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
    const path = globalThis.location ? new URL(wadUrl, globalThis.location.href).pathname : wadUrl;
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
  ['Esc', 'pause & edit'],
];

export function doomCart(options = {}) {
  const engineUrl = options.engineUrl ?? DEFAULT_ENGINE;
  const wadUrl = options.wadUrl ?? DEFAULT_WAD;
  const sound = options.sound !== false;

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
    y = ROWS - 4;
    write(g, y, x, `frame ${handle?.frames ?? 0}`, DEAD_INK);
    write(g, y + 1, x, 'id Software engine,', DEAD_INK);
    write(g, y + 2, x, 'GPL-2.0 · Freedoom data', DEAD_INK);
  }

  function render() {
    const g = makeGrid();
    if (status === 'playing' && handle) {
      paintFramebuffer(g, handle.framebuffer);
      paintedFrames++;
      drawDivider(g);
      drawPanel(g);
      drawChrome(g, 'DOOM — id Software’s own engine, drawing into one Word paragraph');
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
      'run shading is the bottom one. Move **W/S** · strafe **A/D** · turn **←/→** · **Space** ' +
      'fires · **E** opens · **Q** is Doom’s own menu. **Esc** pauses — and then it is only a ' +
      'document again: put your caret in the frame, Undo rewinds it, Save downloads it as .docx.',
    hint: '<b>WASD</b> move · <b>←/→</b> turn · <b>Space</b> fire · <b>E</b> open · <b>Q</b> Doom’s menu — the real engine, one Word paragraph as the screen.',
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
    state: () => ({
      status,
      error,
      progress,
      edited,
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

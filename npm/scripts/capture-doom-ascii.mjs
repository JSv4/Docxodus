#!/usr/bin/env node
// Capture original DOOM shareware running as native text in the real editor.
// Requires the built demo at :8082, fetch-doom-iwad.mjs --shareware, and ffmpeg.
// FFMPEG may name a local executable. DOOM_ENGINE_PATH optionally mirrors the
// exact pinned engine for offline recording; its digest is checked below.
import { chromium } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdirSync, mkdtempSync } from 'node:fs';
import { dirname, resolve, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { spawnSync } from 'node:child_process';

const here = dirname(fileURLToPath(import.meta.url));
const output = resolve(here, '../../docs/images');
const scratch = mkdtempSync(join(tmpdir(), 'docxodus-ascii-capture-'));
const base = process.env.ARCADE_URL ?? 'http://localhost:8082';
const engineSha = 'efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9';
const wadSha = '1d7d43be501e67d927e415e0b8f3e29c3bf33075e859721816f652a526cac771';
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const wad = gunzipSync(readFileSync(resolve(here, '../dist/wasm/vendor/doom1.wad.gz')));
if (sha(wad) !== wadSha) throw new Error('Expected the original DOOM 1.9 shareware IWAD');
mkdirSync(output, { recursive: true });
const browser = await chromium.launch({ headless: true });
const viewport = { width: 1100, height: 1050 };
const context = await browser.newContext({ viewport, deviceScaleFactor: 2 });
const page = await context.newPage();
const errors = [];
page.on('pageerror', e => errors.push(e.message));
if (process.env.DOOM_ENGINE_PATH) {
  const bytes = readFileSync(process.env.DOOM_ENGINE_PATH);
  if (sha(bytes) !== engineSha) throw new Error('Engine mirror does not match the cartridge pin');
  await page.route('https://cdn.jsdelivr.net/gh/grubbyplaya/**', route =>
    route.fulfill({ body: bytes, contentType: 'text/javascript' }));
}

try {
  await page.goto(`${base}/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0`
    + '&cart=doom&render=ascii&wad=./vendor/doom1.wad.gz');
  await page.waitForFunction(() => window.__arcade?.game().doomFrames >= 2,
    null, { timeout: 180000 });
  await page.selectOption('#pace', '0');
  await page.evaluate(() => document.fonts.ready);
  const frameBox = () => page.evaluate(() => {
    const { x, y, width, height } = window.__arcade.canvasElement().getBoundingClientRect();
    return { x, y, width, height };
  });
  // Lossless browser frames keep the small ASCII and static ribbon crisp.
  // Preserve their capture timestamps when encoding; never accelerate play.
  const cdp = await context.newCDPSession(page);
  const recorded = [];
  cdp.on('Page.screencastFrame', event => {
    const bytes = Buffer.from(event.data, 'base64');
    // Screenshot tools can temporarily resize the viewport. Keep the recording
    // at its actual presentation size, holding the preceding frame if needed.
    if (bytes.readUInt32BE(16) !== viewport.width || bytes.readUInt32BE(20) !== viewport.height) {
      cdp.send('Page.screencastFrameAck', { sessionId: event.sessionId }).catch(() => {});
      return;
    }
    const file = `frame-${String(recorded.length).padStart(6, '0')}.png`;
    writeFileSync(join(scratch, file), bytes);
    recorded.push({ file, time: event.metadata.timestamp, received: Date.now() / 1000 });
    cdp.send('Page.screencastFrameAck', { sessionId: event.sessionId }).catch(() => {});
  });
  await cdp.send('Page.startScreencast', { format: 'png', maxWidth: viewport.width,
    maxHeight: viewport.height, everyNthFrame: 1 });
  await page.screenshot({ path: join(output, 'arcade-doom-ascii-title.png') });
  console.log('Captured original DOOM title screen');
  if (process.argv.includes('--preview')) {
    await context.close();
    await browser.close();
    process.exit(0);
  }
  // Exercise the direct title-to-menu transition that players see.
  await page.waitForTimeout(800);
  await page.keyboard.press('Enter'); // main menu
  await page.waitForTimeout(1000);
  await page.screenshot({ path: join(output, 'arcade-doom-ascii-menu.png') });
  await page.keyboard.press('Enter'); // New Game
  await page.waitForTimeout(900);
  await page.keyboard.press('Enter'); // Knee-Deep in the Dead
  await page.waitForTimeout(900);
  await page.keyboard.press('ArrowUp'); // Hey, Not Too Rough
  await page.waitForTimeout(650);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(1500);
  await page.screenshot({ path: join(output, 'arcade-doom-ascii-gameplay.png') });
  console.log('Captured E1M1; recording movement and firing');
  const start = await page.evaluate(() => ({ time: performance.now(), frames: window.__arcade.frames() }));
  const hold = async (key, ms) => {
    await page.keyboard.down(key); await page.waitForTimeout(ms); await page.keyboard.up(key);
  };
  await hold('KeyW', 1100);
  await hold('ArrowLeft', 450);
  await hold('Space', 1200);
  await hold('ArrowRight', 750);
  await hold('KeyW', 1400);
  await hold('Space', 1100);
  await hold('ArrowRight', 650);
  await hold('KeyS', 900);
  await hold('ArrowLeft', 700);
  await hold('Space', 1000);
  const end = await page.evaluate(() => ({ time: performance.now(), frames: window.__arcade.frames() }));
  await page.click('#playpause');
  await page.waitForTimeout(1500);
  const clipEnd = Date.now() / 1000;
  await cdp.send('Page.stopScreencast');
  await page.screenshot({ path: join(output, 'arcade-doom-ascii-frame.png'),
    clip: await frameBox() });
  const saved = await page.evaluate(() => Array.from(window.__arcade.save()));
  writeFileSync(join(output, 'arcade-doom-ascii-frame.docx'), Buffer.from(saved));
  const box = await frameBox();
  await cdp.send('Emulation.setDeviceMetricsOverride', { ...viewport, deviceScaleFactor: 4, mobile: false });
  await page.screenshot({ path: join(output, 'arcade-doom-ascii-detail.png'),
    clip: { x: box.x, y: box.y + box.height * .73, width: box.width, height: box.height * .27 } });
  const proof = await page.evaluate(() => {
    const a = window.__arcade, el = a.canvasElement();
    return { engineTitle: a.game().title, rendering: a.game().rendering,
      textCharacters: el.textContent.length, images: el.querySelectorAll('img,canvas,svg').length,
      incremental: a.editor.lastReconcileFallback === null };
  });
  await context.close();
  if (errors.length) throw new Error(errors.join('\n'));
  if (proof.images || proof.textCharacters !== 64200 || !proof.incremental)
    throw new Error('Capture did not remain a complete native ASCII document frame');
  const ffmpeg = process.env.FFMPEG ?? 'ffmpeg';
  const encode = args => {
    const result = spawnSync(ffmpeg, ['-hide_banner', '-loglevel', 'error', '-y', ...args], { encoding: 'utf8' });
    if (result.error || result.status !== 0) throw new Error(result.error?.message ?? result.stderr);
  };
  if (recorded.length < 20) throw new Error('Too few browser frames were recorded');
  let duration = 0;
  const manifest = recorded.map((frame, i) => {
    const seconds = Math.max(.001, i + 1 < recorded.length
      ? recorded[i + 1].time - frame.time : clipEnd - frame.received);
    duration += seconds;
    return `file '${frame.file}'\nduration ${seconds}\n`;
  }).join('') + `file '${recorded.at(-1).file}'\n`;
  const manifestPath = join(scratch, 'frames.txt');
  writeFileSync(manifestPath, manifest);
  const input = ['-f', 'concat', '-safe', '0', '-i', manifestPath, '-t', String(duration), '-an'];
  encode([...input, '-vf', 'fps=30', '-c:v', 'libx264', '-crf', '18', '-pix_fmt', 'yuv420p',
    '-movflags', '+faststart', join(output, 'arcade-doom-ascii.mp4')]);
  encode([...input, '-filter_complex',
    'fps=8,scale=880:-2:flags=lanczos,format=rgb24,split[a][b];[a]palettegen=max_colors=256:reserve_transparent=0:stats_mode=diff[p];[b][p]paletteuse=dither=bayer:bayer_scale=4',
    // Opaque update rectangles replace changed pixels completely. Ordered
    // dithering keeps soft UI shadows smooth without noise changing each frame.
    '-gifflags', 'offsetting', '-loop', '0', join(output, 'arcade-doom-ascii.gif')]);
  const metadata = { ...proof, source: 'Original DOOM v1.9 shareware', engineSha256: engineSha,
    wadSha256: wadSha, viewport, gif: { width: 880, height: 840, framesPerSecond: 8,
      paletteColors: 256, transparentFrames: false },
    durationSeconds: duration,
    deliveredGameplayFps: (end.frames - start.frames) * 1000 / (end.time - start.time),
    capture: 'Lossless Chromium PNG screencast of the live editor, driven by Playwright; original timestamps preserved.' };
  writeFileSync(join(output, 'arcade-doom-ascii-capture.json'), JSON.stringify(metadata, null, 2) + '\n');
  console.log(JSON.stringify(metadata, null, 2));
  console.log(`Media written to ${output}; source frames at ${scratch}`);
} finally {
  await browser.close();
}

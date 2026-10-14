#!/usr/bin/env node
// Record Original -> ASCII gameplay -> a real OS clipboard paste into Writer.
// Requires the built demo at :8082, shareware IWAD, ffmpeg with x11grab/libass,
// Xvfb, LibreOffice, and a Python with python3-uno/Pillow (normally /usr/bin/python3).
// FFMPEG, XVFB, LIBREOFFICE, LO_PYTHON and DOOM_CAPTURE_OUTPUT override tools/output.
// DOOM_ENGINE_PATH optionally mirrors the exact pinned engine; its hash is checked.
import { chromium } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdirSync, mkdtempSync, openSync, closeSync } from 'node:fs';
import { dirname, resolve, join } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { spawn, spawnSync } from 'node:child_process';
import { createInterface } from 'node:readline';
import { once } from 'node:events';
import { createServer } from 'node:net';

const here = dirname(fileURLToPath(import.meta.url));
const output = resolve(process.env.DOOM_CAPTURE_OUTPUT ?? join(here, '../../docs/images'));
const scratch = mkdtempSync(join(tmpdir(), 'docxodus-ascii-capture-'));
const ffmpeg = process.env.FFMPEG ?? 'ffmpeg';
const viewport = { width: 1100, height: 1050 };
const preview = process.argv.includes('--preview');
const engineSha = 'efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9';
const wadSha = '1d7d43be501e67d927e415e0b8f3e29c3bf33075e859721816f652a526cac771';
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const wad = gunzipSync(readFileSync(resolve(here, '../dist/wasm/vendor/doom1.wad.gz')));
if (sha(wad) !== wadSha) throw new Error('Expected the original DOOM 1.9 shareware IWAD');
mkdirSync(output, { recursive: true });
const file = name => join(output, `arcade-doom-ascii-${name}`);
const run = (exe, args) => {
  const r = spawnSync(exe, args, { encoding: 'utf8' });
  if (r.error || r.status !== 0) throw new Error(r.error?.message ?? r.stderr);
  return r.stdout;
};
let xserver, writer, browser, recorder, recordExit, rpc, recordingStarted;
const logs = [];
const logFd = name => { const fd = openSync(join(scratch, name), 'w'); logs.push(fd); return fd; };
const captions = [];
const stage = text => {
  captions.push({ time: (performance.now() - recordingStarted) / 1000, text });
  console.log(text);
};
const errors = [];
try {
  let env = { ...process.env };
  if (!preview) {
    run(ffmpeg, ['-version']);
    xserver = spawn(process.env.XVFB ?? 'Xvfb', ['-displayfd', '1', '-screen', '0', '1100x1050x24', '-nolisten', 'tcp'],
      { stdio: ['ignore', 'pipe', logFd('xvfb.log')] });
    const lines = createInterface({ input: xserver.stdout });
    const display = await Promise.race([
      once(lines, 'line').then(([line]) => line.trim()),
      once(xserver, 'error').then(([error]) => { throw error; }),
      once(xserver, 'exit').then(() => { throw new Error(`Xvfb exited; see ${scratch}`); }),
    ]);
    env = { ...env, DISPLAY: `:${display}`, WAYLAND_DISPLAY: '', XDG_SESSION_TYPE: 'x11', SAL_USE_VCLPLUGIN: 'gen' };
    // Reserve an unused local port for this private Writer profile.
    const listener = createServer();
    listener.listen(0, '127.0.0.1'); await once(listener, 'listening');
    const port = listener.address().port;
    await new Promise(resolve => listener.close(resolve));
    writer = spawn(process.env.LO_PYTHON ?? '/usr/bin/python3', [join(here, 'capture-doom-libreoffice.py'), scratch, String(port)],
      { env, stdio: ['pipe', 'pipe', logFd('writer-helper.log')] });
    const pending = [];
    const replies = createInterface({ input: writer.stdout });
    replies.on('line', line => {
      const next = pending.shift();
      if (!next) return;
      try { const value = JSON.parse(line); value.error ? next.reject(new Error(value.error)) : next.resolve(value); }
      catch (error) { next.reject(error); }
    });
    writer.on('error', error => { for (const item of pending.splice(0)) item.reject(error); });
    writer.on('exit', () => { for (const item of pending.splice(0)) item.reject(new Error(`Writer exited; see ${scratch}`)); });
    const response = () => new Promise((resolve, reject) => pending.push({ resolve, reject }));
    rpc = request => { const result = response(); writer.stdin.write(JSON.stringify(request) + '\n'); return result; };
    await response();
  }
  browser = await chromium.launch({ headless: preview, env,
    args: preview ? [] : ['--ozone-platform=x11', '--kiosk', '--window-position=0,0', '--window-size=1100,1050'] });
  const context = await browser.newContext({ viewport, deviceScaleFactor: preview ? 2 : 1 });
  const page = await context.newPage();
  if (!preview) {
    // Context windows do not always inherit --kiosk. Match the physical desktop
    // to the emulated viewport so browser chrome cannot crop the bottom dock.
    const windowSession = await context.newCDPSession(page);
    const { windowId } = await windowSession.send('Browser.getWindowForTarget');
    await windowSession.send('Browser.setWindowBounds', { windowId, bounds: { windowState: 'fullscreen' } });
  }
  page.on('pageerror', e => errors.push(e.message));
  if (process.env.DOOM_ENGINE_PATH) {
    const bytes = readFileSync(process.env.DOOM_ENGINE_PATH);
    if (sha(bytes) !== engineSha) throw new Error('Engine mirror does not match the cartridge pin');
    await page.route('https://cdn.jsdelivr.net/gh/grubbyplaya/**', route =>
      route.fulfill({ body: bytes, contentType: 'text/javascript' }));
  }
  const url = new URL('/demo-arcade.html', process.env.ARCADE_URL ?? 'http://localhost:8082');
  url.search = new URLSearchParams({ engine: './embed.bundle.js', intro: '0', sound: '0',
    cart: 'doom', render: 'image', wad: './vendor/doom1.wad.gz' }).toString();
  await page.goto(url.href);
  await page.waitForFunction(() => window.__arcade?.game().doomFrames >= 2, null, { timeout: 180000 });
  await page.selectOption('#pace', '0');
  await page.waitForFunction(() => document.querySelector('[data-dxr="loader"]').hidden);
  await page.evaluate(() => document.fonts.ready);
  const frameBox = () => page.evaluate(() => {
    const { x, y, width, height } = window.__arcade.canvasElement().getBoundingClientRect();
    return { x, y, width, height };
  });
  await page.screenshot({ path: file('title.png') });
  if (preview) { console.log(`Original-mode title written to ${file('title.png')}`); }
  else {
    const raw = join(scratch, 'desktop.mkv');
    recorder = spawn(ffmpeg, ['-hide_banner', '-loglevel', 'error', '-y', '-f', 'x11grab',
      '-framerate', '20', '-video_size', '1100x1050', '-i', env.DISPLAY, '-an',
      '-c:v', 'libx264rgb', '-preset', 'ultrafast', '-crf', '0', '-threads', '2', raw],
      { env, stdio: ['pipe', 'ignore', logFd('recording.log')] });
    recordExit = once(recorder, 'exit');
    recorder.on('error', error => errors.push(error.message));
    recordingStarted = performance.now();
    stage('Original mode: Doom inside a Word document');
    await page.waitForTimeout(1300);
    await page.keyboard.press('Enter');
    await page.waitForTimeout(800);
    await page.screenshot({ path: file('menu.png') });
    const sourceMenu = await page.evaluate(() => {
      const canvas = document.createElement('canvas'); canvas.width = 320; canvas.height = 200;
      const ctx = canvas.getContext('2d'), data = ctx.createImageData(320, 200);
      for (let i = 0; i < 64000; i++) data.data.set([...window.__arcade.game().pixel(i % 320, Math.floor(i / 320)), 255], i * 4);
      ctx.putImageData(data, 0, 0); return canvas.toDataURL().split(',')[1];
    });
    writeFileSync(file('menu-source.png'), Buffer.from(sourceMenu, 'base64'));
    for (let i = 0; i < 2; i++) { await page.keyboard.press('Enter'); await page.waitForTimeout(500); }
    await page.keyboard.press('ArrowUp');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(1200);
    const hold = async (key, ms) => {
      await page.keyboard.down(key); await page.waitForTimeout(ms); await page.keyboard.up(key);
    };
    await hold('KeyW', 900); await hold('ArrowLeft', 400); await hold('Space', 800);
    await page.screenshot({ path: file('original-gameplay.png') });
    const original = await page.evaluate(() => ({ rendering: window.__arcade.game().rendering,
      images: window.__arcade.canvasElement().querySelectorAll('img').length }));
    if (original.rendering !== 'image' || original.images !== 1) throw new Error('Original mode was not recorded');
    stage('Switch to ASCII, then keep moving and firing');
    // Use the visible selector, so the recording shows the actual mode switch.
    await page.locator('#rendering').click();
    await page.waitForTimeout(600);
    await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
    await page.waitForFunction(() => window.__arcade.game().rendering === 'ascii');
    await page.waitForTimeout(1000);
    const start = await page.evaluate(() => ({ time: performance.now(), frames: window.__arcade.frames(),
      sig: window.__arcade.canvasElement().getAttribute('data-render-sig') }));
    await hold('ArrowRight', 600); await hold('KeyW', 1000); await hold('Space', 1000);
    await hold('ArrowLeft', 550); await hold('KeyS', 750); await hold('Space', 900);
    const end = await page.evaluate(() => ({ time: performance.now(), frames: window.__arcade.frames(),
      sig: window.__arcade.canvasElement().getAttribute('data-render-sig') }));
    if (start.sig === end.sig || end.frames <= start.frames) throw new Error('ASCII gameplay did not advance');
    await page.screenshot({ path: file('gameplay.png') });
    stage('Pause. Select the dense ASCII frame and copy it');
    await page.click('#playpause'); await page.waitForTimeout(700);
    await page.screenshot({ path: file('frame.png'), clip: await frameBox() });
    const snapshot = await page.evaluate(async () => {
      const a = window.__arcade, el = a.canvasElement();
      const { rowsFromXml } = await import('/ascii-arcade.js');
      return { text: rowsFromXml(a.session.raw.getXml(a.canvasAnchor())).join('\n'),
        saved: Array.from(a.save()), engineTitle: a.game().title, rendering: a.game().rendering,
        textCharacters: el.textContent.length, images: el.querySelectorAll('img,canvas,svg').length,
        incremental: a.editor.lastReconcileFallback === null, playing: a.playing() };
    });
    if (snapshot.images || snapshot.textCharacters !== 64200 || !snapshot.incremental || snapshot.playing)
      throw new Error('Paused frame is not complete native ASCII text');
    writeFileSync(file('frame.docx'), Buffer.from(snapshot.saved)); delete snapshot.saved;
    writeFileSync(join(scratch, 'clipboard-source.txt'), snapshot.text, 'ascii');
    await page.evaluate(() => {
      const el = window.__arcade.canvasElement(); el.focus();
      const range = document.createRange(); range.selectNodeContents(el);
      const selection = getSelection(); selection.removeAllRanges(); selection.addRange(range);
    });
    await page.waitForTimeout(1000);
    await page.keyboard.press('Control+c'); await page.waitForTimeout(700);
    stage('LibreOffice Writer: paste as unformatted text');
    await rpc({ action: 'focus' }); await page.waitForTimeout(1100);
    const proof = await rpc({ action: 'paste', expected: join(scratch, 'clipboard-source.txt'),
      odt: file('libreoffice.odt'), text: file('clipboard.txt') });
    if (proof.printableCharacters !== 64200 || proof.lineBreaks !== 199 || proof.textFrames)
      throw new Error(`Unexpected Writer content: ${JSON.stringify(proof)}`);
    stage('64,200 printable ASCII characters + 199 line breaks. Zero images.');
    await page.waitForTimeout(3200);
    await rpc({ action: 'screenshot', path: file('libreoffice.png') });
    stage('Select a HUD text sample and enlarge it to 12 pt');
    const offset = 174 * 322 + 55, length = 90;
    const detail = await rpc({ action: 'select', offset, length });
    if (detail.selectedText !== snapshot.text.slice(offset, offset + length)) throw new Error('HUD selection does not match source');
    await page.waitForTimeout(1500);
    const enlarged = await rpc({ action: 'enlarge' });
    if (!enlarged.charactersUnchanged || enlarged.fontPoints !== 12) throw new Error('Font change did not preserve the ASCII text');
    await page.waitForTimeout(1700);
    stage('Editable letters and punctuation, copied from the live game');
    await rpc({ action: 'screenshot', path: file('libreoffice-detail.png') });
    await page.waitForTimeout(3000);
    const duration = (performance.now() - recordingStarted) / 1000;
    recorder.stdin.write('q\n'); await recordExit; recorder = null;
    // Preserve high-resolution source stills separately from the desktop video.
    await page.bringToFront();
    await page.evaluate(() => getSelection().removeAllRanges());
    const box = await frameBox();
    const cdp = await context.newCDPSession(page);
    await cdp.send('Emulation.setDeviceMetricsOverride', { ...viewport, deviceScaleFactor: 4, mobile: false });
    await page.screenshot({ path: file('detail.png'),
      clip: { x: box.x, y: box.y + box.height * .73, width: box.width, height: box.height * .27 } });
    if (errors.length) throw new Error(errors.join('\n'));
    const assTime = seconds => {
      const cs = Math.round(seconds * 100);
      return `${Math.floor(cs / 360000)}:${String(Math.floor(cs / 6000) % 60).padStart(2, '0')}:${String(Math.floor(cs / 100) % 60).padStart(2, '0')}.${String(cs % 100).padStart(2, '0')}`;
    };
    const subtitles = join(scratch, 'captions.ass');
    writeFileSync(subtitles, `[Script Info]\nScriptType: v4.00+\nPlayResX: 1100\nPlayResY: 1140\n\n[V4+ Styles]\nFormat: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding\nStyle: Default,DejaVu Sans,25,&H00FFFFFF,&H00FFFFFF,&H00000000,&H00000000,0,0,0,0,100,100,0,0,1,0,0,2,20,20,29,1\n\n[Events]\nFormat: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text\n` +
      captions.map((c, i) => `Dialogue: 0,${assTime(c.time)},${assTime(captions[i + 1]?.time ?? duration)},Default,,0,0,0,,${c.text}\n`).join(''));
    const filter = `pad=1100:1140:0:0:color=0x111827,ass=${subtitles}`;
    run(ffmpeg, ['-hide_banner', '-loglevel', 'error', '-y', '-i', raw, '-t', String(duration), '-an',
      '-vf', filter, '-c:v', 'libx264', '-preset', 'slow', '-crf', '18', '-pix_fmt', 'yuv420p',
      '-movflags', '+faststart', join(output, 'arcade-doom-ascii.mp4')]);
    run(ffmpeg, ['-hide_banner', '-loglevel', 'error', '-y', '-i', raw, '-t', String(duration), '-an',
      '-filter_complex', `${filter},fps=12,scale=880:-2:flags=lanczos,format=rgb24,split[a][b];[a]palettegen=max_colors=256:reserve_transparent=0:stats_mode=diff[p];[b][p]paletteuse=dither=bayer:bayer_scale=4`,
      '-gifflags', 'offsetting', '-loop', '0', join(output, 'arcade-doom-ascii.gif')]);
    delete snapshot.text;
    const metadata = { ...snapshot, source: 'Original DOOM v1.9 shareware', engineSha256: engineSha,
      wadSha256: wadSha, viewport, gif: { width: 880, height: 912, framesPerSecond: 12,
        paletteColors: 256, transparentFrames: false }, durationSeconds: duration,
      documentGameplayFps: (end.frames - start.frames) * 1000 / (end.time - start.time),
      desktopCaptureFramesPerSecond: 20,
      original, libreoffice: { ...proof, ...enlarged, selectedHudText: detail.selectedText }, captions,
      capture: 'Continuous real-time X11 recording of Chromium and LibreOffice Writer. Actual OS clipboard copy and unformatted paste; captions added below the desktop. No time acceleration.' };
    writeFileSync(file('capture.json'), JSON.stringify(metadata, null, 2) + '\n');
    console.log(JSON.stringify(metadata, null, 2));
    console.log(`Media written to ${output}; lossless desktop recording at ${scratch}`);
  }
} finally {
  if (recorder) { recorder.stdin.write('q\n'); await recordExit; }
  await browser?.close();
  if (writer && writer.exitCode === null) { try { await rpc({ action: 'quit' }); } catch {} writer.stdin.end(); await once(writer, 'exit'); }
  xserver?.kill();
  for (const fd of logs) closeSync(fd);
}

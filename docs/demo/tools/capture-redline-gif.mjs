// Capture a frame sequence of REDLINE THEATER running, for the demo GIF.
//
// NOT part of the test suite and not run by CI — a one-shot recorder, kept so the
// committed GIF can be regenerated when the show changes rather than being a
// mystery binary. Frames land as PNGs; `frames-to-gif.py` beside it assembles
// them. The behaviour the GIF depicts is what npm/tests/demo-redline.spec.ts
// actually asserts.
import { chromium } from '@playwright/test';
import { spawn } from 'node:child_process';
import { mkdirSync, rmSync } from 'node:fs';

// Run from the npm/ directory, which is where @playwright/test and the built
// webroot both live:
//   cd npm && node ../docs/demo/tools/capture-redline-gif.mjs
//   FRAME_DIR=/tmp/redline-frames WIDTH=900 FPS=8 COLORS=48 \
//     python3 ../docs/demo/tools/frames-to-gif.py ../docs/images/demo/redline-theater.gif

const OUT = process.env.FRAME_DIR ?? '/tmp/redline-frames';
const FPS = Number(process.env.FPS ?? 10);
const SECONDS = Number(process.env.SECONDS ?? 26);
const SPEED = process.env.SPEED ?? 'brisk';

rmSync(OUT, { recursive: true, force: true });
mkdirSync(OUT, { recursive: true });

const server = spawn('python3', ['-m', 'http.server', '8899'], {
  cwd: process.env.WEBROOT ?? 'dist/wasm', stdio: 'ignore',
});
await new Promise((r) => setTimeout(r, 1200));

const browser = await chromium.launch();
const page = await browser.newPage({
  viewport: {
    width: Number(process.env.VW ?? 1280),
    height: Number(process.env.VH ?? 720),
  },
});
page.on('pageerror', (e) => console.log('[pageerror]', String(e).slice(0, 300)));

await page.goto('http://127.0.0.1:8899/demo-redline.html?engine=./embed.bundle.js');
await page.waitForFunction(
  () => window.__theater !== undefined || window.__theaterError !== undefined,
  { timeout: 120000 });
const err = await page.evaluate(() => window.__theaterError);
if (err) { console.error('boot failed:', err); process.exit(1); }
await page.evaluate(() => window.__theater.ready);
await page.evaluate((s) => window.__theater.setSpeed(s), SPEED);

// Let the clean baseline sit on screen for a beat before the first mark lands,
// so the GIF opens on "an ordinary contract" rather than mid-edit.
const total = FPS * SECONDS;
let frame = 0;
const shoot = async () => {
  await page.screenshot({ path: `${OUT}/f${String(frame).padStart(4, '0')}.png` });
  frame++;
};
for (let i = 0; i < FPS; i++) { await shoot(); await page.waitForTimeout(1000 / FPS); }

// Fire and forget: `run()` resolves only when the whole negotiation is over, so
// awaiting it here would start the recording after the show had finished.
await page.evaluate(() => { void window.__theater.run(); });

const interval = 1000 / FPS;
while (frame < total) {
  const started = Date.now();
  await shoot();
  const done = await page.evaluate(() => window.__theater.redline() != null);
  // Once the proof is up, hold on it for a couple of seconds and stop — the
  // verdict is the payoff and the loop should end there, not on a blank tail.
  if (done) {
    const hold = FPS * 3;
    for (let i = 0; i < hold && frame < total; i++) {
      await shoot();
      await page.waitForTimeout(interval);
    }
    break;
  }
  const elapsed = Date.now() - started;
  if (elapsed < interval) await page.waitForTimeout(interval - elapsed);
}

console.log(`captured ${frame} frames at ${FPS}fps into ${OUT}`);
console.log('stats =', JSON.stringify(await page.evaluate(() => window.__theater.stats())));
await browser.close();
server.kill();

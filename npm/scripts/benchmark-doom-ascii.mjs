#!/usr/bin/env node
// Run after npm run build && npm run pretest, with dist/wasm served at :8082.
// DOOM_ENGINE_PATH can mirror the pinned Doom engine for an offline run.
// ARCADE_URL, DOOM_BENCH_WAD, DOOM_BENCH_OUTPUT and DOOM_BENCH_BROWSER are optional.
// DOOM_BENCH_BROWSER selects chromium (default), firefox or webkit.
// DOOM_BENCH_EXECUTABLE can select a matching custom Playwright browser build.
// The CPU multiplier is Chromium's CPU throttling rate (1 = this machine).
import { chromium, firefox, webkit } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { cpus, platform } from 'node:os';

const output = resolve(process.env.DOOM_BENCH_OUTPUT ?? 'test-results/doom-ascii-benchmark');
const cpuRate = Number(process.env.DOOM_BENCH_CPU ?? 1);
if (!Number.isFinite(cpuRate) || cpuRate < 1) throw new Error('DOOM_BENCH_CPU must be at least 1');
const browserName = process.env.DOOM_BENCH_BROWSER ?? 'chromium';
const browserType = { chromium, firefox, webkit }[browserName];
if (!browserType) throw new Error('DOOM_BENCH_BROWSER must be chromium, firefox or webkit');
if (browserName !== 'chromium' && cpuRate !== 1) throw new Error('CPU throttling requires Chromium');
mkdirSync(output, { recursive: true });
const executablePath = process.env.DOOM_BENCH_EXECUTABLE
  ?? (browserName === 'chromium' ? process.env.DOCXODUS_CHROMIUM_PATH : undefined);
const browser = await browserType.launch({ headless: true,
  ...(executablePath ? { executablePath } : {}) });
try {
  const page = await browser.newPage({ viewport: { width: 1300, height: 1100 }, deviceScaleFactor: 2 });
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  if (process.env.DOOM_ENGINE_PATH) {
    const bytes = readFileSync(process.env.DOOM_ENGINE_PATH);
    if (createHash('sha256').update(bytes).digest('hex') !==
        'efd7c34714e3753a84cf6873f2c2ae0e1a41bde1f10aaa2b0a95cb6935e8cce9')
      throw new Error('DOOM_ENGINE_PATH does not match the cartridge pin');
    await page.route('https://cdn.jsdelivr.net/gh/grubbyplaya/**', route =>
      route.fulfill({ body: bytes, contentType: 'text/javascript' }));
  }
  const url = new URL('/demo-arcade.html', process.env.ARCADE_URL ?? 'http://localhost:8082');
  url.search = new URLSearchParams({ engine: './embed.bundle.js', intro: '0', sound: '0',
    cart: 'doom', wad: process.env.DOOM_BENCH_WAD ?? './vendor/freedoom1.wad.gz' }).toString();
  await page.goto(url.href);
  await page.waitForFunction(() => {
    const state = window.__arcade?.game();
    if (state?.status === 'error') throw new Error(state.error);
    return state?.doomFrames >= 4;
  }, null, { timeout: 180000 });
  await page.selectOption('#pace', '0');
  for (let i = 0; i < 4; i++) {
    await page.keyboard.press('Enter');
    await page.waitForTimeout(700);
  }
  await page.waitForTimeout(1500);
  await page.selectOption('#rendering', 'ascii');
  await page.evaluate(() => document.fonts.ready);
  const cdp = browserName === 'chromium' ? await page.context().newCDPSession(page) : null;
  await cdp?.send('Emulation.setCPUThrottlingRate', { rate: cpuRate });
  await page.waitForTimeout(3000);

  // Observe the frame number at animation frames, after synchronous document
  // refresh. Multiple mutations before the browser can present count once.
  // Read a signature too, so a frozen document cannot pass a moving workload.
  await page.evaluate(() => {
    window.__doomBench = { samples: [], stopped: false };
    let prior = -1;
    const observe = t => {
      const b = window.__doomBench;
      if (b.stopped) return;
      const a = window.__arcade, n = a.frames();
      if (n !== prior) {
        b.samples.push({ t, n, sig: a.canvasElement().getAttribute('data-render-sig') });
        prior = n;
      }
      requestAnimationFrame(observe);
    };
    requestAnimationFrame(observe);
  });
  const phases = [];
  for (const [name, keys] of [['stationary', []], ['turn', ['ArrowRight']],
    ['advance-fire', ['KeyW', 'Space']], ['turn-fire', ['ArrowLeft', 'Space']]]) {
    for (const key of keys) await page.keyboard.down(key);
    await page.evaluate(() => { window.__doomBench.samples = []; });
    await page.waitForTimeout(12000);
    const result = await page.evaluate(() => {
      const samples = window.__doomBench.samples;
      const first = samples[0], last = samples.at(-1);
      if (samples.length < 2) throw new Error('No presented frames');
      const gaps = samples.slice(1).map((s, i) => s.t - samples[i].t).sort((a, b) => a - b);
      return { seconds: (last.t - first.t) / 1000,
        presentedFps: (samples.length - 1) * 1000 / (last.t - first.t),
        documentFps: (last.n - first.n) * 1000 / (last.t - first.t),
        p95FrameMs: gaps[Math.floor(gaps.length * .95)],
        changedFrames: samples.slice(1).filter((s, i) => s.sig !== samples[i].sig).length,
        timings: window.__arcade.timings() };
    });
    for (const key of keys) await page.keyboard.up(key);
    phases.push({ name, ...result });
    console.log(`${name}: ${result.presentedFps.toFixed(2)} presented fps, p95 ${result.p95FrameMs.toFixed(1)} ms`);
  }
  await page.evaluate(() => { window.__doomBench.stopped = true; window.__arcade.pause(); });
  await cdp?.send('Emulation.setCPUThrottlingRate', { rate: 1 });
  await page.screenshot({ path: resolve(output, 'gameplay.png') });
  const proof = await page.evaluate(async () => {
    const a = window.__arcade;
    const { asciiFramebuffer } = await import('/doom-ascii.js');
    const { rowsFromXml } = await import('/ascii-arcade.js');
    const fb = new Uint8Array(320 * 200 * 4);
    const source = document.createElement('canvas'); source.width = 320; source.height = 200;
    const ctx = source.getContext('2d'), pixels = ctx.createImageData(320, 200);
    for (let i = 0; i < 320 * 200; i++) {
      const [r, g, b] = a.game().pixel(i % 320, Math.floor(i / 320));
      fb.set([b, g, r, 255], i * 4); pixels.data.set([r, g, b, 255], i * 4);
    }
    ctx.putImageData(pixels, 0, 0);
    const expected = asciiFramebuffer(fb).chars.map(row => row.join(''));
    const xml = a.session.raw.getXml(a.canvasAnchor()), rows = rowsFromXml(xml);
    const el = a.canvasElement();
    const saved = a.save();
    const handle = a.bridge.OpenSession(saved, '');
    let html;
    try { html = a.bridge.RenderHtml(handle, 'verify-', false, false, 1); }
    finally { a.bridge.CloseSession(handle); }
    const reopened = new DOMParser().parseFromString(html, 'text/html');
    return { title: a.game().title, rows: rows.length, columns: 320,
      exactSourceCells: JSON.stringify(rows) === JSON.stringify(expected),
      domMatchesDocument: el.textContent.replace(/\u00a0/g, ' ') === rows.join(''),
      savedTextMatches: [...reopened.querySelectorAll('p')].some(p =>
        p.textContent.replace(/\u00a0/g, ' ') === rows.join('')),
      images: el.querySelectorAll('img, canvas, svg').length,
      breaks: el.querySelectorAll('br').length, fallback: a.editor.lastReconcileFallback,
      sourcePng: source.toDataURL().split(',')[1], saved: Array.from(saved) };
  });
  writeFileSync(resolve(output, 'source.png'), Buffer.from(proof.sourcePng, 'base64'));
  writeFileSync(resolve(output, 'frame.docx'), Buffer.from(proof.saved));
  delete proof.sourcePng; delete proof.saved;
  const report = { date: new Date().toISOString(), browserName, browser: browser.version(), executablePath,
    platform: platform(), cpu: cpus()[0]?.model, cpuRate, url: url.href, phases, proof, errors };
  writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2) + '\n');
  if (errors.length || !proof.exactSourceCells || !proof.domMatchesDocument || !proof.savedTextMatches
      || proof.rows !== 200 || proof.breaks !== 199 || proof.images || proof.fallback)
    throw new Error(`Fidelity check failed; see ${output}/report.json`);
  if (phases.some(p => p.presentedFps < 10 || (p.name !== 'stationary' && p.changedFrames < 20)))
    throw new Error(`10 FPS gameplay target not met; see ${output}/report.json`);
  console.log(`10 FPS and full-resolution document checks passed. Artifacts: ${output}`);
} finally {
  await browser.close();
}

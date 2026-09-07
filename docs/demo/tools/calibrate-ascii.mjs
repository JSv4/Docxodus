// Measure glyph coverage in the actual document surface, including the pinned
// font, authored metrics, synthetic bold, and browser rasterization.
// Serve npm/dist/wasm on :8082 after npm run pretest. Run at DPR=1 and DPR=2;
// --all measures every printable ASCII glyph when evaluating a new ramp.
import { chromium } from '../../../npm/node_modules/@playwright/test/index.mjs';
import { ASCII_RAMP } from '../doom-ascii.js';

const glyphs = process.argv.includes('--all')
  ? Array.from({ length: 95 }, (_, i) => String.fromCharCode(i + 32)).join('') : ASCII_RAMP;
const dpr = Number(process.env.DPR ?? 2);
const browser = await chromium.launch({ headless: true });
try {
  const page = await browser.newPage({ viewport: { width: 1100, height: 1050 }, deviceScaleFactor: dpr });
  await page.goto(`${process.env.ARCADE_URL ?? 'http://localhost:8082'}/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0`);
  await page.waitForFunction(() => window.__arcade?.frames() > 0);
  await page.evaluate(() => window.__arcade.pause());
  await page.waitForFunction(() => document.querySelector('[data-dxr="loader"]').hidden);
  const samples = [];
  for (const char of glyphs) {
    const anchor = await page.evaluate(async char => {
      const a = window.__arcade;
      const { frameXml } = await import('/ascii-scenes.js');
      const { ASCII_METRICS } = await import('/doom-ascii.js');
      const chars = Array.from({ length: 200 }, () => Array(321).fill(char));
      const colors = chars.map(row => row.map(() => 'FFFFFF'));
      const old = a.session.raw.getXml(a.canvasAnchor());
      const frame = frameXml(old.slice(0, old.indexOf('>') + 1), { chars, colors }, '000000', ASCII_METRICS);
      const result = a.session.raw.replaceXml(a.canvasAnchor(), frame.xml);
      if (!result.success) throw new Error(JSON.stringify(result));
      a.editor.refresh();
      await document.fonts.ready;
      return a.canvasElement().getAttribute('data-anchor');
    }, char);
    const png = await page.locator(`[data-anchor="${anchor}"]`).screenshot();
    const mean = await page.evaluate(async base64 => {
      const img = new Image(); img.src = `data:image/png;base64,${base64}`; await img.decode();
      const canvas = document.createElement('canvas'); canvas.width = img.width; canvas.height = img.height;
      const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
      const pixels = ctx.getImageData(Math.floor(img.width * .1), Math.floor(img.height * .2),
        Math.floor(img.width * .8), Math.floor(img.height * .5)).data;
      let sum = 0;
      for (let i = 0; i < pixels.length; i += 4) sum += (pixels[i] + pixels[i + 1] + pixels[i + 2]) / 3;
      return sum / (pixels.length / 4);
    }, png.toString('base64'));
    samples.push({ char, mean });
  }
  const peak = Math.max(...samples.map(s => s.mean));
  console.log(JSON.stringify({ dpr, peak, samples: samples.map(s => ({ ...s,
    coverage: Math.round(s.mean / peak * 1000) })) }, null, 2));
} finally {
  await browser.close();
}

import { test, expect } from '@playwright/test';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';

// LibreOffice is intentionally not a dependency of the normal browser suite.
// The scheduled interoperability job opts in and tests one real exported frame
// from each cartridge against LO's 96-dpi PDF rendering.
test.skip(process.env.DOCXODUS_LO_PARITY !== '1',
  'set DOCXODUS_LO_PARITY=1 on a host with libreoffice and pdftoppm');

// intro=0 skips the attract screen: these specs test the cartridges, and the
// intro has its own dedicated coverage.
const OVERRIDE = 'engine=./embed.bundle.js&intro=0';
const backgrounds = {
  quest: [14, 28, 48],
  dungeon: [10, 15, 26],
} as const;

async function loadArcade(page: import('@playwright/test').Page, cart: 'quest' | 'dungeon') {
  await page.goto(`/demo-arcade.html?${OVERRIDE}&cart=${cart}`);
  await page.waitForFunction(() =>
    (window as any).__arcade !== undefined || (window as any).__arcadeError !== undefined,
    { timeout: 90000 },
  );
  const error = await page.evaluate(() => (window as any).__arcadeError);
  expect(error, `arcade boot failed: ${error}`).toBeUndefined();
  await page.waitForFunction(() => (window as any).__arcade.frames() >= 2, { timeout: 30000 });
  return page.evaluate(() => {
    const a = (window as any).__arcade;
    a.pause();
    const unid = (a.canvasAnchor() as string).split(':')[2];
    const canvas = document.querySelector(`[data-anchor="${unid}"]`) as HTMLElement | null;
    return {
      bytes: Array.from(a.save() as Uint8Array),
      unid,
      fallback: a.editor.lastReconcileFallback as string | null,
      canvasStyle: canvas ? {
        tag: canvas.tagName,
        inline: canvas.getAttribute('style'),
        background: getComputedStyle(canvas).backgroundColor,
      } : null,
    };
  });
}

for (const cart of ['quest', 'dungeon'] as const) {
  test(`${cart}: exported frame matches LibreOffice geometry and ink`, async ({ page }, testInfo) => {
    test.setTimeout(120000);
    const work = mkdtempSync(join(tmpdir(), `docxodus-lo-${cart}-`));
    try {
      const frame = await loadArcade(page, cart);
      expect(frame.fallback).toBeNull();

      const docxPath = join(work, `${cart}.docx`);
      const browserPath = join(work, `${cart}-browser.png`);
      const pdfDir = join(work, 'pdf');
      const profileDir = join(work, 'profile');
      const runtimeDir = join(work, 'runtime');
      mkdirSync(pdfDir);
      mkdirSync(profileDir);
      mkdirSync(runtimeDir, { mode: 0o700 });
      writeFileSync(docxPath, Buffer.from(frame.bytes));
      // Fixed demo chrome overlaps the document in normal play. Remove it from this
      // element-level renderer capture so the PNG contains only the DOCX canvas.
      await page.addStyleTag({ content: '#dock, #home, .dxr-loader { display: none !important; }' });
      await page.locator(`[data-anchor="${frame.unid}"]`).screenshot({ path: browserPath });

      execFileSync('libreoffice', [
        `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
        '--headless', '--nologo', '--nodefault', '--nofirststartwizard', '--norestore',
        '--convert-to', 'pdf', '--outdir', pdfDir, docxPath,
      ], {
        env: { ...process.env, XDG_RUNTIME_DIR: runtimeDir },
        stdio: 'pipe',
        timeout: 60000,
      });
      const pdfPath = join(pdfDir, `${cart}.pdf`);
      const loPrefix = join(work, `${cart}-libreoffice`);
      execFileSync('pdftoppm', [
        '-f', '1', '-singlefile', '-r', '96', '-png', pdfPath, loPrefix,
      ], { stdio: 'pipe', timeout: 30000 });
      const loPath = `${loPrefix}.png`;

      const browserPng = readFileSync(browserPath);
      const loPng = readFileSync(loPath);
      await testInfo.attach(`${cart}-frame.docx`, {
        body: Buffer.from(frame.bytes), contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      });
      await testInfo.attach(`${cart}-docxodus.png`, { body: browserPng, contentType: 'image/png' });
      await testInfo.attach(`${cart}-libreoffice.png`, { body: loPng, contentType: 'image/png' });

      const metrics = await page.evaluate(async ({ browser, libreoffice, bg }) => {
        const load = (base64: string) => new Promise<HTMLImageElement>((resolve, reject) => {
          const image = new Image();
          image.onload = () => resolve(image);
          image.onerror = reject;
          image.src = `data:image/png;base64,${base64}`;
        });
        const pixels = (image: HTMLImageElement) => {
          const canvas = document.createElement('canvas');
          canvas.width = image.naturalWidth;
          canvas.height = image.naturalHeight;
          canvas.getContext('2d')!.drawImage(image, 0, 0);
          return { canvas, data: canvas.getContext('2d')!.getImageData(0, 0, canvas.width, canvas.height) };
        };
        const bbox = (data: ImageData) => {
          let x0 = data.width, y0 = data.height, x1 = -1, y1 = -1;
          for (let y = 0; y < data.height; y++) for (let x = 0; x < data.width; x++) {
            const i = (y * data.width + x) * 4;
            if (data.data[i] === bg[0] && data.data[i + 1] === bg[1] && data.data[i + 2] === bg[2]) {
              x0 = Math.min(x0, x); y0 = Math.min(y0, y);
              x1 = Math.max(x1, x); y1 = Math.max(y1, y);
            }
          }
          if (x1 < x0) throw new Error(`canvas background rgb(${bg.join(',')}) not found`);
          return { x: x0, y: y0, width: x1 - x0 + 1, height: y1 - y0 + 1 };
        };
        const crop = (source: HTMLCanvasElement, box: { x: number; y: number; width: number; height: number },
          width = box.width, height = box.height) => {
          const canvas = document.createElement('canvas');
          canvas.width = width; canvas.height = height;
          canvas.getContext('2d')!.drawImage(source,
            box.x, box.y, box.width, box.height, 0, 0, width, height);
          return canvas.getContext('2d')!.getImageData(0, 0, width, height);
        };
        const dominantColor = (data: ImageData) => {
          const counts = new Map<number, number>();
          let best = 0, bestCount = 0;
          for (let i = 0; i < data.data.length; i += 4) {
            const key = (data.data[i] << 16) | (data.data[i + 1] << 8) | data.data[i + 2];
            const count = (counts.get(key) ?? 0) + 1;
            counts.set(key, count);
            if (count > bestCount) { best = key; bestCount = count; }
          }
          return [(best >> 16) & 255, (best >> 8) & 255, best & 255] as [number, number, number];
        };
        const mask = (data: ImageData, background: readonly number[]) => {
          const out = new Uint8Array(data.width * data.height);
          for (let p = 0; p < out.length; p++) {
            const i = p * 4;
            const delta = Math.max(
              Math.abs(data.data[i] - background[0]),
              Math.abs(data.data[i + 1] - background[1]),
              Math.abs(data.data[i + 2] - background[2]),
            );
            out[p] = delta > 20 ? 1 : 0;
          }
          return out;
        };
        const dilate = (input: Uint8Array, width: number, height: number) => {
          const out = new Uint8Array(input.length);
          for (let y = 0; y < height; y++) for (let x = 0; x < width; x++) {
            let on = 0;
            for (let dy = -2; dy <= 2 && !on; dy++) for (let dx = -2; dx <= 2; dx++) {
              const xx = x + dx, yy = y + dy;
              if (xx >= 0 && xx < width && yy >= 0 && yy < height && input[yy * width + xx]) {
                on = 1; break;
              }
            }
            out[y * width + x] = on;
          }
          return out;
        };

        const browserImage = pixels(await load(browser));
        const loImage = pixels(await load(libreoffice));
        // Playwright captures the canvas element itself, so its PNG dimensions are
        // already the browser geometry. Chromium's PNG color management can shift a
        // dark CSS background by a channel value, making exact-color bbox detection
        // unreliable. The LibreOffice PNG is a whole page and still needs cropping.
        const browserBox = {
          x: 0,
          y: 0,
          width: browserImage.data.width,
          height: browserImage.data.height,
        };
        const loBox = bbox(loImage.data);
        // Normalize the (at most one-pixel) raster-boundary difference only for
        // the glyph metric. Geometry is asserted independently below.
        const a = crop(browserImage.canvas, browserBox, loBox.width, loBox.height);
        const b = crop(loImage.canvas, loBox);
        const browserBackground = dominantColor(a);
        const libreofficeBackground = dominantColor(b);
        const rawA = mask(a, browserBackground);
        const rawB = mask(b, libreofficeBackground);
        const am = dilate(rawA, loBox.width, loBox.height);
        const bm = dilate(rawB, loBox.width, loBox.height);
        let intersection = 0, union = 0, matchedA = 0, matchedB = 0;
        for (let i = 0; i < am.length; i++) {
          if (am[i] && bm[i]) intersection++;
          if (am[i] || bm[i]) union++;
          if (rawA[i] && bm[i]) matchedA++;
          if (rawB[i] && am[i]) matchedB++;
        }
        const activeA = rawA.reduce((sum, value) => sum + value, 0);
        const activeB = rawB.reduce((sum, value) => sum + value, 0);
        const coverageA = activeA ? matchedA / activeA : 1;
        const coverageB = activeB ? matchedB / activeB : 1;
        return {
          browser: { width: browserBox.width, height: browserBox.height },
          libreoffice: { width: loBox.width, height: loBox.height },
          backgrounds: { browser: browserBackground, libreoffice: libreofficeBackground },
          activePixels: {
            browser: activeA,
            libreoffice: activeB,
          },
          tolerantInkIoU5x5: union ? intersection / union : 1,
          tolerantInkF1_5x5: coverageA + coverageB
            ? 2 * coverageA * coverageB / (coverageA + coverageB)
            : 1,
        };
      }, {
        browser: browserPng.toString('base64'),
        libreoffice: loPng.toString('base64'),
        bg: [...backgrounds[cart]],
      });

      testInfo.annotations.push({ type: 'libreoffice-parity', description: JSON.stringify(metrics) });
      const details = JSON.stringify({ ...metrics, canvasStyle: frame.canvasStyle });
      expect(Math.abs(metrics.browser.width - metrics.libreoffice.width), details).toBeLessThanOrEqual(1);
      expect(Math.abs(metrics.browser.height - metrics.libreoffice.height), details).toBeLessThanOrEqual(1);
      expect(metrics.backgrounds.browser, details).toEqual([...backgrounds[cart]]);
      expect(metrics.backgrounds.libreoffice, details).toEqual([...backgrounds[cart]]);
      expect(metrics.tolerantInkF1_5x5, details).toBeGreaterThan(0.85);
    } finally {
      rmSync(work, { recursive: true, force: true });
    }
  });
}

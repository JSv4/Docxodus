import { expect, test } from '@playwright/test';
import { compareImages } from './visual-parity/metrics.js';
import { decodePng, encodePng, type RgbaImage } from './visual-parity/png.js';

function image(width: number, height: number, ink?: { x: number; y: number; width: number; height: number }): RgbaImage {
  const data = new Uint8Array(width * height * 4).fill(255);
  if (ink) {
    for (let y = ink.y; y < ink.y + ink.height; y++) {
      for (let x = ink.x; x < ink.x + ink.width; x++) {
        const i = (y * width + x) * 4;
        data[i] = data[i + 1] = data[i + 2] = 0;
      }
    }
  }
  return { width, height, data };
}

test.describe('visual parity metrics', () => {
  test('PNG encoding round-trips RGBA pixels deterministically', () => {
    const original = image(9, 7, { x: 2, y: 3, width: 4, height: 2 });
    const encoded = encodePng(original);
    const decoded = decodePng(encoded);
    expect(decoded.width).toBe(original.width);
    expect(decoded.height).toBe(original.height);
    expect(decoded.data).toEqual(original.data);
    expect(encodePng(decoded)).toEqual(encoded);
  });

  test('PNG decoding rejects corrupt chunks and trailing bytes', () => {
    const encoded = encodePng(image(2, 2));
    const corruptHeader = Buffer.from(encoded);
    corruptHeader[16] ^= 1;
    expect(() => decodePng(corruptHeader)).toThrow(/CRC/);
    expect(() => decodePng(Buffer.concat([encoded, Buffer.from([0])]))).toThrow(/trailing bytes/);
  });

  test('PNG dimensions are bounded before allocation or inflation', () => {
    expect(() => encodePng({ width: 4_000_001, height: 1, data: new Uint8Array() }))
      .toThrow(/pixel limit/);
  });

  test('identical pages are exact, structurally identical, and close', () => {
    const original = image(32, 32, { x: 8, y: 8, width: 10, height: 12 });
    const comparison = compareImages(original, original);
    const result = comparison.metrics;
    expect(result.alignment).toMatchObject({ dx: 0, dy: 0 });
    expect(result.exactDiffPixelRatio).toBe(0);
    expect(result.perceptualDiffPixelRatio).toBe(0);
    expect(result.ssim).toBeCloseTo(1, 10);
    expect(result.tolerantInkPrecision).toBeCloseTo(1, 10);
    expect(result.tolerantInkRecall).toBeCloseTo(1, 10);
    expect(result.tolerantInkF1).toBeCloseTo(1, 10);
    expect(result.severity).toBe('close');
    expect(comparison.overlay.data[0]).toBe(255);
    expect(comparison.overlay.data[1]).toBe(255);
    expect(comparison.overlay.data[2]).toBe(255);
  });

  test('bounded alignment normalizes a one-pixel raster-origin shift and reports it', () => {
    const a = image(40, 40, { x: 12, y: 10, width: 9, height: 13 });
    const b = image(40, 40, { x: 13, y: 11, width: 9, height: 13 });
    const result = compareImages(a, b).metrics;
    expect(result.alignment).toMatchObject({ dx: -1, dy: -1, searchRadius: 2 });
    expect(result.perceptualDiffPixelRatio).toBe(0);
    expect(result.ssim).toBeCloseTo(1, 10);
  });

  test('a material content change produces perceptual differences and a red heatmap', () => {
    const a = image(32, 32, { x: 5, y: 5, width: 8, height: 8 });
    const b = image(32, 32, { x: 18, y: 18, width: 8, height: 8 });
    const result = compareImages(a, b);
    expect(result.metrics.perceptualDiffPixelRatio).toBeGreaterThan(0.05);
    expect(result.metrics.ssim).toBeLessThan(0.9);
    const redPixels = Array.from(result.overlay.data)
      .filter((value, index) => index % 4 === 0 && value === 255).length;
    expect(redPixels).toBeGreaterThan(0);
  });

  test('ink F1 is zero when only one renderer produces content', () => {
    const blank = image(32, 32);
    const content = image(32, 32, { x: 8, y: 8, width: 12, height: 12 });
    const result = compareImages(blank, content).metrics;
    expect(result.tolerantInkPrecision).toBe(0);
    expect(result.tolerantInkRecall).toBe(0);
    expect(result.tolerantInkF1).toBe(0);
    expect(result.severity).toBe('severe');
  });

  test('ink precision and recall retain their Docxodus/reference direction', () => {
    const smallDocxodus = image(32, 32, { x: 10, y: 10, width: 4, height: 4 });
    const largeReference = image(32, 32, { x: 8, y: 8, width: 12, height: 12 });
    const docxodusSubset = compareImages(smallDocxodus, largeReference).metrics;
    expect(docxodusSubset.tolerantInkPrecision).toBe(1);
    expect(docxodusSubset.tolerantInkRecall).toBeLessThan(1);

    const docxodusSuperset = compareImages(largeReference, smallDocxodus).metrics;
    expect(docxodusSuperset.tolerantInkPrecision).toBeLessThan(1);
    expect(docxodusSuperset.tolerantInkRecall).toBe(1);
    expect(docxodusSuperset.tolerantInkF1).toBeCloseTo(docxodusSubset.tolerantInkF1, 10);
  });
});

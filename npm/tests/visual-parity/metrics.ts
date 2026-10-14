import type { RgbaImage } from './png.js';

export type VisualSeverity = 'close' | 'minor' | 'major' | 'severe';

export const VISUAL_THRESHOLDS = {
  alignmentSearchPx: 2,
  justNoticeableDeltaE: 2.3,
  inkBackgroundDistance: 24,
  inkToleranceRadiusPx: 2,
  geometryTolerancePx: 1,
  close: { ssim: 0.98, inkF1: 0.95, perceptualDiffRatio: 0.02 },
  minor: { ssim: 0.95, inkF1: 0.90, perceptualDiffRatio: 0.05 },
  major: { ssim: 0.85, inkF1: 0.75, perceptualDiffRatio: 0.15 },
} as const;

export interface PageMetrics {
  docxodus: { width: number; height: number; background: [number, number, number] };
  libreoffice: { width: number; height: number; background: [number, number, number] };
  alignment: { dx: number; dy: number; searchRadius: number };
  widthDelta: number;
  heightDelta: number;
  exactDiffPixelRatio: number;
  perceptualDiffPixelRatio: number;
  meanDeltaE76: number;
  ssim: number;
  /** Fraction of Docxodus ink lying within the tolerance radius of reference ink. */
  tolerantInkPrecision: number;
  /** Fraction of reference ink lying within the tolerance radius of Docxodus ink. */
  tolerantInkRecall: number;
  tolerantInkF1: number;
  severity: VisualSeverity;
}

interface Pixel {
  r: number;
  g: number;
  b: number;
}

/** Dominant border color — the page background. Exported so the Word-reference capture
 * (word-reference.ts) measures ink with the SAME background model as the pairwise metrics. */
export function background(image: RgbaImage): [number, number, number] {
  const counts = new Map<number, number>();
  const border = Math.min(12, Math.ceil(Math.min(image.width, image.height) / 4));
  let best = 0xffffff;
  let bestCount = -1;
  for (let y = 0; y < image.height; y++) {
    for (let x = 0; x < image.width; x++) {
      if (x >= border && x < image.width - border && y >= border && y < image.height - border) continue;
      const i = (y * image.width + x) * 4;
      const key = (image.data[i] << 16) | (image.data[i + 1] << 8) | image.data[i + 2];
      const count = (counts.get(key) ?? 0) + 1;
      counts.set(key, count);
      if (count > bestCount) {
        bestCount = count;
        best = key;
      }
    }
  }
  return [
    (best >> 16) & 255,
    (best >> 8) & 255,
    best & 255,
  ];
}

function pixelAt(image: RgbaImage, x: number, y: number, fallback: readonly number[]): Pixel {
  if (x < 0 || y < 0 || x >= image.width || y >= image.height) {
    return { r: fallback[0], g: fallback[1], b: fallback[2] };
  }
  const i = (y * image.width + x) * 4;
  const alpha = image.data[i + 3] / 255;
  return {
    r: image.data[i] * alpha + fallback[0] * (1 - alpha),
    g: image.data[i + 1] * alpha + fallback[1] * (1 - alpha),
    b: image.data[i + 2] * alpha + fallback[2] * (1 - alpha),
  };
}

function colorDistance(pixel: Pixel, bg: readonly number[]): number {
  return Math.max(Math.abs(pixel.r - bg[0]), Math.abs(pixel.g - bg[1]), Math.abs(pixel.b - bg[2]));
}

function alignmentScore(
  a: RgbaImage,
  b: RgbaImage,
  bgA: readonly number[],
  bgB: readonly number[],
  dx: number,
  dy: number,
): number {
  const width = Math.max(a.width, b.width);
  const height = Math.max(a.height, b.height);
  let difference = 0;
  let active = 0;
  for (let y = 0; y < height; y += 2) {
    for (let x = 0; x < width; x += 2) {
      const pa = pixelAt(a, x, y, bgA);
      const pb = pixelAt(b, x - dx, y - dy, bgB);
      const inkA = Math.min(255, colorDistance(pa, bgA));
      const inkB = Math.min(255, colorDistance(pb, bgB));
      if (inkA > 8 || inkB > 8) {
        difference += Math.abs(inkA - inkB);
        active++;
      }
    }
  }
  return active ? difference / active : 0;
}

function bestAlignment(
  a: RgbaImage,
  b: RgbaImage,
  bgA: readonly number[],
  bgB: readonly number[],
): { dx: number; dy: number } {
  const radius = VISUAL_THRESHOLDS.alignmentSearchPx;
  let best = { dx: 0, dy: 0, score: Number.POSITIVE_INFINITY, movement: 0 };
  for (let dy = -radius; dy <= radius; dy++) {
    for (let dx = -radius; dx <= radius; dx++) {
      const score = alignmentScore(a, b, bgA, bgB, dx, dy);
      const movement = Math.abs(dx) + Math.abs(dy);
      if (score < best.score - 1e-9 || (Math.abs(score - best.score) <= 1e-9 && movement < best.movement)) {
        best = { dx, dy, score, movement };
      }
    }
  }
  return { dx: best.dx, dy: best.dy };
}

function srgbToLinear(value: number): number {
  const v = value / 255;
  return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
}

function rgbToLab(pixel: Pixel): [number, number, number] {
  const r = srgbToLinear(pixel.r);
  const g = srgbToLinear(pixel.g);
  const b = srgbToLinear(pixel.b);
  const x = (r * 0.4124564 + g * 0.3575761 + b * 0.1804375) / 0.95047;
  const y = r * 0.2126729 + g * 0.7151522 + b * 0.0721750;
  const z = (r * 0.0193339 + g * 0.1191920 + b * 0.9503041) / 1.08883;
  const f = (v: number) => v > 216 / 24389 ? Math.cbrt(v) : (24389 / 27 * v + 16) / 116;
  const fx = f(x);
  const fy = f(y);
  const fz = f(z);
  return [116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz)];
}

function deltaE76(a: Pixel, b: Pixel): number {
  const aa = rgbToLab(a);
  const bb = rgbToLab(b);
  return Math.hypot(aa[0] - bb[0], aa[1] - bb[1], aa[2] - bb[2]);
}

function luminance(pixel: Pixel): number {
  return 0.2126 * pixel.r + 0.7152 * pixel.g + 0.0722 * pixel.b;
}

function blockSsim(a: number[], b: number[]): number {
  const n = a.length;
  if (!n) return 1;
  const meanA = a.reduce((sum, value) => sum + value, 0) / n;
  const meanB = b.reduce((sum, value) => sum + value, 0) / n;
  let varianceA = 0;
  let varianceB = 0;
  let covariance = 0;
  for (let i = 0; i < n; i++) {
    const da = a[i] - meanA;
    const db = b[i] - meanB;
    varianceA += da * da;
    varianceB += db * db;
    covariance += da * db;
  }
  const denominator = Math.max(1, n - 1);
  varianceA /= denominator;
  varianceB /= denominator;
  covariance /= denominator;
  const c1 = (0.01 * 255) ** 2;
  const c2 = (0.03 * 255) ** 2;
  return ((2 * meanA * meanB + c1) * (2 * covariance + c2)) /
    ((meanA * meanA + meanB * meanB + c1) * (varianceA + varianceB + c2));
}

function ssim(
  a: RgbaImage,
  b: RgbaImage,
  bgA: readonly number[],
  bgB: readonly number[],
  dx: number,
  dy: number,
): number {
  const width = Math.max(a.width, b.width);
  const height = Math.max(a.height, b.height);
  const block = 8;
  let total = 0;
  let blocks = 0;
  for (let top = 0; top < height; top += block) {
    for (let left = 0; left < width; left += block) {
      const valuesA: number[] = [];
      const valuesB: number[] = [];
      for (let y = top; y < Math.min(height, top + block); y++) {
        for (let x = left; x < Math.min(width, left + block); x++) {
          valuesA.push(luminance(pixelAt(a, x, y, bgA)));
          valuesB.push(luminance(pixelAt(b, x - dx, y - dy, bgB)));
        }
      }
      total += blockSsim(valuesA, valuesB);
      blocks++;
    }
  }
  return blocks ? total / blocks : 1;
}

function dilate(mask: Uint8Array, width: number, height: number, radius: number): Uint8Array {
  const out = new Uint8Array(mask.length);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      let found = false;
      for (let yy = Math.max(0, y - radius); yy <= Math.min(height - 1, y + radius) && !found; yy++) {
        for (let xx = Math.max(0, x - radius); xx <= Math.min(width - 1, x + radius); xx++) {
          if (mask[yy * width + xx]) { found = true; break; }
        }
      }
      out[y * width + x] = found ? 1 : 0;
    }
  }
  return out;
}

interface InkMetrics {
  precision: number;
  recall: number;
  f1: number;
}

function inkMetrics(
  a: RgbaImage,
  b: RgbaImage,
  bgA: readonly number[],
  bgB: readonly number[],
  dx: number,
  dy: number,
): InkMetrics {
  const width = Math.max(a.width, b.width);
  const height = Math.max(a.height, b.height);
  const maskA = new Uint8Array(width * height);
  const maskB = new Uint8Array(width * height);
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const i = y * width + x;
      maskA[i] = colorDistance(pixelAt(a, x, y, bgA), bgA) > VISUAL_THRESHOLDS.inkBackgroundDistance ? 1 : 0;
      maskB[i] = colorDistance(pixelAt(b, x - dx, y - dy, bgB), bgB) > VISUAL_THRESHOLDS.inkBackgroundDistance ? 1 : 0;
    }
  }
  const tolerantA = dilate(maskA, width, height, VISUAL_THRESHOLDS.inkToleranceRadiusPx);
  const tolerantB = dilate(maskB, width, height, VISUAL_THRESHOLDS.inkToleranceRadiusPx);
  let activeA = 0;
  let activeB = 0;
  let matchedA = 0;
  let matchedB = 0;
  for (let i = 0; i < maskA.length; i++) {
    if (maskA[i]) { activeA++; if (tolerantB[i]) matchedA++; }
    if (maskB[i]) { activeB++; if (tolerantA[i]) matchedB++; }
  }
  const precision = activeA ? matchedA / activeA : activeB ? 0 : 1;
  const recall = activeB ? matchedB / activeB : activeA ? 0 : 1;
  return {
    precision,
    recall,
    f1: precision + recall ? 2 * precision * recall / (precision + recall) : 0,
  };
}

function severity(metrics: Omit<PageMetrics, 'severity'>): VisualSeverity {
  if (Math.abs(metrics.widthDelta) > VISUAL_THRESHOLDS.geometryTolerancePx ||
      Math.abs(metrics.heightDelta) > VISUAL_THRESHOLDS.geometryTolerancePx ||
      metrics.ssim < VISUAL_THRESHOLDS.major.ssim ||
      metrics.tolerantInkF1 < VISUAL_THRESHOLDS.major.inkF1 ||
      metrics.perceptualDiffPixelRatio > VISUAL_THRESHOLDS.major.perceptualDiffRatio) return 'severe';
  if (metrics.ssim < VISUAL_THRESHOLDS.minor.ssim ||
      metrics.tolerantInkF1 < VISUAL_THRESHOLDS.minor.inkF1 ||
      metrics.perceptualDiffPixelRatio > VISUAL_THRESHOLDS.minor.perceptualDiffRatio) return 'major';
  if (metrics.ssim < VISUAL_THRESHOLDS.close.ssim ||
      metrics.tolerantInkF1 < VISUAL_THRESHOLDS.close.inkF1 ||
      metrics.perceptualDiffPixelRatio > VISUAL_THRESHOLDS.close.perceptualDiffRatio) return 'minor';
  return 'close';
}

export function compareImages(a: RgbaImage, b: RgbaImage): { metrics: PageMetrics; overlay: RgbaImage } {
  const bgA = background(a);
  const bgB = background(b);
  const alignment = bestAlignment(a, b, bgA, bgB);
  const width = Math.max(a.width, b.width);
  const height = Math.max(a.height, b.height);
  const overlay = new Uint8Array(width * height * 4);
  let exactDifferent = 0;
  let perceptuallyDifferent = 0;
  let deltaTotal = 0;

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const pa = pixelAt(a, x, y, bgA);
      const pb = pixelAt(b, x - alignment.dx, y - alignment.dy, bgB);
      if (Math.round(pa.r) !== Math.round(pb.r) ||
          Math.round(pa.g) !== Math.round(pb.g) ||
          Math.round(pa.b) !== Math.round(pb.b)) exactDifferent++;
      const delta = deltaE76(pa, pb);
      deltaTotal += delta;
      if (delta > VISUAL_THRESHOLDS.justNoticeableDeltaE) perceptuallyDifferent++;
      const target = (y * width + x) * 4;
      const gray = Math.round((luminance(pa) + luminance(pb)) / 2);
      if (delta > VISUAL_THRESHOLDS.justNoticeableDeltaE) {
        const strength = Math.min(1, delta / 35);
        overlay[target] = 255;
        overlay[target + 1] = Math.round(gray * (1 - strength) * 0.45);
        overlay[target + 2] = Math.round(gray * (1 - strength) * 0.45);
      } else {
        const faded = Math.min(255, Math.round(220 + gray * 0.14));
        overlay[target] = overlay[target + 1] = overlay[target + 2] = faded;
      }
      overlay[target + 3] = 255;
    }
  }

  const pixels = width * height;
  const ink = inkMetrics(a, b, bgA, bgB, alignment.dx, alignment.dy);
  const withoutSeverity: Omit<PageMetrics, 'severity'> = {
    docxodus: { width: a.width, height: a.height, background: bgA },
    libreoffice: { width: b.width, height: b.height, background: bgB },
    alignment: { ...alignment, searchRadius: VISUAL_THRESHOLDS.alignmentSearchPx },
    widthDelta: a.width - b.width,
    heightDelta: a.height - b.height,
    exactDiffPixelRatio: pixels ? exactDifferent / pixels : 0,
    perceptualDiffPixelRatio: pixels ? perceptuallyDifferent / pixels : 0,
    meanDeltaE76: pixels ? deltaTotal / pixels : 0,
    ssim: ssim(a, b, bgA, bgB, alignment.dx, alignment.dy),
    tolerantInkPrecision: ink.precision,
    tolerantInkRecall: ink.recall,
    tolerantInkF1: ink.f1,
  };
  return {
    metrics: { ...withoutSeverity, severity: severity(withoutSeverity) },
    overlay: { width, height, data: overlay },
  };
}

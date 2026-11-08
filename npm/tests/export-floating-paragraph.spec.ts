import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { floatingParagraphDocx } from './docx-page-fitting-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const [wrap, after, prefix] of [['none', 0, 0], ['none', 1, 0], ['none', 39, 0], ['none', 40, 0], ['none', 41, 0], ['none', 82, 0],
  ['square', 0, 0], ['tight', 0, 0], ['top-bottom', 0, 0], ['none', [0, 41, 82], 0], ['square', 0, 38], ['tight', 0, 38],
  ['square', 1, 0], ['square', 39, 0], ['square', 40, 0], ['square', 41, 0], ['square', 82, 0], ['tight', 41, 0]] as const) {
  test(`a long paragraph retains its ${wrap} picture after ${after} lines with ${prefix} prefix lines`, async ({ browser, page }, testInfo) => {
    test.setTimeout(120_000);
    const bytes = floatingParagraphDocx(wrap, after, 82, prefix);
    const positions = typeof after === 'number' ? [after] : after;
    writeFileSync(testInfo.outputPath('source.docx'), bytes);
    const result = await renderDocxArtifacts(bytes, { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, labels: string[] = [], imagePages: number[] = [], pageCounts: number[] = [], leadLabels: string[] = [];
      for (let index = 1; index <= pdf.numPages; index++) {
        const pdfPage = await pdf.getPage(index), content = await pdfPage.getTextContent();
        const text = content.items.filter(item => 'str' in item).map(item => 'str' in item ? item.str : '').join(' ');
        const onPage = text.match(/Willow\d+/g) ?? [];
        labels.push(...onPage); pageCounts.push(onPage.length); leadLabels.push(...text.match(/WillowLead\d+/g) ?? []);
        if ((wrap === 'square' || wrap === 'tight') && typeof after === 'number' && after > 0) {
          const picturePage = Math.min(3, Math.floor(after / 40) + 1);
          const lines = content.items.filter(item => 'str' in item && /Willow\d{3}/.test(item.str));
          for (const [lineIndex, line] of lines.entries()) if ('str' in line)
            expect(line.transform[4]).toBeCloseTo(index === picturePage && lineIndex < 5 ? 153 : 72, 1);
        }
        for (const op of (await pdfPage.getOperatorList()).fnArray)
          if (op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject) imagePages.push(index);
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(labels).toEqual(Array.from({ length: 82 }, (_, i) => `Willow${String(i + 1).padStart(3, '0')}`));
      expect(leadLabels).toEqual(Array.from({ length: prefix }, (_, i) => `WillowLead${i + 1}`));
      expect(imagePages).toEqual(positions.map(n => Math.min(3, Math.floor(n / 40) + 1) + (prefix ? 1 : 0)));
      // A 72pt float cannot fit below the 38-line prefix. Preserve the existing
      // body-fit retry policy; LibreOffice instead lets it overhang the bottom margin.
      if (wrap !== 'top-bottom') expect(pageCounts).toEqual(prefix ? [0, 40, 40, 2] : [40, 40, 2]);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const images = await page.locator('.page-box img').evaluateAll(images => images.map(image => {
      const box = image.closest('.page-box')!.getBoundingClientRect(), rect = image.getBoundingClientRect();
      return { x: (rect.x - box.x) * 0.75, y: (rect.y - box.y) * 0.75, width: rect.width * 0.75, height: rect.height * 0.75 };
    }));
    expect(images).toHaveLength(positions.length);
    for (const image of images) {
      expect(image.width).toBeCloseTo(72, 1); expect(image.height).toBeCloseTo(72, 1);
      if (wrap !== 'top-bottom') { expect(image.x).toBeCloseTo(wrap === 'none' ? 372 : 72, 1); expect(image.y).toBeCloseTo(72, 1); }
    }
    expect(await page.locator('.page-box').last().textContent()).toContain('Willow inventory complete.');
  });
}

test('a late picture in a short paragraph retains its full height at the page boundary', async ({ browser, page }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(floatingParagraphDocx('square', 1, 2, 38), {
    browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
  try {
    const pdf = await task.promise;
    expect(pdf.numPages).toBe(2);
    const text: string[] = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const content = await (await pdf.getPage(i)).getTextContent();
      text.push(content.items.map(item => 'str' in item ? item.str : '').join(' '));
    }
    expect(text[0].match(/WillowLead\d+/g)).toHaveLength(38);
    // The first source line fits before the anchor; only the picture-owning tail moves.
    expect(text[0].match(/Willow\d+/g)).toEqual(['Willow001']);
    expect(text[1].match(/Willow\d+/g)).toEqual(['Willow002']);
    expect(text[1]).toContain('Willow inventory complete.');
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const image = await page.locator('.page-box').nth(1).locator('img').evaluate(image => {
    const box = image.closest('.page-box')!.getBoundingClientRect(), rect = image.getBoundingClientRect();
    return { x: (rect.x - box.x) * 0.75, y: (rect.y - box.y) * 0.75, height: rect.height * 0.75 };
  });
  expect(image.x).toBeCloseTo(72, 1); expect(image.y).toBeCloseTo(72, 1); expect(image.height).toBeCloseTo(72, 1);
});

for (const after of [1, 41]) {
  test(`a late picture after ${after} lines cannot overlap a preceding paragraph's float`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(floatingParagraphDocx('square', after, 82, 1, true), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise, labels: string[] = [], imagePages: number[] = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i), content = await page.getTextContent();
        labels.push(...content.items.flatMap(item => 'str' in item ? item.str.match(/Willow\d+/g) ?? [] : []));
        for (const op of (await page.getOperatorList()).fnArray)
          if (op === OPS.paintImageXObject || op === OPS.paintInlineImageXObject) imagePages.push(i);
      }
      expect(labels).toEqual(Array.from({ length: 82 }, (_, i) => `Willow${String(i + 1).padStart(3, '0')}`));
      expect(imagePages).toEqual([1, 2]);
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const offsets = await page.locator('[data-docx-fragment-edge]').evaluateAll(drawings => drawings.map(drawing => {
      const proxy = drawing.closest('p')!.querySelector('[data-docx-edge-exclusion]')!;
      const ink = drawing.getBoundingClientRect(), exclusion = proxy.getBoundingClientRect();
      return [ink.left - exclusion.left, ink.top - exclusion.top];
    }));
    expect(offsets).toHaveLength(1);
    for (const delta of offsets[0]) expect(Math.abs(delta)).toBeLessThan(0.1);
  });
}

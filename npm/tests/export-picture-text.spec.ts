import { expect, test } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { pictureDocx } from './docx-picture-fixture.js';

const font = readFileSync(new URL('fixtures/fonts/baseline-test-b.woff2', import.meta.url));
const digest = (bytes: Uint8Array | string) => createHash('sha256').update(bytes).digest('hex');
const texts = [Array.from({ length: 120 }, (_, i) => `token${i}`).join(' '),
  'Following paragraph preserves its original position and source identity.'];

for (const tight of [false, true]) {
  test(`precision text preserves ${tight ? 'tight' : 'square'} floating pictures through public export and PDF`, async ({ page, context }, testInfo) => {
    const source = pictureDocx(tight, true);
    await page.goto('/standalone-export-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
    const result = await page.evaluate(async ({ bytes, plan }) =>
      (window as any).DocxodusStandalone.convertWithFontResolver(
        bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, plan),
    { bytes: Array.from(source), plan: {
      mode: 'exact', format: 'woff2', mediaType: 'font/woff2',
      byteLength: font.byteLength, sha256: digest(font), bytesBase64: font.toString('base64'),
      licenseIdentity: 'b'.repeat(64),
    } });
    expect(result.renderReport.status).toBe('complete');
    expect(result.pageCount).toBe(1);
    expect(result.renderReport.source.rawPackageBytesDigest).toBe(digest(source));
    expect(result.renderReport.bindings.htmlDigest).toBe(digest(result.html));
    expect(result.renderReport.fonts.length).toBeGreaterThan(0);
    expect(result.renderReport.fonts.every((face: any) =>
      face.source === 'configured' && face.fileSha256 === digest(font))).toBe(true);

    const htmlPath = testInfo.outputPath('picture-text.html');
    writeFileSync(htmlPath, result.html);
    const offline = await context.newPage();
    await offline.goto(pathToFileURL(htmlPath).href);
    await offline.evaluate(() => document.fonts.ready);
    const reopened = await offline.evaluate(() => {
      const content = document.querySelector<HTMLElement>('.page-content')!;
      const pageBox = content.closest('.page-box')!.getBoundingClientRect();
      const image = content.querySelector('img')!;
      const box = image.getBoundingClientRect(); const band = content.getBoundingClientRect();
      const paragraphs = Array.from(content.querySelectorAll<HTMLElement>('p[data-source-anchor-id]'));
      return { imageCount: content.querySelectorAll('img').length,
        image: { x: (box.left - band.left) * 0.75, y: (box.top - band.top) * 0.75,
          width: box.width * 0.75, height: box.height * 0.75 },
        rows: Array.from(paragraphs[0].querySelectorAll('svg > text')).map(line =>
          Number(line.firstElementChild!.getAttribute('x')) / 4 * 0.75),
        paragraphs: paragraphs.map(p => {
          const rect = p.getBoundingClientRect();
          return { text: p.textContent, anchorId: p.dataset.sourceAnchorId,
            painted: p.querySelectorAll('svg').length,
            geometry: { x: (rect.left - pageBox.left) * 0.75, y: (rect.top - pageBox.top) * 0.75,
              width: rect.width * 0.75, height: rect.height * 0.75 } };
        }) };
    });
    expect(reopened.imageCount).toBe(1);
    expect(reopened.image.x).toBeCloseTo(0, 1);
    expect(reopened.image.y).toBeCloseTo(0, 1);
    expect(reopened.image.width).toBeCloseTo(72, 1);
    expect(reopened.image.height).toBeCloseTo(36, 1);
    expect(reopened.paragraphs.map(p => p.text?.trim())).toEqual(texts);
    expect(reopened.paragraphs.every(p => p.painted === 1)).toBe(true);
    expect(reopened.rows.length).toBeGreaterThan(3);
    if (tight) expect(reopened.rows[1]).toBeLessThan(reopened.rows[0] - 5);
    else expect(reopened.rows[1]).toBeCloseTo(reopened.rows[0], 1);
    expect(reopened.rows.at(-1)).toBeCloseTo(0, 1);
    for (const p of reopened.paragraphs) {
      const fragment = result.pageMap.fragments.find((f: any) => f.anchorId === p.anchorId);
      expect(fragment).toBeDefined();
      expect(fragment.pageNumber).toBe(1);
      for (const dimension of ['x', 'y', 'width', 'height'] as const)
        expect(Math.abs(fragment.geometry[dimension] - p.geometry[dimension])).toBeLessThan(0.1);
    }

    const pdf = await offline.pdf({ preferCSSPageSize: true, printBackground: true });
    writeFileSync(testInfo.outputPath('picture-text.pdf'), pdf);
    const task = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
    const printed = await task.promise;
    expect(printed.numPages).toBe(1);
    let inventory = '';
    for (const item of (await (await printed.getPage(1)).getTextContent()).items)
      if ('str' in item) inventory += item.str + (item.hasEOL ? '\n' : '');
    expect(inventory.split(/\s+/).filter(Boolean)).toEqual(texts.join(' ').split(/\s+/));
    await task.destroy();
    await offline.close();
  });
}

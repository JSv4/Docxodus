import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { bodyFillDocx } from './docx-body-fill-fixture.js';
import { outOfFlowDocx } from './docx-out-of-flow-fixture.js';

const here = dirname(fileURLToPath(import.meta.url));
const directory = join(here, 'fixtures', 'exact-line-spacing');
const word = JSON.parse(readFileSync(join(directory, 'word.json'), 'utf8')) as {
  topMarginPt: number;
  cases: Array<{
    fixture: string; sha256: string; fontSizePt: number; lineHeightPt: number;
    pageCount: number; baselinesPt: number[];
  }>;
  body: { fixture: string; sha256: string; pageCount: number; lineCounts: number[] };
};
const font = readFileSync(join(here, 'fixtures', 'fonts', 'baseline-test-c.woff2'));
const fontPlan = {
  mode: 'exact', format: 'woff2', mediaType: 'font/woff2', byteLength: font.length,
  sha256: createHash('sha256').update(font).digest('hex'), bytesBase64: font.toString('base64'),
  licenseIdentity: 'c'.repeat(64),
};

async function exportFixture(page: Page, fixture: string | Uint8Array) {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(({ bytes, plan }) => (window as any).DocxodusStandalone
    .convertWithFontResolver(bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, plan),
  { bytes: Array.from(typeof fixture === 'string' ? readFileSync(join(directory, fixture)) : fixture), plan: fontPlan });
  expect(result.renderReport.status).toBe('complete');
  await page.setContent(result.html);
  await page.evaluate(() => document.fonts.ready);
  return result;
}

test('the exact-spacing fixtures match the documents measured in Word', () => {
  for (const record of [...word.cases, word.body])
    expect(createHash('sha256').update(readFileSync(join(directory, record.fixture))).digest('hex'))
      .toBe(record.sha256);
  for (const entry of word.cases)
    expect(Buffer.from(bodyFillDocx({
      paragraphs: 3, fontHalfPoints: entry.fontSizePt * 2, lineTwips: entry.lineHeightPt * 20,
    }))).toEqual(readFileSync(join(directory, entry.fixture)));
});

for (const entry of word.cases) {
  test(`exact spacing: ${entry.fontSizePt}pt text on ${entry.lineHeightPt}pt lines follows Word`, async ({ page }) => {
    const result = await exportFixture(page, entry.fixture);
    expect(result.pageCount).toBe(entry.pageCount);
    const geometry = await page.evaluate(() => {
      const box = document.querySelector('.page-box')!;
      return Array.from(box.querySelectorAll('.page-content p')).map((paragraph) => {
        const run = paragraph.querySelector('span')!;
        const marker = document.createElement('span');
        marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0';
        run.prepend(marker);
        const baseline = (marker.getBoundingClientRect().bottom - box.getBoundingClientRect().top) * 0.75;
        marker.remove();
        return { baseline, height: paragraph.getBoundingClientRect().height * 0.75 };
      });
    });
    expect(geometry).toHaveLength(entry.baselinesPt.length);
    geometry.forEach((line, index) => {
      // Word's text origins are quantized; 0.2pt separates that grid from CSS layout precision.
      expect(Math.abs(line.baseline - entry.baselinesPt[index]), `line ${index + 1} baseline`)
        .toBeLessThanOrEqual(0.2);
      expect(Math.abs(line.height - entry.lineHeightPt)).toBeLessThanOrEqual(0.02);
    });
  });
}

test('exact spacing keeps all 54 lines and their descenders on the first page', async ({ page }) => {
  await page.setViewportSize({ width: 1200, height: 1400 });
  const result = await exportFixture(page, word.body.fixture);
  expect(result.pageCount).toBe(word.body.pageCount);
  const pages = await page.evaluate(() => Array.from(document.querySelectorAll('.page-box')).map((box) => {
    const body = box.querySelector<HTMLElement>('.page-content')!;
    const lines = Array.from(body.querySelectorAll('p'));
    const range = document.createRange();
    range.selectNodeContents(lines.at(-1)!);
    const ink = Array.from(range.getClientRects()).sort((a, b) => b.bottom - a.bottom)[0];
    const visible = document.elementFromPoint(ink.left + 2, ink.bottom - 0.25)?.closest('p') === lines.at(-1);
    return { count: lines.length, last: lines.at(-1)!.textContent,
      visible };
  }));
  expect(pages.map(page => page.count)).toEqual(word.body.lineCounts);
  expect(pages[0].last).toBe('Line 54 of the body, typography and spacing.');
  expect(pages[0].visible, 'the last glyph box must remain visible beyond the body line box').toBe(true);
});

test('exact spacing fragments a long paragraph without losing text', async ({ page }) => {
  const text = 'Line typography and spacing. '.repeat(180).trim();
  const source = outOfFlowDocx(`<w:p><w:pPr><w:widowControl w:val="0"/>
    <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="32"/></w:rPr>
      <w:t xml:space="preserve">${text}</w:t></w:r></w:p>`);
  const result = await exportFixture(page, source);
  expect(result.pageCount).toBeGreaterThan(1);
  const actual = await page.locator('.page-content p').allTextContents();
  expect(actual.join('').replace(/\s+/g, ' ').trim()).toBe(text);
});

test('exact spacing aligns mixed font sizes inside a hyperlink', async ({ page }) => {
  const run = (size: number) => `<w:r><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/>
    <w:sz w:val="${size * 2}"/></w:rPr><w:t xml:space="preserve">Line ${size} </w:t></w:r>`;
  const source = outOfFlowDocx(`<w:p><w:pPr>
    <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>
    ${run(10)}<w:hyperlink w:anchor="target">${run(16)}</w:hyperlink>${run(24)}</w:p>`);
  await exportFixture(page, source);
  const baselines = await page.evaluate(() => {
    const top = document.querySelector('.page-box')!.getBoundingClientRect().top;
    return Array.from(document.querySelectorAll('.page-content [data-docx-exact-run]')).map(run => {
      const marker = document.createElement('span');
      marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline';
      run.prepend(marker);
      const y = (marker.getBoundingClientRect().bottom - top) * 0.75;
      marker.remove();
      return y;
    });
  });
  expect(baselines).toHaveLength(3);
  for (const baseline of baselines) expect(Math.abs(baseline - 81.6)).toBeLessThan(0.02);
});

for (const scale of [0.5, 1, 2]) {
  test(`exact ink at scale ${scale} cannot exempt an overflowing block`, async ({ page }) => {
    await page.goto('/standalone-export-harness.html');
    const result = await page.evaluate(async (zoom) => {
      const moduleUrl = 'http://localhost:8083/line-metrics.js';
      const metrics = await import(moduleUrl);
      document.body.innerHTML = `<div id="paper" style="position:relative;width:400px;height:200px;zoom:${zoom}">
        <div id="content" style="position:absolute;top:40px;left:20px;width:360px;height:16px;overflow-x:visible;overflow-y:clip">
          <p style="margin:0;font:24pt 'Times New Roman';line-height:16px;--docx-exact-line-height:12pt">
            <span id="ink" data-docx-exact-run="true" style="vertical-align:top">gyp</span>
          </p><div id="block" style="height:5px">block</div>
        </div></div>`;
      const paper = document.getElementById('paper')!;
      const content = document.getElementById('content')!;
      const ink = document.getElementById('ink')!;
      const block = document.getElementById('block')!;
      metrics.alignExactLineBaselines(content);
      const bounds = paper.getBoundingClientRect();
      metrics.preserveExactLineInk(content, bounds.top, bounds.bottom, bounds);
      const rect = ink.getBoundingClientRect();
      return {
        inkAllowed: metrics.isPreservedExactLineInk(ink, content),
        blockAllowed: metrics.isPreservedExactLineInk(block, content),
        visible: document.elementFromPoint(rect.left + 1, rect.bottom - 0.1)?.id === 'ink',
        height: content.getBoundingClientRect().height / zoom,
      };
    }, scale);
    expect(result).toEqual({ inkAllowed: true, blockAllowed: false, visible: true, height: 16 });
  });
}

test('exact ink cannot extend into a running story', async ({ page }) => {
  await page.goto('/standalone-export-harness.html');
  const allowed = await page.evaluate(async () => {
    const moduleUrl = 'http://localhost:8083/line-metrics.js';
    const metrics = await import(moduleUrl);
    document.body.innerHTML = `<div id="paper" style="position:relative;width:400px;height:200px">
      <div id="content" style="position:absolute;top:40px;left:20px;width:360px;height:16px;overflow-y:clip">
        <p style="margin:0;font:24pt 'Times New Roman';line-height:16px;--docx-exact-line-height:12pt">
          <span id="ink" data-docx-exact-run="true" style="vertical-align:top">gyp</span>
        </p></div></div>`;
    const content = document.getElementById('content')!;
    metrics.alignExactLineBaselines(content);
    const paper = document.getElementById('paper')!.getBoundingClientRect();
    metrics.preserveExactLineInk(content, paper.top, content.getBoundingClientRect().bottom, paper);
    return metrics.isPreservedExactLineInk(document.getElementById('ink')!, content);
  });
  expect(allowed).toBe(false);
});

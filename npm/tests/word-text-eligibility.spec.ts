import { expect, test, type Page } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const module = readFileSync(join(here, '../dist/word-text.js'), 'utf8');

// Issue #1123: single-spaced documents with running stories or right-to-left paragraphs reach the
// precision text path. Lines are already resolved to lengths here, as the export's unrounded
// normal-line pass leaves them before `prepareWordText` runs.
const plain = (text: string, extra = '') =>
  `<p style="--docx-auto-line-spacing: 1; line-height: 16.8667px; margin: 0"${extra}><span style="font: 11pt 'Liberation Serif', serif; line-height: 16.8667px">${text}</span></p>`;

async function prepare(page: Page, body: string): Promise<{ prepared: number; marked: string[] }> {
  await page.setContent(`<div id="root">${body}</div>`);
  await page.addScriptTag({ type: 'module', content: `${module}\nwindow.wordText = { prepareWordText, paintWordText };` });
  await page.waitForFunction(() => (window as any).wordText);
  return page.evaluate(() => {
    const root = document.getElementById('root')!;
    const prepared = (window as any).wordText.prepareWordText(root);
    return { prepared, marked: Array.from(root.querySelectorAll('[data-docx-word-text-prepared]')).map((p) => p.textContent!) };
  });
}

test('running-story paragraphs no longer exclude a single-spaced document', async ({ page }) => {
  const result = await prepare(page, `
    <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="header-default">${plain('MarginHead Chestnut running header.')}</div></div>
    ${plain('MarginTop Chestnut margin records begin here.')}`);
  expect(result.marked).toEqual(['MarginHead Chestnut running header.', 'MarginTop Chestnut margin records begin here.']);
});

test('a plain right-to-left paragraph keeps HTML layout without excluding its neighbours', async ({ page }) => {
  const result = await prepare(page, `
    ${plain('InsetMarker Elm archive entries.')}
    ${plain('LogicalMarker Elm archive entries', ' dir="rtl"')}`);
  expect(result.marked).toEqual(['InsetMarker Elm archive entries.']);
});

test('a right-to-left paragraph with rich inline content still keeps the whole document on HTML layout', async ({ page }) => {
  const result = await prepare(page, `
    ${plain('InsetMarker Elm archive entries.')}
    <p dir="rtl" style="--docx-auto-line-spacing: 1; margin: 0"><span>Elm</span><sup>2</sup></p>`);
  expect(result.prepared).toBe(0);
});

test('painted words keep a fractional paragraph origin the SVG origin snaps away', async ({ page }) => {
  // The paragraph starts 0.32px into a pixel; Chromium paints its SVG from the whole pixel.
  await prepare(page, `<div class="page-box" style="position: relative">
    <div style="padding-left: 100.32px">${plain('BoxMarker Elm archive entries.')}</div></div>`);
  const x = await page.evaluate(() => {
    const root = document.getElementById('root')!;
    (window as any).wordText.paintWordText(root);
    return Number(root.querySelector('tspan')!.getAttribute('x')) / 4;
  });
  // Layout positions are multiples of 1/64 CSS pixel.
  expect(Math.abs(x - 0.32)).toBeLessThanOrEqual(1 / 64);
});

test('later paragraphs keep the twip grid despite layout-unit truncation', async ({ page }) => {
  // 14.5pt lines are 19.333px; layout keeps 19.328125px, so the second paragraph's layout top is
  // early by 0.0052px per line above it. Its lines must still sit whole ideal pitches below the first's.
  const line = '19.333333px';
  const para = (text: string) => `<p style="--docx-auto-line-spacing: 1; line-height: ${line}; margin: 0; width: 120px"><span style="font: 11pt 'Liberation Serif', serif; line-height: ${line}">${text}</span></p>`;
  await prepare(page, `<div class="page-box" style="position: relative">
    ${para('Clause01 parties inspect the sample record')}${para('Clause02 parties inspect the sample record')}</div>`);
  const result = await page.evaluate(() => {
    const root = document.getElementById('root')!;
    const pageTop = root.querySelector('.page-box')!.getBoundingClientRect().top;
    const paragraphs = Array.from(root.querySelectorAll('p'));
    const tops = paragraphs.map((p) => p.getBoundingClientRect().top - pageTop);
    (window as any).wordText.paintWordText(root);
    const baselines = paragraphs.map((p, i) =>
      Math.round(tops[i]) + Number(p.querySelector('text')!.getAttribute('y')) / 4);
    const lines = paragraphs.map((p) => p.querySelectorAll('text').length);
    return { baselines, lines };
  });
  expect(result.lines[0]).toBeGreaterThan(1);
  expect(Math.abs(result.baselines[1] - result.baselines[0] - result.lines[0] * 19.333333)).toBeLessThan(1e-3);
});

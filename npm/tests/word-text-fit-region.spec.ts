import { expect, test } from '@playwright/test';
import { wordTextScript } from './word-text-script.js';

for (const width of [20, 100]) test(`manual ${width}pt group preserves its baseline, runs and ordinary tail`, async ({ page }) => {
  await page.setContent(`<style>p{font:12pt/18pt 'Liberation Serif';margin:0}span{white-space:pre-wrap}</style>
    <main><p><span data-docx-fit-text-width="${width}" data-docx-fit-text-id="0" data-docx-fit-text-count="2">Na</span><span
    data-docx-fit-text-width="${width}" data-docx-fit-text-id="0" data-docx-fit-text-count="2" style="font-weight:bold">me</span><span> tail</span></p></main>`);
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(async () => {
    await document.fonts.ready;
    const root = document.querySelector<HTMLElement>('main')!, p = root.querySelector('p')!;
    const original = Array.from(p.children), height = p.getBoundingClientRect().height;
    const tailY = original[2].getBoundingClientRect().y;
    const api = window as any, count = api.prepareFitText(root), once = p.outerHTML;
    api.prepareFitText(root);
    const group = p.firstElementChild!, inner = group.firstElementChild!;
    return { count, unchanged:once === p.outerHTML, text:p.textContent,
      sameRuns:original.every(e => p.contains(e)), height, afterHeight:p.getBoundingClientRect().height,
      tailY, afterTailY:original[2].getBoundingClientRect().y,
      width:inner.getBoundingClientRect().width * .75, groupWidth:group.getBoundingClientRect().width * .75,
      bold:getComputedStyle(original[1]).fontWeight };
  });
  expect(result.count).toBe(1); expect(result.unchanged).toBe(true); expect(result.sameRuns).toBe(true);
  expect(result.text).toBe('Name tail'); expect(result.bold).toBe('700');
  expect(result.width).toBeCloseTo(width, 2); expect(result.groupWidth).toBeCloseTo(width, 1);
  expect(result.afterHeight).toBe(result.height); expect(result.afterTailY).toBe(result.tailY);
});

test('an incomplete source group is not fitted as separate runs', async ({ page }) => {
  await page.setContent('<main><p><span data-docx-fit-text-width="100" data-docx-fit-text-id="0" data-docx-fit-text-count="2">Na</span>'
    + '<a><span data-docx-fit-text-width="100" data-docx-fit-text-id="0" data-docx-fit-text-count="2">me</span></a></p></main>');
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('main')!, before = root.innerHTML;
    return { count:(window as any).prepareFitText(root), before, after:root.innerHTML };
  });
  expect(result.count).toBe(0); expect(result.after).toBe(result.before);
});

test('a dynamic field keeps its original result nodes for later pagination', async ({ page }) => {
  await page.setContent('<main><span data-field="PAGE"><span data-docx-fit-text-width="100" '
    + 'data-docx-fit-text-id="0" data-docx-fit-text-count="1">1</span></span></main>');
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('main')!, before = root.innerHTML;
    return { count:(window as any).prepareFitText(root), before, after:root.innerHTML };
  });
  expect(result.count).toBe(0); expect(result.after).toBe(result.before);
});

test('a long incomplete group is visited once, including its separately wrapped tail', async ({ page }) => {
  const run = '<span data-docx-fit-text-width="100" data-docx-fit-text-id="0" data-docx-fit-text-count="401">N</span>';
  await page.setContent(`<main><p>${run.repeat(400)}<a>${run}</a></p></main>`);
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('main')!, before = root.innerHTML;
    const descriptor = Object.getOwnPropertyDescriptor(Node.prototype, 'nextSibling')!;
    let visits = 0;
    Object.defineProperty(Node.prototype, 'nextSibling', { configurable:true, get() { visits++; return descriptor.get!.call(this); } });
    try { return { count:(window as any).prepareFitText(root), before, after:root.innerHTML, visits }; }
    finally { Object.defineProperty(Node.prototype, 'nextSibling', descriptor); }
  });
  expect(result.count).toBe(0); expect(result.after).toBe(result.before); expect(result.visits).toBeLessThan(1000);
});

for (const transform of ['scaleX(2)', 'rotate(20deg)']) test(`a ${transform} owner keeps the local manual width`, async ({ page }) => {
  await page.setContent(`<main style="transform:${transform};transform-origin:left top"><p style="font:12pt/18pt 'Liberation Serif'">
    <span data-docx-fit-text-width="100" data-docx-fit-text-id="0" data-docx-fit-text-count="1">Name</span><span> tail</span></p></main>`);
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(async () => {
    await document.fonts.ready;
    const root = document.querySelector<HTMLElement>('main')!, count = (window as any).prepareFitText(root);
    const outer = root.querySelector<HTMLElement>('[data-docx-fit-text-group]')!, inner = outer.firstElementChild as HTMLElement;
    return { count, width:Number.parseFloat(getComputedStyle(outer).width)*.75,
      outer:outer.getBoundingClientRect().width, inner:inner.getBoundingClientRect().width };
  });
  expect(result.count).toBe(1); expect(result.width).toBeCloseTo(100, 1);
  expect(result.inner).toBeCloseTo(result.outer, 1);
});

for (const style of ['padding-left:20px', 'margin-left:20px', 'border-left:5px solid', 'display:none'])
test(`${style} leaves the complete fit group unchanged`, async ({ page }) => {
  await page.setContent(`<main><p><span style="${style}" data-docx-fit-text-width="100" data-docx-fit-text-id="0" data-docx-fit-text-count="1">Name</span><span> tail</span></p></main>`);
  await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const root = document.querySelector<HTMLElement>('main')!, before = root.innerHTML;
    return { count:(window as any).prepareFitText(root), before, after:root.innerHTML };
  });
  expect(result.count).toBe(0); expect(result.after).toBe(result.before);
});

import { test, expect, type Page } from '@playwright/test';
import { build } from 'esbuild';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// Laying a document out again keeps the scroll offsets around it (issue #1863). Replacing the pages emptied
// the element for one layout, and the browser clamped the scroll offsets around it to that size.
const npmRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
// The staging and page-container rules the converter writes in paginated mode.
const converterCss = `<style>.page-staging{position:absolute;left:-9999px;visibility:hidden}
.page-container{display:flex;flex-direction:column;align-items:center;gap:20px;padding:20px;min-height:100vh}</style>`;
const documentOf = (word: string, paragraphs = 120) => `${converterCss}
<div id="pagination-staging" class="page-staging">
  <div data-section-index="0" data-page-width="612" data-page-height="792"
       data-content-width="468" data-content-height="648" data-margin-top="72"
       data-margin-right="72" data-margin-bottom="72" data-margin-left="72">${
  Array.from({ length: paragraphs }, (_, i) =>
    `<p style="margin:0 0 12pt 0">${word} ${i} ${'lorem ipsum dolor sit amet '.repeat(6)}</p>`).join('')
}</div></div><div id="pagination-container" class="page-container"></div>`;
const first = documentOf('Alder'), second = documentOf('Birch'), short = documentOf('Cedar', 3);

let bundle = '';
test.beforeAll(async () => {
  const output = await build({
    stdin: { resolveDir: npmRoot, loader: 'js', contents: `
      import React from 'react';
      import { createRoot } from 'react-dom/client';
      import { flushSync } from 'react-dom';
      import { usePagination, PaginatedDocument } from './src/react.ts';
      import { paginateHtml } from './src/pagination.ts';
      function Viewer({ html, scale }) {
        const ref = React.useRef(null);
        window.paginationError = usePagination(html, ref, { scale }).error?.message;
        return React.createElement('div', { id: 'scroller', style: { height: 900, width: 500, overflow: 'auto' } },
          React.createElement('div', { ref, id: 'target' }));
      }
      let root;
      const render = element => {
        root ??= createRoot(document.getElementById('root'));
        flushSync(() => root.render(element));
      };
      window.hook = (html, scale = 1) => render(React.createElement(Viewer, { html, scale }));
      window.component = html => render(React.createElement(PaginatedDocument, { html, style: { height: 600, minHeight: 0 } }));
      window.paginateHtml = paginateHtml;
    ` },
    bundle: true, format: 'iife', write: false,
    define: { 'process.env.NODE_ENV': '"development"' },
    plugins: [{ name: 'unused-core', setup(b) {
      b.onResolve({ filter: /^\.\/index\.js$/ }, args => args.importer.endsWith('/src/react.ts') ?
        { path: 'unused-core', namespace: 'unused-core' } : undefined);
      b.onLoad({ filter: /.*/, namespace: 'unused-core' }, () => ({ loader: 'js', contents:
        'export const ' + ['initialize', 'convertDocxToHtml', 'compareDocuments', 'compareDocumentsToHtml',
          'getRevisions', 'getAnnotations', 'addAnnotation', 'addAnnotationWithTarget', 'removeAnnotation',
          'hasAnnotations', 'getDocumentStructure', 'getDocumentMetadata', 'isInitialized']
          .map(name => `${name}=()=>{throw new Error("Unused engine call")}`).join(',') + ';' }));
    } }],
  });
  bundle = output.outputFiles[0].text;
});

async function open(page: Page) {
  await page.setViewportSize({ width: 1000, height: 700 });
  // A scroll container taller than the window, so the page container's 100vh does not keep its range.
  await page.setContent(`<!doctype html><div id="root"></div>
    <div id="plain" style="height:900px;overflow:auto"><div id="host"></div></div>
    <div id="self" style="height:600px;overflow:auto"></div>`);
  await page.addScriptTag({ content: bundle });
}
const firstPageText = (page: Page) => page.evaluate(() => document.querySelector('.page-box')?.textContent ?? '');
const offsets = (page: Page, selector: string) => page.evaluate(s => {
  const box = document.querySelector(s)!;
  return [box.scrollTop, box.scrollLeft];
}, selector);
const scrollTo = (page: Page, selector: string, top: number, left = 0) => page.evaluate(([s, y, x]) => {
  const box = document.querySelector(s as string)!;
  box.scrollTop = y as number;
  box.scrollLeft = x as number;
}, [selector, top, left] as const);

test('the hook keeps a scroll container around it in place, down and across, when the HTML or the scale changes', async ({ page }) => {
  await open(page);
  await page.evaluate(html => (window as any).hook(html), first);
  await expect.poll(() => firstPageText(page)).toContain('Alder');
  await scrollTo(page, '#scroller', 1500, 150); // the 816px pages are wider than the 500px scroller

  await page.evaluate(html => (window as any).hook(html), second);
  await expect.poll(() => firstPageText(page)).toContain('Birch');
  expect(await offsets(page, '#scroller')).toEqual([1500, 150]);

  await page.evaluate(html => (window as any).hook(html, 0.8), second);
  // The page box keeps its 612pt width and scales by zoom or transform, so wait for the drawn width.
  await expect.poll(() => page.evaluate(() => Math.round(document.querySelector('.page-box')!.getBoundingClientRect().width)))
    .toBe(Math.round(816 * 0.8));
  // The narrower pages stop the horizontal offset at their edge; the vertical one is kept.
  expect(await page.evaluate(() => {
    const box = document.getElementById('scroller')!;
    return [box.scrollTop, box.scrollLeft === Math.min(150, box.scrollWidth - box.clientWidth)];
  })).toEqual([1500, true]);
});

test('paginateHtml keeps a scroll container around it in place without a scroll event', async ({ page }) => {
  await open(page);
  const result = await page.evaluate(async ([a, b]) => {
    const host = document.getElementById('host')!, scroller = document.getElementById('plain')!;
    (window as any).paginateHtml(a, host);
    scroller.scrollTop = 1500;
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    let events = 0;
    scroller.addEventListener('scroll', () => events++);
    (window as any).paginateHtml(b, host);
    const top = scroller.scrollTop;
    // A clamp would dispatch its scroll event with the next frame.
    await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));
    return { top, events };
  }, [first, second]);
  expect(result).toEqual({ top: 1500, events: 0 });
});

test('an element that scrolls itself keeps its own scroll position', async ({ page }) => {
  await open(page);
  // PaginatedDocument with a fixed height is its own scroll container.
  await page.evaluate(html => (window as any).component(html), first);
  await expect.poll(() => firstPageText(page)).toContain('Alder');
  await scrollTo(page, '#root > div', 1500);
  await page.evaluate(html => (window as any).component(html), second);
  await expect.poll(() => firstPageText(page)).toContain('Birch');
  expect(await offsets(page, '#root > div')).toEqual([1500, 0]);

  expect(await page.evaluate(([a, b]) => {
    const self = document.getElementById('self')!;
    (window as any).paginateHtml(a, self);
    self.scrollTop = 1500;
    (window as any).paginateHtml(b, self);
    return self.scrollTop;
  }, [first, second])).toBe(1500);
});

test('a shorter document stops the scroll at its end, and a failure is still reported', async ({ page }) => {
  await open(page);
  const facts = await page.evaluate(([long, brief]) => {
    const host = document.getElementById('host')!, scroller = document.getElementById('plain')!;
    (window as any).paginateHtml(long, host);
    scroller.scrollTop = 1500;
    (window as any).paginateHtml(brief, host);
    const clamped = { top: scroller.scrollTop, max: scroller.scrollHeight - scroller.clientHeight };
    (window as any).paginateHtml(long, host);
    let error = '';
    try { (window as any).paginateHtml('<p>Not paginated</p>', host); } catch (e) { error = (e as Error).message; }
    return { clamped, error };
  }, [first, short]);
  expect(facts.clamped.top).toBe(Math.max(0, facts.clamped.max));
  expect(facts.clamped.top).toBeLessThan(1500);
  expect(facts.error).toContain('staging element not found');
});

test('offsets come back past smooth scrolling and shadow roots, and an unclamped one is not touched', async ({ page }) => {
  await open(page);
  const facts = await page.evaluate(([a, b]) => {
    const paginate = (window as any).paginateHtml;
    // A container that scrolls smoothly would animate a plain scrollTop write, and lose it to the next one.
    const smooth = document.getElementById('plain')!;
    smooth.style.scrollBehavior = 'smooth';
    paginate(a, document.getElementById('host'));
    smooth.scrollTo({ top: 1500, behavior: 'instant' });
    paginate(b, document.getElementById('host'));
    // The element in a shadow root, the scroll container outside it.
    const outer = document.body.appendChild(document.createElement('div'));
    outer.style.cssText = 'height:900px;overflow:auto';
    const inner = outer.appendChild(document.createElement('div')).attachShadow({ mode: 'open' })
      .appendChild(document.createElement('div'));
    paginate(a, inner);
    outer.scrollTop = 1500;
    paginate(b, inner);
    // In a layout that clamps nothing, the restore writes nothing.
    const quiet = document.getElementById('self')!;
    paginate(a, quiet.appendChild(document.createElement('div')));
    quiet.scrollTop = 900;
    let writes = 0;
    const scrollTo = Element.prototype.scrollTo;
    Element.prototype.scrollTo = function (this: Element, ...args: any[]) { writes++; return (scrollTo as any).apply(this, args); };
    paginate(b, quiet.firstElementChild);
    Element.prototype.scrollTo = scrollTo;
    return { smooth: smooth.scrollTop, shadow: outer.scrollTop, quiet: quiet.scrollTop, writes };
  }, [first, second]);
  expect(facts).toEqual({ smooth: 1500, shadow: 1500, quiet: 900, writes: 0 });
});

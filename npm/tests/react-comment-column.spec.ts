import { test, expect } from '@playwright/test';
import { build } from 'esbuild';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const npmRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const html = `<div id="pagination-staging">
  <aside id="pagination-comment-margin-registry" style="display:none">
    <div data-comment-id="0"><p>${Array.from({ length: 240 }, (_, i) => `CedarNote${i}`).join(' ')}</p></div>
  </aside>
  <div data-section-index="0" data-page-width="612" data-page-height="792"
       data-content-width="468" data-content-height="648" data-margin-top="72"
       data-margin-right="72" data-margin-bottom="72" data-margin-left="72">
    <p><span data-comment-id="0">Cedar inventory</span></p>
  </div></div><div id="pagination-container"></div>`;

for (const component of [false, true]) test(`React ${component ? 'component' : 'hook'} forwards comment placement and responds to option changes`, async ({ page }) => {
  const bundle = await build({
    stdin: { resolveDir: npmRoot, loader: 'js', contents: `
      import React from 'react';
      import { createRoot } from 'react-dom/client';
      import { flushSync } from 'react-dom';
      import { usePagination, PaginatedDocument } from './src/react.ts';
      const html = ${JSON.stringify(html)};
      function Probe({ options }) {
        const ref = React.useRef(null);
        const value = usePagination(html, ref, options);
        window.paginationError = value.error?.message;
        return React.createElement('div', { ref });
      }
      const root = createRoot(document.getElementById('root'));
      window.render = options => flushSync(() => root.render(${component ?
        'React.createElement(PaginatedDocument, { html, ...options })' : 'React.createElement(Probe, { options })'}));
    ` }, bundle: true, format: 'iife', write: false,
    define: { 'process.env.NODE_ENV': '"development"' },
    plugins: [{ name: 'unused-core', setup(b) {
      b.onResolve({ filter: /^\.\/index\.js$/ }, args => args.importer.endsWith('/src/react.ts') ?
        { path: 'unused-core', namespace: 'unused-core' } : undefined);
      b.onLoad({ filter: /.*/, namespace: 'unused-core' }, () => ({ contents:
        'export const ' + ['initialize', 'convertDocxToHtml', 'compareDocuments', 'compareDocumentsToHtml',
          'getRevisions', 'getAnnotations', 'addAnnotation', 'addAnnotationWithTarget', 'removeAnnotation',
          'hasAnnotations', 'getDocumentStructure', 'getDocumentMetadata', 'isInitialized']
          .map(name => `${name}=()=>{throw new Error("Unused engine call")}`).join(',') + ';', loader: 'js' }));
    } }],
  });
  await page.setContent('<div id="root"></div>');
  await page.addScriptTag({ content: bundle.outputFiles[0].text });
  await page.evaluate(() => (window as any).render({ showPageNumbers: false }));
  await expect(page.locator('[data-docx-comment-panel]')).toHaveCount(1);
  const facts = () => page.evaluate(() => {
    const column = document.querySelector<HTMLElement>('.page-comment-margin')!;
    const box = document.querySelector<HTMLElement>('.page-box')!.getBoundingClientRect();
    const rect = column.getBoundingClientRect();
    column.scrollTop = column.scrollHeight;
    return { width: rect.width, gap: rect.left - box.right, pageWidth: box.width, scrolled: column.scrollTop > 0 };
  });
  expect(await facts()).toMatchObject({ width: 192, gap: 16, pageWidth: 816, scrolled: true });
  await page.evaluate(() => (window as any).render({ showPageNumbers: false,
    commentColumnWidthPt: 180, commentColumnGapPt: 18 }));
  await expect.poll(async () => (await facts()).width).toBe(240);
  expect(await facts()).toMatchObject({ gap: 24, pageWidth: 816 });
  await page.evaluate(() => (window as any).render({ showPageNumbers: false, commentColumnPlacement: 'margin' }));
  await expect(page.locator('[data-docx-comment-panel]')).toHaveCount(0);
  await expect(page.locator('.page-box .page-comment-margin')).toHaveCount(1);
});

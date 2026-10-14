import { test, expect, type Page } from '@playwright/test';
import { build } from 'esbuild';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// The React hooks' asynchronous loads, driven in a real page with a stub engine whose promises the
// test settles in any order (issue #973). A load for a document the caller has since replaced, or
// cleared, must not overwrite the newer state when it finally settles.

const npmRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// Stands in for `src/index.js` as imported by `src/react.ts`. Each engine call parks a promise
// keyed by the document's first byte; the test settles it with `settle(kind, tag, value)`.
const STUB_ENGINE = `
  const pending = [];
  export const calls = { structure: 0, annotations: 0 };
  function park(kind, document) {
    calls[kind]++;
    return new Promise((resolveIt) => pending.push({ kind, tag: document[0], resolveIt }));
  }
  export function settle(kind, tag, value) {
    const i = pending.findIndex((p) => p.kind === kind && p.tag === tag);
    if (i < 0) return false;
    pending.splice(i, 1)[0].resolveIt(value);
    return true;
  }
  export const pendingTags = (kind) => pending.filter((p) => p.kind === kind).map((p) => p.tag);
  export const isInitialized = () => true;
  export const initialize = async () => {};
  export const getDocumentStructure = (document) => park('structure', document);
  export const getAnnotations = (document) => park('annotations', document);
  const unused = () => { throw new Error('not stubbed'); };
  export const convertDocxToHtml = unused, compareDocuments = unused, compareDocumentsToHtml = unused,
    getRevisions = unused, addAnnotation = unused, addAnnotationWithTarget = unused,
    removeAnnotation = unused, hasAnnotations = unused, getDocumentMetadata = unused;
`;

const ENTRY = `
  import React from 'react';
  import { createRoot } from 'react-dom/client';
  import { flushSync } from 'react-dom';
  import { useDocumentStructure, useAnnotations } from './src/react.ts';
  import * as engine from 'stub-engine';

  const docs = { A: new Uint8Array([65]), B: new Uint8Array([66]) };
  window.state = {};
  function Probe({ hook, doc }) {
    const document = doc ? docs[doc] : null;
    const result = hook === 'structure' ? useDocumentStructure(document) : useAnnotations(document);
    window.state = {
      value: hook === 'structure' ? result.structure : result.annotations,
      isLoading: result.isLoading,
    };
    return null;
  }
  const root = createRoot(document.getElementById('root'));
  window.harness = {
    engine,
    render: (hook, doc) => flushSync(() => root.render(React.createElement(Probe, { hook, doc }))),
    unmount: () => root.unmount(),
  };
`;

async function loadHarness(page: Page) {
  const bundle = await build({
    stdin: { contents: ENTRY, resolveDir: npmRoot, loader: 'js' },
    bundle: true,
    format: 'iife',
    write: false,
    define: { 'process.env.NODE_ENV': '"development"' },
    plugins: [{
      name: 'stub-engine',
      setup(b) {
        b.onResolve({ filter: /^stub-engine$/ }, () => ({ path: 'stub-engine', namespace: 'stub' }));
        b.onResolve({ filter: /^\.\/index\.js$/ }, (args) =>
          args.importer.endsWith(`${'src'}/react.ts`) ? { path: 'stub-engine', namespace: 'stub' } : undefined);
        b.onLoad({ filter: /.*/, namespace: 'stub' }, () => ({ contents: STUB_ENGINE, loader: 'js' }));
      },
    }],
  });
  await page.setContent('<div id="root"></div>');
  await page.addScriptTag({ content: bundle.outputFiles[0].text });
}

const A = 65;
const B = 66;

/** Wait until the engine holds a load for `tag`, then return. */
async function waitForPending(page: Page, kind: string, tag: number) {
  await expect.poll(() => page.evaluate(
    ({ kind, tag }) => (window as any).harness.engine.pendingTags(kind).includes(tag), { kind, tag },
  )).toBe(true);
}

function settle(page: Page, kind: string, tag: number, value: unknown) {
  return page.evaluate(
    ({ kind, tag, value }) => (window as any).harness.engine.settle(kind, tag, value), { kind, tag, value },
  );
}

/** A few animation frames, so any state update a settled promise would make has been rendered. */
function frames(page: Page) {
  return page.evaluate(() => new Promise<void>((done) => {
    let n = 0;
    const tick = () => (++n >= 4 ? done() : requestAnimationFrame(tick));
    requestAnimationFrame(tick);
  }));
}

const state = (page: Page) => page.evaluate(() => (window as any).state);
const calls = (page: Page) => page.evaluate(() => ({ ...(window as any).harness.engine.calls }));

for (const kind of ['structure', 'annotations'] as const) {
  // A well-formed result for each hook, tagged with the document it was loaded for. The structure
  // hook derives paragraph and table lists from its result, so it must have their shape.
  const value = (tag: string) => (kind === 'structure'
    ? { root: { id: tag, type: 'Document', children: [] }, elementsById: {}, tableColumns: {} }
    : [{ id: tag }]);

  test.describe(`${kind === 'structure' ? 'useDocumentStructure' : 'useAnnotations'} (issue #973)`, () => {
    test.beforeEach(async ({ page }) => loadHarness(page));

    test('a slower load for a replaced document does not overwrite the newer one', async ({ page }) => {
      await page.evaluate((kind) => (window as any).harness.render(kind, 'A'), kind);
      await waitForPending(page, kind, A);
      await page.evaluate((kind) => (window as any).harness.render(kind, 'B'), kind);
      await waitForPending(page, kind, B);

      expect(await settle(page, kind, B, value('B'))).toBe(true);
      await expect.poll(async () => (await state(page)).value).toEqual(value('B'));

      expect(await settle(page, kind, A, value('A'))).toBe(true);
      await frames(page);
      expect(await state(page)).toEqual({ value: value('B'), isLoading: false });
    });

    test('a load still in flight when the document is cleared is ignored', async ({ page }) => {
      await page.evaluate((kind) => (window as any).harness.render(kind, 'A'), kind);
      await waitForPending(page, kind, A);
      await page.evaluate((kind) => (window as any).harness.render(kind, null), kind);

      expect(await settle(page, kind, A, value('A'))).toBe(true);
      await frames(page);
      const cleared = kind === 'structure' ? null : [];
      expect(await state(page)).toEqual({ value: cleared, isLoading: false });
    });

    test('one document loads once, not once per render', async ({ page }) => {
      await page.evaluate((kind) => (window as any).harness.render(kind, 'A'), kind);
      await waitForPending(page, kind, A);
      expect(await settle(page, kind, A, value('A'))).toBe(true);
      await expect.poll(async () => (await state(page)).value).toEqual(value('A'));
      await frames(page);
      expect((await calls(page))[kind]).toBe(1);
    });
  });
}

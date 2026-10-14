// Executed by check-package-boundary.mjs in an isolated consumer directory.
import assert from 'node:assert/strict';
import * as core from 'docxodus/core';

assert.equal(typeof globalThis.document, 'undefined');
assert.equal(typeof globalThis.window, 'undefined');
assert.equal(core.isInitialized(), false, 'import must not start WASM');
for (const name of ['DocxEditor', 'CommentGutter', 'mountRibbon']) {
  assert.equal(name in core, false, `${name} must remain outside the core entry`);
}

await Promise.all([core.initialize(), core.initialize()]);
assert.equal(core.isInitialized(), true);
assert.equal(core.wasmBasePath, new URL('./wasm/', import.meta.resolve('docxodus/core')).href);

const original = core.createBlankDocx();
assert.ok(original instanceof Uint8Array);
const session = core.openDocxSession(original);
let revised;
try {
  assert.ok(session instanceof core.DocxSession);
  const anchor = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:body:'));
  assert.ok(anchor);
  assert.equal(session.replaceText(anchor, 'Node core regression').success, true);
  revised = session.save();
} finally {
  session.close();
}

const html = await core.convertDocxToHtml(revised);
assert.match(html, /Node core regression/);
const redline = await core.compareDocuments(original, revised);
assert.ok(redline instanceof Uint8Array && redline.length > 0);
assert.ok((await core.docxDiffGetRevisions(original, revised)).length > 0);
assert.deepEqual(await core.docxDiffGetRevisions(revised, revised), []);

const annotations = await core.createExternalAnnotationSet(revised, 'node-core');
const annotation = core.createAnnotationFromSearch('node-annotation', 'NOTE', annotations.content, 'Node core');
assert.ok(annotation);
annotations.labelledText.push(annotation);
const annotated = await core.projectAnnotationsOntoHtml(html, annotations);
assert.match(annotated, /node-annotation/);

const browser = await import('./node_modules/docxodus/dist/browser-test.mjs');
assert.deepEqual(Object.keys(browser).sort(), [...Object.keys(core), 'DocxEditor', 'CommentGutter', 'mountRibbon'].sort());
for (const name of Object.keys(core)) {
  assert.equal(browser[name], core[name], `${name} must be shared across entries`);
}
assert.equal(browser.getWasmExports(), core.getWasmExports());
core.setWasmBasePath('/shared-runtime');
assert.equal(browser.wasmBasePath, '/shared-runtime/');
browser.setWasmBasePath('');
assert.equal(core.wasmBasePath, '');
await browser.initialize();
assert.equal(browser.getWasmExports(), core.getWasmExports());

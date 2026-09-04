// Regenerate the editor screenshots under docs/images/editor/ from the shipped surface.
//
// The pictures document `mountRibbon` as it ships, so they are captured from the real
// example host (npm/examples/editor.html served from npm/dist/wasm) with a real document —
// the NVCA Model Certificate of Incorporation, as every screenshot in
// docs/architecture/editor_ui_surface.md has been — rather than from a mock. Each capture
// drives the surface through the same `window.__demo` / `__selectTab` handles the specs use.
//
//   cd npm && npm run build && npm run pretest   # produces dist/wasm/ + the staged host page
//   cp ../TestFiles/NVCA-Model-COI.docx dist/wasm/sample.docx
//   python3 -m http.server 8082 --directory dist/wasm &
//   node ../tools/screenshots/editor/capture.mjs http://localhost:8082/editor.html ../docs/images/editor
//
// Uses the Chromium Playwright installed for the npm tests (no system Chrome needed).

import path from 'node:path';
import fs from 'node:fs';

const playwrightModule = process.env.PLAYWRIGHT_MODULE
  || new URL('../../../npm/node_modules/playwright/index.mjs', import.meta.url).href;
const { chromium } = await import(playwrightModule);

if (process.argv.length !== 4) {
  console.error('usage: node capture.mjs <editor-url> <output-dir>');
  process.exit(1);
}

const url = process.argv[2];
const outDir = path.resolve(process.argv[3]);
fs.mkdirSync(outDir, { recursive: true });

const browser = await chromium.launch({ headless: true });
const shots = [];
const shoot = async (page, name, options = {}) => {
  const file = path.join(outDir, `${name}.png`);
  await page.screenshot({ path: file, ...options });
  shots.push(file);
};
const settle = (page, ms = 250) => page.waitForTimeout(ms);
const selectTab = (page, name) => page.evaluate((n) => window.__selectTab(n), name);

try {
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 }, deviceScaleFactor: 1.25 });
  const page = await context.newPage();
  await page.goto(url);
  await page.waitForFunction(() => !!window.__demo && !!window.__demo.getEditor(), { timeout: 120000 });
  await page.waitForFunction(() => document.querySelector('.dxr')?.dataset.state === 'ready');
  await page.waitForSelector('#loader', { state: 'hidden' });

  // A body paragraph with a real sub-range selected, so the ribbon reflects a live selection.
  await page.evaluate(() => {
    const blocks = Array.from(document.querySelectorAll('.docx-body-flow [data-anchor][contenteditable="true"]'));
    const target = blocks.find((b) => (b.textContent || '').trim().length > 80) || blocks[0];
    target.scrollIntoView({ block: 'center' });
    target.focus();
    const walk = (n) => { if (n.nodeType === 3 && n.textContent.trim()) return n; for (const c of n.childNodes) { const r = walk(c); if (r) return r; } return null; };
    const tn = walk(target);
    const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, Math.min(24, tn.textContent.length));
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
  });
  await settle(page);
  await shoot(page, 'editor-overview');

  const chrome = () => page.locator('.dxr-chrome');
  await selectTab(page, 'home');
  await settle(page, 100);
  await shoot(page, 'ribbon-home', { clip: await chrome().boundingBox() });
  await selectTab(page, 'insert');
  await settle(page, 100);
  await shoot(page, 'ribbon-insert', { clip: await chrome().boundingBox() });
  await selectTab(page, 'layout');
  await settle(page, 100);
  await shoot(page, 'ribbon-layout', { clip: await chrome().boundingBox() });
  await selectTab(page, 'review');
  await settle(page, 100);
  await shoot(page, 'ribbon-review', { clip: await chrome().boundingBox() });
  await selectTab(page, 'view');
  await settle(page, 100);
  await shoot(page, 'ribbon-view', { clip: await chrome().boundingBox() });

  // Table picker open under its button.
  await selectTab(page, 'insert');
  await page.click('#table');
  await page.hover('#gridcells div[data-r="2"][data-c="3"]');
  await settle(page, 100);
  const picker = await page.locator('#gridpicker').boundingBox();
  const bar = await chrome().boundingBox();
  await shoot(page, 'table-picker', {
    clip: { x: bar.x, y: bar.y, width: bar.width, height: picker.y + picker.height - bar.y + 12 },
  });
  await page.keyboard.press('Escape');
  await page.mouse.click(5, 5);

  // Contextual Table tab: put the caret in the first table (insert one when the document has
  // none, and undo it afterwards so the later captures see the document unchanged).
  const insertedTable = await page.evaluate(() => {
    let cell = document.querySelector('.docx-body-flow table [data-anchor][contenteditable="true"]');
    let inserted = false;
    if (!cell) {
      const blocks = Array.from(document.querySelectorAll('.docx-body-flow [data-anchor][contenteditable="true"]'));
      const target = blocks.find((b) => (b.textContent || '').trim().length > 40) || blocks[0];
      target.focus();
      window.__demo.getEditor().insertTable(2, 3, { borderless: false });
      cell = document.querySelector('.docx-body-flow table [data-anchor][contenteditable="true"]');
      inserted = true;
    }
    cell.scrollIntoView({ block: 'center' });
    cell.focus();
    return inserted;
  });
  await page.waitForFunction(() => !document.querySelector('.dxr-tab[data-tab="table"]').hidden);
  await selectTab(page, 'table');
  await settle(page, 150);
  await shoot(page, 'ribbon-table-contextual', { clip: await chrome().boundingBox() });
  if (insertedTable) await page.evaluate(() => window.__demo.getEditor().undo());

  // Comments: select a range, post a comment from the gutter, capture the bubble beside it.
  await page.evaluate(() => {
    const blocks = Array.from(document.querySelectorAll('.docx-body-flow [data-anchor][contenteditable="true"]'));
    const target = blocks.find((b) => (b.textContent || '').trim().length > 120) || blocks[0];
    target.scrollIntoView({ block: 'center' });
    target.focus();
    const walk = (n) => { if (n.nodeType === 3 && n.textContent.trim()) return n; for (const c of n.childNodes) { const r = walk(c); if (r) return r; } return null; };
    const tn = walk(target);
    const r = document.createRange(); r.setStart(tn, 0); r.setEnd(tn, Math.min(30, tn.textContent.length));
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
  });
  await selectTab(page, 'review');
  await page.click('#comment');
  await page.keyboard.type('Confirm this survives the Series A closing.');
  await page.click('.docx-comment-bubble[data-draft] [data-comment-action="post"]');
  await page.waitForSelector('.docx-comment-bubble[data-thread]');
  await page.click('.docx-comment-bubble[data-thread] [data-comment-action="reply"]');
  await page.keyboard.type('It does — see §4.2(b).');
  await page.click('[data-comment-action="post-reply"]');
  await page.waitForFunction(() => document.querySelectorAll('.docx-comment-entry').length >= 2);
  await settle(page, 300);
  await shoot(page, 'comments');

  // Header band (continuous view): caret in the header with a running head typed, contextual
  // tab up. A document with no default header shows the band's placeholder; clicking it seeds
  // the story, exactly as a user would.
  await page.evaluate(() => {
    document.querySelector('.dxr-scroll').scrollTop = 0;
    const placeholder = document.querySelector('[data-hf-band="header"] [data-hf-placeholder]');
    if (placeholder) {
      placeholder.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    }
    const block = document.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]');
    if (!block) return;
    block.focus();
    if (!(block.textContent || '').trim()) {
      const range = document.createRange(); range.selectNodeContents(block);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
      document.execCommand('insertText', false, 'NVCA Model Legal Documents — Confidential Draft');
      block.dispatchEvent(new Event('blur'));
      document.querySelector('[data-hf-band="header"] [data-anchor][contenteditable="true"]')?.focus();
    }
  });
  await settle(page, 400);
  await shoot(page, 'header-band');

  // Paginated view with the footer being edited in place.
  await selectTab(page, 'view');
  await page.click('#viewpage');
  await page.waitForSelector('.page-box');
  await page.evaluate(() => {
    const area = document.querySelectorAll('.page-box')[1]?.querySelector('.page-footer') || document.querySelector('.page-box .page-footer');
    if (!area) return;
    area.scrollIntoView({ block: 'center' });
    const r = area.getBoundingClientRect();
    area.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: r.left + 20, clientY: r.top + 6 }));
  });
  await settle(page, 400);
  await shoot(page, 'paginated');

  // Footnotes section in the continuous view.
  await selectTab(page, 'view');
  await page.click('#viewweb');
  await page.waitForSelector('.docx-body-flow');
  const hadNotes = await page.evaluate(() => {
    const notes = document.querySelector('section.footnotes');
    if (!notes) return false;
    notes.scrollIntoView({ block: 'start' });
    document.querySelector('.dxr-scroll').scrollTop -= 40;
    return true;
  });
  if (hadNotes) {
    await settle(page, 300);
    await shoot(page, 'footnotes');
  }

  // Compact chrome on a phone-sized viewport.
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForFunction(() => document.querySelector('.dxr').dataset.chrome === 'compact');
  await selectTab(page, 'home');
  await settle(page, 400);
  await shoot(page, 'ribbon-compact');

  await context.close();
} finally {
  await browser.close();
}

for (const file of shots) console.log(`screenshot: ${file}`);

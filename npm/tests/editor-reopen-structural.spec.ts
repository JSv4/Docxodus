import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// ===========================================================================
// Regression: structural editing must survive re-opening on the same container.
//
// The single-contenteditable-root rearchitect wired keydown/beforeinput/blur on
// the editing-host ROOT exactly once (guarded by a dataset flag), with the
// listener closures bound to the FIRST editor instance. The demo's New / Open
// flow re-opens on the SAME container (`#editor` IS the contenteditable host, so
// the flag survived). `close()` didn't unwire, so every 2nd+ instance skipped
// re-wiring and its Enter/Backspace were handled by the first (now-closed)
// instance — whose handlers early-return. Plain Enter then fell through to native
// contenteditable, which cloned the paragraph (and its data-anchor), so the DOM
// showed multiple blocks aliasing ONE session paragraph. The model never gained
// paragraphs; the next remount collapsed the document. Silent + unrecoverable.
//
// Fix: re-wire the root listeners per open (keyed on the element) and unwire on
// close, so the live instance always owns structural input.
// ===========================================================================

test.describe('DocxEditor — structural editing survives re-open on the same container', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // Drive Enter the way the business-letter spec does: a synthetic keydown
  // dispatched on the block bubbles to the root handler. If the handler is bound
  // to the live instance, splitAtCaret → SplitParagraph reaches the SESSION; if it
  // is bound to the closed first instance, nothing happens and the session keeps
  // one paragraph. The session projection (save → reopen) is the ground truth.
  function bodyParaCount(md: string): number {
    return md.split('\n').filter((l) => /\{#p:body:/.test(l)).length;
  }

  test('Enter splits a paragraph in the SESSION after close + re-open on the same container', async ({ page }) => {
    const out = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);

      // The demo's New/Open flow: an editor is opened, then a NEW one is opened on
      // the SAME container (the demo closes the old one first).
      const a = D.DocxEditor.openBlank(container, D, {});
      a.close();
      const b = D.DocxEditor.openBlank(container, D, {});

      const editableEls = () =>
        Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[];

      // Type a line into the first block and put a collapsed caret at its end.
      const block = editableEls()[0];
      block.focus();
      block.textContent = 'hello world';
      const sel = window.getSelection()!;
      const range = document.createRange();
      const tn = block.firstChild ?? block;
      range.setStart(tn, (tn.textContent || '').length);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);

      // Enter — must split in the live (b) session, not be swallowed by closed a.
      block.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

      const saved: Uint8Array = b.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      const domAnchors = editableEls().map((e) => e.getAttribute('data-anchor'));

      b.close();
      container.remove();
      return { md, domBlockCount: domAnchors.length, uniqueAnchors: new Set(domAnchors).size === domAnchors.length };
    });

    // Enter reached the model: two body paragraphs (the original + the split-off).
    expect(bodyParaCount(out.md)).toBeGreaterThanOrEqual(2);
    // DOM blocks address distinct session paragraphs (no cloned/aliased anchors).
    expect(out.uniqueAnchors).toBe(true);
    expect(out.domBlockCount).toBeGreaterThanOrEqual(2);
  });

  // Defensive: even if a caller re-opens WITHOUT closing the previous editor
  // (e.g. a bfcache restore, or a buggy host), the new instance must take over
  // structural input — leftover listeners from a stale instance must not win.
  test('Enter splits in the SESSION when re-opened on the same container WITHOUT closing the old one', async ({ page }) => {
    const out = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);

      const a = D.DocxEditor.openBlank(container, D, {});
      // NOTE: deliberately NOT closing `a`.
      const b = D.DocxEditor.openBlank(container, D, {});

      const editableEls = () =>
        Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]')) as HTMLElement[];

      const block = editableEls()[0];
      block.focus();
      block.textContent = 'alpha beta';
      const sel = window.getSelection()!;
      const range = document.createRange();
      const tn = block.firstChild ?? block;
      range.setStart(tn, (tn.textContent || '').length);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);

      block.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }));

      const saved: Uint8Array = b.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      try { a.close(); } catch { /* a may already be torn down */ }
      b.close();
      container.remove();
      return { md };
    });

    expect(bodyParaCount(out.md)).toBeGreaterThanOrEqual(2);
  });
});

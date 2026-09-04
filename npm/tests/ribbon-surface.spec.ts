import { test, expect, type Page } from '@playwright/test';

/**
 * The ribbon surface itself (`mountRibbon`) — the shared editor UI that the
 * standalone example, the GitHub Pages landing page, the full-bleed app and the
 * compact player all host.
 *
 * social-demo.spec.ts proves the demo PAGES work; this proves the properties the
 * module owns on their behalf: density measured from the container, ids that stay
 * addressable when two surfaces share a page, a loader that can be driven by a
 * host that boots its own runtime, and a clean teardown.
 *
 * It drives /editor.html, which is the thinnest host — it supplies exports and
 * nothing else.
 */

async function openEditorHost(page: Page) {
  await page.goto('/editor.html');
  await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
  await page.click('#new');
  await page.waitForFunction(() => !!(window as any).__demo.getEditor());
}

test.describe('ribbon surface', () => {
  test('the example page is a host: the ribbon supplies the whole surface', async ({ page }) => {
    await openEditorHost(page);

    // Chrome, rail, tabs and document surface all come from the module.
    await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'ready');
    await expect(page.locator('.dxr-tab[data-tab="home"]')).toHaveAttribute('aria-selected', 'true');
    await expect(page.locator('#editor')).toHaveAttribute('data-dxr-surface', '');
    await expect(page.locator('#railSession')).toContainText('#');

    // The rail reports the LIVE session handle, not a made-up one.
    const railMatchesSession = await page.evaluate(() => {
      const editor = (window as any).__demo.getEditor();
      return document.getElementById('railSession')!.textContent === `#${editor.sessionHandle}`;
    });
    expect(railMatchesSession).toBe(true);

    // Every command routes through one timing wrapper, so the rail's "last op" is
    // a measurement rather than a label.
    await page.locator('button[data-cmd="bold"]').click();
    await expect(page.locator('#railOp')).toContainText(/^bold \d/);
  });

  test('density is measured from the container, not the viewport', async ({ page }) => {
    await openEditorHost(page);
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');

    // A wide viewport with a narrow MOUNT is still narrow. This is the case a
    // viewport media query gets wrong, and the reason the module uses a
    // ResizeObserver on its own root.
    await page.evaluate(() => {
      document.getElementById('app')!.style.width = '420px';
    });
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'compact');
    await expect(page.locator('.dxr-rail')).toBeHidden();
    await expect(page.locator('.dxr-hint')).toBeHidden();
    // Compact drops labels and rows, never commands.
    await expect(page.locator('button[data-cmd="subscript"]')).toBeAttached();
    await expect(page.locator('#fontfamily')).toBeAttached();

    await page.evaluate(() => {
      document.getElementById('app')!.style.width = '';
    });
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');
    await expect(page.locator('.dxr-rail')).toBeVisible();

    // An explicit mode wins over measurement, which is what player.html relies on.
    await page.evaluate(() => (window as any).__ribbon.setChrome('compact'));
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'compact');
    await page.evaluate(() => (window as any).__ribbon.setChrome('auto'));
    await expect(page.locator('.dxr')).toHaveAttribute('data-chrome', 'full');
  });

  test('a second surface on the same page takes prefixed ids and stays independent', async ({ page }) => {
    await openEditorHost(page);

    const result = await page.evaluate(() => {
      const { mountRibbon } = (window as any).DocxodusEditor;
      const host = document.createElement('div');
      host.id = 'second';
      host.style.cssText = 'height:400px';
      document.body.appendChild(host);

      const second = mountRibbon(host, {
        exports: (window as any).__demo.exports,
        loader: false,
        rail: false,
      });
      second.openBlank('second.docx');
      (window as any).__second = second;

      return {
        // The first surface keeps the bare ids it had.
        firstSurfaceIntact: document.getElementById('editor') === (window as any).__ribbon.surface,
        // The second could not reuse them, so it generated a prefix.
        secondPrefixed: /^dxr\d+-editor$/.test(second.surface.id),
        secondHasBlocks: second.surface.querySelectorAll('[data-anchor]').length > 0,
        // Opt-outs are honoured per instance.
        secondHasRail: !!second.element.querySelector('.dxr-rail'),
        secondHasLoader: !!second.element.querySelector('.dxr-loader'),
        // Distinct WASM sessions, not a shared one.
        distinctSessions:
          second.editor.sessionHandle !== (window as any).__demo.getEditor().sessionHandle,
      };
    });

    expect(result).toEqual({
      firstSurfaceIntact: true,
      secondPrefixed: true,
      secondHasBlocks: true,
      secondHasRail: false,
      secondHasLoader: false,
      distinctSessions: true,
    });

    // Commands on the second surface leave the first alone: the module scopes its
    // selection handling to its own surface.
    const isolated = await page.evaluate(() => {
      const second = (window as any).__second;
      const target = second.surface.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      const before = document.getElementById('editor')!.innerHTML;
      target.focus();
      const range = document.createRange();
      range.selectNodeContents(target);
      const selection = window.getSelection()!;
      selection.removeAllRanges();
      selection.addRange(range);
      second.element.querySelector('button[data-cmd="bold"]')!.dispatchEvent(
        new MouseEvent('click', { bubbles: true }),
      );
      return { firstUnchanged: document.getElementById('editor')!.innerHTML === before };
    });
    expect(isolated.firstUnchanged).toBe(true);

    // destroy() releases the session and removes the DOM it owns.
    const destroyed = await page.evaluate(() => {
      (window as any).__second.destroy();
      return {
        gone: !document.querySelector('#second .dxr'),
        firstStillLive: !!(window as any).__demo.getEditor(),
      };
    });
    expect(destroyed).toEqual({ gone: true, firstStillLive: true });
  });

  test('a host that boots its own runtime can drive the loading overlay', async ({ page }) => {
    await page.goto('/editor.html');
    // The overlay paints before any runtime exists — that is the point of it.
    await expect(page.locator('#loader')).toBeVisible();
    await expect(page.locator('#loaderTitle')).not.toBeEmpty();
    await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'loading');

    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await expect(page.locator('.dxr')).toHaveAttribute('data-state', 'ready');
    await expect(page.locator('#loader')).toBeHidden();

    // The stage API moves the bar and the copy together, and fail() surfaces the
    // error with a retry rather than leaving a dead surface.
    const driven = await page.evaluate(() => {
      const ribbon = (window as any).__ribbon;
      ribbon.loader.show();
      ribbon.loader.stage({ title: 'Custom stage', copy: 'Doing a thing', progress: 42, label: 'Working' });
      const width = document.getElementById('loaderBar')!.style.width;
      const stagedTitle = document.getElementById('loaderTitle')!.textContent;
      ribbon.loader.fail(new Error('boom'));
      return {
        width,
        stagedTitle,
        title: document.getElementById('loaderTitle')!.textContent,
        state: document.querySelector('.dxr')!.getAttribute('data-state'),
        copy: document.getElementById('loaderCopy')!.textContent,
        retryVisible: getComputedStyle(document.getElementById('loaderRetry')!).display !== 'none',
      };
    });
    expect(driven.width).toBe('42%');
    expect(driven.stagedTitle).toBe('Custom stage');
    expect(driven.title).toBe('The local engine did not start');
    expect(driven.state).toBe('error');
    expect(driven.copy).toBe('boom');
    expect(driven.retryVisible).toBe(true);
  });

  test('a draft bubble sits beside its selection, and its Ctrl+Enter posts without touching the paragraph', async ({ page }) => {
    await openEditorHost(page);
    await page.click('.dxr-tab[data-tab="review"]');
    await page.click('#editor [data-anchor][contenteditable="true"]');
    await page.keyboard.type('The indemnity clause needs review.');
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
    await page.click('button[data-dxr="comment"]');
    const draftBox = page.locator('.docx-comment-bubble[data-draft] textarea');
    await expect(draftBox).toBeFocused();

    // The first comment in a document drafts against a gutter that was display:none a moment
    // ago. Measured against that zero rect, the bubble landed a ribbon's height below its
    // selection; it must open level with the selected paragraph.
    const offset = await page.evaluate(() => {
      const bubble = document.querySelector('.docx-comment-bubble[data-draft]')!.getBoundingClientRect();
      const block = document.querySelector('#editor [data-anchor][contenteditable="true"]')!.getBoundingClientRect();
      return Math.abs(bubble.top - block.top);
    });
    expect(offset).toBeLessThan(40);

    // Ctrl+Enter in the bubble posts the draft. The same chord on a document block is the
    // ribbon's "page break before"; a keystroke inside the bubble must not reach it.
    await page.keyboard.type('Please reconsider this clause.');
    await page.keyboard.press('Control+Enter');
    await expect(page.locator('.docx-comment-bubble[data-thread]')).toHaveCount(1);
    await expect(page.locator('.docx-comment-bubble[data-draft]')).toHaveCount(0);
    const { comments, pageBreak } = await page.evaluate(() => {
      const editor = (window as any).__demo.getEditor();
      const exports = editor.exports;
      const handle = exports.DocxSessionBridge.OpenSession(editor.save(), '');
      try {
        const body = Object.keys(JSON.parse(exports.DocxSessionBridge.Project(handle)).anchorIndex)
          .find((k: string) => k.startsWith('p:body:'))!;
        return {
          comments: JSON.parse(exports.DocxSessionBridge.ListComments(handle)).map((c: any) => c.text),
          pageBreak: /pageBreakBefore/.test(exports.DocxSessionBridge.RawGetXml(handle, body)),
        };
      } finally {
        exports.DocxSessionBridge.CloseSession(handle);
      }
    });
    expect(comments).toHaveLength(1);
    expect(comments[0]).toContain('Please reconsider this clause.');
    expect(pageBreak).toBe(false);
  });

  test('the Review tab authors and resolves native comments through the gutter (issue #580)', async ({ page }) => {
    await openEditorHost(page);

    // The Review tab reveals the comment controls in their empty state.
    await page.click('.dxr-tab[data-tab="review"]');
    await expect(page.locator('.dxr-panel[data-panel="review"]')).toHaveAttribute('data-active', '');
    await expect(page.locator('[data-dxr="commentcount"]')).toContainText('No comments');

    // Give the block text (an empty paragraph is rightly refused with empty_comment_span),
    // leave it uncommitted — addComment must sync the typing before commenting — select a
    // word, and hit New Comment: a draft bubble opens beside the selection, Word-style.
    await page.click('#editor [data-anchor][contenteditable="true"]');
    await page.keyboard.type('The indemnity clause needs review.');
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
    await page.click('button[data-dxr="comment"]');
    await expect(page.locator('.docx-comment-bubble[data-draft] textarea')).toBeFocused();
    await page.keyboard.type('Please reconsider this clause.');
    await page.click('.docx-comment-bubble[data-draft] [data-comment-action="post"]');

    // Session truth: exactly one NATIVE comment thread carrying the typed text.
    const entries = await page.evaluate(() =>
      (window as any).__demo.getEditor().listComments());
    expect(entries).toHaveLength(1);
    expect(entries[0].text).toContain('Please reconsider this clause.');
    expect(entries[0].resolved).toBeFalsy();

    // The gutter shows the thread beside its highlight, and the Review tab counts it.
    await expect(page.locator('.docx-comment-bubble[data-thread]')).toHaveCount(1);
    await expect(page.locator('.docx-comment-bubble[data-thread] .docx-comment-text')).toContainText('Please reconsider');
    await expect(page.locator('#editor span.comment-highlight')).toHaveCount(1);
    await expect(page.locator('[data-dxr="commentcount"]')).toContainText('1 open');

    // Resolve from the ribbon: session truth flips and the bubble shows it.
    await page.click('#editor span.comment-highlight');
    await page.click('button[data-dxr="commentresolve"]');
    await expect(page.locator('.docx-comment-bubble[data-resolved]')).toHaveCount(1);
    const resolved = await page.evaluate(() =>
      (window as any).__demo.getEditor().listComments()[0].resolved);
    expect(resolved).toBe(true);

    // Reopen: the toggle is symmetric.
    await expect(page.locator('button[data-dxr="commentresolve"]')).toHaveText('Reopen');
    await page.click('button[data-dxr="commentresolve"]');
    await expect(page.locator('.docx-comment-bubble[data-resolved]')).toHaveCount(0);

    // The comment survives a save round-trip as native OOXML (not an overlay).
    const persisted = await page.evaluate(async () => {
      const editor = (window as any).__demo.getEditor();
      const bytes: Uint8Array = editor.save();
      const exports = editor.exports;
      const handle = exports.DocxSessionBridge.OpenSession(bytes, '');
      try {
        return JSON.parse(exports.DocxSessionBridge.ListComments(handle));
      } finally {
        exports.DocxSessionBridge.CloseSession(handle);
      }
    });
    expect(persisted).toHaveLength(1);
    expect(persisted[0].text).toContain('Please reconsider this clause.');
  });
});

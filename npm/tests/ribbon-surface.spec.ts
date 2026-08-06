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
});

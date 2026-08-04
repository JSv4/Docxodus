import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

/**
 * CDN embedding — the packaging contract for `dist/embed.bundle.js` / `dist/embed.iife.js`.
 *
 * The page lives on the main test origin (:8082) while the embed bundle AND every
 * WASM asset load from a second, CORS-enabled origin (:8083, `tests/cors-server.py`
 * serving `dist/`). That is byte-for-byte the jsDelivr/unpkg shape: cross-origin
 * `<script type="module">`/dynamic import, cross-origin `_framework` fetches under
 * `Access-Control-Allow-Origin: *` (which is why the build patches the runtime to
 * `credentials:"omit"`), and WASM base-path auto-detection with no configuration.
 *
 * These tests intentionally pass NO `wasmBasePath`: auto-detection working
 * cross-origin is the thing under test. If one of them fails after a loader or
 * build-script change, single-tag CDN embedding is broken for consumers even if
 * every same-origin spec still passes.
 */

const CDN_ORIGIN = 'http://localhost:8083';

function readTestFile(relativePath: string): number[] {
  return Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

test.describe('CDN embed — ESM bundle (embed.bundle.js)', () => {
  test('createViewer renders a DOCX with WASM auto-detected cross-origin', async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(
      async ({ cdn, docBytes }) => {
        const hostStyle = document.createElement('style');
        hostStyle.textContent = 'body { margin: 7px; font-family: monospace; }';
        document.head.appendChild(hostStyle);
        const mod = await import(/* @vite-ignore */ `${cdn}/embed.bundle.js`);
        const div = document.createElement('div');
        div.id = 'viewer';
        document.body.appendChild(div);
        const viewer = await mod.createViewer(div, new Uint8Array(docBytes));
        const root = div.querySelector<HTMLElement>('[data-docxodus-embed-root]')!;
        const documentStyles = Array.from(div.querySelectorAll<HTMLStyleElement>('style'));
        const result = {
          wasmBasePath: mod.wasmBasePath,
          hasContent: div.querySelectorAll('p, h1, h2, h3, table').length,
          hasStyles: div.querySelectorAll('style').length,
          htmlLength: viewer.html.length,
          hostMargin: getComputedStyle(document.body).margin,
          hostFont: getComputedStyle(document.body).fontFamily,
          documentMargin: getComputedStyle(root).margin,
          documentFont: getComputedStyle(root).fontFamily,
          stylesScoped: documentStyles.every((style) =>
            style.hasAttribute('data-docxodus-scoped-style')),
          hasDocumentGlobalAtRules: documentStyles.some((style) =>
            /@(page|import)\b/i.test(style.textContent ?? '')),
        };
        await viewer.reload(new Uint8Array(docBytes));
        const reloadedContent = div.querySelectorAll('p, h1, h2, h3, table').length;
        viewer.destroy();
        return { ...result, reloadedContent, destroyed: div.childNodes.length === 0 };
      },
      { cdn: CDN_ORIGIN, docBytes: readTestFile('HC031-Complicated-Document.docx') },
    );

    // Auto-detection must land on the CDN origin, not the page origin.
    expect(result.wasmBasePath).toBe(`${CDN_ORIGIN}/wasm/`);
    expect(result.hasContent).toBeGreaterThan(10);
    expect(result.hasStyles).toBeGreaterThan(0);
    expect(result.htmlLength).toBeGreaterThan(1000);
    expect(result.hostMargin).toBe('7px');
    expect(result.hostFont).toContain('monospace');
    expect(result.documentMargin).toBe('20px');
    expect(result.documentFont).toContain('Arial');
    expect(result.stylesScoped).toBe(true);
    expect(result.hasDocumentGlobalAtRules).toBe(false);
    expect(result.reloadedContent).toBeGreaterThan(10);
    expect(result.destroyed).toBe(true);
  });

  test('createEditor opens, edits, and saves a document loaded cross-origin', async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(
      async ({ cdn, docBytes }) => {
        const hostStyle = document.createElement('style');
        hostStyle.textContent = 'body { margin: 7px; font-family: monospace; }';
        document.head.appendChild(hostStyle);
        const mod = await import(/* @vite-ignore */ `${cdn}/embed.bundle.js`);
        const div = document.createElement('div');
        document.body.appendChild(div);
        const editor = await mod.createEditor(div, new Uint8Array(docBytes));
        const blocks = div.querySelectorAll('[data-anchor]').length;
        const block = div.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]');
        if (!block) throw new Error('No editable block rendered');
        const marker = ' PR347_REAL_EDIT';
        block.focus();
        block.textContent = (block.textContent ?? '') + marker;
        block.dispatchEvent(new InputEvent('input', {
          bubbles: true,
          inputType: 'insertText',
          data: marker,
        }));
        block.blur();

        // A later editor remount creates a fresh converter stylesheet. It must
        // already be scoped when setPaginated returns, including removal of @page.
        editor.setPaginated(true);
        const saved: Uint8Array = editor.save();
        const hostMargin = getComputedStyle(document.body).margin;
        const hostFont = getComputedStyle(document.body).fontFamily;
        const documentStyles = Array.from(div.querySelectorAll<HTMLStyleElement>('style'));
        const stylesScoped = documentStyles.every((style) =>
          style.hasAttribute('data-docxodus-scoped-style'));
        const hasDocumentGlobalAtRules = documentStyles.some((style) =>
          /@(page|import)\b/i.test(style.textContent ?? ''));
        editor.close();

        const verification = document.createElement('div');
        document.body.appendChild(verification);
        const reopened = await mod.createEditor(verification, saved);
        const persisted = verification.textContent?.includes(marker) ?? false;
        reopened.close();
        return {
          blocks,
          savedLength: saved.length,
          persisted,
          hostMargin,
          hostFont,
          styleCount: documentStyles.length,
          stylesScoped,
          hasDocumentGlobalAtRules,
        };
      },
      { cdn: CDN_ORIGIN, docBytes: readTestFile('HC001-5DayTourPlanTemplate.docx') },
    );

    expect(result.blocks).toBeGreaterThan(0);
    expect(result.savedLength).toBeGreaterThan(1000);
    expect(result.persisted).toBe(true);
    expect(result.hostMargin).toBe('7px');
    expect(result.hostFont).toContain('monospace');
    expect(result.styleCount).toBeGreaterThan(0);
    expect(result.stylesScoped).toBe(true);
    expect(result.hasDocumentGlobalAtRules).toBe(false);
  });

  test('createEditor with no source opens a blank document', async ({ page }) => {
    await page.goto('/');
    const result = await page.evaluate(async (cdn) => {
      const mod = await import(/* @vite-ignore */ `${cdn}/embed.bundle.js`);
      const div = document.createElement('div');
      document.body.appendChild(div);
      const editor = await mod.createEditor(div);
      const blocks = div.querySelectorAll('[data-anchor]').length;
      const saved: Uint8Array = editor.save();
      editor.close();
      return { blocks, savedLength: saved.length };
    }, CDN_ORIGIN);

    expect(result.blocks).toBeGreaterThan(0);
    expect(result.savedLength).toBeGreaterThan(1000);
  });
});

test.describe('CDN embed — wasm-webroot layout fallback', () => {
  // The demo webroot serves the wasm directory itself, so the bundle sits NEXT
  // to _framework/ instead of next to a wasm/ subdirectory. ensureWasm() must
  // probe wasm/ first (package/CDN layout), fail, and fall back to the bundle's
  // own directory — which also proves initialize() is retryable after failure.
  test('embed.html demo boots with assets next to the bundle', async ({ page }) => {
    await page.goto('/embed.html');
    await expect(page.locator('#status')).toContainText('ready', { timeout: 45000 });
    expect(await page.locator('#editor [data-anchor]').count()).toBeGreaterThan(0);
  });
});

test.describe('CDN embed — classic script (embed.iife.js)', () => {
  test('window.Docxodus resolves WASM via document.currentScript', async ({ page }) => {
    await page.goto('/');
    // A classic <script src> from the CDN origin — no modules, no bundler.
    await page.evaluate(async (cdn) => {
      await new Promise<void>((resolve, reject) => {
        const s = document.createElement('script');
        s.src = `${cdn}/embed.iife.js`;
        s.onload = () => resolve();
        s.onerror = () => reject(new Error('embed.iife.js failed to load'));
        document.head.appendChild(s);
      });
    }, CDN_ORIGIN);

    const result = await page.evaluate(async () => {
      const Docxodus = (window as any).Docxodus;
      const div = document.createElement('div');
      document.body.appendChild(div);
      const editor = await Docxodus.createEditor(div);
      const blocks = div.querySelectorAll('[data-anchor]').length;
      editor.close();
      return { hasGlobal: typeof Docxodus.createViewer === 'function', blocks };
    });

    expect(result.hasGlobal).toBe(true);
    expect(result.blocks).toBeGreaterThan(0);
  });
});

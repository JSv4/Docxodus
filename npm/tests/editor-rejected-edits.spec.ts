import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

// When the engine rejects an edit, the editor must not keep showing text the document does not
// hold, and it must say so (issue #969). The bridge is wrapped so its text-replacement calls
// return a real engine rejection on demand; everything else is the live engine.

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  return Array.from(new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath))));
}

async function openWithRejectingBridge(page: Page, withHandler = true) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ bytesArray, withHandler }) => {
    const w = window as any;
    const D = w.Docxodus;
    const rejection = JSON.stringify({
      success: false, created: [], removed: [], modified: [],
      error: { code: 'anchor_not_found', message: 'forced rejection' },
    });
    w.rejectText = false;
    w.bridgeCalls = [] as string[];
    const bridge = new Proxy(D.DocxSessionBridge, {
      get(target, key) {
        const value = target[key];
        if (typeof value !== 'function') return value;
        return (...args: unknown[]) => {
          w.bridgeCalls.push(String(key));
          if (w.rejectText && (key === 'ReplaceText' || key === 'ReplaceTextAtSpan')) return rejection;
          return value.apply(target, args);
        };
      },
    });
    const exportsWithBridge = new Proxy(D, {
      get: (target, key) => (key === 'DocxSessionBridge' ? bridge : target[key]),
    });
    const container = document.createElement('div');
    document.body.appendChild(container);
    w.failures = [] as unknown[];
    w.editor = D.DocxEditor.open(container, new Uint8Array(bytesArray), exportsWithBridge,
      withHandler ? { onEditFailed: (info: unknown) => w.failures.push(info) } : {});
    w.container = container;
    w.firstEditable = () => Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]'))
      .find((e: any) => (e.textContent || '').trim().length > 5) as HTMLElement;
  }, { bytesArray: readTestFile('HC031-Complicated-Document.docx'), withHandler });
}

test.describe('DocxEditor — rejected edits (#969)', () => {
  test('a rejected text commit restores the block and reports the failure', async ({ page }) => {
    await openWithRejectingBridge(page);
    const out = await page.evaluate(() => {
      const w = window as any;
      const target = w.firstEditable();
      const anchor = target.getAttribute('data-anchor');
      const original = target.textContent;
      w.rejectText = true;
      target.focus();
      target.textContent = 'REJECTED-BY-ENGINE';
      target.dispatchEvent(new Event('blur'));
      w.rejectText = false;
      const now = w.container.querySelector(`[data-anchor="${anchor}"]`);
      return {
        restored: now?.textContent === original,
        stillShowsRejected: w.container.textContent.includes('REJECTED-BY-ENGINE'),
        failures: w.failures,
      };
    });
    expect(out.restored).toBe(true);
    expect(out.stillShowsRejected).toBe(false);
    expect(out.failures).toHaveLength(1);
    expect(out.failures[0]).toMatchObject({ code: 'anchor_not_found', message: 'forced rejection' });
    expect(typeof (out.failures[0] as { anchorId: unknown }).anchorId).toBe('string');
  });

  test('a structural op stops when flushing the typed text is rejected', async ({ page }) => {
    await openWithRejectingBridge(page);
    const out = await page.evaluate(() => {
      const w = window as any;
      const target = w.firstEditable();
      const anchor = target.getAttribute('data-anchor');
      const original = target.textContent;
      target.focus();
      target.textContent = 'TYPED-THEN-REJECTED';
      w.rejectText = true;
      w.bridgeCalls.length = 0;
      w.editor.setAlignment('center');
      w.rejectText = false;
      const now = w.container.querySelector(`[data-anchor="${anchor}"]`);
      return {
        restored: now?.textContent === original,
        calls: [...w.bridgeCalls],
        failures: w.failures.length,
      };
    });
    expect(out.restored).toBe(true);
    expect(out.failures).toBe(1);
    // The flush was attempted; the alignment was not applied on top of the rejected text.
    expect(out.calls.some((c: string) => c.startsWith('ReplaceText'))).toBe(true);
    expect(out.calls).not.toContain('SetParagraphFormat');
  });

  test('without a handler a rejection is logged, not swallowed', async ({ page }) => {
    const warnings: string[] = [];
    page.on('console', (m) => { if (m.type() === 'warning') warnings.push(m.text()); });
    await openWithRejectingBridge(page, false);
    await page.evaluate(() => {
      const w = window as any;
      const target = w.firstEditable();
      w.rejectText = true;
      target.focus();
      target.textContent = 'NO-HANDLER';
      target.dispatchEvent(new Event('blur'));
      w.rejectText = false;
    });
    await expect.poll(() => warnings.some((t) => t.includes('rejected an edit') && t.includes('forced rejection')))
      .toBe(true);
  });
});

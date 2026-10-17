import { test, expect } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/**
 * Issue #1151: the paginator gives every single-spaced line its font's unrounded height before it
 * lays pages out. A block the paginated editor renders afresh after an edit gets the same pass, so
 * an edited paragraph keeps the pitch of the page around it instead of Chromium's whole-pixel one.
 */
test('an edited paragraph in the paginated editor keeps its unrounded line height', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  const bytes = fs.readFileSync(path.join(__dirname, '../../TestFiles/HC031-Complicated-Document.docx'));

  const out = await page.evaluate((array: number[]) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    document.body.appendChild(container);
    const editor = D.DocxEditor.open(container, new Uint8Array(array), D, { paginated: true });
    const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
    const target = Array.from(container.querySelectorAll<HTMLElement>(
      '#pagination-container p[data-anchor][contenteditable="true"]'))
      .find((p) => norm(p.textContent || '').length > 5 && getComputedStyle(p).lineHeight !== 'normal')!;
    const before = getComputedStyle(target).lineHeight;
    target.focus();
    target.textContent = 'PITCHEDIT41 content';
    target.dispatchEvent(new Event('blur'));
    const edited = Array.from(container.querySelectorAll<HTMLElement>('#pagination-container p[data-anchor]'))
      .find((p) => norm(p.textContent || '').includes('PITCHEDIT41'))!;
    const result = { before, after: getComputedStyle(edited).lineHeight, swapped: edited !== target };
    editor.close();
    container.remove();
    return result;
  }, Array.from(bytes));

  expect(out.swapped).toBe(true);
  expect(out.before).not.toBe('normal');
  expect(out.after).toBe(out.before);
});

import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue #1677: a bordered paragraph's space before and after are drawn outside its border, on the
// box the renderer wraps it in. A paragraph-format edit therefore has to re-render that box; a
// single-block swap would put the new paragraph inside a box that still has the old spacing.
test.describe('DocxEditor — spacing of a bordered paragraph', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('setParagraphSpacing moves the box, not the text inside it', async ({ page }) => {
    const spacing = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement).focus();
      editor.insertHorizontalRule(12, 'single');
      const boxed = () => Array.from(container.querySelectorAll<HTMLElement>('p[data-anchor]'))
        .find((p) => /border-bottom/.test(p.parentElement?.getAttribute('style') ?? ''))!;
      const read = () => {
        const box = getComputedStyle(boxed().parentElement!), own = getComputedStyle(boxed());
        return { box: [box.marginTop, box.marginBottom], own: [own.marginTop, own.marginBottom] };
      };
      const before = read();
      boxed().focus();
      editor.setParagraphSpacing({ beforePt: 30, afterPt: 15 });
      const after = read();
      editor.close(); container.remove();
      return { before, after };
    });

    expect(spacing.before.own).toEqual(['0px', '0px']);
    expect(spacing.after).toEqual({ box: ['40px', '20px'], own: ['0px', '0px'] });
    expect(spacing.before.box).not.toEqual(spacing.after.box);
  });
});

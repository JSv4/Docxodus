import { test, expect } from '@playwright/test';

// Issue 8 — arbitrary font size (demo). The size control was a <select> capped at 48pt; it is
// now an editable combobox accepting any positive value. Driving the real demo (editor.html):
// typing 72 sets a 72pt run (~96px), well beyond the old cap.
test.describe('Demo — arbitrary font size', () => {
  test('typing 72 in the size field sets a 72pt run (beyond the old 48 cap)', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    // Type "BIG" into the first paragraph.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.keyboard.type('BIG');

    // Select the paragraph, then set the size field to 72 and apply.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.fill('#fontsize', '72');
    await page.locator('#fontsize').dispatchEvent('change');

    const px = await page.evaluate(() => {
      const span = document.querySelector('#editor p[data-anchor] span') as HTMLElement | null;
      return span ? parseFloat(getComputedStyle(span).fontSize) : 0;
    });
    // 72pt ≈ 96px — comfortably beyond the old 48pt (64px) ceiling.
    expect(px).toBeGreaterThan(90);
  });

  // S-1 smoke-test finding 3: focusing the size field blurs the block and collapsed the
  // selection, so the combobox could only size a WHOLE paragraph. The editor now caches the
  // last real selection, so the combobox sizes just the selected sub-range.
  test('sizes only the selected sub-range, not the whole paragraph', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    // Type "BIGsmall" into the first paragraph.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.keyboard.type('BIGsmall');

    // Select only the first three characters ("BIG"), then apply 28pt via the combobox.
    await page.evaluate(() => {
      const tn = document.querySelector('#editor p[data-anchor] span')!.firstChild as Text;
      const r = document.createRange();
      r.setStart(tn, 0);
      r.setEnd(tn, 3);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.waitForTimeout(50); // let the editor's selectionchange cache the sub-range
    await page.fill('#fontsize', '28');
    await page.locator('#fontsize').dispatchEvent('change');

    const sizes = await page.evaluate(() => {
      const block = document.querySelector('#editor p[data-anchor]') as HTMLElement;
      const big = Array.from(block.querySelectorAll('span')).find((s) => /BIG/.test(s.textContent || ''));
      const small = Array.from(block.querySelectorAll('span')).find(
        (s) => /small/.test(s.textContent || '') && !/BIG/.test(s.textContent || ''),
      );
      return {
        bigPx: big ? parseFloat(getComputedStyle(big).fontSize) : 0,
        smallPx: small ? parseFloat(getComputedStyle(small).fontSize) : 0,
      };
    });

    expect(sizes.bigPx).toBeGreaterThan(30);   // "BIG" got the 28pt (~37px)
    expect(sizes.smallPx).toBeGreaterThan(0);  // "small" is its own run
    expect(sizes.smallPx).toBeLessThan(20);    // …left at the default ~11pt, NOT resized
  });

  // S-1 round-3 regression: the size field bound applyFontSize to BOTH `change` and keydown-Enter,
  // so pressing Enter fired it twice → two undo snapshots → one size change took TWO Ctrl+Z to
  // revert. Enter now just commits via blur (single `change`), so ONE undo reverts the size.
  test('applying a size via Enter is reverted by a single undo (no double-fire)', async ({ page }) => {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('#new');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    // Type "WORD" into the first paragraph and select it.
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor][contenteditable="true"]') as HTMLElement;
      p.focus();
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.keyboard.type('WORD');
    await page.evaluate(() => {
      const p = document.querySelector('#editor p[data-anchor]') as HTMLElement;
      const r = document.createRange();
      r.selectNodeContents(p);
      const s = window.getSelection()!;
      s.removeAllRanges();
      s.addRange(r);
    });
    await page.waitForTimeout(50);

    // Apply 40pt via the field, committing with a real Enter keypress.
    await page.click('#fontsize');
    await page.fill('#fontsize', '40');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(50);

    const sizePx = () =>
      page.evaluate(() => {
        const span = document.querySelector('#editor p[data-anchor] span') as HTMLElement | null;
        return span ? parseFloat(getComputedStyle(span).fontSize) : 0;
      });
    const appliedPx = await sizePx();

    // A SINGLE undo must restore the original size.
    await page.evaluate(() => (window as any).__demo.getEditor().undo());
    await page.waitForTimeout(50);
    const afterOneUndoPx = await sizePx();

    expect(appliedPx).toBeGreaterThan(45);     // 40pt ≈ 53px applied
    expect(afterOneUndoPx).toBeLessThan(20);   // reverted to ~11pt with ONE undo (was 2 before)
  });
});

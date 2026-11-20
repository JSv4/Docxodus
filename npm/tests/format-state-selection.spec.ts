import { test, expect } from '@playwright/test';

/**
 * DocxEditor.queryFormatState (what the ribbon's Bold, Italic, ... buttons show) reads the format
 * of the first character the selection covers. After formatting part of a paragraph the editor
 * restores the selection starting at the END of the text node before it ("Hello from |"), and the
 * state was read there: the bold just applied showed as not bold.
 */
test('the format of a selection is read at its first character, not at the text before it', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'fs-editor';
    document.body.appendChild(container);
    (window as any).__ed = D.DocxEditor.openBlank(container, D, {});
  });
  const block = page.locator('#fs-editor [data-anchor][contenteditable="true"]').first();
  await block.click();
  await page.keyboard.type('Hello from world');

  const result = await page.evaluate(() => {
    const ed = (window as any).__ed;
    // (the paragraph's element is replaced when it re-renders: looked up each time)
    const paragraph = () => document.querySelector<HTMLElement>('#fs-editor [data-anchor][contenteditable="true"]')!;
    const textNodes = () => {
      const out: Text[] = [];
      const w = document.createTreeWalker(paragraph(), NodeFilter.SHOW_TEXT);
      for (let t = w.nextNode(); t; t = w.nextNode()) if ((t.textContent ?? '').length) out.push(t as Text);
      return out;
    };
    const select = (startNode: Node, start: number, endNode: Node, end: number) => {
      const r = document.createRange();
      r.setStart(startNode, start);
      r.setEnd(endNode, end);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(r);
    };
    // the typing committed (the editor commits a paragraph's typing when it loses the focus)
    paragraph().blur();
    paragraph().focus();
    // "world" selected and made bold, as the ribbon's Bold button does it
    const whole = textNodes().find((t) => t.textContent!.includes('world'))!;
    const at = whole.textContent!.indexOf('world');
    select(whole, at, whole, at + 5);
    ed.format('bold', true);
    const restored = ed.queryFormatState().bold;

    // the same selection by hand: from the end of "Hello from " to the end of "world"
    const [plain, bold] = textNodes();
    select(plain, plain.length, bold, bold.length);
    const byHand = ed.queryFormatState().bold;
    // a selection that starts in the plain text reads it
    select(plain, 6, bold, bold.length);
    const fromPlain = ed.queryFormatState().bold;
    // a caret takes the format of the text before it
    select(plain, plain.length, plain, plain.length);
    const caretAfterPlain = ed.queryFormatState().bold;
    select(bold, 2, bold, 2);
    const caretInBold = ed.queryFormatState().bold;
    return { texts: textNodes().map((t) => t.textContent), restored, byHand, fromPlain, caretAfterPlain, caretInBold };
  });
  expect(result.texts).toEqual(['Hello from ', 'world']);
  expect(result.restored).toBe(true);
  expect(result.byHand).toBe(true);
  expect(result.fromPlain).toBe(false);
  expect(result.caretAfterPlain).toBe(false);
  expect(result.caretInBold).toBe(true);
});

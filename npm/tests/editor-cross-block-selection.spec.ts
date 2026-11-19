import { test, expect, type Page } from '@playwright/test';

/**
 * Selections across paragraphs, as Word has them. Every paragraph is its own editing host, which
 * fences a browser's selection at its edge: the editor makes Shift+click, Shift+arrow keys and
 * Ctrl+A reach across paragraphs, and makes typing, Backspace, Enter, cut and paste over such a
 * selection one edit (one undo step) instead of an edit to the focused paragraph alone.
 */

const blocks = (page: Page) => page.locator('#xb-host [data-anchor][contenteditable="true"]');
const texts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#xb-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/ /g, ' ').trim()));
/** The indexes of the blocks the selection starts and ends in, and whether it is collapsed. */
const selection = (page: Page) => page.evaluate(() => {
  const all = Array.from(document.querySelectorAll('#xb-host [data-anchor][contenteditable="true"]'));
  const sel = window.getSelection()!;
  const r = sel.getRangeAt(0);
  const index = (n: Node) => all.findIndex((b) => b.contains(n));
  return { from: index(r.startContainer), to: index(r.endContainer), collapsed: r.collapsed };
});

async function mount(page: Page, paragraphs: string[]) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'xb-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('cross.docx');
    ribbon.loader.done();
    (window as any).__xbRibbon = ribbon;
  });
  await blocks(page).first().click();
  for (const [i, p] of paragraphs.entries()) {
    await page.keyboard.type(p);
    if (i < paragraphs.length - 1) await page.keyboard.press('Enter');
  }
  await expect(blocks(page)).toHaveCount(paragraphs.length);
}

/** Focus block `a` and select from character `startOffset` of it to `endOffset` of block `b`. */
async function selectAcross(page: Page, a: number, startOffset: number, b: number, endOffset: number) {
  await blocks(page).nth(a).focus();
  await page.evaluate(({ a, startOffset, b, endOffset }) => {
    const all = document.querySelectorAll<HTMLElement>('#xb-host [data-anchor][contenteditable="true"]');
    const at = (block: HTMLElement, offset: number): [Node, number] => {
      const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      let left = offset;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        const len = (n.nodeValue ?? '').length;
        if (left <= len) return [n, left];
        left -= len;
      }
      return [block, block.childNodes.length];
    };
    const range = document.createRange();
    range.setStart(...at(all[a], startOffset));
    range.setEnd(...at(all[b], endOffset));
    const sel = window.getSelection()!;
    sel.removeAllRanges();
    sel.addRange(range);
  }, { a, startOffset, b, endOffset });
}

const PARAS = ['Alpha one two.', 'Beta three four.', 'Gamma five six.'];

test('Shift+click extends the selection into another paragraph', async ({ page }) => {
  await mount(page, PARAS);
  await blocks(page).nth(0).click({ position: { x: 5, y: 5 } });
  await blocks(page).nth(2).click({ position: { x: 30, y: 5 }, modifiers: ['Shift'] });
  expect(await selection(page)).toEqual({ from: 0, to: 2, collapsed: false });
});

test('Shift+Down and Shift+Up move the selection across paragraphs, from its fixed end', async ({ page }) => {
  await mount(page, PARAS);
  await blocks(page).nth(0).click({ position: { x: 5, y: 5 } });
  await page.keyboard.press('Home');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  expect(await selection(page)).toMatchObject({ from: 0, to: 2, collapsed: false });
  await page.keyboard.press('Shift+ArrowUp');
  expect(await selection(page)).toMatchObject({ from: 0, to: 1, collapsed: false });
  // Ctrl+Shift+End: to the end of the document.
  await page.keyboard.press('Control+Shift+End');
  expect(await selection(page)).toMatchObject({ from: 0, to: 2 });
});

test('Shift+Right at a paragraph end goes on into the next paragraph', async ({ page }) => {
  await mount(page, PARAS);
  await blocks(page).nth(0).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  expect(await selection(page)).toMatchObject({ from: 0, to: 1, collapsed: false });
});

test('Ctrl+A selects every paragraph; typing replaces them all, and one undo brings them back', async ({ page }) => {
  await mount(page, PARAS);
  await blocks(page).nth(1).click();
  await page.keyboard.press('Control+a');
  expect(await selection(page)).toMatchObject({ from: 0, to: 2, collapsed: false });
  await page.keyboard.type('X');
  await expect.poll(() => texts(page)).toEqual(['X']);
  await page.keyboard.press('Control+z');
  await expect.poll(() => texts(page)).toEqual(PARAS);
});

test('typing over part of two paragraphs joins what is left, as Word does', async ({ page }) => {
  await mount(page, PARAS);
  await selectAcross(page, 0, 4, 1, 7);    // "a one two." + "Beta th"
  await page.keyboard.type('X');
  await expect.poll(() => texts(page)).toEqual(['AlphXree four.', 'Gamma five six.']);
  // The caret is after the X: typing goes on there.
  await page.keyboard.type('Y');
  await page.locator('body').click({ position: { x: 1, y: 1 } });
  await expect.poll(() => texts(page)).toEqual(['AlphXYree four.', 'Gamma five six.']);
});

test('Backspace removes a selection over three paragraphs; one undo brings it back', async ({ page }) => {
  await mount(page, PARAS);
  await selectAcross(page, 0, 5, 2, 5);
  await page.keyboard.press('Backspace');
  await expect.poll(() => texts(page)).toEqual(['Alpha five six.']);
  await page.keyboard.press('Control+z');
  await expect.poll(() => texts(page)).toEqual(PARAS);
});

test('Enter over a selection across paragraphs leaves two paragraphs split there', async ({ page }) => {
  await mount(page, PARAS);
  await selectAcross(page, 0, 4, 1, 7);
  await page.keyboard.press('Enter');
  await expect.poll(() => texts(page)).toEqual(['Alph', 'ree four.', 'Gamma five six.']);
});

test('cut and paste over a selection across paragraphs', async ({ page }) => {
  await mount(page, PARAS);
  await selectAcross(page, 0, 4, 1, 7);
  const cut = await page.evaluate(() => {
    const data = new DataTransfer();
    const ev = new ClipboardEvent('cut', { clipboardData: data, bubbles: true, cancelable: true });
    document.activeElement!.dispatchEvent(ev);
    return data.getData('text/plain');
  });
  expect(cut).toBe('a one two.\nBeta th');
  await expect.poll(() => texts(page)).toEqual(['Alphree four.', 'Gamma five six.']);

  await selectAcross(page, 0, 4, 1, 6);    // "ree four." + "Gamma "
  await page.evaluate(() => {
    const data = new DataTransfer();
    data.setData('text/plain', 'one\ntwo');
    document.activeElement!.dispatchEvent(new ClipboardEvent('paste', { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await expect.poll(() => texts(page)).toEqual(['Alphone', 'twofive six.']);
});

test('with changes tracked, typing over paragraphs loses no text and one undo restores them', async ({ page }) => {
  await mount(page, PARAS);
  await page.evaluate(() => document.querySelector<HTMLInputElement>('#xb-host [data-dxr="trackchanges"]')!.click());
  await blocks(page).nth(0).focus();
  await selectAcross(page, 0, 4, 1, 7);
  await page.keyboard.type('X');
  // The deleted text stays on the page, struck through, until the change is accepted or rejected;
  // the X is a tracked insertion after it.
  await expect.poll(() => page.locator('#xb-host ins').filter({ hasText: 'X' }).count()).toBeGreaterThan(0);
  const deleted = await page.locator('#xb-host del').allTextContents();
  expect(deleted.join('|')).toContain('a one two.');
  expect(deleted.join('|')).toContain('Beta th');
  const all = (await texts(page)).join('|');
  expect(all).toContain('ree four.');
  expect(all).toContain('Gamma five six.');
  await page.keyboard.press('Control+z');
  await expect.poll(() => texts(page)).toEqual(PARAS);
});

const header = (page: Page) => page.locator('#xb-host [data-hf-band="header"] [data-anchor][contenteditable="true"]');
const bodyTexts = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#xb-host .docx-body-flow [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').replace(/\u00a0/g, ' ').trim()));

test('in a header, Ctrl+A and typing edit the header (the browser selects and types there)', async ({ page }) => {
  await mount(page, PARAS);
  await page.locator('#xb-host [data-hf-band="header"] [data-hf-placeholder]').click();
  await expect(header(page)).toHaveCount(1);
  await header(page).first().click();
  await page.keyboard.type('Head one');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Head two');
  await expect(header(page)).toHaveCount(2);

  await page.keyboard.press('Control+a');
  await page.keyboard.type('X');

  // The key reaches the header (it used to be dropped): its paragraph now reads X.
  await expect.poll(async () => (await header(page).allTextContents()).map((t) => t.trim())).toContain('X');
  expect(await bodyTexts(page)).toEqual(PARAS);
});

test('over a selection reaching into a table, typing is not dropped', async ({ page }) => {
  await mount(page, PARAS);
  await blocks(page).nth(2).click();
  await page.keyboard.press('End');
  await page.evaluate(() => (window as any).__xbRibbon.editor.insertTable(1, 2, { cellContents: ['Cell A', 'Cell B'] }));
  const all = () => page.evaluate(() =>
    Array.from(document.querySelectorAll<HTMLElement>('#xb-host .docx-body-flow [data-anchor][contenteditable="true"]'))
      .map((b) => (b.textContent ?? '').trim()));
  await expect.poll(async () => (await all()).includes('Cell A')).toBe(true);
  const cell = (await all()).indexOf('Cell A');
  await selectAcross(page, 1, 5, cell, 4);    // from "Beta |three" into "Cell| A"
  const before = (await all()).join('|');

  await page.keyboard.type('Z');

  await expect.poll(async () => (await all()).join('|')).not.toEqual(before);
  expect((await all()).join('|')).toContain('Z');
});

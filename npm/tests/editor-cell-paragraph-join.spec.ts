import { test, expect, type Page } from '@playwright/test';

/**
 * Backspace at a paragraph's start and Delete at its end join two paragraphs, as in Word: in a
 * table cell only paragraphs of the same cell, so an empty paragraph in a cell can be removed while
 * the cell boundary stays.
 */

const HOST = '#cj-host';

async function mount(page: Page) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'cj-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const r = (window as any).DocxodusEditor.mountRibbon(host, {});
    r.setExports((window as any).Docxodus);
    r.openBlank('j.docx');
    r.loader.done();
    (window as any).__ribbon = r;
  });
}

/** Each table cell's paragraphs, as text. */
const cells = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll('#cj-host td')).map((td) =>
    Array.from(td.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))
      .map((p) => (p.textContent ?? '').replace(/ /g, ' ').trim())));

const inCell = (page: Page, cell: number, paragraph: number) =>
  page.locator(`${HOST} td`).nth(cell).locator('[data-anchor][contenteditable="true"]').nth(paragraph);

/** A 1×2 table whose first cell holds an empty paragraph, "Text" and another empty paragraph. */
async function cellWithEmptyParagraphs(page: Page) {
  await mount(page);
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.evaluate(() => (window as any).__ribbon.editor.insertTable(1, 2, { cellContents: ['Text', 'Next'] }));
  await inCell(page, 0, 0).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Enter');
  await inCell(page, 0, 1).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await expect.poll(() => cells(page)).toEqual([['', 'Text', ''], ['Next']]);
}

test('Backspace in an empty paragraph of a cell removes it; one undo brings it back', async ({ page }) => {
  await cellWithEmptyParagraphs(page);
  await inCell(page, 0, 2).click();
  await page.keyboard.press('Backspace');
  await expect.poll(() => cells(page)).toEqual([['', 'Text'], ['Next']]);
  await page.keyboard.press('Control+z');
  await expect.poll(() => cells(page)).toEqual([['', 'Text', ''], ['Next']]);
});

test('Delete in an empty paragraph of a cell removes it', async ({ page }) => {
  await cellWithEmptyParagraphs(page);
  await inCell(page, 0, 0).click();
  await page.keyboard.press('Delete');
  await expect.poll(() => cells(page)).toEqual([['Text', ''], ['Next']]);
});

test('Backspace at the start of a cell and Delete at its end leave the cell boundary alone', async ({ page }) => {
  await cellWithEmptyParagraphs(page);
  await inCell(page, 1, 0).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await inCell(page, 0, 2).click();
  await page.keyboard.press('Delete');
  expect(await cells(page)).toEqual([['', 'Text', ''], ['Next']]);
});

test('Backspace at a cell paragraph\'s start joins it to the one before, in the same cell', async ({ page }) => {
  await cellWithEmptyParagraphs(page);
  await inCell(page, 0, 1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Backspace');
  await expect.poll(() => cells(page)).toEqual([['Text', ''], ['Next']]);
});

test('outside a table, Delete at a paragraph\'s end joins the next paragraph to it', async ({ page }) => {
  await mount(page);
  const blocks = page.locator(`${HOST} [data-anchor][contenteditable="true"]`);
  await blocks.first().click();
  await page.keyboard.type('One');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Two');
  await blocks.first().click();
  await page.keyboard.press('End');
  await page.keyboard.press('Delete');
  const texts = () => page.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>('#cj-host [data-anchor][contenteditable="true"]'))
    .map((b) => (b.textContent ?? '').trim()));
  await expect.poll(texts).toEqual(['One Two']); // (the engine's join, as Backspace's)
  await page.keyboard.press('Control+z');
  await expect.poll(texts).toEqual(['One', 'Two']);
});

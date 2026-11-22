import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { test, expect, type Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * Backspace at a paragraph's start and Delete at its end join two paragraphs, as in Word: in a
 * table cell only paragraphs of the same cell, so an empty paragraph in a cell can be removed while
 * the cell boundary stays. A tab, a line break or a picture is content: beside one, the keys delete
 * it.
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

/** A document of four paragraphs: "Above", a lone tab, "Tail" and a tab, "Next". */
function tabDocument(): Uint8Array {
  const body = '<w:p><w:r><w:t>Above</w:t></w:r></w:p>'
    + '<w:p><w:r><w:tab/></w:r></w:p>'
    + '<w:p><w:r><w:t>Tail</w:t><w:tab/></w:r></w:p>'
    + '<w:p><w:r><w:t>Next</w:t></w:r></w:p>';
  const relationship = (id: string, kind: string, target: string) =>
    `<Relationship Id="${id}" Type="${R_NS}/${kind}" Target="${target}"/>`;
  return storedZip([
    { name: '[Content_Types].xml', data: xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      + '<Default Extension="xml" ContentType="application/xml"/>'
      + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
      + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>') },
    { name: '_rels/.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + relationship('rId1', 'officeDocument', 'word/document.xml') + '</Relationships>') },
    { name: 'word/_rels/document.xml.rels', data: xml('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      + relationship('rId1', 'styles', 'styles.xml') + '</Relationships>') },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"/>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>${body}`
      + '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
      + '</w:body></w:document>') },
  ]);
}

/** Opens `tabDocument` and puts the caret beside the tab of paragraph `index`: before it, or after it. */
async function caretBesideTab(page: Page, index: number, side: 'before' | 'after') {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(({ bytes, index, side }) => {
    const w = window as any;
    const container = document.createElement('div');
    container.id = 'cj-host';
    document.body.appendChild(container);
    w.editor = w.Docxodus.DocxEditor.open(container, new Uint8Array(bytes), w.Docxodus, {});
    const block = Array.from(container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'))[index];
    const tab = block.querySelector('[data-docx-tab]')!;
    block.focus();
    const at = Array.prototype.indexOf.call(tab.parentNode!.childNodes, tab) + (side === 'after' ? 1 : 0);
    window.getSelection()!.collapse(tab.parentNode, at);
  }, { bytes: Array.from(tabDocument()), index, side });
}

/** Each paragraph's run content in the saved package: text as itself, a tab as `|`. */
const savedShapes = (page: Page) => page.evaluate((ns) => {
  const w = window as any;
  (document.activeElement as HTMLElement | null)?.blur();
  const session = w.Docxodus.openTypedSession(w.editor.save());
  try {
    return Object.entries(session.project().anchorIndex as Record<string, { kind: string }>)
      .filter(([, a]) => a.kind === 'p').map(([id]) => {
        const doc = new DOMParser().parseFromString(session.raw.getXml(id), 'application/xml');
        return Array.from(doc.getElementsByTagNameNS(ns, 'r')).flatMap((r) => Array.from(r.children))
          .map((c) => c.localName === 't' ? c.textContent : c.localName === 'tab' ? '|' : '').join('');
      });
  } finally { session.close(); }
}, W_NS);

test('Backspace after the only tab of a paragraph deletes the tab, not the paragraph mark before it', async ({ page }) => {
  await caretBesideTab(page, 1, 'after');
  await page.keyboard.press('Backspace');
  expect(await savedShapes(page)).toEqual(['Above', '', 'Tail|', 'Next']);
});

test('Delete before a tab that ends a paragraph deletes the tab, not the paragraph mark after it', async ({ page }) => {
  await caretBesideTab(page, 2, 'before');
  await page.keyboard.press('Delete');
  expect(await savedShapes(page)).toEqual(['Above', '|', 'Tail', 'Next']);
});

const PICTURE = fs.readFileSync(path.join(
  path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'TestFiles', 'img.png')).toString('base64');

/** Each paragraph as text, prefixed `[img]` when it holds a picture (wherever in it). */
const withPictures = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#cj-host [data-anchor][contenteditable="true"]'))
    .map((p) => (p.querySelector('img') ? '[img]' : '') + (p.textContent ?? '').replace(/[ ‎‏]/g, '').trim()));

/** Puts the caret in paragraph `index` of the host, just before or just after its picture. */
const caretBesidePicture = (page: Page, index: number, side: 'before' | 'after') =>
  page.evaluate(({ index, side }) => {
    const paragraph = Array.from(document.querySelectorAll<HTMLElement>('#cj-host [data-anchor][contenteditable="true"]'))[index];
    paragraph.focus();
    const picture = paragraph.querySelector('img')!;
    const at = Array.prototype.indexOf.call(picture.parentNode!.childNodes, picture) + (side === 'after' ? 1 : 0);
    window.getSelection()!.collapse(picture.parentNode, at);
  }, { index, side });

/** Body paragraphs "One" and "Two", with a picture at the end of "One", or at its start. */
async function onePictureTwo(page: Page, pictureAt: 'end' | 'start') {
  await mount(page);
  const blocks = page.locator(`${HOST} [data-anchor][contenteditable="true"]`);
  await blocks.first().click();
  await page.keyboard.type('One');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Two');
  await blocks.first().click();
  await page.keyboard.press(pictureAt === 'end' ? 'End' : 'Home');
  await page.evaluate((b64) => (window as any).__ribbon.editor.insertImage(b64, { widthPoints: 40, heightPoints: 40 }), PICTURE);
  await expect.poll(() => withPictures(page)).toEqual(['[img]One', 'Two']);
}

/** A cell holding "Top" and a second paragraph "Low" that starts with a picture. */
async function cellWithPictureStartingLow(page: Page) {
  await mount(page);
  await page.locator(`${HOST} [data-anchor][contenteditable="true"]`).first().click();
  await page.evaluate(() => (window as any).__ribbon.editor.insertTable(1, 2, { cellContents: ['Top', 'Next'] }));
  await inCell(page, 0, 0).click();
  await page.keyboard.press('End');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Low');
  await page.keyboard.press('Home');
  await page.evaluate((b64) => (window as any).__ribbon.editor.insertImage(b64, { widthPoints: 40, heightPoints: 40 }), PICTURE);
  await expect.poll(() => cells(page)).toEqual([['Top', 'Low'], ['Next']]);
  await expect.poll(() => page.locator(`${HOST} td img`).count()).toBe(1);
}

test('Delete before a picture that ends a paragraph deletes the picture, not the paragraph mark after it', async ({ page }) => {
  await onePictureTwo(page, 'end');
  await caretBesidePicture(page, 0, 'before');
  await page.keyboard.press('Delete');
  await expect.poll(() => withPictures(page)).toEqual(['One', 'Two']);
});

test('Backspace after a picture that starts a cell paragraph deletes the picture, not the break before it', async ({ page }) => {
  await cellWithPictureStartingLow(page);
  await caretBesidePicture(page, 1, 'after');
  await page.keyboard.press('Backspace');
  await expect.poll(() => cells(page)).toEqual([['Top', 'Low'], ['Next']]);
  await expect.poll(() => page.locator(`${HOST} img`).count()).toBe(0);
});

test('with the picture on the other side of the caret, the keys still join the paragraphs', async ({ page }) => {
  await onePictureTwo(page, 'start');
  await caretBesidePicture(page, 0, 'after');
  await page.keyboard.press('End');
  await page.keyboard.press('Delete');
  await expect.poll(() => withPictures(page)).toEqual(['[img]One Two']);

  await cellWithPictureStartingLow(page);
  await caretBesidePicture(page, 1, 'before');
  await page.keyboard.press('Backspace');
  await expect.poll(() => cells(page)).toEqual([['Top Low'], ['Next']]);
  await expect.poll(() => page.locator(`${HOST} img`).count()).toBe(1);
});

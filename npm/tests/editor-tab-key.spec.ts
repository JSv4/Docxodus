import { test, expect, type Page } from '@playwright/test';
import { W_NS } from './docx-zip.js';

// The Tab key types a tab, as in Word: a real w:tab at the caret (or over the selection), not a
// move of the focus out of the document. In a list item, Tab at its start still changes the level
// and Tab inside its text types a tab; in a table cell it still moves between cells.

/** Each body paragraph's run content in the editor's saved package: text as itself, a tab as `|`. */
const SHAPES = `(bytes) => {
  const session = window.Docxodus.openTypedSession(bytes);
  try {
    return Object.entries(session.project().anchorIndex).filter(([, a]) => a.kind === 'p' || a.kind === 'li').map(([id]) => {
      const doc = new DOMParser().parseFromString(session.raw.getXml(id), 'application/xml');
      return Array.from(doc.getElementsByTagNameNS('${W_NS}', 'r')).flatMap((r) => Array.from(r.children))
        .map((c) => c.localName === 't' ? c.textContent : c.localName === 'tab' ? '|' : '').join('');
    });
  } finally { session.close(); }
}`;

const blocks = (page: Page) => page.locator('#tk-host [data-anchor][contenteditable="true"]:not(#pagination-staging *)');

async function mount(page: Page, paginated: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.evaluate(() => {
    const host = document.createElement('div');
    host.id = 'tk-host';
    host.style.height = '700px';
    document.body.appendChild(host);
    const ribbon = (window as any).DocxodusEditor.mountRibbon(host, {});
    ribbon.setExports((window as any).Docxodus);
    ribbon.openBlank('tab.docx');
    ribbon.loader.done();
    (window as any).__tkRibbon = ribbon;
  });
  if (paginated) {
    await page.evaluate(() => (window as any).__tkRibbon.editor.setPaginated(true));
    await expect(page.locator('#tk-host .page-box [data-anchor][contenteditable="true"]').first()).toBeVisible();
  }
  await blocks(page).first().click();
}

async function saved(page: Page): Promise<string[]> {
  await page.evaluate(() => (document.activeElement as HTMLElement | null)?.blur());
  return page.evaluate((shapes) => (eval(shapes) as (b: Uint8Array) => string[])(
    (window as any).__tkRibbon.editor.save()), SHAPES);
}

/** Each paragraph as the page shows it: text as itself, a rendered tab as `|`. */
const shown = (page: Page) => page.evaluate(() =>
  Array.from(document.querySelectorAll<HTMLElement>('#tk-host [data-anchor][contenteditable="true"]')).map((b) => {
    let out = '';
    const walk = (n: Node): void => {
      if (n.nodeType === 1 && (n as HTMLElement).hasAttribute('data-docx-tab')) { out += '|'; return; }
      if (n.nodeType === 3) { out += n.nodeValue ?? ''; return; }
      n.childNodes.forEach(walk);
    };
    walk(b);
    return out.replace(/[ ‎‏]/g, '').trim();
  }));

const focusInDocument = (page: Page) =>
  page.evaluate(() => !!document.activeElement?.closest('#tk-host [data-anchor]'));

for (const paginated of [false, true]) {
  test(`Tab types a tab at the caret, and typing goes on after it (paginated=${paginated})`, async ({ page }) => {
    await mount(page, paginated);
    await page.keyboard.type('ab');
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('Tab');
    expect(await focusInDocument(page)).toBe(true);
    await page.keyboard.type('c');
    expect(await shown(page)).toEqual(['a|cb']);
    if (!paginated) expect(await saved(page)).toEqual(['a|cb']);
  });

  test(`Tab in an empty paragraph, then text (paginated=${paginated})`, async ({ page }) => {
    await mount(page, paginated);
    await page.keyboard.press('Tab');
    expect(await focusInDocument(page)).toBe(true);
    await page.keyboard.type('Name');
    expect(await shown(page)).toEqual(['|Name']);
    if (!paginated) expect(await saved(page)).toEqual(['|Name']);
  });
}

test('Tab over a selection replaces it with a tab; Undo restores the text', async ({ page }) => {
  await mount(page, false);
  await page.keyboard.type('Name Date');
  for (let i = 0; i < 5; i++) await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Tab');
  expect(await saved(page)).toEqual(['Name|Date']);
  await blocks(page).first().click();
  await page.keyboard.press('Control+z');
  expect(await saved(page)).toEqual(['Name Date']);
});

test('Shift+Tab outside a list or table keeps the focus and changes nothing', async ({ page }) => {
  await mount(page, false);
  await page.keyboard.type('ab');
  await page.keyboard.press('Shift+Tab');
  expect(await focusInDocument(page)).toBe(true);
  expect(await saved(page)).toEqual(['ab']);
});

test('in a list item, Tab at its start changes the level; inside its text it types a tab', async ({ page }) => {
  await mount(page, false);
  await page.keyboard.type('One');
  await page.keyboard.press('Enter');
  await page.keyboard.type('Two');
  for (const i of [0, 1]) {
    await blocks(page).nth(i).click();
    await page.evaluate(() => (window as any).__tkRibbon.editor.toggleList('decimal'));
  }
  const ilvl = () => page.evaluate((ns) => {
    const ed = (window as any).__tkRibbon.editor;
    const session = (window as any).Docxodus.openTypedSession(ed.save());
    try {
      return Object.entries(session.project().anchorIndex).filter(([, a]: any) => a.kind === 'p' || a.kind === 'li').map(([id]) => {
        const doc = new DOMParser().parseFromString(session.raw.getXml(id), 'application/xml');
        return doc.getElementsByTagNameNS(ns, 'ilvl')[0]?.getAttribute('w:val') ?? '-';
      });
    } finally { session.close(); }
  }, W_NS);
  await expect.poll(ilvl).toEqual(['0', '0']);
  await blocks(page).nth(1).click();
  await page.keyboard.press('Home');
  await page.keyboard.press('Tab');
  await expect.poll(ilvl).toEqual(['0', '1']);
  await blocks(page).nth(1).click();
  await page.keyboard.press('End');
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('Tab');
  expect(await saved(page)).toEqual(['One', 'Tw|o']);
});

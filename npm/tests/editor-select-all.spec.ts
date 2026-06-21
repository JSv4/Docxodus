import { test, expect, Page } from '@playwright/test';

// Context-aware Select-All (complex-filing smoke-test #2 + #6). Native Ctrl+A mis-selects the first
// table once a contenteditable=false table island exists, so a global format is impossible. The
// editor intercepts Ctrl+A: in the BODY it selects all body blocks (tables excluded, a v1 boundary);
// in a TABLE it selects that table's cells so one font applies to the whole table.

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

function ctrlA(): KeyboardEvent {
  return new KeyboardEvent('keydown', { key: 'a', ctrlKey: true, bubbles: true, cancelable: true });
}

test.describe('DocxEditor — context-aware Select-All', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Ctrl+A in the body fonts all body blocks (across a table), not the cells', async ({ page }) => {
    const r = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});

      const bodyBlocks = () =>
        Array.from(container.querySelectorAll('[data-anchor][data-editable="1"]'))
          .filter((b) => !(b as HTMLElement).closest('table')) as HTMLElement[];
      const cells = () =>
        Array.from(container.querySelectorAll('table [data-anchor][data-editable="1"]')) as HTMLElement[];
      const fontOf = (el: HTMLElement) =>
        getComputedStyle((el.querySelector('span') as HTMLElement) ?? el).fontFamily;

      // Build: "Alpha" para, then a table below it, then "Beta" in the trailing para.
      const b0 = bodyBlocks()[0];
      b0.focus();
      b0.textContent = 'Alpha';
      b0.dispatchEvent(new Event('blur'));
      bodyBlocks()[0].focus();
      editor.insertTable(1, 2, { borderless: true });
      const trailing = bodyBlocks().find((b) => (b.textContent || '').trim().length === 0)!;
      trailing.focus();
      trailing.textContent = 'Beta';
      trailing.dispatchEvent(new Event('blur'));

      // Ctrl+A from a body block, then apply a font.
      const start = bodyBlocks().find((b) => (b.textContent || '').includes('Alpha'))!;
      start.focus();
      container.dispatchEvent(ctrlA_());
      editor.setFontFamily('Arial');

      const bodyFonts = bodyBlocks()
        .filter((b) => (b.textContent || '').trim().length > 0)
        .map(fontOf);
      const cellFonts = cells().map(fontOf);
      editor.close();
      return { bodyFonts, cellFonts };

      function ctrlA_() {
        return new KeyboardEvent('keydown', { key: 'a', ctrlKey: true, bubbles: true, cancelable: true });
      }
    });
    // Both body paragraphs got Arial (select-all spanned them across the table)…
    expect(r.bodyFonts.length).toBeGreaterThanOrEqual(2);
    expect(r.bodyFonts.every((f) => /Arial/i.test(f))).toBe(true);
    // …and the table cells did NOT (tables are a body-selection boundary).
    expect(r.cellFonts.every((f) => !/Arial/i.test(f))).toBe(true);
  });

  test('Ctrl+A in a table cell fonts every cell of that table', async ({ page }) => {
    const cellFonts = await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});

      const cells = () =>
        Array.from(container.querySelectorAll('table [data-anchor][data-editable="1"]')) as HTMLElement[];
      const fontOf = (el: HTMLElement) =>
        getComputedStyle((el.querySelector('span') as HTMLElement) ?? el).fontFamily;

      const first = container.querySelector('[data-anchor][data-editable="1"]') as HTMLElement;
      first.focus();
      editor.insertTable(1, 2, { borderless: true });
      // Fill both cells (commit each on blur).
      const cs = cells();
      cs.forEach((c, i) => { c.focus(); c.textContent = 'C' + i; c.dispatchEvent(new Event('blur')); });

      // Ctrl+A from inside a cell selects that table's cells; one setFontFamily fonts them all.
      cells()[0].focus();
      container.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', ctrlKey: true, bubbles: true, cancelable: true }));
      editor.setFontFamily('Times New Roman');

      const fonts = cells().map(fontOf);
      editor.close();
      return fonts;
    });
    expect(cellFonts.length).toBe(2);
    expect(cellFonts.every((f) => /Times New Roman/i.test(f))).toBe(true);
  });
});

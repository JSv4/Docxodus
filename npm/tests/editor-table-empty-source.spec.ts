import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// S-1 smoke-test findings 2 + 4: inserting a table from an empty paragraph left that empty
// paragraph stranded ABOVE the table (finding 4), and a table at the end of the body had no
// paragraph after it, so there was nowhere to type below it (finding 2). After the fixes,
// inserting a table on an empty block puts the table first and a (reachable) paragraph below.
test.describe('DocxEditor — table on an empty paragraph', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('no stray empty paragraph above the table, and an editable paragraph below it', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const container = document.createElement('div');
      document.body.appendChild(container);
      const D = (window as any).Docxodus;
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      // Active block is the single empty body paragraph.
      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertTable(2, 3, { borderless: true });

      const tbl = container.querySelector('table') as HTMLElement;
      // Block paragraphs that are NOT inside the table.
      const outerParas = Array.from(container.querySelectorAll('p[data-anchor]')).filter(
        (p) => !tbl.contains(p),
      ) as HTMLElement[];
      const before = outerParas.filter(
        (p) => (tbl.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_PRECEDING) !== 0,
      ).length;
      const after = outerParas.filter(
        (p) => (tbl.compareDocumentPosition(p) & Node.DOCUMENT_POSITION_FOLLOWING) !== 0,
      ).length;

      editor.close();
      container.remove();
      return { hasTable: !!tbl, before, after };
    });

    expect(out.hasTable).toBe(true);
    expect(out.before).toBe(0);             // finding 4: no stray empty line above the table
    expect(out.after).toBeGreaterThanOrEqual(1); // finding 2: a paragraph exists below the table
  });
});

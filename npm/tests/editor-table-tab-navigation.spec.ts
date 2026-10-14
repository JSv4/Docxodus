import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('DocxEditor — table Tab navigation', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('Tab and Shift+Tab move by cell, preserve edits, and append a final row', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(String(error)));

    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'table-tab-editor';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      (window as any).__tableTab = { container, editor, D };

      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      editor.insertTable(2, 2, {
        borderless: true,
        cellContents: ['one', 'two', 'three', 'four'],
      });
    });

    const focusCell = async (index: number, atEnd = false) => {
      await page.evaluate(({ index, atEnd }) => {
        const container = (window as any).__tableTab.container as HTMLElement;
        const table = container.querySelector('table') as HTMLTableElement;
        const cells = Array.from(table.querySelectorAll<HTMLElement>('td, th'))
          .filter((cell) => cell.closest('table') === table);
        const block = Array.from(
          cells[index].querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
        ).find((candidate) => candidate.closest('td, th') === cells[index])!;
        block.focus();
        const range = document.createRange();
        range.selectNodeContents(block);
        range.collapse(!atEnd);
        const selection = getSelection()!;
        selection.removeAllRanges();
        selection.addRange(range);
      }, { index, atEnd });
    };

    const state = () => page.evaluate(() => {
      const container = (window as any).__tableTab.container as HTMLElement;
      const active = document.activeElement as HTMLElement | null;
      const block = active?.closest<HTMLElement>('[data-anchor][contenteditable="true"]') ?? null;
      const cell = block?.closest<HTMLElement>('td, th') ?? null;
      const table = cell?.closest<HTMLTableElement>('table') ?? null;
      const cells = table
        ? Array.from(table.querySelectorAll<HTMLElement>('td, th'))
            .filter((candidate) => candidate.closest('table') === table)
        : [];
      const selection = getSelection();
      let caretAtStart = false;
      if (block && selection?.rangeCount && selection.isCollapsed) {
        const before = document.createRange();
        before.selectNodeContents(block);
        before.setEnd(selection.anchorNode!, selection.anchorOffset);
        caretAtStart = before.toString().length === 0;
      }
      return {
        cellIndex: cell ? cells.indexOf(cell) : -1,
        cellCount: cells.length,
        rowCount: table?.querySelectorAll('tr').length ?? 0,
        caretAtStart,
      };
    });

    // A dirty edit commits during the focus move, and Tab lands at the start of the next cell.
    await focusCell(0, true);
    await page.keyboard.type('!');
    await page.keyboard.press('Tab');
    expect(await state()).toEqual({ cellIndex: 1, cellCount: 4, rowCount: 2, caretAtStart: true });

    // Shift+Tab reverses the move and clamps at the first cell instead of escaping the table.
    await page.keyboard.press('Shift+Tab');
    expect(await state()).toEqual({ cellIndex: 0, cellCount: 4, rowCount: 2, caretAtStart: true });
    await page.keyboard.press('Shift+Tab');
    expect(await state()).toEqual({ cellIndex: 0, cellCount: 4, rowCount: 2, caretAtStart: true });

    // Like Word, Tab from the final cell appends a matching row and enters its first cell.
    await focusCell(3, true);
    await page.keyboard.type('!');
    await page.keyboard.press('Tab');
    expect(await state()).toEqual({ cellIndex: 4, cellCount: 6, rowCount: 3, caretAtStart: true });
    await page.keyboard.press('Shift+Tab');
    expect(await state()).toEqual({ cellIndex: 3, cellCount: 6, rowCount: 3, caretAtStart: true });

    const persisted = await page.evaluate(() => {
      const { container, editor, D } = (window as any).__tableTab;
      (document.activeElement as HTMLElement | null)?.blur();
      const saved: Uint8Array = editor.save();
      const reopenedContainer = document.createElement('div');
      document.body.appendChild(reopenedContainer);
      const reopened = D.DocxEditor.open(reopenedContainer, saved, D, {});
      const texts = Array.from(reopenedContainer.querySelectorAll('td, th'))
        .map((cell) => (cell.textContent ?? '').trim());
      const rows = reopenedContainer.querySelectorAll('table tr').length;
      reopened.close();
      reopenedContainer.remove();
      return { texts, rows };
    });

    expect(persisted.texts[0]).toBe('one!');
    expect(persisted.texts[3]).toBe('four!');
    expect(persisted.rows).toBe(3);
    expect(errors, `uncaught errors:\n${errors.join('\n')}`).toEqual([]);
  });
});

import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

async function openParagraphDocument(page: Page, names: string[], options: Record<string, unknown> = {}) {
  await page.evaluate((options) => {
    const D = (window as any).Docxodus;
    const container = document.createElement('div');
    container.id = 'block-drag-host';
    container.style.cssText = 'width:700px;margin:40px auto;padding:32px;background:white';
    document.body.appendChild(container);
    const moves: unknown[] = [];
    const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {
      blockDrag: true,
      onMove: (info: unknown) => moves.push(info),
      ...options,
    });
    (window as any).__drag = { editor, container, moves };
    (container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement).focus();
  }, options);
  for (let i = 0; i < names.length; i++) {
    if (i > 0) await page.keyboard.press('Enter');
    await page.keyboard.type(names[i]);
  }
  await page.evaluate(() => {
    (document.activeElement as HTMLElement)?.blur();
    // Canonical full paint gives every unit a signature and recreates the drag targets.
    (window as any).__drag.editor['remount']();
  });
}

const unitState = (page: Page) => page.evaluate(() => {
  const { editor } = (window as any).__drag;
  return (editor['bodyUnitNodes']() as HTMLElement[]).map((el) => ({
    tag: el.tagName,
    text: (el.textContent ?? '').replace(/\s+/g, ' ').trim(),
    editable: el.getAttribute('contenteditable'),
  }));
});

test.describe('DocxEditor — block drag handle', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('click menu moves a block accessibly and preserves its DOM node', async ({ page }) => {
    await openParagraphDocument(page, ['Alpha', 'Beta', 'Gamma']);
    await page.evaluate(() => {
      const beta = Array.from(document.querySelectorAll<HTMLElement>('#block-drag-host p[data-anchor]'))
        .find((el) => el.textContent?.includes('Beta'))!;
      (beta as any).__identity = 'same-node';
    });

    const beta = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'Beta' });
    await beta.hover();
    const handle = page.locator('.docx-block-handle');
    await expect(handle).toBeVisible();
    await expect(handle).toHaveAttribute('aria-haspopup', 'menu');
    await handle.click();
    await expect(page.getByRole('menuitem', { name: 'Move to top' })).toBeVisible();
    await page.getByRole('menuitem', { name: 'Move to top' }).click();

    expect((await unitState(page)).map((x) => x.text)).toEqual(['Beta', 'Alpha', 'Gamma']);
    const result = await page.evaluate(() => {
      const { moves } = (window as any).__drag;
      const beta = Array.from(document.querySelectorAll<HTMLElement>('#block-drag-host p[data-anchor]'))
        .find((el) => el.textContent?.includes('Beta'))!;
      return { sameNode: (beta as any).__identity, moves: moves.length };
    });
    expect(result).toEqual({ sameNode: 'same-node', moves: 1 });
  });

  test('dragging uses before/after drop zones and reorders the live session', async ({ page }) => {
    await openParagraphDocument(page, ['One', 'Two', 'Three']);
    const one = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'One' });
    const three = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'Three' });
    await one.hover();
    const handle = page.locator('.docx-block-handle');
    const handleBox = await handle.boundingBox();
    const targetBox = await three.boundingBox();
    expect(handleBox).not.toBeNull();
    expect(targetBox).not.toBeNull();
    await handle.dragTo(three, {
      targetPosition: { x: targetBox!.width / 2, y: targetBox!.height - 2 },
    });
    expect((await unitState(page)).map((x) => x.text)).toEqual(['Two', 'Three', 'One']);
  });

  test('a cell hover selects and moves its whole table', async ({ page }) => {
    await openParagraphDocument(page, ['Before', 'After']);
    await page.evaluate(() => {
      const { editor } = (window as any).__drag;
      const first = editor['editableList']()[0] as HTMLElement;
      first.focus();
      editor.insertTable(2, 2);
    });
    const cell = page.locator('#block-drag-host table td p[contenteditable="true"]').first();
    await cell.hover();
    await page.locator('.docx-block-handle').click();
    await page.getByRole('menuitem', { name: 'Move to bottom' }).click();
    const units = await unitState(page);
    expect(units.at(-1)?.tag).toBe('TABLE');
    expect(units.filter((x) => x.tag === 'TABLE')).toHaveLength(1);
  });

  test('review mode renders a native move pair and keeps the source read-only', async ({ page }) => {
    await openParagraphDocument(page, ['North', 'Middle', 'South']);
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const state = (window as any).__drag;
      const saved = state.editor.save();
      state.editor.close();
      state.container.replaceChildren();
      const editor = D.DocxEditor.open(state.container, saved, D, {
        blockDrag: true,
        trackedChanges: 1, // TrackedChangeMode.RenderInline
        revisionAuthor: 'Drag Tester',
      });
      (window as any).__drag.editor = editor;
      const units = editor['bodyUnitNodes']() as HTMLElement[];
      editor.moveBlock(editor['anchorIdOf'](units[0]), editor['anchorIdOf'](units[2]), 'after');
    });
    const review = await page.evaluate(() => {
      const host = document.querySelector('#block-drag-host')!;
      const from = host.querySelector<HTMLElement>("del[class$='move-from'], del[class*='move-from ']");
      const to = host.querySelector<HTMLElement>("ins[class$='move-to'], ins[class*='move-to ']");
      return {
        from: from?.textContent,
        to: to?.textContent,
        sourceEditable: from?.closest('[data-anchor]')?.getAttribute('contenteditable'),
      };
    });
    expect(review.from).toContain('North');
    expect(review.to).toContain('North');
    expect(review.sourceEditable).toBe('false');
  });
});

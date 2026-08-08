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

  // The handle floats in the page margin, so the natural gesture — press it and pull straight
  // down — never crosses a paragraph box. Element hit testing therefore found no drop target for
  // the whole gesture: no drop line, and a release that silently did nothing. Drop position is
  // resolved from the pointer's vertical position against the blocks instead.
  test('a drag down the left gutter shows the drop line and lands', async ({ page }) => {
    const names = ['One', 'Two', 'Three', 'Four', 'Five', 'Six'];
    await openParagraphDocument(page, names);
    await page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'One' }).hover();
    const handle = page.locator('.docx-block-handle');
    await expect(handle).toBeVisible();

    // The custom drag preview is mounted only for the browser's snapshot and torn down again,
    // so it has to be observed rather than queried.
    await page.evaluate(() => {
      (window as any).__previews = [];
      new MutationObserver((records) => {
        for (const r of records)
          for (const node of Array.from(r.addedNodes))
            if (node instanceof HTMLElement)
              (window as any).__previews.push(
                ...Array.from(node.querySelectorAll('.docx-block-drag-preview')).map(
                  (el) => el.textContent,
                ),
              );
      }).observe(document.body, { childList: true, subtree: true });
    });
    const sample = () => page.evaluate(() => {
      const line = document.querySelector<HTMLElement>('.docx-block-drop-indicator')!;
      const rect = line.getBoundingClientRect();
      return {
        shown: getComputedStyle(line).display !== 'none',
        top: Math.round(rect.top),
        width: Math.round(rect.width),
        dimmed: document.querySelector('.docx-block-drag-source')?.textContent?.trim() ?? null,
      };
    });

    const hb = (await handle.boundingBox())!;
    const last = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'Six' });
    const lb = (await last.boundingBox())!;
    const gutterX = hb.x + hb.width / 2;
    expect(gutterX).toBeLessThan(lb.x); // the pointer never enters the text column

    await page.mouse.move(gutterX, hb.y + hb.height / 2);
    await page.mouse.down();
    const endY = lb.y + lb.height - 2;
    const steps = [];
    for (let i = 1; i <= 6; i++) {
      await page.mouse.move(gutterX, hb.y + ((endY - hb.y) * i) / 6, { steps: 4 });
      steps.push(await sample());
    }
    // Chromium dispatches drag events asynchronously from the synthetic mouse move, so the line
    // can trail a step. Nudge until it settles on the edge the release will actually use.
    await expect.poll(async () => {
      await page.mouse.move(gutterX, endY - 2);
      await page.mouse.move(gutterX, endY);
      return (await sample()).top;
    }).toBeGreaterThan(lb.y + lb.height / 2);
    const previews = await page.evaluate(() => (window as any).__previews as string[]);
    await page.mouse.up();

    // The line appears as soon as the pointer clears the block being dragged (over that block
    // there is no drop position to show), and then tracks continuously to the end of the gesture.
    const firstShown = steps.findIndex((s) => s.shown);
    expect(firstShown).toBeGreaterThanOrEqual(0);
    expect(firstShown).toBeLessThanOrEqual(2);
    const shown = steps.slice(firstShown);
    expect(shown.every((s) => s.shown)).toBe(true);
    expect(new Set(shown.map((s) => s.top)).size).toBeGreaterThan(1);
    expect(shown.map((s) => s.width)).toEqual(shown.map(() => Math.round(lb.width)));
    // The block being carried is dimmed for the whole gesture, and the preview names it.
    // (Sampled from `firstShown`: the browser needs a few pixels of travel before it starts a
    // native drag at all, so the opening sample can precede dragstart.)
    expect(shown.map((s) => s.dimmed)).toEqual(shown.map(() => 'One'));
    expect(previews).toContain('One');

    // What the line predicted is where the block lands.
    expect((await unitState(page)).map((x) => x.text)).toEqual([...names.slice(1), 'One']);
    // Nothing is left dimmed once the move lands.
    expect(await page.locator('.docx-block-drag-source').count()).toBe(0);
    await expect(page.locator('.docx-block-drop-indicator')).toBeHidden();
  });

  // A paragraph's w:spacing becomes a CSS margin, which sits OUTSIDE the border box, so the
  // visual gap between two blocks is not where either box edge is. Drawing the drop line on the
  // target's edge underlines that block's last line instead of reading as a boundary — which
  // every DOM-level assertion above is blind to, since the line is shown, is the right width,
  // and tracks the pointer either way. This pins the position itself.
  test('the drop line is drawn in the gap between blocks, not on a block edge', async ({ page }) => {
    await openParagraphDocument(page, ['Alpha', 'Beta', 'Gamma', 'Delta']);
    // Real spacing, written through the session rather than injected as CSS, so the geometry
    // under test is the geometry a document with `w:spacing` actually produces.
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const { editor } = (window as any).__drag;
      for (const unit of editor['bodyUnitNodes']() as HTMLElement[]) {
        const id = editor['anchorIdOf'](unit);
        if (id) D.DocxSessionBridge.SetParagraphFormat(editor.sessionHandle, id,
          JSON.stringify({ spacingBefore: 240, spacingAfter: 240 }));
      }
      editor['remount']();
    });

    // Driven through the drag internals rather than a native drag: this is a geometry assertion,
    // and Chromium's asynchronous drag-event dispatch would only add settle timing to it.
    const geo = await page.evaluate(() => {
      const { editor } = (window as any).__drag;
      const units = editor['bodyUnitNodes']() as HTMLElement[];
      const source = units[0];
      editor['blockDragSource'] = source;
      editor['refreshBlockMoveTargets'](source);
      editor['captureDropZones']();

      const lineTopFor = (y: number): number | null => {
        const hit = editor['resolveDropAt'](y);
        if (!hit) return null;
        editor['paintDropIndicator']({ zone: hit.zone, position: hit.position });
        const line = document.querySelector<HTMLElement>('.docx-block-drop-indicator')!;
        return getComputedStyle(line).display === 'none'
          ? null
          : line.getBoundingClientRect().top;
      };

      const box = (el: HTMLElement) => el.getBoundingClientRect();
      const gamma = box(units[2]);
      const delta = box(units[3]);
      return {
        // Lower half of Gamma ⇒ insert after Gamma, i.e. in the Gamma/Delta gap.
        betweenBlocks: lineTopFor(gamma.bottom - 2),
        gammaBottom: gamma.bottom,
        deltaTop: delta.top,
        // Below the last block ⇒ insert after Delta, where there is no neighbour to bisect.
        endOfFlow: lineTopFor(delta.bottom + 40),
        deltaBottom: delta.bottom,
      };
    });

    // There is a real gap to land in — otherwise the assertions below prove nothing.
    expect(geo.deltaTop - geo.gammaBottom).toBeGreaterThan(6);
    // Strictly inside the gap, on neither block's edge.
    expect(geo.betweenBlocks).toBeGreaterThan(geo.gammaBottom + 1);
    expect(geo.betweenBlocks).toBeLessThan(geo.deltaTop - 1);
    // Past the end of the flow the line still clears the last block's text.
    expect(geo.endOfFlow).toBeGreaterThan(geo.deltaBottom + 1);
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

  // A section break partitions the body into regions a block cannot move between. The engine has
  // always refused those moves; the UI used to draw a drop indicator over them anyway and only
  // fail on release, and "move to top/bottom" always targeted the document ends — so on a document
  // with section breaks those commands could never succeed.
  test('a section break partitions the document into move regions', async ({ page }) => {
    await openParagraphDocument(page, ['Alpha', 'Beta', 'Gamma']);
    await page.evaluate(() => {
      const D = (window as any).Docxodus;
      const { editor } = (window as any).__drag;
      const units = editor['bodyUnitNodes']() as HTMLElement[];
      const afterBeta = editor['anchorIdOf'](units[1]);
      D.DocxSessionBridge.RawInsertXml(editor.sessionHandle, afterBeta, 'after',
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:pPr><w:sectPr/></w:pPr></w:p>');
      editor['remount']();
    });

    // The section-break paragraph itself can never be moved, so it ends up with no handle.
    // Hovering does not ask the engine — that query is document-scale work and hovering happens
    // on every pointer move — so the handle withdraws on the idle callback that follows.
    await page.evaluate(() => {
      const { editor } = (window as any).__drag;
      const units = editor['bodyUnitNodes']() as HTMLElement[];
      const breakUnit = units.find((el) => (el.textContent ?? '').trim() === '')!;
      editor['showBlockHandle'](breakUnit);
    });
    await expect(page.locator('.docx-block-handle')).toBeHidden();

    const state = await page.evaluate(() => {
      const { editor } = (window as any).__drag;
      const units = editor['bodyUnitNodes']() as HTMLElement[];
      // Alpha may reach Beta (same region) but not Gamma (across the break).
      const alpha = units.find((el) => el.textContent?.includes('Alpha'))!;
      editor['refreshBlockMoveTargets'](alpha);
      const targets: Map<string, { before: boolean; after: boolean }> = editor['blockMoveTargets'];
      const idOf = (text: string) =>
        editor['anchorIdOf'](units.find((el: HTMLElement) => el.textContent?.includes(text))!);
      return {
        canReachBeta: targets.has(idOf('Beta')),
        canReachGamma: targets.has(idOf('Gamma')),
      };
    });
    expect(state).toEqual({ canReachBeta: true, canReachGamma: false });

    // Dragging Alpha onto Gamma must not even offer a drop: no indicator, no reorder.
    const alpha = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'Alpha' });
    const gamma = page.locator('#block-drag-host p[data-anchor]').filter({ hasText: 'Gamma' });
    await alpha.hover();
    await page.evaluate(() => {
      const indicator = document.querySelector<HTMLElement>('.docx-block-drop-indicator')!;
      (window as any).__shown = [];
      new MutationObserver(() => {
        if (indicator.style.display === 'block') (window as any).__shown.push(1);
      }).observe(indicator, { attributes: true, attributeFilter: ['style'] });
    });
    const gammaBox = (await gamma.boundingBox())!;
    await page.locator('.docx-block-handle').dragTo(gamma, {
      targetPosition: { x: gammaBox.width / 2, y: gammaBox.height - 2 },
    });
    expect(await page.evaluate(() => (window as any).__shown.length)).toBe(0);
    expect((await unitState(page)).map((x) => x.text).filter(Boolean)).toEqual(['Alpha', 'Beta', 'Gamma']);

    // "Move to bottom" means the end of Alpha's OWN region — Beta — not the document end.
    await alpha.hover();
    await page.locator('.docx-block-handle').click();
    await page.getByRole('menuitem', { name: 'Move to bottom' }).click();
    expect((await unitState(page)).map((x) => x.text).filter(Boolean)).toEqual(['Beta', 'Alpha', 'Gamma']);
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
      const { editor } = (window as any).__drag;
      const host = document.querySelector('#block-drag-host')!;
      const from = host.querySelector<HTMLElement>("del[class$='move-from'], del[class*='move-from ']");
      const to = host.querySelector<HTMLElement>("ins[class$='move-to'], ins[class*='move-to ']");
      return {
        from: from?.textContent,
        to: to?.textContent,
        sourceEditable: from?.closest('[data-anchor]')?.getAttribute('contenteditable'),
        fallback: editor.lastReconcileFallback,
      };
    });
    expect(review.from).toContain('North');
    expect(review.to).toContain('North');
    expect(review.sourceEditable).toBe('false');
    // A tracked move repaints INCREMENTALLY. The move-from source keeps its unid, so only the
    // plan's per-unit content signature makes it diff as changed; if that ever stops holding the
    // op falls back to a whole-document remount, which is seconds on a real document.
    expect(review.fallback).toBeNull();
  });
});

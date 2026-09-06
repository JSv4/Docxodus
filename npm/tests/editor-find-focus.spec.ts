import { test, expect, type Page } from '@playwright/test';

/**
 * Find and replace must not take the keyboard.
 *
 * The regression this pins: the find field re-scanned on every `input` event and then SELECTED the
 * first hit, and selecting inside a contenteditable block focuses that block. So the moment a
 * partial query matched, the caret jumped into the document and the REST OF THE QUERY WAS TYPED
 * INTO THE DOCUMENT — searching for "fox" left an "ox" in the text.
 *
 * The fix keeps the jump-to-hit (the document still scrolls and the match is painted) but expresses
 * it as a highlight rather than a selection, so focus never leaves the search field. The caret is
 * handed the match only when the bar closes.
 *
 * Two layers, both needed:
 *   - the DocxEditor API contract (`showFindMatches` paints, `selectMatch` commits), driven
 *     directly against a session on the bare harness;
 *   - the shipped find bar in `ribbon.ts`, driven with real keystrokes on the editor host.
 */

const SENTENCE = 'The quick brown fox jumps over the lazy fox.';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

/** Names of the registered CSS highlights, so a spec can see what was painted. */
function paintedHighlights(): { active: number; others: number } {
  const registry = (CSS as any).highlights;
  return {
    active: registry?.get('docxodus-find-active')?.size ?? 0,
    others: registry?.get('docxodus-find')?.size ?? 0,
  };
}

test.describe('DocxEditor — find painting does not move focus', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('showFindMatches paints the hits and leaves the keyboard where it was', async ({ page }) => {
    const out = await page.evaluate(async (sentence) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      // A stand-in for the find field: any control OUTSIDE the document surface will do.
      const field = document.createElement('input');
      field.id = 'query';
      document.body.appendChild(field);

      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const first = () => container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      const block = first();
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('insertText', false, sentence);
      block.dispatchEvent(new Event('blur'));

      field.focus();
      field.value = 'fox';
      const matches = editor.find('fox');
      editor.showFindMatches(matches, 0);
      const registry = (CSS as any).highlights;
      const painted = {
        matches: matches.length,
        focusId: (document.activeElement as HTMLElement)?.id ?? '',
        active: registry?.get('docxodus-find-active')?.size ?? 0,
        others: registry?.get('docxodus-find')?.size ?? 0,
      };

      // Riding to the SECOND hit is still focus-free.
      editor.showFindMatches(matches, 1);
      const stepped = { focusId: (document.activeElement as HTMLElement)?.id ?? '' };

      editor.clearFindMatches();
      const cleared = {
        active: registry?.get('docxodus-find-active')?.size ?? 0,
        others: registry?.get('docxodus-find')?.size ?? 0,
      };

      const text = (first().textContent ?? '').trim();
      editor.close();
      container.remove();
      field.remove();
      return { painted, stepped, cleared, text };
    }, SENTENCE);

    // Both hits found, one painted as active and the other as a plain match.
    expect(out.painted.matches).toBe(2);
    expect(out.painted.active).toBe(1);
    expect(out.painted.others).toBe(1);
    // The whole point: the search field still owns the keyboard.
    expect(out.painted.focusId).toBe('query');
    expect(out.stepped.focusId).toBe('query');
    // Clearing takes the painting away again.
    expect(out.cleared).toEqual({ active: 0, others: 0 });
    // And nothing was typed into the document along the way.
    expect(out.text).toBe(SENTENCE);
  });

  test('selectMatch still commits the caret onto the hit', async ({ page }) => {
    const out = await page.evaluate(async (sentence) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const field = document.createElement('input');
      field.id = 'query';
      document.body.appendChild(field);

      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const first = () => container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      const block = first();
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('insertText', false, sentence);
      block.dispatchEvent(new Event('blur'));

      field.focus();
      const matches = editor.find('fox');
      editor.showFindMatches(matches, 0);
      editor.selectMatch(matches[0]);
      const registry = (CSS as any).highlights;
      const after = {
        selected: (window.getSelection()?.toString() ?? ''),
        inDocument: first().contains(document.activeElement) || first() === document.activeElement,
        // Committing to the caret drops the painting — one visible "current match", not two.
        active: registry?.get('docxodus-find-active')?.size ?? 0,
      };

      editor.close();
      container.remove();
      field.remove();
      return after;
    }, SENTENCE);

    expect(out.selected).toBe('fox');
    expect(out.inDocument).toBe(true);
    expect(out.active).toBe(0);
  });

  test('replaceMatch with focus:false edits the text without pulling the caret in', async ({ page }) => {
    const out = await page.evaluate(async (sentence) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const field = document.createElement('input');
      field.id = 'query';
      document.body.appendChild(field);

      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const first = () => container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      const block = first();
      block.focus();
      const range = document.createRange();
      range.selectNodeContents(block);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('insertText', false, sentence);
      block.dispatchEvent(new Event('blur'));

      field.focus();
      const replaced = editor.replaceMatch(editor.find('fox')[0], 'cat', { focus: false });
      const after = {
        replaced,
        focusId: (document.activeElement as HTMLElement)?.id ?? '',
        text: (first().textContent ?? '').trim(),
      };

      editor.close();
      container.remove();
      field.remove();
      return after;
    }, SENTENCE);

    expect(out.replaced).toBe(true);
    expect(out.text).toBe('The quick brown cat jumps over the lazy fox.');
    expect(out.focusId).toBe('query');
  });
});

test.describe('ribbon find bar — typing a query keeps typing in the query', () => {
  /** New blank document, one paragraph of known text, committed. */
  async function seedDocument(page: Page) {
    await page.goto('/editor.html');
    await page.waitForFunction(() => !!(window as any).__demo, { timeout: 60000 });
    await page.click('[data-dxr="new"]');
    await page.waitForFunction(() => !!(window as any).__demo.getEditor());

    const block = page.locator('[data-dxr-surface] [data-anchor][contenteditable="true"]').first();
    await block.click();
    await page.evaluate(() => {
      const el = document.querySelector(
        '[data-dxr-surface] [data-anchor][contenteditable="true"]',
      ) as HTMLElement;
      const range = document.createRange();
      range.selectNodeContents(el);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(range);
    });
    await page.keyboard.type(SENTENCE);
  }

  /** The document's visible text, whitespace-normalised. */
  function surfaceText(): string {
    const surface = document.querySelector('[data-dxr-surface]') as HTMLElement;
    return (surface.textContent ?? '').replace(/\s+/g, ' ').trim();
  }

  test('every character of the query reaches the search field, not the document', async ({ page }) => {
    await seedDocument(page);
    await page.click('[data-dxr="findtoggle"]');
    await expect(page.locator('[data-dxr="findbar"]')).toBeVisible();

    const before = await page.evaluate(surfaceText);
    expect(before).toContain(SENTENCE);

    // Character by character. "f" already matches, which is exactly when the old code jumped the
    // caret into the document and swallowed the "ox".
    await page.keyboard.type('fox', { delay: 30 });

    await expect(page.locator('[data-dxr="findtext"]')).toHaveValue('fox');
    await expect(page.locator('[data-dxr="findcount"]')).toHaveText('1 of 2');
    expect(await page.evaluate(surfaceText)).toBe(before);
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-dxr'))).toBe('findtext');

    // The hit is shown by painting it, which is what makes the focus-free jump possible.
    expect(await page.evaluate(paintedHighlights)).toEqual({ active: 1, others: 1 });
  });

  test('Enter and the Next button step matches without ending the query', async ({ page }) => {
    await seedDocument(page);
    await page.click('[data-dxr="findtoggle"]');
    await page.keyboard.type('fox');

    await page.keyboard.press('Enter');
    await expect(page.locator('[data-dxr="findcount"]')).toHaveText('2 of 2');
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-dxr'))).toBe('findtext');

    await page.click('[data-dxr="findnext"]');
    await expect(page.locator('[data-dxr="findcount"]')).toHaveText('1 of 2');
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-dxr'))).toBe('findtext');

    // Typing continues to refine the same query rather than editing the document.
    await page.keyboard.type('y');
    await expect(page.locator('[data-dxr="findtext"]')).toHaveValue('foxy');
    await expect(page.locator('[data-dxr="findcount"]')).toHaveText('0 of 0');
  });

  test('closing the bar hands the current match to the caret', async ({ page }) => {
    await seedDocument(page);
    await page.click('[data-dxr="findtoggle"]');
    await page.keyboard.type('lazy');
    await page.keyboard.press('Escape');

    await expect(page.locator('[data-dxr="findbar"]')).toBeHidden();
    const landed = await page.evaluate(() => {
      const surface = document.querySelector('[data-dxr-surface]') as HTMLElement;
      return {
        selection: window.getSelection()?.toString() ?? '',
        inDocument: surface.contains(document.activeElement),
        painted: (CSS as any).highlights?.get('docxodus-find-active')?.size ?? 0,
      };
    });
    // Jumped to the result, as asked — but only once the search itself is over.
    expect(landed.selection).toBe('lazy');
    expect(landed.inDocument).toBe(true);
    expect(landed.painted).toBe(0);
  });

  test('Replace edits the document and leaves the keyboard in the replace field', async ({ page }) => {
    await seedDocument(page);
    await page.click('[data-dxr="replacetoggle"]');
    await page.keyboard.type('quick');
    await page.click('[data-dxr="replacetext"]');
    await page.keyboard.type('nimble');
    await page.click('[data-dxr="replaceone"]');

    await expect
      .poll(async () => page.evaluate(surfaceText))
      .toContain('The nimble brown fox');
    expect(await page.evaluate(() => document.activeElement?.getAttribute('data-dxr'))).toBe('replacetext');
    await expect(page.locator('[data-dxr="replacetext"]')).toHaveValue('nimble');
  });
});

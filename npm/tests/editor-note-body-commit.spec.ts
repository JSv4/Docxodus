import { test, expect, Page } from '@playwright/test';

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// #583 — editing a footnote's body must commit like any block edit.
//
// Repro found by playing DOCX GOLF's footnote hole by hand: type into the
// note body at the bottom of the document, click away, and the session still
// holds the old note text; the next interaction reverts the DOM, silently
// discarding the user's edit. Body-paragraph edits through the identical
// gestures commit fine.
test.describe('DocxEditor — footnote body edits commit', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('typing into a footnote body and blurring reaches the session', async ({ page }) => {
    const out = await page.evaluate(async () => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      document.body.appendChild(container);
      const blank: Uint8Array = D.DocxSessionBridge.CreateBlankDocx();
      const editor = D.DocxEditor.open(container, blank, D, {});

      // Seed a body paragraph, then a footnote through the editor's own command
      // (the ribbon's Footnote button calls this).
      const body = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      body.focus();
      const sel = window.getSelection()!;
      let r = document.createRange();
      r.selectNodeContents(body);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'The Lender may rely on the opinions.');
      body.dispatchEvent(new Event('blur'));

      const fresh = container.querySelector('p[data-anchor][contenteditable="true"]') as HTMLElement;
      fresh.focus();
      r = document.createRange();
      r.selectNodeContents(fresh); r.collapse(false);
      sel.removeAllRanges(); sel.addRange(r);
      editor.insertFootnote();
      await new Promise((res) => setTimeout(res, 300));

      // Find the rendered note body and retype its text — the human gesture.
      const blocks = Array.from(
        container.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'),
      );
      const note = blocks.find((b) => (b.textContent || '').includes('New footnote'));
      if (!note) {
        editor.close(); container.remove();
        return { err: 'note body not rendered as an editable block' };
      }
      note.focus();
      const walker = document.createTreeWalker(note, NodeFilter.SHOW_TEXT);
      let target: Text | null = null;
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if ((n as Text).data.includes('New footnote')) { target = n as Text; break; }
      }
      if (!target) {
        editor.close(); container.remove();
        return { err: 'note text node not found' };
      }
      r = document.createRange();
      r.setStart(target, 0); r.setEnd(target, target.data.length);
      sel.removeAllRanges(); sel.addRange(r);
      document.execCommand('insertText', false, 'As defined in the Master Agreement.');
      note.dispatchEvent(new Event('blur'));
      await new Promise((res) => setTimeout(res, 300));

      // Session truth: what does the footnote hold now?
      const proj = JSON.parse(D.DocxSessionBridge.Project(editor.sessionHandle));
      const fn = Object.entries(proj.anchorIndex as Record<string, any>)
        .filter(([id]) => id.startsWith('p:fn:'))
        .map(([, t]) => t.textPreview as string);
      const domNow = (container.querySelector('[data-anchor][contenteditable="true"]:last-child') as HTMLElement)?.textContent;
      editor.close(); container.remove();
      return { fn, domNow };
    });

    expect(out.err, out.err).toBeUndefined();
    expect(
      out.fn!.some((t: string) => t.includes('As defined in the Master Agreement')),
      `the typed note text must reach the session; footnotes held: ${JSON.stringify(out.fn)}`,
    ).toBe(true);
  });
});

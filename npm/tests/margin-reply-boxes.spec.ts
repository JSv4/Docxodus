import { test, expect, type Page } from '@playwright/test';
import { commentThreadFixture } from './docx-comment-thread-fixture.js';

// A margin reply nests inside its parent's note. It must read as part of that one note: no
// second fill, border or shadow, and a body as wide as the root note's body at every depth.
async function measureNotes(page: Page, threaded: boolean, paginated: boolean) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  return page.evaluate(({ bytes, paginated }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Cedar thread', 'docx-', false, '', 2, 'comment-',
      paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, {
      showPageNumbers: false, commentColumnPlacement: 'margin',
    });
    else host.innerHTML = parsed.body.innerHTML;
    const scope = paginated ? '.page-comment-margin ' : '';
    return Array.from(host.querySelectorAll<HTMLElement>(`${scope}.comment-margin-note`)).map(note => {
      const style = getComputedStyle(note);
      const body = note.querySelector<HTMLElement>(':scope > .comment-margin-note-body')!;
      let depth = 0;
      for (let p = note.parentElement?.closest('.comment-margin-note'); p; p = p.parentElement?.closest('.comment-margin-note')) depth++;
      return { depth, bodyWidth: body.getBoundingClientRect().width, borderLeft: style.borderLeftWidth,
        background: style.backgroundColor, shadow: style.boxShadow };
    });
  }, { bytes: Array.from(commentThreadFixture(threaded, false)), paginated });
}

for (const paginated of [false, true]) {
  test(`replies keep the root note's text width and draw no box (paginated=${paginated})`, async ({ page }) => {
    const notes = await measureNotes(page, true, paginated);
    expect(notes.map(n => n.depth)).toEqual([0, 1, 2]);
    const [root, ...replies] = notes;
    expect(root.borderLeft).toBe('3px');
    for (const reply of replies) {
      expect(reply.bodyWidth).toBeCloseTo(root.bodyWidth, 0);
      expect(reply.borderLeft).toBe('0px');
      expect(reply.background).toBe('rgba(0, 0, 0, 0)');
      expect(reply.shadow).toBe('none');
    }
  });

  test(`flat notes each keep the full note box (paginated=${paginated})`, async ({ page }) => {
    const notes = await measureNotes(page, false, paginated);
    expect(notes.map(n => n.depth)).toEqual([0, 0, 0]);
    for (const note of notes) {
      expect(note.borderLeft).toBe('3px');
      expect(note.background).toBe('rgb(255, 249, 196)');
      expect(note.shadow).not.toBe('none');
      expect(note.bodyWidth).toBeCloseTo(notes[0].bodyWidth, 0);
    }
  });
}

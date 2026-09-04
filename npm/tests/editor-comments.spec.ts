import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  return Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

/**
 * Word-style comments in the editor: the commented range is highlighted inline and every
 * thread is a bubble in the gutter beside the sheet, positioned at its highlight, with
 * reply / resolve / edit / delete acting on the session's native comment family.
 *
 * Driven through the bare `DocxEditor` on the test harness (no ribbon), so what is pinned is
 * the engine + gutter contract every host gets.
 */

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

async function installHelpers(page: Page) {
  await page.evaluate(() => {
    const D = (window as any).Docxodus;
    (window as any).__cmMount = (bytes: Uint8Array, opts: object) => {
      const container = document.createElement('div');
      container.style.cssText = 'position:relative; width:1100px; padding-right:280px;';
      document.body.appendChild(container);
      return { container, editor: D.DocxEditor.open(container, bytes, D, opts) };
    };
    /** Select `length` content characters from offset `from`, walking every text node of the
     *  block (a highlighted range splits the text into several) and skipping engine chrome. */
    (window as any).__cmSelect = (el: HTMLElement, from: number, length: number) => {
      el.focus();
      const points: Array<{ node: Text; offset: number }> = [];
      const walk = (n: Node): void => {
        if (n.nodeType === 3) {
          if ((n.parentElement as HTMLElement | null)?.closest('a.comment-marker, [data-list-marker]')) return;
          const text = n.textContent || '';
          for (let i = 0; i <= text.length; i++) points.push({ node: n as Text, offset: i });
          return;
        }
        for (const c of Array.from(n.childNodes)) walk(c);
      };
      walk(el);
      // Consecutive text nodes share a boundary point; collapse duplicates by content offset.
      const byOffset: Array<{ node: Text; offset: number }> = [];
      let last: Text | null = null;
      for (const p of points) {
        if (p.offset === 0 && last !== null) continue; // boundary already counted at the previous node's end
        byOffset.push(p);
        last = p.node;
      }
      const start = byOffset[from];
      const end = byOffset[from + length];
      const r = document.createRange();
      r.setStart(start.node, start.offset);
      r.setEnd(end.node, end.offset);
      const sel = window.getSelection()!;
      sel.removeAllRanges();
      sel.addRange(r);
    };
    (window as any).__frames = (n: number) => new Promise<void>((resolve) => {
      const tick = () => (n-- <= 0 ? resolve() : requestAnimationFrame(tick));
      tick();
    });
  });
}

test.describe('DocxEditor — comment gutter', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
    await installHelpers(page);
  });

  test('an existing comment renders as an inline highlight plus a bubble at its position', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate(async (raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__cmMount(new Uint8Array(raw), {});
      await w.__frames(3);
      editor.layoutComments();
      const highlight = container.querySelector('span.comment-highlight[data-comment-id]') as HTMLElement;
      const bubble = container.querySelector('.docx-comment-bubble[data-thread]') as HTMLElement;
      const gutter = container.querySelector('.docx-comment-gutter') as HTMLElement;
      const out = {
        comments: editor.listComments().map((c: any) => ({ id: c.id, author: c.author, text: c.text })),
        highlightText: highlight?.textContent,
        highlightId: highlight?.getAttribute('data-comment-id'),
        bubbleId: bubble?.dataset.commentId,
        bubbleText: bubble?.textContent?.replace(/\s+/g, ' ').trim(),
        orphan: bubble?.hasAttribute('data-orphan'),
        // The bubble sits at the highlight's vertical position, in the gutter to the right.
        bubbleTop: bubble?.getBoundingClientRect().top,
        highlightTop: highlight?.getBoundingClientRect().top,
        bubbleLeft: bubble?.getBoundingClientRect().left,
        highlightRight: highlight?.getBoundingClientRect().right,
        gutterVisible: !!gutter && !gutter.hidden && !gutter.hasAttribute('data-empty'),
        leaders: container.querySelectorAll('.docx-comment-leaders path').length,
        // The reference marker is engine chrome: present for addressing, never editable/visible.
        markerHidden: getComputedStyle(container.querySelector('a.comment-marker') as HTMLElement).display === 'none',
        markerEditable: (container.querySelector('a.comment-marker') as HTMLElement).getAttribute('contenteditable'),
      };
      editor.close();
      container.remove();
      return out;
    }, bytes);

    expect(res.comments).toHaveLength(1);
    expect(res.comments[0].author).toBe('Eric White');
    expect(String(res.comments[0].id)).toBe(res.highlightId);
    expect(res.bubbleId).toBe(res.highlightId);
    expect(res.bubbleText).toContain('This is a comment.');
    expect(res.orphan).toBe(false);
    expect(res.gutterVisible).toBe(true);
    expect(Math.abs(res.bubbleTop! - res.highlightTop!)).toBeLessThan(4);
    expect(res.bubbleLeft!).toBeGreaterThan(res.highlightRight!);
    expect(res.leaders).toBe(1);
    expect(res.markerHidden).toBe(true);
    expect(res.markerEditable).toBe('false');
  });

  test('a new comment on a selection highlights exactly that range and gets a bubble', async ({ page }) => {
    const res = await page.evaluate(async () => {
      const w = window as any;
      const D = w.Docxodus;
      const { container, editor } = w.__cmMount(D.DocxSessionBridge.CreateBlankDocx(), { commentAuthor: 'Ada' });
      const block = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      block.focus();
      document.execCommand('insertText', false, 'The indemnity clause needs review.');
      block.dispatchEvent(new Event('blur'));
      const fresh = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      w.__cmSelect(fresh, 4, 9); // "indemnity"
      const created = editor.addComment('Please reconsider.', 'Ada');
      await w.__frames(3);
      editor.layoutComments();
      const highlight = container.querySelector('span.comment-highlight') as HTMLElement;
      const bubble = container.querySelector('.docx-comment-bubble[data-thread]') as HTMLElement;
      const out = {
        created: created && { author: created.author, text: created.text },
        highlightText: highlight?.textContent,
        bubbleAuthor: bubble?.querySelector('.docx-comment-author')?.textContent,
        bubbleText: bubble?.querySelector('.docx-comment-text')?.textContent,
        // Editing the commented paragraph keeps the range: type after the comment, commit,
        // and the highlight must survive the re-render.
        textAfterEdit: '',
        highlightAfterEdit: '',
      };
      const again = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      again.focus();
      const sel = window.getSelection()!;
      const r = document.createRange();
      r.selectNodeContents(again);
      r.collapse(false);
      sel.removeAllRanges();
      sel.addRange(r);
      document.execCommand('insertText', false, ' Urgently.');
      again.dispatchEvent(new Event('blur'));
      await w.__frames(3);
      const after = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      out.textAfterEdit = (after.textContent || '').trim();
      out.highlightAfterEdit = container.querySelector('span.comment-highlight')?.textContent || '';
      // And the saved file carries a native comment with the range markers.
      const saved: Uint8Array = editor.save();
      const h = D.DocxSessionBridge.OpenSession(saved, '{}');
      const persisted = JSON.parse(D.DocxSessionBridge.ListComments(h));
      const body = Object.keys(JSON.parse(D.DocxSessionBridge.Project(h)).anchorIndex).find((k) => k.startsWith('p:body:'))!;
      const xml = D.DocxSessionBridge.RawGetXml(h, body);
      D.DocxSessionBridge.CloseSession(h);
      editor.close();
      container.remove();
      return { ...out, persisted: persisted.map((c: any) => c.text), rangeMarkers: /commentRangeStart/.test(xml) && /commentRangeEnd/.test(xml) };
    });

    expect(res.created).toEqual({ author: 'Ada', text: 'Please reconsider.' });
    expect(res.highlightText).toBe('indemnity');
    expect(res.bubbleAuthor).toBe('Ada');
    expect(res.bubbleText).toBe('Please reconsider.');
    expect(res.textAfterEdit).toContain('needs review. Urgently.');
    expect(res.highlightAfterEdit).toBe('indemnity');
    expect(res.persisted).toEqual(['Please reconsider.']);
    expect(res.rangeMarkers).toBe(true);
  });

  test('reply, edit, resolve and delete act on the thread through the bubble', async ({ page }) => {
    const res = await page.evaluate(async () => {
      const w = window as any;
      const D = w.Docxodus;
      const { container, editor } = w.__cmMount(D.DocxSessionBridge.CreateBlankDocx(), { commentAuthor: 'Ada' });
      const block = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      block.focus();
      document.execCommand('insertText', false, 'Payment is due in thirty days.');
      block.dispatchEvent(new Event('blur'));
      const fresh = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      w.__cmSelect(fresh, 0, 7);
      const root = editor.addComment('Net 30?', 'Ada');
      await w.__frames(3);
      editor.layoutComments();
      const bubble = () => container.querySelector('.docx-comment-bubble[data-thread]') as HTMLElement;
      const click = (sel: string) => bubble().querySelector(sel)!.dispatchEvent(new MouseEvent('click', { bubbles: true }));

      // Reply.
      click('[data-comment-action="reply"]');
      // The reply box must exist the moment the click returns — a keystroke that follows the
      // click in the same frame has to land in it, not in the page.
      const boxImmediate = !!bubble().querySelector('textarea[data-comment-reply-text]');
      await w.__frames(2);
      const replyBox = bubble().querySelector('textarea[data-comment-reply-text]') as HTMLTextAreaElement;
      const boxFocused = document.activeElement === replyBox;
      replyBox.value = 'Yes, per the term sheet.';
      click('[data-comment-action="post-reply"]');
      await w.__frames(3);
      const afterReply = {
        boxImmediate,
        boxFocused,
        entries: bubble().querySelectorAll('.docx-comment-entry').length,
        comments: editor.listComments().map((c: any) => ({ text: c.text, reply: !!c.parentAnchorId })),
      };

      // Edit the root's text.
      bubble().querySelector('.docx-comment-root [data-comment-action="edit"]')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await w.__frames(2);
      const editBox = bubble().querySelector('textarea[data-comment-edit-text]') as HTMLTextAreaElement;
      editBox.value = 'Net 30 or net 45?';
      click('[data-comment-action="save"]');
      await w.__frames(3);
      const afterEdit = editor.listComments()[0].text;

      // Resolve — thread state, mirrored on the highlight.
      click('[data-comment-action="resolve"]');
      await w.__frames(3);
      const afterResolve = {
        resolved: editor.listComments()[0].resolved,
        bubbleResolved: bubble().hasAttribute('data-resolved'),
        highlightResolved: container.querySelector('span.comment-highlight')!.hasAttribute('data-comment-resolved'),
      };

      // Delete the thread: root and reply go, the highlight goes with them.
      bubble().querySelector('.docx-comment-root [data-comment-action="delete"]')!.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      await w.__frames(3);
      const afterDelete = {
        comments: editor.listComments().length,
        bubbles: container.querySelectorAll('.docx-comment-bubble[data-thread]').length,
        highlights: container.querySelectorAll('span.comment-highlight').length,
        text: (container.querySelector('[data-anchor][contenteditable="true"]')!.textContent || '').trim(),
      };
      editor.close();
      container.remove();
      return { rootId: root?.anchorId, afterReply, afterEdit, afterResolve, afterDelete };
    });

    expect(res.rootId).toMatch(/^cmt:cmt:/);
    expect(res.afterReply.boxImmediate).toBe(true);
    expect(res.afterReply.boxFocused).toBe(true);
    expect(res.afterReply.entries).toBe(2);
    expect(res.afterReply.comments).toEqual([
      { text: 'Net 30?', reply: false },
      { text: 'Yes, per the term sheet.', reply: true },
    ]);
    expect(res.afterEdit).toBe('Net 30 or net 45?');
    expect(res.afterResolve).toEqual({ resolved: true, bubbleResolved: true, highlightResolved: true });
    expect(res.afterDelete).toEqual({ comments: 0, bubbles: 0, highlights: 0, text: 'Payment is due in thirty days.' });
  });

  test('bubbles stack without overlapping and stay in document order', async ({ page }) => {
    const res = await page.evaluate(async () => {
      const w = window as any;
      const D = w.Docxodus;
      const { container, editor } = w.__cmMount(D.DocxSessionBridge.CreateBlankDocx(), {});
      const block = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
      block.focus();
      document.execCommand('insertText', false, 'Alpha beta gamma delta epsilon zeta eta theta.');
      block.dispatchEvent(new Event('blur'));
      // Three comments on one line: their highlights share a top, so the bubbles must stack.
      for (const [from, len, text] of [[0, 5, 'first'], [6, 4, 'second'], [11, 5, 'third']] as const) {
        const b = container.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
        w.__cmSelect(b, from, len);
        editor.addComment(text, 'Ada');
        await w.__frames(2);
      }
      editor.layoutComments();
      const bubbles = Array.from(container.querySelectorAll('.docx-comment-bubble[data-thread]')) as HTMLElement[];
      const boxes = bubbles.map((b) => ({ text: b.querySelector('.docx-comment-text')!.textContent, top: b.getBoundingClientRect().top, bottom: b.getBoundingClientRect().bottom }));
      boxes.sort((a, b) => a.top - b.top);
      editor.close();
      container.remove();
      return boxes;
    });

    expect(res.map((b) => b.text)).toEqual(['first', 'second', 'third']);
    expect(res[1].top).toBeGreaterThanOrEqual(res[0].bottom);
    expect(res[2].top).toBeGreaterThanOrEqual(res[1].bottom);
  });

  test('the gutter follows into page view and can be hidden', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate(async (raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__cmMount(new Uint8Array(raw), { paginated: true });
      await w.__frames(3);
      editor.layoutComments();
      const highlight = container.querySelector('.page-box span.comment-highlight') as HTMLElement;
      const bubble = container.querySelector('.docx-comment-bubble[data-thread]') as HTMLElement;
      const paginated = {
        bubbleTop: bubble?.getBoundingClientRect().top,
        highlightTop: highlight?.getBoundingClientRect().top,
        orphan: bubble?.hasAttribute('data-orphan'),
      };
      editor.showComments(false);
      const hidden = {
        gutterHidden: (container.querySelector('.docx-comment-gutter') as HTMLElement).hidden,
        highlightsStill: container.querySelectorAll('span.comment-highlight').length,
        flag: container.hasAttribute('data-comments-hidden'),
      };
      editor.showComments(true);
      editor.close();
      container.remove();
      return { paginated, hidden };
    }, bytes);

    expect(res.paginated.orphan).toBe(false);
    expect(Math.abs(res.paginated.bubbleTop! - res.paginated.highlightTop!)).toBeLessThan(4);
    expect(res.hidden.gutterHidden).toBe(true);
    expect(res.hidden.highlightsStill).toBeGreaterThan(0);
    expect(res.hidden.flag).toBe(true);
  });

  test('comments: false renders no comment markup and no gutter', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');
    const res = await page.evaluate(async (raw: number[]) => {
      const w = window as any;
      const { container, editor } = w.__cmMount(new Uint8Array(raw), { comments: false });
      await w.__frames(2);
      const out = {
        highlights: container.querySelectorAll('span.comment-highlight').length,
        gutter: container.querySelectorAll('.docx-comment-gutter').length,
        comments: editor.listComments().length,
      };
      editor.close();
      container.remove();
      return out;
    }, bytes);
    expect(res).toEqual({ highlights: 0, gutter: 0, comments: 1 });
  });
});

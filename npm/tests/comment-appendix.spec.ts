import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const fragmentParagraphs of [undefined, false, true]) test(`appendix fragments its comments with body fragmentation=${fragmentParagraphs}`, async ({ page }) => {
  await page.setContent(`<style>p{margin:0;font:12pt/1.15 "Liberation Serif"}span{white-space:pre-wrap}</style>
    <div id="s"><div id="pagination-comment-margin-registry" style="display:none">
      <div class="comment-margin-note" data-comment-id="0" data-source-anchor-id="cmt:cmt:11111111111111111111111111111111">
        <div class="comment-margin-note-header">Reviewer</div><div class="comment-margin-note-body">
          <p data-source-anchor-id="p:cmt:22222222222222222222222222222222">Leading        spaced\nline ${'Cedar inventory '.repeat(900)}TAIL COMPLETE</p>
        </div>
      </div></div>
      <section data-section-index="0" data-page-width="612" data-page-height="792" data-content-width="468"
        data-content-height="648" data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72">
        <p data-source-anchor-id="p:body:33333333333333333333333333333333"><span data-comment-id="0">Body anchor</span></p>
      </section></div><div id="o"></div>`);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const facts = await page.evaluate(fragmentParagraphs => {
    const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('s'), document.getElementById('o'), {
      commentOverflow: 'appendix', ...(fragmentParagraphs === undefined ? {} : { fragmentParagraphs }),
      layoutToken: { documentVersion: 1, rendererFingerprint: 'comment-appendix-test' },
    }).paginate();
    const text = document.querySelector('[data-comment-appendix] p span')!.firstChild!;
    const range = document.createRange();
    range.setStart(text, 7); range.setEnd(text, 15);
    const spacesWidth = range.getBoundingClientRect().width;
    range.setStart(text, 15); range.setEnd(text, 21);
    const firstWordY = range.getBoundingClientRect().top;
    range.setStart(text, 22); range.setEnd(text, 26);
    return { spacesWidth, nextWordY: range.getBoundingClientRect().top - firstWordY,
      pages: result.totalPages, availability: result.pageMap.availability,
      overflowing: Array.from(document.querySelectorAll<HTMLElement>('#o .page-content')).filter(c => c.scrollHeight > c.clientHeight + 1).length,
      text: document.getElementById('o')!.textContent };
  }, fragmentParagraphs);
  expect(facts.availability).toBe('available');
  expect(facts.pages).toBeGreaterThan(2);
  expect(facts.overflowing).toBe(0);
  expect(facts.text).toContain('TAIL COMPLETE');
  expect(facts.spacesWidth).toBeLessThan(8);
  expect(facts.nextWordY).toBe(0);
});

test('a reply header on its own page retains the root thread owner and flex layout', async ({ page }) => {
  const root = 'cmt:cmt:11111111111111111111111111111111';
  const reply = 'cmt:cmt:44444444444444444444444444444444';
  const header = '<div class="comment-margin-note-header" style="display:flex;justify-content:space-between"><span>Reviewer</span><span>Date</span></div>';
  await page.setContent(`<style>p{margin:0;font:12pt/1.15 "Liberation Serif"}</style>
    <div id="s"><div id="pagination-comment-margin-registry" style="display:none">
      <div class="comment-margin-note" data-comment-id="0" data-source-anchor-id="${root}">${header}
        <p>${'Root message. '.repeat(12)}</p>
        <div class="comment-margin-reply" data-comment-id="1" data-source-anchor-id="${reply}">${header}<p>Reply message.</p></div>
      </div></div>
      <section data-section-index="0" data-page-width="612" data-page-height="168" data-content-width="468"
        data-content-height="24" data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72">
        <p><span data-comment-id="0">Body anchor</span></p>
      </section></div><div id="o"></div>`);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const facts = await page.evaluate(({ root, reply }) => {
    const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('s'), document.getElementById('o'), {
      commentOverflow: 'appendix', layoutToken: { documentVersion: 1, rendererFingerprint: 'comment-owner-test' },
    }).paginate();
    const header = document.querySelector<HTMLElement>(`#o .comment-margin-note-header[data-source-anchor-id="${reply}"]`)!;
    const page = Number(header.closest<HTMLElement>('.page-box')!.dataset.pageNumber);
    return { rootPages: result.pageMap.fragments.filter((f: any) => f.anchorId === root).map((f: any) => f.pageNumber),
      page, display: getComputedStyle(header).display, children: header.children.length };
  }, { root, reply });
  expect(facts.rootPages).toContain(facts.page);
  expect(facts.display).toBe('flex');
  expect(facts.children).toBe(2);
});

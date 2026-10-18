import { test, expect, type Page } from '@playwright/test';
import { commentThreadFixture } from './docx-comment-thread-fixture.js';

async function render(page: Page, threaded: boolean, paginated: boolean, classes: boolean, longAuthor: boolean, extra: { prefix?: string; rootFontSize?: number; resolved?: boolean } = {}) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true);
  await page.addScriptTag({ url: '/pagination.bundle.js' });
  return page.evaluate(({ bytes, paginated, classes, prefix, rootFontSize }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Cedar thread', 'docx-', classes, '', 2, prefix,
      paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
    if (html.startsWith('{')) throw new Error(html);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    document.head.innerHTML = parsed.head.innerHTML;
    document.documentElement.style.fontSize = `${rootFontSize}px`;
    document.body.innerHTML = '<div id="host"></div>';
    const host = document.querySelector<HTMLElement>('#host')!;
    if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, {
      showPageNumbers: false, commentColumnPlacement: 'margin',
    });
    else host.innerHTML = parsed.body.innerHTML;
    return Array.from(host.querySelectorAll<HTMLElement>(paginated ? `.page-comment-margin .${prefix}margin-note` : `.${prefix}margin-note`)).map(note => {
      const n = note.getBoundingClientRect();
      const header = note.querySelector<HTMLElement>(`.${prefix}margin-note-header`)!;
      return { fontSize: parseFloat(getComputedStyle(note).fontSize), body: note.textContent,
        headerParts: Array.from(header.children).map(child => {
          const r = child.getBoundingClientRect();
          const range = document.createRange(); range.selectNodeContents(child);
          const glyphs = range.getBoundingClientRect();
          return { text: child.textContent, glyphLeft: glyphs.left - n.left, glyphRight: glyphs.right - n.left, left: r.left - n.left, right: r.right - n.left,
            top: r.top - n.top, bottom: r.bottom - n.top, width: r.width, height: r.height };
        }), width: n.width, height: n.height };
    });
  }, { bytes: Array.from(commentThreadFixture(threaded, longAuthor, extra.resolved ?? false)), paginated, classes,
    prefix: extra.prefix ?? 'comment-', rootFontSize: extra.rootFontSize ?? 16 });
}

for (const classes of [false, true]) for (const threaded of [false, true]) {
  test(`all comment levels retain one note text size (threaded=${threaded}, classes=${classes})`, async ({ page }) => {
    const notes = await render(page, threaded, false, classes, true);
    expect(notes).toHaveLength(3);
    for (const note of notes) expect(note.fontSize).toBeCloseTo(notes[0].fontSize, 2);
  });
}

for (const longAuthor of [false, true]) test(`physical-margin header stays inside the note (long author=${longAuthor})`, async ({ page }) => {
  const notes = await render(page, longAuthor, true, false, longAuthor);
  expect(notes).toHaveLength(3);
  for (const note of notes) {
    expect(note.headerParts).toHaveLength(3);
    for (const part of note.headerParts) {
      expect(part.width).toBeGreaterThan(0); expect(part.height).toBeGreaterThan(0);
      expect(part.left).toBeGreaterThanOrEqual(-1); expect(part.right).toBeLessThanOrEqual(note.width + 1);
      expect(part.top).toBeGreaterThanOrEqual(-1); expect(part.bottom).toBeLessThanOrEqual(note.height + 1);
    }
  }
});

test('custom-prefix notes follow the root text scale at every thread level', async ({ page }) => {
  const notes = await render(page, true, false, true, true, { prefix: 'review-', rootFontSize: 20 });
  expect(notes).toHaveLength(3);
  for (const note of notes) expect(note.fontSize).toBeCloseTo(17, 2);
});

test('resolved header glyphs stay inside a narrow nested note', async ({ page }) => {
  const notes = await render(page, true, true, true, true, { resolved: true });
  expect(notes).toHaveLength(3);
  for (const note of notes) {
    expect(note.headerParts).toHaveLength(4);
    for (const part of note.headerParts) {
      expect(part.glyphLeft).toBeGreaterThanOrEqual(-1);
      expect(part.glyphRight).toBeLessThanOrEqual(note.width + 1);
    }
  }
});

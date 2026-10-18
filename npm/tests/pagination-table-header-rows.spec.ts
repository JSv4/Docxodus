import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { generateExactRowDocx, type ExactRowOptions } from './docx-exact-row-fixture.js';

/**
 * Word repeats a table's header rows (`w:tblHeader`) at the top of every page the table continues
 * on, whatever their height rule. The paginated view did not repeat them at all (issue #1171).
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

async function paginate(page: Page, options: ExactRowOptions, commentRenderMode = 0) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate(({ input, commentRenderMode }) => {
    const D = (window as any).Docxodus;
    const bin = new Uint8Array(input);
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      bin, 'Document', 'docx-', false, '', commentRenderMode, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, /* stampAnchors */ true,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    const headerRow = staging.querySelector<HTMLElement>('tr[data-docx-header-row]');
    const headerAnchor = headerRow?.dataset.sourceAnchorId ?? null;
    // The editor anchors of the header row's cell paragraphs.
    const editorAnchors = Array.from(headerRow?.querySelectorAll<HTMLElement>('[data-anchor]') ?? [])
      .map((element) => element.dataset.anchor!);
    const engine = new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
      showPageNumbers: false,
      layoutToken: { documentVersion: 0, rendererFingerprint: 'header-rows' },
    });
    const pageMap = engine.paginate().pageMap;
    const bridge = D.DocxSessionBridge;
    const handle = bridge.OpenSession(bin, '');
    let registration;
    try {
      registration = JSON.parse(bridge.RegisterPageMap(handle, JSON.stringify(pageMap), 'header-rows'));
    } finally {
      bridge.CloseSession(handle);
    }
    const pages = Array.from(container.querySelectorAll<HTMLElement>('.page-box')).map((box) =>
      Array.from(box.querySelectorAll<HTMLTableRowElement>('.page-content tr'))
        .map((row) => (row.cells[0]?.textContent ?? '').trim()));
    return {
      headerAnchor,
      pages,
      editorAnchors: editorAnchors.length,
      activeEditorAnchors: editorAnchors.map((anchor) =>
        container.querySelectorAll(`[data-anchor="${anchor}"]`).length),
      headerFragmentPages: pageMap.fragments
        .filter((fragment: any) => fragment.anchorId === headerAnchor)
        .map((fragment: any) => fragment.pageNumber),
      registration,
      marginComments: Array.from(container.querySelectorAll<HTMLElement>('.page-box')).map((box) =>
        box.querySelectorAll('.page-comment-margin [data-comment-id]').length),
      commentMarks: Array.from(container.querySelectorAll<HTMLElement>('.page-box')).map((box) =>
        box.querySelectorAll('.page-content a[id^="comment-ref-"]').length),
    };
  }, { input: Array.from(generateExactRowDocx(options)), commentRenderMode });
}

for (const rule of ['exact', 'atLeast'] as const) {
  test(`a ${rule}-height header row opens every page its table continues on`, async ({ page }) => {
    const result = await paginate(page, { headerRule: rule, bodyRows: 70 });

    expect(result.headerAnchor).not.toBeNull();
    expect(result.pages.length).toBeGreaterThan(1);
    // Every page's first table row is the header; the body rows appear once each, in order.
    for (const rows of result.pages) expect(rows[0]).toBe('Heading Ledger');
    const body = result.pages.flat().filter((text) => text !== 'Heading Ledger');
    expect(body.filter((text) => /^Row\d\d/.test(text)))
      .toEqual(Array.from({ length: 69 }, (_, i) => `Row${String(i + 2).padStart(2, '0')} ledger entry.`));
    // The repeats are presentation copies: one editor anchor, one PageMap fragment per page.
    expect(result.editorAnchors).toBeGreaterThan(0);
    expect(result.activeEditorAnchors).toEqual(Array(result.editorAnchors).fill(1));
    expect(result.headerFragmentPages).toEqual(result.pages.map((_, index) => index + 1));
    expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);
  });
}

test('a table without a header row does not repeat its first row', async ({ page }) => {
  const result = await paginate(page, { bodyRows: 70 });

  expect(result.headerAnchor).toBeNull();
  expect(result.pages.length).toBeGreaterThan(1);
  expect(result.pages.flat().filter((text) => text.startsWith('Alder'))).toHaveLength(1);
  expect(result.pages[1][0]).toMatch(/^Row\d\d ledger entry\.$/);
});

test('a comment in a header row is shown once, on the page that holds the row itself', async ({ page }) => {
  // Margin comments: the copies that open later pages carry neither the comment nor its mark.
  const result = await paginate(page, { headerRule: 'exact', bodyRows: 70, headerComment: true }, 2);

  expect(result.pages.length).toBeGreaterThan(1);
  // The first page's row carries the comment's mark ("Heading Ledger[7]").
  for (const rows of result.pages) expect(rows[0]).toMatch(/^Heading Ledger/);
  const once = result.pages.map((_, index) => (index === 0 ? 1 : 0));
  expect(result.commentMarks).toEqual(once);
  expect(result.marginComments.map((count) => Math.min(count, 1))).toEqual(once);
  expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);
});

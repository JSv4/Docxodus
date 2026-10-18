import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { generateExactRowDocx } from './docx-exact-row-fixture.js';

/**
 * An exact-height row (`w:hRule="exact"`) hides whatever its height cannot hold, as Word does. A
 * paragraph it hides whole has no visible piece, so the PageMap has no fragment for it; it used to
 * fail the whole layout instead ("source anchor … has no measurable fragment"), which the PDF export
 * reported as output_verification_failure (issue #1168).
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

interface Result {
  error: string | null;
  paragraphs: Array<{ anchor: string; text: string }>;
  mapped: string[];
  registration: { success: boolean } | null;
  /** Each cell paragraph's citation availability, through GetPageCitation. */
  citations: Record<string, { availability: string; unavailableReason?: string }>;
}

async function paginate(page: Page, bytes: Uint8Array): Promise<Result> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const D = (window as any).Docxodus;
    const bin = new Uint8Array(input);
    const html: string = D.DocumentConverter.ConvertDocxToHtmlComplete(
      bin, 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    const paragraphs = Array.from(staging.querySelectorAll<HTMLElement>('td p[data-source-anchor-id]'))
      .map((p) => ({ anchor: p.dataset.sourceAnchorId!, text: p.textContent ?? '' }));
    try {
      const engine = new (window as any).DocxodusPagination.PaginationEngine(staging, container, {
        showPageNumbers: false,
        fragmentParagraphs: true,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'exact-row' },
      });
      const pageMap = engine.paginate().pageMap;
      const bridge = D.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      let registration;
      const citations: Record<string, { availability: string; unavailableReason?: string }> = {};
      try {
        registration = JSON.parse(bridge.RegisterPageMap(handle, JSON.stringify(pageMap), 'exact-row'));
        const request = JSON.stringify({ documentVersion: 0, rendererFingerprint: 'exact-row' });
        for (const { anchor } of paragraphs) citations[anchor] = JSON.parse(bridge.GetPageCitation(handle, anchor, request));
      } finally {
        bridge.CloseSession(handle);
      }
      return { error: null, paragraphs, mapped: pageMap.fragments.map((f: any) => f.anchorId), registration, citations };
    } catch (error) {
      return { error: String(error), paragraphs, mapped: [], registration: null, citations: {} };
    }
  }, Array.from(bytes));
}

test('a paragraph an exact-height row hides whole has no fragment and does not fail the map', async ({ page }) => {
  const result = await paginate(page, generateExactRowDocx({ clippedParagraphs: 3 }));

  expect(result.error).toBeNull();
  expect(result.registration?.success, JSON.stringify(result.registration)).toBe(true);
  const anchorOf = (word: string) => result.paragraphs.find((p) => p.text.startsWith(word))!.anchor;
  // Alder fills the row, Birch is cut by its bottom edge, Cedar lies wholly below it.
  expect(result.mapped).toContain(anchorOf('Alder'));
  expect(result.mapped).toContain(anchorOf('Birch'));
  expect(result.mapped).not.toContain(anchorOf('Cedar'));
  // A citation of the hidden paragraph is unavailable for that reason; its neighbours cite normally.
  expect(result.citations[anchorOf('Cedar')]).toMatchObject({ availability: 'unavailable', unavailableReason: 'anchor_not_mapped' });
  expect(result.citations[anchorOf('Alder')].availability).toBe('available');
});

/** Synthetic staging for one section; `body` is its content. */
function section(body: string): string {
  return `<div id="pagination-staging"><div data-section-index="0"
       data-page-width="122" data-page-height="80" data-content-width="120" data-content-height="78"
       data-margin-top="1" data-margin-right="1" data-margin-bottom="1" data-margin-left="1">${body}</div></div>
    <div id="pagination-container"></div>`;
}

async function paginateSynthetic(page: Page, html: string): Promise<{ error: string; mapped: string[] }> {
  await page.setContent('<div id="viewer"></div>');
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    try {
      const result = (window as any).DocxodusPagination.paginateHtml(input, 'viewer', {
        showPageNumbers: false,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'synthetic' },
      });
      return { error: '', mapped: result.pageMap.fragments.map((f: any) => f.anchorId) };
    } catch (error) {
      return { error: error instanceof Error ? error.message : String(error), mapped: [] };
    }
  }, html);
}

test('a paragraph a nested exact row lies below its outer exact row has no fragment either', async ({ page }) => {
  // The inner row's clip box holds the paragraph, but the outer row hides the inner row whole.
  const result = await paginateSynthetic(page, section(`
    <div data-docx-exact-row-cell="true" style="height:12pt;overflow:hidden">
      <p data-source-anchor-id="p:body:outer" style="height:12pt;margin:0">outer</p>
      <div data-docx-exact-row-cell="true" style="height:12pt;overflow:hidden">
        <p data-source-anchor-id="p:body:inner" style="height:12pt;margin:0">inner</p>
      </div>
    </div>`));

  expect(result.error).toBe('');
  expect(result.mapped).toContain('p:body:outer');
  expect(result.mapped).not.toContain('p:body:inner');
});

/** A borderless table whose rows are `rows`; `zeroExact` rows carry the converter's zero-height clip. */
function zeroTable(id: string, rows: Array<'zeroExact' | 'empty'>): string {
  const row = (kind: string, index: number) => `<tr data-source-anchor-id="tr:body:${id}${index}">`
    + `<td data-source-anchor-id="tc:body:${id}${index}" style="padding:0">`
    + (kind === 'zeroExact'
      ? `<div data-docx-exact-row-cell="true" style="height:0;overflow:hidden">`
        + `<p data-source-anchor-id="p:body:${id}${index}" style="margin:0">hidden</p></div>`
      : '')
    + '</td></tr>';
  return `<table data-source-anchor-id="tbl:body:${id}" style="border-collapse:collapse;border-spacing:0">`
    + `${rows.map(row).join('')}</table>`;
}

test('a table made only of zero-height exact rows has no fragment either', async ({ page }) => {
  const result = await paginateSynthetic(page, section(
    `${zeroTable('z', ['zeroExact'])}<p data-source-anchor-id="p:body:after" style="margin:0">after</p>`));

  expect(result.error).toBe('');
  expect(result.mapped).toEqual(['p:body:after']);
});

for (const [name, rows] of [['no exact row', ['empty']], ['an exact row beside an empty one', ['zeroExact', 'empty']]] as const) {
  test(`a table of no height with ${name} still fails the map`, async ({ page }) => {
    const result = await paginateSynthetic(page, section(
      `${zeroTable('g', [...rows])}<p data-source-anchor-id="p:body:after" style="margin:0">after</p>`));

    expect(result.error).toContain('source anchor tbl:body:g has no measurable fragment');
  });
}

test('content any other clip hides whole still fails the map', async ({ page }) => {
  const result = await paginateSynthetic(page, section(`
    <div data-source-anchor-id="tbl:body:box" style="height:12pt;overflow:hidden;margin:0">
      <p data-source-anchor-id="p:body:shown" style="height:12pt;margin:0">shown</p>
      <p data-source-anchor-id="p:body:lost" style="height:12pt;margin:0">lost</p>
    </div>`));
  expect(result.error).toContain('source anchor p:body:lost has no measurable fragment');
});

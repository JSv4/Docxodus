import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { tableRemainderDocx } from './docx-table-remainder-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const kind of ['rows', 'oversized', 'row-cant-split', 'cell', 'kept', 'kept-bookmarks', 'kept-borders', 'cell-cant-split', 'columns', 'columns-fit'] as const) {
  test(`public ${kind} table preserves PDF inventory and reports column fallback`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(tableRemainderDocx(kind), {
      browser, outputs: ['html', 'pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    });
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    const texts: string[] = [];
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      for (let number = 1; number <= pdf.numPages; number++) {
        const content = await (await pdf.getPage(number)).getTextContent();
        texts.push(content.items.filter(item => 'str' in item).map(item => item.str).join(' '));
      }
    } finally { await task.destroy(); }
    const count = kind === 'oversized' ? 18 : kind === 'columns' ? 12 : kind === 'columns-fit' ? 4 : 8;
    const tokens = Array.from({ length: count }, (_, i) => `Cedar${i + 1}`);
    const deferred = kind.startsWith('kept') || kind === 'cell-cant-split' || kind === 'columns';
    const firstPage = texts[0].match(/Cedar\d+/g) ?? [];
    if (deferred) expect(firstPage).toEqual([]);
    else {
      expect(firstPage.length).toBeGreaterThanOrEqual(4);
      expect(firstPage).toEqual(tokens.slice(0, firstPage.length));
    }
    expect(texts.join(' ').match(/Cedar\d+/g)).toEqual(tokens);
    expect(texts.join(' ').match(/Complete\./g)).toHaveLength(1);
    const warnings = result.renderReport.warnings.filter(w => w.code === 'column_table_page_remainder');
    expect(warnings).toHaveLength(kind === 'columns' ? 1 : 0);
    if (kind === 'columns') expect(warnings[0]).toMatchObject({ phase: 'pagination', partUri: '/word/document.xml' });
    await page.setContent(result.html!);
    if (kind === 'kept-bookmarks') {
      for (let i = 0; i < count; i++) await expect(page.locator(`[id="RowEnd${i}"]`)).toHaveCount(1);
    }
    const overflow = await page.locator('.page-content').evaluateAll(bodies => bodies.flatMap(body =>
      Array.from(body.querySelectorAll('td p'), p => p.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom)));
    expect(overflow.every(height => height <= 1)).toBe(true);
  });
}

test('strict unsupported-content policy rejects the intact column-table move', async ({ browser }) => {
  test.setTimeout(120_000);
  await expect(renderDocxArtifacts(tableRemainderDocx('columns'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
    unsupportedContent: 'strict',
  })).rejects.toMatchObject({ code: 'resource_policy_failure', phase: 'pagination' });
});

test('column fallback identifies a relocated main part', async ({ browser }) => {
  test.setTimeout(120_000);
  const result = await renderDocxArtifacts(tableRemainderDocx('columns', 'ledger/body.xml'), {
    browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000,
  });
  expect(result.renderReport.warnings.filter(w => w.code === 'column_table_page_remainder'))
    .toEqual([expect.objectContaining({ partUri: '/ledger/body.xml', phase: 'pagination' })]);
});

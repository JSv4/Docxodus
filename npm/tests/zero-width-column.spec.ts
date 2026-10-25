import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { readFileSync } from 'node:fs';

// Native Word captures pin a measurable one-point sliver; retain its anchors.

const __dirname = path.dirname(fileURLToPath(import.meta.url));

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

function input(id: string): Uint8Array {
  return new Uint8Array(readFileSync(path.join(__dirname, '../../TestFiles/ZeroWidthColumns', id + '.docx')));
}

interface Mapped { error: string | null; fragments: Array<{ anchorId: string; height: number; width: number; text: string }> }

async function paginate(page: Page, bytes: Uint8Array): Promise<Mapped> {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlWithPagination(
      new Uint8Array(input), 'Document', 'docx-', true, '', -1, 'comment-', 1, 1, 'page-');
    const viewer = document.createElement('div');
    viewer.id = 'zero-width-viewer';
    document.body.append(viewer);
    try {
      const { pageMap } = (window as any).DocxodusPagination.paginateHtml(html, 'zero-width-viewer', {
        fragmentParagraphs: true, showPageNumbers: false,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'zero-width-column' },
      });
      const text = (anchorId: string) => viewer.querySelector(`[data-source-anchor-id="${anchorId}"]`)?.textContent ?? '';
      return { error: null, fragments: pageMap.fragments.map((f: any) =>
        ({ anchorId: f.anchorId, height: f.geometry.height, width: f.geometry.width, text: text(f.anchorId) })) };
    } catch (error) {
      return { error: String(error), fragments: [] };
    }
  }, Array.from(bytes));
}

const rowHeight = (result: Mapped) => result.fragments.find(f => f.anchorId.startsWith('tr:'))!.height;
const mappedTexts = (result: Mapped) => result.fragments.filter(f => f.anchorId.startsWith('p:')).map(f => f.text);

for (const id of ['zero-exact', 'zero-atleast', 'width-control']) {
  test(`${id}: the paginated view preserves the narrow cell and its paragraph`, async ({ page }) => {
    const result = await paginate(page, input(id));
    expect(result.error).toBeNull();
    expect(mappedTexts(result)).toEqual(['Alder', 'Birch']);
    const cells = result.fragments.filter(f => f.anchorId.startsWith('tc:'));
    expect(cells).toHaveLength(2);
    expect(cells[0].width).toBeCloseTo(id === 'width-control' ? 24 : 1, 1);
    expect(rowHeight(result)).toBeCloseTo(id === 'zero-atleast' ? 75 : 18, 0);
  });
}

for (const id of ['zero-exact', 'zero-atleast', 'width-control',
  'tables-zero-width-column-times-new-roman', 'tables-zero-width-column-control-times-new-roman',
  'zero-default-margins', 'zero-explicit-margins', 'zero-spanned-columns', 'zero-separate-columns']) {
  test(`${id}: the captured Word input exports with measurable cell and paragraph anchors`, async ({ browser }) => {
    test.setTimeout(120_000);
    const actual = await renderDocxArtifacts(input(id), {
      browser, outputs: ['pdf'], reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 100_000,
    });
    expect(actual.pageCount).toBe(1);
    expect(actual.pdf!.byteLength).toBeGreaterThan(1000);
    const cells = actual.pageMap.fragments.filter(f => f.anchorId.startsWith('tc:'));
    const singleTable = ['zero-exact', 'zero-atleast', 'width-control'].includes(id);
    expect(cells).toHaveLength(singleTable ? 2 : id === 'zero-separate-columns' ? 9 : 6);
    expect(cells.every(f => f.geometry.width > 0 && f.geometry.height > 0)).toBe(true);
  });
}

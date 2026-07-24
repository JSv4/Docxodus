import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, {
    timeout: 30000,
  });
}

function countPdfPages(pdf: Buffer): number {
  // Chromium writes one uncompressed page dictionary per physical PDF page. The word
  // boundary deliberately excludes the /Type /Pages tree node.
  return (pdf.toString('latin1').match(/\/Type\s*\/Page\b/g) ?? []).length;
}

// Keep this browser-level proof independent of a WASM rebuild. HC008e locks the
// converter's generated form of these same rules, while this test proves that
// Chromium honors them for real PDF output at both viewer scales.
const paginationPrintCss = `
  @media print {
    .page-container {
      display: block;
      gap: 0;
      padding: 0;
      background: transparent;
      min-height: 0;
    }
    .page-box {
      zoom: 1 !important;
      transform: none !important;
      margin: 0 !important;
      box-shadow: none;
    }
    .page-box + .page-box {
      break-before: page;
      page-break-before: always;
    }
  }
`;

test.describe('Pagination print layout', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  for (const scale of [1, 0.8]) {
    test(`prints one physical page per paginator page at scale ${scale}`, async ({ page }) => {
      // Get the product's screen pagination CSS, then replace the document body with a
      // tiny synthetic two-page section. This keeps the regression focused on print
      // layout and avoids relying on any external benchmark fixture.
      const source = readTestFile('HC006-Test-01.docx');
      const conversion = await page.evaluate(
        ([bytes, paginationScale]) => (window as any).DocxodusTests.convertToHtmlWithPagination(
          new Uint8Array(bytes), 1, paginationScale),
        [Array.from(source), scale]
      );

      expect(conversion.error).toBeUndefined();
      expect(conversion.html).toBeDefined();

      await page.setContent(conversion.html!);
      await page.addStyleTag({ content: paginationPrintCss });
      await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });

      const pagination = await page.evaluate((paginationScale) => {
        document.body.style.margin = '0';

        const staging = document.getElementById('pagination-staging') as HTMLElement;
        const pageContainer = document.getElementById('pagination-container') as HTMLElement;
        if (!staging || !pageContainer) {
          return { error: 'Pagination elements not found' };
        }

        staging.innerHTML = `
          <div data-section-index="0"
               data-page-width="612" data-page-height="792"
               data-content-width="468" data-content-height="648"
               data-margin-top="72" data-margin-right="72"
               data-margin-bottom="72" data-margin-left="72">
            <p>First page</p>
            <div class="page-break" data-page-break="true"></div>
            <p>Second page</p>
          </div>`;
        pageContainer.innerHTML = '';

        const { PaginationEngine } = (window as any).DocxodusPagination;
        const engine = new PaginationEngine(staging, pageContainer, {
          scale: paginationScale,
          showPageNumbers: false,
        });
        const result = engine.paginate();

        return {
          totalPages: result.totalPages,
          pageBoxes: pageContainer.querySelectorAll('.page-box').length,
        };
      }, scale);

      if ('error' in pagination) {
        throw new Error(pagination.error);
      }
      expect(pagination.totalPages).toBe(2);
      expect(pagination.pageBoxes).toBe(2);

      await page.emulateMedia({ media: 'print' });
      const printStyles = await page.evaluate(() => {
        const container = document.getElementById('pagination-container') as HTMLElement;
        const pageBoxes = Array.from(container.querySelectorAll('.page-box')) as HTMLElement[];
        return {
          containerDisplay: getComputedStyle(container).display,
          containerPaddingTop: getComputedStyle(container).paddingTop,
          secondPageBreakBefore: getComputedStyle(pageBoxes[1]).breakBefore,
          firstPageZoom: getComputedStyle(pageBoxes[0]).zoom,
          firstPageTransform: getComputedStyle(pageBoxes[0]).transform,
        };
      });

      expect(printStyles.containerDisplay).toBe('block');
      expect(printStyles.containerPaddingTop).toBe('0px');
      expect(printStyles.secondPageBreakBefore).toBe('page');
      expect(printStyles.firstPageZoom).toBe('1');
      expect(printStyles.firstPageTransform).toBe('none');

      const pdf = await page.pdf({
        format: 'Letter',
        printBackground: true,
        margin: { top: '0', right: '0', bottom: '0', left: '0' },
      });
      expect(countPdfPages(pdf)).toBe(2);
    });
  }
});

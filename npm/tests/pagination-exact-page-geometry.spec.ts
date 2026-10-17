import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * The paginator lays every page out from the section's page-geometry attributes. They were written
 * to one decimal of a point, so a 1417-twip margin (70.85pt) arrived as 70.8pt and the text column
 * was laid out a twentieth of a point off what the document declares (issue #1164).
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** An A4 page with the 1417 / 1134 / 1134 / 1701-twip margins common to European templates. */
function a4Docx(): Uint8Array {
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Larch ledger on an A4 page.</w:t></w:r></w:p>`
      + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

async function paginatedGeometry(page: Page, bytes: Uint8Array) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      false, false, false, false, false, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    new (window as any).DocxodusPagination.PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
    const box = container.querySelector<HTMLElement>('.page-box')!.getBoundingClientRect();
    const content = container.querySelector<HTMLElement>('.page-content')!.getBoundingClientRect();
    const pt = (px: number) => px * 72 / 96;
    return {
      pageWidth: pt(box.width), pageHeight: pt(box.height),
      contentLeft: pt(content.left - box.left), contentTop: pt(content.top - box.top),
      contentWidth: pt(content.width),
    };
  }, Array.from(bytes));
}

test('a section with 1417-twip margins lays its text column out at the exact points', async ({ page }) => {
  const geometry = await paginatedGeometry(page, a4Docx());
  // Browsers lay out in 1/64 CSS px (about 0.012pt); one decimal of a point put each edge 0.05pt off.
  const close = (actual: number, expected: number) => expect(Math.abs(actual - expected), `${actual} vs ${expected}`)
    .toBeLessThan(0.02);
  close(geometry.pageWidth, 595.3);
  close(geometry.pageHeight, 841.9);
  close(geometry.contentTop, 70.85);
  close(geometry.contentLeft, 85.05);
  close(geometry.contentWidth, 453.55);
});

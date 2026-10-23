import { expect, test, type Page } from '@playwright/test';
import * as path from 'path';
import { fileURLToPath } from 'url';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

/**
 * A PAGE / NUMPAGES footer deleted under tracked changes keeps its field instructions in
 * w:delInstrText. Paginated markup lost the fields' page-number markers, so every page showed the
 * cached "Folio 1 / 2" (issue #1272). The undeleted footer is the control.
 */

const __dirname = path.dirname(fileURLToPath(import.meta.url));

function folioDocx(deleted: boolean): Uint8Array {
  const [t, instr] = deleted ? ['delText', 'delInstrText'] : ['t', 'instrText'];
  const text = (value: string) => `<w:r><w:${t} xml:space="preserve">${value}</w:${t}></w:r>`;
  const field = (code: string, cached: string) => `<w:r><w:fldChar w:fldCharType="begin"/></w:r>`
    + `<w:r><w:${instr} xml:space="preserve"> ${code} </w:${instr}></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>`
    + `${text(cached)}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
  const folio = text('Folio ') + field('PAGE', '1') + text(' / ') + field('NUMPAGES', '2');
  const footer = deleted ? `<w:del w:id="31" w:author="Review Example" w:date="2026-09-12T10:00:00Z">${folio}</w:del>` : folio;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="footer" Type="${R_NS}/footer" Target="footer1.xml"/></Relationships>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${W_NS}"><w:p>${footer}</w:p></w:ftr>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>`
      + `<w:p><w:r><w:t>Lunar receiver inspection</w:t></w:r></w:p>`
      + `<w:p><w:r><w:br w:type="page"/></w:r><w:r><w:t>Signal record</w:t></w:r></w:p>`
      + `<w:sectPr><w:footerReference w:type="default" r:id="footer"/><w:pgSz w:w="12240" w:h="15840"/>`
      + `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

async function paginatedFooters(page: Page, bytes: Uint8Array) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
  await page.addScriptTag({ path: path.join(__dirname, '../dist/pagination.bundle.js') });
  return page.evaluate((input) => {
    const html: string = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(input), 'Document', 'docx-', false, '', 0, 'comment-',
      /* paginationMode */ 1, 1, 'page-', false, 0, 'annot-',
      /* footnotes */ false, /* headersAndFooters */ true, /* trackedChanges */ true,
      /* showDeleted */ true, /* moves */ true, false, null, false,
    );
    const host = document.createElement('div');
    host.innerHTML = html;
    document.body.appendChild(host);
    const staging = host.querySelector<HTMLElement>('#pagination-staging')!;
    const container = host.querySelector<HTMLElement>('#pagination-container')!;
    new (window as any).DocxodusPagination.PaginationEngine(staging, container, { showPageNumbers: false }).paginate();
    return Array.from(container.querySelectorAll<HTMLElement>('.page-box')).map((box) => {
      const footer = box.querySelector<HTMLElement>('.page-footer')!;
      const text = (node: Element | null) => node?.textContent!.replace(/\s+/g, ' ').trim() ?? null;
      return { text: text(footer), deleted: text(footer.querySelector('del')) };
    });
  }, Array.from(bytes));
}

test('a deleted PAGE / NUMPAGES footer shows each page its own number, still marked deleted', async ({ page }) => {
  expect(await paginatedFooters(page, folioDocx(true))).toEqual([
    { text: 'Folio 1 / 2', deleted: 'Folio 1 / 2' },
    { text: 'Folio 2 / 2', deleted: 'Folio 2 / 2' },
  ]);
});

test('the undeleted footer control numbers its pages the same way', async ({ page }) => {
  expect(await paginatedFooters(page, folioDocx(false))).toEqual([
    { text: 'Folio 1 / 2', deleted: null },
    { text: 'Folio 2 / 2', deleted: null },
  ]);
});

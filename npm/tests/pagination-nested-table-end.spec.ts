import { expect, test } from '@playwright/test';
import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

function nestedTableDocx(formatted: boolean): Uint8Array {
  const tail = formatted ? '<w:pPr><w:spacing w:after="240" w:line="400" w:lineRule="exact"/></w:pPr>' : '';
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
      </Types>`) },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="office" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`) },
    { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"><w:docDefaults>
      <w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
      <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>
      </w:docDefaults></w:styles>`) },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>
      <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>
        <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>
          <w:p><w:r><w:t>Inner text</w:t></w:r></w:p>
        </w:tc></w:tr></w:tbl><w:p>${tail}<w:r><w:t/></w:r></w:p>
      </w:tc></w:tr></w:tbl><w:p><w:r><w:t>Following body</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`) },
  ]);
}

for (const formatted of [false, true]) {
  test(`${formatted ? 'formatted' : 'structural'} nested-table end paragraph retains PageMap identity`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, false, false,
      ) as string;
      if (html.startsWith('{')) throw new Error(html);
      const source = new DOMParser().parseFromString(html, 'text/html');
      const sourceTail = Array.from(source.querySelectorAll('p')).find(p => p.parentElement?.localName === 'td' && !p.textContent?.trim())!;
      const anchor = sourceTail.getAttribute('data-source-anchor-id');
      document.body.innerHTML = `<main id="fixture">${html}</main>`;
      const engine = new (window as any).DocxodusPagination.PaginationEngine(
        document.getElementById('pagination-staging'), document.getElementById('pagination-container'),
        { showPageNumbers: false, pageGap: 0 });
      const pagination = engine.paginate();
      const map = engine.materializePageMap(0, 'nested-table-end-test');
      const content = document.querySelector<HTMLElement>('#fixture .page-content')!;
      const tail = Array.from(content.querySelectorAll('p')).find(p => p.getAttribute('data-source-anchor-id') === anchor)!;
      const table = content.querySelector('table')!;
      const next = Array.from(content.querySelectorAll('p')).find(p => p.textContent === 'Following body')!;
      return { pages: pagination.totalPages, anchor, height: tail.getBoundingClientRect().height * 0.75,
        after: parseFloat(getComputedStyle(tail).marginBottom) * 0.75,
        gap: (next.getBoundingClientRect().top - table.getBoundingClientRect().bottom) * 0.75,
        text: Array.from(content.querySelectorAll('p'), p => p.textContent?.trim()).filter(Boolean).join(' '),
        fragments: map.fragments.filter((f: any) => f.anchorId === anchor),
      };
    }, { bytes: Array.from(nestedTableDocx(formatted)) });
    expect(result.pages).toBe(1);
    expect(result.text).toBe('Inner text Following body');
    expect(result.anchor).toBeTruthy();
    expect(result.fragments).toHaveLength(1);
    expect(result.fragments[0].inTableCell).toBe(true);
    expect(result.fragments[0].geometry.height).toBeGreaterThan(0);
    expect(result.gap).toBeLessThan(2);
    if (formatted) {
      expect(result.height).toBeCloseTo(20, 1);
      expect(result.after).toBeCloseTo(12, 1);
    } else {
      expect(result.height).toBeGreaterThan(0);
      expect(result.height).toBeLessThan(0.1);
      expect(result.after).toBe(0);
    }
  });
}

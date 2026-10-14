import { expect, Page, test } from '@playwright/test';
import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

/**
 * Issue #827 — Word pulls an over-wide table into the left margin with a negative w:tblInd.
 * The paginated view must keep the negative indent and must not clip the table at the body
 * column: Word only clips at the paper edge.
 */

// US Letter landscape, 1in margins: a 648pt text column holding a 750pt table indented -36.25pt.
function generateNegativeIndentDocx(rtl = false, rowCount = 1): Uint8Array {
  const rows = Array.from({ length: rowCount }, (_, i) => {
    const merged = rowCount > 1 && i % 8 === 0;
    const cell = (text: string, width: number, span = '') => `<w:tc><w:tcPr>
      <w:tcW w:w="${width}" w:type="dxa"/>${span}</w:tcPr>
      <w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
    return `<w:tr>${rowCount > 1 ? '<w:trPr><w:trHeight w:val="600" w:hRule="exact"/></w:trPr>' : ''}
      ${merged ? cell(`ROW ${i}`, 15000, '<w:gridSpan w:val="2"/>')
        : cell(`ROW ${i}`, 7500) + cell('SECOND CELL', 7500)}</w:tr>`;
  }).join('');
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:tbl>
    <w:tblPr>${rtl ? '<w:bidiVisual/>' : ''}<w:tblW w:w="15000" w:type="dxa"/><w:tblInd w:w="-725" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="7500"/><w:gridCol w:w="7500"/></w:tblGrid>
    ${rows}
  </w:tbl>
  <w:p/>
  <w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    { name: 'word/document.xml', data: xml(documentXml) },
  ]);
}

async function paginateDocx(page: Page, rtl: boolean, scale: number, rowCount = 1) {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });

  return page.evaluate(async ({ bytes, scale }) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false,
    ) as string;
    if (html.startsWith('{')) throw new Error(`conversion failed: ${html.slice(0, 300)}`);

    document.body.innerHTML = '<main id="fixture"></main>';
    const result = (window as any).DocxodusPagination.paginateHtml(
      html,
      document.getElementById('fixture'),
      { scale, showPageNumbers: false, pageGap: 0,
        layoutToken: { documentVersion: 0, rendererFingerprint: 'negative-indent' } },
    );
    await document.fonts.ready;

    return {
      totalPages: result.totalPages,
      tableFragments: result.pageMap.fragments.filter((f: any) => f.anchorId.startsWith('tbl:')),
    };
  }, { bytes: Array.from(generateNegativeIndentDocx(rtl, rowCount)), scale });
}

for (const rtl of [false, true]) {
  for (const scale of [1, 0.65]) {
    test(`negative ${rtl ? 'RTL' : 'LTR'} indent stays visible at scale ${scale}`, async ({ page }) => {
      const result = await paginateDocx(page, rtl, scale);
      const geometry = await page.locator('#fixture .page-content table').evaluate((table) => {
        const bodyRect = table.closest('.page-content')!.getBoundingClientRect();
        const tableRect = table.getBoundingClientRect();
        const midY = tableRect.top + tableRect.height / 2;
        // Hit-testing honours clipping; bounding boxes alone cannot prove visibility.
        const hits = (x: number) => table.contains(document.elementFromPoint(x, midY));
        return {
          leftOverhang: bodyRect.left - tableRect.left,
          rightOverhang: tableRect.right - bodyRect.right,
          leftMarginVisible: hits(tableRect.left + 4),
          rightMarginVisible: hits(tableRect.right - 4),
        };
      });

      expect(result.totalPages).toBe(1);
      expect(rtl ? geometry.rightOverhang : geometry.leftOverhang)
        .toBeCloseTo(36.25 * (4 / 3) * scale, 0);
      expect(geometry.leftOverhang).toBeGreaterThan(0);
      expect(geometry.rightOverhang).toBeGreaterThan(0);
      expect(geometry.leftMarginVisible).toBe(true);
      expect(geometry.rightMarginVisible).toBe(true);
      expect(result.tableFragments).toHaveLength(1);
      expect(result.tableFragments[0].geometry.x).toBeCloseTo(rtl ? 6.25 : 35.75, 0);
      expect(result.tableFragments[0].geometry.width).toBeCloseTo(750, 0);
    });
  }
}

test('split tables keep their indent, merged rows, and visible margin overhangs on every page', async ({ page }) => {
  const rowCount = 40;
  const result = await paginateDocx(page, false, 0.65, rowCount);
  expect(result.totalPages).toBeGreaterThan(2);
  const tables = page.locator('#fixture .page-content table');
  expect(await tables.count()).toBe(result.totalPages);
  expect(await tables.locator('tr').count()).toBe(rowCount);
  expect(await tables.locator('td[colspan="2"]').count()).toBe(5);
  expect((await tables.allTextContents()).join(' ').match(/ROW \d+/g))
    .toEqual(Array.from({ length: rowCount }, (_, i) => `ROW ${i}`));
  expect(result.tableFragments).toHaveLength(result.totalPages);
  for (const fragment of result.tableFragments) {
    expect(fragment.geometry.x).toBeCloseTo(35.75, 0);
    expect(fragment.geometry.width).toBeCloseTo(750, 0);
  }
  for (const table of await tables.all()) {
    await table.scrollIntoViewIfNeeded();
    const visible = await table.evaluate((el) => {
      const r = el.getBoundingClientRect();
      return [r.left + 4, r.right - 4]
        .every((x) => el.contains(document.elementFromPoint(x, r.top + r.height / 2)));
    });
    expect(visible).toBe(true);
  }
});

test('overhangs clip at the paper edges and body bottom, including PageMap geometry', async ({ page }) => {
  await page.setContent('<main id="fixture"></main>');
  await page.addScriptTag({ url: 'http://localhost:8082/pagination.bundle.js' });
  const result = await page.evaluate(() => {
    const html = `<div id="pagination-staging"><div data-section-index="0"
      data-page-width="200" data-page-height="120" data-content-width="160" data-content-height="80"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      <div><table data-source-anchor-id="tbl:body:oversized"
        style="width:260pt; margin-left:-50pt; border-collapse:collapse; table-layout:fixed">
        <tbody><tr style="height:160pt"><td>OVERSIZED</td></tr></tbody>
      </table></div></div></div><div id="pagination-container"></div>`;
    const pagination = (window as any).DocxodusPagination.paginateHtml(html, 'fixture', {
      showPageNumbers: false,
      layoutToken: { documentVersion: 0, rendererFingerprint: 'negative-indent-clipping' },
    });
    const table = document.querySelector('#fixture .page-content table')!;
    const paper = table.closest('.page-box')!.getBoundingClientRect();
    const body = table.closest('.page-content')!.getBoundingClientRect();
    const rect = table.getBoundingClientRect();
    const hits = (x: number, y: number) => table.contains(document.elementFromPoint(x, y));
    return {
      extendsPastPaper: rect.left < paper.left && rect.right > paper.right,
      extendsPastBody: rect.bottom > body.bottom,
      insideLeft: hits(paper.left + 4, body.top + 4),
      insideRight: hits(paper.right - 4, body.top + 4),
      outsideLeft: hits(paper.left - 4, body.top + 4),
      outsideRight: hits(paper.right + 4, body.top + 4),
      aboveBottom: hits(body.left + 4, body.bottom - 4),
      belowBottom: hits(body.left + 4, body.bottom + 4),
      geometry: pagination.pageMap.fragments[0].geometry,
    };
  });
  expect(result.extendsPastPaper).toBe(true);
  expect(result.extendsPastBody).toBe(true);
  expect(result.insideLeft).toBe(true);
  expect(result.insideRight).toBe(true);
  expect(result.outsideLeft).toBe(false);
  expect(result.outsideRight).toBe(false);
  expect(result.aboveBottom).toBe(true);
  expect(result.belowBottom).toBe(false);
  expect(result.geometry.x).toBeCloseTo(0, 0);
  expect(result.geometry.y).toBeCloseTo(20, 0);
  expect(result.geometry.width).toBeCloseTo(200, 0);
  expect(result.geometry.height).toBeCloseTo(80, 0);
});

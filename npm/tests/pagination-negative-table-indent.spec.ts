import { expect, test } from '@playwright/test';
import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

/**
 * Issue #827 — Word pulls an over-wide table into the left margin with a negative w:tblInd.
 * The paginated view must keep the negative indent and must not clip the table at the body
 * column: Word only clips at the paper edge.
 */

// US Letter landscape, 1in margins: a 648pt text column holding a 750pt table indented -36.25pt.
function generateNegativeIndentDocx(): Uint8Array {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>
  <w:tbl>
    <w:tblPr><w:tblW w:w="15000" w:type="dxa"/><w:tblInd w:w="-725" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
    <w:tblGrid><w:gridCol w:w="7500"/><w:gridCol w:w="7500"/></w:tblGrid>
    <w:tr>
      <w:tc><w:tcPr><w:tcW w:w="7500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>LEFT CELL</w:t></w:r></w:p></w:tc>
      <w:tc><w:tcPr><w:tcW w:w="7500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>RIGHT CELL</w:t></w:r></w:p></w:tc>
    </w:tr>
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

test('negative table indent extends into the page margin without being clipped', async ({ page }) => {
  await page.goto('/test-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusReady === true, undefined,
    { timeout: 30_000 });
  await page.addScriptTag({ url: '/pagination.bundle.js' });

  const geometry = await page.evaluate(async (bytes) => {
    const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
      new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
      1, 1, 'page-', false, 0, 'annot-', true, true, false, false, false,
    ) as string;
    if (html.startsWith('{')) throw new Error(`conversion failed: ${html.slice(0, 300)}`);

    document.body.innerHTML = '<main id="fixture"></main>';
    (window as any).DocxodusPagination.paginateHtml(
      html,
      document.getElementById('fixture'),
      { scale: 1, showPageNumbers: false, pageGap: 0 },
    );
    await document.fonts.ready;

    const body = document.querySelector<HTMLElement>('#fixture .page-content')!;
    const table = body.querySelector<HTMLElement>('table')!;
    const bodyRect = body.getBoundingClientRect();
    const tableRect = table.getBoundingClientRect();
    const midY = tableRect.top + tableRect.height / 2;
    // Hit-testing honours overflow clipping, so these prove the overhangs are actually painted.
    const hits = (x: number) => table.contains(document.elementFromPoint(x, midY));
    return {
      leftOverhang: bodyRect.left - tableRect.left,
      leftMarginVisible: hits(tableRect.left + 4),
      rightMarginVisible: hits(tableRect.right - 4),
      extendsPastRight: tableRect.right > bodyRect.right,
    };
  }, Array.from(generateNegativeIndentDocx()));

  const PT_TO_PX = 4 / 3;
  expect(geometry.leftOverhang).toBeCloseTo(36.25 * PT_TO_PX, 0);
  expect(geometry.extendsPastRight).toBe(true);
  expect(geometry.leftMarginVisible).toBe(true);
  expect(geometry.rightMarginVisible).toBe(true);
});

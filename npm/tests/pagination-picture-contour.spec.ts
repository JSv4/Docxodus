import { expect, test } from '@playwright/test';
import { storedZip, xml, R_NS, W_NS } from './docx-zip.js';

function pictureDocx(tight: boolean): Uint8Array {
  const wp = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  const pic = 'http://schemas.openxmlformats.org/drawingml/2006/picture';
  const wrap = tight ? `<wp:wrapTight wrapText="right" distR="114300"><wp:wrapPolygon edited="1">
    <wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/>
    </wp:wrapPolygon></wp:wrapTight>` : '<wp:wrapSquare wrapText="right" distR="114300"/>';
  const words = Array.from({ length: 120 }, (_, i) => `token${i}`).join(' ');
  return storedZip([
    { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
      </Types>`) },
    { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="office" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`) },
    { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="picture" Type="${R_NS}/image" Target="media/black.png"/></Relationships>`) },
    { name: 'word/media/black.png', data: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAC0lEQVR4nGNgQAYAAA4AAamRc7EAAAAASUVORK5CYII=', 'base64') },
    { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${wp}" xmlns:a="${a}" xmlns:pic="${pic}">
      <w:body><w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="68580" distR="68580" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="914400" cy="457200"/>${wrap}<wp:docPr id="1" name="Authored triangle"/>
      <a:graphic><a:graphicData uri="${pic}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="black.png"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
      <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
      </pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r><w:r><w:t>${words}</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
      </w:body></w:document>`) },
  ]);
}

for (const tight of [false, true]) {
  test(`${tight ? 'tight contour' : 'square wrap'} preserves picture placement and all source text`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Document', 'docx-', true, '', -1, 'comment-',
        1, 1, 'page-', false, 0, 'annot-', false, false, false, false, false,
      ) as string;
      if (html.startsWith('{')) throw new Error(html);
      document.body.innerHTML = '<main id="fixture"></main>';
      const pagination = (window as any).DocxodusPagination.paginateHtml(html, document.getElementById('fixture'),
        { showPageNumbers: false, pageGap: 0 });
      const content = document.querySelector<HTMLElement>('#fixture .page-content')!;
      const image = content.querySelector<HTMLImageElement>('img')!;
      const paragraph = content.querySelector('p')!;
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      const starts = new Map<number, number>();
      for (let node = walker.nextNode(); node; node = walker.nextNode()) {
        const text = node.textContent ?? '';
        for (let i = 0; i < text.length; i++) {
          if (/\s/.test(text[i])) continue;
          const range = document.createRange(); range.setStart(node, i); range.setEnd(node, i + 1);
          const rect = range.getBoundingClientRect();
          const row = Math.round(rect.top);
          starts.set(row, Math.min(starts.get(row) ?? Infinity, rect.left - content.getBoundingClientRect().left));
        }
      }
      return { pages: pagination.totalPages, text: paragraph.textContent,
        imageX: image.getBoundingClientRect().left - content.getBoundingClientRect().left,
        imageWidth: image.getBoundingClientRect().width,
        rows: Array.from(starts).sort(([a], [b]) => a - b).map(([, x]) => x),
      };
    }, { bytes: Array.from(pictureDocx(tight)) });
    expect(result.pages).toBe(1);
    expect(result.text?.trim()).toBe(Array.from({ length: 120 }, (_, i) => `token${i}`).join(' '));
    expect(result.imageX).toBeCloseTo(0, 1);
    expect(result.imageWidth * 0.75).toBeCloseTo(72, 1);
    expect(result.rows.length).toBeGreaterThan(3);
    if (tight) expect(result.rows[1]).toBeLessThan(result.rows[0] - 5);
    else expect(result.rows[1]).toBeCloseTo(result.rows[0], 1);
    expect(result.rows[result.rows.length - 1]).toBeCloseTo(0, 1);
  });
}

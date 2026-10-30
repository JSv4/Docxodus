import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { pageFittingDocx } from './docx-page-fitting-fixture.js';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
test.use({ launchOptions: { chromiumSandbox: true, ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });
for (const distance of [4294964756, -2540, 25400]) {
  test(`anchored rule clearance ${distance} preserves the whole document`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const expected = distance === 25400 ? 2 : -0.2;
    const drawing = `<w:p><w:r><w:drawing><wp:anchor distT="${distance}" distB="${distance}" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>38100</wp:posOffset></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>152400</wp:posOffset></wp:positionV><wp:extent cx="1219200" cy="0"/>
      <wp:wrapTopAndBottom/><wp:docPr id="7" name="Juniper inventory rule"/>
      <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="0"/></a:xfrm>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="12700"><a:solidFill><a:srgbClr val="315E42"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/>
      </wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>`;
    const input = pageFittingDocx(`<w:p><w:r><w:t>Juniper signed-clearance inventory.</w:t></w:r></w:p>${drawing}<w:p><w:r><w:t>Juniper signed-clearance complete.</w:t></w:r></w:p>`);
    const result = await renderDocxArtifacts(input, { browser, outputs: ['html', 'pdf'], reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin', commentColumnPlacement: 'balloons', timeoutMs: 120_000 });
    expect(result.pageMap.pages).toHaveLength(1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      expect(pdf.numPages).toBe(1);
      const text = (await (await pdf.getPage(1)).getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
      expect(text).toContain('Juniper signed-clearance inventory.');
      expect(text).toContain('Juniper signed-clearance complete.');
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    await expect(page.locator('.page-box [data-docx-shape-anchor] svg')).toHaveCount(1);
    const geometry = await page.locator('.page-box [data-docx-shape-anchor]').evaluate(element => ({
      top: element.getAttribute('data-docx-anchor-wrap-top'), bottom: element.getAttribute('data-docx-anchor-wrap-bottom'),
      width: element.getBoundingClientRect().width * .75,
    }));
    expect(Number(geometry.top)).toBe(expected); expect(Number(geometry.bottom)).toBe(expected); expect(geometry.width).toBeCloseTo(96, 1);
  });
}

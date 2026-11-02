import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { writeFileSync } from 'node:fs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

function fixture(rtl: boolean, bordered: boolean) {
  const border = bordered ? '<w:pBdr><w:top w:val="single" w:sz="8"/></w:pBdr>' : '';
  const direction = rtl ? '<w:bidi/>' : '';
  const body = [-540, 0, 540].map((indent, i) => `<w:p><w:pPr>
    ${border}${direction}<w:ind w:left="${indent}"/></w:pPr>
    <w:r><w:t>Laurel${i}</w:t></w:r></w:p>`).join('') +
    `<w:p><w:pPr>${border}${direction}<w:ind w:hanging="180"/></w:pPr>
      <w:r><w:t>HangingFirst</w:t><w:br/><w:t>HangingNext</w:t></w:r></w:p>`;
  const parts = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      ${['document', 'styles', 'settings'].map(p => `<Override PartName="/word/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${p}${p === 'document' ? '.main' : ''}+xml"/>`).join('')}</Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings'].map(p => `<Relationship Id="${p}" Type="${R_NS}/${p}" Target="${p}.xml"/>`).join('')}</Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

for (const rtl of [false, true]) for (const bordered of [false, true]) {
  test(`${rtl ? 'RTL' : 'LTR'} ${bordered ? 'bordered' : 'plain'} paragraph keeps signed leading PDF origins`, async ({ browser }, testInfo) => {
    test.setTimeout(120_000);
    const input = fixture(rtl, bordered);
    const actual = await renderDocxArtifacts(input, { browser, outputs: ['pdf'],
      reviewProfile: 'final', commentProfile: 'hidden', timeoutMs: 120_000 });
    expect(actual.pageCount).toBe(1);
    writeFileSync(testInfo.outputPath('signed.docx'), input);
    writeFileSync(testInfo.outputPath('signed.pdf'), actual.pdf!);
    const task = getDocument({ data: new Uint8Array(actual.pdf!), verbosity: 0 });
    try {
      const page = await (await task.promise).getPage(1);
      const items = (await page.getTextContent()).items.filter(item => 'str' in item).filter(item => item.str.length > 0);
      expect(items).toHaveLength(5);
      for (let i = 0; i < 3; i++) {
        const item = items.find(item => item.str === `Laurel${i}`)!;
        const edge = item.transform[4] + (rtl ? item.width : 0);
        expect(edge).toBeCloseTo(rtl ? 567 - i * 27 : 45 + i * 27, 0);
      }
      for (const [text, offset] of [['HangingFirst', -9], ['HangingNext', 0]] as const) {
        const item = items.find(item => item.str === text)!;
        expect(item.transform[4] + (rtl ? item.width : 0)).toBeCloseTo(rtl ? 540 - offset : 72 + offset, 0);
      }
    } finally { await task.destroy(); }
  });
}

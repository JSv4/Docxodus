// Issue #1704. compareDocuments failed about once in a hundred calls after the runtime had
// warmed up: "Comparison failed - empty result", from an InvalidCastException thrown while the
// renderer deep-cloned a freshly parsed styles part (the step it takes when the two documents'
// w:docDefaults differ). The comparison was not at fault. A nursery collection landing inside
// that clone, with the jiterpreter's interp-entry thunks carrying the calls from compiled code
// into interpreted code, left a frame holding a stale object reference.
//
// The failure needs a collection to land in a narrow window, so this spec makes collections
// frequent (a 1 MB nursery instead of the default) and repeats the comparison. Against the
// runtime built before the fix this test failed in three runs of three, and the same loop under
// Node lost 9 and 11 calls of 600; with interp-entry thunks off, which is what
// scripts/build-wasm.sh now ships, 6,600 such calls have not lost one.
// wasm-runtime-config.spec.ts pins the switch itself; this pins what the switch is for.
import { test, expect, Page } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const COMPARISONS = 400;

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// One heading, a table and a sentence; the two versions differ in that sentence and in the
// document-default font size. Authored here.
function rota(fontSize: number, closing: string): Uint8Array {
  const cell = (text: string) => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="${fontSize}"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault></w:docDefaults>`
      + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>`
      + `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>`
      + `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Rota for the loading dock</w:t></w:r></w:p>`
      + `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>`
      + `<w:tr>${cell('Shift')}${cell('Lead')}</w:tr><w:tr>${cell('Early')}${cell('Priya')}</w:tr><w:tr>${cell('Late')}${cell('Tomasz')}</w:tr></w:tbl>`
      + `<w:p><w:r><w:t>${closing}</w:t></w:r></w:p>`
      + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
  };
  return storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) })));
}

test.describe('WASM comparison under frequent garbage collection (#1704)', () => {
  test(`${COMPARISONS} comparisons of one pair with differing document defaults all succeed`, async ({ page }) => {
    test.setTimeout(600000);
    await page.goto('/test-harness.html?gcParams=nursery-size%3D1m');
    await waitForDocxodus(page);

    const failed = await page.evaluate(([left, right, count]) => {
      const comparer = (window as any).Docxodus.DocumentComparer;
      const a = new Uint8Array(left), b = new Uint8Array(right);
      const failures: number[] = [];
      for (let i = 0; i < count; i++)
        if (comparer.CompareDocuments(a, b, null).length === 0)
          failures.push(i);
      return failures;
    }, [
      Array.from(rota(22, 'Swap requests go to the duty manager by Thursday.')),
      Array.from(rota(24, 'Swap requests go to the duty manager by Wednesday noon.')),
      COMPARISONS,
    ] as const);

    expect(failed, 'indices of the comparisons that returned an empty result').toEqual([]);
  });
});

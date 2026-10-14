import { expect, test, type Page } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { storedZip, W_NS, xml } from './docx-zip.js';

/**
 * A paragraph's own line box is sized from the configured face of its family even when the
 * paragraph's weight or style differs from its runs' (issue #956).
 *
 * The export points every element whose font matches a resolved request at that request's face, so
 * a paragraph's strut, its share of each line box and the `lh` its auto line multiple is built on,
 * use the configured font. The match used to require the same weight and style. A heading's `h1`
 * inherits the browser's bold while its runs are normal, and a paragraph of bold runs keeps a normal
 * mark, so neither paragraph matched its runs' request. Its strut fell back to whatever the machine
 * resolves the family to: in the export benchmark, a Calibri Light heading laid out at Liberation
 * Sans's line pitch instead of the contract's Carlito.
 *
 * The family here is one no machine has, served through the export's font resolver as the
 * repository's test font, so a fallback is always detectable. Its `hhea` metrics give a natural line
 * height of (1901 + 483 + 0) / 2048 em.
 */
const FAMILY = 'Docxodus Strut Probe';
/** The document default, a different family, so a paragraph cannot inherit the probe's face from its section. */
const DEFAULT_FAMILY = 'Docxodus Strut Default';
const NATURAL_EM = 2384 / 2048;
const fontBytes = readFileSync(join(dirname(fileURLToPath(import.meta.url)),
  '..', '..', 'docs', 'demo', 'fonts', 'docxodus-canvas-mono.woff2'));
const fontPlan = {
  mode: 'exact',
  format: 'woff2',
  mediaType: 'font/woff2',
  byteLength: fontBytes.byteLength,
  sha256: createHash('sha256').update(fontBytes).digest('hex'),
  bytesBase64: fontBytes.toString('base64'),
  licenseIdentity: 'a'.repeat(64),
};
const TOLERANCE_PT_PER_LINE = 0.05;
const TEXT = 'The quick brown fox jumps over the lazy dog while the export counts every line. '.repeat(8);

/**
 * A Heading 1 paragraph or a paragraph of bold runs (one per document, so neither weight's request
 * exists for the other to borrow), 16 pt at Word's 1.08 auto spacing, each
 * with its family set on its own style as Word's headings carry theirs, so neither inherits a face.
 */
function docx(which: 'HEADING' | 'BOLD'): Uint8Array {
  const spacing = '<w:spacing w:before="0" w:after="0" w:line="259" w:lineRule="auto"/>';
  return storedZip([
    {
      name: '[Content_Types].xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`),
    },
    {
      name: '_rels/.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: 'word/_rels/document.xml.rels',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`),
    },
    {
      name: 'word/styles.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
  <w:rFonts w:ascii="${DEFAULT_FAMILY}" w:hAnsi="${DEFAULT_FAMILY}" w:cs="${DEFAULT_FAMILY}" w:eastAsia="${DEFAULT_FAMILY}"/>
  <w:sz w:val="32"/><w:szCs w:val="32"/>
</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
  <w:pPr><w:outlineLvl w:val="0"/></w:pPr>
  <w:rPr><w:rFonts w:ascii="${FAMILY}" w:hAnsi="${FAMILY}"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="Body"/>
  <w:rPr><w:rFonts w:ascii="${FAMILY}" w:hAnsi="${FAMILY}"/></w:rPr></w:style></w:styles>`),
    },
    {
      name: 'word/document.xml',
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  ${which === 'HEADING'
    ? `<w:p><w:pPr><w:pStyle w:val="Heading1"/>${spacing}</w:pPr>
    <w:r><w:t xml:space="preserve">HEADING ${TEXT}</w:t></w:r></w:p>`
    : `<w:p><w:pPr><w:pStyle w:val="Body"/>${spacing}</w:pPr>
    <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">BOLD ${TEXT}</w:t></w:r></w:p>`}
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
  </w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

/** Mean distance between consecutive line tops of the block whose text starts with `prefix`, in pt. */
async function pitches(page: Page, which: 'HEADING' | 'BOLD'): Promise<Record<string, number>> {
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, plan }) =>
    (window as any).DocxodusStandalone.convertWithFontResolver(
      bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, plan),
  { bytes: Array.from(docx(which)), plan: fontPlan });
  expect(result.renderReport.status).toBe('complete');
  await page.setContent(result.html);
  return page.evaluate(() => {
    const out: Record<string, number> = {};
    for (const block of Array.from(document.querySelectorAll('h1, p'))) {
      const prefix = (block.textContent ?? '').trimStart().split(' ')[0];
      if (prefix !== 'HEADING' && prefix !== 'BOLD') continue;
      const range = document.createRange();
      range.selectNodeContents(block);
      const tops = [...new Set(Array.from(range.getClientRects())
        .filter((rect) => rect.width > 0)
        .map((rect) => Math.round(rect.top * 64) / 64))].sort((a, b) => a - b);
      out[prefix] = ((tops[tops.length - 1] - tops[0]) / (tops.length - 1)) * 0.75;
    }
    return out;
  });
}

test.describe('a paragraph sizes its lines from its configured face (#956)', () => {
  test('a heading and a paragraph of bold runs keep the font metrics', async ({ page }) => {
    const expected = NATURAL_EM * 16 * (259 / 240);

    for (const prefix of ['HEADING', 'BOLD'] as const) {
      const measured = await pitches(page, prefix);
      expect(measured[prefix], `${prefix} block not found`).toBeDefined();
      expect(Math.abs(measured[prefix] - expected), `${prefix}: ${measured[prefix]} pt vs ${expected} pt`)
        .toBeLessThanOrEqual(TOLERANCE_PT_PER_LINE);
    }
  });
});

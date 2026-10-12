import { expect, test } from '@playwright/test';
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { storedZip, W_NS, xml } from './docx-zip.js';

const texts = Array.from({ length: 42 }, (_, i) =>
  `Record${String(i).padStart(3, '0')} AV A plain paragraph retains its source text.`);
function sourceFor(flow: 'single' | 'balanced' | 'continuous' | 'explicit'): Uint8Array { return storedZip([
  { name: '[Content_Types].xml', data: xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    </Types>`) },
  { name: '_rels/.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`) },
  { name: 'word/_rels/document.xml.rels', data: xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    </Relationships>`) },
  { name: 'word/styles.xml', data: xml(`<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>
    </w:rPr></w:rPrDefault></w:docDefaults></w:styles>`) },
  { name: 'word/document.xml', data: xml(`<w:document xmlns:w="${W_NS}"><w:body>
    ${texts.map((text, index) => `<w:p><w:pPr><w:keepLines/><w:spacing w:before="0" w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
      ${flow === 'explicit' && index === 18 ? '<w:r><w:br w:type="column"/></w:r>' : ''}
      <w:r><w:t>${text}</w:t></w:r></w:p>`).join('')}
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      ${flow === 'single' ? '' : '<w:cols w:num="2" w:space="480"/>'}
      ${flow === 'continuous' || flow === 'explicit' ? '<w:type w:val="continuous"/>' : ''}
    </w:sectPr>
    </w:body></w:document>`) },
]); }
const font = readFileSync(new URL('fixtures/fonts/baseline-test-b.woff2', import.meta.url));
const digest = (bytes: Uint8Array | string) => createHash('sha256').update(bytes).digest('hex');

for (const flow of ['single', 'balanced', 'continuous', 'explicit'] as const) {
test(`precision ${flow} text survives public export, offline reopen and PDF with its font and PageMap identities`, async ({ page, context }, testInfo) => {
  const source = sourceFor(flow);
  await page.goto('/standalone-export-harness.html');
  await page.waitForFunction(() => (window as any).DocxodusStandaloneReady === true);
  const result = await page.evaluate(async ({ bytes, plan }) =>
    (window as any).DocxodusStandalone.convertWithFontResolver(
      bytes, { reviewProfile: 'final', commentProfile: 'hidden' }, plan),
  { bytes: Array.from(source), plan: {
    mode: 'exact', format: 'woff2', mediaType: 'font/woff2',
    byteLength: font.byteLength, sha256: digest(font), bytesBase64: font.toString('base64'),
    licenseIdentity: 'b'.repeat(64),
  } });
  expect(result.renderReport.status).toBe('complete');
  expect(result.pageCount).toBe(2);
  expect(result.renderReport.source.rawPackageBytesDigest).toBe(digest(source));
  expect(result.renderReport.bindings.htmlDigest).toBe(digest(result.html));
  expect(result.renderReport.fonts.length).toBeGreaterThan(0);
  expect(result.renderReport.fonts.every((face: any) =>
    face.source === 'configured' && face.fileSha256 === digest(font))).toBe(true);
  expect(result.pageMap.fragments).toHaveLength(texts.length);
  expect(result.pageMap.rendererFingerprint).toBe(result.rendererFingerprint);

  const offlinePath = testInfo.outputPath('precision-text.html');
  writeFileSync(offlinePath, result.html);
  const offline = await context.newPage();
  await offline.goto(pathToFileURL(offlinePath).href);
  await offline.evaluate(() => document.fonts.ready);
  const reopened = await offline.evaluate(() => Array.from(document.querySelectorAll<HTMLElement>(
    '.page-box p[data-source-anchor-id]')).map(p => {
    const page = p.closest<HTMLElement>('.page-box')!;
    const box = p.getBoundingClientRect(); const pageBox = page.getBoundingClientRect();
    return { text: p.textContent, painted: p.querySelectorAll('svg').length,
      anchorId: p.dataset.sourceAnchorId, pageNumber: Number(page.dataset.pageNumber),
      geometry: { x: (box.left - pageBox.left) * 0.75, y: (box.top - pageBox.top) * 0.75,
        width: box.width * 0.75, height: box.height * 0.75 } };
  }));
  expect(reopened.map(p => p.text)).toEqual(texts);
  expect(reopened.every(p => p.painted === 1)).toBe(true);
  if (flow === 'explicit') {
    expect(reopened[18].geometry.x - reopened[0].geometry.x).toBeCloseTo(246, 0);
    expect(reopened[36].pageNumber).toBe(2);
  }
  for (const [index, p] of reopened.entries()) {
    const fragment = result.pageMap.fragments[index];
    expect(fragment.anchorId).toBe(p.anchorId);
    expect(fragment.pageNumber).toBe(p.pageNumber);
    for (const dimension of ['x', 'y', 'width', 'height'] as const)
      expect(Math.abs(fragment.geometry[dimension] - p.geometry[dimension])).toBeLessThan(0.1);
  }

  const pdf = await offline.pdf({ preferCSSPageSize: true, printBackground: true });
  writeFileSync(testInfo.outputPath('precision-text.pdf'), pdf);
  const loadingTask = getDocument({ data: new Uint8Array(pdf), verbosity: 0 });
  const printed = await loadingTask.promise;
  expect(printed.numPages).toBe(result.pageCount);
  let inventory = '';
  for (let index = 1; index <= printed.numPages; index++) {
    const pdfPage = await printed.getPage(index);
    for (const item of (await pdfPage.getTextContent()).items)
      if ('str' in item) inventory += item.str + (item.hasEOL ? '\n' : '');
    inventory += '\n';
  }
  // PDF parsers can split a word into several drawing items. Concatenate those
  // items before tokenizing, retaining their actual spaces and line endings.
  expect(inventory.split(/\s+/).filter(Boolean)).toEqual(texts.join(' ').split(/\s+/));
  await loadingTask.destroy();
  await offline.close();
});
}

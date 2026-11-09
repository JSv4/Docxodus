import { test, expect } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original minimal body fields also prove this option alone selects the complete WASM entry point.
const bytes = Array.from(storedZip(Object.entries({
  '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
  '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
  'word/styles.xml': `<w:styles xmlns:w="${W_NS}"/>`,
  'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p>
    <w:fldSimple w:instr=" PAGE \\* roman "><w:r><w:rPr><w:i/></w:rPr><w:t>iii</w:t></w:r></w:fldSimple>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> NUMPAGES </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>
    </w:p><w:sectPr/></w:body></w:document>`,
}).map(([name, value]) => ({ name, data: xml(value) }))));

for (const worker of [false, true]) {
  test(`flow page fields cross the ${worker ? 'worker' : 'main-thread'} API`, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    const results = await page.evaluate(async ({ input, worker }) => {
      const coreModule = 'http://localhost:8083/core.js';
      const api = worker
        ? await (window as any).createDocxodusWorker()
        : await import(coreModule);
      if (!worker) await api.initialize(`${location.origin}/`);
      const shapes = [];
      for (const stampPageNumberFields of [undefined, false, true]) {
        const html = await api.convertDocxToHtml(new Uint8Array(input),
          stampPageNumberFields === undefined ? undefined : { stampPageNumberFields });
        const doc = new DOMParser().parseFromString(html, 'text/html');
        shapes.push({
          text: doc.body.textContent!.trim(),
          fields: Array.from(doc.querySelectorAll('[data-field]')).map(field => ({
            kind: field.getAttribute('data-field'), format: field.getAttribute('data-field-format'),
            text: field.textContent,
          })),
        });
      }
      if (worker) api.terminate();
      return shapes;
    }, { input: bytes, worker });
    expect(results).toEqual([
      { text: 'iii', fields: [] },
      { text: 'iii', fields: [] },
      { text: 'iii', fields: [
        { kind: 'PAGE', format: 'roman', text: 'iii' },
        { kind: 'NUMPAGES', format: null, text: '' },
      ] },
    ]);
  });
}

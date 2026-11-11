import { test, expect } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const input = Array.from(storedZip(Object.entries({
  '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
  '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="settings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
  'word/styles.xml': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:pPrDefault/></w:docDefaults></w:styles>`,
  'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Anchor</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ed</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:bCs/></w:rPr><w:t>Ord</w:t></w:r><w:r><w:t>inary</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r></w:p>
    <w:sectPr/></w:body></w:document>`,
}).map(([name, value]) => ({ name, data: xml(value) }))));

for (const worker of [false, true]) {
  test(`ordinary text joins through the ${worker ? 'worker' : 'main-thread'} API`, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    const result = await page.evaluate(async ({ input, worker }) => {
      const moduleUrl = 'http://localhost:8083/core.js';
      const api = worker ? await (window as any).createDocxodusWorker() : await import(moduleUrl);
      if (!worker) await api.initialize(`${location.origin}/`);
      const html = await api.convertDocxToHtml(new Uint8Array(input), { stampAnchors: true });
      if (worker) api.terminate();
      const doc = new DOMParser().parseFromString(html, 'text/html');
      return Array.from(doc.querySelectorAll('p')).map(p => ({
        anchored: p.hasAttribute('data-anchor'),
        spans: Array.from(p.children).map(span => span.textContent),
      }));
    }, { input, worker });
    expect(result).toEqual([
      { anchored: true, spans: ['Anchored'] },
      { anchored: true, spans: ['Ordinary'] },
      { anchored: true, spans: ['Bold', 'Italic'] },
    ]);
  });
}

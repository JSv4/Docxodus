import { test, expect } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

const parts: Record<string, string> = {
  'word/styles.xml': `<w:styles xmlns:w="${W_NS}"/>`,
  'word/settings.xml': `<w:settings xmlns:w="${W_NS}"><w:evenAndOddHeaders/></w:settings>`,
  'word/numbering.xml': `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
};
const types: Record<string, string> = { document: 'document.main', styles: 'styles', settings: 'settings', numbering: 'numbering' };
let relationships = ['styles', 'settings', 'numbering'].map(name =>
  `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('');
let references = '';
const variants = ['header-default', 'header-first', 'header-even', 'footer-default', 'footer-first', 'footer-even'];
for (const name of variants) {
  const [kind, variant] = name.split('-');
  const tag = kind === 'header' ? 'hdr' : 'ftr';
  const field = name === 'header-default' ? '<w:fldSimple w:instr=" PAGE "/>' : '';
  parts[`word/${name}.xml`] = `<w:${tag} xmlns:w="${W_NS}"><w:p><w:r><w:t>${name}</w:t></w:r>${field}</w:p></w:${tag}>`;
  types[name] = kind;
  references += `<w:${kind}Reference w:type="${variant}" r:id="${name}"/>`;
  relationships += `<Relationship Id="${name}" Type="${R_NS}/${kind}" Target="${name}.xml"/>`;
}
parts['word/document.xml'] = `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>ListItem</w:t></w:r></w:p><w:sectPr>${references}<w:titlePg/></w:sectPr></w:body></w:document>`;
parts['word/_rels/document.xml.rels'] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`;
parts['_rels/.rels'] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`;
parts['[Content_Types].xml'] = '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + Object.entries(types).map(([name, type]) => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>`).join('') + '</Types>';
const input = Array.from(storedZip(Object.entries(parts).map(([name, value]) => ({ name, data: xml(value) }))));

for (const worker of [false, true]) {
  test(`flow variants preserve semantic lists and fields through ${worker ? 'worker' : 'main-thread'} API`, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    const result = await page.evaluate(async ({ input, worker }) => {
      const moduleUrl = 'http://localhost:8083/core.js';
      const api = worker ? await (window as any).createDocxodusWorker() : await import(moduleUrl);
      if (!worker) await api.initialize(`${location.origin}/`);
      const shapes = [];
      for (const renderHeaderFooterRegistry of [false, true]) {
        const html = await api.convertDocxToHtml(new Uint8Array(input), {
          renderHeadersAndFooters: true, renderHeaderFooterRegistry, semanticLists: true, stampPageNumberFields: true,
        });
        const doc = new DOMParser().parseFromString(html, 'text/html');
        shapes.push({
          stories: doc.querySelectorAll('header, footer').length,
          variants: Array.from(doc.querySelectorAll('[data-hf-type]')).map(e => e.getAttribute('data-hf-type')).sort(),
          lists: doc.querySelectorAll('ol > li').length,
          fields: Array.from(doc.querySelectorAll('[data-field]')).map(e => e.getAttribute('data-field')),
          paginated: !!doc.querySelector('#pagination-staging'),
        });
      }
      if (worker) api.terminate();
      return shapes;
    }, { input, worker });
    expect(result).toEqual([
      { stories: 2, variants: [], lists: 1, fields: ['PAGE'], paginated: false },
      { stories: 0, variants: [...variants].sort(), lists: 1, fields: ['PAGE'], paginated: false },
    ]);
  });
}

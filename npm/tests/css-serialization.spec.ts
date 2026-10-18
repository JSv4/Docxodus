import { test, expect } from '@playwright/test';
import { storedZip, xml, W_NS, R_NS } from './docx-zip.js';

// Original, schema-valid minimal document. The adjacent plain paragraph is a control.
function cascadeFixture(): Uint8Array {
  return storedZip(Object.entries({
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="${R_NS}/styles" Target="styles.xml"/><Relationship Id="rSettings" Type="${R_NS}/settings" Target="settings.xml"/></Relationships>`,
    'word/document.xml': `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>Cedar cascade sample.</w:t></w:r></w:p><w:p><w:r><w:t>Plain control.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    'word/styles.xml': `<w:styles xmlns:w="${W_NS}"/>`,
    'word/settings.xml': `<w:settings xmlns:w="${W_NS}"/>`,
  }).map(([name, value]) => ({ name, data: xml(value) })));
}

for (const paginated of [false, true]) {
  test(`WASM CSS child selectors apply in ${paginated ? 'paginated' : 'continuous'} HTML`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const observed = await page.evaluate(async ({ bytes, paginated }) => {
      const css = 'p:first-child > span { color: rgb(255, 0, 0) !important; background-color: rgb(0, 255, 0); }' +
        'p:first-child::after { content: "A&B > C"; }' +
        'p:nth-child(2)::after { content: "</StYlE><p id=escape>A&B ]]> C"; }';
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cascade', 'docx-', false, css, -1, 'comment-',
        paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      const host = document.createElement('div'); document.body.replaceChildren(host);
      if (paginated) {
        const { paginateHtml } = (window as any).DocxodusPagination;
        paginateHtml(html, host, { showPageNumbers: false });
      } else host.innerHTML = html;
      const paragraphs = host.querySelectorAll(paginated ? '.page-content p' : 'p');
      const target = paragraphs[0].querySelector('span')!;
      const control = paragraphs[1].querySelector('span')!;
      return { color: getComputedStyle(target).color, background: getComputedStyle(target).backgroundColor,
        control: getComputedStyle(control).color,
        content: getComputedStyle(paragraphs[0], '::after').content,
        escapedContent: getComputedStyle(paragraphs[1], '::after').content,
        injected: host.querySelector('#escape') !== null };
    }, { bytes: Array.from(cascadeFixture()), paginated });
    expect(observed.color).toBe('rgb(255, 0, 0)');
    expect(observed.background).toBe('rgb(0, 255, 0)');
    expect(observed.control).not.toBe('rgb(255, 0, 0)');
    expect(observed.content).toBe('"A&B > C"');
    expect(observed.escapedContent).toBe('"</StYlE><p id=escape>A&B ]]> C"');
    expect(observed.injected).toBe(false);
  });
}

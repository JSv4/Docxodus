import { expect, test } from '@playwright/test';

for (const shape of ['hyperlink', 'field', 'mixed']) {
  test(`renders deep ${shape} wrappers through the shipping WASM converter`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const result = await page.evaluate((shape) => {
      const D = (window as any).Docxodus;
      const session = D.openTypedSession(D.DocxSessionBridge.CreateBlankDocx());
      let bytes: Uint8Array;
      try {
        const anchor = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'))!;
        const start = session.raw.getXml(anchor).match(/^<w:p\b[^>]*>/)![0].replace(/\/>$/, '>');
        let content = '<w:r><w:t>Willow browser path.</w:t></w:r>';
        for (let i = 0; i < 64; i++) {
          const kind = shape === 'mixed' ? ['hyperlink', 'customXml', 'sdt'][i % 3] : shape;
          content = kind === 'hyperlink' ? `<w:hyperlink w:anchor="Target">${content}</w:hyperlink>`
            : kind === 'field' ? `<w:fldSimple w:instr=" DOCPROPERTY Title ">${content}</w:fldSimple>`
              : kind === 'customXml' ? `<w:customXml w:element="clause">${content}</w:customXml>`
                : `<w:sdt><w:sdtPr><w:id w:val="${i + 1}"/></w:sdtPr><w:sdtContent>${content}</w:sdtContent></w:sdt>`;
        }
        const edit = session.raw.replaceXml(anchor, start + content + '</w:p>');
        if (!edit.success) throw new Error(JSON.stringify(edit));
        bytes = session.save();
      } finally { session.close(); }
      const start = performance.now();
      const html = D.DocumentConverter.ConvertDocxToHtml(bytes);
      const milliseconds = performance.now() - start;
      const text = new DOMParser().parseFromString(html, 'text/html').body.textContent ?? '';
      return { milliseconds, occurrences: text.split('Willow browser path.').length - 1 };
    }, shape);
    expect(result.occurrences).toBe(1);
    expect(result.milliseconds).toBeLessThan(10_000);
  });
}

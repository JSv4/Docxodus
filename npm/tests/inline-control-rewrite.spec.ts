import { expect, test } from '@playwright/test';
import type { DocxSession } from '../src/session.js';

for (const citation of [false, true]) for (const outsideEdit of [false, true]) {
  test(`paragraph rewrite keeps ${citation ? 'citation' : 'bound'} control on ${outsideEdit ? 'outside edit' : 'echo'}`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    const result = await page.evaluate(({ citation, outsideEdit }) => {
      const D = (window as any).Docxodus;
      let session: DocxSession = D.openTypedSession(D.DocxSessionBridge.CreateBlankDocx());
      try {
        const anchor = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'))!;
        const start = session.raw.getXml(anchor)!.match(/^<w:p\b[^>]*>/)![0].replace(/\/>$/, '>');
        const fieldStart = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> CITATION Willow </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>';
        const control = '<w:sdt><w:sdtPr><w:id w:val="7"/>' + (citation ? '<w:citation/>'
          : '<w:dataBinding w:xpath="/record[1]/version[1]" w:storeItemID="{00000000-0000-0000-0000-000000000001}"/><w:text/>')
          + '</w:sdtPr><w:sdtContent>' + (citation ? fieldStart : '') + '<w:r><w:t>3.1</w:t></w:r>'
          + (citation ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : '') + '</w:sdtContent></w:sdt>';
        const setup = session.raw.replaceXml(anchor, start + '<w:r><w:t xml:space="preserve">Version </w:t></w:r>'
          + control + '<w:r><w:t xml:space="preserve"> approved.</w:t></w:r></w:p>');
        if (!setup.success) throw new Error(JSON.stringify(setup));
        const source = session.save(); session.close(); session = D.openTypedSession(source);
        const target = Object.keys(session.project().anchorIndex).find(id => id.startsWith('p:'))!;
        const original = session.listContentControls()[0];
        const before = session.raw.getXml(original.anchorId);
        const payload = outsideEdit ? 'Updated version 3.1 approved today.' : 'Version 3.1 approved.';
        const edit = session.replaceText(target, payload);
        const saved = session.save(); session.close(); session = D.openTypedSession(saved);
        const controls = session.listContentControls();
        return { edit, count: controls.length, before,
          after: controls.length === 1 ? session.raw.getXml(controls[0].anchorId) : null,
          markdown: session.project().markdown, payload };
      } finally { session.close(); }
    }, { citation, outsideEdit });
    expect(result.edit.success, JSON.stringify(result.edit.error)).toBe(true);
    expect(result.count).toBe(1);
    expect(result.after).toBe(result.before);
    expect(result.markdown).toContain(result.payload);
  });
}

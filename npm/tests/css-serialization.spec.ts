import { test, expect } from '@playwright/test';
import { cssCascadeFixture as cascadeFixture } from './docx-css-cascade-fixture.js';


for (const paginated of [false, true]) for (const escaped of [false, true]) {
  test(`WASM CSS child selectors apply in ${paginated ? 'paginated' : 'continuous'} HTML (escaped=${escaped})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const observed = await page.evaluate(async ({ bytes, paginated, escaped }) => {
      const closing = JSON.stringify('</StYlE><p id=escape>A&B ]]> C').replace(/</g, escaped ? '\\<' : '<');
      const css = 'p:first-child > span { color: rgb(255, 0, 0) !important; background-color: rgb(0, 255, 0); }' +
        'p:first-child::after { content: "A&B > C"; }' +
        'p:nth-child(2)::after { content: ' + closing + '; }';
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
    }, { bytes: Array.from(cascadeFixture()), paginated, escaped });
    expect(observed.color).toBe('rgb(255, 0, 0)');
    expect(observed.background).toBe('rgb(0, 255, 0)');
    expect(observed.control).not.toBe('rgb(255, 0, 0)');
    expect(observed.content).toBe('"A&B > C"');
    expect(observed.escapedContent).toBe('"</StYlE><p id=escape>A&B ]]> C"');
    expect(observed.injected).toBe(false);
  });
}

import { test, expect } from '@playwright/test';
import { changeBarFixture } from './docx-change-bar-fixture.js';

for (const paginated of [false, true]) for (const classes of [false, true]) {
  test(`change bars stay outside the column (${paginated ? 'paginated' : 'continuous'}, classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const observed = await page.evaluate(({ bytes, paginated, classes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar ledger', 'docx-', classes, '', -1, 'comment-',
        paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, true, true, true,
        false, null, false, 1);
      if (html.startsWith('{')) throw new Error(html);
      const host = document.createElement('div'); document.body.replaceChildren(host);
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      else host.innerHTML = html;
      return Array.from(host.querySelectorAll(paginated ? '.page-content .rev-changed-line' : '.rev-changed-line'))
        .map(p => ({ left: Number.parseFloat(getComputedStyle(p, '::before').left),
          paragraphLeft: p.getBoundingClientRect().left }));
    }, { bytes: Array.from(changeBarFixture()), paginated, classes });
    expect(observed).toHaveLength(2);
    expect(observed.map(p => p.left)).toEqual([-40, -88]);
    expect(observed[0].paragraphLeft + observed[0].left)
      .toBeCloseTo(observed[1].paragraphLeft + observed[1].left, 1);
  });
}

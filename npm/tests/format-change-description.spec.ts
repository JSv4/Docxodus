import { test, expect } from '@playwright/test';
import { formatRevisionFixture } from './docx-format-revision-fixture.js';

for (const paginated of [false, true]) for (const classes of [false, true]) {
  test(`formatting labels describe only effective changes (${paginated ? 'paginated' : 'continuous'}, classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const labels = await page.evaluate(({ bytes, paginated, classes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar revisions', 'docx-', classes, '', -1, 'comment-',
        paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, true, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const host = document.createElement('div'); document.body.replaceChildren(host);
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      else host.innerHTML = html;
      const root = paginated ? host.querySelector('.page-content')! : host;
      return { descriptions: Array.from(root.querySelectorAll('.rev-format-change')).map(e => e.getAttribute('title')),
        families: Array.from(root.querySelectorAll('.rev-format-change')).map(e => getComputedStyle(e.querySelector('span') ?? e).fontFamily),
        control: Array.from(root.querySelectorAll('p')).find(p => p.textContent!.startsWith('PlainExample'))!
          .querySelector('.rev-format-change') !== null };
    }, { bytes: Array.from(formatRevisionFixture()), paginated, classes });
    expect(labels.descriptions).toEqual(['Bold added', 'Font size changed', 'Font changed']);
    expect(labels.control).toBe(false);
    expect(labels.families[2]).toContain('Liberation Sans');
  });
}

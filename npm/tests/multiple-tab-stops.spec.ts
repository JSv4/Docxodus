import { test, expect } from '@playwright/test';
import { multipleTabFixture, tabEntries } from './docx-multiple-tab-fixture.js';

for (const paginated of [false, true]) for (const classes of [false, true]) {
  test(`manual tab labels retain left titles and right page numbers (${paginated ? 'paginated' : 'continuous'}, classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const results = await page.evaluate(({ bytes, entries, paginated, classes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar tabs', 'docx-', classes, '', -1, 'comment-',
        paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      else {
        host.innerHTML = parsed.body.innerHTML;
        document.body.style.cssText = 'margin:0;padding:0';
        host.style.cssText = 'width:6.5in;margin:0;padding:0';
      }
      const root = paginated ? host.querySelector<HTMLElement>('.page-content')! : host;
      const origin = root.getBoundingClientRect().left;
      return entries.map(([label, title, number]) => {
        const paragraph = Array.from(root.querySelectorAll('p')).find(p => p.textContent?.includes(label))!;
        const leaf = (text: string) => Array.from(paragraph.querySelectorAll<HTMLElement>('span'))
          .find(e => e.childElementCount === 0 && e.textContent === text)!.getBoundingClientRect();
        const titleRect = leaf(title), numberRect = leaf(number);
        const tabs = Array.from(paragraph.querySelectorAll<HTMLElement>('[data-docx-tab]'));
        return { titleLeft: titleRect.left - origin, numberRight: numberRect.right - origin,
          titleTop: titleRect.top, numberTop: numberRect.top,
          alignments: tabs.map(t => t.dataset.docxTab), leaders: tabs.map(t => t.dataset.docxTabLeader ?? null) };
      });
    }, { bytes: Array.from(multipleTabFixture()), entries: tabEntries, paginated, classes });
    for (const entry of results) {
      expect(entry.alignments).toEqual(['left', 'right']);
      expect(entry.leaders).toEqual([null, 'dot']);
      expect(entry.titleLeft).toBeCloseTo(120, 0);
      expect(entry.numberRight).toBeCloseTo(600, 0);
      expect(entry.numberTop).toBeCloseTo(entry.titleTop, 0);
    }
  });
}

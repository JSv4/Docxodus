import { test, expect } from '@playwright/test';
import { hiddenMarkFixture } from './docx-hidden-mark-fixture.js';

for (const paginated of [false, true]) for (const classes of [false, true]) {
  test(`hidden paragraph mark joins visible text and numbering (${paginated ? 'paginated' : 'continuous'}, classes=${classes})`, async ({ page }) => {
    await page.goto('/test-harness.html');
    await page.waitForFunction(() => (window as any).DocxodusReady === true);
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const result = await page.evaluate(({ bytes, paginated, classes }) => {
      const html = (window as any).Docxodus.DocumentConverter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytes), 'Cedar marks', 'docx-', classes, '', -1, 'comment-', paginated ? 1 : 0, 1, 'page-', false, 0, 'annot-', false, false, false, true, true);
      if (html.startsWith('{')) throw new Error(html);
      const host = document.createElement('div'); document.body.replaceChildren(host);
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
      else host.innerHTML = html;
      const root = paginated ? host.querySelector('.page-content')! : host;
      const text = (needle: string) => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        let node: Node | null;
        while ((node = walker.nextNode())) if (node.textContent?.includes(needle)) {
          const range = document.createRange(); range.selectNodeContents(node);
          const rect = range.getBoundingClientRect();
          const style = getComputedStyle(node.parentElement!);
          return { top: rect.top, width: rect.width, weight: style.fontWeight };
        }
        throw new Error(`Missing ${needle}`);
      };
      return { heading: text('Cedar.'), body: text('Return the cedar'), number: text('1.'),
        controlHeading: text('Maple.'), controlBody: text('Keep the maple'),
        headingText: root.querySelector('h2')!.textContent,
        separateBody: Array.from(root.querySelectorAll('p')).some(p => p.textContent?.includes('Return the cedar')) };
    }, { bytes: Array.from(hiddenMarkFixture()), paginated, classes });
    expect(result.headingText).toContain('Return the cedar folder before dusk.');
    expect(result.separateBody).toBe(false);
    expect(Math.abs(result.body.top - result.heading.top)).toBeLessThan(2);
    expect(Math.abs(result.number.top - result.heading.top)).toBeLessThan(2);
    expect(result.number.width).toBeGreaterThan(1);
    expect(Number(result.heading.weight)).toBeGreaterThanOrEqual(600);
    expect(Number(result.body.weight)).toBeLessThan(600);
    expect(result.controlBody.top - result.controlHeading.top).toBeGreaterThan(10);
  });
}

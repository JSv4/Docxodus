import { test, expect } from '@playwright/test';
import { balloonDocx, commented } from './docx-markup-balloons-fixture.js';

for (const prefix of ['Oriel', "Oriel's"]) {
  test(`declared font classes reach every story and balloon (${prefix})`, async ({ page }) => {
    const comments = Array.from({ length: 16 }, (_, id) => ({ id, author: 'Oriel Reviewer', initials: 'OR',
      text: 'Review the harbor ledger and confirm the totals. '.repeat(35) }));
    const bytes = balloonDocx({
      body: comments.map(({ id }) => `<w:p>${commented(id, `Section${id}`)}</w:p>`).join(''), comments,
      bodyFont: `${prefix} Text`, balloonFont: `${prefix} Caption`,
      header: `<w:p><w:r><w:rPr><w:rFonts w:ascii="${prefix} Running" w:hAnsi="${prefix} Running"/></w:rPr><w:t>OrielHeading</w:t></w:r></w:p>`,
      fontTable: `<w:font w:name="${prefix} Text"><w:family w:val="modern"/></w:font>` +
        `<w:font w:name="${prefix} Caption"><w:family w:val="roman"/></w:font>` +
        `<w:font w:name="${prefix} Running"><w:family w:val="swiss"/></w:font>`,
    });
    await page.goto('/worker-test-harness.html');
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const fonts = await page.evaluate(async input => {
      const moduleUrl = 'http://localhost:8083/core.js';
      const api = await import(moduleUrl);
      await api.initialize(`${location.origin}/`);
      const html = await api.convertDocxToHtml(new Uint8Array(input), {
        paginationMode: 1, commentRenderMode: 2, renderHeadersAndFooters: true,
        renderTrackedChanges: true, revisionPresentation: 1,
      });
      const parsed = new DOMParser().parseFromString(html, 'text/html');
      document.head.insertAdjacentHTML('beforeend', parsed.head.innerHTML);
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      (window as any).DocxodusPagination.paginateHtml(parsed.body.innerHTML, host, {
        showPageNumbers: false, commentColumnPlacement: 'balloons', balloonOverflow: 'appendix',
      });
      const family = (elements: Element[]) => elements.map(e => getComputedStyle(e).fontFamily);
      const spans = Array.from(host.querySelectorAll('.page-box span'));
      const selected = (selector: string) => family(Array.from(host.querySelectorAll(selector)));
      return {
        body: family(spans.filter(e => e.textContent === 'Section0')),
        header: family(spans.filter(e => e.textContent === 'OrielHeading')),
        comments: selected('.page-balloon'), labels: selected('.page-balloon b'),
        cuts: selected('.page-balloon-cut'), appendix: selected('.page-balloon-appendix-text'),
        bars: selected('.page-balloon-appendix-bar'),
      };
    }, Array.from(bytes));
    for (const [key, name, generic] of [
      ['body', `${prefix} Text`, 'monospace'], ['header', `${prefix} Running`, 'sans-serif'],
      ['comments', `${prefix} Text`, 'monospace'], ['labels', `${prefix} Caption`, 'serif'],
      ['cuts', `${prefix} Caption`, 'serif'], ['appendix', `${prefix} Text`, 'monospace'], ['bars', `${prefix} Caption`, 'serif'],
    ] as const) {
      expect(fonts[key].length, key).toBeGreaterThan(0);
      for (const stack of fonts[key]) {
        expect(stack, key).toContain(name);
        expect(stack.split(',').at(-1)!.trim(), key).toBe(generic);
      }
    }
  });
}

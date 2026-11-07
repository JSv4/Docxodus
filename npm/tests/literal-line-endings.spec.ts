import { expect, test } from '@playwright/test';
import { literalLineEndingCases, literalLineEndingFixture } from './docx-literal-tab-fixture.js';

for (const [kind, name, ending, spaces] of literalLineEndingCases) {
  test(`${kind} text ${name} has the same visible line and tab geometry as spaces`, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const facts = await page.evaluate(async (inputs) => {
      await (window as any).createDocxodusWorker();
      const worker = (window as any).DocxodusWorker, result = [];
      for (const bytes of inputs) {
        const html = await worker.convertDocxToHtml(new Uint8Array(bytes), {
          paginationMode: 1, renderHeadersAndFooters: true, renderTrackedChanges: true, showDeletedContent: true,
        });
        const parsed = new DOMParser().parseFromString(html, 'text/html');
        document.head.innerHTML = parsed.head.innerHTML;
        document.body.innerHTML = '<div id="host"></div>';
        const host = document.querySelector<HTMLElement>('#host')!;
        (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
        await document.fonts.ready;
        const box = host.querySelector('.page-box')!, markers = [];
        for (const marker of ['NorthMark', 'HarborMark', 'TailMark']) {
          const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT), found = [];
          for (let node = walker.nextNode(); node; node = walker.nextNode()) {
            const start = (node.textContent ?? '').indexOf(marker);
            if (start < 0) continue;
            const range = document.createRange(); range.setStart(node, start); range.setEnd(node, start + marker.length);
            const rect = range.getBoundingClientRect();
            found.push({ x: rect.x, y: rect.y, width: rect.width, height: rect.height });
          }
          markers.push(found);
        }
        result.push({ pages: host.querySelectorAll('.page-box').length, markers,
          text: box.textContent, breaks: box.querySelectorAll('br').length,
          tabs: box.querySelectorAll('[data-docx-tab]').length });
      }
      worker.terminate(); return result;
    }, [Array.from(literalLineEndingFixture(kind, ending)), Array.from(literalLineEndingFixture(kind, spaces))]);
    expect(facts[0]).toEqual(facts[1]);
    expect(facts[0].pages).toBe(1);
    expect(facts[0].tabs).toBe(1);
    for (const marker of facts[0].markers) expect(marker).toHaveLength(1);
    expect(facts[0].markers.map(m => m[0].y)).toEqual(Array(3).fill(facts[0].markers[0][0].y));
  });
}

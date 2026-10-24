import { expect, test } from '@playwright/test';
import { literalTabFixture, type LiteralTabCase } from './docx-literal-tab-fixture.js';

for (const kind of ['body', 'header', 'footer', 'multiple', 'deleted', 'inserted'] as LiteralTabCase[]) {
  test('literal tabs use paragraph stops in ' + kind, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    await page.addScriptTag({ url: '/pagination.bundle.js' });
    const facts = await page.evaluate(async ({ actual, control }) => {
      await (window as any).createDocxodusWorker();
      const worker = (window as any).DocxodusWorker;
      const result = [];
      for (const bytes of [actual, control]) {
        const html = await worker.convertDocxToHtml(new Uint8Array(bytes), {
          paginationMode: 1, renderHeadersAndFooters: true, renderTrackedChanges: true, showDeletedContent: true,
        });
        const parsed = new DOMParser().parseFromString(html, 'text/html');
        document.head.innerHTML = parsed.head.innerHTML;
        document.body.innerHTML = '<div id="host"></div>';
        const host = document.querySelector<HTMLElement>('#host')!;
        (window as any).DocxodusPagination.paginateHtml(html, host, { showPageNumbers: false });
        await document.fonts.ready;
        const box = host.querySelector('.page-box')!;
        const column = box.querySelector('.page-content')!.getBoundingClientRect();
        const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
        const matches: Text[] = [];
        for (let node = walker.nextNode(); node; node = walker.nextNode())
          if (node.textContent?.includes('HarborMark')) matches.push(node as Text);
        if (matches.length !== 1) throw new Error('Expected one visible HarborMark, got ' + matches.length);
        const node = matches[0], start = node.data.indexOf('HarborMark');
        const range = document.createRange(); range.setStart(node, start); range.setEnd(node, start + 'HarborMark'.length);
        const rect = range.getBoundingClientRect();
        result.push({ right: rect.right - column.left, width: rect.width,
          tabs: Array.from(box.querySelectorAll<HTMLElement>('[data-docx-tab]')).map(t => ({ alignment: t.dataset.docxTab, leader: t.dataset.docxTabLeader })),
          decoration: Array.from(box.querySelectorAll<HTMLElement>('del > span, ins > span, del > span > span, ins > span > span'))
            .filter(e => !e.hasAttribute('data-docx-exact-run')).map(e => getComputedStyle(e).textDecorationLine),
          weight: getComputedStyle(node.parentElement!).fontWeight, text: box.textContent });
      }
      worker.terminate();
      return result;
    }, { actual: Array.from(literalTabFixture(kind)), control: Array.from(literalTabFixture(kind, false)) });
    expect(facts[0].tabs.map(t => t.alignment)).toEqual(kind === 'multiple' ? ['left', 'right'] : ['right']);
    expect(facts[0].right).toBeCloseTo(624, 0);
    if (kind === 'deleted' || kind === 'inserted') {
      const line = kind === 'deleted' ? 'line-through' : 'underline';
      expect(facts[0].decoration).toEqual([line, line, line]);
    }
    expect(facts[0].weight).toBe('700');
    expect(facts[0]).toEqual(facts[1]);
  });
}

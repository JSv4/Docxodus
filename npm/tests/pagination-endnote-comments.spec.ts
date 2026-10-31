import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const layouts = [
  ...['comment-', 'review:', 'review', 'review--'].map(prefix => ({ prefix, band: 'plain' })),
  ...['continuous', 'header'].map(band => ({ prefix: 'comment-', band })),
];

for (const { prefix, band } of layouts) test(`endnote ${prefix} threads retain presentation and ownership in ${band} bands`, async ({ page }) => {
  const base = prefix.replace(/-+$/, '');
  const escaped = await page.evaluate(value => CSS.escape(value), prefix);
  const geometry = `data-page-width="400" data-page-height="300" data-content-width="256" data-content-height="156"
    data-margin-top="72" data-margin-bottom="72" data-margin-left="72" data-margin-right="72"`;
  const preceding = band === 'continuous' ? `<section data-section-index="0" ${geometry}><p>Preceding section</p></section>` : '';
  const header = band === 'header' ? `<div id="pagination-hf-registry"><div data-section="0" data-hf-type="header-default"><p style="height:90pt">Tall running header</p></div></div>` : '';
  const words = Array.from({ length: 240 }, (_, i) => `Cedar${i}`);
  await page.setContent(`<style>
    p{margin:0;font:12pt/15pt 'Liberation Serif'} aside{margin-top:12pt;border-top:2px solid #ccc;padding-top:8pt}
    ol{list-style:none;padding:0}li{padding:9pt;background:#f5f5f5;border-left:3px solid blue;margin-bottom:12pt}
    .${escaped}replies{padding-left:18pt}.resolved{opacity:.65}
    .${escaped}header{display:flex;gap:6pt;font:10pt serif;margin-bottom:6pt}
    </style><div id="s">${header}${preceding}<section data-section-index="${band === 'continuous' ? 1 : 0}"
      data-section-type="continuous" data-header-height="12" ${geometry}>
      <p id="comment-ref-0"><a href="#comment-0">[1]</a><a href="#comment-1">[2]</a> Body</p>
      <aside class="${base}s-section"><h2>Comments</h2><ol class="${base}s-list">
        <li class="${base}" id="comment-0" data-source-anchor-id="cmt:cmt:root" data-author="Cedar Reviewer">
          <div class="${prefix}header">Cedar Reviewer<a href="#comment-ref-0">Return</a></div>
          <div class="${prefix}body"><p data-source-anchor-id="p:cmt:first">First message</p></div>
          <ol class="${prefix}replies"><li class="${base} resolved" id="comment-1" data-source-anchor-id="cmt:cmt:reply">
            <div class="${prefix}header">Reply Reviewer<a href="#comment-ref-0">Return</a></div>
            <div class="${prefix}body"><p data-source-anchor-id="p:cmt:long">${words.join(' ')}</p></div>
          </li></ol>
        </li></ol></aside>
    </section></div><div id="o"></div>`);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const facts = await page.evaluate(() => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('s'), document.getElementById('o'), {
      fragmentParagraphs: true, showPageNumbers: false,
      layoutToken: { documentVersion: 1, rendererFingerprint: 'endnote-comments' },
    });
    const result = engine.paginate();
    const output = document.getElementById('o')!;
    const comments = Array.from(output.querySelectorAll<HTMLElement>('li'));
    return { pages: result.totalPages, map: result.pageMap,
      bodyHeights: Array.from(output.querySelectorAll<HTMLElement>('.page-content')).map(c => c.clientHeight * .75),
      sections: result.pageMap.pages.map((p: any) => p.sectionIndex),
      text: output.textContent, titles: output.querySelectorAll('h2').length,
      links: Array.from(output.querySelectorAll<HTMLAnchorElement>('a[href^="#"]')).map(a =>
        output.querySelectorAll(`[id="${a.hash.slice(1)}"]`).length),
      overflowing: Array.from(output.querySelectorAll<HTMLElement>('.page-content')).some(c => c.scrollHeight > c.clientHeight + 1),
      owners: comments.map(c => ({ author: c.dataset.author, background: getComputedStyle(c).backgroundColor })),
    };
  });
  if (band === 'header') expect(Math.max(...facts.bodyHeights)).toBeLessThan(156);
  if (band === 'continuous') expect(new Set(facts.sections)).toEqual(new Set([0, 1]));
  expect(facts.overflowing).toBe(false);
  expect(facts.pages).toBeGreaterThan(2);
  expect(facts.text!.match(/Cedar\d+/g)).toEqual(words);
  expect(facts.titles).toBe(1);
  expect(facts.links.every(count => count === 1)).toBe(true);
  expect(facts.owners.every(owner => owner.background === 'rgb(245, 245, 245)')).toBe(true);
  expect(facts.owners.some(owner => owner.author === 'Cedar Reviewer')).toBe(true);
  const paragraphPages = facts.map.fragments.filter((f: any) => f.anchorId === 'p:cmt:long').map((f: any) => f.pageNumber);
  for (const id of ['cmt:cmt:root', 'cmt:cmt:reply']) {
    const pages = facts.map.fragments.filter((f: any) => f.anchorId === id).map((f: any) => f.pageNumber);
    for (const page of paragraphPages) expect(pages).toContain(page);
  }
});

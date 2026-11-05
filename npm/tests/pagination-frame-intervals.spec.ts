import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const cases = [
  ...['center', 'interior', 'right', 'page'].map(position => ({ name: position, position })),
  ...['fragmented', 'long word', 'justified', 'distant table', 'mixed RTL', 'RTL span'].map(name => ({ name, position: 'center' })),
];

for (const { name, position } of cases) {
  test(`${name} inline frame uses verified intervals or retains its diagnostic`, async ({ page }) => {
    const words = Array.from({ length: name === 'fragmented' ? 240 : 24 }, (_, i) => `w${String(i).padStart(2, '0')}`);
    if (name === 'long word') words[0] = 'w00000000000000';
    const alignment = position === 'center' || position === 'right' ? position : '';
    const offset = position === 'interior' ? 60 : position === 'page' ? 6 : 0;
    await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style>
      <div id="staging"><section data-section-index="0" data-page-width="228" data-page-height="240"
        data-content-width="180" data-content-height="192" data-margin-top="24" data-margin-bottom="24"
        data-margin-left="24" data-margin-right="24">
        <div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0">
          <div data-docx-frame-exclusion="true" style="float:left;display:grid">
            <div data-docx-text-frame="true" data-docx-frame-pending="true" data-content-type="Other" data-docx-frame-h-anchor="${position === 'page' ? 'page' : 'text'}"
              data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline" data-docx-frame-wrap="around"
              ${alignment ? `data-docx-frame-xalign="${alignment}"` : `data-docx-frame-x="${offset * 20}"`}
              style="position:relative;display:flow-root;grid-area:1/1;width:60pt;height:24pt;margin-left:${offset}pt">
              <p>BoxOne</p><p>BoxTwo</p>
            </div>
          </div>
        </div>
        <p data-source-anchor-id="interval-source" style="text-align:${name === 'justified' ? 'justify' : 'left'}"><span ${name === 'RTL span' ? 'dir="rtl"' : ''}>${name === 'mixed RTL' ? 'אבג דהו ' : ''}${words.join(' ')}</span></p>
        ${name === 'distant table' ? '<div style="height:80pt"></div><table><tr><td>Later table</td></tr></table>' : ''}
      </section></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const measured = await page.evaluate(() => {
      const engine = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true,
          layoutToken: { documentVersion: 1, rendererFingerprint: 'frame-intervals' } });
      const pagination = engine.paginate();
      const box = document.querySelector<HTMLElement>('#output .page-box')!;
      const origin = box.getBoundingClientRect(), scale = origin.width / 228;
      const frame = box.querySelector('[data-docx-text-frame]')!.getBoundingClientRect();
      const found: Array<{ text: string; x: number; y: number; right: number; page: number }> = [];
      for (const p of Array.from(document.querySelectorAll('#output [data-source-anchor-id="interval-source"]'))) {
        const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          for (const match of node.textContent!.matchAll(/w\d+/g)) {
            const range = document.createRange(); range.setStart(node, match.index!); range.setEnd(node, match.index! + match[0].length);
            const r = range.getBoundingClientRect();
            const pageBox = p.closest<HTMLElement>('.page-box')!, pageOrigin = pageBox.getBoundingClientRect();
            found.push({ text: match[0], x: (r.left - pageOrigin.left) / scale,
              y: (r.top - pageOrigin.top) / scale, right: (r.right - pageOrigin.left) / scale,
              page: Number(pageBox.dataset.pageNumber) });
          }
        }
      }
      return { pages: pagination.totalPages, map: pagination.pageMap,
        pending: box.querySelectorAll('[data-docx-frame-pending]').length, frames: box.querySelectorAll('[data-docx-text-frame]').length,
        left: (frame.left - origin.left) / scale, right: (frame.right - origin.left) / scale,
        top: (frame.top - origin.top) / scale, bottom: (frame.bottom - origin.top) / scale, words: found };
    });
    expect(measured.frames).toBe(1);
    expect(measured.left).toBeCloseTo(position === 'page' ? 6 : position === 'right' ? 144 : 84, 1);
    expect(measured.words.map(word => word.text)).toEqual(words);
    const unsupported = name === 'mixed RTL' || name === 'RTL span';
    expect(measured.pending).toBe(unsupported ? 1 : 0);
    if (unsupported) return;
    if (name === 'fragmented') {
      expect(measured.pages).toBeGreaterThan(1);
      expect(new Set(measured.map.fragments.filter((f: any) => f.anchorId === 'interval-source')
        .map((f: any) => f.pageNumber)).size).toBe(measured.pages);
    }
    if (name === 'long word') {
      expect(measured.words[0].y).toBeGreaterThanOrEqual(measured.bottom - 1);
      return;
    }
    const first = measured.words.filter(word => word.page === 1 && Math.abs(word.y - measured.words[0].y) < 1);
    expect(first[0].y).toBeLessThan(measured.top + 12);
    expect(first[0].x).toBeCloseTo(position === 'page' ? 66 : 24, 1);
    if (position === 'center' || position === 'interior') {
      expect(first.some(word => word.right <= measured.left + .1)).toBe(true);
      expect(first.some(word => word.x >= measured.right - .1)).toBe(true);
    }
    for (const word of measured.words.filter(word => word.page === 1 && word.y < measured.bottom - 1)) {
      expect(word.right <= measured.left + .1 || word.x >= measured.right - .1).toBe(true);
    }
  });
}

for (const alignment of ['inside', 'outside']) test(`header ${alignment} uses physical parity after displayed numbering restarts`, async ({ page }) => {
  await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
    <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="header-default">
      <div data-docx-frame-container="true" style="position:relative;height:0"><div data-docx-frame-exclusion="true" style="float:left;display:grid">
        <div data-docx-text-frame="true" data-docx-frame-pending="true" data-content-type="Other"
          data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
          data-docx-frame-xalign="${alignment}" style="position:relative;display:flow-root;grid-area:1/1;width:60pt;height:96pt"><p>Frame</p></div>
      </div></div><p>Header words</p></div></div>
    <section data-section-index="0" data-page-num-start="2" data-page-width="228" data-page-height="240"
      data-content-width="180" data-content-height="192" data-margin-top="24" data-margin-bottom="24"
      data-margin-left="24" data-margin-right="24" data-header-height="5">
      <p data-body="true">body0 body1 body2</p><p data-body="true" data-page-break-before="true">body3 body4 body5</p>
    </section></div><div id="output"></div>`);
  await page.addScriptTag({ path: bundle });
  const pages = await page.evaluate(() => {
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
      document.getElementById('output'), { fragmentParagraphs: true, showPageNumbers: false }).paginate();
    return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
      const origin = box.getBoundingClientRect(), scale = origin.width / 228;
      const frame = box.querySelector('[data-docx-text-frame]')!.getBoundingClientRect();
      const body = box.querySelector('[data-body]')!, walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
      const text = walker.nextNode()!, range = document.createRange(); range.setStart(text, 0); range.setEnd(text, 5);
      const word = range.getBoundingClientRect();
      return { displayed: Number(box.dataset.displayedPageNumber), left: (frame.left - origin.left) / scale,
        first: (word.left - origin.left) / scale, overlaps: word.top < frame.bottom,
        pending: box.querySelectorAll('[data-docx-frame-pending]').length };
    });
  });
  expect(pages.map(p => p.displayed)).toEqual([2, 3]);
  for (const [i, facts] of pages.entries()) {
    const left = (i === 0) === (alignment === 'inside');
    expect(facts.left).toBeCloseTo(left ? 24 : 144, 1);
    expect(facts.first).toBeCloseTo(left ? 84 : 24, 1);
    expect(facts.overlaps).toBe(true);
    expect(facts.pending).toBe(0);
  }
});

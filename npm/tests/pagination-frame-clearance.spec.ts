import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
for (const mode of ['around', 'notBeside', 'long', 'page-end', 'siblings', 'separated', 'fragmented', 'footnote']) test(`${mode} vertical clearance reflows the prefix only when its frame is admitted`, async ({ page }) => {
  const words = ['long', 'fragmented'].includes(mode) ? Array.from({ length: mode === 'long' ? 60 : 150 }, (_, i) => `B${String(i).padStart(2, '0')}`) : ['Before'];
  const fillers = ['page-end', 'footnote'].includes(mode) ? Array.from({ length: mode === 'footnote' ? 10 : 13 },
    (_, i) => `<p>Filler${i}${mode === 'footnote' && i === 0 ? '<sup data-footnote-id="f1">1</sup>' : ''}</p>`).join('') : '';
  const extraFrame = `<div data-docx-text-frame="true" data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text"
    data-docx-frame-yalign="inline" data-docx-frame-wrap="notBeside" data-docx-frame-clearance="4"
    style="position:relative;display:flow-root;grid-area:1/1;width:60pt;height:24pt;margin-left:108pt"><p>OtherFrame</p></div>`;
  await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
    ${mode === 'footnote' ? '<div id="pagination-footnote-registry" style="display:none"><div class="footnote-item" data-footnote-id="f1"><span class="footnote-number">1</span><div class="footnote-content"><p>Note</p></div></div></div>' : ''}
    <section data-section-index="0" data-page-width="228" data-page-height="240" data-content-width="180"
      data-content-height="192" data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
      ${fillers}<p data-source-anchor-id="before">${words.join(' ')}</p>
      <div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0">
        <div data-docx-frame-exclusion="true" style="float:left;display:grid">
          <div data-docx-text-frame="true" data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text"
            data-docx-frame-yalign="inline" data-docx-frame-wrap="${mode === 'around' ? 'around' : 'notBeside'}"
            data-docx-frame-pending="true" data-docx-shape="partial" data-content-type="Other" data-element-name="framePr" data-docx-frame-vspace="80" data-docx-frame-clearance="4" style="position:relative;display:flow-root;grid-area:1/1;width:60pt;height:24pt;margin-left:12pt">
            <p data-source-anchor-id="frame-first">FrameOne</p><p>FrameTwo</p>
          </div>${mode === 'siblings' ? extraFrame : ''}
        </div>
      </div>${mode === 'separated' ? `<p>Between</p><div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0"><div data-docx-frame-exclusion="true" style="float:left;display:grid">${extraFrame}</div></div>` : ''}<p data-source-anchor-id="after">After</p>
    </section></div><div id="output"></div>`);
  await page.addScriptTag({ path: bundle });
  const facts = await page.evaluate(() => {
    const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
      document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true,
        layoutToken: { documentVersion: 1, rendererFingerprint: 'frame-clearance' } }).paginate();
    const found: Array<{ text: string; x: number; y: number; page: number }> = [];
    for (const paragraph of Array.from(document.querySelectorAll('#output p'))) {
      const box = paragraph.closest<HTMLElement>('.page-box')!, origin = box.getBoundingClientRect(), scale = origin.width / 228;
      const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
      for (let text = walker.nextNode(); text; text = walker.nextNode()) for (const match of text.textContent!.matchAll(/\S+/g)) {
        const range = document.createRange(); range.setStart(text, match.index!); range.setEnd(text, match.index! + match[0].length);
        const rect = range.getBoundingClientRect();
        found.push({ text: match[0], x: (rect.left - origin.left) / scale, y: (rect.top - origin.top) / scale,
          page: Number(box.dataset.pageNumber) });
      }
    }
    return { pages: result.totalPages, words: found, map: result.pageMap, pending: document.querySelectorAll('#output [data-docx-frame-pending]').length };
  });
  expect(facts.pending).toBe(0);
  const word = (text: string) => {
    const found = facts.words.filter(w => w.text === text); expect(found, text).toHaveLength(1); return found[0];
  };
  words.forEach(word);
  const first = word('FrameOne'), last = word('FrameTwo'), after = word('After'), before = word(words.at(-1)!);
  expect(last.y - first.y).toBeCloseTo(12, 1);
  expect(first.x).toBeCloseTo(36, 1);
  expect(facts.pages).toBe(['page-end', 'footnote', 'fragmented'].includes(mode) ? 2 : 1);
  if (mode === 'fragmented') {
    expect(new Set(facts.map.fragments.filter((f: any) => f.anchorId === 'before').map((f: any) => f.pageNumber)).size).toBe(2);
    expect(after.page).toBe(first.page); expect(after.y - first.y).toBeCloseTo(40, 1);
    return;
  }
  if (['siblings', 'separated'].includes(mode)) {
    const other = word('OtherFrame');
    expect(other.x).toBeCloseTo(132, 1);
    expect(other.y - first.y).toBeCloseTo(mode === 'siblings' ? 0 : 52, 1);
    expect(after.y - other.y).toBeCloseTo(40, 1);
    if (mode === 'separated') expect(word('Between').y - other.y).toBeCloseTo(28, 1);
  }
  if (['page-end', 'footnote'].includes(mode)) {
    expect(before.page).toBe(1); expect(first.page).toBe(2); expect(after.page).toBe(2);
    if (mode === 'footnote') expect(word('Note').page).toBe(1);
    expect(before.y - word('Filler0').y).toBeCloseTo(mode === 'footnote' ? 120 : 156, 1);
    expect(after.y - first.y).toBeCloseTo(28, 1);
  } else {
    expect(before.y - first.y).toBeCloseTo(mode === 'around' ? -12 : 28, 1);
    expect(before.x).toBeCloseTo(mode === 'around' ? 96 : mode === 'long' ? 96 : 24, 1);
    expect(after.y - first.y).toBeCloseTo(mode === 'around' ? 0 : mode === 'separated' ? 92 : 40, 1);
    if (mode === 'long') expect(first.y - word('B00').y).toBeCloseTo(108, 1);
  }
  for (const [anchorId, pageNumber] of [['before', 1], ['frame-first', first.page], ['after', after.page]]) {
    expect(facts.map.fragments.filter((fragment: any) => fragment.anchorId === anchorId))
      .toEqual([expect.objectContaining({ pageNumber })]);
  }
});

for (const [beforeCount, afterCount, bottom, top, second] of [
  [6, 5, 0, 3, false], [7, 3, 0, 15, true], [5, 5, 0, 9, true], [2, 4, 0, 15, false],
] as const) test(`clearance admission budgets the ${beforeCount}-paragraph prefix and ${second ? 'two' : 'one'} frames`, async ({ page }) => {
  const paragraph = (text: string, margin = '') => `<p style="${margin}">${text}</p>`;
  const frame = (label: string, clearance: number) => `<div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0">
    <div data-docx-frame-exclusion="true" style="float:left;display:grid"><div data-docx-text-frame="true"
    data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
    data-docx-frame-wrap="notBeside" data-docx-frame-clearance="${clearance}"
    style="position:relative;display:flow-root;width:60pt;height:24pt;margin-left:12pt">${paragraph(label)}</div></div></div>`;
  const before = Array.from({ length: beforeCount }, (_, i) => `Pre${i}`), after = Array.from({ length: afterCount }, (_, i) => `Post${i}`);
  await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
    <section data-section-index="0" data-page-width="228" data-page-height="240" data-content-width="180" data-content-height="192"
    data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
    ${before.map(text => paragraph(text, `margin-bottom:${bottom}pt`)).join('')}${frame('FrameA', 4)}
    ${paragraph('Middle', `margin-top:${top}pt`)}${second ? frame('FrameB', 8) : ''}
    ${after.map(text => paragraph(text, `margin-top:${top}pt`)).join('')}</section></div><div id="output"></div>`);
  await page.addScriptTag({ path: bundle });
  const pages = await page.evaluate(() => {
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('output'),
      { showPageNumbers: false, fragmentParagraphs: true }).paginate();
    return Array.from(document.querySelectorAll('#output .page-box')).map(box => {
      const body = box.querySelector('.page-content')!.getBoundingClientRect();
      const paragraphs = Array.from(box.querySelectorAll('p'));
      return { text: paragraphs.map(p => p.textContent), overflow: Math.max(...paragraphs.map(p => p.getBoundingClientRect().bottom - body.bottom)) };
    });
  });
  for (const text of [...before, 'FrameA', 'Middle', ...(second ? ['FrameB'] : []), ...after])
    expect(pages.flatMap(page => page.text).filter(value => value === text)).toHaveLength(1);
  for (const page of pages) expect(page.overflow).toBeLessThan(0.5);
});

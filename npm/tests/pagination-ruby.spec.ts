import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const story of ['body', 'table', 'footnote'] as const) for (const ruby of [false, true]) {
  test(`${story} continuation keeps ruby base and annotation together=${ruby}`, async ({ page }) => {
    const words = Array.from({ length: 100 }, (_, i) => `Pine${String(i).padStart(3, '0')}`);
    const pair = ruby
      ? '<span data-docx-ruby="true" data-docx-ruby-raise="11" style="display:inline-block"><ruby style="ruby-align:center"><span data-docx-ruby-base="true">CEDAR</span><rp>(</rp><rt style="font-size:6pt">reading</rt><rp>)</rp></ruby></span>'
      : 'CEDAR reading';
    const paragraph = `<p data-source-anchor-id="p:${story === 'footnote' ? 'fn' : 'body'}:long">${words.slice(0, 35).join(' ')} ${pair} ${words.slice(35).join(' ')}</p>`;
    const table = `<div><table style="width:188pt;table-layout:fixed;border-collapse:collapse"><tbody>${Array.from({ length: 10 }, (_, row) =>
      `<tr><td style="padding:0"><p data-source-anchor-id="p:body:long">${words.slice(row * 10, row * 10 + 10).join(' ')}${row === 3 ? ` ${pair}` : ''}</p></td></tr>`).join('')}</tbody></table></div>`;
    const note = `<div id="pagination-footnote-registry" style="display:none"><div class="footnote-item" data-footnote-id="f1"><span class="footnote-number">1</span><span class="footnote-content">${paragraph}</span></div></div>`;
    await page.setContent(`<style>p{font:11pt/normal 'Liberation Serif';margin:0}td{vertical-align:top}rt{font-family:'Liberation Serif'}</style>
      <div id="staging">${story === 'footnote' ? note : ''}
        <div data-section-index="0" data-page-width="228" data-page-height="200" data-content-width="188" data-content-height="160"
          data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
          ${story === 'body' ? paragraph : story === 'table' ? table : '<p data-source-anchor-id="p:body:host">Cite<sup data-footnote-id="f1">1</sup></p>'}
          <p data-source-anchor-id="p:body:tail">PineTail</p>
        </div></div><div id="container"></div>`);
    await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
    const result = await page.evaluate(() => {
      const container = document.getElementById('container')!;
      try {
        const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), container,
          { showPageNumbers: false, fragmentParagraphs: true, layoutToken: { documentVersion: 1, rendererFingerprint: 'ruby-test' } }).paginate();
        return { error: '', pages: result.totalPages, text: container.textContent!,
          overflow: Array.from(container.querySelectorAll<HTMLElement>('[data-source-anchor-id$=":long"]')).map(p => {
            const paper = p.closest('.page-box')!.getBoundingClientRect(), ink = p.getBoundingClientRect();
            return Math.max(ink.bottom - (paper.bottom - 20 * 96 / 72), paper.top + 20 * 96 / 72 - ink.top);
          }),
          pairs: Array.from(container.querySelectorAll('ruby')).map(r => ({ base: r.firstElementChild?.textContent, annotation: r.querySelector('rt')?.textContent })) };
      } catch (e) { return { error: String(e), pages: 0, text: '', overflow: [], pairs: [] }; }
    });
    expect(result.error).toBe('');
    expect(result.pages).toBeGreaterThan(1);
    expect(result.text.match(/Pine\d{3}/g)).toEqual(words);
    expect(result.text.match(/CEDAR/g)).toHaveLength(1);
    expect(result.text.match(/reading/g)).toHaveLength(1);
    expect(result.overflow.every(value => value < 1)).toBe(true);
    expect(result.pairs).toEqual(ruby ? [{ base: 'CEDAR', annotation: 'reading' }] : []);
  });
}

for (const raise of [0, 11, 22]) test(`ruby reserves annotation ink at ${raise}pt without accumulating offsets`, async ({ page }) => {
  await page.setContent(`<style>p{font:11pt/normal 'Liberation Serif';margin:0}</style><div id="staging">
    <section data-section-index="0" data-page-width="300" data-page-height="300" data-content-width="260" data-content-height="260"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20">
      <p data-source-anchor-id="p:body:ruby">Before <span data-docx-ruby="true" data-docx-ruby-raise="${raise}" style="display:inline-block">
        <ruby><span data-docx-ruby-base="true">CEDAR</span><rp>(</rp><rt style="font-size:6pt">reading</rt><rp>)</rp></ruby></span> After</p>
      <p>Tail</p></section></div><div id="container"></div>`);
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const results = await page.evaluate(() => {
    return [1, 2].map(() => {
      const engine = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { showPageNumbers: false, fragmentParagraphs: true });
      engine.paginate();
      const paragraph = document.querySelector<HTMLElement>('#container [data-source-anchor-id="p:body:ruby"]')!;
      const base = paragraph.querySelector<HTMLElement>('[data-docx-ruby-base]')!, rt = paragraph.querySelector<HTMLElement>('rt')!;
      const baseline = (element: HTMLElement) => {
        const marker = document.createElement('span');
        marker.style.cssText = 'display:inline-block;width:0;height:0;vertical-align:baseline;margin:0;padding:0;border:0';
        element.append(marker); const y = marker.getBoundingClientRect().top; marker.remove(); return y;
      };
      const gap = (baseline(base) - baseline(rt)) * 0.75;
      const line = paragraph.getBoundingClientRect();
      return { gap, height: line.height, inkFits: [base, rt].every(element => {
        const rect = element.getBoundingClientRect(); return rect.top >= line.top - 1 && rect.bottom <= line.bottom + 1;
      }), falseLine: engine.startsVisualLine(paragraph, { node: paragraph, offset: 2 }) };
    });
  });
  for (const result of results) {
    expect(result.gap).toBeCloseTo(raise, 1);
    expect(result.inkFits).toBe(true);
    expect(result.falseLine).toBe(false);
  }
  expect(results[1]).toEqual(results[0]);
});

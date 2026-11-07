import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const kind of ['rows', 'oversized', 'cant-split-rows', 'single-row', 'kept', 'single-row-kept', 'headers', 'nested', 'partial-keeps', 'nested-kept', 'header-kept', 'oversized-kept', 'short-long-kept'] as const) {
  test(`a ${kind} table uses the remaining page without losing content`, async ({ page }) => {
    const entries = Array.from({ length: kind === 'oversized' || kind === 'oversized-kept' || kind === 'short-long-kept' ? 18 : 8 }, (_, i) => `Cedar${i + 1}`);
    const kept = kind === 'kept' || kind === 'single-row-kept' || kind === 'nested-kept' || kind === 'oversized-kept' || kind === 'short-long-kept';
    const header = kind === 'headers' || kind === 'header-kept';
    const row = (text: string, index: number) => `<tr ${kind === 'cant-split-rows' ? 'data-docx-cant-split="true"' : ''}><td><p ${kept || (kind === 'partial-keeps' && (index === 3 || index === 4)) ? 'data-keep-with-next="true"' : ''}>${text}</p></td></tr>`;
    const rows = kind === 'short-long-kept'
      ? row(entries[0], 0) + `<tr><td><p>${entries.slice(1).join('<br>')}</p></td></tr>`
      : kind === 'single-row' || kind === 'single-row-kept'
      ? `<tr ${kept ? 'data-docx-cant-split="true"' : ''}><td><p>${entries.join('<br>')}</p></td></tr>`
      : entries.map(row).join('');
    const table = (content: string) => `<div><table><colgroup><col style="width:100%"></colgroup><tbody>${content}</tbody></table></div>`;
    const content = kind === 'nested' || kind === 'nested-kept' ? table(`<tr><td>${table(rows)}</td></tr>`)
      : table((header ? `<tr data-docx-header-row="true"><td><p ${kind === 'header-kept' ? 'data-keep-with-next="true"' : ''}>Header</p></td></tr>` : '') + rows);
    await page.setContent(`<style>
      #pagination-staging, #pagination-container { font: 12pt/14pt Arial; }
      p { margin:0; } table { width:300pt; border-collapse:collapse; table-layout:fixed; }
      td { vertical-align:top; padding:0; border:0; }
    </style><div id="pagination-staging"><div data-section-index="0"
      data-page-width="340" data-page-height="180" data-content-width="300" data-content-height="140"
      data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
      ${Array.from({ length:5 }, (_, i) => `<p>Lead${i + 1}</p>`).join('')}
      ${content}
      <p>Complete.</p></div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('pagination-staging'), container,
        { fragmentParagraphs:true, showPageNumbers:false }).paginate();
      return Array.from(container.querySelectorAll<HTMLElement>('.page-content')).map(body => ({
        text:body.innerText, overflow:Math.max(0, ...Array.from(body.querySelectorAll('p'), p =>
          p.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom)),
      }));
    });
    expect(result[0].text.match(/Cedar\d+/g) ?? []).toEqual(kept ? [] : entries.slice(0, kind === 'partial-keeps' ? 3 : header ? 4 : 5));
    expect(result.flatMap(p => p.text.match(/Cedar\d+/g) ?? [])).toEqual(entries);
    expect(result.every(p => p.overflow <= 1)).toBe(true);
    expect(result.map(p => p.text).join(' ').match(/Complete\./g)).toHaveLength(1);
    if (header) expect(result.filter(p => /Cedar/.test(p.text)).every(p => p.text.includes('Header'))).toBe(true);
  });
}

for (const remeasure of [false, true]) for (const guard of ['none', 'cant-split', 'note']) {
  const cantSplit = guard === 'cant-split', note = guard === 'note';
  test(`column-table remainder respects guard=${guard} through remeasurement=${remeasure}`, async ({ page }) => {
    const geometry = 'data-page-width="340" data-page-height="180" data-content-width="300" data-content-height="140" data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20"';
    await page.setContent(`<style>#pagination-staging,#pagination-container {font:12pt/14pt Arial}
      p {margin:0} table {width:130pt;border-collapse:collapse;table-layout:fixed} td {padding:0;border:0}</style>
      <div id="pagination-staging"><section data-section-index="0" ${geometry}>
      ${Array.from({ length:5 }, (_, i) => `<p ${remeasure && i === 0 ? 'data-docx-edge-picture="true"' : ''}>Lead${i + 1}</p>`).join('')}</section>
      <section data-section-index="1" ${geometry} data-section-type="continuous" data-cols="2" data-col-gap="20" style="column-count:2;column-gap:20pt">
      <div><table><colgroup><col style="width:100%"></colgroup><tbody>
      ${Array.from({ length:12 }, (_, i) => `<tr ${cantSplit ? 'data-docx-cant-split="true"' : ''}><td><p>Cedar${i + 1}</p></td></tr>`).join('')}
      </tbody></table></div><p>Complete.${note ? '<sup data-footnote-id="note">1</sup>' : ''}</p></section>
      ${note ? '<div id="pagination-footnote-registry"><div class="footnote-item" data-footnote-id="note"><span class="footnote-number">1</span><div class="footnote-content"><p>Note text.</p></div></div></div>' : ''}
      </div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('pagination-container')!;
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('pagination-staging'), container,
        { fragmentParagraphs:true, showPageNumbers:false }).paginate();
      return Array.from(container.querySelectorAll<HTMLElement>('.page-content')).map(body => ({
        text:body.innerText, warnings:body.querySelectorAll('[data-docx-column-table-deferred]').length,
      }));
    });
    expect(result).toHaveLength(2);
    const labels = result.flatMap(page => page.text.match(/Cedar\d+/g) ?? []);
    expect(labels).toEqual(Array.from({ length:12 }, (_, i) => `Cedar${i + 1}`));
    if (cantSplit || note) expect(result[0].text).not.toContain('Cedar');
    else expect(result[0].text).toContain('Cedar');
    expect(result.map(page => page.warnings)).toEqual(cantSplit || note ? [0, 1] : [0, 0]);
  });
}

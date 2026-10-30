import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle=fileURLToPath(new URL('../dist/pagination.bundle.js',import.meta.url));

for(const mode of ['recover','editor','exact','header'] as const) {
  test(`oversized merged grids retain ${mode} semantics`,async({page})=>{
    const rows=Array.from({length:6},(_,i)=>`<tr ${mode==='exact'&&i===0?'data-docx-row-height-rule="exact"':''}
      ${mode==='header'&&i===0?'data-docx-header-row="true"':''} data-docx-cant-split="true">
      ${i===0?'<td id="merge-cell" data-anchor="cell:body:merge" rowspan="0" colspan="2"><p>WillowSpan</p></td>':''}
      <td id="cell-${i}"><p>Willow${i}</p></td></tr>`).join('');
    await page.setContent(`<style>table{border-collapse:collapse;width:260pt}td{padding:0;vertical-align:top}p{margin:0;font:12pt/40pt Arial}</style>
      <div id="staging"><section data-section-index="0" data-page-width="300" data-page-height="160"
      data-content-width="260" data-content-height="120" data-margin-top="20" data-margin-bottom="20"
      data-margin-left="20" data-margin-right="20"><div><table>${rows}</table></div></section></div><div id="container"></div>`);
    await page.addScriptTag({path:bundle});
    const result=await page.evaluate(fragment=>{
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),document.getElementById('container'),
        {fragmentParagraphs:fragment,showPageNumbers:false}).paginate();
      const root=document.getElementById('container')!;
      return{text:root.textContent,tables:root.querySelectorAll('table').length,
        spans:Array.from(root.querySelectorAll('td'),cell=>cell.rowSpan),
        ids:Array.from(root.querySelectorAll('[id]'),element=>element.id),
        cells:root.querySelectorAll('td').length,
        kept:Array.from(root.querySelectorAll('tr')).every(row=>row.dataset.docxCantSplit==='true')};
    },mode!=='editor');
    expect(result.text?.match(/Willow\d/g)).toEqual(Array.from({length:6},(_,i)=>`Willow${i}`));
    expect(result.text?.match(/WillowSpan/g)).toHaveLength(1);
    expect(new Set(result.ids).size).toBe(result.ids.length);
    expect(result.ids.filter(id=>id==='merge-cell')).toHaveLength(1);
    expect(result.kept).toBe(true);
    if(mode==='recover') {
      expect(result.tables).toBeGreaterThan(1);
      expect(result.spans.every(span=>span===1)).toBe(true);
      expect(result.cells).toBe(12);
    } else {
      expect(result.tables).toBe(1);
      expect(result.spans).toContain(0);
      expect(result.cells).toBe(7);
    }
  });
}

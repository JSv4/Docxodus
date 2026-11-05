import { expect, test } from '@playwright/test';
import { getDocument, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { oversizedMergeDocx } from './docx-oversized-merge-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const rows of [6,46]) {
  test(`a ${rows}-row overlapping merged ledger preserves every cell and picture`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(oversizedMergeDocx(rows), { browser, outputs: ['html','pdf'],
      reviewProfile:'markup', markupPresentation:'word', commentProfile:'margin', commentColumnPlacement:'balloons', timeoutMs:120_000 });
    const task = getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
    try {
      const pdf=await task.promise, texts:string[]=[];let images=0;
      for(let index=1;index<=pdf.numPages;index++) {
        const p=await pdf.getPage(index), text=await p.getTextContent();
        texts.push(text.items.filter(item=>'str' in item).map(item=>item.str).join(' '));
        images+=(await p.getOperatorList()).fnArray.filter(op=>op===OPS.paintImageXObject||op===OPS.paintInlineImageXObject).length;
      }
      expect(images).toBe(2);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      expect(texts.join(' ').match(/Cedar\d+[ABC]/g)).toEqual(Array.from({length:rows},(_,i)=>
        ['A','B','C'].map(suffix=>`Cedar${String(i+1).padStart(2,'0')}${suffix}`)).flat());
      expect(texts.join(' ').match(/Copse\d+/g)).toEqual(Array.from({length:Math.floor(rows/2)},(_,i)=>`Copse${String(i+1).padStart(2,'0')}`));
      for(const label of ['CedarSpan','CedarOverlap','Cedar ledger complete.']) expect(texts.join(' ').split(label)).toHaveLength(2);
    } finally {await task.destroy();}
    await page.setContent(result.html!);
    const geometry=await page.locator('.page-content table').evaluateAll(tables=>tables.map(table=>{
      const band=table.closest('.page-content')!.getBoundingClientRect();
      return{overflow:table.getBoundingClientRect().bottom-band.bottom,spans:Array.from(table.querySelectorAll('td'),cell=>({rows:cell.rowSpan,columns:cell.colSpan})),
        headers:table.querySelectorAll('[data-docx-header-row=true]').length};
    }));
    expect(geometry.every(table=>table.overflow<=1&&table.headers===1)).toBe(true);
    if(rows===6) expect(geometry.flatMap(t=>t.spans).filter(cell=>cell.rows>1).map(cell=>cell.rows)).toEqual([6,3]);
    else expect(geometry.length).toBeGreaterThan(1);
    expect(geometry.flatMap(t=>t.spans).some(cell=>cell.columns===2)).toBe(true);
    const fragments=result.pageMap.fragments.filter(fragment=>fragment.anchorId.startsWith('tbl:'));
    expect(fragments).toHaveLength(geometry.length);
    expect(new Set(fragments.map(f=>f.anchorId)).size).toBe(1);
    expect(new Set(fragments.map(f=>f.fragmentId)).size).toBe(fragments.length);
  });
}

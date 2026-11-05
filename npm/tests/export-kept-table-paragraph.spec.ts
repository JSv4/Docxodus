import {expect,test} from '@playwright/test';
import {getDocument} from 'pdfjs-dist/legacy/build/pdf.mjs';
import {renderDocxArtifacts} from '../../npm-export/dist/index.js';
import {keptTableParagraphDocx} from './docx-kept-table-fixture.js';
test.use({launchOptions:{chromiumSandbox:true,...(process.env.DOCXODUS_CHROMIUM_PATH?{executablePath:process.env.DOCXODUS_CHROMIUM_PATH}:{})}});
for(const mode of ['lines','next','fitting'] as const) {
  test(`table cell paragraph ${mode} keeps preserve all source text`,async({browser,page})=>{
    test.setTimeout(120_000);
    const result=await renderDocxArtifacts(keptTableParagraphDocx(mode),{browser,outputs:['html','pdf'],reviewProfile:'markup',markupPresentation:'word',commentProfile:'margin',commentColumnPlacement:'balloons',timeoutMs:120_000});
    const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
    try{
      const pdf=await task.promise,pages:string[]=[];
      for(let i=1;i<=pdf.numPages;i++)pages.push((await(await pdf.getPage(i)).getTextContent()).items.filter(item=>'str'in item).map(item=>item.str).join(' '));
      const count=mode==='fitting'?20:83;
      expect(pages.join(' ').match(/Elm\d+/g)).toEqual(Array.from({length:count},(_,i)=>`Elm${String(i+1).padStart(3,'0')}`));
      expect(pages.join(' ').match(/Elm adjacent record\./g)).toHaveLength(1);
      expect(pages.join(' ').match(/Elm inventory complete\./g)).toHaveLength(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      if(mode==='fitting') {
        expect(pages.findIndex(text=>text.includes('Elm001'))).toBe(1);
        expect(pages[1].match(/Elm\d+/g)).toHaveLength(20);
      } else expect(pages.filter(text=>/Elm\d/.test(text)).every(text=>(text.match(/Elm\d+/g)?.length??0)>=2)).toBe(true);
    }finally{await task.destroy();}
    await page.setContent(result.html!);
    const geometry=await page.locator('.page-content table').evaluateAll(tables=>tables.map(table=>table.getBoundingClientRect().bottom-table.closest('.page-content')!.getBoundingClientRect().bottom));
    expect(geometry.every(overflow=>overflow<=1)).toBe(true);
  });
}

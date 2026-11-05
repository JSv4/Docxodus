import {expect,test} from '@playwright/test';
import {getDocument,OPS} from 'pdfjs-dist/legacy/build/pdf.mjs';
import {renderDocxArtifacts} from '../../npm-export/dist/index.js';
import {inlineLayoutAtomDocx} from './docx-inline-layout-atom-fixture.js';
test.use({launchOptions:{chromiumSandbox:true,...(process.env.DOCXODUS_CHROMIUM_PATH?{executablePath:process.env.DOCXODUS_CHROMIUM_PATH}:{})}});
for(const kind of ['shape','tab','tab-cell'] as const) {
  test(`whole ${kind} layout boxes continue without losing source content`,async({browser,page})=>{
    test.setTimeout(120_000);
    const result=await renderDocxArtifacts(inlineLayoutAtomDocx(kind),{browser,outputs:['html','pdf'],reviewProfile:'markup',markupPresentation:'word',commentProfile:'margin',commentColumnPlacement:'balloons',timeoutMs:120_000});
    const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
    try{
      const pdf=await task.promise,pages:string[]=[];let images=0;
      for(let i=1;i<=pdf.numPages;i++){
        const p=await pdf.getPage(i);
        pages.push((await p.getTextContent()).items.filter(item=>'str'in item).map(item=>item.str).join(' '));
        images+=(await p.getOperatorList()).fnArray.filter(op=>op===OPS.paintImageXObject||op===OPS.paintInlineImageXObject).length;
      }
      expect(pdf.numPages).toBeGreaterThan(1);
      expect(pdf.numPages).toBe(result.pageMap.pages.length);
      if(kind==='shape'){
        expect(pages.join(' ').match(/\[SHAPE\]/g)).toHaveLength(5);
        expect(images).toBe(1);
        expect(pages.join(' ').match(/Alder shape inventory complete\./g)).toHaveLength(1);
      }else{
        expect(pages.join(' ').match(/Willow\d+/g)).toEqual(Array.from({length:83},(_,i)=>`Willow${String(i+1).padStart(3,'0')}`));
        expect(pages.join(' ').match(/R7\./g)).toHaveLength(1);
        expect(pages.join(' ').match(/Willow tab inventory complete\./g)).toHaveLength(1);
        if(kind==='tab-cell')expect(pages.join(' ').match(/Willow neighbor\./g)).toHaveLength(1);
      }
    }finally{await task.destroy();}
    await page.setContent(result.html!);
    if(kind==='shape'){
      await expect(page.locator('.page-content [data-docx-shape="unrendered"]')).toHaveCount(5);
      const size=await page.locator('.page-content img').evaluate(image=>({width:image.getBoundingClientRect().width*0.75,height:image.getBoundingClientRect().height*0.75}));
      expect(size.width).toBeCloseTo(320,1);expect(size.height).toBeCloseTo(590,1);
    }else await expect(page.locator('.page-content [data-docx-tab]')).toHaveCount(1);
  });
}

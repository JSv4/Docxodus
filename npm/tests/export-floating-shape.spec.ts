import {expect, test} from '@playwright/test';
import {getDocument, OPS} from 'pdfjs-dist/legacy/build/pdf.mjs';
import {renderDocxArtifacts} from '../../npm-export/dist/index.js';
import {floatingShapeDocx, vmlShapeDocx} from './docx-floating-shape-fixture.js';

test.use({launchOptions:{chromiumSandbox:true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? {executablePath:process.env.DOCXODUS_CHROMIUM_PATH} : {}),
}});

for(const grouped of [false,true]) for(const wrap of ['none','topAndBottom'] as const)
  for(const reviewProfile of ['markup','final'] as const) {
  test(`${grouped?'group':'shape'} ${wrap} exports with its ink (${reviewProfile})`,async({browser})=>{
    test.setTimeout(120_000);
    const result=await renderDocxArtifacts(floatingShapeDocx({grouped,wrap}),{
      browser,outputs:['html','pdf'],reviewProfile,commentProfile:'hidden',timeoutMs:120_000,
    });
    expect(result.pageCount).toBe(1);
    expect(result.pdf!.byteLength).toBeGreaterThan(0);
    expect(result.html).toContain('fill="#19647E"');
    expect(result.renderReport.unsupportedContent).toEqual([]);
    expect(result.pageMap.fragments).toHaveLength(3);
  });
}

for(const [name,source] of [
  ['unsupported floating pentagon',()=>floatingShapeDocx({geometry:'pentagon'})],
  ['interior wrapping',()=>floatingShapeDocx({wrap:'square',x:100})],
  ['VML without an extent',()=>vmlShapeDocx(false)],
] as const) test(`${name} reaches the unsupported-content policy`,async({browser})=>{
  test.setTimeout(120_000);
  const input=source();
  const options={browser,outputs:['pdf'] as const,reviewProfile:'markup' as const,
    commentProfile:'hidden' as const,timeoutMs:120_000};
  const result=await renderDocxArtifacts(input,options);
  expect(result.pageCount).toBe(1);
  expect(result.warnings.map(w=>w.code)).toContain('unsupported_content_placeholder');
  expect(result.renderReport.unsupportedContent).toEqual([
    expect.objectContaining({contentType:'Other',action:'placeholder'}),
  ]);
  await expect(renderDocxArtifacts(input,{...options,unsupportedContent:'strict'}))
    .rejects.toMatchObject({code:'resource_policy_failure'});
});

for(const [name,source,element] of [
  ['hexagon',()=>floatingShapeDocx({geometry:'hexagon'}),'polygon'],
  ['plain VML rectangle',()=>vmlShapeDocx(),'rect'],
] as const) test(`${name} exports its native geometry without a placeholder label`,async({browser})=>{
  const result=await renderDocxArtifacts(source(),{browser,outputs:['html','pdf'],
    reviewProfile:'final',commentProfile:'hidden',timeoutMs:120_000});
  expect(result.pageCount).toBe(1);
  expect(result.html).toContain(`<${element}`);
  expect(result.renderReport.unsupportedContent).toEqual([]);
  const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
  try {
    const page=await(await task.promise).getPage(1);
    const text=await page.getTextContent();
    const words=text.items.filter(item=>'str' in item).map(item=>item.str).join('');
    expect(words).not.toContain('[SHAPE]');
    const ops=await page.getOperatorList();
    expect(ops.fnArray.some((op,i)=>op===OPS.setFillRGBColor && ops.argsArray[i][0]==='#19647e')).toBe(true);
  } finally {await task.destroy();}
});

for(const behind of [false,true]) test(`PDF paints a ${behind?'background':'foreground'} shape above the paper and on the correct side of text`,async({browser})=>{
  const result=await renderDocxArtifacts(floatingShapeDocx({wrap:'none',mixed:true,behind}),{
    browser,outputs:['pdf'],reviewProfile:'final',commentProfile:'hidden',timeoutMs:120_000,
  });
  const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
  try {
    const page=await(await task.promise).getPage(1);
    const operators=await page.getOperatorList();
    const colorIndex=(color:string)=>operators.fnArray.findIndex((op,i)=>
      op===OPS.setFillRGBColor && operators.argsArray[i][0]===color);
    const teal=colorIndex('#19647e'),paper=colorIndex('#ffffff');
    const glyphs=operators.fnArray.flatMap((op,i)=>op===OPS.showText
      ? operators.argsArray[i][0].flatMap((glyph:{unicode?:string}|number)=>
        typeof glyph==='number'?[]:Array.from(glyph.unicode??'',letter=>({letter,index:i}))) : []);
    const offset=glyphs.map(glyph=>glyph.letter).join('').indexOf('AnchorWord');
    const text=offset<0?-1:glyphs[offset].index;
    expect(paper).toBeGreaterThanOrEqual(0);
    expect(teal).toBeGreaterThan(paper);
    expect(text).toBeGreaterThanOrEqual(0);
    if(behind) expect(teal).toBeLessThan(text);
    else expect(teal).toBeGreaterThan(text);
  } finally {await task.destroy();}
});

test('relative size changes the PDF rectangle without scaling its physical outline',async({browser})=>{
  const result=await renderDocxArtifacts(floatingShapeDocx({wrap:'none',relativeWidth:50000,relativeHeight:12500,outline:true}),{
    browser,outputs:['pdf'],reviewProfile:'final',commentProfile:'hidden',timeoutMs:120_000,
  });
  const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
  try {
    const ops=await(await(await task.promise).getPage(1)).getOperatorList();
    let sx=1,sy=1,lineWidth=1;const stack:[number,number,number][]=[];const widths:number[]=[];
    ops.fnArray.forEach((op,i)=>{
      const a=ops.argsArray[i];
      if(op===OPS.save) stack.push([sx,sy,lineWidth]);
      else if(op===OPS.restore) [sx,sy,lineWidth]=stack.pop()!;
      else if(op===OPS.transform) {sx*=Math.hypot(a[0],a[1]);sy*=Math.hypot(a[2],a[3]);}
      else if(op===OPS.setLineWidth) lineWidth=a[0];
      else if(op===OPS.setGState) {for(const [key,value] of a[0]) if(key==='LW') lineWidth=value;}
      else if(op===OPS.constructPath && a[0]===OPS.stroke) widths.push(lineWidth*Math.sqrt(sx*sy));
    });
    expect(widths).toHaveLength(1);
    expect(widths[0]).toBeCloseTo(1.5,2);
  } finally {await task.destroy();}
});

for(const story of ['header','footer'] as const) test(`${story} drawings retain their page position and source mapping`,async({browser})=>{
  const top=story==='header'?36:660;
  const result=await renderDocxArtifacts(floatingShapeDocx({story,wrap:'none',behind:true,vertical:'page',y:top}),{
    browser,outputs:['html','pdf'],reviewProfile:'final',commentProfile:'hidden',timeoutMs:120_000,
  });
  expect(result.pageCount).toBe(1);
  expect(result.renderReport.unsupportedContent).toEqual([]);
  expect(result.pageMap.fragments.some(f=>f.story!=='body')).toBe(true);
  const page=await browser.newPage();
  try {
    await page.setContent(result.html!);await page.evaluate(()=>document.fonts.ready);
    const f=await page.evaluate(()=>{
      const anchor=document.querySelector<HTMLElement>('.page-box [data-docx-shape-anchor]')!;
      const paper=anchor.closest<HTMLElement>('.page-box')!.getBoundingClientRect();
      const ink=anchor.querySelector('svg')!.getBoundingClientRect();
      return {top:(ink.top-paper.top)*.75,height:ink.height*.75,parent:anchor.parentElement?.className};
    });
    expect(f.top).toBeCloseTo(top,0);expect(f.height).toBeCloseTo(72,0);expect(f.parent).toBe('page-box');
  } finally {await page.close();}
});

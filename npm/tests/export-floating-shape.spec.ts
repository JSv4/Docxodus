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
  ['floating hexagon',()=>floatingShapeDocx({geometry:'hexagon'})],
  ['interior wrapping',()=>floatingShapeDocx({wrap:'square',x:100})],
  ['plain VML rectangle',()=>vmlShapeDocx()],
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

import { expect, test } from '@playwright/test';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';
import { decodePng } from './visual-parity/png.js';

let script: string;
test.beforeAll(async () => {
  const result = await build({ entryPoints: [fileURLToPath(new URL('../dist/table-borders.js',import.meta.url))],
    bundle:true,format:'iife',globalName:'TableBorders',write:false });
  script = result.outputFiles[0].text;
});
const edgeMap = (top=3,bottom=3) => JSON.stringify([
  ['top',0,1,top,'#205080'],['bottom',0,1,bottom,'#205080'],
  ['vertical',0,0,3,'#205080'],['vertical',1,1,3,'#205080'],
]);
const row = (top=3,bottom=3,label='Cypress') => `<tr data-docx-table-border-edges='${edgeMap(top,bottom)}'>`
  + `<td data-docx-exact-cell-borders="true" style="--docx-cell-border-top:${top}pt;--docx-cell-border-bottom:${bottom}pt;`
  + '--docx-cell-border-left:3pt;--docx-cell-border-right:3pt;border:3pt solid #205080;padding:0">'
  + `<p style="margin:0;line-height:42pt">${label}</p></td></tr>`;
const table = (rows=row()) => '<table data-docx-exact-table-borders="0" style="border-collapse:collapse;table-layout:fixed;width:120pt">'
  + '<colgroup><col style="width:120pt"></colgroup><tbody>'+rows+'</tbody></table>';

for (const mode of ['ordinary','rotated','reflected','clipped','rotated-clip'] as const) {
  test(`${mode} borders keep local geometry, closed corners, and ancestor clips`, async ({page}) => {
    const transform = mode === 'rotated' || mode === 'rotated-clip' ? 'rotate(30deg)' : mode === 'reflected' ? 'scale(-1,1)' : 'none';
    const clipped = mode.includes('clip');
    await page.setContent(`<style>body{margin:0;background:white}.page-box{position:relative;width:600px;height:400px}
      #holder{position:absolute;left:200px;top:100px;transform:${transform};transform-origin:0 0;${clipped?'height:20px;overflow:hidden;':''}}
      </style><div class="page-box"><div id="holder">${table()}</div></div>`);
    await page.addScriptTag({content:script});
    await page.evaluate(() => {
      const api = (window as any).TableBorders;
      api.prepareTableBorders(document.body); api.paintTableBorders(document.body);
    });
    // Independent local points: both outer top corners and the left side must be painted;
    // the bottom edge must disappear under the 20px ancestor clip.
    const points = await page.evaluate(({transform,clipped}) => {
      const m = new DOMMatrix(transform === 'none' ? undefined : getComputedStyle(document.querySelector('#holder')!).transform);
      const project = (x:number,y:number) => { const p=m.transformPoint({x,y}); return {x:p.x+200,y:p.y+100}; };
      return {paint:[project(clipped?1:-1,1),project(clipped?159:161,1),project(clipped?1:-1,10)],
        bottom:project(80,62),clipped};
    },{transform,clipped});
    const png = decodePng(await page.screenshot());
    const blue = (point:{x:number;y:number}) => {
      const index=(Math.floor(point.y)*png.width+Math.floor(point.x))*4;
      return [...png.data.subarray(index,index+3)];
    };
    for (const point of points.paint) expect(blue(point)).toEqual([32,80,128]);
    expect(blue(points.bottom)).toEqual(clipped ? [255,255,255] : [32,80,128]);
  });
}

for (const alignment of ['start','end','rtl-end'] as const) test(`${alignment} continuations reserve their visible outer edges before fitting`,async ({page}) => {
  const side=alignment==='rtl-end'?'left':'right';
  const rows=[row(.25,3,'CypressOne'),row(3,.25,'CypressTwo'),row(.25,.25,'CypressThree')].map((html,index)=>{
    const edges=JSON.parse(edgeMap([.25,3,.25][index],[3,.25,.25][index]));
    for (const edge of edges) if (edge[0]==='vertical' && edge[1]===1) edge[3]=[.25,3,.25][index];
    return html.replace(/data-docx-table-border-edges='[^']*'/,`data-docx-table-border-edges='${JSON.stringify(edges)}'`);
  }).join('');
  let markup=table(rows);
  if (alignment!=='start') markup=markup.replace('<table ',`<table data-docx-grid-align-end="true" data-docx-grid-margin-side="${side}" `)
    .replace('table-layout:fixed;',`table-layout:fixed;direction:${alignment==='rtl-end'?'rtl':'ltr'};margin-${side}:1.5pt;margin-${side==='left'?'right':'left'}:auto;`);
  await page.setContent(`<style>body{margin:0}p{font:12pt Arial;line-height:80pt!important}</style>
    <div id="staging"><section data-section-index="0" data-page-width="300" data-page-height="160"
      data-content-width="260" data-content-height="120" data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20">
      <div>${markup}</div>
    </section></div><div id="container"></div>`);
  await page.addScriptTag({content:script});
  await page.addScriptTag({path:fileURLToPath(new URL('../dist/pagination.bundle.js',import.meta.url))});
  const pages = await page.evaluate((alignment) => {
    const api=(window as any).TableBorders;
    api.prepareTableBorders(document.getElementById('staging'));
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),document.getElementById('container'),
      {fragmentParagraphs:true,showPageNumbers:false,prepareTableFragment:api.refreshTableBorderLayout}).paginate();
    api.paintTableBorders(document.getElementById('container'));
    return Array.from(document.querySelectorAll('#container .page-box'),page => {
      const y=page.getBoundingClientRect().top;
      const boxes=Array.from(page.querySelectorAll('svg[data-docx-table-border-paint] g rect'),r=>r.getBoundingClientRect());
      return {text:page.textContent,top:Math.min(...boxes.map(r=>r.top-y))*.75,
        end:(alignment==='rtl-end'?Math.min(...boxes.map(r=>r.left)):Math.max(...boxes.map(r=>r.right)))-page.getBoundingClientRect().left};
    });
  },alignment);
  expect(pages).toHaveLength(3);
  for (const [index,page] of pages.entries()) {
    expect(page.text).toContain(['CypressOne','CypressTwo','CypressThree'][index]);
    expect(page.top).toBeCloseTo(20,1);
    if (alignment!=='start') expect(page.end*.75).toBeCloseTo(alignment==='rtl-end'?20:280,1);
  }
});

test('repeated headers resolve the new shared edge once and reserve its winning width',async ({page}) => {
  const colors=['#c02020','#205080'];
  const join=JSON.stringify([['top',0,1,3,colors[0]]]);
  const rows=[row(.25,3,'CypressHeader'),row(3,.75,'CypressFirst'),row(.75,.25,'CypressLast')].map((html,index)=>{
    const edges=JSON.parse(edgeMap([.25,3,.75][index],[3,.75,.25][index]));
    for (const edge of edges) if (edge[3]===3 && edge[0]!=='vertical') edge[4]=colors[0];
    return html.replace(/data-docx-table-border-edges='[^']*'/,`data-docx-table-border-edges='${JSON.stringify(edges)}'`)
      .replace('<tr ',`<tr ${index===0?'data-docx-header-row="true"':`data-docx-table-header-edges='${join}'`} `)
      .replace('line-height:42pt',`line-height:${index===0?20:80}pt`);
  }).join('');
  await page.setContent(`<style>body{margin:0}p{font:12pt Arial}</style><div id="staging">
    <section data-section-index="0" data-page-width="300" data-page-height="160" data-content-width="260" data-content-height="120"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><div>${table(rows)}</div></section>
    </div><div id="container"></div>`);
  await page.addScriptTag({content:script});
  await page.addScriptTag({path:fileURLToPath(new URL('../dist/pagination.bundle.js',import.meta.url))});
  const result=await page.evaluate(()=>{
    const api=(window as any).TableBorders;
    api.prepareTableBorders(document.getElementById('staging'));
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),document.getElementById('container'),
      {fragmentParagraphs:true,showPageNumbers:false,prepareTableFragment:api.refreshTableBorderLayout}).paginate();
    api.paintTableBorders(document.getElementById('container'));
    return Array.from(document.querySelectorAll('#container .page-box'),page=>{
      const rows=Array.from(page.querySelectorAll('tr')),boundary=rows[0].getBoundingClientRect().bottom;
      return {text:page.textContent,padding:(rows[1].querySelector('p')!.getBoundingClientRect().top-rows[1].getBoundingClientRect().top)*.75,
        borders:Array.from(page.querySelectorAll('svg[data-docx-table-border-paint] g rect')).filter(rect=>{
          const box=rect.getBoundingClientRect();return Math.abs(box.top+box.height/2-boundary)<.01;
        }).map(rect=>({color:rect.getAttribute('fill'),height:rect.getBoundingClientRect().height*.75}))};
    });
  });
  expect(result).toHaveLength(2);
  expect(result[1].text).toContain('CypressLast');
  for (const page of result) {
    expect(page.borders).toEqual([{color:'#c02020',height:3}]);
    expect(page.padding).toBeCloseTo(1.5,2);
  }
});

for (const nested of [false,true]) test(`foreground drawings cover ${nested?'transformed textbox':'body'} table borders`,async ({page})=>{
  await page.setContent(`<style>body{margin:0}.page-box{position:relative;width:600px;height:400px}
    #story{position:absolute;left:120.375px;top:80.625px;transform-origin:0 0;${nested?'transform:rotate(15deg);z-index:1;':''}}
    #cover{position:absolute;left:30px;top:0;width:60px;height:30px;background:#fff0c0;z-index:1}
    </style><div class="page-box"><div id="story"><div>${table()}</div><div id="cover"></div></div></div>`);
  await page.addScriptTag({content:script});
  const point=await page.evaluate(()=>{
    const api=(window as any).TableBorders;api.prepareTableBorders(document.body);api.paintTableBorders(document.body);
    const style=getComputedStyle(document.getElementById('story')!);
    const matrix=new DOMMatrix(style.transform==='none'?undefined:style.transform);
    const point=matrix.transformPoint({x:60,y:2});return {x:point.x+120.375,y:point.y+80.625};
  });
  const png=decodePng(await page.screenshot());
  const index=(Math.floor(point.y)*png.width+Math.floor(point.x))*4;
  expect([...png.data.subarray(index,index+3)]).toEqual([255,240,192]);
});

// HTML transform layers round their paint origins. The border must share that origin with
// surrounding HTML ink, while retaining its declared subpixel width inside the same context.
for (const mode of ['fraction','rotate','nested-rotate','nested-reflection'] as const) {
  test(`${mode} PDF borders stay aligned with neighboring HTML paint`,async ({page})=>{
    const outer=mode==='fraction'?'none':mode==='rotate'?'rotate(30deg)':'rotate(17deg)';
    const inner=mode==='nested-rotate'?'rotate(-31deg)':mode==='nested-reflection'?'scale(-1.25,.8)':'none';
    const thin=table(row(.25,.25)).replace(/3pt/g,'.25pt').replace(/,3,/g,',0.25,');
    await page.setContent(`<style>@page{size:612pt 792pt;margin:0}body{margin:0}.page-box{position:relative;width:612pt;height:792pt;overflow:hidden}
      #outer{position:absolute;left:240.3125px;top:150.4375px;transform:${outer};transform-origin:0 0}
      #inner{position:relative;left:20.21875px;top:30.34375px;transform:${inner};transform-origin:0 0}
      #reference{position:absolute;left:0;top:100px;width:16px;height:16px;background:#d06020}
      </style><div class="page-box"><div id="outer"><div id="inner">${thin}<div id="reference"></div></div></div>`);
    await page.addScriptTag({content:script});
    const expected=await page.evaluate(()=>{
      const api=(window as any).TableBorders;api.prepareTableBorders(document.body);api.paintTableBorders(document.body);
      const rect=document.querySelector<SVGRectElement>('[data-docx-table-border-paint] g rect')!;
      const c=rect.getScreenCTM()!,m=new DOMMatrix([c.a,c.b,c.c,c.d,c.e,c.f]);
      const x=rect.x.baseVal.value,y=rect.y.baseVal.value,w=rect.width.baseVal.value,h=rect.height.baseVal.value;
      const rule=[[x,y],[x+w,y],[x+w,y+h],[x,y+h]].map(([x,y])=>{const p=m.transformPoint({x,y});return[p.x*.75,p.y*.75];});
      const reference=[[0,0],[16,0],[16,16],[0,16]].map(([x,y])=>{
        const probe=document.createElement('span');probe.style.cssText=`position:absolute;left:${x}px;top:${y}px;width:0;height:0`;
        document.getElementById('reference')!.append(probe);const box=probe.getBoundingClientRect();probe.remove();return[box.left*.75,box.top*.75];
      });
      return {rule,reference};
    });
    const {getDocument,OPS,Util}=await import('pdfjs-dist/legacy/build/pdf.mjs');
    const task=getDocument({data:new Uint8Array(await page.pdf({preferCSSPageSize:true,printBackground:true})),verbosity:0});
    try {
      const pdf=await task.promise, pdfPage=await pdf.getPage(1),ops=await pdfPage.getOperatorList();
      let matrix=[1,0,0,1,0,0],color='';
      const stack:{matrix:number[];color:string}[]=[],shapes:{color:string;points:number[][]}[]=[];
      ops.fnArray.forEach((op,index)=>{
        const args=ops.argsArray[index];
        if(op===OPS.save)stack.push({matrix:[...matrix],color});
        else if(op===OPS.restore)({matrix,color}=stack.pop()!);
        else if(op===OPS.transform)matrix=Util.transform(matrix,args);
        else if(op===OPS.setFillRGBColor)color=args[0];
        else if(op===OPS.constructPath && (color==='#205080'||color==='#d06020')){
          const values=Array.from(args[1][0]) as number[],points:number[][]=[];
          for(let i=0;i<values.length;){
            const command=values[i++];if(command===4)continue;
            expect([0,1]).toContain(command);
            const p=[values[i++],values[i++]];Util.applyTransform(p,matrix);points.push([p[0],792-p[1]]);
          }
          shapes.push({color,points});
        }
      });
      const reference=shapes.find(shape=>shape.color==='#d06020')!.points;
      const centroid=(points:number[][])=>[0,1].map(axis=>points.reduce((n,p)=>n+p[axis],0)/points.length);
      const a=centroid(reference),b=centroid(expected.reference),shift=a.map((n,i)=>n-b[i]);
      const wanted=expected.rule.map(p=>p.map((n,i)=>n+shift[i]));
      const errors=shapes.filter(s=>s.color==='#205080').map(shape=>Math.max(...wanted.map(p=>Math.min(...shape.points.map(q=>Math.hypot(p[0]-q[0],p[1]-q[1]))))));
      expect(Math.min(...errors)).toBeLessThan(.015);
      // Matching all four corners also checks both endpoints and the transformed 0.25pt width.
    } finally {await task.destroy();}
  });
}

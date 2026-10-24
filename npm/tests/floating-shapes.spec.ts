import {expect, test, type Page} from '@playwright/test';
import {floatingShapeDocx, type FloatingShapeOptions} from './docx-floating-shape-fixture.js';

async function layout(page: Page, options: FloatingShapeOptions, paginated = true) {
  await page.goto('/worker-test-harness.html');
  await page.addScriptTag({url:'/pagination.bundle.js'});
  const facts = await page.evaluate(async ({bytes, paginated}) => {
    await (window as any).createDocxodusWorker();
    const worker = (window as any).DocxodusWorker;
    try {
      const html = await worker.convertDocxToHtml(new Uint8Array(bytes), {
        paginationMode:paginated ? 1 : 0, renderTrackedChanges:true,
        renderUnsupportedContentPlaceholders:true,
      });
      const parsed = new DOMParser().parseFromString(html,'text/html');
      document.head.innerHTML = parsed.head.innerHTML;
      document.body.innerHTML = '<div id="host"></div>';
      const host = document.querySelector<HTMLElement>('#host')!;
      if (paginated) (window as any).DocxodusPagination.paginateHtml(html,host,{showPageNumbers:false});
      else host.innerHTML = parsed.body.innerHTML;
      await document.fonts.ready;
      // paginateHtml retains a hidden staging copy. Inspect only physical page content.
      const firstShape = host.querySelector<HTMLElement>(paginated
        ? '.page-box [data-docx-shape-anchor]' : '[data-docx-shape-anchor]')!;
      const root = firstShape.closest<HTMLElement>('.page-box') ?? host;
      const paragraphs = Array.from(root.querySelectorAll<HTMLElement>('p'));
      const anchorParagraph = firstShape.closest<HTMLElement>('p') ?? root.querySelector<HTMLElement>(
        `[data-source-anchor-id="${CSS.escape(firstShape.dataset.docxAnchorHost!)}"]`)!;
      const afterParagraph = paragraphs[paragraphs.indexOf(anchorParagraph)+1];
      const origin = root.getBoundingClientRect();
      const rect = (box:DOMRect) => ({left:(box.left-origin.left)*.75,top:(box.top-origin.top)*.75,
        width:box.width*.75,height:box.height*.75,right:(box.right-origin.left)*.75,bottom:(box.bottom-origin.top)*.75});
      const textRect = (p:HTMLElement) => {
        const walker = document.createTreeWalker(p,NodeFilter.SHOW_TEXT);
        for(let n=walker.nextNode();n;n=walker.nextNode()) if(n.textContent?.trim()) {
          const range=document.createRange();range.selectNodeContents(n);return rect(range.getBoundingClientRect());
        }
        return rect(p.getBoundingClientRect());
      };
      const shapes = Array.from(root.querySelectorAll<HTMLElement>('[data-docx-drawing-anchor]'));
      const svg = shapes.map(s=>s.querySelector<SVGSVGElement>('svg')!).filter(Boolean);
      const painted = svg[0]?.getBoundingClientRect();
      const atCenter = painted ? document.elementsFromPoint(painted.left+painted.width/2,painted.top+painted.height/2)
        .map(e=>e.closest('[data-docx-drawing-anchor]')?.querySelector('rect')?.getAttribute('fill')).filter(Boolean) : [];
      const anchorText = textRect(anchorParagraph);
      const textPoint = {x:(anchorText.left+anchorText.width/2)/.75+origin.left,
        y:(anchorText.top+anchorText.height/2)/.75+origin.top};
      const textPaint = document.elementsFromPoint(textPoint.x,textPoint.y).map(e=>
        e.closest('[data-docx-drawing-anchor]') ? 'shape' : e.textContent==='AnchorWord' ? 'text' : null).filter(Boolean);
      return {pages:host.querySelectorAll('.page-box').length,page:Number(root.dataset.pageNumber??1), shapes:svg.map((s,i)=>({
        ...rect(s.getBoundingClientRect()),fill:s.querySelector('rect')?.getAttribute('fill'),
        z:getComputedStyle(shapes[i]).zIndex,
      })),anchor:rect(anchorParagraph.getBoundingClientRect()),anchorText,after:textRect(afterParagraph),
        lineHeight:parseFloat(getComputedStyle(anchorParagraph).lineHeight)*.75,
        painted:Array.from(new Set(atCenter)),textPaint:Array.from(new Set(textPaint)),
        unsupported:root.querySelectorAll('[data-content-type]').length};
    } finally { worker.terminate(); }
  },{bytes:Array.from(floatingShapeDocx(options)),paginated});
  await test.info().attach('layout.json',{body:JSON.stringify(facts,null,2),contentType:'application/json'});
  return facts;
}

for(const grouped of [false,true]) for(const paginated of [false,true]) {
  test(`${grouped?'group':'shape'} top/bottom wrap keeps its extent clear (${paginated?'pages':'flow'})`,async({page})=>{
    const f=await layout(page,{grouped,wrap:'topAndBottom'},paginated);
    expect(f.shapes).toHaveLength(1);
    expect(f.shapes[0].width).toBeCloseTo(144,0);expect(f.shapes[0].height).toBeCloseTo(72,0);
    expect(f.after.top).toBeGreaterThanOrEqual(f.shapes[0].bottom-.5);
    expect(f.unsupported).toBe(0);
  });
  test(`${grouped?'group':'shape'} no-wrap does not consume a body line (${paginated?'pages':'flow'})`,async({page})=>{
    const f=await layout(page,{grouped,wrap:'none'},paginated);
    expect(f.shapes).toHaveLength(1);
    expect(f.after.top-f.anchor.top).toBeLessThan(f.lineHeight+4);
    expect(f.unsupported).toBe(0);
  });
}

for(const paginated of [false,true]) {
test(`relativeHeight orders graphics independently of source order (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap:'none',layers:[{order:4294967295,fill:'19647E'},{order:2147483648,fill:'FF9900'}]},paginated);
  expect(f.shapes).toHaveLength(2);
  expect(f.painted[0]).toBe('#19647E');
});

test(`behindDoc places a graphic below a foreground drawing (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap:'none',layers:[{order:1,fill:'19647E'},{order:90,fill:'FF9900',behind:true}]},paginated);
  expect(f.shapes).toHaveLength(2);
  expect(f.painted[0]).toBe('#19647E');
});

for(const behind of [false,true]) test(`a ${behind?'background':'foreground'} drawing respects body text (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap:'none',mixed:true,behind},paginated);
  expect(f.textPaint[0]).toBe(behind?'text':'shape');
});

for(const wrap of ['none','topAndBottom'] as const) test(`mixed text follows ${wrap} clearance (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap,mixed:true},paginated);
  if(wrap==='topAndBottom') expect(f.anchorText.top).toBeGreaterThanOrEqual(f.shapes[0].bottom-.5);
  else expect(f.anchorText.top).toBeLessThan(f.shapes[0].bottom-30);
  expect(f.unsupported).toBe(0);
});

test(`overlapping top/bottom drawings reserve one extent (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap:'topAndBottom',layers:[{order:1,fill:'19647E'},{order:2,fill:'FF9900'}]},paginated);
  expect(f.shapes).toHaveLength(2);
  expect(f.after.top).toBeGreaterThanOrEqual(f.shapes[0].bottom-.5);
  expect(f.after.top-f.shapes[0].bottom).toBeLessThan(30);
});

test(`top/bottom clearance includes an offset and bottom distance (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap:'topAndBottom',y:24,bottomDistance:9},paginated);
  expect(f.shapes[0].top-f.anchor.top).toBeCloseTo(24,0);
  expect(f.after.top).toBeGreaterThanOrEqual(f.shapes[0].bottom+8.5);
});

for(const wrap of ['square','tight','through'] as const) test(`column-edge ${wrap} wraps text beside its extent (${paginated?'pages':'flow'})`,async({page})=>{
  const f=await layout(page,{wrap},paginated);
  expect(f.after.top).toBeLessThan(f.shapes[0].bottom);
  expect(f.after.left).toBeGreaterThan(f.shapes[0].left+20);
  if(wrap==='square') expect(f.after.left).toBeGreaterThanOrEqual(f.shapes[0].right-.5);
  else expect(f.after.left).toBeLessThan(f.shapes[0].right-2);
  expect(f.unsupported).toBe(0);
});
}

test('top/bottom clearance and its host move together at a page boundary',async({page})=>{
  const f=await layout(page,{wrap:'topAndBottom',before:Array.from({length:38},(_,i)=>`Inspection line ${i+1}`)});
  expect(f.page).toBe(2);
  expect(f.after.top).toBeGreaterThanOrEqual(f.shapes[0].bottom-.5);
  expect(f.shapes[0].top).toBeCloseTo(72,0);
});

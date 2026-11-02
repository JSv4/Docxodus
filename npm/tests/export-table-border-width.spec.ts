import { expect, test } from '@playwright/test';
import { getDocument, OPS, Util } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { tableBorderDocx, type TableBorderCase } from './docx-table-border-fixture.js';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}),
} });

for (const size of [2, 4, 8, 12]) for (const kind of ['ordinary', 'style', 'merged', 'none'] as TableBorderCase[]) {
  if (kind === 'none' && size !== 2) continue;
  test(`${kind} table paints its declared ${size}/8pt borders`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(tableBorderDocx(size, kind), {
      browser, outputs: ['html','pdf'], reviewProfile: 'final', commentProfile: 'hidden',
    });
    expect(result.pageCount).toBe(1);
    await page.setContent(result.html!);
    const widths = await page.locator('td').evaluateAll(cells => cells.map(cell => cell.getBoundingClientRect().width*.75));
    expect(widths).toEqual(kind === 'merged' ? [240,120,120] : [120,120]);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdfPage = await (await task.promise).getPage(1);
      const operators = await pdfPage.getOperatorList();
      let transform = [1,0,0,1,0,0], color = '';
      const stack: { transform: number[]; color: string }[] = [], thickness: number[] = [];
      operators.fnArray.forEach((op,i) => {
        const args = operators.argsArray[i];
        if (op === OPS.save) stack.push({ transform: [...transform], color });
        else if (op === OPS.restore) ({ transform, color } = stack.pop()!);
        else if (op === OPS.transform) transform = Util.transform(transform,args);
        else if (op === OPS.setFillRGBColor) color = args[0];
        else if (op === OPS.constructPath && color === '#205080') {
          const bounds = args[2], a = [bounds[0],bounds[1]], b = [bounds[2],bounds[3]];
          Util.applyTransform(a,transform); Util.applyTransform(b,transform);
          thickness.push(Math.min(Math.abs(b[0]-a[0]),Math.abs(b[1]-a[1])));
        }
      });
      if (kind === 'none') expect(thickness).toEqual([]);
      else {
        expect(thickness.length).toBeGreaterThanOrEqual(5);
        for (const width of thickness) expect(width).toBeCloseTo(size/8, 2);
      }
      const text = (await pdfPage.getTextContent()).items.filter(item => 'str' in item).map(item => item.str).join(' ');
      for (const label of ['CypressNorth','CypressSouth','CypressAfter']) expect(text).toContain(label);
    } finally { await task.destroy(); }
  });
}

test('export reserves a repeated header\'s winning border before fitting each continuation',async ({browser,page})=>{
  const result=await renderDocxArtifacts(tableBorderDocx(2,'repeated'),{
    browser,outputs:['html','pdf'],reviewProfile:'final',commentProfile:'hidden',
  });
  expect(result.pageCount).toBe(2);
  await page.setContent(result.html!);
  const continuation=page.locator('tr[data-docx-repeated-header-row="true"] + tr').first();
  expect(await continuation.locator('td').first().evaluate(cell=>parseFloat(getComputedStyle(cell).paddingTop)*.75)).toBe(6);
  const task=getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
  try {
    const pdf=await task.promise;
    for(let i=1;i<=2;i++) {
      const text=(await (await pdf.getPage(i)).getTextContent()).items.filter(item=>'str' in item).map(item=>item.str).join(' ');
      expect(text.match(/CypressHeader/g)).toEqual(['CypressHeader']);
      expect(text.match(/CypressRow\d+/g)).toEqual(Array.from({length:3},(_,row)=>`CypressRow${(i-1)*3+row+1}`));
    }
  } finally {await task.destroy();}
});

// Word's captured #1219 rule: margins are measured from the grid center, while zero
// margins still keep text outside the border. Legacy text-aligned tables put text at indent.
for (const [placement, margin, size, offset] of [
  ['legacy',90,2,4.5], ['legacy',90,8,4.5], ['modern',90,8,5],
  ['legacy',0,24,1.5], ['modern',0,24,3], ['text',0,24,0], ['text',90,8,0],
] as const) {
  test(`${placement} ${size}/8pt table keeps the ${margin}-twip text margin`, async ({ browser, page }) => {
    const result = await renderDocxArtifacts(tableBorderDocx(size,'ordinary',margin,placement), {
      browser, outputs: ['html','pdf'], reviewProfile: 'final', commentProfile: 'hidden',
    });
    await page.setContent(result.html!);
    const width = await page.locator('td p').first().evaluate(p => p.getBoundingClientRect().width*.75);
    expect(width).toBeCloseTo(120 - 2*Math.max(margin/20,size/16),1);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const text = (await (await (await task.promise).getPage(1)).getTextContent()).items.filter(item => 'str' in item);
      const start = text.find(item => item.str === 'CypressNorth')!;
      expect(start.transform[4]).toBeCloseTo(72+offset,1);
    } finally { await task.destroy(); }
  });
}

test('an auto row height does not impose the recorded minimum',async ({browser,page}) => {
  const heights:number[]=[];
  for (const kind of ['ordinary','auto'] as const) {
    const result=await renderDocxArtifacts(tableBorderDocx(2,kind),{browser,outputs:['html'],reviewProfile:'final',commentProfile:'hidden'});
    await page.setContent(result.html!);
    heights.push(await page.locator('tr').evaluate(row=>row.getBoundingClientRect().height));
  }
  expect(heights[1]).toBeCloseTo(heights[0],2);
});

for (const placement of ['end','rtl-end'] as const) {
  test(`${placement} keeps the outer painted edge at the content boundary`,async ({browser,page}) => {
    const result=await renderDocxArtifacts(tableBorderDocx(24,'ordinary',90,placement),{
      browser,outputs:['html'],reviewProfile:'final',commentProfile:'hidden',
    });
    await page.setContent(result.html!);
    const edge=await page.locator('.page-box').evaluate((page,rtl)=>{
      const rects=Array.from(page.querySelectorAll('svg[data-docx-table-border-paint] g rect'),r=>r.getBoundingClientRect());
      const box=page.getBoundingClientRect();
      return (rtl ? Math.min(...rects.map(r=>r.left))-box.left : Math.max(...rects.map(r=>r.right))-box.left)*.75;
    },placement==='rtl-end');
    expect(edge).toBeCloseTo(placement==='rtl-end'?72:540,1);
  });
}

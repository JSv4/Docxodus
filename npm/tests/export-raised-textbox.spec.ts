import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { raisedTextboxDocx } from './docx-raised-textbox-fixture.js';

test.use({launchOptions:{chromiumSandbox:true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? {executablePath:process.env.DOCXODUS_CHROMIUM_PATH} : {})}});

test('raised and lowered inline textboxes preserve text and the following paragraph position', async ({browser,page}) => {
  const labels: number[] = [], receipts: number[] = [];
  for (const position of [0,48,-48]) {
  const result = await renderDocxArtifacts(raisedTextboxDocx(position), {
    browser, outputs:['html','pdf'],reviewProfile:'final',commentProfile:'hidden', timeoutMs:120_000,
  });
  expect(result.pageCount).toBe(1);
  const task = getDocument({data:new Uint8Array(result.pdf!),verbosity:0});
  try {
    const content = await (await (await task.promise).getPage(1)).getTextContent();
    const items = content.items.filter(item=>'str' in item);
    const text = items.map(item=>item.str).join(' ');
    labels.push(items.find(item=>item.str.includes('Cobalt label'))!.transform[5]);
    receipts.push(items.find(item=>item.str.includes('Cobalt inventory'))!.transform[5]);
    for (const label of ['Cobalt label inside.','Cobalt inventory complete.']) expect(text.split(label)).toHaveLength(2);
  } finally { await task.destroy(); }
  await page.setContent(result.html!);
  const geometry = await page.locator('[data-docx-textbox="true"]').evaluate(box=>{
    const rect=box.getBoundingClientRect(), body=document.querySelector('.page-content')!.getBoundingClientRect();
    return {width:rect.width*.75,height:rect.height*.75,top:(rect.top-body.top)*.75};
  });
  expect(geometry.width).toBeCloseTo(180,1);
  expect(geometry.height).toBeCloseTo(36,1);
  if (position>0) expect(geometry.top).toBeLessThan(-20);
  }
  expect(labels[1]-labels[0]).toBeCloseTo(24,2);
  expect(labels[2]-labels[0]).toBeCloseTo(-24,2);
  expect(receipts[1]).toBeCloseTo(receipts[0],2);
  expect(receipts[2]).toBeCloseTo(receipts[0],2);
});

import { chromium } from '../../../npm/node_modules/@playwright/test/index.mjs';
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
const out = process.env.OUTPUT_DIR ?? '/tmp/docxodus-render-compare'; mkdirSync(out,{recursive:true});
const browser = await chromium.launch({headless:true});
try {
 const page = await browser.newPage({viewport:{width:1100,height:1050},deviceScaleFactor:Number(process.env.DPR ?? 2)});
 if (process.env.PROJECTION_PATH) await page.route('**/doom-ascii.js', r=>r.fulfill({body:readFileSync(process.env.PROJECTION_PATH),contentType:'text/javascript'}));
 await page.goto('http://localhost:8082/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0&cart=doom&render=ascii&wad=./vendor/doom1.wad.gz');
 await page.waitForFunction(()=>window.__arcade?.game().doomFrames>=2,null,{timeout:180000});
 await page.selectOption('#pace','0');
 const capture = async label => {
  await page.evaluate(()=>window.__arcade.pause());
  const raw = await page.evaluate(()=>{
   const a=window.__arcade,c=document.createElement('canvas');c.width=320;c.height=200;
   const ctx=c.getContext('2d'),d=ctx.createImageData(320,200);
   for(let y=0;y<200;y++)for(let x=0;x<320;x++) d.data.set([...a.game().pixel(x,y),255],(y*320+x)*4);
   ctx.putImageData(d,0,0);return c.toDataURL('image/png').split(',')[1];
  });
  writeFileSync(`${out}/${label}-source.png`,Buffer.from(raw,'base64'));
  for(const rendering of ['ascii','image']) {
   await page.selectOption('#rendering',rendering);
   await page.evaluate(()=>document.fonts.ready);
   const anchor=await page.evaluate(()=>window.__arcade.canvasElement().getAttribute('data-anchor'));
   await page.locator(`[data-anchor="${anchor}"]`).screenshot({path:`${out}/${label}-${rendering}.png`});
  }
  await page.selectOption('#rendering','ascii');
  const proof = await page.evaluate(async()=>{
   const a=window.__arcade;
   const {asciiFramebuffer}=await import('/doom-ascii.js');
   const {rowsFromXml}=await import('/ascii-arcade.js');
   const fb=new Uint8Array(320*200*4);
   for(let y=0;y<200;y++)for(let x=0;x<320;x++){
    const [r,g,b]=a.game().pixel(x,y); fb.set([b,g,r,255],(y*320+x)*4);
   }
   const expected=asciiFramebuffer(fb), xml=a.session.raw.getXml(a.canvasAnchor());
   const rows=rowsFromXml(xml),want=expected.chars.flat().join(''),el=a.canvasElement();
   const doc = new DOMParser().parseFromString('<root xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+xml+'</root>','application/xml');
   const expectedColors=expected.colors.flat(); let actualColors=[];
   for(const run of doc.getElementsByTagName('w:r')){
    const col=run.getElementsByTagName('w:color')[0]?.getAttribute('w:val');
    for(const t of run.getElementsByTagName('w:t'))actualColors.push(...Array(t.textContent.length).fill(col));
   }
   let htmlColors=[];
   for(const span of el.querySelectorAll('span')){
    if(span.querySelector('span'))continue;
    const rgb=getComputedStyle(span).color.match(/\d+/g).slice(0,3).map(Number);
    const hex=rgb.map(v=>v.toString(16).padStart(2,'0')).join('').toUpperCase();
    htmlColors.push(...Array(span.textContent.length).fill(hex));
   }
   return {frames:a.game().doomFrames,chars:el.textContent.length,xmlText:rows.join('')===want,
    domText:el.textContent===want, xmlColors:actualColors.length, htmlColors:htmlColors.length,
    wrongXmlColors:actualColors.filter((c,i)=>c!==expectedColors[i]).length,
    wrongHtmlColors:htmlColors.filter((c,i)=>c!==actualColors[i]).length,
    xml,html:el.outerHTML};
  });
  writeFileSync(`${out}/${label}-proof.json`,JSON.stringify(proof));
  const anchor=await page.evaluate(()=>window.__arcade.canvasElement().getAttribute('data-anchor'));
  await page.locator(`[data-anchor="${anchor}"]`).screenshot({path:`${out}/${label}-ascii-fresh.png`});
  const {xml,html,...summary}=proof; console.log(label,summary);
 };
 await capture('title');
 await page.click('#playpause');
 await page.keyboard.press('Enter'); await page.waitForTimeout(1000);
 await capture('menu');
 await page.click('#playpause');
 for(let i=0;i<3;i++){await page.keyboard.press('Enter');await page.waitForTimeout(800);}
 await page.waitForTimeout(1200);
 await capture('gameplay');
} finally {await browser.close();}

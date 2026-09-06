import { chromium } from '../../../npm/node_modules/@playwright/test/index.mjs';
import { writeFileSync } from 'node:fs';
const browser=await chromium.launch({headless:true});
try{
 const page=await browser.newPage({viewport:{width:1100,height:1050},deviceScaleFactor:Number(process.env.DPR ?? 2)});
 await page.goto('http://localhost:8082/demo-arcade.html?engine=./embed.bundle.js&intro=0&sound=0&cart=platformer');
 await page.waitForFunction(()=>window.__arcade?.frames()>0);
 await page.evaluate(()=>window.__arcade.pause());
 const measures=[];
 for(let c=32;c<127;c++){
  const anchor=await page.evaluate(async c=>{
   const a=window.__arcade;
   const {frameXml}=await import('/ascii-scenes.js');
   const {ASCII_METRICS}=await import('/doom-ascii.js');
   const chars=Array.from({length:200},()=>Array(321).fill(String.fromCharCode(c)));
   const colors=chars.map(row=>row.map(()=>'FFFFFF'));
   const xml=a.session.raw.getXml(a.canvasAnchor());
   const frame=frameXml(xml.slice(0,xml.indexOf('>')+1),{chars,colors},'000000',ASCII_METRICS);
   const result=a.session.raw.replaceXml(a.canvasAnchor(),frame.xml);if(!result.success)throw new Error(JSON.stringify(result));
   a.editor.refresh();await document.fonts.ready;
   return a.canvasElement().getAttribute('data-anchor');
  },c);
  const png=await page.locator(`[data-anchor="${anchor}"]`).screenshot();
  const mean=await page.evaluate(async b64=>{
   const img=new Image();img.src='data:image/png;base64,'+b64;await img.decode();
   const canvas=document.createElement('canvas');canvas.width=img.width;canvas.height=img.height;
   const ctx=canvas.getContext('2d');ctx.drawImage(img,0,0);
   const p=ctx.getImageData(Math.floor(img.width*.1),Math.floor(img.height*.2),Math.floor(img.width*.8),Math.floor(img.height*.5)).data;
   let sum=0;for(let i=0;i<p.length;i+=4)sum+=(p[i]+p[i+1]+p[i+2])/3;
   return sum/(p.length/4);
  },png.toString('base64'));
  measures.push(mean);if(c%10===0)console.log(c,mean);
 }
 const peak=Math.max(...measures);const values=measures.map(v=>Math.round(v/peak*1000));
 writeFileSync(process.env.COVERAGE_PATH ?? '/tmp/docxodus-glyph-coverage.json',JSON.stringify(values));
 console.log(JSON.stringify({peak,values}));
}finally{await browser.close();}

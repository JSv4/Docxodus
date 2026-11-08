import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { wordTextScript } from './word-text-script.js';

for (const owner of ['header','footer','footnotes'])
  test(`${owner} glyph baselines do not inherit body rounding`, async ({ page }) => {
    const observed = [];
    for (const count of [0,40]) {
      await page.setContent(`<style>body{margin:0}.page-box{position:relative;width:816px;height:1056px}
        p{font:16px/18.3984375px 'Liberation Serif';margin:0;width:624px;--docx-auto-line-spacing:1}
        p span{white-space:pre-wrap;line-height:18.3984375px}
        .page-header{position:absolute;top:48px;left:96px}.page-content{position:absolute;top:120px;left:96px}
        .page-footer{position:absolute;top:990px;left:96px}.page-footnotes{position:absolute;top:920px;left:96px}</style>
        <main class="page-box"><div class="page-header"></div><div class="page-content">
        ${Array.from({ length:count },(_,i) => `<p><span>Body${i} text.</span></p>`).join('')}</div>
        <div class="page-footer"></div><div class="page-footnotes"></div></main>
        <div id="prepare"><p><span style="--docx-character-scale:.8">Copper target.</span></p></div>`);
      await page.addScriptTag({ content:wordTextScript });
      observed.push(await page.evaluate(async owner => {
        await document.fonts.ready;
        const api = window as any, prep = document.getElementById('prepare')!, target = prep.querySelector('p')!;
        api.prepareWordText(prep,undefined,'running');
        document.querySelector('.page-' + owner)!.append(target);prep.remove();
        api.prepareWordText(document.querySelector('.page-content')!);
        const before = target.getBoundingClientRect().toJSON();
        api.paintWordText(document.querySelector('main')!);
        return { before,after:target.getBoundingClientRect().toJSON(),glyph:target.querySelector('text')!.getBoundingClientRect().top };
      },owner));
      if (count) {
        // The vector includes print-origin compensation. Read actual PDF
        // baselines to verify recovery still accumulates inside the body band.
        const task = getDocument({ data:new Uint8Array(await page.pdf({ width:'816px',height:'1056px' })),verbosity:0 });
        try {
          const pdf = await task.promise;
          const body = (await (await pdf.getPage(1)).getTextContent()).items.flatMap(i => 'str' in i && /^Body\d/.test(i.str) ? [i] : []);
          expect(body).toHaveLength(40);
          expect(Math.abs(body[0].transform[5] - body[39].transform[5] - 39 * 18.3984375 * .75)).toBeLessThan(.02);
        } finally { await task.destroy(); }
      }
    }
    expect(observed[1].before).toEqual(observed[0].before);
    for (const record of observed) expect(record.after).toEqual(record.before);
    expect(observed[1].glyph).toBeCloseTo(observed[0].glyph,4);
  });

const noteWords = Array.from({ length:24 }, (_, i) => `N${String(i).padStart(4,'0')}`).join(' ');
const noteMarkup = (text = noteWords, scale = 1.5) => `<style>
  p{font:16px/18px 'Liberation Mono';margin:0;--docx-auto-line-spacing:1}
  p span{white-space:pre-wrap;line-height:18px;--docx-character-scale:${scale}}</style>
  <aside id="registry"><div data-footnote-id="1"><p data-source-anchor-id="p:fn:one" style="line-height:18px!important"><span style="line-height:18px">${text}</span></p></div></aside>
  <main class="page-footnotes" style="width:216pt"></main>`;

test('carried note lines refit 216→144→216pt without losing fragment styles or text', async ({ page }) => {
  await page.setContent(noteMarkup());await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const api = window as any, registry = document.getElementById('registry')!, band = document.querySelector<HTMLElement>('main')!;
    let checkpoints = 0;
    const prepare = api.createFootnoteTextPreparer(registry, () => checkpoints++);
    band.append(registry.firstElementChild!.cloneNode(true));
    const first = prepare(band), p = band.querySelector<HTMLElement>('p')!;
    const firstLines = Array.from(p.querySelectorAll('[data-docx-word-line]')).map(line => line.textContent);
    // The continuation owns these properties; restoration must not replace its paragraph style.
    p.style.setProperty('margin-top','0','important');p.style.setProperty('text-indent','0','important');
    p.style.textAlignLast = 'justify';
    const before = p.outerHTML;
    checkpoints = 0;
    const run = p.firstElementChild!;
    Object.defineProperty(run,'textContent',{ configurable:true,get() { throw new Error('Same-width note rescanned its fitted text'); } });
    const same = prepare(p), unchanged = p.outerHTML === before, sameCheckpoints = checkpoints;
    Reflect.deleteProperty(run,'textContent');
    const clone = p.cloneNode(true) as HTMLElement;p.replaceWith(clone);
    band.style.width = '144pt';const narrow = prepare(clone);
    const narrowLines = Array.from(clone.querySelectorAll('[data-docx-word-line]')).map(line => line.textContent);
    band.style.width = '216pt';const wide = prepare(clone);
    return { first,same,unchanged,sameCheckpoints,narrow,wide,firstLines,narrowLines,
      restoredLines:Array.from(clone.querySelectorAll('[data-docx-word-line]')).map(line => line.textContent),
      text:clone.textContent,margin:clone.style.marginTop,indent:clone.style.textIndent,last:clone.style.textAlignLast,
      marginPriority:clone.style.getPropertyPriority('margin-top') };
  });
  expect(result.first && result.same && result.unchanged && result.narrow && result.wide).toBe(true);
  expect(result.sameCheckpoints).toBeLessThanOrEqual(2);
  expect(result.narrowLines.length).toBeGreaterThan(result.firstLines.length);
  expect(result.restoredLines).toEqual(result.firstLines);expect(result.text).toBe(noteWords);
  expect([result.margin,result.indent,result.last,result.marginPriority]).toEqual(['0px','0px','justify','important']);
});

test('an owned scaled word that cannot refit refuses instead of becoming unscaled', async ({ page }) => {
  // Ten glyphs fit 80pt at 100%, but exceed it at 150%.
  await page.setContent(noteMarkup('ABCDEFGHIJ'));await page.addScriptTag({ content:wordTextScript });
  const result = await page.evaluate(() => {
    const api = window as any, registry = document.getElementById('registry')!, band = document.querySelector<HTMLElement>('main')!;
    const prepare = api.createFootnoteTextPreparer(registry);
    band.append(registry.firstElementChild!.cloneNode(true));const initial = prepare(band);
    band.style.width = '80pt';const fitted = band.cloneNode(true), refit = prepare(band), rollback = band.isEqualNode(fitted);
    const untouched = registry.firstElementChild!.cloneNode(true) as HTMLElement;band.replaceChildren(untouched);
    const before = untouched.outerHTML, unsupported = prepare(band);
    return { initial,refit,rollback,unsupported,initialFallbackUnchanged:untouched.outerHTML === before };
  });
  expect(result).toEqual({ initial:true,refit:false,rollback:true,unsupported:true,initialFallbackUnchanged:true });
});

for (const source of ['inline','inline-block','reference','link','mixed','missing identity','duplicate identity'])
  test(`splitting an unsupported ${source} note does not admit its plain continuation`, async ({ page }) => {
    await page.setContent(noteMarkup());await page.addScriptTag({ content:wordTextScript });
    const result = await page.evaluate(source => {
      const registry = document.getElementById('registry')!, original = registry.querySelector<HTMLElement>('p')!;
      if (source.startsWith('inline')) original.style.display = source;
      if (source === 'reference') original.insertAdjacentHTML('afterbegin','<sup>1</sup>');
      if (source === 'link') original.insertAdjacentHTML('afterbegin','<a href="#target">Link</a>');
      if (source === 'mixed') original.insertAdjacentHTML('afterbegin','<span style="font-weight:bold">Bold</span>');
      if (source === 'missing identity') original.removeAttribute('data-source-anchor-id');
      if (source === 'duplicate identity') original.after(original.cloneNode(true));
      const prepare = (window as any).createFootnoteTextPreparer(registry), band = document.querySelector('main')!;
      band.append(registry.firstElementChild!.cloneNode(true));
      const p = band.querySelector<HTMLElement>('p')!;
      // A continuation starts a block and no longer contains the leading reference/formatting.
      p.style.display = 'block';while (p.children.length > 1) p.firstElementChild!.remove();
      const before = band.innerHTML;
      return { success:prepare(band),unchanged:band.innerHTML === before };
    },source);
    expect(result).toEqual({ success:true,unchanged:true });
  });

for (const metadata of [null,'{','null','{"width":192}'])
  test(`owned note with invalid fit metadata ${metadata === null ? 'missing' : JSON.stringify(metadata)} refuses without mutation`, async ({ page }) => {
    await page.setContent(noteMarkup());await page.addScriptTag({ content:wordTextScript });
    const result = await page.evaluate(metadata => {
      const registry = document.getElementById('registry')!, band = document.querySelector('main')!;
      const prepare = (window as any).createFootnoteTextPreparer(registry);
      band.append(registry.firstElementChild!.cloneNode(true));const initial = prepare(band), p = band.querySelector('p')!;
      if (metadata === null) p.removeAttribute('data-docx-note-fit');else p.setAttribute('data-docx-note-fit',metadata);
      const before = p.outerHTML;
      return { initial,refit:prepare(band),unchanged:p.outerHTML === before };
    },metadata);
    expect(result).toEqual({ initial:true,refit:false,unchanged:true });
  });

for (const scale of [.8,1.5]) test(`a body picture leaves ${scale * 100}% justified note word positions unchanged`, async ({ page }) => {
  const positions = [];
  for (const picture of [false,true]) {
    await page.setContent(noteMarkup(noteWords,scale) + (picture ? '<img width="20" height="20">' : ''));
    await page.addScriptTag({ content:wordTextScript });
    positions.push(await page.evaluate(async () => {
      await document.fonts.ready;document.body.dataset.docxCompatibilityMode = '15';
      const registry = document.getElementById('registry')!, band = document.querySelector('main')!;
      registry.querySelector<HTMLElement>('p')!.style.textAlign = 'justify';
      const prepare = (window as any).createFootnoteTextPreparer(registry);
      band.append(registry.firstElementChild!.cloneNode(true));prepare(band);
      (window as any).paintWordText(document.body);
      return Array.from(band.querySelectorAll('tspan')).map(t => [t.textContent,Number(t.getAttribute('x')),Number(t.parentElement!.getAttribute('y'))]);
    }));
  }
  expect(positions[0]).toHaveLength(24);expect(positions[1]).toEqual(positions[0]);
});

for (const graphic of ['<img width="12" height="12">','<div data-docx-text-frame="true">Frame</div>'])
  test(`graphic registry retains all note paragraphs (${graphic})`, async ({ page }) => {
    await page.setContent(noteMarkup() + `<template id="graphic">${graphic}</template>`);
    await page.addScriptTag({ content:wordTextScript });
    const result = await page.evaluate(() => {
      const api = window as any, registry = document.getElementById('registry')!, band = document.querySelector('main')!;
      const separator = document.createElement('div');separator.dataset.footnoteSeparator = 'continuation';
      separator.append((document.getElementById('graphic') as HTMLTemplateElement).content.cloneNode(true));registry.append(separator);
      const prepare = api.createFootnoteTextPreparer(registry);
      band.append(registry.firstElementChild!.cloneNode(true));const before = band.innerHTML;
      return { success:prepare(band),unchanged:band.innerHTML === before };
    });
    expect(result).toEqual({ success:true,unchanged:true });
  });

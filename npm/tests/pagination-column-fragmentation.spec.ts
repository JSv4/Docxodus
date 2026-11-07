import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
const geometry = `data-page-width="340" data-page-height="200" data-content-width="300" data-content-height="160"
  data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20"`;
const labels = Array.from({ length: 47 }, (_, i) => `Birch${String(i + 1).padStart(3, '0')}`);

for (const sequential of [false, true]) for (const group of [false, true]) {
  test(`column fragments retain identities and spacing below a title: sequential=${sequential}, border=${group}`, async ({ page }) => {
    const paragraph = `<p id="source" data-anchor="editor:source" data-source-anchor-id="p:body:source"
      data-widow-control="false" style="margin-top:6pt;text-indent:9pt"><strong id="run" data-anchor="editor:run">${labels.join('<br>')}</strong></p>`;
    await page.setContent(`<style>#staging,#container{font:11pt/16pt 'Liberation Serif'}p{margin:0}
      .border{border:1px solid black;padding:0}</style><div id="staging">
      <section data-section-index="0" ${geometry}><p style="height:48pt">Title</p></section>
      <section data-section-index="1" ${geometry} data-section-type="continuous" data-cols="2" data-col-gap="20"
        style="column-count:2;column-gap:20pt">${group ? `<div class="border">${paragraph}</div>` : paragraph}
        <p data-source-anchor-id="p:body:end">BirchEnd</p></section>
      ${sequential ? '' : `<section data-section-index="2" ${geometry} data-section-type="continuous"><p>After columns</p></section>`}
      </div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('container')!;
      const engine = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), container,
        { fragmentParagraphs: true, showPageNumbers: false });
      engine.paginate();
      const fragments = Array.from(container.querySelectorAll<HTMLElement>('[data-source-anchor-id="p:body:source"]'));
      const overflow: string[] = [];
      for (const body of Array.from(container.querySelectorAll<HTMLElement>('.page-content'))) {
        const band = body.getBoundingClientRect();
        const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT);
        for (let node = walker.nextNode(); node; node = walker.nextNode()) {
          if (!node.textContent?.trim()) continue;
          const range = document.createRange(); range.selectNodeContents(node);
          if (Array.from(range.getClientRects()).some(r => r.top < band.top - 1 || r.bottom > band.bottom + 1
            || r.left < band.left - 1 || r.right > band.right + 1)) overflow.push(node.textContent);
        }
      }
      return {
        text: container.textContent,
        pages: Array.from(container.querySelectorAll<HTMLElement>('.page-content'), body => body.innerText),
        ids: ['source', 'run'].map(id => container.querySelectorAll(`[id="${id}"]`).length),
        anchors: ['editor:source', 'editor:run'].map(id => container.querySelectorAll(`[data-anchor="${id}"]`).length),
        spacing: fragments.map(p => ({ indent: parseFloat(getComputedStyle(p).textIndent), top: parseFloat(getComputedStyle(p).marginTop) })),
        overflow, map: engine.materializePageMap(0, 'column-fragmentation'),
      };
    });
    expect(result.text?.match(/Birch\d+/g)).toEqual(labels);
    expect(result.ids).toEqual([1, 1]);
    expect(result.anchors).toEqual([1, 1]);
    expect(result.pages[0]).toContain('Birch001');
    expect(result.pages.find(text => text.includes('BirchEnd'))).toContain('Birch047');
    expect(result.spacing.length).toBeGreaterThan(1);
    expect(result.spacing[0].indent).toBeCloseTo(12, 1);
    expect(result.spacing.slice(1).every(p => p.indent === 0 && p.top === 0)).toBe(true);
    expect(result.overflow).toEqual([]);
    const mapped = result.map.fragments.filter((f: any) => f.anchorId === 'p:body:source');
    expect(mapped.length).toBeGreaterThan(1);
    expect(mapped.map((f: any) => f.fragmentIndex)).toEqual(mapped.map((_: any, i: number) => i));
  });
}

for (const restriction of ['editor', 'column-break', 'fixed-height', 'keep-lines']) {
  test(`unsupported column paragraph retains its whole-block fallback: ${restriction}`, async ({ page }) => {
    await page.setContent(`<style>#staging,#container{font:11pt/16pt 'Liberation Serif'}p{margin:0}</style>
      <div id="staging"><section data-section-index="0" ${geometry} data-cols="2" data-col-gap="20" style="column-count:2;column-gap:20pt">
      <p data-source-anchor-id="p:body:source" ${restriction === 'column-break' ? 'data-column-break-before="true" style="break-before:column"' : restriction === 'fixed-height' ? 'style="height:752pt"' : restriction === 'keep-lines' ? 'data-keep-lines="true" style="break-inside:avoid"' : ''}>${labels.join('<br>')}</p>
      ${Array.from({ length:30 }, (_, i) => `<p>End${i + 1}</p>`).join('')}
      </section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(restriction => {
      const container = document.getElementById('container')!;
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), container,
        { fragmentParagraphs: restriction !== 'editor', showPageNumbers: false }).paginate();
      return { text: container.textContent, count: container.querySelectorAll('[data-source-anchor-id="p:body:source"]').length,
        endOverflow: Array.from(container.querySelectorAll<HTMLElement>('p')).filter(p => /^End/.test(p.textContent ?? '')).map(p => {
          const box = p.getBoundingClientRect(), body = p.closest('.page-content')!.getBoundingClientRect();
          return box.bottom > body.bottom + 1 || box.right > body.right + 1;
        }) };
    }, restriction);
    expect(result.text?.match(/Birch\d+/g)).toEqual(labels);
    expect(result.count).toBe(1);
    expect(result.text?.match(/End\d+/g)).toEqual(Array.from({ length:30 }, (_, i) => `End${i + 1}`));
    expect(result.endOverflow).toEqual(Array(30).fill(false));
  });
}


test('successive column tails keep filling the current band', async ({ page }) => {
  const groups = Array.from({ length: 3 }, (_, i) => Array.from({ length: 21 }, (_, j) => `Fir${i * 22 + j + 1}`));
  await page.setContent(`<style>#staging,#container{font:11pt/16pt 'Liberation Serif'}p{margin:0}</style>
    <div id="staging"><section data-section-index="0" ${geometry} data-section-type="continuous" data-cols="2" data-col-gap="20" style="column-count:2;column-gap:20pt">
    ${groups.map((lines, i) => `<p data-widow-control="false">${lines.join('<br>')}</p><p>Fir${(i + 1) * 22}</p>`).join('')}
    </section></div><div id="container"></div>`);
  await page.addScriptTag({ path: bundle });
  const pages = await page.evaluate(() => {
    const container = document.getElementById('container')!;
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), container,
      { fragmentParagraphs: true, showPageNumbers: false }).paginate();
    return Array.from(container.querySelectorAll('.page-content'), body => body.textContent?.match(/Fir\d+/g) ?? []);
  });
  expect(pages.map(labels => labels.length)).toEqual([20, 20, 20, 6]);
  expect(pages.flat()).toEqual(Array.from({ length:66 }, (_, i) => `Fir${i + 1}`));
});

test('ordinary column streams keep cloning work proportional to their length', async ({ page }) => {
  await page.setContent(`<style>#staging,#container{font:11pt/16pt 'Liberation Serif'}p{margin:0}</style><div id="staging"></div><div id="container"></div>`);
  await page.addScriptTag({ path: bundle });
  const samples = await page.evaluate(geometry => {
    const clone = Node.prototype.cloneNode;
    let clonedCharacters = 0;
    Node.prototype.cloneNode = function (deep?: boolean) {
      if (deep) clonedCharacters += this.textContent?.length ?? 0;
      return clone.call(this, deep);
    };
    try {
      return [300, 900].map(count => {
        const staging = document.getElementById('staging')!, container = document.getElementById('container')!;
        staging.innerHTML = `<section data-section-index="0" ${geometry} data-cols="2" data-col-gap="20" style="column-count:2;column-gap:20pt">
          ${'<p>Row</p>'.repeat(count)}</section>`;
        container.replaceChildren(); clonedCharacters = 0;
        new (window as any).DocxodusPagination.PaginationEngine(staging, container,
          { fragmentParagraphs: true, showPageNumbers: false }).paginate();
        return { work: clonedCharacters, count: container.querySelectorAll('p').length };
      });
    } finally { Node.prototype.cloneNode = clone; }
  }, geometry);
  expect(samples.map(sample => sample.count)).toEqual([300, 900]);
  expect(samples[1].work).toBeLessThan(samples[0].work * 4);
});


for (const oneParagraph of [false, true]) {
  test(`joining a column chunk retains the next paragraph's space-before: oneParagraph=${oneParagraph}`, async ({ page }) => {
    const labels = Array.from({ length:21 }, (_, i) => `Ash${i + 1}`);
    await page.setContent(`<style>#staging,#container{font:11pt/16pt 'Liberation Serif'}p{margin:0;orphans:1;widows:1}</style>
      <div id="staging"><section data-section-index="0" ${geometry} data-section-type="continuous" data-cols="2" data-col-gap="20" style="column-count:2;column-gap:20pt">
      ${oneParagraph ? `<p data-widow-control="false">${labels.join('<br>')}</p>` : labels.map(t => `<p>${t}</p>`).join('')}
      <p id="receipt" style="margin-top:32pt">AshReceipt</p></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      const container = document.getElementById('container')!;
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), container,
        { fragmentParagraphs:true, showPageNumbers:false }).paginate();
      const receipt = container.querySelector('#receipt')!, body = receipt.closest('.page-content')!;
      return { pages:container.querySelectorAll('.page-content').length,
        top:(receipt.getBoundingClientRect().top - body.getBoundingClientRect().top) * .75,
        text:body.textContent };
    });
    expect(result.pages).toBe(2);
    expect(result.text).toContain('Ash21');
    expect(result.top).toBeCloseTo(48, 1);
  });
}

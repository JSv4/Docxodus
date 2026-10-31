import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

const wrappedFrame = (width: number, height = 60) => `<div data-docx-frame-container="true" style="position:relative;height:0">
      <div data-docx-frame-exclusion="true" style="float:left;display:grid"><div data-docx-text-frame="true"
      data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline" data-docx-frame-wrap="around"
      style="position:relative;display:flow-root;width:${width}pt;height:${height}pt;grid-area:1/1;align-self:start"><p>Box</p></div></div></div>`;

for (const story of ['body', 'header', 'footer']) for (const { widths, between } of [
  { widths: [60, 60], between: 'paragraph' }, { widths: [30, 90], between: 'paragraph' },
  { widths: [90, 30], between: 'paragraph' }, { widths: [30, 90], between: 'table' },
  { widths: [30, 90], between: 'flow-root' }, { widths: [30, 90], between: 'margin' },
  { widths: [90, 30], between: 'gap' },
]) {
  test(`${story} separated ${widths.join('/')}pt frames preserve the intervening ${between}`, async ({ page }) => {
    const block = between === 'table' ? '<table data-row style="font:inherit;width:80pt;border-collapse:collapse"><tr><td style="padding:0;height:12pt">One</td></tr></table>'
      : between === 'flow-root' ? '<div data-row style="display:flow-root;width:80pt;height:12pt">One</div>'
        : `<p data-row style="margin-bottom:${between === 'margin' ? 6 : 0}pt">One${between === 'gap' ? '<br>Gap<br>Gap' : ''}</p>`;
    const secondY = between === 'margin' ? 18 : between === 'gap' ? 36 : 12;
    const content = `${wrappedFrame(widths[0], between === 'gap' ? 12 : 60)}${block}${wrappedFrame(widths[1])}<p data-row>Two</p><p data-row>Tri</p>`;
    for (const scale of [1, .75, 1.25]) {
      await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
        ${story === 'body' ? '' : `<div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="${story}-default">${content}</div></div>`}
        <section data-section-index="0" data-page-width="168" data-page-height="240"
        data-content-width="120" data-content-height="144" data-header-height="5" data-footer-height="5"
        data-margin-top="48" data-margin-bottom="48" data-margin-left="24" data-margin-right="24">
        ${story === 'body' ? content : '<p>Body</p>'}</section></div><div id="output"></div>`);
      await page.addScriptTag({ path: bundle });
      const pages = await page.evaluate(({ scale, story }) => {
        new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
          document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true, scale }).paginate();
        return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
          const band = box.querySelector<HTMLElement>(story === 'body' ? '.page-content' : `.page-${story}`)!;
          const origin = band.getBoundingClientRect(), units = box.getBoundingClientRect().width / 168;
          const rect = (node: Element | Range) => {
            const r = node.getBoundingClientRect(); return { x: (r.left - origin.left) / units, y: (r.top - origin.top) / units };
          };
          return { frames: Array.from(box.querySelectorAll('[data-docx-text-frame]')).map(rect),
            rows: Array.from(band.querySelectorAll('[data-row]')).map(p => {
              const range = document.createRange(); range.selectNodeContents(p);
              const line = range.getClientRects()[0];
              return { x: (line.left - origin.left) / units, blockY: rect(p).y };
            }) };
        });
      }, { scale, story });
      expect(pages).toHaveLength(1);
      expect(pages[0].frames).toHaveLength(2);
      for (const [i, frame] of pages[0].frames.entries()) {
        expect(frame.x).toBeCloseTo(0, 1);
        expect(frame.y).toBeCloseTo(i === 0 ? 0 : secondY, 1);
      }
      expect(pages[0].rows).toHaveLength(3);
      for (const [i, row] of pages[0].rows.entries()) {
        expect(row.blockY).toBeCloseTo(i === 0 ? 0 : secondY + (i - 1) * 12, 1);
        expect(row.x).toBeCloseTo(i === 0 ? widths[0] : between === 'gap' ? widths[1] : Math.max(...widths), 1);
      }
    }
  });
}

for (const story of ['header', 'body']) for (const wide of [false, true]) {
  test(`${story} adjacent frames retain the ${wide ? 'stepped' : 'narrow control'} exclusion`, async ({ page }) => {
    const frame = (width: number, height: number) => `<div data-docx-text-frame="true"
      data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
      style="position:relative;display:flow-root;width:${width}pt;height:${height}pt;grid-area:1/1;align-self:start"><p>Box</p></div>`;
    const frames = `<div data-docx-frame-container="true" style="position:relative;height:0">
      <div data-docx-frame-exclusion="true" style="float:left;display:grid">${wide ? frame(90, 12) : ''}${frame(30, 72)}</div></div>`;
    for (const scale of [1, .75]) {
      await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
        ${story === 'header' ? `<div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="header-default">${frames}<p>Head</p></div></div>` : ''}
        <section data-section-index="0" data-page-width="168" data-page-height="144"
          data-content-width="120" data-content-height="96" data-header-height="5"
          data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
          ${story === 'body' ? frames : ''}<p>Row0</p><p>Row1</p><p>Row2</p>
        </section></div><div id="output"></div>`);
      await page.addScriptTag({ path: bundle });
      const pages = await page.evaluate(scale => {
        new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
          document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true, scale }).paginate();
        return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
          const body = box.querySelector<HTMLElement>('.page-content')!, left = body.getBoundingClientRect().left;
          const units = box.getBoundingClientRect().width / 168;
          return Array.from(body.querySelectorAll('p')).map(p => {
            const range = document.createRange(); range.selectNodeContents(p);
            return { text: p.textContent, x: (range.getBoundingClientRect().left - left) / units };
          });
        });
      }, scale);
      expect(pages).toHaveLength(1);
      expect(pages[0].map(p => p.text)).toEqual(['Row0', 'Row1', 'Row2']);
      for (const [i, p] of pages[0].entries()) expect(p.x).toBeCloseTo(story === 'body' && wide && i === 0 ? 90 : 30, 1);
    }
  });
}

test('body frames share a header exclusion without charging its overhang to body height', async ({ page }) => {
  for (const scale of [1, .75, 1.25]) {
    await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
      <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="header-default">
      ${wrappedFrame(60, 120)}<p>Head</p></div></div>
      <section data-section-index="0" data-page-width="168" data-page-height="144"
      data-content-width="120" data-content-height="96" data-header-height="5"
      data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
      ${wrappedFrame(30, 24)}<p>R0</p>${wrappedFrame(90, 24)}${Array.from({ length: 7 }, (_, i) => `<p>R${i + 1}</p>`).join('')}
      </section></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(scale => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true, scale }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
        const body = box.querySelector<HTMLElement>('.page-content')!.getBoundingClientRect();
        const units = box.getBoundingClientRect().width / 168;
        return Array.from(box.querySelectorAll('.page-content p')).map(p => {
          const range = document.createRange(); range.selectNodeContents(p);
          const rect = range.getBoundingClientRect();
          return { text: p.textContent, x: (rect.left - body.left) / units };
        });
      });
    }, scale);
    expect(pages).toHaveLength(1);
    expect(pages[0].map(row => row.text)).toEqual(Array.from({ length: 8 }, (_, i) => `R${i}`));
    for (const [i, row] of pages[0].entries()) expect(row.x).toBeCloseTo(i === 1 || i === 2 ? 90 : 60, 1);
  }
});

for (const kind of ['table', 'flow-root']) for (const framed of [false, true]) {
  test(`a full-width ${kind} accounts for ${framed ? 'header-frame clearance' : 'ordinary header control'} when paginating`, async ({ page }) => {
    const content = kind === 'table'
      ? `<table style="font:inherit;width:120pt;border-collapse:collapse"><tbody>${[0, 1, 2].map(i =>
        `<tr><td style="height:16pt;padding:0"><p data-row>Row${i}</p></td></tr>`).join('')}</tbody></table>`
      : '<div data-row style="display:flow-root;width:120pt;height:48pt">Block</div>';
    await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
      <div id="pagination-hf-registry" style="display:none">
      <div data-section="0" data-hf-type="header-first">${framed ? wrappedFrame(60, 72) : ''}<p>Head</p></div>
      <div data-section="0" data-hf-type="header-default"><p>Head</p></div></div>
      <section data-section-index="0" data-page-width="168" data-page-height="144"
      data-content-width="120" data-content-height="96" data-header-height="5"
      data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
      <p style="margin-bottom:12pt">Intro</p>${content}</section></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-content')).map(body => {
        const bounds = body.getBoundingClientRect();
        return Array.from(body.querySelectorAll<HTMLElement>('[data-row]')).map(row => ({
          text: row.textContent, outside: row.getBoundingClientRect().bottom - bounds.bottom,
        }));
      });
    });
    expect(pages).toHaveLength(framed ? 2 : 1);
    expect(pages.flat().map(row => row.text)).toEqual(kind === 'table' ? ['Row0', 'Row1', 'Row2'] : ['Block']);
    for (const row of pages.flat()) expect(row.outside).toBeLessThanOrEqual(.5);
  });
}

for (const height of [114, 120]) {
  test(`a ${height}pt header frame does not evict body text that fits beside it`, async ({ page }) => {
    await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
      <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="header-default">
        <div data-docx-frame-container="true" style="position:relative;height:0">
          <div data-docx-frame-exclusion="true" style="float:left"><div data-docx-text-frame="true"
          data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
          style="position:relative;display:flow-root;width:60pt;height:${height}pt"><p>Frame</p></div></div>
        </div><p>Head</p></div></div>
      <section data-section-index="0" data-page-width="168" data-page-height="144"
        data-content-width="120" data-content-height="96" data-header-height="5"
        data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
        <p>Row0</p><p>Row1</p><p>Row2</p>
      </section></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
        const body = box.querySelector<HTMLElement>('.page-content')!, left = body.getBoundingClientRect().left;
        return Array.from(body.querySelectorAll('p')).map(p => {
          const range = document.createRange(); range.selectNodeContents(p);
          return { text: p.textContent, x: (range.getBoundingClientRect().left - left) * .75 };
        });
      });
    });
    expect(pages).toHaveLength(1);
    expect(pages[0].map(p => p.text)).toEqual(['Row0', 'Row1', 'Row2']);
    for (const paragraph of pages[0]) expect(paragraph.x).toBeCloseTo(60, 1);
  });
}

for (const framed of [false, true]) {
  test(`a first-page header ${framed ? 'frame' : 'control'} uses the same wrapping when splitting body text`, async ({ page }) => {
    const words = Array.from({ length: 32 }, (_, i) => `W${String(i).padStart(3, '0')}`);
    const frame = framed ? `<div data-docx-frame-container="true" style="position:relative;height:0">
      <div data-docx-frame-exclusion="true" style="float:left"><div data-docx-text-frame="true"
      data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
      style="position:relative;display:flow-root;width:60pt;height:72pt;left:0;top:0"><p>Frame</p></div></div></div>` : '';
    await page.setContent(`<style>#staging,#output{font:10pt/12pt 'Courier New'}p{margin:0}</style><div id="staging">
      <div id="pagination-hf-registry" style="display:none">
        <div data-section="0" data-hf-type="header-first">${frame}<p>Head</p></div>
        <div data-section="0" data-hf-type="header-default"><p>Head</p></div>
      </div><section data-section-index="0" data-page-width="168" data-page-height="144"
        data-content-width="120" data-content-height="96" data-header-height="5"
        data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24">
        <p data-anchor="p:body:sage" data-source-anchor-id="p:body:sage" data-word-paragraph="true">${words.join(' ')}</p>
      </section></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
        const body = box.querySelector<HTMLElement>('.page-content')!, rect = body.getBoundingClientRect();
        const p = body.querySelector<HTMLElement>('p')!, range = document.createRange();
        range.selectNodeContents(p);
        const lines = Array.from(range.getClientRects());
        return { text: body.textContent!.trim(), firstX: (lines[0].left - rect.left) * .75,
          bottom: (lines[lines.length - 1].bottom - rect.bottom) * .75,
          frames: box.querySelectorAll('[data-docx-text-frame]').length,
          anchor: p.dataset.sourceAnchorId ?? p.dataset.anchor };
      });
    });
    expect(pages).toHaveLength(framed ? 2 : 1);
    expect(pages.flatMap(p => p.text.split(/\s+/))).toEqual(words);
    expect(pages[0].firstX).toBeCloseTo(framed ? 60 : 0, 1);
    expect(pages.every(p => p.bottom < 1)).toBe(true);
    expect(pages.every(p => p.anchor === 'p:body:sage')).toBe(true);
    if (framed) {
      expect(pages.map(p => p.frames)).toEqual([1, 0]);
      expect(pages[1].firstX).toBeCloseTo(0, 1);
    }
  });
}

for (const story of ['header', 'footer'] as const) for (const framed of [false, true]) {
  test(`${story} running style remeasures a changed ${framed ? 'wrapped' : 'ordinary'} story`, async ({ page }) => {
    const title = 'Sage calibration record with several descriptive words for the northern research station';
    const frame = framed ? `<div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0">
      <div data-docx-frame-exclusion="true" style="float:left"><div data-docx-text-frame="true"
      data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
      style="position:relative;display:flow-root;width:30pt;height:12pt;left:0;top:0"><p>Frame</p></div></div></div>` : '';
    await page.setContent(`<style>#staging,#output{font:12pt/12pt 'Liberation Serif'}p{margin:0}</style><div id="staging">
      <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="${story}-default">${frame}
        <p data-running="true"><span data-field="STYLEREF" data-field-style="Chapter">ALDER</span></p></div></div>
      <div data-section-index="0" data-page-width="228" data-page-height="240" data-content-width="180" data-content-height="192"
        data-margin-top="24" data-margin-bottom="24" data-margin-left="24" data-margin-right="24" data-header-height="5" data-footer-height="5">
        <p data-anchor="a" data-paragraph-style="Chapter" data-styleref-text="ALDER">ALDER</p>
        <p data-anchor="b" data-paragraph-style="Chapter" data-styleref-text="${title}" data-page-break-before="true">${title}</p><p>Body end</p>
      </div></div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(story => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
        const band = box.querySelector<HTMLElement>(`.page-${story}`)!.getBoundingClientRect();
        const running = box.querySelector<HTMLElement>('[data-running]')!;
        const text = running.getBoundingClientRect(), body = box.querySelector<HTMLElement>('.page-content')!.getBoundingClientRect();
        return { text: running.textContent, textHeight: text.height, bandHeight: band.height,
          clearance: story === 'header' ? body.top - text.bottom : text.top - body.bottom };
      });
    }, story);
    expect(pages).toHaveLength(2);
    expect(pages[1].text).toBe(title);
    expect(pages[1].textHeight).toBeGreaterThan(pages[0].textHeight);
    expect(pages[1].bandHeight).toBeGreaterThanOrEqual(pages[1].textHeight - 0.5);
    expect(pages[1].clearance).toBeGreaterThanOrEqual(-0.5);
  });
}

for (const framed of [false, true]) {
  test(`a first-line note stays with its fragmented ${framed ? 'wrapped' : 'ordinary'} paragraph`, async ({ page }) => {
    const frame = framed ? `<div data-docx-frame-container="true" data-keep-with-next="true" style="position:relative;height:0">
      <div data-docx-frame-exclusion="true" style="float:left"><div data-docx-text-frame="true"
      data-docx-frame-h-anchor="text" data-docx-frame-v-anchor="text" data-docx-frame-yalign="inline"
      style="position:relative;display:flow-root;width:60pt;height:24pt;left:0;top:0"><p>Frame</p></div></div></div>` : '';
    const lines = Array.from({ length: 12 }, (_, i) => `ROW${i}${i === 0 ? '<sup data-footnote-id="f1">1</sup>' : ''}`).join('<br>');
    await page.setContent(`<style>#staging,#output{font:10pt/12pt Arial}p{margin:0}.footnote-content p{height:12pt}</style>
      <div id="staging"><div id="pagination-footnote-registry" style="display:none"><div class="footnote-item" data-footnote-id="f1">
      <span class="footnote-number">1</span><div class="footnote-content"><p>NOTE0</p></div></div></div>
      <section data-section-index="0" data-page-width="240" data-page-height="240" data-content-width="180" data-content-height="180"
      data-margin-top="30" data-margin-bottom="30" data-margin-left="30" data-margin-right="30">${frame}<p>${lines}</p><p>END</p></section>
      </div><div id="output"></div>`);
    await page.addScriptTag({ path: bundle });
    const pages = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'),
        document.getElementById('output'), { showPageNumbers: false, fragmentParagraphs: true }).paginate();
      return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
        const frame = box.querySelector<HTMLElement>('[data-docx-text-frame]')?.getBoundingClientRect();
        const note = box.querySelector<HTMLElement>('.page-footnotes');
        return { body: box.querySelector('.page-content')?.textContent, note: note?.textContent,
          collision: !!frame && !!note && frame.bottom > note.getBoundingClientRect().top + 0.5 };
      });
    });
    expect(pages).toHaveLength(2);
    expect(pages[0].note).toContain('NOTE0');
    expect(pages[0].body).toContain('ROW10');
    expect(pages[0].body).not.toContain('ROW11');
    expect(pages[0].collision).toBe(false);
    expect(pages[1].body).toContain('ROW11');
    expect(pages[1].body).toContain('END');
  });
}

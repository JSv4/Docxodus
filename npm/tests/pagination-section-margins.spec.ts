import { test, expect, type Page } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));
type Margins = Partial<Record<'top' | 'right' | 'bottom' | 'left', number>>;

function section(index: number, content: string, changed: Margins = {}, type = 'continuous', width = 300): string {
  const m = { top: 30, right: 30, bottom: 30, left: 30, ...changed };
  return `<section data-section-index="${index}" data-section-type="${type}"
    data-page-width="${width}" data-page-height="240"
    data-content-width="${width - m.left - m.right}" data-content-height="${240 - m.top - m.bottom}"
    ${Object.entries(m).map(([side, value]) => `data-margin-${side}="${value}"`).join(' ')}>${content}</section>`;
}

const paragraph = (id: string, height = 20, style = '') =>
  `<p data-marker="${id}" style="height:${height}pt;${style}">${id}</p>`;

async function layout(page: Page, sections: string) {
  await page.setContent(`<style>#staging,#output{font:10pt/12pt Arial}p{margin:0;padding:0}</style>
    <div id="staging">${sections}</div><div id="output"></div>`);
  await page.addScriptTag({ path: bundle });
  return page.evaluate(() => {
    new (window as any).DocxodusPagination.PaginationEngine(
      document.getElementById('staging'), document.getElementById('output'),
      { showPageNumbers: false, fragmentParagraphs: true },
    ).paginate();
    return Array.from(document.querySelectorAll<HTMLElement>('#output .page-box')).map(box => {
      const paper = box.getBoundingClientRect();
      const unit = paper.width / parseFloat(box.style.width);
      return {
        owner: Number(box.dataset.sectionIndex), pageInSection: Number(box.dataset.pageInSection),
        bodyTop: parseFloat(box.querySelector<HTMLElement>('.page-content')!.style.top),
        bodyHeight: parseFloat(box.querySelector<HTMLElement>('.page-content')!.style.height),
        notes: Array.from(box.querySelectorAll<HTMLElement>('.page-footnotes .footnote-item')).map(note => ({
          width: note.getBoundingClientRect().width / unit,
          x: (note.getBoundingClientRect().left - paper.left) / unit,
          text: note.textContent ?? '',
        })),
        images: Array.from(box.querySelectorAll<HTMLImageElement>('img')).map(image => ({
          width: image.getBoundingClientRect().width / unit,
          height: image.getBoundingClientRect().height / unit,
        })),
        exclusions: Array.from(box.querySelectorAll<HTMLElement>('[data-docx-shape-exclusion]'))
          .map(element => element.getBoundingClientRect().width / unit),
        blocks: Array.from(box.querySelectorAll<HTMLElement>('[data-marker]')).map(p => {
          const r = p.getBoundingClientRect();
          const words: Array<{ text: string; x: number; y: number; right: number; bottom: number }> = [];
          const walker = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
          for (let n = walker.nextNode(); n; n = walker.nextNode()) {
            for (const match of n.textContent!.matchAll(/\S+/g)) {
              const range = document.createRange();
              range.setStart(n, match.index!); range.setEnd(n, match.index! + match[0].length);
              const word = range.getBoundingClientRect();
              words.push({ text: match[0], x: (word.left - paper.left) / unit,
                y: (word.top - paper.top) / unit, right: (word.right - paper.left) / unit,
                bottom: (word.bottom - paper.top) / unit });
            }
          }
          return { id: p.dataset.marker, x: (r.left - paper.left) / unit,
            y: (r.top - paper.top) / unit, width: r.width / unit, height: r.height / unit, words };
        }),
      };
    });
  });
}

for (const side of ['left', 'right', 'top', 'bottom'] as const) {
  test(`a continuous ${side} margin change shares the existing page`, async ({ page }) => {
    const pages = await layout(page, section(0, paragraph('before'))
      + section(1, paragraph('after'), { [side]: 60 }));
    expect(pages).toHaveLength(1);
    expect(pages[0].blocks.map(p => p.id)).toEqual(['before', 'after']);
    expect(pages[0].blocks[1].x).toBeCloseTo(side === 'left' ? 60 : 30, 1);
    expect(pages[0].blocks[1].width).toBeCloseTo(side === 'left' || side === 'right' ? 210 : 240, 1);
    expect(pages[0].blocks[1].y).toBeCloseTo(50, 1);
    expect(pages[0].bodyTop).toBe(30);
    expect(pages[0].bodyHeight).toBe(180);
  });
}

test('section width preserves authored paragraph width and indentation', async ({ page }) => {
  const pages = await layout(page, section(0, paragraph('before'))
    + section(1, paragraph('inset', 20, 'width:50%;margin-left:12pt'), { left: 60, right: 60 })
    + section(2, paragraph('restored')));
  expect(pages).toHaveLength(1);
  expect(pages[0].blocks.map(p => p.id)).toEqual(['before', 'inset', 'restored']);
  expect(pages[0].blocks[1].x).toBeCloseTo(72, 1);
  expect(pages[0].blocks[1].width).toBeCloseTo(90, 1);
  expect(pages[0].blocks[2].x).toBeCloseTo(30, 1);
  expect(pages[0].blocks[2].width).toBeCloseTo(240, 1);
});

for (const side of ['top', 'bottom', 'left'] as const) {
  test(`a ${side} margin change adopts page geometry only on spill`, async ({ page }) => {
    const pages = await layout(page, section(0, paragraph('before', 100))
      + section(1, paragraph('shared', 60) + paragraph('spill', 80), { [side]: 60 }));
    expect(pages.map(p => p.blocks.map(b => b.id))).toEqual([['before', 'shared'], ['spill']]);
    expect(pages.map(p => p.owner)).toEqual([0, 1]);
    expect(pages.map(p => p.pageInSection)).toEqual([1, 2]);
    expect(pages.map(p => p.bodyTop)).toEqual([30, side === 'top' ? 60 : 30]);
    expect(pages.map(p => p.bodyHeight)).toEqual([180, side === 'left' ? 180 : 150]);
    expect(pages[0].blocks[1].x).toBeCloseTo(side === 'left' ? 60 : 30, 1);
    expect(pages[1].blocks[0].x).toBeCloseTo(side === 'left' ? 60 : 30, 1);
  });
}

test('explicit nextPage still separates equal-sized sections', async ({ page }) => {
  const pages = await layout(page, section(0, paragraph('before'))
    + section(1, paragraph('after'), {}, 'nextPage'));
  expect(pages.map(p => p.blocks.map(b => b.id))).toEqual([['before'], ['after']]);
});

test('the existing paper-size boundary remains separate', async ({ page }) => {
  const pages = await layout(page, section(0, paragraph('before'))
    + section(1, paragraph('after'), {}, 'continuous', 360));
  expect(pages.map(p => p.blocks.map(b => b.id))).toEqual([['before'], ['after']]);
});

test('horizontal section containers preserve collapsed paragraph spacing', async ({ page }) => {
  const pages = await layout(page, section(0, paragraph('before', 20, 'margin-bottom:12pt'))
    + section(1, paragraph('narrow', 20, 'margin-top:18pt;margin-bottom:10pt'), { left: 60 })
    + section(2, paragraph('restored', 20, 'margin-top:15pt')));
  expect(pages).toHaveLength(1);
  for (const [i, y] of [30, 68, 103].entries()) expect(pages[0].blocks[i].y).toBeCloseTo(y, 1);
});

for (const float of [false, true]) {
  test(`narrowed paragraph fragments retain all words${float ? ' beside an earlier float' : ''}`, async ({ page }) => {
    const words = Array.from({ length: 150 }, (_, i) => `Sequoia${String(i).padStart(3, '0')}`);
    const before = float
      ? '<p data-marker="before"><span data-docx-shape-exclusion="true" style="float:left;width:90pt;height:72pt"></span>before</p>'
      : paragraph('before', 40);
    const pages = await layout(page, section(0, before)
      + section(1, `<p data-marker="long">${words.join(' ')}</p>`, { left: 60 }));
    expect(pages.length).toBeGreaterThan(1);
    expect(pages[0].blocks.map(b => b.id)).toEqual(['before', 'long']);
    const fragments = pages.flatMap(p => p.blocks.filter(b => b.id === 'long'));
    expect(fragments.flatMap(p => p.words.map(w => w.text))).toEqual(words);
    for (const p of pages) {
      for (const block of p.blocks.filter(b => b.id === 'long')) {
        expect(block.x).toBeCloseTo(60, 1);
        expect(block.width).toBeCloseTo(210, 1);
        for (const word of block.words) {
          expect(word.x).toBeGreaterThanOrEqual(59.9);
          expect(word.right).toBeLessThanOrEqual(270.1);
          expect(word.bottom).toBeLessThanOrEqual(p.bodyTop + p.bodyHeight + 0.5);
          if (float && p === pages[0] && word.y < 102) expect(word.x).toBeGreaterThanOrEqual(119.9);
        }
      }
    }
  });
}

test('page-margin anchors retain page geometry while column anchors follow the new section', async ({ page }) => {
  const anchors = ['margin', 'column'].map((basis, i) => `<span data-marker="${basis}"
    data-docx-drawing-anchor="true" data-docx-anchor-h-relative="${basis}" data-docx-anchor-h-offset="0"
    data-docx-anchor-v-relative="page" data-docx-anchor-v-offset="${80 + i * 40}"
    style="position:absolute;width:30pt;height:12pt"></span>`).join('');
  const pages = await layout(page, section(0, paragraph('before'))
    + section(1, `<p>after${anchors}</p>`, { left: 60 }));
  expect(pages).toHaveLength(1);
  expect(pages[0].blocks.find(b => b.id === 'margin')!.x).toBeCloseTo(30, 1);
  expect(pages[0].blocks.find(b => b.id === 'column')!.x).toBeCloseTo(60, 1);
});

test('a widened section measures the actual earlier float before admitting its words', async ({ page }) => {
  const words = Array.from({ length: 10 }, (_, i) => `LongWord${String(i).padStart(3, '0')}`);
  const floated = '<p data-marker="float"><span data-docx-shape-exclusion="true" style="float:left;width:90pt;height:72pt"></span>before</p>';
  const pages = await layout(page, section(0, paragraph('filler', 100) + floated, { left: 60 })
    + section(1, `<p data-marker="long">${words.join(' ')}</p>`));
  expect(pages.length).toBeGreaterThan(1);
  expect(pages.flatMap(p => p.blocks.filter(b => b.id === 'long').flatMap(b => b.words.map(w => w.text)))).toEqual(words);
  for (const p of pages) for (const b of p.blocks.filter(b => b.id === 'long')) {
    for (const word of b.words) expect(word.bottom).toBeLessThanOrEqual(p.bodyTop + p.bodyHeight + 0.5);
  }
});

for (const kept of [true, false]) {
  test(`a widened section ${kept ? 'keeps its heading with floated text' : 'discards float height after a page turn'}`, async ({ page }) => {
    const words = Array.from({ length: 10 }, (_, i) => `LongWord${String(i).padStart(3, '0')}`);
    const floated = '<p data-marker="float"><span data-docx-shape-exclusion="true" style="float:left;width:90pt;height:72pt"></span>before</p>';
    const heading = '<p data-marker="heading" data-keep-with-next="true">HEADING</p>';
    const pages = await layout(page, section(0, paragraph('filler', kept ? 90 : 100) + floated, { left: 60 })
      + section(1, `${kept ? heading : ''}<p data-marker="long" data-keep-lines="true">${words.join(' ')}</p>${kept ? '' : paragraph('tail', 120)}`));
    expect(pages.map(p => p.blocks.map(b => b.id))).toEqual([
      ['filler', 'float'], kept ? ['heading', 'long'] : ['long', 'tail'],
    ]);
    expect(pages[1].blocks.find(b => b.id === 'long')!.words.map(w => w.text)).toEqual(words);
    for (const block of pages[1].blocks) expect(block.y + block.height).toBeLessThanOrEqual(210.1);
  });
}

test('relative shape exclusion and painted width use the same page margins', async ({ page }) => {
  const anchor = `<span data-docx-shape-exclusion="wrapSquare" style="float:left;position:relative;width:30pt;height:48pt">
    <span data-marker="shape" data-docx-shape-anchor="true" data-docx-drawing-anchor="true"
      data-docx-anchor-wrap="wrapSquare" data-docx-anchor-width-relative="margin" data-docx-anchor-width-percent="50"
      data-docx-anchor-h-relative="column" data-docx-anchor-h-offset="0"
      data-docx-anchor-v-relative="paragraph" data-docx-anchor-v-offset="0"
      data-docx-anchor-extent-width="30" data-docx-anchor-extent-height="48"
      style="position:absolute;width:30pt;height:48pt;background:blue"></span></span>`;
  const words = 'ALDER BIRCH CEDAR DOGWOOD ELM FIR GINKGO HAZEL IVY JUNIPER'.split(' ');
  const pages = await layout(page, section(0, paragraph('before', 12))
    + section(1, `<p data-marker="text">${anchor}${words.join(' ')}</p>`, { left: 90 }));
  expect(pages).toHaveLength(1);
  const shape = pages[0].blocks.find(b => b.id === 'shape')!;
  expect(shape.x).toBeCloseTo(90, 1);
  expect(shape.width).toBeCloseTo(120, 1);
  expect(pages[0].exclusions).toEqual([120]);
  const text = pages[0].blocks.find(b => b.id === 'text')!.words;
  expect(text.map(w => w.text)).toEqual(words);
  for (const word of text) {
    expect(word.y >= shape.y + shape.height - 0.1 || word.x >= shape.x + shape.width - 0.1).toBe(true);
  }
});

test('top/bottom shape clearance uses the shared page height', async ({ page }) => {
  const shape = `<span data-docx-shape-clearance="true" aria-hidden="true" data-page-map-exclude="true"
      style="float:left;width:100%;height:155pt"></span>
    <span data-marker="shape" data-docx-shape-anchor="true" data-docx-drawing-anchor="true"
      data-docx-anchor-wrap="wrapTopAndBottom" data-docx-anchor-h-relative="column" data-docx-anchor-h-offset="0"
      data-docx-anchor-v-relative="paragraph" data-docx-anchor-v-offset="0"
      data-docx-anchor-extent-width="120" data-docx-anchor-extent-height="155"
      style="position:absolute;width:120pt;height:155pt;background:blue"></span>`;
  const pages = await layout(page, section(0, paragraph('before', 12))
    + section(1, `<p data-marker="anchor">${shape}ANCHOR</p><p data-marker="after">AFTER</p>`, { top: 60 }));
  expect(pages).toHaveLength(2);
  const box = pages[0].blocks.find(b => b.id === 'shape')!;
  expect(box.y).toBeCloseTo(42, 1);
  expect(box.height).toBeCloseTo(155, 1);
  expect(pages[0].blocks.find(b => b.id === 'anchor')!.words[0].y).toBeGreaterThanOrEqual(196.9);
  expect(pages[1].blocks.map(b => b.id)).toEqual(['after']);
});

for (const count of [15, 60]) {
for (const left of [0, 90]) {
  test(`a ${count}-word note uses its section width within the page band (left ${left})`, async ({ page }) => {
    const words = Array.from({ length: count }, (_, i) => `Note${String(i).padStart(3, '0')}`);
    const registry = `<div id="pagination-footnote-registry"><div class="footnote-item" data-footnote-id="note"
      style="font:10pt/12pt monospace"><span class="footnote-content"><p>${words.join(' ')}</p></span></div></div>`;
    const pages = await layout(page, registry + section(0, paragraph('before'))
      + section(1, '<p data-marker="after">after<sup data-footnote-id="note">1</sup></p>', { left }));
    expect(pages[0].blocks.map(b => b.id)).toEqual(['before', 'after']);
    expect(pages.flatMap(p => p.notes.flatMap(n => n.text.match(/Note\d{3}/g) ?? []))).toEqual(words);
    if (count === 15) expect(pages).toHaveLength(1);
    else expect(pages.length).toBeGreaterThan(1);
    for (const note of pages.flatMap(p => p.notes)) {
      expect(note.width).toBeCloseTo(left === 0 ? 240 : 180, 1);
      expect(note.x).toBeCloseTo(30, 1);
    }
  });
}
}

test('an inline picture that fits the shared page retains its authored height', async ({ page }) => {
  const image = '<img alt="" style="width:20pt;height:155pt" src="data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%228%22 height=%228%22%3E%3Crect width=%228%22 height=%228%22 fill=%22blue%22/%3E%3C/svg%3E">';
  const pages = await layout(page, section(0, paragraph('before', 12))
    + section(1, `<p>${image}</p>`, { top: 60 }));
  expect(pages).toHaveLength(1);
  expect(pages[0].images[0].height).toBeCloseTo(155, 1);
  expect(pages[0].images[0].width).toBeCloseTo(20, 1);
});

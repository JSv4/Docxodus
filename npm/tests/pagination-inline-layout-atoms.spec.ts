import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const kind of ['tab', 'shape', 'unmarked', 'positioned', 'nonempty-tab'] as const) {
  test(`known ${kind} boxes preserve complete content at paragraph boundaries`, async ({ page }) => {
    const content = Array.from({ length: 29 }, (_, i) => {
      const shape = kind === 'shape';
      const box = shape ? `<span data-docx-shape="unrendered" style="display:inline-block;width:40pt;height:12pt">shape${i}</span>`
        : `<span data-atom="${i}" style="display:inline-flex;${kind === 'positioned' ? 'position:relative;left:1pt' : ''}"><span style="display:block;white-space:nowrap">tag${i}</span>
          <span ${kind === 'unmarked' ? '' : 'data-docx-tab="left"'} style="display:block;height:0;width:24pt;border-bottom:1px dotted">${kind === 'nonempty-tab' ? 'filled' : ''}</span></span>`;
      return `${box} Alder${i} recorded e\u0301 and \u{1f332}`;
    }).join('<br>');
    await page.setContent(`<style>p{margin:0;font:12pt/16pt Arial}</style><div id="staging"><section data-section-index="0"
      data-page-width="300" data-page-height="160" data-content-width="260" data-content-height="120"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><p>${content}</p></section></div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(() => {
      new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
        { fragmentParagraphs: true, showPageNumbers: false }).paginate();
      return { paragraphs: document.querySelectorAll('#container p').length, text: document.getElementById('container')!.textContent,
        atoms: Array.from(document.querySelectorAll<HTMLElement>('#container [data-atom], #container [data-docx-shape]'), el => ({
          width: el.getBoundingClientRect().width, text: el.textContent, id: el.dataset.atom })) };
    });
    expect(result.text?.match(/Alder\d+/g)).toEqual(Array.from({ length: 29 }, (_, i) => `Alder${i}`));
    expect(result.text?.match(/e\u0301 and \u{1f332}/gu)).toHaveLength(29);
    expect(result.atoms).toHaveLength(29);
    if (kind === 'tab' || kind === 'shape') expect(result.paragraphs).toBeGreaterThan(1);
    else expect(result.paragraphs).toBe(1);
    if (kind === 'shape') expect(result.atoms.every(atom => Math.abs(atom.width - 40 * 4 / 3) < 0.1)).toBe(true);
    else expect(result.atoms.map(atom => atom.id)).toEqual(Array.from({ length: 29 }, (_, i) => String(i)));
  });
}

test('fitted Word line wrappers remain whole across paragraph continuations', async ({ page }) => {
  const lines = Array.from({ length: 29 }, (_, i) => `<span id="line-${i}" data-docx-word-line="true" style="display:block;white-space:pre">Cedar${String(i).padStart(2, '0')}  recorded. </span>`).join('');
  await page.setContent(`<style>p{margin:0;font:12pt/16pt Arial}</style><div id="staging"><section data-section-index="0"
    data-page-width="300" data-page-height="160" data-content-width="260" data-content-height="120"
    data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><p><span data-docx-word-fit="1">${lines}</span></p></section></div><div id="container"></div>`);
  await page.addScriptTag({ path: bundle });
  const result = await page.evaluate(() => {
    new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
      { fragmentParagraphs: true, showPageNumbers: false }).paginate();
    return { paragraphs: document.querySelectorAll('#container p').length,
      lines: Array.from(document.querySelectorAll('#container [data-docx-word-line]'), line => ({ id: line.id, text: line.textContent })) };
  });
  expect(result.paragraphs).toBeGreaterThan(1);
  expect(result.lines).toEqual(Array.from({ length: 29 }, (_, i) => ({ id: `line-${i}`, text: `Cedar${String(i).padStart(2, '0')}  recorded. ` })));
});

import { expect, test } from '@playwright/test';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { renderDocxArtifacts } from '../../npm-export/dist/index.js';
import { storyTableWidthDocx } from './docx-story-table-width-fixture.js';
import { fileURLToPath } from 'node:url';

test.use({ launchOptions: { chromiumSandbox: true,
  ...(process.env.DOCXODUS_CHROMIUM_PATH ? { executablePath: process.env.DOCXODUS_CHROMIUM_PATH } : {}) } });

for (const kind of ['wide', 'normal', 'margin'] as const) {
  test(`${kind} running-story tables preserve all labels on both pages`, async ({ browser, page }) => {
    test.setTimeout(120_000);
    const result = await renderDocxArtifacts(storyTableWidthDocx(kind), { browser, outputs: ['html', 'pdf'],
      reviewProfile: 'markup', markupPresentation: 'word', commentProfile: 'margin', commentColumnPlacement: 'balloons' });
    expect(result.pageCount).toBe(2);
    const task = getDocument({ data: new Uint8Array(result.pdf!), verbosity: 0 });
    try {
      const pdf = await task.promise;
      for (let index = 1; index <= pdf.numPages; index++) {
        const content = await (await pdf.getPage(index)).getTextContent();
        const text = content.items.filter(item => 'str' in item).map(item => item.str).join(' ');
        for (const story of ['top','base']) for (const label of ['amber','blue','copper'])
          expect(text.split(`Moss ${story} ${label}`)).toHaveLength(2);
      }
    } finally { await task.destroy(); }
    await page.setContent(result.html!);
    const grids = await page.locator('.page-header table, .page-footer table').evaluateAll(tables => tables.map(table => ({
      width: table.getBoundingClientRect().width * 0.75,
      font: getComputedStyle(table.querySelector('span')!).fontSize,
      fitted: (table as HTMLElement).style.getPropertyPriority('width') === 'important',
    })));
    expect(grids).toHaveLength(4);
    for (const grid of grids) {
      expect(grid.width).toBeCloseTo(kind === 'wide' ? 288 : kind === 'normal' ? 252 : 330, 0);
      expect(grid.font).toBe('12px');
      expect(grid.fitted).toBe(kind === 'wide');
    }
    for (const story of ['header','footer']) {
      const paragraphs = result.pageMap.fragments.filter(f => f.story === story && f.anchorId.startsWith('p:'));
      expect(paragraphs).toHaveLength(6);
      expect(paragraphs.every(f => f.geometry.x >= 0 && f.geometry.x + f.geometry.width <= 360.5)).toBe(true);
    }
  });
}

for (const fragmentParagraphs of [false, true]) {
  test(`running-story fitting follows each section and the read-only option=${fragmentParagraphs}`, async ({ page }) => {
    const grid = '<table style="width:450pt;table-layout:fixed;border-collapse:collapse"><colgroup><col style="width:450pt"></colgroup><tbody><tr><td style="width:450pt;padding:0">Moss checks.</td></tr></tbody></table>';
    await page.setContent(`<div id="viewer"></div>`);
    await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
    const result = await page.evaluate(({ grid, fragmentParagraphs }) => {
      const html = `<div id="pagination-staging"><div id="pagination-hf-registry" style="display:none">
        ${[0,1].map(index => `<div data-section="${index}" data-hf-type="header-default">${grid}</div>`).join('')}</div>
        ${[300,400].map((width,index) => `<section data-section-index="${index}" data-page-width="${width}" data-page-height="360"
          data-content-width="${width - 72}" data-content-height="288" data-margin-left="36" data-margin-right="36"
          data-margin-top="36" data-margin-bottom="36"><p>Inventory ${index + 1}.</p></section>`).join('')}</div><div id="pagination-container"></div>`;
      const result = (window as any).DocxodusPagination.paginateHtml(html, 'viewer', { fragmentParagraphs, showPageNumbers:false });
      return { pages: result.totalPages,
        widths: Array.from(document.querySelectorAll('.page-header table'), table => table.getBoundingClientRect().width * 0.75) };
    }, { grid, fragmentParagraphs });
    expect(result.pages).toBe(2);
    expect(result.widths.map(width => Math.round(width))).toEqual(fragmentParagraphs ? [228,328] : [450,450]);
  });
}

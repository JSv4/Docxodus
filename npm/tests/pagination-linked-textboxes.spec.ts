import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const mode of ['flow', 'repeat', 'overset', 'wide', 'duplicate', 'disabled'] as const) {
  test(`linked textbox ${mode} keeps story content explicit`, async ({ page }) => {
    const words = Array.from({ length: mode === 'overset' ? 700 : 70 }, (_, i) => `Birch${String(i).padStart(3, '0')}`).join(' ');
    const boxes = (repeated: boolean) => `<span data-docx-textbox-linked="true" data-docx-textbox-story="original#8" data-docx-textbox-seq="0" ${repeated ? 'data-docx-textbox-repeat="true"' : ''}
      style="position:absolute;left:0;top:0;display:inline-block;box-sizing:border-box;width:144pt;height:60pt;padding:3pt;overflow:hidden">
      <span data-source-anchor-id="p:original:story" data-docx-textbox-paragraph="true" data-widow-control="false" style="display:block;font:11pt/12pt Arial">${mode === 'wide' ? `<span style="white-space:nowrap">Birch${'X'.repeat(200)}</span>` : words}</span></span>
      ${Array.from({ length: mode === 'duplicate' ? 2 : 1 }, () => `<span data-docx-textbox-linked="true" data-docx-textbox-story="original#8" data-docx-textbox-seq="1" ${repeated ? 'data-docx-textbox-repeat="true"' : ''}
      style="position:absolute;left:240pt;top:0;display:inline-block;box-sizing:border-box;width:216pt;height:216pt;padding:3pt;overflow:hidden"></span>`).join('')}`;
    const section = (index: number) => `<section data-section-index="${index}" data-page-width="500" data-page-height="340" data-content-width="460" data-content-height="300"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><p style="margin:0;position:relative">Carrier${boxes(mode === 'repeat')}</p></section>`;
    await page.setContent(`<div id="staging">${section(0)}${mode === 'repeat' ? section(1) : ''}</div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(enabled => {
      try {
        const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
          { fragmentParagraphs: enabled, showPageNumbers: false }).paginate();
        return { pages: result.totalPages, boxes: Array.from(document.querySelectorAll('#container [data-docx-textbox-linked]'), box => box.textContent), error: null };
      } catch (error) { return { pages: 0, boxes: [], error: String(error) }; }
    }, mode !== 'disabled');
    if (mode === 'overset' || mode === 'wide') expect(result.error).toContain('does not fit');
    else if (mode === 'duplicate') expect(result.error).toContain('repeated frame sequence');
    else {
      expect(result.error).toBeNull();
      expect(result.pages).toBe(mode === 'repeat' ? 2 : 1);
      const all = result.boxes.join(' ').match(/Birch\d+/g);
      expect(all).toEqual(Array.from({ length: mode === 'repeat' ? 2 : 1 }, () => words.split(' ')).flat());
      if (mode === 'disabled') expect(result.boxes[1]?.trim()).toBe('');
      else { expect(result.boxes[0]).toContain('Birch000'); expect(result.boxes[1]).toContain('Birch069'); }
    }
  });
}

import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';
const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const mode of ['flow', 'repeat', 'overset', 'wide', 'duplicate', 'disabled', 'bottom', 'center', 'kept', 'impossible-keep', 'empty', 'repeat-custom'] as const) {
  test(`linked textbox ${mode} keeps story content explicit`, async ({ page }) => {
    const repeat = mode === 'repeat' || mode === 'repeat-custom';
    const words = Array.from({ length: mode === 'overset' ? 700 : mode === 'kept' ? 16 : mode === 'impossible-keep' ? 60 : 70 }, (_, i) => `Birch${String(i).padStart(3, '0')}`).join(' ');
    const paragraph = (text: string, keep = false) => `<span data-docx-textbox-paragraph="true" data-keep-with-next="${keep}" data-widow-control="false" style="display:block;flex-shrink:0;font:11pt/12pt Arial">${text}</span>`;
    const content = mode === 'kept' || mode === 'impossible-keep'
      ? paragraph('BirchIntro') + Array.from({ length: mode === 'kept' ? 2 : 6 }, (_, i) => {
        const count = mode === 'kept' ? 8 : 10;
        return paragraph(words.split(' ').slice(i * count, (i + 1) * count).join(' '), i < (mode === 'kept' ? 1 : 5));
      }).join('')
      : (mode === 'empty' ? paragraph('') : '') + paragraph(mode === 'wide' ? `<span style="white-space:nowrap">Birch${'X'.repeat(200)}</span>` : words);
    const boxes = (repeated: boolean) => `<span data-docx-textbox-linked="true" data-docx-textbox-story="original#8" data-docx-textbox-seq="0" ${repeated ? 'data-docx-textbox-repeat="true"' : ''}
      style="position:absolute;left:0;top:0;display:${mode === 'bottom' || mode === 'center' ? `flex;flex-direction:column;justify-content:${mode === 'bottom' ? 'flex-end' : 'center'}` : 'inline-block'};box-sizing:border-box;width:144pt;height:60pt;padding:3pt;overflow:hidden">
      ${content}</span>
      ${Array.from({ length: mode === 'duplicate' ? 2 : 1 }, () => `<span data-docx-textbox-linked="true" data-docx-textbox-story="original#8" data-docx-textbox-seq="1" ${repeated ? 'data-docx-textbox-repeat="true"' : ''}
      style="position:absolute;left:240pt;top:0;display:inline-block;box-sizing:border-box;width:216pt;height:216pt;padding:3pt;overflow:hidden"></span>`).join('')}`;
    const section = (index: number) => `<section data-section-index="${index}" data-page-width="500" data-page-height="340" data-content-width="460" data-content-height="300"
      data-margin-top="20" data-margin-bottom="20" data-margin-left="20" data-margin-right="20"><p style="margin:0;position:relative">Carrier${boxes(repeat)}</p></section>`;
    await page.setContent(`<div id="staging">${section(0)}${repeat ? section(1) : ''}</div><div id="container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(({ enabled, custom }) => {
      try {
        const result = new (window as any).DocxodusPagination.PaginationEngine(document.getElementById('staging'), document.getElementById('container'),
          { fragmentParagraphs: enabled, showPageNumbers: false, cssPrefix: custom ? 'leaf-' : undefined }).paginate();
        return { pages: result.totalPages, boxes: Array.from(document.querySelectorAll('#container [data-docx-textbox-linked]'), box => box.textContent), error: null };
      } catch (error) { return { pages: 0, boxes: [], error: String(error) }; }
    }, { enabled: mode !== 'disabled', custom: mode === 'repeat-custom' });
    if (mode === 'overset' || mode === 'wide') expect(result.error).toContain('does not fit');
    else if (mode === 'duplicate') expect(result.error).toContain('repeated frame sequence');
    else {
      expect(result.error).toBeNull();
      expect(result.pages).toBe(repeat ? 2 : 1);
      const all = result.boxes.join(' ').match(/Birch\d+/g);
      expect(all).toEqual(Array.from({ length: repeat ? 2 : 1 }, () => words.split(' ')).flat());
      if (mode === 'disabled') expect(result.boxes[1]?.trim()).toBe('');
      else if (mode === 'kept') { expect(result.boxes[0]).not.toContain('Birch000'); expect(result.boxes[1]).toContain('Birch015'); }
      else { expect(result.boxes[0]).toContain('Birch000'); expect(result.boxes[1]).toContain(mode === 'impossible-keep' ? 'Birch059' : 'Birch069'); }
    }
  });
}

import { expect, test } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const bundle = join(dirname(fileURLToPath(import.meta.url)), '../dist/pagination.bundle.js');

for (const kind of ['raised', 'ordinary', 'lowered', 'clipped-parent', 'clipped-inside', 'transformed'] as const) {
  test(`inline textbox ${kind} preserves painting and source completeness`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(kind => {
      const position = kind === 'ordinary' ? 0 : kind === 'lowered' ? -24 : 24;
      const box = `<span data-docx-textbox="true" data-docx-textbox-clip="true"
        style="${kind === 'transformed' ? 'transform:translateX(5pt);' : ''}display:inline-block;box-sizing:border-box;width:180pt;height:36pt;overflow:hidden;vertical-align:top;line-height:normal">
        <span data-source-anchor-id="p:body:label" style="display:block;font-size:12pt;line-height:14pt;${kind === 'clipped-inside' ? 'margin-top:40pt' : ''}">Cobalt label inside.</span></span>`;
      const run = `<span style="vertical-align:${position}pt">${box}</span>`;
      const html = `<div id="pagination-staging"><div data-section-index="0"
        data-page-width="432" data-page-height="576" data-content-width="360" data-content-height="504"
        data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
        <p data-source-anchor-id="p:body:host" style="font:10pt/12pt 'Liberation Serif';margin:0">${kind === 'clipped-parent' ? `<span style="display:inline-block;overflow:hidden;height:12pt">${run}</span>` : run}</p>
        <p data-source-anchor-id="p:body:tail" style="font:10pt/12pt 'Liberation Serif';margin:0">Cobalt inventory complete.</p>
      </div></div><div id="pagination-container"></div>`;
      try {
        const result = (window as any).DocxodusPagination.paginateHtml(html, 'viewer', {
          showPageNumbers: false, layoutToken: {documentVersion:0,rendererFingerprint:'raised-box'},
        });
        const label = document.querySelector<HTMLElement>('[data-source-anchor-id="p:body:label"]')!;
        const body = document.querySelector('.page-content')!.getBoundingClientRect();
        return {error:null, anchors:result.pageMap.fragments.map((f:any)=>f.anchorId),
          top:label.getBoundingClientRect().top-body.top, font:getComputedStyle(label).fontFamily};
      } catch (error) { return {error:String(error),anchors:[],top:0,font:''}; }
    }, kind);
    if (kind === 'clipped-parent' || kind === 'clipped-inside' || kind === 'transformed') {
      expect(result.error).toContain('has no measurable fragment');
    } else {
      expect(result.error).toBeNull();
      expect(result.anchors).toEqual(expect.arrayContaining(['p:body:host','p:body:label','p:body:tail']));
      expect(result.font).toContain('Liberation Serif');
      if (kind === 'raised') expect(result.top).toBeLessThan(-20);
    }
  });
}

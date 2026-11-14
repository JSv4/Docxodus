import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

for (const kind of ['right', 'left', 'spaces', 'hidden', 'source-clip', 'source-svg', 'paint-svg', 'border-svg', 'clipped-svg'] as const) {
  test(`table clipping distinguishes ${kind} text from painted ink`, async ({ page }) => {
    await page.setContent('<main id="fixture"></main>');
    await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
    const result = await page.evaluate(kind => {
      const content = kind.endsWith('svg')
        ? `<svg ${kind === 'source-svg' || kind === 'clipped-svg' ? 'data-docx-visual-resource-id="1"' : kind === 'border-svg' ? 'aria-hidden="true"' : ''} style="width:180pt;height:12pt"><rect width="240" height="16" fill="blue"/></svg>`
        : kind === 'spaces' ? `ABC${' '.repeat(50)}` : 'ABCDEFGHIJKLMNO';
      const style = kind === 'hidden' ? 'visibility:hidden' : (kind === 'source-clip' || kind === 'clipped-svg')
        ? 'display:inline-block;width:10pt;overflow:hidden;vertical-align:top' : '';
      document.getElementById('fixture')!.innerHTML = `<div id="staging"><div data-section-index="0"
        data-page-width="200" data-page-height="120" data-content-width="160" data-content-height="80"
        data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
        <table data-docx-fixed-width="true" style="width:240pt;table-layout:fixed;border-collapse:collapse;margin-left:${kind === 'left' ? -40 : 0}pt">
          <tr><td style="padding:0"><p data-source-anchor-id="p:body:sample" style="margin:0;font:10pt monospace;white-space:pre">${kind === 'left' ? '' : '<span>OK</span><span style="display:inline-block;width:140pt"></span>'}<span style="${style}">${content}</span></p></td></tr>
        </table></div></div><div id="pages"></div>`;
      const engine = new (window as any).DocxodusPagination.PaginationEngine(
        document.getElementById('staging'), document.getElementById('pages'),
        { showPageNumbers: false, fragmentParagraphs: false });
      engine.paginate();
      engine.materializePageMap(0, 'table-clipping');
      const clipped = engine.getClippedTableContentAnchors();
      // A second materialization must replace, rather than accumulate, diagnostics.
      document.querySelector('#pages p')!.textContent = kind === 'left' ? ' '.repeat(12) + 'OK' : 'OK';
      engine.materializePageMap(0, 'table-clipping');
      return { clipped, after: engine.getClippedTableContentAnchors() };
    }, kind);
    expect(result.clipped).toEqual(['right', 'left', 'source-svg'].includes(kind) ? ['p:body:sample'] : []);
    expect(result.after).toEqual([]);
  });
}

// Only the paper may hide a whole cell, and only of an authored fixed-width table (issue #1733):
// a table the layout sized, a cell the body band cuts off, and a cell with no box are still lost.
for (const kind of ['fixed', 'unfixed', 'below-band', 'no-box'] as const) {
  test(`a ${kind} table cell beyond the paper ${kind === 'fixed' ? 'is clipped by the paper' : 'fails completeness'}`, async ({ page }) => {
    await page.setContent('<main id="fixture"></main>');
    await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
    const result = await page.evaluate(kind => {
      const far = kind === 'below-band' ? 'position:relative;top:120pt' : kind === 'no-box' ? 'height:0;overflow:hidden' : '';
      document.getElementById('fixture')!.innerHTML = `<div id="staging"><div data-section-index="0"
        data-page-width="200" data-page-height="120" data-content-width="160" data-content-height="80"
        data-margin-top="20" data-margin-right="20" data-margin-bottom="20" data-margin-left="20">
        <table ${kind === 'unfixed' ? '' : 'data-docx-fixed-width="true"'} style="width:300pt;table-layout:fixed;border-collapse:collapse">
          <tr>${['near', 'edge', 'far'].map(cell => `<td data-source-anchor-id="tc:body:${cell}" style="width:100pt;padding:0">
            <p data-source-anchor-id="p:body:${cell}" style="margin:0;font:10pt monospace;${cell === 'far' ? far : ''}">Wren ${cell}</p></td>`).join('')}</tr>
        </table></div></div><div id="pages"></div>`;
      try {
        const engine = new (window as any).DocxodusPagination.PaginationEngine(
          document.getElementById('staging'), document.getElementById('pages'),
          { showPageNumbers: false, fragmentParagraphs: false });
        engine.paginate();
        const map = engine.materializePageMap(0, 'off-paper-cell');
        return { error: '', mapped: map.fragments.map((f: any) => f.anchorId), clipped: engine.getSourceClippedTextAnchors() };
      } catch (error) { return { error: String(error), mapped: [], clipped: [] }; }
    }, kind);
    if (kind !== 'fixed') {
      expect(result.error).toContain('has no measurable fragment');
      return;
    }
    expect(result.error).toBe('');
    expect(result.mapped).toEqual(expect.arrayContaining(['tc:body:near', 'p:body:near', 'tc:body:edge', 'p:body:edge']));
    expect(result.mapped).not.toContain('tc:body:far');
    // The cell's text is its paragraph's, so only the paragraph is reported.
    expect(result.clipped).toEqual(['p:body:far']);
  });
}

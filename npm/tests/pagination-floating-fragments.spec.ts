import { test, expect } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const tile = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7';

for (const scope of ['body', 'cell', 'note', 'textbox', 'layout']) {
  for (const kind of ['leading-box', 'later-box', 'character-box', 'line-box', 'aligned-box', 'wrapped-box',
    'leading-edge', 'later-edge', 'foreign-float', 'shifted-edge', 'transformed-edge', 'text-edge', 'two-image-edge',
    'positioned-image-edge', 'transformed-image-edge']) {
    test(`${scope} fragmentation preserves the ${kind} boundary`, async ({ page }) => {
      await page.setContent('<div id="pagination-staging"></div><div id="pagination-container"></div>');
      await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
      const result = await page.evaluate(({ scope, kind, tile }) => {
        const p = document.createElement(scope === 'textbox' ? 'span' : 'p');
        p.style.cssText = 'display:block;position:relative;width:400px;margin:0;font:12px/16px Arial';
        if (scope === 'textbox') p.dataset.docxTextboxParagraph = 'true';
        const img = `<img src="${tile}" style="width:40px;height:40px;${kind === 'positioned-image-edge' ? 'position:absolute;' : ''}${kind === 'transformed-image-edge' ? 'transform:scale(2);' : ''}">`;
        const box = `<span data-docx-drawing-anchor="true" data-docx-anchor-h-relative="${kind === 'character-box' ? 'character' : 'column'}"
          data-docx-anchor-v-relative="${kind === 'line-box' ? 'line' : 'paragraph'}" data-docx-anchor-wrap="${kind === 'wrapped-box' ? 'wrapSquare' : 'wrapNone'}"
          ${kind === 'aligned-box' ? 'data-docx-anchor-v-align="bottom"' : ''}
          style="position:absolute;left:200px;top:0;width:100px;height:40px"><span data-docx-textbox="true"><span
          data-docx-textbox-paragraph="true" style="display:block">Box story</span></span></span>`;
        const edge = `<span ${kind === 'foreign-float' ? '' : 'data-docx-edge-picture="true"'}
          style="float:left;${kind === 'shifted-edge' ? 'position:relative;top:1px;' : ''}${kind === 'transformed-edge' ? 'transform:scale(2);' : ''}">${img}${kind === 'two-image-edge' ? img : kind === 'text-edge' ? 'Extra text' : ''}</span>`;
        p.innerHTML = (kind.startsWith('later') ? '<span>Before. </span>' : '')
          + (kind.endsWith('box') ? box : edge) + '<span>After the drawing.</span>';
        document.getElementById('pagination-staging')!.append(p);
        const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false });
        if (scope === 'layout') return engine.hasRangeFragmentSafeLayout(p);
        return engine.canRangeFragmentParagraph(p, scope === 'note', scope === 'body', false, scope === 'cell', scope === 'textbox');
      }, { scope, kind, tile });
      // The standalone layout check leaves textless leading atoms to its caller's image gate.
      const leadingLayoutAtom = scope === 'layout' && !kind.endsWith('box') && !['later-edge', 'text-edge'].includes(kind);
      expect(result).toBe(kind === 'leading-box' || leadingLayoutAtom
        || (scope === 'body' && ['later-box', 'leading-edge'].includes(kind)));
    });
  }
}

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

for (const owner of ['body', 'endnote', 'footnote', 'header', 'comment', 'column', 'positioned-run', 'transformed-run']) {
  test(`late edge preparation preserves ${owner} ownership`, async ({ page }) => {
    await page.setContent('<div id="pagination-staging"></div><div id="pagination-container"></div>');
    await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
    const result = await page.evaluate(({ owner, tile }) => {
      const host = document.createElement('div');
      const wrappers: Record<string, string> = {
        endnote: '<section class="endnotes"><ol><li></li></ol></section>',
        footnote: '<div id="pagination-footnote-registry"></div>',
        header: '<div id="pagination-hf-registry"></div>',
        comment: '<div id="pagination-comment-margin-registry"></div>',
        column: '<div data-pagination-columns="true"></div>',
      };
      host.innerHTML = wrappers[owner] ?? '<div></div>';
      const p = document.createElement('p');
      p.style.cssText = 'width:400px;margin:0;font:12px/16px Arial';
      const shell = owner === 'positioned-run' ? 'position:relative;top:-2px' : owner === 'transformed-run' ? 'transform:translateY(-2px)' : '';
      p.innerHTML = `<span>First line<br>Second line<br></span><span style="${shell}"><span data-docx-edge-picture="true" style="float:left;margin-right:12px"><img id="source-image" src="${tile}" style="width:96px;height:96px"></span></span><span>Third line</span>`;
      (host.querySelector('li') ?? host.firstElementChild)!.append(p);
      document.getElementById('pagination-staging')!.append(host);
      const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false });
      const original = host.innerHTML;
      const rawEligible = engine.canRangeFragmentParagraph(p, false, true);
      engine.prepareBodyEdgePictures(host);
      const once = host.innerHTML;
      engine.prepareBodyEdgePictures(host);
      const image = p.querySelector('img')!, box = image.getBoundingClientRect(), paragraph = p.getBoundingClientRect();
      return { rawEligible, eligible: engine.canRangeFragmentParagraph(p, false, true),
        unchanged: original === once, idempotent: once === host.innerHTML,
        text: p.textContent, ids: Array.from(host.querySelectorAll('[id]')).map(el => el.id),
        imageTop: box.top - paragraph.top, imageLeft: box.left - paragraph.left };
    }, { owner, tile });
    expect(result.rawEligible).toBe(false);
    expect(result.idempotent).toBe(true);
    expect(result.unchanged).toBe(owner !== 'body');
    expect(result.eligible).toBe(owner === 'body');
    expect(result.text).toBe('First lineSecond lineThird line');
    expect(result.ids.filter(id => id === 'source-image')).toHaveLength(1);
    if (owner === 'body') { expect(result.imageTop).toBe(0); expect(result.imageLeft).toBe(0); }
  });
}

test('a persistent picture exclusion mismatch cannot escape through oversized fallback', async ({ page }) => {
  await page.setContent(`<div id="pagination-staging"><p style="margin:0;font:12px/16px Arial" data-widow-control="false">
    Before<br><span data-docx-edge-picture="true" style="float:left"><img src="${tile}" style="width:40px;height:40px"></span>After
    </p></div><div id="pagination-container"></div>`);
  await page.addStyleTag({ content: '[data-docx-edge-exclusion] { margin-left: 6px !important; }' });
  await page.addScriptTag({ path: fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url)) });
  const error = await page.evaluate(() => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', {
      fragmentParagraphs: true, showPageNumbers: false,
    });
    try { engine.paginate(); return null; } catch (error) { return (error as Error).message; }
  });
  expect(error).toBe('A wrapping picture does not match its destination exclusion');
});

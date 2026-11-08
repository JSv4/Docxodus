import { expect, test } from '@playwright/test';
import { fileURLToPath } from 'node:url';

const bundle = fileURLToPath(new URL('../dist/pagination.bundle.js', import.meta.url));

for (const relative of ['page', 'paragraph']) for (const scenario of [
  'none', 'zoom', 'balloons', 'blank', 'nbsp', 'wrong-origin', 'offset-before', 'offset-after',
  'member-before', 'member-after', 'transform-before', 'transform-after', 'font-before', 'font-after',
  'text-after', 'host-after', 'removed', 'partial', 'linked', 'percent', 'nested-anchor', 'raw-wrapper', 'partial-ink',
  'blank-wrong-origin', 'blank-member-before', 'blank-visible-before', 'blank-member-after',
  'blank-transform-before', 'blank-font-before', 'blank-font-after', 'blank-host-after', 'blank-text-after',
  'internal', 'internal-percent',
]) test(`group source placement binds ${relative} coordinates: ${scenario}`, async ({ page }) => {
  const change = scenario.replace(/^blank-/, ''), blank = scenario === 'blank' || scenario.startsWith('blank-');
  const internal = change.startsWith('internal');
  const member = `<span data-docx-drawing-member="true" ${change === 'partial' ? 'data-docx-shape="partial"' : ''}
    ${change === 'nested-anchor' ? 'data-docx-drawing-anchor="true" data-docx-anchor-h-relative="page" data-docx-anchor-v-relative="page" data-docx-anchor-h-offset="320" data-docx-anchor-v-offset="20"' : ''}
    style="position:absolute;left:${internal ? 20 : change === 'partial-ink' ? 180 : change === 'visible-before' ? 60 : 320}pt;top:0;width:200pt;height:60pt">
    <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-clip="true"
      style="display:inline-block;position:relative;width:200pt;height:60pt;overflow:hidden">
      ${change === 'partial-ink' ? '<span style="display:block;white-space:nowrap">Birch edge extends across the paper boundary</span>' : ''}
      <span data-source-anchor-id="p:body:hidden" style="display:block;white-space:nowrap;${internal ? 'margin-top:70pt' : change === 'partial-ink' ? 'margin-left:160pt' : ''}">
        ${blank ? '<br>' : change === 'nbsp' ? '&nbsp;' : 'Birch outside'}</span>
    </span></span>`;
  await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
    data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
    data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
    <p data-source-anchor-id="p:body:host" style="font:10pt/12pt Arial;margin:0">Birch host
      <span data-docx-drawing-group="true" data-docx-drawing-anchor="true"
        data-docx-anchor-h-relative="${relative === 'page' ? 'page' : 'column'}" data-docx-anchor-v-relative="${relative}"
        data-docx-anchor-h-offset="${relative === 'page' ? 36 : 0}" data-docx-anchor-v-offset="20"
        ${change === 'linked' ? 'data-docx-textbox-linked="true"' : ''}
        ${change === 'percent' || change === 'internal-percent' ? 'data-docx-anchor-width-percent="20" data-docx-anchor-width-relative="page"' : ''}
        style="display:inline-block;position:absolute;width:100pt;height:60pt">
        ${change === 'raw-wrapper' ? `<span>${member}</span>` : member}
      </span></p>
    <p data-source-anchor-id="p:body:closing" data-format-description="Font: Bold" data-format-key="1"
      style="font:10pt/12pt Arial;margin:0">Birch closing</p>
    </div></div><div id="pagination-container"></div>`);
  await page.addScriptTag({ path: bundle });
  const result = await page.evaluate(({ change }) => {
    const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', {
      showPageNumbers: false, scale: change === 'zoom' ? 0.7 : 1,
      commentColumnPlacement: change === 'balloons' ? 'balloons' : 'margin',
    });
    if (change === 'wrong-origin') {
      const resolve = engine.resolveAnchorAxis.bind(engine);
      engine.resolveAnchorAxis = (...args: any[]) => resolve(...args) + 10;
    }
    if (change.endsWith('-before')) {
      const position = engine.positionDrawingAnchors.bind(engine);
      engine.positionDrawingAnchors = (...args: any[]) => {
        const root = args[1].querySelector('[data-docx-drawing-group]');
        if (root) {
          if (change === 'offset-before') root.dataset.docxAnchorHOffset = '50';
          if (change === 'member-before' || change === 'visible-before') root.firstElementChild.style.left = '350pt';
          if (change === 'transform-before') root.firstElementChild.style.transform = 'translateX(10pt)';
          if (change === 'font-before') root.querySelector('[data-source-anchor-id]').style.font = '16pt/20pt Arial';
        }
        return position(...args);
      };
    }
    try {
      engine.paginate();
      const container = document.getElementById('pagination-container')!;
      const root = container.querySelector<HTMLElement>('[data-docx-drawing-group]')!;
      const member = root.firstElementChild as HTMLElement;
      const text = root.querySelector<HTMLElement>('[data-source-anchor-id]')!;
      if (change === 'offset-after') root.dataset.docxAnchorHOffset = '50';
      if (change === 'member-after') member.style.left = '350pt';
      if (change === 'transform-after') member.style.transform = 'translateX(0)';
      if (change === 'font-after') text.style.font = '16pt/20pt Arial';
      if (change === 'text-after') text.textContent = 'Birch replaced';
      if (change === 'host-after') container.querySelector<HTMLElement>('[data-source-anchor-id="p:body:host"]')!.style.marginTop = '12pt';
      if (change === 'removed') text.remove();
      const before = container.innerHTML;
      const map = engine.materializePageMap(0, 'group-source');
      const again = engine.materializePageMap(0, 'group-source');
      return { error: '', clipped: engine.getSourceClippedTextAnchors(),
        readOnly: before === container.innerHTML && JSON.stringify(map) === JSON.stringify(again),
        anchors: map.fragments.map((f: any) => f.anchorId) };
    } catch (error) { return { error: String(error), clipped: [], readOnly: false, anchors: [] }; }
  }, { change });
  if (internal || ['none', 'zoom', 'balloons', 'blank', 'nbsp'].includes(change) || (change === 'host-after' && relative === 'page')) {
    expect(result.error).toBe('');
    expect(result.clipped).toEqual(blank || change === 'nbsp' ? [] : ['p:body:hidden']);
    expect(result.anchors).toContain('p:body:closing');
    expect(result.anchors).not.toContain('p:body:hidden');
    expect(result.readOnly).toBe(true);
  } else expect(result.error).toContain('has no measurable fragment');
});

import { expect, test } from '@playwright/test';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const bundle = join(dirname(fileURLToPath(import.meta.url)), '../dist/pagination.bundle.js');

for (const relative of ['page', 'paragraph']) for (const scale of [0.7, 1, 1.25])
for (const change of ['none', 'empty', 'blank', 'before', 'after', 'font', 'text', 'host']) {
  test(`balloon shrink preserves the ${relative} paper clip at zoom ${scale}: ${change}`, async ({ page }) => {
    await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
      data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
      data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
      <p data-source-anchor-id="p:body:host" style="font:10pt/12pt Arial;margin:0">Birch host
        <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-position="true"
          data-docx-drawing-anchor="true" data-docx-anchor-h-relative="${relative === 'page' ? 'page' : 'column'}" data-docx-anchor-v-relative="${relative}"
          data-docx-anchor-h-offset="320" data-docx-anchor-v-offset="20"
          style="display:inline-block;position:absolute;width:160pt;height:70pt;overflow:hidden">
          <span data-source-anchor-id="p:body:hidden" style="display:block;font:10pt/12pt Arial">${change === 'empty' ? '<br>' : change === 'blank' ? '&nbsp;' : 'Birch outside'}</span>
        </span></p>
      <p data-source-anchor-id="p:body:closing" data-format-description="Font: Bold" data-format-key="1"
        style="font:10pt/12pt Arial;margin:0">Birch closing</p>
      </div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(({ scale, change }) => {
      try {
        const prefix = scale === 1.25 ? 'quail-' : 'page-';
        const { PaginationEngine, measureRenderedSource } = (window as any).DocxodusPagination;
        const engine = new PaginationEngine('pagination-staging', 'pagination-container', {
          showPageNumbers: false, scale, cssPrefix: prefix, commentColumnPlacement: 'balloons', deferMarkupBalloons: true,
        });
        engine.paginate();
        const root = document.getElementById('pagination-container')!;
        const box = root.querySelector<HTMLElement>('[data-docx-textbox]')!;
        const hidden = box.querySelector<HTMLElement>('[data-source-anchor-id]')!;
        if (change === 'before') box.style.left = '330pt';
        engine.layoutMarkupBalloons();
        if (change === 'after') box.style.left = '330pt';
        if (change === 'font') hidden.style.fontSize = '15pt';
        if (change === 'text') hidden.textContent = 'Birch replaced';
        if (change === 'host') root.querySelector<HTMLElement>('[data-source-anchor-id="p:body:host"]')!.style.marginTop = '10pt';
        const page = root.querySelector<HTMLElement>(`.${prefix}box`)!;
        const { visible } = measureRenderedSource(window, hidden, page, page.getBoundingClientRect());
        const painted = visible.right > visible.left && visible.bottom > visible.top;
        const before = root.innerHTML;
        const map = engine.materializePageMap(0, 'balloon-source-clip');
        const repeated = engine.materializePageMap(0, 'balloon-source-clip');
        const paper = page.getBoundingClientRect();
        const closing = root.querySelector<HTMLElement>('[data-source-anchor-id="p:body:closing"]')!;
        const ink = measureRenderedSource(window, closing, page, paper).visible;
        const expected = { x: (ink.left - paper.left) * 300 / paper.width, y: (ink.top - paper.top) * 360 / paper.height,
          width: (ink.right - ink.left) * 300 / paper.width, height: (ink.bottom - ink.top) * 360 / paper.height };
        const geometry = map.fragments.find((fragment: any) => fragment.anchorId === 'p:body:closing').geometry;
        return { error: '', anchors: map.fragments.map((f: any) => f.anchorId), clipped: engine.getSourceClippedTextAnchors(), painted,
          readOnly: before === root.innerHTML && JSON.stringify(map) === JSON.stringify(repeated),
          physical: Object.entries(expected).every(([key, value]) => Math.abs(geometry[key] - value) < 0.01),
          balloons: root.querySelectorAll(`.${prefix}balloon`).length };
      } catch (error) { return { error: String(error), anchors: [], clipped: [], painted: false, readOnly: false, physical: false, balloons: 0 }; }
    }, { scale, change });
    if (['none', 'empty', 'blank'].includes(change)) {
      expect(result.error).toBe('');
      expect(result.anchors).toContain('p:body:closing');
      expect(result.clipped).toEqual(change === 'none' ? ['p:body:hidden'] : []);
      expect(result.painted).toBe(false);
      expect(result.readOnly).toBe(true);
      expect(result.physical).toBe(true);
      expect(result.balloons).toBeGreaterThan(0);
    } else if (change === 'host' && relative === 'page') expect(result.error).toBe('');
    else expect(result.error).toContain('has no measurable fragment');
  });
}

for (const scale of [0.7, 1, 1.25]) for (const change of [false, true]) {
  test(`balloon clipping cannot excuse newly lost sibling ink at zoom ${scale}, mutation ${change}`, async ({ page }) => {
    await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
      data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
      data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
      <p data-source-anchor-id="p:body:host" data-format-description="Font: Bold" data-format-key="1" style="font:10pt/12pt Arial;margin:0">Birch host
        <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-position="true"
          data-docx-drawing-anchor="true" data-docx-anchor-h-relative="page" data-docx-anchor-v-relative="page"
          data-docx-anchor-h-offset="36" data-docx-anchor-v-offset="330"
          style="display:inline-block;position:absolute;width:160pt;height:70pt;overflow:hidden">
          <span data-source-anchor-id="p:body:visible" style="display:block;font:10pt/12pt Arial;margin-top:10pt">Birch visible</span>
          <span data-source-anchor-id="p:body:hidden" style="display:block;font:10pt/12pt Arial;margin-top:30pt">Birch outside</span>
        </span></p></div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(({ scale, change }) => {
      const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', {
        showPageNumbers: false, scale, commentColumnPlacement: 'balloons', deferMarkupBalloons: true,
      });
      engine.paginate();
      engine.layoutMarkupBalloons();
      const before = engine.materializePageMap(0, 'before');
      if (change) document.querySelector<HTMLElement>('#pagination-container [data-source-anchor-id="p:body:visible"]')!.style.transform = 'translateY(14pt)';
      let error = '';
      try { engine.materializePageMap(0, 'after'); } catch (failure) { error = String(failure); }
      return { error, anchors: before.fragments.map((fragment: any) => fragment.anchorId) };
    }, { scale, change });
    expect(result.anchors).toContain('p:body:visible');
    if (change) expect(result.error).toContain('has no measurable fragment');
    else expect(result.error).toBe('');
  });
}

for (const change of ['none', 'zoom-half', 'zoom-small', 'zoom-large', 'zoom-double', 'upper-edge', 'partial-source', 'partial-before', 'partial-after', 'before-placement', 'offset-before', 'transform-before', 'after-placement', 'font', 'box'] as const) {
  test(`partial source placement preserves original ink geometry: ${change}`, async ({ page }) => {
    await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
      data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
      data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
      <p data-source-anchor-id="p:body:host" style="font:10pt/12pt Arial;margin:0">Birch host
        <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-position="true"
          data-docx-drawing-anchor="true" data-docx-anchor-h-relative="page" data-docx-anchor-v-relative="page"
          data-docx-anchor-h-offset="36" data-docx-anchor-v-offset="330"
          style="display:inline-block;position:absolute;width:160pt;height:70pt;overflow:hidden">
          <span data-source-anchor-id="p:body:visible" style="display:block;font:10pt/12pt Arial;margin-top:10pt">Birch visible</span>
          <span data-source-anchor-id="p:body:marker" style="display:block;font:10pt/12pt Arial;margin-top:30pt">¶</span>
        </span></p></div></div><div id="pagination-container"></div>`);
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(change => {
      try {
        if (change === 'partial-source')
          document.querySelector<HTMLElement>('[data-source-anchor-id="p:body:visible"]')!.style.transform = 'translateY(14pt)';
        if (change === 'upper-edge') {
          const box = document.querySelector<HTMLElement>('[data-docx-textbox]')!;
          const marker = box.querySelector('[data-source-anchor-id="p:body:marker"]')!;
          const range = document.createRange();
          range.selectNodeContents(marker);
          const bottom = range.getBoundingClientRect().bottom - box.getBoundingClientRect().top;
          box.dataset.docxAnchorVOffset = String((-bottom + 0.1) * 0.75);
        }
        const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false, scale: change === 'zoom-half' ? 0.5 : change === 'zoom-small' || change === 'upper-edge' ? 0.7 : change === 'zoom-large' ? 1.25 : change === 'zoom-double' ? 2 : 1 });
        if (['partial-before', 'before-placement', 'offset-before', 'transform-before'].includes(change)) {
          const position = engine.positionDrawingAnchors.bind(engine);
          engine.positionDrawingAnchors = (...args: any[]) => {
            for (const box of args[1].querySelectorAll('[data-docx-textbox]')) {
              if (change === 'partial-before') box.firstElementChild.style.transform = 'translateY(14pt)';
              if (change === 'before-placement') box.firstElementChild.style.marginTop = '40pt';
              if (change === 'offset-before') box.dataset.docxAnchorVOffset = '359';
              if (change === 'transform-before') box.style.transform = 'translateY(29pt)';
            }
            return position(...args);
          };
        }
        engine.paginate();
        const root = document.getElementById('pagination-container')!;
        const box = root.querySelector<HTMLElement>('[data-docx-textbox]')!;
        if (change === 'partial-after') engine.materializePageMap(0, 'before-mutation');
        if (change === 'partial-after') root.querySelector<HTMLElement>('[data-source-anchor-id="p:body:visible"]')!.style.transform = 'translateY(14pt)';
        if (change === 'after-placement') root.querySelector<HTMLElement>('[data-source-anchor-id="p:body:visible"]')!.style.marginTop = '40pt';
        if (change === 'font') root.querySelector<HTMLElement>('[data-source-anchor-id="p:body:marker"]')!.style.fontSize = '15pt';
        if (change === 'box') box.style.top = '340pt';
        const map = engine.materializePageMap(0, 'birch-edge');
        return { error: '', anchors: map.fragments.map((f: any) => f.anchorId), clipped: engine.getSourceClippedTextAnchors() };
      } catch (error) { return { error: String(error), anchors: [], clipped: [] }; }
    }, change);
    if (change === 'none' || change.startsWith('zoom-')) {
      expect(result.error).toBe('');
      expect(result.anchors).toContain('p:body:visible');
      expect(result.clipped).toEqual(['p:body:marker']);
    } else if (change === 'upper-edge') {
      expect(result.clipped).not.toContain('p:body:marker');
      if (!result.error) expect(result.anchors).toContain('p:body:marker');
      else expect(result.error).toContain('has no measurable fragment');
    } else expect(result.error).toContain('has no measurable fragment');
  });
}

test('off-paper textbox verification scans source text linearly', async ({ page }) => {
  const paragraphs = 100;
  await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
    data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
    data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
    <p data-source-anchor-id="p:body:host">Birch host<span data-docx-textbox="true" data-docx-textbox-clip="true"
      data-docx-textbox-source-position="true" data-docx-drawing-anchor="true"
      data-docx-anchor-h-relative="page" data-docx-anchor-v-relative="page"
      data-docx-anchor-h-offset="36" data-docx-anchor-v-offset="400"
      style="display:inline-block;position:absolute;width:160pt;height:70pt;overflow:hidden">
      ${Array.from({ length: paragraphs }, (_, i) => `<span data-source-anchor-id="p:body:t${i}" style="display:block">Birch ${i}</span>`).join('')}
    </span></p></div></div><div id="pagination-container"></div>`);
  await page.addScriptTag({ path: bundle });
  const result = await page.evaluate(() => {
    let ranges = 0;
    const createRange = document.createRange.bind(document);
    document.createRange = () => { ranges++; return createRange(); };
    const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false });
    engine.paginate();
    engine.materializePageMap(0, 'linear-source-scan');
    return { ranges, clipped: engine.getSourceClippedTextAnchors().length };
  });
  expect(result.clipped).toBe(paragraphs);
  expect(result.ranges).toBeLessThanOrEqual(paragraphs * 10);
});

for (const kind of ['fixed', 'unmarked', 'linked', 'autofit', 'partial', 'percentage', 'inactive-percentage', 'visible-zero-width', 'blank-zero-outside', 'blank-zero-inside', 'blank-zero-unmarked', 'unclipped', 'new-loss', 'removed', 'replaced', 'moved', 'ordinary-off-paper'] as const) {
  test(`source textbox clipping preserves the ${kind} completeness boundary`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(kind => {
      const host = document.getElementById('viewer')!;
      host.innerHTML = `<div id="pagination-staging"><div data-section-index="0"
        data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
        data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
        <p data-source-anchor-id="p:body:host" style="font:10pt/14pt Arial;margin:0">
          <span data-docx-textbox="true" data-docx-textbox-clip="true"
            ${kind === 'unmarked' || kind === 'blank-zero-unmarked' || kind === 'visible-zero-width' ? '' : 'data-docx-textbox-source-clip="true"'}
            ${kind === 'linked' ? 'data-docx-textbox-linked="true"' : ''}
            ${kind === 'autofit' ? 'data-docx-anchor-autofit="true"' : ''}
            ${kind === 'partial' ? 'data-docx-shape="partial"' : ''}
            ${kind === 'percentage' ? 'data-docx-anchor-height-percent="50"' : kind === 'inactive-percentage' ? 'data-docx-anchor-width-percent="0" data-docx-anchor-height-percent="0"' : ''}
            style="display:inline-block;width:160pt;height:28pt;overflow:${kind === 'unclipped' ? 'visible' : 'hidden'};vertical-align:top;${kind === 'ordinary-off-paper' ? 'position:absolute;left:-400pt' : ''}">
            <span data-source-anchor-id="p:body:red" style="display:block;font:10pt/14pt Arial">Quail red marker</span>
            <span data-source-anchor-id="p:body:blue" style="display:block;font:10pt/14pt Arial">Quail blue marker</span>
            <span data-source-anchor-id="p:body:last" style="display:block;font:10pt/14pt Arial;${kind.startsWith('blank-zero') || kind === 'visible-zero-width' ? `width:0;position:relative;top:-28pt;margin-left:${kind === 'blank-zero-inside' ? 80 : 180}pt` : ''}">${kind.startsWith('blank-zero') ? '&nbsp;' : 'Quail hidden marker'}</span>
          </span>
        </p><p data-source-anchor-id="p:body:second-host" style="font:10pt/14pt Arial;margin:0">
          <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-clip="true"
            style="display:inline-block;width:160pt;height:14pt;overflow:hidden;vertical-align:top">
            <span data-source-anchor-id="p:body:other" style="display:block;font:10pt/14pt Arial">Quail second box</span>
          </span></p><p data-source-anchor-id="p:body:caption" style="font:10pt/14pt Arial;margin:0">Quail ledger complete</p>
      </div></div><div id="pagination-container"></div>`;
      try {
        const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', {
          showPageNumbers: false,
        });
        engine.paginate();
        const box = host.querySelector<HTMLElement>('#pagination-container [data-docx-textbox]')!;
        if (kind === 'new-loss') box.style.height = '14pt';
        if (kind === 'removed') box.querySelector('[data-source-anchor-id="p:body:last"]')!.remove();
        if (kind === 'replaced') box.querySelector('[data-source-anchor-id="p:body:last"]')!.textContent = 'Quail altered marker';
        if (kind === 'moved') {
          const second = host.querySelector<HTMLElement>('#pagination-container [data-source-anchor-id="p:body:other"]')!.parentElement!;
          second.append(box.querySelector('[data-source-anchor-id="p:body:last"]')!);
        }
        const map = engine.materializePageMap(0, 'quail-clip');
        return { error: null, anchors: map.fragments.map((f: any) => f.anchorId) };
      } catch (error) { return { error: String(error), anchors: [] }; }
    }, kind);
    if (['unmarked', 'linked', 'autofit', 'partial', 'percentage', 'visible-zero-width', 'blank-zero-inside', 'blank-zero-unmarked', 'new-loss', 'removed', 'replaced', 'moved', 'ordinary-off-paper'].includes(kind)) {
      expect(result.error).toContain('has no measurable fragment');
    } else {
      expect(result.error).toBeNull();
      expect(result.anchors).toContain('p:body:red');
      expect(result.anchors).toContain('p:body:blue');
      expect(result.anchors).toContain('p:body:caption');
      expect(result.anchors.includes('p:body:last')).toBe(kind === 'unclipped');
    }
  });
}

for (const story of ['header', 'footer'] as const) for (const mutate of [false, true]) {
  test(`${story} source clipping ${mutate ? 'rejects new loss' : 'retains its pristine story proof'}`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(({ story, mutate }) => {
      document.getElementById('viewer')!.innerHTML = `<div id="pagination-staging">
        <div id="pagination-hf-registry" style="display:none"><div data-section="0" data-hf-type="${story}-default">
          <p data-source-anchor-id="p:${story}:host" style="margin:0;font:10pt/14pt Arial">
          <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-clip="true"
            style="display:inline-block;width:150pt;height:28pt;overflow:hidden;vertical-align:top">
            <span data-source-anchor-id="p:${story}:visible" style="display:block;font:10pt/14pt Arial">Quail running light</span>
            <span style="display:block;height:20pt"></span>
            <span data-source-anchor-id="p:${story}:hidden" style="display:block;font:10pt/14pt Arial">Quail running shade</span>
          </span></p></div></div>
        <div data-section-index="0" data-page-width="300" data-page-height="360"
          data-content-width="228" data-content-height="252" data-margin-top="54" data-margin-right="36"
          data-margin-bottom="54" data-margin-left="36" data-header-distance="12" data-footer-distance="12">
          <p data-source-anchor-id="p:body:caption" style="margin:0;font:10pt/14pt Arial">Quail running story</p>
        </div></div><div id="pagination-container"></div>`;
      try {
        const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false });
        engine.paginate();
        if (mutate) document.querySelector<HTMLElement>(`#pagination-container [data-source-anchor-id="p:${story}:visible"]`)!.style.marginTop = '35pt';
        const map = engine.materializePageMap(0, 'quail-running-clip');
        return { error: null, anchors: map.fragments.map((f: any) => f.anchorId), warnings: engine.getSourceClippedTextAnchors() };
      } catch (error) { return { error: String(error), anchors: [], warnings: [] }; }
    }, { story, mutate });
    if (mutate) expect(result.error).toContain('has no measurable fragment');
    else {
      expect(result.error).toBeNull();
      expect(result.anchors).toContain(`p:${story}:visible`);
      expect(result.anchors).not.toContain(`p:${story}:hidden`);
      expect(result.warnings).toEqual([`p:${story}:hidden`]);
    }
  });
}

for (const kind of ['outside', 'column-paragraph', 'margin-paragraph', 'changed-paragraph', 'changed-margin-paragraph', 'unsupported-basis', 'straddling', 'ink-crosses-paper', 'fixed-ink-crosses-paper', 'autofit', 'linked', 'partial', 'changed-position', 'changed-font'] as const) {
  test(`source placement preserves the ${kind} paper boundary`, async ({ page }) => {
    await page.setContent('<div id="viewer"></div>');
    await page.addScriptTag({ path: bundle });
    const result = await page.evaluate(kind => {
      document.getElementById('viewer')!.innerHTML = `<div id="pagination-staging"><div data-section-index="0"
        data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
        data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
        <p data-source-anchor-id="p:body:host" style="font:10pt/14pt Arial;margin:0">Quail outer host
          <span data-docx-textbox="true" data-docx-textbox-clip="true" data-docx-textbox-source-position="true"
            ${kind === 'fixed-ink-crosses-paper' ? 'data-docx-textbox-source-clip="true"' : ''}
            data-docx-drawing-anchor="true" data-docx-anchor-h-relative="${['column-paragraph', 'changed-paragraph'].includes(kind) ? 'column' : ['margin-paragraph', 'changed-margin-paragraph'].includes(kind) ? 'margin' : kind === 'unsupported-basis' ? 'unknown' : 'page'}" data-docx-anchor-v-relative="${['column-paragraph', 'changed-paragraph', 'margin-paragraph', 'changed-margin-paragraph'].includes(kind) ? 'paragraph' : 'page'}"
            data-docx-anchor-h-offset="${kind === 'straddling' ? -40 : -170}" data-docx-anchor-v-offset="60"
            ${kind === 'autofit' ? 'data-docx-anchor-autofit="true"' : ''}
            ${kind === 'linked' ? 'data-docx-textbox-linked="true"' : ''}
            ${kind === 'partial' ? 'data-docx-shape="partial"' : ''}
            style="display:inline-block;width:100pt;height:28pt;overflow:hidden;vertical-align:top">
            <span data-source-anchor-id="p:body:hidden" style="display:block;font:10pt/14pt Arial;white-space:nowrap">${kind.includes('ink-crosses-paper') ? 'Quail unwound cord '.repeat(12) : 'Quail short cord'}</span>
          </span></p></div></div><div id="pagination-container"></div>`;
      try {
        const engine = new (window as any).DocxodusPagination.PaginationEngine('pagination-staging', 'pagination-container', { showPageNumbers: false });
        engine.paginate();
        const box = document.querySelector<HTMLElement>('#pagination-container [data-docx-textbox]')!;
        if (kind === 'changed-paragraph' || kind === 'changed-margin-paragraph') document.querySelector<HTMLElement>('#pagination-container [data-source-anchor-id="p:body:host"]')!.style.transform = 'translateY(10pt)';
        if (kind === 'changed-position') box.style.left = '-180pt';
        if (kind === 'changed-font') box.firstElementChild!.setAttribute('style', 'display:block;font:11pt/14pt Arial');
        const map = engine.materializePageMap(0, 'quail-paper-clip');
        return { error: null, anchors: map.fragments.map((f: any) => f.anchorId), warnings: engine.getSourceClippedTextAnchors() };
      } catch (error) { return { error: String(error), anchors: [], warnings: [] }; }
    }, kind);
    if (kind === 'outside' || kind === 'column-paragraph' || kind === 'margin-paragraph' || kind === 'fixed-ink-crosses-paper') {
      expect(result.error).toBeNull();
      expect(result.anchors).toContain('p:body:host');
      expect(result.anchors).not.toContain('p:body:hidden');
      expect(result.warnings).toEqual(['p:body:hidden']);
    } else if (kind === 'straddling') {
      expect(result.error).toBeNull();
      expect(result.anchors).toContain('p:body:hidden');
      expect(result.warnings).toEqual([]);
    } else expect(result.error).toContain('has no measurable fragment');
  });
}

for (const fit of ['auto', 'percent']) for (const change of [
  'right', 'bottom', 'visible', 'straddling', 'left', 'above', 'zoom', 'balloons',
  'offset-before', 'font-before', 'wrong-origin', 'move-after', 'font-after', 'text-after', 'removed', 'linked', 'partial',
]) test(`resizing page-offset proof binds ${fit} source geometry: ${change}`, async ({ page }) => {
  const left = change === 'left' ? -170 : change === 'straddling' ? 290 : ['bottom', 'above', 'visible'].includes(change) ? 36 : 320;
  const top = change === 'bottom' ? 400 : change === 'above' ? -200 : 20;
  await page.setContent(`<div id="pagination-staging"><div data-section-index="0"
    data-page-width="300" data-page-height="360" data-content-width="228" data-content-height="288"
    data-margin-top="36" data-margin-right="36" data-margin-bottom="36" data-margin-left="36">
    <p data-source-anchor-id="p:body:host" style="font:10pt/12pt Arial;margin:0">Birch host
      <span data-docx-textbox="true" data-docx-textbox-source-position="page-offset" data-docx-drawing-anchor="true"
        data-docx-anchor-h-relative="page" data-docx-anchor-v-relative="page"
        data-docx-anchor-h-offset="${left}" data-docx-anchor-v-offset="${top}"
        ${fit === 'auto' ? 'data-docx-anchor-autofit="true"' : 'data-docx-textbox-clip="true" data-docx-anchor-width-percent="8" data-docx-anchor-width-relative="page"'}
        ${change === 'linked' ? 'data-docx-textbox-linked="true"' : change === 'partial' ? 'data-docx-shape="partial"' : ''}
        style="display:inline-block;position:absolute;width:160pt;${fit === 'auto' ? '' : 'height:180pt;'}overflow:hidden">
        <span data-source-anchor-id="p:body:resized" style="display:block;font:10pt/12pt Arial">Birch amber red green blue<br>Birch amber red green blue<br>Birch amber red green blue</span>
      </span></p><p data-source-anchor-id="p:body:closing" data-format-description="Font: Bold" data-format-key="1"
        style="font:10pt/12pt Arial;margin:0">Birch closing</p>
    </div></div><div id="pagination-container"></div>`);
  await page.addScriptTag({ path: bundle });
  const result = await page.evaluate(({ fit, change }) => {
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
        const box = args[1].querySelector('[data-docx-textbox]');
        if (box) {
          if (change === 'offset-before') box.dataset.docxAnchorHOffset = '350';
          else box.firstElementChild.style.fontSize = '15pt';
        }
        return position(...args);
      };
    }
    try {
      engine.paginate();
      const root = document.getElementById('pagination-container')!;
      const box = root.querySelector<HTMLElement>('[data-docx-textbox]')!;
      const text = box.firstElementChild as HTMLElement;
      if (change === 'move-after') box.style.left = '350pt';
      if (change === 'font-after') text.style.fontSize = '15pt';
      if (change === 'text-after') text.textContent = 'Birch replacement';
      if (change === 'removed') text.remove();
      const map = engine.materializePageMap(0, 'resizing-source');
      return { error: '', clipped: engine.getSourceClippedTextAnchors(),
        mapped: map.fragments.some((f: any) => f.anchorId === 'p:body:resized'),
        resized: fit === 'auto' ? parseFloat(getComputedStyle(box).height) > 40 : box.style.width === '24pt' };
    } catch (error) { return { error: String(error), clipped: [], mapped: false, resized: false }; }
  }, { fit, change });
  if (['right', 'bottom', 'zoom', 'balloons'].includes(change)) {
    expect(result.error).toBe('');
    expect(result.clipped).toEqual(['p:body:resized']);
    expect(result.resized).toBe(true);
    expect(result.mapped).toBe(false);
  } else if (['visible', 'straddling'].includes(change)) {
    expect(result.clipped).toEqual([]);
    if (result.error) expect(result.error).toContain('has no measurable fragment');
    else expect(result.mapped).toBe(true);
  } else expect(result.error).toContain('has no measurable fragment');
});

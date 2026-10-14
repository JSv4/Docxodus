import { test, expect, Page } from '@playwright/test';
import { imageOnlyDelta, patchImageAttributes } from '../src/editor-image-patch';

// The reconciler's image-only patch, exercised on real DOM in a bare page — no WASM,
// no editor. The two helpers are written without free variables precisely so they
// can be injected here by source; the same functions ship in the editor bundle.
//
// The property under test: `imageOnlyDelta` returns the <img> pairs when, and only
// when, the fresh render is the live node with different image attributes (editor
// stamps aside). Anything else — a text edit, a run attribute, a lost child — must
// return null so the reconciler keeps swapping nodes exactly as it did before.

async function installHelpers(page: Page) {
  await page.setContent('<div id="host"></div>');
  await page.addScriptTag({
    content:
      `window.__imageOnlyDelta = ${imageOnlyDelta.toString()};\n` +
      `window.__patchImageAttributes = ${patchImageAttributes.toString()};`,
  });
}

/** Parse `live` and `fresh` HTML into elements and run imageOnlyDelta over them.
 *  Returns the pair count, or null. */
function delta(page: Page, live: string, fresh: string) {
  return page.evaluate(
    ({ live, fresh }) => {
      const parse = (html: string) =>
        new DOMParser().parseFromString(html, 'text/html').body.firstElementChild!;
      const pairs = (window as any).__imageOnlyDelta(parse(live), parse(fresh));
      return pairs === null ? null : pairs.length;
    },
    { live, fresh },
  );
}

const PIX_A = 'data:image/png;base64,AAAA';
const PIX_B = 'data:image/png;base64,BBBB';

test.describe('editor-image-patch — imageOnlyDelta', () => {
  test.beforeEach(async ({ page }) => installHelpers(page));

  test('identical trees are an empty delta, not null', async ({ page }) => {
    const html = `<p data-anchor="a1" class="x"><span>text</span></p>`;
    expect(await delta(page, html, html)).toBe(0);
  });

  test('a changed image source pairs the two <img> elements', async ({ page }) => {
    const live = `<p data-anchor="a1"><img src="${PIX_A}" style="width: 10pt; height: 5pt" alt="frame"></p>`;
    const fresh = `<p data-anchor="a1"><img src="${PIX_B}" style="width: 10pt; height: 5pt" alt="frame"></p>`;
    expect(await delta(page, live, fresh)).toBe(1);
  });

  test('image dimension and alt changes are still image-only', async ({ page }) => {
    const live = `<p data-anchor="a1"><img src="${PIX_A}" style="width: 10pt; height: 5pt"></p>`;
    const fresh = `<p data-anchor="a1"><img src="${PIX_A}" style="width: 20pt; height: 9pt" alt="bigger"></p>`;
    expect(await delta(page, live, fresh)).toBe(1);
  });

  test('images nested inside runs and links are found', async ({ page }) => {
    const live = `<p data-anchor="a1"><span class="r"><a href="#x"><img src="${PIX_A}"></a></span><span class="r"><img src="${PIX_A}"></span></p>`;
    const fresh = `<p data-anchor="a1"><span class="r"><a href="#x"><img src="${PIX_B}"></a></span><span class="r"><img src="${PIX_B}"></span></p>`;
    expect(await delta(page, live, fresh)).toBe(2);
  });

  test('editor-owned stamps on the live node are not a difference', async ({ page }) => {
    const live =
      `<p data-anchor="a1" contenteditable="true" data-committed-text="" data-render-sig="s1">` +
      `<span class="r"><img src="${PIX_A}"></span></p>`;
    const fresh = `<p data-anchor="a1"><span class="r"><img src="${PIX_B}"></span></p>`;
    expect(await delta(page, live, fresh)).toBe(1);
  });

  test('the injected list-marker separator on the live side is skipped', async ({ page }) => {
    const live =
      `<p data-anchor="a1"><span data-list-marker="true">1.</span>` +
      `<span data-editor-list-separator="" data-list-marker="true" aria-hidden="true"> </span>` +
      `<span class="r"><img src="${PIX_A}"></span></p>`;
    const fresh =
      `<p data-anchor="a1"><span data-list-marker="true">1.</span><span class="r"><img src="${PIX_B}"></span></p>`;
    expect(await delta(page, live, fresh)).toBe(1);
  });

  test('a text change is not image-only', async ({ page }) => {
    const live = `<p data-anchor="a1"><span>old</span><img src="${PIX_A}"></p>`;
    const fresh = `<p data-anchor="a1"><span>new</span><img src="${PIX_B}"></p>`;
    expect(await delta(page, live, fresh)).toBeNull();
  });

  test('a changed run attribute is not image-only', async ({ page }) => {
    const live = `<p data-anchor="a1"><span style="color: red"><img src="${PIX_A}"></span></p>`;
    const fresh = `<p data-anchor="a1"><span style="color: blue"><img src="${PIX_B}"></span></p>`;
    expect(await delta(page, live, fresh)).toBeNull();
  });

  test('a changed block attribute (its unid, a class) is not image-only', async ({ page }) => {
    const live = `<p data-anchor="a1" class="para"><img src="${PIX_A}"></p>`;
    expect(await delta(page, live, `<p data-anchor="a2" class="para"><img src="${PIX_B}"></p>`)).toBeNull();
    expect(await delta(page, live, `<p data-anchor="a1" class="para centred"><img src="${PIX_B}"></p>`)).toBeNull();
  });

  test('a gained or lost child is not image-only', async ({ page }) => {
    const live = `<p data-anchor="a1"><img src="${PIX_A}"></p>`;
    expect(await delta(page, live, `<p data-anchor="a1"><img src="${PIX_B}"><span>caption</span></p>`)).toBeNull();
    expect(await delta(page, `<p data-anchor="a1"><img src="${PIX_A}"><span>c</span></p>`, live)).toBeNull();
  });

  test('an image replaced by text, or a leaf render against a wrapper, is not image-only', async ({ page }) => {
    expect(await delta(page, `<p data-anchor="a1"><img src="${PIX_A}"></p>`, `<p data-anchor="a1">gone</p>`)).toBeNull();
    expect(await delta(page, `<p data-anchor="a1"><img src="${PIX_A}"></p>`, `<div><p data-anchor="a1"><img src="${PIX_B}"></p></div>`)).toBeNull();
  });

  test('a wrapper-shaped pair whose inner block differs only in an image is image-only', async ({ page }) => {
    const live = `<div class="border-group"><p data-anchor="a1" contenteditable="true"><img src="${PIX_A}"></p></div>`;
    const fresh = `<div class="border-group"><p data-anchor="a1"><img src="${PIX_B}"></p></div>`;
    expect(await delta(page, live, fresh)).toBe(1);
  });
});

test.describe('editor-image-patch — patchImageAttributes', () => {
  test.beforeEach(async ({ page }) => installHelpers(page));

  test('the live element takes the fresh attributes and keeps its identity and editor stamps', async ({ page }) => {
    const result = await page.evaluate(({ a, b }) => {
      const host = document.getElementById('host')!;
      host.innerHTML = `<p><img id="live" src="${a}" style="width: 10pt; height: 5pt" title="stale" contenteditable="false"></p>`;
      const live = host.querySelector('img')!;
      (live as any).__sentinel = 'same node';
      const fresh = new DOMParser().parseFromString(
        `<img src="${b}" style="width: 20pt; height: 9pt" alt="frame 2">`, 'text/html',
      ).body.firstElementChild as HTMLImageElement;
      const order: string[] = [];
      const observer = new MutationObserver((muts) => {
        for (const m of muts) order.push(m.attributeName ?? '');
      });
      observer.observe(live, { attributes: true });
      (window as any).__patchImageAttributes(live, fresh);
      const flushed = observer.takeRecords().map((m) => m.attributeName ?? '');
      observer.disconnect();
      return {
        sentinel: (host.querySelector('img') as any).__sentinel,
        src: live.getAttribute('src'),
        style: live.getAttribute('style'),
        alt: live.getAttribute('alt'),
        title: live.getAttribute('title'),
        contenteditable: live.getAttribute('contenteditable'),
        id: live.getAttribute('id'),
        order: order.concat(flushed),
      };
    }, { a: PIX_A, b: PIX_B });
    expect(result.sentinel).toBe('same node');
    expect(result.src).toBe(PIX_B);
    expect(result.style).toBe('width: 20pt; height: 9pt');
    expect(result.alt).toBe('frame 2');
    expect(result.title).toBeNull(); // stale attribute removed
    expect(result.id).toBeNull(); // not editor-owned, not in the fresh render: removed
    expect(result.contenteditable).toBe('false'); // editor-owned: kept
    // Dimensions land before the new source starts loading.
    expect(result.order.indexOf('style')).toBeLessThan(result.order.indexOf('src'));
    expect(result.order.filter((n) => n === 'src')).toHaveLength(1);
  });

  test('an unchanged source is not re-set', async ({ page }) => {
    const touched = await page.evaluate((a) => {
      const host = document.getElementById('host')!;
      host.innerHTML = `<p><img src="${a}"></p>`;
      const live = host.querySelector('img')!;
      const fresh = new DOMParser().parseFromString(`<img src="${a}" alt="named">`, 'text/html')
        .body.firstElementChild as HTMLImageElement;
      const observer = new MutationObserver(() => {});
      observer.observe(live, { attributes: true });
      (window as any).__patchImageAttributes(live, fresh);
      const names = observer.takeRecords().map((m) => m.attributeName);
      observer.disconnect();
      return names;
    }, PIX_A);
    expect(touched).toEqual(['alt']);
  });
});

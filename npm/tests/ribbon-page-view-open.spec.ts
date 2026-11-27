import { test, expect, type Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const TEST_FILES_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../TestFiles');

/**
 * Opening a document while the ribbon is in page view (issue #1868).
 *
 * The ribbon renders a document it opens into a stand-in for its surface and only then retires
 * the current one, so that a file it cannot open leaves the current document, its selection and
 * its undo history alone. Page view has to measure what it lays out, and the stand-in was never
 * in the document: every open in page view threw "Pagination requires staging and container
 * elements attached to their document", and a ribbon created in page view could not start.
 *
 * The oracle for the pages an open lays out is the same document switched to page view after it
 * opened in continuous view, which paginates in the live surface.
 */

async function openDemo(page: Page) {
  await page.addInitScript(() => {
    // How one ribbon action ended: what it threw, and whether the ribbon now shows another editor.
    (window as any).__attempt = async (action: (ribbon: any) => unknown) => {
      const ribbon = (window as any).__ribbon;
      const before = ribbon.editor;
      let error = '';
      try { await action(ribbon); } catch (thrown) { error = String(thrown); }
      return { error, replaced: ribbon.editor !== before };
    };
  });
  await page.goto('/demo-app.html?engine=./embed.bundle.js&history=0');
  await page.waitForFunction(() => !!(window as any).__ribbon?.editor, { timeout: 90000 });
}

/** The page view as the reader sees it, plus anything an open could leave behind in the page. */
async function pageView(page: Page) {
  return page.evaluate(() => {
    const ribbon = (window as any).__ribbon;
    const surface = ribbon.surface as HTMLElement;
    const pages = Array.from(surface.querySelectorAll<HTMLElement>('.page-box'));
    const anchors = Array.from(document.querySelectorAll('[data-anchor]'), (el) => el.getAttribute('data-anchor'));
    return {
      view: surface.dataset.view,
      text: pages.map((box) => (box.textContent || '').replace(/\s+/g, ' ').trim()),
      surfaces: document.querySelectorAll('[data-dxr-surface]').length,
      pageContainers: document.querySelectorAll('#pagination-container').length,
      staging: document.querySelectorAll('#pagination-staging, .page-staging').length,
      repeatedAnchors: anchors.length - new Set(anchors).size,
      firstPageWidth: pages[0]?.getBoundingClientRect().width ?? 0,
      zoom: ribbon.editor.zoom as number,
      disabledSheets: surface.querySelectorAll('style[media="not all"], link[media="not all"]').length,
    };
  });
}

const CLEAN = { surfaces: 1, pageContainers: 1, staging: 0, repeatedAnchors: 0, disabledSheets: 0 };

/** Retype the first editable paragraph and commit it, as a user leaving the paragraph would. */
async function editFirstBlock(page: Page, text: string) {
  return page.evaluate((text) => {
    const ribbon = (window as any).__ribbon;
    const block = ribbon.surface.querySelector('[data-anchor][contenteditable="true"]') as HTMLElement;
    const original = block.textContent;
    block.focus();
    block.textContent = text;
    block.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' }));
    block.blur();
    return { anchor: block.getAttribute('data-anchor')!, original, session: ribbon.editor.sessionHandle as number };
  }, text);
}

/** Undo once, reporting the paragraph's text before and after. */
async function undo(page: Page, anchor: string) {
  return page.evaluate((anchor) => {
    const ribbon = (window as any).__ribbon;
    const text = () => ribbon.surface.querySelector(`[data-anchor="${anchor}"]`)?.textContent;
    const draft = text();
    ribbon.editor.undo();
    return { session: ribbon.editor.sessionHandle as number, draft, undone: text() };
  }, anchor);
}

test.describe('Ribbon — opening a document in page view (issue #1868)', () => {
  test('a new document and a file open in page view, laid out as page view lays them out', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    await openDemo(page);
    await page.evaluate(() => {
      const ribbon = (window as any).__ribbon;
      (window as any).__guide = ribbon.editor.save();
      ribbon.control('viewpage').click();
    });
    const switched = await pageView(page);
    expect(switched.view).toBe('paginated');
    expect(switched.text.length).toBeGreaterThan(1);
    expect(switched).toMatchObject(CLEAN);

    expect(await page.evaluate(() => (window as any).__attempt((ribbon: any) => ribbon.openBlank())))
      .toEqual({ error: '', replaced: true });
    const blank = await pageView(page);
    expect(blank.view).toBe('paginated');
    expect(blank.text.length).toBe(1);
    expect(blank).toMatchObject(CLEAN);

    expect(await page.evaluate(() => (window as any).__attempt(
      (ribbon: any) => ribbon.open((window as any).__guide, 'guide.docx'),
    ))).toEqual({ error: '', replaced: true });
    const opened = await pageView(page);
    expect(opened.view).toBe('paginated');
    // Page for page what switching to page view lays out.
    expect(opened.text).toEqual(switched.text);
    expect(opened.firstPageWidth).toBeCloseTo(switched.firstPageWidth, 0);
    // One surface, one page stack, one copy of each block: nothing of the stand-in or of the
    // retired document is left in the page.
    expect(opened).toMatchObject(CLEAN);

    // The opened document is the editor's, in the ribbon's own surface, and it edits and undoes.
    expect(await page.evaluate(() => (window as any).__ribbon.editor.root === (window as any).__ribbon.surface))
      .toBe(true);
    const edit = await editFirstBlock(page, 'Edited after the open');
    expect(await undo(page, edit.anchor))
      .toEqual({ session: edit.session, draft: 'Edited after the open', undone: edit.original });
    expect(errors).toEqual([]);
  });

  test('a long document opened in page view gets the pages switching to page view lays out', async ({ page }) => {
    test.setTimeout(180000);
    await openDemo(page);
    // The charter's page breaks follow from its line metrics, so they move if the new document is
    // measured anywhere but under the rules that style the ribbon's surface: laid out at the end
    // of <body>, outside the embed's scoped stylesheet, it comes to 53 pages instead of 52.
    const bytes = Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, 'NVCA-Model-COI.docx')));
    const result = await page.evaluate((array) => {
      const ribbon = (window as any).__ribbon;
      const charter = new Uint8Array(array);
      const pages = () => Array.from(ribbon.surface.querySelectorAll('.page-box') as NodeListOf<HTMLElement>,
        (box) => (box.textContent || '').replace(/\s+/g, ' ').trim());
      ribbon.open(charter, 'charter.docx');
      ribbon.control('viewpage').click();
      const switched = pages();
      ribbon.open(charter, 'charter.docx');
      return { switched, opened: pages() };
    }, bytes);
    expect(result.switched.length).toBeGreaterThan(40);
    expect(result.opened).toEqual(result.switched);
  });

  test('continuous view still opens a document', async ({ page }) => {
    await openDemo(page);
    expect(await page.evaluate(() => (window as any).__attempt(
      (ribbon: any) => ribbon.open(ribbon.editor.save(), 'guide.docx'),
    ))).toEqual({ error: '', replaced: true });
    expect(await pageView(page)).toMatchObject({
      view: 'continuous', text: [], surfaces: 1, pageContainers: 0, repeatedAnchors: 0,
    });
    expect(await page.evaluate(() => (window as any).__ribbon.editor.root === (window as any).__ribbon.surface))
      .toBe(true);
  });

  test('zoom, fit-to-width and the header/footer toggle work on a document opened in page view', async ({ page }) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(String(error)));
    await openDemo(page);
    await page.evaluate(() => (window as any).__ribbon.control('viewpage').click());
    const switched = await pageView(page);

    expect(await page.evaluate(() => (window as any).__attempt(
      (ribbon: any) => ribbon.open(ribbon.editor.save(), 'guide.docx'),
    ))).toEqual({ error: '', replaced: true });
    const opened = await pageView(page);
    expect(opened.zoom).toBe(1);

    await page.evaluate(() => {
      const zoom = (window as any).__ribbon.control('zoom') as HTMLSelectElement;
      zoom.value = '0.5';
      zoom.dispatchEvent(new Event('change'));
    });
    const halved = await pageView(page);
    expect(halved.zoom).toBe(0.5);
    expect(halved.firstPageWidth / opened.firstPageWidth).toBeCloseTo(0.5, 2);

    // A window narrower than the page zooms it to fit instead of clipping it.
    await page.evaluate(() => (window as any).__ribbon.control('zoomfit').click());
    await page.setViewportSize({ width: 520, height: 900 });
    await expect.poll(() => page.evaluate(() => {
      const scroll = document.querySelector<HTMLElement>('.dxr-scroll')!.getBoundingClientRect();
      const box = document.querySelector<HTMLElement>('[data-dxr-surface] .page-box')!.getBoundingClientRect();
      return box.left >= scroll.left && box.right <= scroll.right;
    })).toBe(true);
    expect((await pageView(page)).zoom).toBeLessThan(1);

    // The header/footer toggle re-opens the live document, so it goes through the same open.
    expect(await page.evaluate(() => (window as any).__attempt(
      (ribbon: any) => ribbon.control('headerfooter').click(),
    ))).toEqual({ error: '', replaced: true });
    const withBands = await pageView(page);
    expect(withBands.view).toBe('paginated');
    expect(withBands.text.length).toBe(switched.text.length);
    expect(withBands).toMatchObject(CLEAN);
    expect(errors).toEqual([]);
  });

  test('a ribbon created in page view starts in page view', async ({ page }) => {
    await openDemo(page);
    const started = await page.evaluate(async () => {
      const { createRibbonEditor } = await import('./embed.bundle.js' as string);
      const host = document.createElement('div');
      host.style.cssText = 'position:fixed;inset:0;z-index:10;background:#fff';
      document.body.append(host);
      try {
        const ribbon = await createRibbonEditor(host, null, { paginated: true, loader: false });
        return { error: '', pages: ribbon.surface.querySelectorAll('.page-box').length, view: ribbon.surface.dataset.view };
      } catch (error) {
        return { error: String(error), pages: 0, view: '' };
      }
    });
    expect(started).toEqual({ error: '', pages: 1, view: 'paginated' });
  });

  // Laid out beside the live document, the new one is measured under its own styles alone: with fabricated
  // classes two documents share class names, and the live one's rules would move the new one's page breaks.
  test('a document opened in page view over another is laid out under its own styles', async ({ page }) => {
    await openDemo(page);
    const read = (name: string) => Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, name)));
    const result = await page.evaluate(async ([live, opened]) => {
      const { createRibbonEditor } = await import('./embed.bundle.js' as string);
      const host = document.createElement('div');
      host.style.cssText = 'position:fixed;inset:0;z-index:10;background:#fff';
      document.body.append(host);
      const ribbon = await createRibbonEditor(host, new Uint8Array(opened), { fabricateClasses: true, loader: false });
      const pages = () => Array.from(ribbon.surface.querySelectorAll('.page-box') as NodeListOf<HTMLElement>,
        (box) => (box.textContent || '').replace(/\s+/g, ' ').trim());
      ribbon.control('viewpage').click();
      const switched = pages();
      ribbon.open(new Uint8Array(live), 'live.docx');
      ribbon.open(new Uint8Array(opened), 'opened.docx');
      return { switched, overAnother: pages() };
    }, [read('HC031-Complicated-Document.docx'), read('NVCA-Model-COI.docx')]);
    expect(result.switched.length).toBeGreaterThan(40);
    expect(result.overAnother).toEqual(result.switched);
  });

  // Re-opening in page view keeps the reader's place, even at the very end of a long document, while the live
  // document's stylesheets are off and its pages briefly shorter.
  test('the header/footer toggle in page view keeps the scroll position at the end of a long document', async ({ page }) => {
    await openDemo(page);
    const charter = Array.from(fs.readFileSync(path.join(TEST_FILES_DIR, 'NVCA-Model-COI.docx')));
    const result = await page.evaluate(async (bytes) => {
      const { createRibbonEditor } = await import('./embed.bundle.js' as string);
      const host = document.createElement('div');
      host.style.cssText = 'position:fixed;inset:0;z-index:10;background:#fff';
      document.body.append(host);
      const ribbon = await createRibbonEditor(host, new Uint8Array(bytes), { paginated: true, comments: false, loader: false });
      const scroller = host.querySelector<HTMLElement>('.dxr-scroll')!;
      scroller.scrollTop = scroller.scrollHeight;
      const before = scroller.scrollTop;
      ribbon.control('headerfooter').click();
      return { before, after: scroller.scrollTop };
    }, charter);
    expect(result.before).toBeGreaterThan(10000);
    expect(result.after).toBe(result.before);
  });

  const failures: Array<[string, () => Promise<unknown>]> = [
    ['a file that is not a document', () => (window as any).__attempt(
      (ribbon: any) => ribbon.open(new Uint8Array([1, 2, 3]), 'broken.docx'),
    )],
    // The paginator can refuse a document after the stand-in already holds its markup. This
    // render has no page view in it, and an unscoped stylesheet that would reflow the current
    // document if the stand-in stayed in the page.
    ['a document that fails while it is laid out', () => (window as any).__attempt(async (ribbon: any) => {
      const wasm = (await import('./embed.bundle.js' as string)).getWasmExports();
      ribbon.setExports({
        DocumentConverter: wasm.DocumentConverter,
        DocxSessionBridge: {
          ...wasm.DocxSessionBridge,
          RenderEditorHtml: () => '<style>p { margin: 300px 0 }</style><p>No page view here</p>',
        },
      });
      ribbon.open(ribbon.editor.save(), 'broken.docx');
    })],
  ];
  for (const [name, failure] of failures) {
    test(`${name} leaves the document open in page view, with its undo history`, async ({ page }) => {
      await openDemo(page);
      await page.evaluate(() => (window as any).__ribbon.control('viewpage').click());

      const edit = await editFirstBlock(page, 'Undo survives a rejected open');
      const edited = await pageView(page);
      expect(edited.text.join(' ')).toContain('Undo survives a rejected open');

      const result = await page.evaluate(failure) as { error: string; replaced: boolean };
      expect(result.error).not.toBe('');
      expect(result.replaced).toBe(false);
      expect(await pageView(page)).toEqual(edited);
      expect(await undo(page, edit.anchor))
        .toEqual({ session: edit.session, draft: 'Undo survives a rejected open', undone: edit.original });
    });
  }
});

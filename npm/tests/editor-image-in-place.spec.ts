import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(__dirname, '..', '..', 'TestFiles', relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Replacing an image's media through the session and calling `editor.refresh()` must
// repaint by patching the <img> that is already on screen, never by swapping the
// paragraph node. Firefox and WebKit paint a freshly inserted <img> as an empty box
// until its data URI is decoded — the arcade's Doom cartridge, which replaces its
// frame image at frame rate, strobed white on every frame there. An in-place `src`
// change keeps the previous bitmap up until the new one is ready.
//
// The proof is DOM identity: a sentinel property set on the live <img> (and on its
// paragraph) must survive the refresh while `src` moves to the new media. The last
// test pins the boundary: a change that is NOT image-only still swaps the node.
test.describe('DocxEditor — image replacement patches the live <img> in place', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  /** Fresh editor over a blank doc plus a typed session over the same handle; inserts
   *  `first` as an inline image in the first paragraph and paints it. */
  async function openWithImage(page: Page, first: Uint8Array) {
    return page.evaluate((firstArray: number[]) => {
      const D = (window as any).Docxodus;
      const container = document.createElement('div');
      container.id = 'image-host';
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, D.DocxSessionBridge.CreateBlankDocx(), D, {});
      const session = new D.DocxSession(editor.sessionHandle, D.DocxSessionBridge);
      const anchor = JSON.parse(D.DocxSessionBridge.FindByKind(editor.sessionHandle, 'p', 'body'))[0].id as string;
      const inserted = session.insertImage(anchor, 0, new Uint8Array(firstArray), {
        widthPoints: 120,
        heightPoints: 80,
        preserveAspect: false,
        altText: 'frame',
      });
      if (!inserted.success) throw new Error(`insertImage: ${inserted.error?.code} ${inserted.error?.message}`);
      editor.refresh();
      const img = container.querySelector('img') as HTMLImageElement | null;
      if (!img) throw new Error('the refreshed paragraph carries no <img>');
      (img as any).__sentinel = 'live image';
      (img.closest('[data-anchor]') as any).__sentinel = 'live paragraph';
      (window as any).__img = { editor, session, container, imageId: inserted.imageId as string };
      return { src: img.getAttribute('src'), anchor: img.closest('[data-anchor]')!.getAttribute('data-anchor') };
    }, Array.from(first));
  }

  test('replaceImage + refresh keeps the <img> node and moves its src', async ({ page }) => {
    const first = readTestFile('img.png');
    const second = readTestFile('img2.png');
    const before = await openWithImage(page, first);
    expect(before.src).toMatch(/^data:image\/png;base64,/);

    const after = await page.evaluate((secondArray: number[]) => {
      const { editor, session, container, imageId } = (window as any).__img;
      const replaced = session.replaceImage(imageId, new Uint8Array(secondArray));
      if (!replaced.success) throw new Error(`replaceImage: ${replaced.error?.code} ${replaced.error?.message}`);
      editor.refresh();
      const img = container.querySelector('img') as HTMLImageElement;
      const block = img.closest('[data-anchor]') as HTMLElement;
      return {
        imgSentinel: (img as any).__sentinel ?? null,
        blockSentinel: (block as any).__sentinel ?? null,
        src: img.getAttribute('src'),
        fallback: editor.lastReconcileFallback ?? null,
        images: container.querySelectorAll('img').length,
        contenteditable: block.getAttribute('contenteditable'),
      };
    }, Array.from(second));

    expect(after.fallback).toBeNull(); // still an incremental repaint
    expect(after.images).toBe(1);
    expect(after.imgSentinel).toBe('live image'); // same element, not a swap
    expect(after.blockSentinel).toBe('live paragraph');
    expect(after.src).toMatch(/^data:image\/png;base64,/);
    expect(after.src).not.toBe(before.src); // and it really shows the new media
    expect(after.contenteditable).toBe('true'); // the wired node is the one kept
  });

  test('the patched block is stamped, so the next refresh is a no-op for it', async ({ page }) => {
    const first = readTestFile('img.png');
    const second = readTestFile('img2.png');
    await openWithImage(page, first);
    const result = await page.evaluate((secondArray: number[]) => {
      const { editor, session, container, imageId } = (window as any).__img;
      session.replaceImage(imageId, new Uint8Array(secondArray));
      editor.refresh();
      const srcAfterReplace = container.querySelector('img')!.getAttribute('src');
      const attributeWrites: string[] = [];
      const observer = new MutationObserver((muts) => {
        for (const m of muts) attributeWrites.push(`${(m.target as Element).tagName}.${m.attributeName}`);
      });
      observer.observe(container, { attributes: true, subtree: true, childList: true });
      editor.refresh(); // nothing changed in the session
      const records = observer.takeRecords();
      observer.disconnect();
      const img = container.querySelector('img') as HTMLImageElement;
      return {
        srcStable: img.getAttribute('src') === srcAfterReplace,
        imgSentinel: (img as any).__sentinel ?? null,
        childListChanges: records.filter((m) => m.type === 'childList').length,
        imgAttributeWrites: attributeWrites.filter((w) => w.startsWith('IMG.')),
        fallback: editor.lastReconcileFallback ?? null,
      };
    }, Array.from(second));
    expect(result.fallback).toBeNull();
    expect(result.srcStable).toBe(true);
    expect(result.imgSentinel).toBe('live image');
    expect(result.childListChanges).toBe(0);
    expect(result.imgAttributeWrites).toEqual([]);
  });

  test('a frame-rate loop of replacements never leaves the paragraph without its image', async ({ page }) => {
    // The arcade's shape: many replaceImage + refresh rounds. Every round must find the
    // same <img> in place — a swap would show up as a lost sentinel on some round.
    const frames = [readTestFile('img.png'), readTestFile('img2.png')];
    await openWithImage(page, frames[0]);
    const result = await page.evaluate((frameArrays: number[][]) => {
      const { editor, session, container, imageId } = (window as any).__img;
      let lostIdentity = 0;
      let missing = 0;
      const sources = new Set<string>();
      for (let i = 0; i < 20; i++) {
        session.replaceImage(imageId, new Uint8Array(frameArrays[(i + 1) % 2]));
        editor.refresh();
        const img = container.querySelector('img') as HTMLImageElement | null;
        if (!img) { missing++; continue; }
        if ((img as any).__sentinel !== 'live image') lostIdentity++;
        sources.add(img.getAttribute('src')!);
      }
      return { lostIdentity, missing, distinctSources: sources.size, fallback: editor.lastReconcileFallback ?? null };
    }, frames.map((f) => Array.from(f)));
    expect(result.missing).toBe(0);
    expect(result.lostIdentity).toBe(0);
    expect(result.distinctSources).toBe(2); // the two media alternate on the one element
    expect(result.fallback).toBeNull();
  });

  test('a change that is not image-only still swaps the node', async ({ page }) => {
    const first = readTestFile('img.png');
    await openWithImage(page, first);
    const result = await page.evaluate(() => {
      const { editor, session, container } = (window as any).__img;
      const anchor = container.querySelector('img')!.closest('[data-anchor]')!;
      const id = Object.keys(session.project().anchorIndex).find((k: string) =>
        k.endsWith(':' + anchor.getAttribute('data-anchor')));
      // Centre the paragraph: the block's own attributes change, not just the image's.
      const res = session.setParagraphFormat(id, { alignment: 'center' });
      if (!res.success) throw new Error(`setParagraphFormat: ${res.error?.code} ${res.error?.message}`);
      editor.refresh();
      const img = container.querySelector('img') as HTMLImageElement | null;
      const block = img?.closest('[data-anchor]') as HTMLElement | null;
      return {
        hasImage: img !== null,
        textAlign: block ? getComputedStyle(block).textAlign : null,
        imgSentinel: (img as any)?.__sentinel ?? null,
        fallback: editor.lastReconcileFallback ?? null,
      };
    });
    expect(result.hasImage).toBe(true);
    expect(result.textAlign).toBe('center');
    expect(result.imgSentinel).toBeNull(); // swapped, as before
    expect(result.fallback).toBeNull(); // and still incrementally
  });
});

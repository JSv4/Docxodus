import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): Uint8Array {
  return new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath)));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue #453 — native image ops through the BROWSER transport.
//
// The .NET, MCP, and Python surfaces each have their own image tests. This file covers the one
// hop none of them touch: `Uint8Array` -> `imageBytesToBase64` (a chunked `String.fromCharCode`
// loop over 0x8000-byte windows, then `globalThis.btoa`) -> the `[JSExport]` string parameter.
// That chunking exists specifically because spreading a large array into `fromCharCode` blows
// the argument limit, so at least one fixture here MUST exceed one 32768-byte chunk or the loop
// is never exercised past its first iteration.
//
// These go through `window.Docxodus.openTypedSession(...)`, which returns the real bundled
// `DocxSession` wrapper, so the encoder under test is the shipped one — not a re-implementation.
test.describe('DocxSession native images (typed wrapper — Issue #453)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('getImageCapabilities reports the browser runtime and the writable format set', async ({
    page,
  }) => {
    const docx = readTestFile('HC006-Test-01.docx');
    const capabilities = await page.evaluate(async (docxArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
      try {
        return session.getImageCapabilities();
      } finally {
        session.close();
      }
    }, Array.from(docx));

    expect(capabilities.runtime).toBe('browser-wasm');
    expect(capabilities.schemaVersion).toBe(1);
    expect(capabilities.acceptsBinaryBytes).toBe(true);
    // A browser build must never claim it can reach the network or the filesystem.
    expect(capabilities.supportsNetworkFetch).toBe(false);
    expect(capabilities.supportsFileIo).toBe(false);
    expect(capabilities.usesHeaderParsingOnly).toBe(true);

    const png = capabilities.formats.find((f: any) => f.format === 'png');
    expect(png.canInsert).toBe(true);
    expect(png.canReplace).toBe(true);
    const webp = capabilities.formats.find((f: any) => f.format === 'webp');
    expect(webp.canInspect).toBe(true);
    expect(webp.canInsert).toBe(false);
  });

  test('insertImage sends a real multi-chunk PNG across the base64 boundary intact', async ({
    page,
  }) => {
    // 15128 bytes — a genuinely decodable 180x174 PNG, not a header stub. The dimensions come
    // back out of the *package* media part, so a mangled encode cannot produce them by accident.
    const docx = readTestFile('HC006-Test-01.docx');
    const png = readTestFile('img.png');
    expect(png.length).toBeGreaterThan(15000);

    const result = await page.evaluate(
      async ([docxArray, pngArray]: number[][]) => {
        const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
        try {
          const proj = session.project();
          const anchor = Object.keys(proj.anchorIndex).find((k: string) =>
            k.startsWith('p:body:'),
          )!;

          const inserted = session.insertImage(anchor, 0, new Uint8Array(pngArray), {
            altText: 'browser transport',
            widthPoints: 90,
          });
          const images = session.listImages();

          // Reopen the saved bytes: proves the media survived the whole round trip, not just
          // that the in-memory session reported success.
          const saved = session.save();
          const reopened = (window as any).Docxodus.openTypedSession(saved);
          let reopenedImages: any[] = [];
          try {
            reopenedImages = reopened.listImages();
          } finally {
            reopened.close();
          }

          return {
            success: inserted.success,
            errorCode: inserted.error?.code,
            imageId: inserted.imageId,
            images,
            reopenedImages,
          };
        } finally {
          session.close();
        }
      },
      [Array.from(docx), Array.from(png)],
    );

    expect(result.errorCode).toBeUndefined();
    expect(result.success).toBe(true);
    expect(result.images).toHaveLength(1);

    const image = result.images[0];
    expect(image.id).toBe(result.imageId);
    expect(image.format).toBe('png');
    expect(image.contentType).toBe('image/png');
    // The bytes crossed the boundary byte-exact, or the header parser could not read these.
    expect(image.intrinsicWidthPixels).toBe(180);
    expect(image.intrinsicHeightPixels).toBe(174);
    expect(image.contentTypeMatchesBytes).toBe(true);
    expect(image.isBroken).toBe(false);
    expect(image.canMutate).toBe(true);
    expect(image.renderedWidthPoints).toBeCloseTo(90, 6);

    expect(result.reopenedImages).toHaveLength(1);
    expect(result.reopenedImages[0].intrinsicWidthPixels).toBe(180);
    expect(result.reopenedImages[0].intrinsicHeightPixels).toBe(174);
  });

  test('replace / resize / describe / remove all round-trip through the wrapper', async ({
    page,
  }) => {
    const docx = readTestFile('HC006-Test-01.docx');
    const first = readTestFile('img.png');
    const second = readTestFile('img2.png');

    const result = await page.evaluate(
      async ([docxArray, firstArray, secondArray]: number[][]) => {
        const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
        try {
          const proj = session.project();
          const anchor = Object.keys(proj.anchorIndex).find((k: string) =>
            k.startsWith('p:body:'),
          )!;
          const inserted = session.insertImage(anchor, 0, new Uint8Array(firstArray), {
            widthPoints: 100,
            heightPoints: 100,
            preserveAspect: false,
          });
          const id = inserted.imageId as string;

          const replaced = session.replaceImage(id, new Uint8Array(secondArray));
          const afterReplace = session.listImages()[0];

          const resized = session.setImageDimensions(id, {
            widthPoints: 50,
            heightPoints: 40,
            preserveAspect: false,
          });
          const described = session.setImageMetadata(id, 'alt text', 'the title');
          const afterEdits = session.listImages()[0];

          const removed = session.removeImage(id);
          const afterRemove = session.listImages();

          return {
            replaced: replaced.success,
            afterReplace,
            resized: resized.success,
            described: described.success,
            afterEdits,
            removed: removed.success,
            afterRemoveCount: afterRemove.length,
            undone: session.undo(),
            afterUndoCount: session.listImages().length,
          };
        } finally {
          session.close();
        }
      },
      [Array.from(docx), Array.from(first), Array.from(second)],
    );

    expect(result.replaced).toBe(true);
    // img2.png is 199x147: the *new* intrinsic size is visible immediately…
    expect(result.afterReplace.intrinsicWidthPixels).toBe(199);
    expect(result.afterReplace.intrinsicHeightPixels).toBe(147);
    // …while the rendered box deliberately survives the byte swap unchanged.
    expect(result.afterReplace.renderedWidthPoints).toBeCloseTo(100, 6);
    expect(result.afterReplace.renderedHeightPoints).toBeCloseTo(100, 6);

    expect(result.resized).toBe(true);
    expect(result.described).toBe(true);
    expect(result.afterEdits.renderedWidthPoints).toBeCloseTo(50, 6);
    expect(result.afterEdits.renderedHeightPoints).toBeCloseTo(40, 6);
    expect(result.afterEdits.altText).toBe('alt text');
    expect(result.afterEdits.title).toBe('the title');

    expect(result.removed).toBe(true);
    expect(result.afterRemoveCount).toBe(0);
    expect(result.undone).toBe(true);
    expect(result.afterUndoCount).toBe(1);
  });

  test('setImageFloatingLayout writes the anchored subset and reads it back', async ({ page }) => {
    const docx = readTestFile('HC006-Test-01.docx');
    const png = readTestFile('img.png');

    const result = await page.evaluate(
      async ([docxArray, pngArray]: number[][]) => {
        const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
        try {
          const proj = session.project();
          const anchor = Object.keys(proj.anchorIndex).find((k: string) =>
            k.startsWith('p:body:'),
          )!;
          const inserted = session.insertImage(anchor, 0, new Uint8Array(pngArray), {
            placement: 'floating',
            floatingLayout: {
              horizontalRelativeFrom: 'page',
              horizontalOffsetEmu: 914400,
              verticalRelativeFrom: 'margin',
              // Negative offsets are legal and must survive JSON serialization.
              verticalOffsetEmu: -457200,
              wrapMode: 'square',
              wrapSide: 'left',
            },
          });
          const id = inserted.imageId as string;
          const afterInsert = session.listImages()[0];

          const applied = session.setImageFloatingLayout(id, {
            horizontalRelativeFrom: 'margin',
            horizontalAlignment: 'right',
            verticalRelativeFrom: 'line',
            verticalAlignment: 'top',
            wrapMode: 'none',
            wrapSide: 'both_sides',
          });
          const afterLayout = session.listImages()[0];

          return {
            success: inserted.success,
            errorCode: inserted.error?.code,
            afterInsert,
            applied: applied.success,
            appliedError: applied.error?.code,
            afterLayout,
          };
        } finally {
          session.close();
        }
      },
      [Array.from(docx), Array.from(png)],
    );

    expect(result.errorCode).toBeUndefined();
    expect(result.success).toBe(true);
    expect(result.afterInsert.placement).toBe('floating');
    expect(result.afterInsert.floatingLayoutSupported).toBe(true);
    expect(result.afterInsert.floatingLayout.horizontalOffsetEmu).toBe(914400);
    // The negative-number path: culture-sensitive formatting here would have produced JSON
    // that `JSON.parse` rejects outright, so arriving as a number at all is the assertion.
    expect(result.afterInsert.floatingLayout.verticalOffsetEmu).toBe(-457200);
    expect(result.afterInsert.floatingLayout.wrapMode).toBe('square');

    expect(result.appliedError).toBeUndefined();
    expect(result.applied).toBe(true);
    expect(result.afterLayout.floatingLayout.horizontalAlignment).toBe('right');
    expect(result.afterLayout.floatingLayout.verticalAlignment).toBe('top');
    expect(result.afterLayout.floatingLayout.wrapMode).toBe('none');
  });

  test('rejected input returns a typed error rather than throwing across the boundary', async ({
    page,
  }) => {
    const docx = readTestFile('HC006-Test-01.docx');
    const png = readTestFile('img.png');

    const result = await page.evaluate(
      async ([docxArray, pngArray]: number[][]) => {
        const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
        try {
          const proj = session.project();
          const anchor = Object.keys(proj.anchorIndex).find((k: string) =>
            k.startsWith('p:body:'),
          )!;
          return {
            empty: session.insertImage(anchor, 0, new Uint8Array([])).error?.code,
            // Valid base64, but not any image signature we accept.
            garbage: session.insertImage(anchor, 0, new Uint8Array([1, 2, 3, 4, 5, 6, 7, 8]))
              .error?.code,
            // Real bytes, so validation gets past the payload and reaches the anchor.
            badAnchor: session.insertImage(
              'p:body:deadbeefdeadbeef',
              0,
              new Uint8Array(pngArray),
            ).error?.code,
            missingImage: session.removeImage('img:body:deadbeefdeadbeef').error?.code,
            stillEmpty: session.listImages().length,
          };
        } finally {
          session.close();
        }
      },
      [Array.from(docx), Array.from(png)],
    );

    expect(result.empty).toBe('invalid_image_data');
    expect(result.garbage).toBe('unsupported_image_format');
    expect(result.badAnchor).toBe('anchor_not_found');
    expect(result.missingImage).toBe('image_not_found');
    // None of the rejections may have partially mutated the document.
    expect(result.stillEmpty).toBe(0);
  });
});

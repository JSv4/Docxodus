import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
// An embedded PNG in the first paragraph and an external linked picture in the second.
const fixture = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/IM762-ImageCoverage.docx'),
));

function png(width: number, height: number): number[] {
  const bytes = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52,
    0, 0, 0, 0, 0, 0, 0, 0];
  bytes[19] = width; bytes[23] = height;
  return bytes;
}

// A minimal lossless (VP8L) WebP header the engine's header parser reads dimensions from.
function webp(width: number, height: number): number[] {
  const w = width - 1, h = height - 1;
  const ascii = (s: string) => Array.from(s).map(c => c.charCodeAt(0));
  return [...ascii('RIFF'), 22, 0, 0, 0, ...ascii('WEBP'), ...ascii('VP8L'), 10, 0, 0, 0, 0x2f,
    w & 0xff, ((w >> 8) & 0x3f) | ((h & 0x03) << 6), (h >> 2) & 0xff, (h >> 10) & 0x0f, 0, 0, 0, 0, 0];
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Image coverage matrix, WebP, wrap polygons and tracked edits (#762)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('the matrix names what applies, embed_linked converts, and WebP is a real media part', async ({ page }) => {
    const outcome = await page.evaluate(({ bytes, webpBytes, pngBytes }) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const capabilities = session.getImageCapabilities();
        const linked = session.listImages().find((image: any) => image.isLinked);
        const embedded = session.listImages().find((image: any) => !image.isLinked);
        const op = (image: any, name: string) => image.operations.find((o: any) => o.operation === name);
        const refused = session.replaceImage(linked.id, new Uint8Array(pngBytes));
        const converted = session.embedLinkedImage(linked.id, new Uint8Array(webpBytes));
        const after = session.listImages().find((image: any) => image.id === linked.id);
        return {
          webpInsert: capabilities.formats.find((f: any) => f.format === 'webp').canInsert,
          wrapModes: capabilities.mutableWrapModes,
          trackedOperations: capabilities.trackedOperations,
          markups: capabilities.markups.map((m: any) => m.markup),
          linkedCanMutate: linked.canMutate,
          linkedReplace: op(linked, 'replace'),
          linkedEmbed: op(linked, 'embed_linked').canMutate,
          embeddedEmbedReason: op(embedded, 'embed_linked').reason,
          refusedCode: refused.error?.code,
          convertedSuccess: converted.success,
          afterLinked: after.isLinked,
          afterFormat: after.format,
          afterContentType: after.contentType,
          afterWidth: after.intrinsicWidthPixels,
        };
      } finally { session.close(); }
    }, { bytes: Array.from(fixture), webpBytes: webp(6, 7), pngBytes: png(6, 7) });
    expect(outcome.webpInsert).toBe(true);
    expect(outcome.wrapModes).toEqual(['none', 'square', 'tight', 'through', 'top_and_bottom']);
    expect(outcome.trackedOperations).toContain('embed_linked');
    expect(outcome.markups).toContain('alternate_content');
    expect(outcome.linkedCanMutate).toBe(false);
    expect(outcome.linkedReplace.canMutate).toBe(false);
    expect(outcome.linkedReplace.reason).toContain('embed_linked');
    expect(outcome.linkedEmbed).toBe(true);
    expect(outcome.embeddedEmbedReason).toBe('picture is already embedded');
    expect(outcome.refusedCode).toBe('linked_image_read_only');
    expect(outcome.convertedSuccess).toBe(true);
    expect(outcome.afterLinked).toBe(false);
    expect(outcome.afterFormat).toBe('webp');
    expect(outcome.afterContentType).toBe('image/webp');
    expect(outcome.afterWidth).toBe(6);
  });

  test('tight wrap polygons round-trip and tracked replace records a revision pair', async ({ page }) => {
    const outcome = await page.evaluate(({ bytes, pngBytes }) => {
      const D = (window as any).Docxodus;
      const session = D.openTypedSession(new Uint8Array(bytes));
      const outline = { points: [{ x: 0, y: 10800 }, { x: 10800, y: 0 }, { x: 21600, y: 10800 }, { x: 0, y: 10800 }], edited: true };
      let polygon: any; let rectangle: any;
      try {
        const anchor = session.listImages().find((image: any) => !image.isLinked).anchorId;
        const inserted = session.insertImage(anchor, 0, new Uint8Array(pngBytes), {
          placement: 'floating', floatingLayout: { wrapMode: 'tight', wrapPolygon: outline },
        });
        polygon = session.listImages().find((image: any) => image.id === inserted.imageId).floatingLayout.wrapPolygon;
        session.setImageFloatingLayout(inserted.imageId, { wrapMode: 'through' });
        rectangle = session.listImages().find((image: any) => image.id === inserted.imageId).floatingLayout.wrapPolygon;
      } finally { session.close(); }
      const tracked = D.openTypedSession(new Uint8Array(bytes), JSON.stringify({ trackedChanges: 'render_inline' }));
      try {
        const embedded = tracked.listImages().find((image: any) => !image.isLinked);
        const replaced = tracked.replaceImage(embedded.id, new Uint8Array(pngBytes));
        const images = tracked.listImages();
        const deleted = images.find((image: any) => image.id === embedded.id);
        const inserted = images.find((image: any) => image.id === replaced.imageId);
        const types = tracked.listRevisions().map((r: any) => r.type);
        tracked.rejectAllRevisions();
        const restored = tracked.listImages().find((image: any) => !image.isLinked);
        return {
          polygon, rectangle,
          replacedSuccess: replaced.success,
          newId: replaced.imageId !== embedded.id,
          deletedCanMutate: deleted.canMutate,
          deletedReason: deleted.unsupportedReason,
          insertedCanMutate: inserted.canMutate,
          insertedWidth: inserted.intrinsicWidthPixels,
          types,
          restoredId: restored.id === embedded.id,
          restoredWidth: restored.intrinsicWidthPixels,
        };
      } finally { tracked.close(); }
    }, { bytes: Array.from(fixture), pngBytes: png(9, 9) });
    expect(outcome.polygon).toEqual({ points: [{ x: 0, y: 10800 }, { x: 10800, y: 0 }, { x: 21600, y: 10800 }, { x: 0, y: 10800 }], edited: true });
    expect(outcome.rectangle.edited).toBe(false);
    expect(outcome.rectangle.points).toHaveLength(5);
    expect(outcome.replacedSuccess).toBe(true);
    expect(outcome.newId).toBe(true);
    expect(outcome.deletedCanMutate).toBe(false);
    expect(outcome.deletedReason).toContain('tracked deletion');
    expect(outcome.insertedCanMutate).toBe(true);
    expect(outcome.insertedWidth).toBe(9);
    expect(outcome.types).toContain('insert');
    expect(outcome.types).toContain('delete');
    expect(outcome.restoredId).toBe(true);
    expect(outcome.restoredWidth).toBe(2);
  });
});

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

/** Return the first N body p/h/li anchors sorted in markdown document order. */
function firstBodyAnchors(proj: any, count: number): string[] {
  return (Object.entries(proj.anchorIndex) as [string, any][])
    .map(([id, t]) => ({ id, ...t }))
    .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
    .map(t => ({ t, idx: (proj.markdown as string).indexOf('{#' + t.id + '}') }))
    .filter(x => x.idx >= 0)
    .sort((a, b) => a.idx - b.idx)
    .slice(0, count)
    .map(x => x.t.id);
}

test.describe('annotation write surface', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  // ─── Test A ──────────────────────────────────────────────────────────────
  // add → list → remove → save → reopen → empty
  test('add → list → remove → save → reopen → empty', async ({ page }) => {
    const bytes = readTestFile('DA001-TemplateDocument.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;

      // ── helpers ──────────────────────────────────────────────────────────
      function bodyAnchors(proj: any, n: number): string[] {
        return (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map(t => ({ t, idx: (proj.markdown as string).indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a: any, b: any) => a.idx - b.idx)
          .slice(0, n)
          .map((x: any) => x.t.id);
      }

      // ── Phase 1: open, add, list, remove ─────────────────────────────────
      const h1 = bridge.OpenSession(bin, '');
      let anchorId: string;
      let addResult: any;
      let listAfterAdd: any[];
      let removeResult: any;
      let savedBytes: Uint8Array;
      try {
        const proj = JSON.parse(bridge.Project(h1));
        const anchors = bodyAnchors(proj, 1);
        if (anchors.length === 0) throw new Error('no body anchors found');
        anchorId = anchors[0];

        const annJson = JSON.stringify({ id: 'ws-1', labelId: 'LBL', label: 'Label', color: '#FF0000', bookmarkName: '' });
        addResult = JSON.parse(bridge.AddAnnotation(h1, anchorId, '', annJson));

        listAfterAdd = JSON.parse(bridge.ListAnnotations(h1));
        removeResult = JSON.parse(bridge.SessionRemoveAnnotation(h1, 'ws-1'));
        savedBytes = bridge.Save(h1);
      } finally {
        bridge.CloseSession(h1);
      }

      // ── Phase 2: reopen the saved bytes, verify annotation is gone ────────
      const h2 = bridge.OpenSession(savedBytes, '');
      let listAfterReopen: any[];
      try {
        listAfterReopen = JSON.parse(bridge.ListAnnotations(h2));
      } finally {
        bridge.CloseSession(h2);
      }

      return {
        anchorId,
        addSuccess: addResult.success,
        addAnnotationId: addResult.annotationId,
        addError: addResult.error,
        listAfterAddIncludesId: listAfterAdd.some((a: any) => a.id === 'ws-1'),
        removeSuccess: removeResult.success,
        savedByteCount: savedBytes.length,
        listAfterReopenIncludesId: listAfterReopen.some((a: any) => a.id === 'ws-1'),
      };
    }, Array.from(bytes));

    expect(result.addSuccess, `addAnnotation failed: ${JSON.stringify(result.addError)}`).toBe(true);
    expect(result.addAnnotationId).toBe('ws-1');
    expect(result.listAfterAddIncludesId).toBe(true);
    expect(result.removeSuccess).toBe(true);
    expect(result.savedByteCount).toBeGreaterThan(0);
    expect(result.listAfterReopenIncludesId).toBe(false);
  });

  // ─── Test B ──────────────────────────────────────────────────────────────
  // updateAnnotation mutates label and applies a metadata patch; both surface
  // through ListAnnotations now that SerializeAnnotations emits the metadata bag.
  test('update mutates label and metadata patch is observable via list', async ({ page }) => {
    const bytes = readTestFile('DA001-TemplateDocument.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;

      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        const anchors = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map(t => ({ t, idx: (proj.markdown as string).indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a: any, b: any) => a.idx - b.idx)
          .map((x: any) => x.t.id);
        if (anchors.length === 0) throw new Error('no body anchors found');
        const anchorId = anchors[0];

        // Add with original label and metadata
        const annJson = JSON.stringify({
          id: 'ws-2',
          labelId: 'OLD',
          label: 'Old',
          color: '#000000',
          bookmarkName: '',
          metadata: { keep: 'yes', drop: 'old' },
        });
        const addResult = JSON.parse(bridge.AddAnnotation(h, anchorId, '', annJson));

        // Update: change label, patch metadata (drop "drop", add "new")
        const updateJson = JSON.stringify({
          label: 'New',
          metadataPatch: { drop: null, new: 'fresh' },
        });
        const updateResult = JSON.parse(bridge.UpdateAnnotation(h, 'ws-2', updateJson));

        // ListAnnotations now surfaces metadata, so we can verify the patch landed.
        const listed = (JSON.parse(bridge.ListAnnotations(h)) as any[]).find(a => a.id === 'ws-2');

        return {
          addSuccess: addResult.success,
          addError: addResult.error,
          updateSuccess: updateResult.success,
          updateAnnotationId: updateResult.annotationId,
          updateError: updateResult.error,
          listedLabel: listed?.label,
          listedLabelId: listed?.labelId,
          listedMetadata: listed?.metadata,
        };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    expect(result.addSuccess, `addAnnotation failed: ${JSON.stringify(result.addError)}`).toBe(true);
    expect(result.updateSuccess, `updateAnnotation failed: ${JSON.stringify(result.updateError)}`).toBe(true);
    expect(result.updateAnnotationId).toBe('ws-2');
    // Label changed from "Old" to "New"
    expect(result.listedLabel).toBe('New');
    // labelId was not patched so should remain "OLD"
    expect(result.listedLabelId).toBe('OLD');
    // Metadata patch: "drop" removed, "new" added, "keep" untouched
    expect(result.listedMetadata).toEqual({ keep: 'yes', new: 'fresh' });
  });

  // ─── Test C ──────────────────────────────────────────────────────────────
  // moveAnnotation re-targets the bookmark to a different block
  // A cross-block move reports both old and new anchors in EditResult.modified.
  test('move re-targets the bookmark (cross-block → modified.length === 2)', async ({ page }) => {
    const bytes = readTestFile('DA001-TemplateDocument.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;

      const h = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(h));
        // Need at least 2 distinct body anchors so the move is genuinely cross-block
        const paragraphs = (Object.entries(proj.anchorIndex) as [string, any][])
          .map(([id, t]) => ({ id, ...t }))
          .filter(t => t.scope === 'body' && ['p', 'h', 'li'].includes(t.kind))
          .map(t => ({ t, idx: (proj.markdown as string).indexOf('{#' + t.id + '}') }))
          .filter(x => x.idx >= 0)
          .sort((a: any, b: any) => a.idx - b.idx)
          .slice(0, 2)
          .map((x: any) => x.t.id);

        if (paragraphs.length < 2) throw new Error('fixture has fewer than 2 body anchors');

        const [p0, p1] = paragraphs;

        // Add to first paragraph
        const annJson = JSON.stringify({
          id: 'ws-3',
          labelId: 'L',
          label: 'L',
          color: '#000000',
          bookmarkName: '',
        });
        const addResult = JSON.parse(bridge.AddAnnotation(h, p0, '', annJson));

        // Move to second paragraph with explicit span { start:0, length:2 }
        const spanJson = JSON.stringify({ start: 0, length: 2 });
        const moveResult = JSON.parse(bridge.MoveAnnotation(h, 'ws-3', p1, spanJson));

        return {
          addSuccess: addResult.success,
          addError: addResult.error,
          moveSuccess: moveResult.success,
          moveError: moveResult.error,
          moveAnnotationId: moveResult.annotationId,
          modifiedCount: moveResult.modified?.length ?? -1,
          p0,
          p1,
        };
      } finally {
        bridge.CloseSession(h);
      }
    }, Array.from(bytes));

    expect(result.addSuccess, `addAnnotation failed: ${JSON.stringify(result.addError)}`).toBe(true);
    expect(result.moveSuccess, `moveAnnotation failed: ${JSON.stringify(result.moveError)}`).toBe(true);
    expect(result.moveAnnotationId).toBe('ws-3');
    // Cross-block move: both old and new anchors appear in modified[]
    expect(result.modifiedCount).toBe(2);
  });
});

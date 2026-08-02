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

// Issue #318 — markup-native revision listing + selective per-revision accept/reject.
test.describe('DocxSession revision review (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('ListRevisions reads markup identity; AcceptRevision/RejectRevision resolve selectively', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(
        bin,
        JSON.stringify({ trackedChanges: 'render_inline', revisionAuthor: 'Spec Reviewer' }),
      );
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:'),
        )!;

        const edit = JSON.parse(bridge.ReplaceText(handle, bodyAnchor, 'Tracked rewrite.'));

        const listed: any[] = JSON.parse(bridge.ListRevisions(handle));
        const insertRev = listed.find((r) => r.type === 'insert');
        const deleteRev = listed.find((r) => r.type === 'delete');

        const accepted = JSON.parse(bridge.AcceptRevision(handle, insertRev.id));
        const remaining: any[] = JSON.parse(bridge.ListRevisions(handle));
        const rejected = JSON.parse(bridge.RejectRevision(handle, deleteRev.id));
        const afterBoth: any[] = JSON.parse(bridge.ListRevisions(handle));

        const missing = JSON.parse(bridge.AcceptRevision(handle, 'rev999999'));

        return {
          editOk: edit.success,
          listedCount: listed.length,
          insertAuthor: insertRev?.author,
          insertText: insertRev?.text,
          insertHasAnchor: typeof insertRev?.anchorId === 'string',
          idsStartWithRev: listed.every((r) => typeof r.id === 'string' && r.id.startsWith('rev')),
          acceptOk: accepted.success,
          remainingIds: remaining.map((r: any) => r.id),
          deleteId: deleteRev?.id,
          rejectOk: rejected.success,
          afterBothCount: afterBoth.length,
          missingOk: missing.success,
          missingCode: missing.error?.code,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.editOk).toBe(true);
    expect(result.listedCount).toBe(2); // one delete (old text) + one insert (new text)
    expect(result.insertAuthor).toBe('Spec Reviewer');
    expect(result.insertText).toBe('Tracked rewrite.');
    expect(result.insertHasAnchor).toBe(true);
    expect(result.idsStartWithRev).toBe(true);
    expect(result.acceptOk).toBe(true);
    // Resolving one revision leaves the other's id untouched.
    expect(result.remainingIds).toEqual([result.deleteId]);
    expect(result.rejectOk).toBe(true);
    expect(result.afterBothCount).toBe(0);
    expect(result.missingOk).toBe(false);
    expect(result.missingCode).toBe('revision_not_found');
  });
});

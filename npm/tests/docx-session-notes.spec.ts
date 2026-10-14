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

// Issue #276 — footnote/endnote authoring through the WASM bridge.
test.describe('DocxSession footnote/endnote authoring (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('InsertFootnote creates the note, cites it, projects it, and round-trips', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:'),
        )!;

        const made = JSON.parse(bridge.InsertFootnote(handle, bodyAnchor, 0, 'Source: 2025 annual report.'));
        const created: { id: string; kind: string; scope: string }[] = made.created ?? [];

        const after = JSON.parse(bridge.Project(handle));
        const saved = bridge.Save(handle);

        // Reopen the saved bytes in a fresh session to prove the note persisted.
        const handle2 = bridge.OpenSession(saved, '');
        let reopenedHasNote = false;
        try {
          reopenedHasNote = JSON.parse(bridge.Project(handle2)).markdown.includes(
            'Source: 2025 annual report.',
          );
        } finally {
          bridge.CloseSession(handle2);
        }

        return {
          success: made.success,
          errorCode: made.error?.code,
          hasNoteDefAnchor: created.some((a) => a.kind === 'fn' && a.scope === 'fn'),
          hasNoteParaAnchor: created.some((a) => a.kind === 'p' && a.scope === 'fn'),
          modifiedHost: (made.modified ?? []).some((a: any) => a.id === bodyAnchor),
          markdownHasFootnotesSection: after.markdown.includes('# Footnotes'),
          markdownHasNoteText: after.markdown.includes('Source: 2025 annual report.'),
          savedBytes: saved.length,
          reopenedHasNote,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.success, `error=${result.errorCode}`).toBe(true);
    expect(result.hasNoteDefAnchor).toBe(true);
    expect(result.hasNoteParaAnchor).toBe(true);
    expect(result.modifiedHost).toBe(true);
    expect(result.markdownHasFootnotesSection).toBe(true);
    expect(result.markdownHasNoteText).toBe(true);
    expect(result.savedBytes).toBeGreaterThan(0);
    expect(result.reopenedHasNote).toBe(true);
  });

  test('an authored footnote is editable and deletable through the existing ops', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const made = JSON.parse(bridge.InsertFootnote(handle, bodyAnchor, 0, 'Original.'));
        const created: { id: string; kind: string; scope: string }[] = made.created ?? [];
        const notePara = created.find((a) => a.kind === 'p' && a.scope === 'fn')!.id;
        const noteDef = created.find((a) => a.kind === 'fn')!.id;

        const edited = JSON.parse(bridge.ReplaceText(handle, notePara, 'Rewritten.'));
        const afterEdit = JSON.parse(bridge.Project(handle)).markdown;

        const deleted = JSON.parse(bridge.DeleteBlock(handle, noteDef));
        const afterDelete = JSON.parse(bridge.Project(handle)).markdown;

        return {
          editSuccess: edited.success,
          editShows: afterEdit.includes('Rewritten.'),
          deleteSuccess: deleted.success,
          noteGone: !afterDelete.includes('Rewritten.') && !afterDelete.includes('# Footnotes'),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.editSuccess).toBe(true);
    expect(result.editShows).toBe(true);
    expect(result.deleteSuccess).toBe(true);
    expect(result.noteGone).toBe(true);
  });

  test('InsertEndnote error envelope: a note-scope anchor gives a typed error code', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const made = JSON.parse(bridge.InsertEndnote(handle, bodyAnchor, 0, 'A note.'));
        const notePara = (made.created ?? []).find((a: any) => a.kind === 'p' && a.scope === 'en')!.id;

        // Word does not allow a note reference inside another note's story.
        const nested = JSON.parse(bridge.InsertEndnote(handle, notePara, 0, 'Nested.'));
        // Nor is an out-of-range offset accepted.
        const badOffset = JSON.parse(bridge.InsertEndnote(handle, bodyAnchor, 100000, 'Nope.'));

        return {
          madeSuccess: made.success,
          nestedSuccess: nested.success,
          nestedCode: nested.error?.code,
          badOffsetSuccess: badOffset.success,
          badOffsetCode: badOffset.error?.code,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.madeSuccess).toBe(true);
    expect(result.nestedSuccess).toBe(false);
    expect(result.nestedCode).toBe('anchor_wrong_kind');
    expect(result.badOffsetSuccess).toBe(false);
    expect(result.badOffsetCode).toBe('offset_out_of_range');
  });
});

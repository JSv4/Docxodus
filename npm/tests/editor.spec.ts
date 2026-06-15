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

// End-to-end DocxEditor (Option B): render faithful pages → edit a block →
// ONLY that block re-renders from the live session → save is lossless for
// untouched content. The complete editor experience, in a real browser.
test.describe('DocxEditor — block editor end-to-end', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('renders, edits one block incrementally, and saves losslessly', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);

      const editor = D.DocxEditor.open(container, bin, D, {});

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const editablePs = () =>
        Array.from(container.querySelectorAll('p[data-anchor][contenteditable="true"]'))
          .filter((e) => norm(e.textContent || '').length > 5) as HTMLElement[];

      const before = editablePs();
      const blockCountBefore = container.querySelectorAll('[data-anchor]').length;

      // Target = first editable paragraph; witness = a different one (must survive).
      const target = before[0];
      const witness = before[before.length - 1];
      const targetUnidBefore = target.getAttribute('data-anchor');
      const witnessText = norm(witness.textContent || '');
      const anchorsBefore = Array.from(container.querySelectorAll('[data-anchor]'))
        .map((e) => e.getAttribute('data-anchor'));

      // Edit the target block and commit (blur fires the editor's commit path).
      const MARKER = 'EDITORLOOP99 brand new content';
      target.focus();
      target.textContent = MARKER;
      target.dispatchEvent(new Event('blur'));

      // After commit: the target block was re-rendered in place (its anchor changed,
      // the witness is untouched). Find the block now carrying the marker.
      const edited = editablePs().find((e) => norm(e.textContent || '').includes('EDITORLOOP99'));
      const editedText = edited ? norm(edited.textContent || '') : '(missing)';
      const editedAnchor = edited ? edited.getAttribute('data-anchor') : null;
      const witnessStillPresent = editablePs().some((e) => norm(e.textContent || '') === witnessText);
      const blockCountAfter = container.querySelectorAll('[data-anchor]').length;

      // How many block anchors changed? Only the edited one should differ.
      const anchorsAfter = Array.from(container.querySelectorAll('[data-anchor]'))
        .map((e) => e.getAttribute('data-anchor'));
      const changed = anchorsBefore.filter((a) => !anchorsAfter.includes(a)).length;

      // Save (lossless) and reopen to confirm persistence + untouched survival.
      const saved: Uint8Array = editor.save();
      const reopened = D.DocxSessionBridge.OpenSession(saved, '');
      const md = JSON.parse(D.DocxSessionBridge.Project(reopened)).markdown as string;
      D.DocxSessionBridge.CloseSession(reopened);

      // Compare on alphanumeric-normalized text — robust to markdown escaping/whitespace.
      const alnum = (s: string) => (s || '').toLowerCase().replace(/[^a-z0-9]/g, '');
      const mdAlnum = alnum(md);

      editor.close();
      container.remove();

      return {
        blockCountBefore,
        blockCountAfter,
        targetUnidBefore,
        editedAnchor,
        editedText,
        witnessText,
        witnessStillPresent,
        anchorsChanged: changed,
        savedLen: saved.length,
        savedHasEdit: mdAlnum.includes('editorloop99'),
        savedHasWitness: mdAlnum.includes(alnum(witnessText).slice(0, 30)),
        editableCount: before.length,
      };
    }, Array.from(bytes));

    // Rendered as many addressable, editable blocks.
    expect(out.editableCount).toBeGreaterThan(1);
    expect(out.blockCountBefore).toBeGreaterThan(1);
    // The edit is visible in the re-rendered block...
    expect(out.editedText).toContain('EDITORLOOP99');
    // ...and the block's anchor is STABLE within the live session (ReplaceText mutates
    // runs in place; the Unid attribute is not re-derived per edit) — so no anchor churn
    // and the editor needs no stable-key remap mid-session.
    expect(out.editedAnchor).toBe(out.targetUnidBefore);
    expect(out.anchorsChanged).toBe(0);
    expect(out.blockCountAfter).toBe(out.blockCountBefore);
    // An untouched block survived in the DOM and in the saved document (lossless).
    expect(out.witnessStillPresent).toBe(true);
    expect(out.savedLen).toBeGreaterThan(0);
    expect(out.savedHasEdit).toBe(true);
    expect(out.savedHasWitness).toBe(true);
  });

  // Paginated mode: blocks flow into real page boxes (margins/headers via the converter
  // + pagination.ts), and those page blocks remain editable with incremental re-render.
  test('paginated mode renders page boxes with editable blocks', async ({ page }) => {
    const bytes = readTestFile('HC031-Complicated-Document.docx');

    const out = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const D = (window as any).Docxodus;

      const container = document.createElement('div');
      document.body.appendChild(container);
      const editor = D.DocxEditor.open(container, bin, D, { paginated: true });

      const norm = (s: string) => (s || '').replace(/\s+/g, ' ').trim();
      const pageContainer = container.querySelector('#pagination-container');
      const pageBoxes = container.querySelectorAll('.page-box').length;
      const editable = pageContainer
        ? (Array.from(
            pageContainer.querySelectorAll('p[data-anchor][contenteditable="true"]'),
          ) as HTMLElement[]).filter((e) => norm(e.textContent || '').length > 5)
        : [];

      let editedText = '(none)';
      const target = editable[0];
      if (target) {
        target.focus();
        target.textContent = 'PAGINATEDEDIT77 content';
        target.dispatchEvent(new Event('blur'));
        const edited = (Array.from(
          (pageContainer as HTMLElement).querySelectorAll('p[data-anchor]'),
        ) as HTMLElement[]).find((e) => norm(e.textContent || '').includes('PAGINATEDEDIT77'));
        editedText = edited ? norm(edited.textContent || '') : '(missing)';
      }

      editor.close();
      container.remove();
      return { pageBoxes, editableCount: editable.length, editedText };
    }, Array.from(bytes));

    expect(out.pageBoxes).toBeGreaterThan(0); // real page boxes rendered
    expect(out.editableCount).toBeGreaterThan(0); // blocks inside pages stay editable
    expect(out.editedText).toContain('PAGINATEDEDIT77'); // incremental edit inside a page
  });
});

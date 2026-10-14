import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/HC001-5DayTourPlanTemplate.docx'),
));

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Guarded commit of a retained preview (#760)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('a retained preview commits exactly as previewed and is one undo step', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        // The first block the projection lists at body level (a table-cell paragraph would not
        // surface its text in the top-level markdown read back below).
        const anchor = /\{#((?:p|h):body:[0-9a-f]+)\}/.exec(session.project().markdown as string)![1];
        const baseVersion = session.getVersion();
        const preview = session.previewBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: (shadow: any) => shadow.insertParagraph(anchor, 'after', 'committed as previewed') },
        ], 'atomic', { retain: true });
        const previewedId = preview.steps[0].results[0].created[0].id as string;
        const untouched = session.getVersion() === baseVersion
          && !(session.project().markdown as string).includes('committed as previewed');

        const commit = session.commitPreview(preview.retention.previewId);
        const versionMatches = commit.resultVersion === preview.resultVersion
          && session.getVersion() === commit.resultVersion;
        const markdown = session.project().markdown as string;
        const undone = session.undo();
        const afterUndo = session.project().markdown as string;
        return {
          previewSuccess: preview.success,
          retention: preview.retention,
          previewHash: preview.packageHash,
          previewCaveat: preview.warnings.some((w: string) => w.includes('may be generated')),
          untouched,
          commitSuccess: commit.success,
          commitPreviewFlag: commit.preview,
          commitHash: commit.packageHash,
          commitRetention: commit.retention,
          commitCaveat: commit.warnings.some((w: string) => w.includes('may be generated')),
          commitStepCount: commit.steps.length,
          committedId: commit.steps[0]?.results[0]?.created[0]?.id,
          previewedId,
          versionMatches,
          landed: markdown.includes('committed as previewed') && markdown.includes(previewedId),
          undone,
          undoRemoved: !afterUndo.includes('committed as previewed'),
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.previewSuccess).toBe(true);
    expect(outcome.retention.previewId).toMatch(/^pv-/);
    expect(outcome.retention.baseVersion).toBe(0);
    expect(outcome.retention.expiresAt).toMatch(/Z$/);
    expect(outcome.previewCaveat).toBe(true);
    expect(outcome.untouched).toBe(true);
    expect(outcome.commitSuccess).toBe(true);
    expect(outcome.commitPreviewFlag).toBe(false);
    expect(outcome.commitHash).toBe(outcome.previewHash);
    expect(outcome.commitRetention).toEqual(outcome.retention);
    expect(outcome.commitCaveat).toBe(false);
    expect(outcome.commitStepCount).toBe(1);
    expect(outcome.committedId).toBe(outcome.previewedId);
    expect(outcome.versionMatches).toBe(true);
    expect(outcome.landed).toBe(true);
    expect(outcome.undone).toBe(true);
    expect(outcome.undoRemoved).toBe(true);
  });

  test('a stale or consumed preview is refused without editing', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const anchor = /\{#((?:p|h):body:[0-9a-f]+)\}/.exec(session.project().markdown as string)![1];
        const preview = session.previewBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: (shadow: any) => shadow.insertParagraph(anchor, 'after', 'never lands') },
        ], 'atomic', { retain: true });
        session.insertParagraph(anchor, 'after', 'intervening');
        const version = session.getVersion();
        const stale = session.commitPreview(preview.retention.previewId);
        const staleVersionUnchanged = session.getVersion() === version;

        const second = session.previewBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: (shadow: any) => shadow.insertParagraph(anchor, 'after', 'once') },
        ], 'atomic', { retain: true });
        const first = session.commitPreview(second.retention.previewId, {
          transactionId: 'tx-commit', request: { commit: second.retention.previewId },
        });
        const replay = session.commitPreview(second.retention.previewId, {
          transactionId: 'tx-commit', request: { commit: second.retention.previewId },
        });
        const consumed = session.commitPreview(second.retention.previewId);
        const plain = session.previewBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: (shadow: any) => shadow.insertParagraph(anchor, 'after', 'not retained') },
        ]);
        const markdown = session.project().markdown as string;
        return {
          staleCode: stale.failure?.error.code,
          staleSuccess: stale.success,
          staleVersionUnchanged,
          neverLanded: !markdown.includes('never lands'),
          firstSuccess: first.success,
          firstTransaction: first.transaction?.transactionId,
          replayIdentical: JSON.stringify(first) === JSON.stringify(replay),
          onceCount: markdown.split('once').length - 1,
          consumedCode: consumed.failure?.error.code,
          plainRetention: plain.retention,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.staleSuccess).toBe(false);
    expect(outcome.staleCode).toBe('preview_stale');
    expect(outcome.staleVersionUnchanged).toBe(true);
    expect(outcome.neverLanded).toBe(true);
    expect(outcome.firstSuccess).toBe(true);
    expect(outcome.firstTransaction).toBe('tx-commit');
    expect(outcome.replayIdentical).toBe(true);
    expect(outcome.onceCount).toBe(1);
    expect(outcome.consumedCode).toBe('preview_not_found');
    expect(outcome.plainRetention).toBeUndefined();
  });
});

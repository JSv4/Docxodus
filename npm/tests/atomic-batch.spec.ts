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

test.describe('DocxSession atomic batches (#445)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('rollback is exact and success is one version/undo unit', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchors = (Object.entries(projection.anchorIndex) as [string, any][])
          .filter(([id, value]) => value.scope === 'body'
            && ['p', 'h', 'li'].includes(value.kind)
            && projection.markdown.includes(`{#${id}}`))
          .map(([id]) => id);
        const before = projection.markdown;

        const failed = session.executeBatch([
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText(anchors[0], 'Speculative npm edit.') },
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText('p:body:missing', 'failure') },
        ]);
        const afterFailure = session.project().markdown;
        const versionAfterFailure = session.getVersion();
        const undoAfterFailure = session.undo();

        const committed = session.executeBatch([
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText(anchors[0], 'Committed npm first.') },
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText(anchors[1], 'Committed npm second.') },
        ]);
        const committedMarkdown = session.project().markdown;
        const committedVersion = session.getVersion();
        const undoCommitted = session.undo();

        return {
          failed,
          restored: afterFailure === before,
          versionAfterFailure,
          undoAfterFailure,
          committed,
          committedMarkdown,
          committedVersion,
          undoCommitted,
          restoredAfterUndo: session.project().markdown === before,
        };
      } finally {
        session.close();
      }
    }, Array.from(fixture));

    expect(result.failed.success).toBe(false);
    expect(result.failed.rolledBack).toBe(true);
    expect(result.failed.failure.index).toBe(1);
    expect(result.failed.failure.error.code).toBe('anchor_not_found');
    expect(result.restored).toBe(true);
    expect(result.versionAfterFailure).toBe(0);
    expect(result.undoAfterFailure).toBe(false);

    expect(result.committed.success).toBe(true);
    expect(result.committedVersion).toBe(1);
    expect(result.committedMarkdown).toContain('Committed npm first.');
    expect(result.committedMarkdown).toContain('Committed npm second.');
    expect(result.undoCommitted).toBe(true);
    expect(result.restoredAfterUndo).toBe(true);
  });

  test('best-effort preflight observes earlier sequential state', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchors = (Object.entries(projection.anchorIndex) as [string, any][])
          .filter(([id, value]) => value.scope === 'body'
            && ['p', 'h', 'li'].includes(value.kind)
            && projection.markdown.includes(`{#${id}}`))
          .map(([id]) => id);
        return session.executeBatch([
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText(anchors[0], 'Sequential npm state.') },
          { tool: 'docx_edit', action: 'replace_text',
            preflight: () => session.project().markdown.includes('Sequential npm state.')
              ? undefined
              : { code: 'precondition_failed', message: 'prior state missing' },
            mutation: () => session.replaceText(anchors[1], 'Observed npm state.') },
        ], 'best_effort');
      } finally {
        session.close();
      }
    }, Array.from(fixture));

    expect(result.success).toBe(true);
    expect(result.steps).toHaveLength(2);
  });
});

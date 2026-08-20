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

  test('an empty mutation result list is a successful no-op step (core parity)', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchors = (Object.entries(projection.anchorIndex) as [string, any][])
          .filter(([id, value]) => value.scope === 'body'
            && ['p', 'h', 'li'].includes(value.kind)
            && projection.markdown.includes(`{#${id}}`))
          .map(([id]) => id);
        const batch = session.executeBatch([
          { tool: 'docx_edit', action: 'replace_text_range',
            mutation: () => [] },
          { tool: 'docx_edit', action: 'replace_text',
            mutation: () => session.replaceText(anchors[0], 'Committed after the noop step.') },
        ]);
        return {
          success: batch.success,
          status: batch.status,
          stepSuccess: batch.steps.map((step: any) => step.success),
          markdown: session.project().markdown.includes('Committed after the noop step.'),
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(result.success).toBe(true);
    expect(result.status).toBe('ok');
    expect(result.stepSuccess).toEqual([true, true]);
    expect(result.markdown).toBe(true);
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

  test('isolated preview returns a rich shadow receipt and preserves live redo', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(
        new Uint8Array(bytes),
        JSON.stringify({ undoDepth: 1, persistAnchorIds: true }),
      );
      try {
        const projection = session.project();
        const anchors = (Object.entries(projection.anchorIndex) as [string, any][])
          .filter(([id, value]) => value.scope === 'body'
            && ['p', 'h', 'li'].includes(value.kind)
            && projection.markdown.includes(`{#${id}}`))
          .map(([id]) => id);
        session.replaceText(anchors[0], 'npm redo target');
        session.undo();
        const beforeVersion = session.getVersion();
        const before = Array.from(session.save());

        const preview = session.previewBatch([
          { tool: 'docx_edit', action: 'replace_text',
            mutation: (shadow: any) => shadow.replaceText(anchors[0], 'Predicted npm first.') },
          { tool: 'docx_edit', action: 'replace_text',
            mutation: (shadow: any) => shadow.replaceText(anchors[1], 'Predicted npm second.') },
        ], 'atomic', { html: 'full' });

        const after = Array.from(session.save());
        const liveMarkdown = session.project().markdown;
        const liveVersion = session.getVersion();
        const undo = session.undo();
        const redo = session.redo();
        return {
          preview,
          beforeVersion,
          bytesEqual: before.length === after.length
            && before.every((value, index) => value === after[index]),
          liveMarkdown,
          liveVersion,
          undo,
          redo,
          redoneMarkdown: session.project().markdown,
        };
      } finally {
        session.close();
      }
    }, Array.from(fixture));

    expect(result.preview.preview).toBe(true);
    expect(result.preview.success).toBe(true);
    expect(result.preview.baseVersion).toBe(result.beforeVersion);
    expect(result.preview.resultVersion).toBe(result.beforeVersion + 1);
    expect(result.preview.packageHash).toMatch(/^[0-9a-f]{64}$/);
    expect(result.preview.steps).toHaveLength(2);
    expect(result.preview.revisionChanges).toEqual({ added: [], removed: [], modified: [] });
    expect(result.preview.commentChanges).toEqual({ added: [], removed: [], modified: [] });
    expect(result.preview.annotationChanges).toEqual({ added: [], removed: [], modified: [] });
    expect(result.preview.html).toContain('Predicted npm first.');
    expect(result.bytesEqual).toBe(true);
    expect(result.liveVersion).toBe(result.beforeVersion);
    expect(result.liveMarkdown).not.toContain('Predicted npm');
    expect(result.undo).toBe(false);
    expect(result.redo).toBe(true);
    expect(result.redoneMarkdown).toContain('npm redo target');
  });

  test('preview validates HTML mode first and treats renderer error envelopes as warnings', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const api = (window as any).Docxodus;
      const live = api.openTypedSession(new Uint8Array(bytes));
      try {
        const anchor = Object.keys(live.project().anchorIndex)
          .find(id => id.startsWith('p:body:'))!;
        let invalidInvoked = false;
        let invalidError = '';
        try {
          live.previewBatch([
            { tool: 'docx_edit', action: 'replace_text', mutation: (shadow: any) => {
              invalidInvoked = true;
              return shadow.replaceText(anchor, 'must not execute');
            } },
          ], 'atomic', { html: 'invalid' as any });
        } catch (error) {
          invalidError = error instanceof Error ? error.message : String(error);
        }

        // Intercept EVERY export the full-preview path may reach: the client prefers the
        // shared preview profile (RenderPreviewHtml) and falls back to the editor-profile
        // exports only on a bundle that predates it. Naming just one leaves this test green
        // for the wrong reason on whichever bundle the fallback does not apply to.
        const renderFailure = new Set([
          'RenderPreviewHtml', 'RenderHtmlForReview', 'RenderHtml',
        ]);
        const bridge = new Proxy(api.DocxSessionBridge, {
          get(target, property, receiver) {
            if (typeof property === 'string' && renderFailure.has(property)) {
              return () => JSON.stringify({ error: 'simulated renderer failure' });
            }
            return Reflect.get(target, property, receiver);
          },
        });
        const wrapped = new api.DocxSession(
          bridge.OpenSession(new Uint8Array(bytes), ''),
          bridge,
        );
        try {
          const wrappedAnchor = Object.keys(wrapped.project().anchorIndex)
            .find(id => id.startsWith('p:body:'))!;
          const preview = wrapped.previewBatch([
            { tool: 'docx_edit', action: 'replace_text',
              mutation: (shadow: any) => shadow.replaceText(wrappedAnchor, 'shadow only') },
          ], 'atomic', { html: 'full' });
          return { invalidInvoked, invalidError, preview, liveVersion: live.getVersion() };
        } finally {
          wrapped.close();
        }
      } finally {
        live.close();
      }
    }, Array.from(fixture));

    expect(result.invalidInvoked).toBe(false);
    expect(result.invalidError).toContain('unknown preview HTML mode');
    expect(result.liveVersion).toBe(0);
    expect(result.preview.success).toBe(true);
    expect(result.preview.html).toBeNull();
    expect(result.preview.warnings).toContain(
      'Preview HTML could not be generated: simulated renderer failure',
    );
  });
});

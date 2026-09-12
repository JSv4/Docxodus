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

test.describe('DocxSession mutation transactions (#761)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('an identical retry replays the original result without executing again', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchor = (Object.entries(projection.anchorIndex) as [string, any][])
          .find(([, value]) => value.scope === 'body' && value.kind === 'p')![0];
        let executions = 0;
        const batch = () => session.executeBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: () => { executions++; return session.insertParagraph(anchor, 'after', 'inserted exactly once'); } },
        ], 'atomic', { transactionId: 'tx-1', request: { insertAfter: anchor, text: 'inserted exactly once' } });
        const first = batch();
        const version = session.getVersion();
        const retry = batch();
        const markdown = session.project().markdown as string;
        return {
          executions,
          same: JSON.stringify(first) === JSON.stringify(retry),
          success: first.success,
          transaction: first.transaction,
          versionUnchanged: session.getVersion() === version,
          occurrences: markdown.split('inserted exactly once').length - 1,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.success).toBe(true);
    expect(outcome.executions).toBe(1);
    expect(outcome.same).toBe(true);
    expect(outcome.transaction.transactionId).toBe('tx-1');
    expect(outcome.transaction.requestFingerprint).toMatch(/^sha256:/);
    expect(outcome.versionUnchanged).toBe(true);
    expect(outcome.occurrences).toBe(1);
  });

  test('reusing an id for a different request is a conflict that changes nothing', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const anchor = (Object.entries(projection.anchorIndex) as [string, any][])
          .find(([, value]) => value.scope === 'body' && value.kind === 'p')![0];
        session.executeBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: () => session.insertParagraph(anchor, 'after', 'first') },
        ], 'atomic', { transactionId: 'tx-2', request: { text: 'first' } });
        const version = session.getVersion();
        let executed = false;
        const conflict = session.executeBatch([
          { tool: 'docx_create', action: 'insert_paragraph',
            mutation: () => { executed = true; return session.insertParagraph(anchor, 'after', 'second'); } },
        ], 'atomic', { transactionId: 'tx-2', request: { text: 'second' } });
        return {
          executed,
          success: conflict.success,
          code: conflict.failure?.error.code,
          transactionId: conflict.transaction?.transactionId,
          versionUnchanged: session.getVersion() === version,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.executed).toBe(false);
    expect(outcome.success).toBe(false);
    expect(outcome.code).toBe('transaction_conflict');
    expect(outcome.transactionId).toBe('tx-2');
    expect(outcome.versionUnchanged).toBe(true);
  });

  test('a transaction needs a request object and a non-blank id', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const attempt = (transaction: any) => {
          try { session.executeBatch([], 'atomic', transaction); return 'ok'; }
          catch (error) { return error instanceof Error ? error.message : String(error); }
        };
        return {
          noRequest: attempt({ transactionId: 'tx-3' }),
          blank: attempt({ transactionId: '   ', request: {} }),
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.noRequest).toContain('request object');
    expect(outcome.blank).toContain('whitespace');
  });
});

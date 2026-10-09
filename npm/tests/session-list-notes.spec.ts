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

// Issue #1026: listNotes was reachable only through the raw bridge export the editor calls.
test.describe('DocxSession.listNotes (#1026)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('lists footnotes and endnotes in citation order', async ({ page }) => {
    const result = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const projection = session.project();
        const host = (Object.entries(projection.anchorIndex) as [string, any][])
          .find(([, v]) => v.scope === 'body' && ['p', 'h', 'li'].includes(v.kind))![0];
        const before = session.listNotes(false);
        const first = session.insertFootnote(host, 0, 'First note.');
        // Cited at the same offset, so it sits before the first and is cited first.
        const second = session.insertFootnote(host, 0, 'Earlier note.');
        const endnote = session.insertEndnote(host, 0, 'An endnote.');
        const def = (r: any, kind: string) => r.created.find((a: any) => a.kind === kind).id;
        return {
          before,
          footnotes: session.listNotes(false),
          endnotes: session.listNotes(true),
          expectedFootnotes: [def(second, 'fn'), def(first, 'fn')],
          expectedEndnote: def(endnote, 'en'),
        };
      } finally {
        session.close();
      }
    }, Array.from(fixture));

    expect(result.before).toEqual([]);
    expect(result.footnotes.map((n: any) => n.ordinal)).toEqual([1, 2]);
    expect(result.footnotes.map((n: any) => n.defAnchorId)).toEqual(result.expectedFootnotes);
    expect(result.footnotes.every((n: any) => typeof n.id === 'string' && n.id.length > 0)).toBe(true);
    expect(result.endnotes.map((n: any) => n.defAnchorId)).toEqual([result.expectedEndnote]);
  });
});

import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const TEST_FILES_DIR = path.join(__dirname, '../../TestFiles');

function readTestFile(relativePath: string): number[] {
  return Array.from(new Uint8Array(fs.readFileSync(path.join(TEST_FILES_DIR, relativePath))));
}

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

// Issue #1024 — an argument the caller leaves out takes the engine's default, declared once in
// DocxSessionOps. The npm wrapper sends null for it through a nullable `[JSExport]` parameter instead
// of restating the default, so each omitted call must answer exactly as the call that names the
// default does. These go through `window.Docxodus.openTypedSession(...)`, the shipped wrapper.
test.describe('DocxSession omitted arguments take the engine default (Issue #1024)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('queries answer the same with the argument omitted and named', async ({ page }) => {
    const result = await page.evaluate(async (docxArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
      try {
        const anchor = Object.keys(session.project().anchorIndex).find((id) => id.startsWith('p:body:'))!;
        const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
        return {
          diffIsArray: Array.isArray(session.getDiff()),
          diff: same(session.getDiff(), session.getDiff(0)),
          projectAnchor: same(session.projectAnchor(anchor), session.projectAnchor(anchor, 2)),
          findByRegex: same(session.findByRegex('the'), session.findByRegex('the', 0)),
          findPlaceholders: same(session.findPlaceholders(), session.findPlaceholders(7, 1, undefined, 0)),
          remainingPlaceholders: same(session.remainingPlaceholders(), session.remainingPlaceholders(7)),
          renderBlock: same(
            session.renderBlock(anchor),
            session.renderBlock(anchor, { cssPrefix: 'docx-', fabricateClasses: false }),
          ),
        };
      } finally {
        session.close();
      }
    }, readTestFile('HC001-5DayTourPlanTemplate.docx'));

    expect(result).toEqual({
      diffIsArray: true,
      diff: true,
      projectAnchor: true,
      findByRegex: true,
      findPlaceholders: true,
      remainingPlaceholders: true,
      renderBlock: true,
    });
  });

  test('the wrapper sends null for an omitted argument instead of its own default', async ({ page }) => {
    const sent = await page.evaluate(async (docxArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
      try {
        const anchor = Object.keys(session.project().anchorIndex).find((id) => id.startsWith('p:body:'))!;
        const wasm = session.wasm;
        const calls: Record<string, unknown[]> = {};
        for (const name of ['GetDiff', 'ProjectAnchor', 'FindByRegex', 'FindPlaceholders',
          'RemainingPlaceholders', 'RenderBlockHtml', 'InsertPageNumberField']) {
          const real = wasm[name];
          wasm[name] = (...args: unknown[]) => { calls[name] = args.slice(1); return real.apply(wasm, args); };
        }
        session.getDiff();
        session.projectAnchor(anchor);
        session.findByRegex('the');
        session.findPlaceholders();
        session.remainingPlaceholders();
        session.renderBlock(anchor);
        session.insertPageNumberField(anchor);
        return calls;
      } finally {
        session.close();
      }
    }, readTestFile('HC001-5DayTourPlanTemplate.docx'));

    expect(sent.GetDiff).toEqual([null]);
    expect(sent.ProjectAnchor![1]).toBeNull();
    expect(sent.FindByRegex![1]).toBeNull();
    expect(sent.FindPlaceholders).toEqual([null, null, null, null]);
    expect(sent.RemainingPlaceholders).toEqual([null]);
    expect(sent.RenderBlockHtml!.slice(1)).toEqual([null, null]);
    // The bridge's positional convention: an empty string is an omitted token.
    expect(sent.InsertPageNumberField![1]).toBe('');
  });

  test('setRepeatHeaderRow without repeat marks the row, and false unmarks it', async ({ page }) => {
    const result = await page.evaluate(async (docxArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
      try {
        const anchor = Object.keys(session.project().anchorIndex).find((id) => id.startsWith('p:body:'))!;
        const inserted = session.insertTable(anchor, 'after', 2, 2);
        const cell = inserted.created[0].id;
        const table = inserted.tableAnchors.added.find((a: any) => a.entityKind === 'table').anchor.id;
        const marked = session.setRepeatHeaderRow(cell);
        const markedXml = session.raw.getXml(table);
        const unmarked = session.setRepeatHeaderRow(cell, false);
        return {
          marked: marked.success,
          headerAfterMark: markedXml.includes('tblHeader'),
          unmarked: unmarked.success,
          headerAfterUnmark: session.raw.getXml(table).includes('tblHeader'),
        };
      } finally {
        session.close();
      }
    }, readTestFile('HC001-5DayTourPlanTemplate.docx'));

    expect(result).toEqual({ marked: true, headerAfterMark: true, unmarked: true, headerAfterUnmark: false });
  });

  test('insertPageNumberField without a field inserts the current page', async ({ page }) => {
    const result = await page.evaluate(async (docxArray: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(docxArray));
      try {
        const anchor = Object.keys(session.project().anchorIndex).find((id) => id.startsWith('p:body:'))!;
        const inserted = session.insertPageNumberField(anchor);
        return { success: inserted.success, xml: session.raw.getXml(anchor) };
      } finally {
        session.close();
      }
    }, readTestFile('HC001-5DayTourPlanTemplate.docx'));

    expect(result.success).toBe(true);
    expect(result.xml).toMatch(/PAGE/);
    expect(result.xml).not.toMatch(/NUMPAGES/);
  });
});

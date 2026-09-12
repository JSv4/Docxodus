import { test, expect, Page } from '@playwright/test';
import * as fs from 'fs';
import * as path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const fixture = new Uint8Array(fs.readFileSync(
  path.join(__dirname, '../../TestFiles/CC763-NestedContentControls.docx'),
));

async function waitForDocxodus(page: Page) {
  await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
}

test.describe('Content controls: nested fills and tracked mutations (#763)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('preserve fills around a nested child and the matrix names what applies', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(new Uint8Array(bytes));
      try {
        const control = (nativeId: string) =>
          session.listContentControls().find((c: any) => c.nativeId === nativeId);
        const outer = control('100');
        const inner = control('101');
        const refused = session.fillContentControlText(outer.anchorId, 'x');
        const preserved = session.fillContentControlText(outer.anchorId, 'Outer via npm', {
          nestedControls: 'preserve',
          childFills: { [inner.anchorId]: 'Inner via npm' },
        });
        return {
          nested: outer.nestedControlAnchorIds,
          innerAnchor: inner.anchorId,
          preserveEntry: outer.operations.find((op: any) => op.operation === 'fill_text' && op.nestedControls === 'preserve'),
          refuseEntry: outer.operations.find((op: any) => op.operation === 'fill_text' && op.nestedControls === 'refuse'),
          refusedCode: refused.error?.code,
          preservedSuccess: preserved.success,
          modified: preserved.modified.map((a: any) => a.id),
          innerText: control('101').text,
          outerText: control('100').text,
          tag: control('100').tag,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.nested).toEqual([outcome.innerAnchor]);
    expect(outcome.preserveEntry.canMutate).toBe(true);
    expect(outcome.refuseEntry.canMutate).toBe(false);
    expect(outcome.refusedCode).toBe('content_control_nested_fill_unsupported');
    expect(outcome.preservedSuccess).toBe(true);
    expect(outcome.modified).toEqual([outcome.modified[0], outcome.innerAnchor]);
    expect(outcome.innerText).toBe('Inner via npm');
    expect(outcome.outerText).toContain('Outer via npm');
    expect(outcome.tag).toBe('outer-tag');
  });

  test('tracked text fills record revisions and state changes refuse with a reason', async ({ page }) => {
    const outcome = await page.evaluate((bytes: number[]) => {
      const session = (window as any).Docxodus.openTypedSession(
        new Uint8Array(bytes), JSON.stringify({ trackedChanges: 'render_inline' }));
      try {
        const control = (nativeId: string) =>
          session.listContentControls().find((c: any) => c.nativeId === nativeId);
        const inner = control('101');
        const checkbox = control('102');
        const filled = session.fillContentControlText(inner.anchorId, 'tracked via npm');
        const types = session.listRevisions().map((r: any) => r.type);
        const refused = session.setContentControlChecked(checkbox.anchorId, true);
        return {
          innerCanMutate: inner.canMutate,
          checkboxCanMutate: checkbox.canMutate,
          checkboxReason: checkbox.unsupportedReason,
          filledSuccess: filled.success,
          types,
          refusedCode: refused.error?.code,
        };
      } finally { session.close(); }
    }, Array.from(fixture));
    expect(outcome.innerCanMutate).toBe(true);
    expect(outcome.checkboxCanMutate).toBe(false);
    expect(outcome.checkboxReason).toContain('w14:checked');
    expect(outcome.filledSuccess).toBe(true);
    expect(outcome.types).toContain('insert');
    expect(outcome.types).toContain('delete');
    expect(outcome.refusedCode).toBe('tracked_operation_unsupported');
  });
});

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

// A PNG signature + IHDR is all the bridge's format/dimension sniffing needs.
function png(width: number, height: number): number[] {
  const bytes = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52];
  for (const value of [width, height]) {
    bytes.push((value >>> 24) & 0xff, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff);
  }
  return bytes;
}

// Issue #452 — native content controls across the WASM bridge. HC030 is Word-authored and
// carries five controls: rich text, plain text, picture, checkbox, combo box.
test.describe('DocxSession content controls (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('ListContentControls decodes the Word-authored registry and honors scopes', async ({ page }) => {
    const bytes = readTestFile('HC030-Content-Controls.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(new Uint8Array(bytesArray), '');
      try {
        const all = JSON.parse(bridge.ListContentControls(handle, 0x3f));
        const body = JSON.parse(bridge.ListContentControls(handle, 0x01));
        const headers = JSON.parse(bridge.ListContentControls(handle, 0x02));
        return {
          types: all.map((c: any) => c.type),
          placements: all.map((c: any) => c.placement),
          anchors: all.map((c: any) => c.anchorId),
          scopes: all.map((c: any) => c.scope),
          canMutate: all.map((c: any) => c.canMutate),
          unsupported: all.map((c: any) => c.unsupportedReason ?? null),
          bodyCount: body.length,
          headerCount: headers.length,
          comboItems: all.find((c: any) => c.type === 'combo_box')?.itemValues,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.types).toEqual(['rich_text', 'plain_text', 'picture', 'checkbox', 'combo_box']);
    expect(result.placements).toEqual(['block', 'inline', 'block', 'block', 'block']);
    expect(result.anchors.every((a: string) => a.startsWith('sdt:body:'))).toBe(true);
    expect(result.scopes).toEqual(['body', 'body', 'body', 'body', 'body']);
    expect(result.canMutate).toEqual([true, true, true, true, true]);
    expect(result.unsupported).toEqual([null, null, null, null, null]);
    expect(result.bodyCount).toBe(5);
    expect(result.headerCount).toBe(0);
    expect(result.comboItems).toEqual(['One', 'Two', 'Three']);
  });

  test('every fill route mutates through the bridge and survives save/reopen', async ({ page }) => {
    const bytes = readTestFile('HC030-Content-Controls.docx');

    const result = await page.evaluate(
      async ({ bytesArray, image }: { bytesArray: number[]; image: number[] }) => {
        const bridge = (window as any).Docxodus.DocxSessionBridge;
        const handle = bridge.OpenSession(new Uint8Array(bytesArray), '');
        try {
          const controls = JSON.parse(bridge.ListContentControls(handle, 0x3f));
          const byType = (type: string) =>
            controls.find((c: any) => c.type === type).anchorId as string;
          const plain = byType('plain_text');
          const rich = byType('rich_text');
          const checkbox = byType('checkbox');
          const combo = byType('combo_box');
          const picture = byType('picture');
          const options = '{}';

          const text = JSON.parse(
            bridge.FillContentControlText(handle, plain, 'bridge plain value', options),
          );
          const markdown = JSON.parse(
            bridge.FillContentControlRichText(handle, rich, 'bridge **rich** value', options),
          );
          const checked = JSON.parse(
            bridge.SetContentControlChecked(handle, checkbox, true, options),
          );
          const selected = JSON.parse(
            bridge.SelectContentControlItem(handle, combo, 'Three', options),
          );
          const filledPicture = JSON.parse(
            bridge.FillContentControlPicture(
              handle,
              picture,
              btoa(String.fromCharCode(...image)),
              options,
            ),
          );

          const saved = bridge.Save(handle);
          const reopenedHandle = bridge.OpenSession(saved, '');
          let reopened: any[] = [];
          try {
            reopened = JSON.parse(bridge.ListContentControls(reopenedHandle, 0x3f));
          } finally {
            bridge.CloseSession(reopenedHandle);
          }

          return {
            successes: [text, markdown, checked, selected, filledPicture].map((r) => r.success),
            errors: [text, markdown, checked, selected, filledPicture].map(
              (r) => r.error?.code ?? null,
            ),
            modifiedIsTarget: text.modified?.[0]?.id === plain,
            // The anchor derives from the native w:sdtPr/w:id, so it survives a clean save.
            reopenedText: Object.fromEntries(reopened.map((c: any) => [c.anchorId, c.text])),
            plain,
            rich,
            checkbox,
            combo,
          };
        } finally {
          bridge.CloseSession(handle);
        }
      },
      { bytesArray: Array.from(bytes), image: png(4, 5) },
    );

    expect(result.errors).toEqual([null, null, null, null, null]);
    expect(result.successes).toEqual([true, true, true, true, true]);
    expect(result.modifiedIsTarget).toBe(true);
    expect(result.reopenedText[result.plain]).toBe('bridge plain value');
    expect(result.reopenedText[result.rich]).toBe('bridge rich value');
    expect(result.reopenedText[result.checkbox]).toBe('☒');
    expect(result.reopenedText[result.combo]).toBe('Three');
  });

  test('the date and repeating-section routes are wired and typed', async ({ page }) => {
    const bytes = readTestFile('HC030-Content-Controls.docx');

    // HC030 has no date or repeating-section control, so these routes are proved reachable
    // by the engine's typed rejection of a well-formed call against a wrong-typed target.
    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(new Uint8Array(bytesArray), '');
      try {
        const controls = JSON.parse(bridge.ListContentControls(handle, 0x3f));
        const plain = controls.find((c: any) => c.type === 'plain_text').anchorId as string;

        return {
          wrongTypeDate: JSON.parse(
            bridge.SetContentControlDate(handle, plain, '2026-08-14T00:00:00Z', null, '{}'),
          ).error?.code,
          badDateValue: JSON.parse(
            bridge.SetContentControlDate(handle, plain, 'not-a-timestamp', 'August 2026', '{}'),
          ).error?.code,
          addItem: JSON.parse(bridge.AddRepeatingSectionItem(handle, plain, '', '{}')).error?.code,
          removeItem: JSON.parse(bridge.RemoveRepeatingSectionItem(handle, plain)).error?.code,
          unknownAnchor: JSON.parse(
            bridge.RemoveRepeatingSectionItem(handle, 'sdt:body:deadbeef'),
          ).error?.code,
          badOptions: JSON.parse(
            bridge.FillContentControlText(handle, plain, 'x', '{"bindingPolicy":"nonsense"}'),
          ).error?.code,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.wrongTypeDate).toBe('content_control_wrong_type');
    expect(result.badDateValue).toBe('invalid_content_control_value');
    expect(result.addItem).toBe('content_control_wrong_type');
    expect(result.removeItem).toBe('content_control_wrong_type');
    expect(result.unknownAnchor).toBe('content_control_not_found');
    expect(result.badOptions).toBe('invalid_content_control_value');
  });

  test('render_inline tracked mode is reported by the registry, not only at mutation time', async ({
    page,
  }) => {
    const bytes = readTestFile('HC030-Content-Controls.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(new Uint8Array(bytesArray), '');
      try {
        const before = JSON.parse(bridge.ListContentControls(handle, 0x3f));
        bridge.SetTrackedChanges(handle, 1); // TrackedChangeMode.RenderInline
        const tracked = JSON.parse(bridge.ListContentControls(handle, 0x3f));
        const plain = tracked.find((c: any) => c.type === 'plain_text').anchorId as string;
        const attempted = JSON.parse(bridge.FillContentControlText(handle, plain, 'x', '{}'));
        return {
          beforeMutable: before.map((c: any) => c.canMutate),
          trackedMutable: tracked.map((c: any) => c.canMutate),
          trackedReasons: tracked.map((c: any) => c.unsupportedReason ?? ''),
          attemptedCode: attempted.error?.code,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.beforeMutable).toEqual([true, true, true, true, true]);
    // Discovery must agree with what a fill actually does — an agent plans off this registry.
    expect(result.trackedMutable).toEqual([false, false, false, false, false]);
    expect(result.trackedReasons.every((r: string) => r.includes('tracked revisions'))).toBe(true);
    expect(result.attemptedCode).toBe('tracked_operation_unsupported');
  });
});

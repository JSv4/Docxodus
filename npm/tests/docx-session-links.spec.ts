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

// ProjectionScopes flag mask, mirrored from npm/src/types.ts so the spec can pass it
// to the bridge without importing the bundle.
const SCOPE_ALL = 63;
const SCOPE_BODY = 1;
const SCOPE_HEADERS = 2;

// Issue #451 — native hyperlinks and bookmarks through the WASM bridge. Each spec normalizes
// the paragraphs it touches to known text, so every span below is a literal offset rather than
// an assumption about the fixture's wording.
test.describe('DocxSession hyperlinks and bookmarks (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('external hyperlink CRUD reuses one relationship and survives save/reopen', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const paragraphs = Object.keys(proj.anchorIndex).filter(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:') || k.startsWith('li:body:'),
        );
        bridge.ReplaceText(handle, paragraphs[0], 'Alpha beta gamma delta');
        bridge.ReplaceText(handle, paragraphs[1], 'Second paragraph text');

        const first = JSON.parse(
          bridge.AddHyperlink(handle, paragraphs[0], 0, 5, 'external', 'https://example.test/shared'),
        );
        const second = JSON.parse(
          bridge.AddHyperlink(handle, paragraphs[1], 0, 6, 'external', 'https://example.test/shared'),
        );
        const listed = JSON.parse(bridge.ListHyperlinks(handle, 63));

        const updated = JSON.parse(
          bridge.UpdateHyperlink(handle, first.hyperlinkId, 'external', 'https://example.test/moved'),
        );
        const afterUpdate = JSON.parse(bridge.ListHyperlinks(handle, 63));
        const removed = JSON.parse(bridge.RemoveHyperlink(handle, first.hyperlinkId));
        const afterRemove = JSON.parse(bridge.ListHyperlinks(handle, 63));

        const saved = bridge.SaveWithAnchorIds(handle);
        const handle2 = bridge.OpenSession(saved, '');
        let reopened: any[] = [];
        try {
          reopened = JSON.parse(bridge.ListHyperlinks(handle2, 63));
        } finally {
          bridge.CloseSession(handle2);
        }

        return {
          firstOk: first.success,
          firstError: first.error?.code,
          secondOk: second.success,
          listedCount: listed.length,
          listedText: listed[0]?.text,
          listedKinds: listed.map((l: any) => l.kind),
          distinctRelationships: new Set(listed.map((l: any) => l.relationshipId)).size,
          anyBroken: listed.some((l: any) => l.isBroken),
          updatedOk: updated.success,
          targetsAfterUpdate: afterUpdate.map((l: any) => l.target).sort(),
          removedOk: removed.success,
          idsAfterRemove: afterRemove.map((l: any) => l.id),
          survivorId: second.hyperlinkId,
          reopenedIds: reopened.map((l: any) => l.id),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.firstOk, `error=${result.firstError}`).toBe(true);
    expect(result.secondOk).toBe(true);
    expect(result.listedCount).toBe(2);
    expect(result.listedText).toBe('Alpha');
    expect(result.listedKinds).toEqual(['external', 'external']);
    // One URI, one owning-part relationship, reused by both links.
    expect(result.distinctRelationships).toBe(1);
    expect(result.anyBroken).toBe(false);
    expect(result.updatedOk).toBe(true);
    expect(result.targetsAfterUpdate).toEqual([
      'https://example.test/moved',
      'https://example.test/shared',
    ]);
    expect(result.removedOk).toBe(true);
    expect(result.idsAfterRemove).toEqual([result.survivorId]);
    expect(result.reopenedIds).toEqual([result.survivorId]);
  });

  test('rename retargets inbound internal links and removal refuses a live target', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const paragraphs = Object.keys(proj.anchorIndex).filter(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:') || k.startsWith('li:body:'),
        );
        bridge.ReplaceText(handle, paragraphs[0], 'Alpha beta gamma delta');
        bridge.ReplaceText(handle, paragraphs[1], 'Second paragraph text');

        const added = JSON.parse(
          bridge.AddBookmark(handle, 'TargetOne', paragraphs[0], 0, paragraphs[0], 5),
        );
        const link = JSON.parse(
          bridge.AddHyperlink(handle, paragraphs[1], 0, 6, 'internal', 'TargetOne'),
        );
        const internal = JSON.parse(bridge.ListHyperlinks(handle, 63)).find(
          (l: any) => l.id === link.hyperlinkId,
        );

        const blocked = JSON.parse(bridge.RemoveBookmark(handle, 'TargetOne'));
        const renamed = JSON.parse(bridge.RenameBookmark(handle, 'TargetOne', 'TargetTwo'));
        const retargeted = JSON.parse(bridge.ListHyperlinks(handle, 63)).find(
          (l: any) => l.id === link.hyperlinkId,
        );

        // Releasing the last inbound reference releases the bookmark.
        const detached = JSON.parse(
          bridge.UpdateHyperlink(handle, link.hyperlinkId, 'external', 'https://example.test/out'),
        );
        const finallyRemoved = JSON.parse(bridge.RemoveBookmark(handle, 'TargetTwo'));

        return {
          addedOk: added.success,
          addedError: added.error?.code,
          linkOk: link.success,
          internalKind: internal?.kind,
          internalTarget: internal?.target,
          internalRelationshipId: internal?.relationshipId ?? null,
          internalBroken: internal?.isBroken,
          blockedOk: blocked.success,
          blockedCode: blocked.error?.code,
          renamedOk: renamed.success,
          retargetedTarget: retargeted?.target,
          retargetedBroken: retargeted?.isBroken,
          detachedOk: detached.success,
          finallyRemovedOk: finallyRemoved.success,
          remainingBookmarks: JSON.parse(bridge.ListBookmarks(handle, 63)).map((b: any) => b.name),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.addedOk, `error=${result.addedError}`).toBe(true);
    expect(result.linkOk).toBe(true);
    expect(result.internalKind).toBe('internal');
    expect(result.internalTarget).toBe('TargetOne');
    // Internal links are relationship-free w:anchor markup.
    expect(result.internalRelationshipId).toBeNull();
    expect(result.internalBroken).toBe(false);
    expect(result.blockedOk).toBe(false);
    expect(result.blockedCode).toBe('bookmark_in_use');
    expect(result.renamedOk).toBe(true);
    expect(result.retargetedTarget).toBe('TargetTwo');
    expect(result.retargetedBroken).toBe(false);
    expect(result.detachedOk).toBe(true);
    expect(result.finallyRemovedOk).toBe(true);
    expect(result.remainingBookmarks).not.toContain('TargetTwo');
  });

  test('bookmark ranges report per-paragraph segments, move, and keep their id', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const paragraphs = Object.keys(proj.anchorIndex).filter(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:') || k.startsWith('li:body:'),
        );
        bridge.ReplaceText(handle, paragraphs[0], 'Alpha beta gamma delta');
        bridge.ReplaceText(handle, paragraphs[1], 'Second paragraph text');

        const added = JSON.parse(
          bridge.AddBookmark(handle, 'AcrossParas', paragraphs[0], 6, paragraphs[1], 6),
        );
        const before = JSON.parse(bridge.ListBookmarks(handle, 63)).find(
          (b: any) => b.name === 'AcrossParas',
        );
        const moved = JSON.parse(
          bridge.MoveBookmark(handle, 'AcrossParas', paragraphs[1], 0, paragraphs[1], 6),
        );
        const after = JSON.parse(bridge.ListBookmarks(handle, 63)).find(
          (b: any) => b.name === 'AcrossParas',
        );

        return {
          addedOk: added.success,
          addedError: added.error?.code,
          isValid: before?.isValid,
          validationError: before?.validationError,
          isPaired: before?.isPaired,
          isManaged: before?.isManaged,
          segmentCount: before?.segments?.length,
          text: before?.text,
          movedOk: moved.success,
          idBefore: before?.bookmarkId,
          idAfter: after?.bookmarkId,
          textAfter: after?.text,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.addedOk, `error=${result.addedError}`).toBe(true);
    expect(result.isValid, result.validationError).toBe(true);
    expect(result.isPaired).toBe(true);
    expect(result.isManaged).toBe(false);
    expect(result.segmentCount).toBe(2);
    expect(result.text).toBe('beta gamma delta\nSecond');
    expect(result.movedOk).toBe(true);
    // A same-part move keeps the pair's numeric id; only the coordinates change.
    expect(result.idAfter).toBe(result.idBefore);
    expect(result.textAfter).toBe('Second');
  });

  test('scoped listing, reserved names, and missing targets are structured', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(
      async ({ bytesArray, all, body, headers }: any) => {
        const bin = new Uint8Array(bytesArray);
        const bridge = (window as any).Docxodus.DocxSessionBridge;
        const handle = bridge.OpenSession(bin, '');
        try {
          const proj = JSON.parse(bridge.Project(handle));
          const paragraphs = Object.keys(proj.anchorIndex).filter(
            (k) => k.startsWith('p:body:') || k.startsWith('h:body:') || k.startsWith('li:body:'),
          );
          bridge.ReplaceText(handle, paragraphs[0], 'Alpha beta gamma delta');

          bridge.AddHyperlink(handle, paragraphs[0], 0, 5, 'external', 'https://example.test/body');
          bridge.SetHeaderText(handle, paragraphs[0], 'default', 'header line');
          const headerAnchor = Object.keys(JSON.parse(bridge.Project(handle)).anchorIndex).find(
            (k) => k.startsWith('p:hdr') || k.startsWith('h:hdr'),
          )!;
          bridge.AddHyperlink(handle, headerAnchor, 0, 6, 'external', 'https://example.test/header');

          const reserved = JSON.parse(
            bridge.AddBookmark(handle, '_Toc12345', paragraphs[0], 0, paragraphs[0], 5),
          );
          const missing = JSON.parse(
            bridge.AddHyperlink(handle, paragraphs[0], 6, 4, 'internal', 'NoSuchBookmark'),
          );
          const unknown = JSON.parse(bridge.RemoveHyperlink(handle, 'hl:body:deadbeef'));

          return {
            allScopes: JSON.parse(bridge.ListHyperlinks(handle, all)).length,
            bodyScopes: JSON.parse(bridge.ListHyperlinks(handle, body)).map((l: any) => l.scope),
            headerScopes: JSON.parse(bridge.ListHyperlinks(handle, headers)).map((l: any) => l.scope),
            reservedOk: reserved.success,
            reservedCode: reserved.error?.code,
            missingOk: missing.success,
            missingCode: missing.error?.code,
            unknownOk: unknown.success,
            unknownCode: unknown.error?.code,
            bookmarkNames: JSON.parse(bridge.ListBookmarks(handle, all)).map((b: any) => b.name),
          };
        } finally {
          bridge.CloseSession(handle);
        }
      },
      { bytesArray: Array.from(bytes), all: SCOPE_ALL, body: SCOPE_BODY, headers: SCOPE_HEADERS },
    );

    expect(result.allScopes).toBe(2);
    expect(result.bodyScopes).toEqual(['body']);
    expect(result.headerScopes.length).toBe(1);
    expect(result.headerScopes[0].startsWith('hdr')).toBe(true);
    // Word owns the _Toc*/_Ref*/_Hlt*/_Hlk*/_GoBack namespace and reallocates it.
    expect(result.reservedOk).toBe(false);
    expect(result.reservedCode).toBe('invalid_bookmark_name');
    expect(result.missingOk).toBe(false);
    expect(result.missingCode).toBe('missing_bookmark_target');
    expect(result.unknownOk).toBe(false);
    expect(result.unknownCode).toBe('hyperlink_not_found');
    expect(result.bookmarkNames).not.toContain('_Toc12345');
  });

  test('a hyperlink drawn over a bookmark relocates its markers instead of stranding them', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const paragraph = Object.keys(proj.anchorIndex).find(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:') || k.startsWith('li:body:'),
        )!;
        bridge.ReplaceText(handle, paragraph, 'Alpha beta gamma delta');

        // The bookmark sits strictly INSIDE the span the hyperlink will cover.
        const added = JSON.parse(bridge.AddBookmark(handle, 'Inner', paragraph, 6, paragraph, 10));
        const wrapped = JSON.parse(
          bridge.AddHyperlink(handle, paragraph, 0, 16, 'external', 'https://example.test/wrap'),
        );
        const bookmark = JSON.parse(bridge.ListBookmarks(handle, 63)).find(
          (b: any) => b.name === 'Inner',
        );
        // A stranded start would land after its own end and stop resolving, which makes the
        // bookmark permanently unmutatable — so mutability is the real assertion.
        const renamed = JSON.parse(bridge.RenameBookmark(handle, 'Inner', 'Renamed'));
        const removed = JSON.parse(bridge.RemoveBookmark(handle, 'Renamed'));

        return {
          addedOk: added.success,
          wrappedOk: wrapped.success,
          wrappedError: wrapped.error?.code,
          isValid: bookmark?.isValid,
          validationError: bookmark?.validationError,
          span: bookmark?.segments?.[0]?.span,
          text: bookmark?.text,
          renamedOk: renamed.success,
          removedOk: removed.success,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.addedOk).toBe(true);
    expect(result.wrappedOk, `error=${result.wrappedError}`).toBe(true);
    expect(result.isValid, result.validationError).toBe(true);
    expect(result.span).toEqual({ start: 6, length: 4 });
    expect(result.text).toBe('beta');
    expect(result.renamedOk).toBe(true);
    expect(result.removedOk).toBe(true);
  });
});

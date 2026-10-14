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

// Issues #300/#317 — native Word comment authoring and threading through the WASM bridge.
test.describe('DocxSession comment authoring (WASM bridge)', () => {
  test.beforeEach(async ({ page }) => {
    await page.goto('/test-harness.html');
    await waitForDocxodus(page);
  });

  test('AddComment creates the comment, lists it, projects it, and round-trips', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find(
          (k) => k.startsWith('p:body:') || k.startsWith('h:body:'),
        )!;

        const made = JSON.parse(
          bridge.AddComment(handle, bodyAnchor, '', 'Alice', 'AL', '2026-08-01T00:00:00Z', 'Needs review.'),
        );
        const created: { id: string; kind: string; scope: string }[] = made.created ?? [];

        const listed = JSON.parse(bridge.ListComments(handle));
        const after = JSON.parse(bridge.Project(handle));
        const saved = bridge.Save(handle);

        // Reopen the saved bytes in a fresh session to prove the comment persisted.
        const handle2 = bridge.OpenSession(saved, '');
        let reopened: any[] = [];
        try {
          reopened = JSON.parse(bridge.ListComments(handle2));
        } finally {
          bridge.CloseSession(handle2);
        }

        return {
          success: made.success,
          errorCode: made.error?.code,
          hasDefAnchor: created.some((a) => a.kind === 'cmt' && a.scope === 'cmt'),
          hasParaAnchor: created.some((a) => a.kind === 'p' && a.scope === 'cmt'),
          modifiedHost: (made.modified ?? []).some((a: any) => a.id === bodyAnchor),
          listedAuthor: listed[0]?.author,
          listedInitials: listed[0]?.initials,
          listedDate: listed[0]?.date,
          listedText: listed[0]?.text,
          markdownHasCommentsSection: after.markdown.includes('# Comments'),
          markdownHasCommentText: after.markdown.includes('Needs review.'),
          reopenedCount: reopened.length,
          reopenedText: reopened[0]?.text,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.success, `error=${result.errorCode}`).toBe(true);
    expect(result.hasDefAnchor).toBe(true);
    expect(result.hasParaAnchor).toBe(true);
    expect(result.modifiedHost).toBe(true);
    expect(result.listedAuthor).toBe('Alice');
    expect(result.listedInitials).toBe('AL');
    expect(result.listedDate).toBe('2026-08-01T00:00:00Z');
    expect(result.listedText).toBe('Needs review.');
    expect(result.markdownHasCommentsSection).toBe(true);
    expect(result.markdownHasCommentText).toBe(true);
    expect(result.reopenedCount).toBe(1);
    expect(result.reopenedText).toBe('Needs review.');
  });

  test('UpdateComment and RemoveComment round-trip through the bridge', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const made = JSON.parse(bridge.AddComment(handle, bodyAnchor, '', 'Bob', '', '', 'Original.'));
        const cmtAnchor = (made.created ?? []).find((a: any) => a.kind === 'cmt')!.id;

        const updated = JSON.parse(bridge.UpdateComment(handle, cmtAnchor, 'Revised body.'));
        const afterUpdate = JSON.parse(bridge.ListComments(handle));

        const removed = JSON.parse(bridge.RemoveComment(handle, cmtAnchor));
        const afterRemove = JSON.parse(bridge.ListComments(handle));
        const markdownAfter = JSON.parse(bridge.Project(handle)).markdown;

        return {
          updateSuccess: updated.success,
          updatedText: afterUpdate[0]?.text,
          authorPreserved: afterUpdate[0]?.author,
          removeSuccess: removed.success,
          removedListEmpty: afterRemove.length === 0,
          commentsSectionGone: !markdownAfter.includes('# Comments'),
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.updateSuccess).toBe(true);
    expect(result.updatedText).toBe('Revised body.');
    expect(result.authorPreserved).toBe('Bob');
    expect(result.removeSuccess).toBe(true);
    expect(result.removedListEmpty).toBe(true);
    expect(result.commentsSectionGone).toBe(true);
  });

  test('AddCommentReply and SetCommentResolved preserve native thread state', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;
        const made = JSON.parse(bridge.AddComment(handle, bodyAnchor, '', 'Alice', '', '', 'Parent.'));
        const parentAnchor = (made.created ?? []).find((a: any) => a.kind === 'cmt')!.id;

        const replied = JSON.parse(
          bridge.AddCommentReply(handle, parentAnchor, 'Bob', 'BO', '', 'Reply.'),
        );
        const replyAnchor = (replied.created ?? []).find((a: any) => a.kind === 'cmt')!.id;
        const beforeResolve = JSON.parse(bridge.ListComments(handle));

        const resolved = JSON.parse(bridge.SetCommentResolved(handle, replyAnchor, true));
        const afterResolve = JSON.parse(bridge.ListComments(handle));
        const reopened = JSON.parse(bridge.SetCommentResolved(handle, replyAnchor, false));
        const saved = bridge.Save(handle);

        const handle2 = bridge.OpenSession(saved, '');
        let afterReopen: any[] = [];
        try {
          afterReopen = JSON.parse(bridge.ListComments(handle2));
        } finally {
          bridge.CloseSession(handle2);
        }

        const find = (entries: any[], anchorId: string) =>
          entries.find((entry: any) => entry.anchorId === anchorId);
        const reopenedParent = afterReopen.find((entry: any) => entry.author === 'Alice');
        const reopenedReply = afterReopen.find((entry: any) => entry.author === 'Bob');
        return {
          replySuccess: replied.success,
          parentResolved: find(beforeResolve, parentAnchor)?.resolved,
          replyParentMatches: find(beforeResolve, replyAnchor)?.parentAnchorId === parentAnchor,
          replyInitiallyResolved: find(beforeResolve, replyAnchor)?.resolved,
          resolveSuccess: resolved.success,
          replyResolved: find(afterResolve, replyAnchor)?.resolved,
          reopenSuccess: reopened.success,
          reopenedParentMatches: reopenedReply?.parentAnchorId === reopenedParent?.anchorId,
          reopenedResolved: reopenedReply?.resolved,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.replySuccess).toBe(true);
    expect(result.parentResolved).toBe(false);
    expect(result.replyParentMatches).toBe(true);
    expect(result.replyInitiallyResolved).toBe(false);
    expect(result.resolveSuccess).toBe(true);
    expect(result.replyResolved).toBe(true);
    expect(result.reopenSuccess).toBe(true);
    expect(result.reopenedParentMatches).toBe(true);
    expect(result.reopenedResolved).toBe(false);
  });

  test('AddComment error envelope: typed codes for empty span and wrong-kind anchors', async ({ page }) => {
    const bytes = readTestFile('HC001-5DayTourPlanTemplate.docx');

    const result = await page.evaluate(async (bytesArray: number[]) => {
      const bin = new Uint8Array(bytesArray);
      const bridge = (window as any).Docxodus.DocxSessionBridge;
      const handle = bridge.OpenSession(bin, '');
      try {
        const proj = JSON.parse(bridge.Project(handle));
        const bodyAnchor = Object.keys(proj.anchorIndex).find((k) => k.startsWith('p:body:'))!;

        const emptySpan = JSON.parse(
          bridge.AddComment(handle, bodyAnchor, '{"start":0,"length":0}', 'A', '', '', 'x'),
        );

        const made = JSON.parse(bridge.AddComment(handle, bodyAnchor, '', 'A', '', '', 'Host.'));
        const cmtPara = (made.created ?? []).find((a: any) => a.kind === 'p' && a.scope === 'cmt')!.id;
        // Word has no comments-on-comments: a cmt-scope paragraph is not a legal host.
        const nested = JSON.parse(bridge.AddComment(handle, cmtPara, '', 'A', '', '', 'Nested.'));

        return {
          emptySpanSuccess: emptySpan.success,
          emptySpanCode: emptySpan.error?.code,
          nestedSuccess: nested.success,
          nestedCode: nested.error?.code,
        };
      } finally {
        bridge.CloseSession(handle);
      }
    }, Array.from(bytes));

    expect(result.emptySpanSuccess).toBe(false);
    expect(result.emptySpanCode).toBe('empty_comment_span');
    expect(result.nestedSuccess).toBe(false);
    expect(result.nestedCode).toBe('anchor_wrong_kind');
  });
});

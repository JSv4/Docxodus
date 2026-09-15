import { test, expect } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

/**
 * Main-thread `getComments` from docxodus/core: the document-level read must return what a
 * session's `listComments()` returns — minus anchors, which are minted per session — and must
 * surface the bridge's error envelope instead of swallowing it.
 */
test("getComments matches session.listComments() without opening a session", async ({ page }) => {
  const here = path.dirname(fileURLToPath(import.meta.url));
  const bytes = fs.readFileSync(path.join(here, "../../TestFiles/DD/DD002-DenseComments.docx"));

  await page.goto("http://localhost:8083/");
  const result = await page.evaluate(async (byteValues) => {
    const moduleUrl = "http://localhost:8083/embed.bundle.js";
    const api = await import(moduleUrl);
    await api.initialize("http://localhost:8083/wasm/");
    const input = new Uint8Array(byteValues);

    const stateless = await api.getComments(input);
    const session = api.openDocxSession(input);
    const viaSession = session.listComments();
    session.close();

    // Anchors differ per session, so compare content and the reply link by index.
    const shape = (list: any[]) => list.map((c) => ({
      id: c.id, author: c.author, initials: c.initials, date: c.date, text: c.text, resolved: c.resolved,
      parentIndex: c.parentAnchorId ? list.findIndex((p) => p.anchorId === c.parentAnchorId) : -1,
    }));

    let invalidError = "";
    try { await api.getComments(new Uint8Array([1, 2, 3])); } catch (e) { invalidError = String((e as Error).message); }
    return { stateless: shape(stateless), viaSession: shape(viaSession), invalidError };
  }, Array.from(bytes));

  expect(result.stateless).toHaveLength(10);
  expect(result.stateless).toEqual(result.viaSession);
  expect(result.stateless[7].parentIndex).toBe(6);
  expect(result.invalidError).toMatch(/^Failed to get comments: /);
});

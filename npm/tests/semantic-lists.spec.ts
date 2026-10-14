import { test, expect, Page } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";

/**
 * Semantic lists (issue #895): `semanticLists` turns Word list paragraphs into <ol>/<ul>/<li>, through
 * the WASM export and through the worker. DB012 holds three short decimal lists separated by plain
 * paragraphs.
 */
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const fixture = Array.from(
  fs.readFileSync(path.join(__dirname, "../../TestFiles/DB012-Lists-With-Different-Numberings.docx")),
);

interface ListShape {
  lists: number;
  items: number;
  itemsOutsideLists: number;
  listParagraphs: number;
}

/** Parse converted HTML in the page and count its list structure. */
async function shape(page: Page, html: string): Promise<ListShape> {
  return page.evaluate((source) => {
    const doc = new DOMParser().parseFromString(source, "text/html");
    const items = Array.from(doc.querySelectorAll("li"));
    return {
      lists: doc.querySelectorAll("ol, ul").length,
      items: items.length,
      itemsOutsideLists: items.filter((li) => !li.parentElement?.matches("ol, ul")).length,
      listParagraphs: doc.querySelectorAll("p > [data-list-marker]").length,
    };
  }, html);
}

test.describe("semantic lists (#895)", () => {
  test("the WASM export groups list paragraphs only when asked", async ({ page }) => {
    await page.goto("/test-harness.html");
    await page.waitForFunction(() => (window as any).DocxodusReady === true, { timeout: 30000 });
    const [plain, semantic] = await page.evaluate((bytesArray: number[]) => {
      const converter = (window as any).Docxodus.DocumentConverter;
      const convert = (...trailing: unknown[]) => converter.ConvertDocxToHtmlComplete(
        new Uint8Array(bytesArray), "Document", "docx-", true, "", -1, "comment-", 0, 1.0, "page-",
        false, 0, "annot-", false, false, false, true, true, false, null, true, 0, ...trailing);
      return [convert(false), convert(true)];
    }, fixture);

    expect(await shape(page, plain)).toEqual({ lists: 0, items: 0, itemsOutsideLists: 0, listParagraphs: 9 });
    const grouped = await shape(page, semantic);
    expect(grouped).toEqual({ lists: 3, items: 9, itemsOutsideLists: 0, listParagraphs: 0 });
    // Each item keeps its block anchor.
    const anchored = await page.evaluate((source) =>
      Array.from(new DOMParser().parseFromString(source, "text/html").querySelectorAll("li"))
        .every((li) => li.hasAttribute("data-anchor")), semantic);
    expect(anchored).toBe(true);
  });

  test("the worker carries semanticLists to the export", async ({ page }) => {
    await page.goto("/worker-test-harness.html");
    const html = await page.evaluate(async (bytesArray: number[]) => {
      await (window as any).createDocxodusWorker();
      return (window as any).DocxodusWorker.convertDocxToHtml(new Uint8Array(bytesArray), { semanticLists: true });
    }, fixture);

    expect(await shape(page, html)).toMatchObject({ lists: 3, items: 9, itemsOutsideLists: 0 });
  });
});

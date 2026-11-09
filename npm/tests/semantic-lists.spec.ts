import { test, expect, Page } from "@playwright/test";
import * as fs from "fs";
import * as path from "path";
import { fileURLToPath } from "url";
import { storedZip, xml, W_NS, R_NS } from "./docx-zip.js";

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

// Original two-level Word-style bullets followed by a checkbox control.
function familiarBullets(): number[] {
  const level = (depth: number, glyph: string, family: string) => `<w:lvl w:ilvl="${depth}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${glyph}"/><w:pPr><w:ind w:left="${720 * (depth + 1)}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="${family}" w:hAnsi="${family}"/></w:rPr></w:lvl>`;
  const item = (id: number, depth: number, text: string) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="${depth}"/><w:numId w:val="${id}"/></w:numPr></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const parts: Record<string, string> = {
    'document': `<w:document xmlns:w="${W_NS}"><w:body>${item(1, 0, 'Harbor')}${item(1, 1, 'Ledger')}${item(2, 0, 'Confirm')}<w:sectPr/></w:body></w:document>`,
    'styles': `<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
    'settings': `<w:settings xmlns:w="${W_NS}"/>`,
    'numbering': `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0">${level(0, '\uf0b7', 'Symbol')}${level(1, 'o', 'Courier New')}</w:abstractNum><w:abstractNum w:abstractNumId="1">${level(0, '☐', 'Liberation Serif')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`,
  };
  return Array.from(storedZip(Object.entries({
    ...Object.fromEntries(Object.entries(parts).map(([name, content]) => [`word/${name}.xml`, content])),
    'word/_rels/document.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${['styles', 'settings', 'numbering'].map(name => `<Relationship Id="${name}" Type="${R_NS}/${name}" Target="${name}.xml"/>`).join('')}</Relationships>`,
    '_rels/.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="document" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>${Object.keys(parts).map(name => `<Override PartName="/word/${name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${name === 'document' ? 'document.main' : name}+xml"/>`).join('')}</Types>`,
  }).map(([name, value]) => ({ name, data: xml(value) }))));
}

for (const worker of [false, true]) {
  test(`familiar bullets normalize only on request through ${worker ? 'worker' : 'core'}`, async ({ page }) => {
    await page.goto('/worker-test-harness.html');
    const results = await page.evaluate(async ({ input, worker }) => {
      const moduleUrl = 'http://localhost:8083/core.js';
      const api = worker ? await (window as any).createDocxodusWorker() : await import(moduleUrl);
      if (!worker) await api.initialize(`${location.origin}/`);
      const results = [];
      for (const normalizeBulletMarkers of [false, true]) {
        const html = await api.convertDocxToHtml(new Uint8Array(input), {
          semanticLists: true, normalizeBulletMarkers, fabricateCssClasses: false, stampAnchors: true,
        });
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const items = Array.from(doc.querySelectorAll('li'));
        results.push({
          types: Array.from(doc.querySelectorAll('ul')).map(e => e.style.listStyleType),
          markers: items.map(e => !!e.querySelector(':scope > [data-list-marker]')),
          anchors: items.every(e => !!e.dataset.anchor),
          text: items.map(e => Array.from(e.childNodes).filter(n => !(n instanceof Element && n.matches('ul')))
            .map(n => n.textContent).join('').trim()),
        });
      }
      if (worker) api.terminate();
      return results;
    }, { input: familiarBullets(), worker });
    expect(results[0].types).toEqual(['none', 'none', 'none']);
    expect(results[0].markers).toEqual([true, true, true]);
    expect(results[1].types).toEqual(['disc', 'circle', 'none']);
    expect(results[1].markers).toEqual([false, false, true]);
    expect(results.every(r => r.anchors)).toBe(true);
    expect(results[1].text.slice(0, 2)).toEqual(['Harbor', 'Ledger']);
    expect(results[1].text[2]).toContain('☐');
    expect(results[1].text[2]).toContain('Confirm');
  });
}

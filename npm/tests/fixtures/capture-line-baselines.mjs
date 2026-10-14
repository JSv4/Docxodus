// Records Word's line baselines for line-baselines.docx (issue #908).
//
// Usage, from npm/:
//   node tests/fixtures/capture-line-baselines.mjs <word-export.pdf> "<Word version>" "<how it was exported>"
//
// Export the PDF from Word itself (File > Save As / Export > PDF, not a print driver), from the committed
// fixture, unedited. Only the numbers this prints are committed (line-baselines.word.json), never the PDF,
// following tests/visual-parity/WORD_REFERENCE.md. Each page holds one CASE paragraph; a baseline is the
// y of a text item's origin (its text matrix), in pt down from the top of the page.
import { createHash } from 'node:crypto';
import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

const here = dirname(fileURLToPath(import.meta.url));
const [pdfPath, wordVersion, exportPath] = process.argv.slice(2);
if (!pdfPath || !wordVersion || !exportPath) {
  console.error('usage: capture-line-baselines.mjs <word.pdf> "<Word version>" "<export path>"');
  process.exit(2);
}

/** Baselines per page: one entry per distinct text-origin y, top to bottom, with the text on it. */
async function pageBaselines(bytes) {
  const document = await getDocument({ data: new Uint8Array(bytes), verbosity: 0 }).promise;
  const pages = [];
  for (let number = 1; number <= document.numPages; number++) {
    const page = await document.getPage(number);
    const height = page.view[3];
    const lines = new Map();
    for (const item of (await page.getTextContent()).items) {
      if (!item.str || !item.str.trim()) continue;
      const y = Math.round((height - item.transform[5]) * 1000) / 1000;
      lines.set(y, (lines.get(y) ?? '') + item.str);
    }
    pages.push([...lines.entries()].sort((a, b) => a[0] - b[0]).map(([y, text]) => ({ y, text })));
  }
  return pages;
}

const fixture = readFileSync(join(here, 'line-baselines.docx'));
const pages = await pageBaselines(readFileSync(pdfPath));
const cases = pages.map((lines, index) => {
  const id = lines.map((line) => /CASE(\d)/.exec(line.text)?.[1]).find((match) => match !== undefined);
  return {
    case: Number(id),
    page: index + 1,
    // The y of each line's text origin; on a line holding a raised run, the run sits on its own origin.
    baselinesPt: lines.map((line) => line.y),
    lineStarts: lines.map((line) => line.text.slice(0, 12)),
  };
});
const record = {
  fixture: 'npm/tests/fixtures/line-baselines.docx',
  fixtureSha256: createHash('sha256').update(fixture).digest('hex'),
  word: wordVersion,
  export: exportPath,
  capturedAt: new Date().toISOString().slice(0, 10),
  topMarginPt: 72,
  cases,
};
writeFileSync(join(here, 'line-baselines.word.json'), `${JSON.stringify(record, null, 2)}\n`);
console.log(JSON.stringify(record, null, 2));

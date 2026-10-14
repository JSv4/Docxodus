import assert from "node:assert/strict";
import { test } from "node:test";
import { pdfLineText, joinHyphenatedLines } from "./pdf-line-text.mjs";

test("joins a hyphenated token only at a reported PDF line boundary", () => {
  assert.equal(joinHyphenatedLines(pdfLineText([
    { str: "footnote-", hasEOL: true }, { str: "1-1-599", hasEOL: false },
  ])).trim(), "footnote-1-1-599");
  assert.equal(joinHyphenatedLines(pdfLineText([
    { str: "footnote-", hasEOL: false }, { str: "1-1-599", hasEOL: false },
  ])).trim(), "footnote- 1-1-599");
});

test("retains missing, duplicated, and reordered text as inventory failures", () => {
  const token = /\bfootnote-1-1-599\b/g;
  assert.equal([...joinHyphenatedLines("footnote-\n1-1-59").matchAll(token)].length, 0);
  assert.equal([...joinHyphenatedLines("footnote-\n1-1-599 footnote-1-1-599").matchAll(token)].length, 2);
  assert.equal(joinHyphenatedLines("next\nprevious"), "next\nprevious");
});

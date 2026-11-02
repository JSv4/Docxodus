import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { renderDocxArtifacts, DocxodusExportError } from "../dist/index.js";
import { normalizeMarkupVisibility } from "docxodus/export-browser";

test("visibility defaults and object key order have one canonical representation", () => {
  for (const value of [undefined, {}, { formatting: true, comments: true, insertionsAndDeletions: true }])
    assert.deepEqual(normalizeMarkupVisibility(value), {});
  const normalized = normalizeMarkupVisibility({ comments: false, formatting: false, insertionsAndDeletions: true });
  assert.equal(JSON.stringify(normalized), '{"formatting":false,"comments":false}');
  assert.equal(Object.isFrozen(normalized), true);
});

for (const markupVisibility of [null, true, [], "hidden", { unknown: false }, { formatting: 0 }, { comments: "false" }]) {
  test(`invalid visibility ${JSON.stringify(markupVisibility)} is refused before browser startup`, async () => {
    await assert.rejects(renderDocxArtifacts(new Uint8Array([1]), {
      outputs: ["html"], reviewProfile: "markup", commentProfile: "hidden", markupVisibility,
    }), error => error instanceof DocxodusExportError && error.code === "invalid_argument"
      && error.phase === "input_validation" && /markupVisibility/.test(error.message));
  });
}

test("CLI advertises all three boolean visibility flags", () => {
  const cli = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
  const result = spawnSync(process.execPath, [cli, "--help"], { encoding: "utf8" });
  for (const flag of ["--hide-formatting-changes", "--hide-insertions-and-deletions", "--hide-comments"])
    assert.ok(result.stderr.includes(flag), flag);
});

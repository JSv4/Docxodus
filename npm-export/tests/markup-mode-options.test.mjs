import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { convertDocxToPdf, DocxodusExportError } from "../dist/index.js";

const repositoryRoot = fileURLToPath(new URL("../../", import.meta.url));
const cliEntry = fileURLToPath(new URL("../dist/cli.js", import.meta.url));
const options = { reviewProfile: "final", commentProfile: "margin", timeoutMs: 60_000 };

// Both are refused before a browser is launched.
for (const [name, extra, message] of [
  ["an unknown value", { commentColumnPlacement: "balloons", markupMode: "balloons" }, /markupMode is invalid/],
  ["inline without the balloon layout", { markupMode: "inline" }, /needs the balloon layout/],
]) {
  test(`markupMode with ${name} is an invalid argument`, async () => {
    const source = new Uint8Array(await readFile(join(repositoryRoot, "TestFiles", "CA", "CA001-Plain.docx")));
    await assert.rejects(convertDocxToPdf(source, { ...options, ...extra }),
      (error) => error instanceof DocxodusExportError && error.code === "invalid_argument" && error.phase === "input_validation"
        && message.test(error.message));
  });
}

test("the CLI documents --markup-mode and refuses a value it does not know", () => {
  const run = (args) => spawnSync(process.execPath, [cliEntry, ...args], { encoding: "utf8" });
  assert.match(run(["--help"]).stderr, /--markup-mode <mixed\|inline>/);
  const invalid = run(["convert", "in.docx", "--to", "pdf", "--output", "out.pdf", "--review-profile", "final",
    "--comments", "margin", "--comment-column-placement", "balloons", "--markup-mode", "balloons"]);
  assert.notEqual(invalid.status, 0);
  assert.match(invalid.stderr, /--markup-mode/);
  assert.doesNotMatch(invalid.stderr, /Unknown option/);
});

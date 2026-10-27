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
  ["an unknown value", { commentColumnPlacement: "balloons", balloonOverflow: "pages" }, /balloonOverflow is invalid/],
  ["the appendix without the balloon layout", { balloonOverflow: "appendix" }, /needs the balloon layout/],
]) {
  test(`balloonOverflow with ${name} is an invalid argument`, async () => {
    const source = new Uint8Array(await readFile(join(repositoryRoot, "TestFiles", "CA", "CA001-Plain.docx")));
    await assert.rejects(convertDocxToPdf(source, { ...options, ...extra }),
      (error) => error instanceof DocxodusExportError && error.code === "invalid_argument" && error.phase === "input_validation"
        && message.test(error.message));
  });
}

test("the CLI documents --balloon-overflow and refuses a value it does not know", () => {
  const run = (args) => spawnSync(process.execPath, [cliEntry, ...args], { encoding: "utf8" });
  assert.match(run(["--help"]).stderr, /--balloon-overflow <truncate\|appendix>/);
  const invalid = run(["convert", "in.docx", "--to", "pdf", "--output", "out.pdf", "--review-profile", "final",
    "--comments", "margin", "--comment-column-placement", "balloons", "--balloon-overflow", "pages"]);
  assert.notEqual(invalid.status, 0);
  assert.match(invalid.stderr, /--balloon-overflow/);
  assert.doesNotMatch(invalid.stderr, /Unknown option/);
});

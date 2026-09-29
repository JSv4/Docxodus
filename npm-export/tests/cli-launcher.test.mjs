import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { mkdir, mkdtemp, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, relative } from "node:path";
import { test } from "node:test";
import { fileURLToPath, pathToFileURL } from "node:url";

const cliEntry = fileURLToPath(new URL("../dist/cli.js", import.meta.url));

function run(args, cwd) {
  const result = spawnSync(process.execPath, args, {
    cwd,
    encoding: "utf8",
    timeout: 15_000,
  });
  assert.ifError(result.error);
  assert.equal(result.signal, null);
  return result;
}

test("the direct CLI entry prints help", () => {
  const result = run([cliEntry, "--help"]);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stderr, /^Usage:/);
});

test("the npm executable symlink runs commands and preserves their exit codes", {
  skip: process.platform === "win32" ? "npm uses command shims on Windows" : false,
}, async (t) => {
  const consumer = await mkdtemp(join(tmpdir(), "docxodus-cli-consumer-"));
  t.after(() => rm(consumer, { recursive: true, force: true }));
  const executable = join(consumer, "node_modules", ".bin", "docxodus");
  await mkdir(dirname(executable), { recursive: true });
  await symlink(relative(dirname(executable), cliEntry), executable);

  const help = run([executable, "--help"], consumer);
  assert.equal(help.status, 0, help.stderr);
  assert.match(help.stderr, /^Usage:/);

  const invalid = run([executable, "--unknown-option"], consumer);
  assert.equal(invalid.status, 2);
  assert.match(invalid.stderr, /Unknown option/);
});

test("importing runCli from another script does not execute the CLI", async (t) => {
  const consumer = await mkdtemp(join(tmpdir(), "docxodus-cli-import-"));
  t.after(() => rm(consumer, { recursive: true, force: true }));
  const importer = join(consumer, "importer.mjs");
  await writeFile(importer,
    `import { runCli } from ${JSON.stringify(pathToFileURL(cliEntry).href)};\n`
    + "console.log(typeof runCli);\n");
  const result = run([importer, "--help"], consumer);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.stdout, "function\n");
  assert.equal(result.stderr, "");
});

test("importing runCli tolerates an argv entry that is not a file", () => {
  const script = `const { runCli } = await import(${JSON.stringify(pathToFileURL(cliEntry).href)});`
    + "console.log(typeof runCli);";
  const result = run(["--input-type=module", "--eval", script, "--", "--help"]);
  assert.equal(result.status, 0, result.stderr);
  assert.equal(result.stdout, "function\n");
  assert.equal(result.stderr, "");
});

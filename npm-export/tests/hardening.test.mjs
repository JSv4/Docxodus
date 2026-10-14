import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import {
  mkdtemp,
  readFile,
  readdir,
  rm,
  stat,
  writeFile,
} from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { after, before, describe, test } from "node:test";
import { fileURLToPath } from "node:url";
import { PDFDocument } from "pdf-lib";
import { canonicalJson, canonicalJsonBytes } from "../dist/canonical.js";
import {
  convertDocxToPdf,
  DocxodusExportError,
} from "../dist/index.js";
import { chromiumSandboxUnavailable } from "../dist/browser-session.js";
import { humanDiagnostic } from "../dist/diagnostics.js";
import {
  prepareDestinations,
  publishNoReplace,
  readStableInputFile,
} from "../dist/files.js";
import { verifyPdf } from "../dist/pdf.js";
import { decodeStrictUtf8, strictJsonParse } from "../dist/strict-json.js";

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = dirname(here);
const repositoryRoot = dirname(packageRoot);
const hostEntry = join(packageRoot, "dist", "host.js");
const cliEntry = join(packageRoot, "dist", "cli.js");
const fixture = join(repositoryRoot, "TestFiles", "CA", "CA001-Plain.docx");
const baseOptions = Object.freeze({ reviewProfile: "final", commentProfile: "hidden" });
let scratch;

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function frame(bytes) {
  const payload = typeof bytes === "string" ? Buffer.from(bytes) : Buffer.from(bytes);
  const header = Buffer.alloc(4);
  header.writeUInt32BE(payload.byteLength);
  return Buffer.concat([header, payload]);
}

function requestFrame(control, blobs = []) {
  return Buffer.concat([frame(JSON.stringify(control)), ...blobs.map(frame)]);
}

function parseControlFrame(bytes) {
  assert.ok(bytes.byteLength >= 4);
  const length = bytes.readUInt32BE(0);
  assert.equal(bytes.byteLength, length + 4);
  return JSON.parse(bytes.subarray(4).toString("utf8"));
}

before(async () => {
  scratch = await mkdtemp(join(tmpdir(), "docxodus-hardening-test-"));
});

after(async () => {
  await rm(scratch, { recursive: true, force: true });
});

describe("hardening boundaries", { concurrency: false }, () => {
  test("canonical JSON is exact UTF-8 without newline and rejects unsafe values", () => {
    const value = { z: -0, a: "é", omitted: undefined, nested: { b: 2, a: 1 } };
    assert.equal(canonicalJson(value), '{"a":"é","nested":{"a":1,"b":2},"z":0}');
    const bytes = canonicalJsonBytes(value);
    assert.equal(bytes.toString("utf8"), canonicalJson(value));
    assert.equal(bytes.at(-1), "}".charCodeAt(0));
    assert.throws(() => canonicalJson({ broken: "\ud800" }), /surrogate/i);
    assert.throws(() => canonicalJson(new Date()), /plain objects/i);
    assert.throws(() => canonicalJson({ value: Number.NaN }), /non-finite/i);
    assert.equal(
      canonicalJson(JSON.parse('{"__proto__":{"polluted":true},"safe":1}')),
      '{"__proto__":{"polluted":true},"safe":1}',
    );
  });

  test("strict JSON rejects duplicate properties, excessive depth, and malformed UTF-8", () => {
    assert.throws(() => strictJsonParse('{"a":1,"a":2}'), /duplicate/i);
    assert.throws(() => strictJsonParse(`${"[".repeat(130)}0${"]".repeat(130)}`), /deep/i);
    assert.throws(() => strictJsonParse('{"value":"\\ud800"}'), /surrogate/i);
    assert.throws(() => strictJsonParse('{"value":1e999}'), /non-finite/i);
    assert.throws(() => decodeStrictUtf8(Uint8Array.of(0xc3, 0x28), "fixture"), /strict UTF-8/i);
  });

  test("source reads and destination preflight enforce limits and portable collisions", async () => {
    const inputPath = join(scratch, "source.docx");
    await writeFile(inputPath, Buffer.from("docx"));
    await assert.rejects(
      readStableInputFile(inputPath, 3),
      (error) => error instanceof DocxodusExportError
        && error.code === "resource_limit"
        && error.phase === "package_preflight",
    );
    const input = await readStableInputFile(inputPath, 4);
    await assert.rejects(
      prepareDestinations(input, {
        htmlPath: join(scratch, "Report.json"),
        reportPath: join(scratch, "report.JSON"),
      }),
      (error) => error instanceof DocxodusExportError
        && /same path/i.test(error.message),
    );
  });

  test("multi-output publication rolls back every still-owned commit", async () => {
    const firstPath = join(scratch, "transaction-first.html");
    const secondPath = join(scratch, "transaction-second.pdf");
    await writeFile(secondPath, "preexisting");
    const destination = (kind, path) => ({
      kind,
      requestedPath: path,
      absolutePath: path,
      resolvedPath: path,
      parentPath: scratch,
    });
    await assert.rejects(
      publishNoReplace([
        { destination: destination("htmlPath", firstPath), bytes: Buffer.from("first") },
        { destination: destination("pdfPath", secondPath), bytes: Buffer.from("second") },
      ]),
      (error) => error instanceof DocxodusExportError
        && error.phase === "filesystem_commit"
        && error.committedDestinations.length === 0,
    );
    await assert.rejects(stat(firstPath), { code: "ENOENT" });
    assert.equal(await readFile(secondPath, "utf8"), "preexisting");
    assert.deepEqual(
      (await readdir(scratch)).filter((name) => name.startsWith(".docxodus-")),
      [],
    );
  });

  test("pre-cancelled publication creates neither destinations nor staging files", async () => {
    const output = join(scratch, "cancelled-publication.html");
    const controller = new AbortController();
    controller.abort(new Error("test cancellation"));
    await assert.rejects(
      publishNoReplace([{
        destination: {
          kind: "htmlPath",
          requestedPath: output,
          absolutePath: output,
          resolvedPath: output,
          parentPath: scratch,
        },
        bytes: Buffer.from("must not be published"),
      }], controller.signal),
      (error) => error instanceof DocxodusExportError
        && error.code === "operation_cancelled"
        && error.phase === "output_write",
    );
    await assert.rejects(stat(output), { code: "ENOENT" });
    assert.deepEqual(
      (await readdir(scratch)).filter((name) => name.startsWith(".docxodus-")),
      [],
    );
  });

  test("byte APIs reject cancellation and invalid runtime attestations before browser work", async () => {
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(
      convertDocxToPdf(Uint8Array.of(1), { ...baseOptions, signal: controller.signal }),
      (error) => error instanceof DocxodusExportError
        && error.code === "operation_cancelled"
        && error.phase === "input_validation",
    );
    await assert.rejects(
      convertDocxToPdf(Uint8Array.of(1), {
        ...baseOptions,
        environmentAttestation: {
          chromiumProduct: "Chromium",
          chromiumBuild: "1",
          launchFlags: [],
          hostFonts: [],
          basis: "missing discriminators",
        },
      }),
      (error) => error instanceof DocxodusExportError && error.code === "invalid_argument",
    );
    await assert.rejects(
      convertDocxToPdf(Uint8Array.of(1), {
        ...baseOptions,
        browserExecutablePath: "relative/chromium",
      }),
      (error) => error instanceof DocxodusExportError && error.code === "invalid_argument",
    );
  });

  test("PDF verifier rejects parser over-admission and trailing polyglot bytes", async () => {
    const document = await PDFDocument.create();
    document.addPage([612, 792]);
    const bytes = new Uint8Array(await document.save({ useObjectStreams: false }));
    await assert.rejects(
      verifyPdf(bytes, [{ pageNumber: 1, width: 612, height: 792 }], bytes.byteLength - 1),
      (error) => error instanceof DocxodusExportError && error.code === "resource_limit",
    );
    const polyglot = Buffer.concat([bytes, Buffer.from("evil")]);
    await assert.rejects(
      verifyPdf(polyglot, [{ pageNumber: 1, width: 612, height: 792 }], polyglot.byteLength),
      (error) => error instanceof DocxodusExportError
        && error.code === "output_verification_failure"
        && /terminal|trailing/i.test(error.message),
    );
  });

  test("framed host rejects duplicate JSON, oversized control, digest mismatch, and trailing bytes", () => {
    const duplicate = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      input: frame('{"schemaVersion":1,"sources":[],"batches":[],"schemaVersion":1}'),
    });
    assert.equal(duplicate.status, 0, duplicate.stderr.toString());
    assert.match(parseControlFrame(duplicate.stdout).fatal.message, /duplicate/i);

    const oversizedHeader = Buffer.alloc(4);
    oversizedHeader.writeUInt32BE(8_388_609);
    const oversized = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      input: oversizedHeader,
    });
    assert.equal(oversized.status, 0, oversized.stderr.toString());
    assert.match(parseControlFrame(oversized.stdout).fatal.message, /frame length/i);

    const source = Buffer.from("not a docx");
    const digestMismatch = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      input: requestFrame({
        schemaVersion: 1,
        sources: [{
          id: "s1",
          byteLength: source.byteLength,
          sha256: "0".repeat(64),
          mediaType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        }],
        batches: [{
          id: "b1",
          sourceId: "s1",
          artifactRequestIds: [],
          options: { ...baseOptions, outputs: [] },
        }],
      }, [source]),
    });
    assert.equal(digestMismatch.status, 0, digestMismatch.stderr.toString());
    assert.match(parseControlFrame(digestMismatch.stdout).fatal.message, /digest mismatch/i);

    const trailing = spawnSync(process.execPath, [hostEntry], {
      cwd: packageRoot,
      input: Buffer.concat([
        requestFrame({ schemaVersion: 1, sources: [], batches: [] }),
        Buffer.from([1]),
      ]),
    });
    assert.equal(trailing.status, 0, trailing.stderr.toString());
    assert.match(parseControlFrame(trailing.stdout).fatal.message, /trailing bytes/i);
  });

  test("CLI bounds and strictly parses configuration before launching Chromium", async () => {
    const duplicateConfig = join(scratch, "duplicate-attestation.json");
    await writeFile(duplicateConfig, '{"schemaVersion":1,"schemaVersion":1}');
    const output = join(scratch, "must-not-exist.pdf");
    const args = [
      cliEntry,
      "convert",
      fixture,
      "--to", "pdf",
      "--output", output,
      "--review-profile", "final",
      "--comments", "hidden",
      "--environment-attestation", duplicateConfig,
    ];
    const duplicate = spawnSync(process.execPath, args, { cwd: packageRoot });
    assert.equal(duplicate.status, 2);
    assert.match(duplicate.stderr.toString(), /duplicate/i);
    await assert.rejects(stat(output), { code: "ENOENT" });

    const oversizedConfig = join(scratch, "oversized-attestation.json");
    await writeFile(oversizedConfig, Buffer.alloc(1_048_577, 0x20));
    const oversized = spawnSync(process.execPath, [
      ...args.slice(0, -1), oversizedConfig,
    ], { cwd: packageRoot });
    assert.equal(oversized.status, 2);
    assert.match(oversized.stderr.toString(), /exceeds 1048576 bytes/i);

    const conflicting = spawnSync(process.execPath, [
      cliEntry,
      "convert",
      fixture,
      "--to", "pdf",
      "--output", output,
      "--review-profile", "final",
      "--comments", "hidden",
      "--browser-executable", "/runtime/a/chromium",
    ], {
      cwd: packageRoot,
      env: { ...process.env, DOCXODUS_CHROMIUM_PATH: "/runtime/b/chromium" },
    });
    assert.equal(conflicting.status, 2);
    assert.match(conflicting.stderr.toString(), /conflicts/i);
  });

  test("human diagnostics render the cause chain, terminal-safe and newline-preserving", () => {
    const launchLog = [
      "browserType.launch: Target page, context or browser has been closed",
      "Browser logs:",
      "Chromium sandboxing failed!",
      "================================",
      "  - (preferred): Configure your environment to support sandboxing",
      "================================",
    ].join("\n");
    const decorated = `\u001b[2m${launchLog}\u001b[22m\r\nCall\tlog:\u0000`;
    const rendered = humanDiagnostic(new DocxodusExportError(
      "browser_launch_failure",
      "browser_launch",
      "Chromium could not be launched because this host denies its process sandbox.",
      "Permit unprivileged user namespaces on the render host.",
      {
        cause: new AggregateError([
          new Error(decorated),
          new Error("the private temporary directory could not be removed"),
        ]),
      },
    ));
    assert.match(rendered, /^browser_launch_failure \(browser_launch\): /);
    assert.match(rendered, /^Cause: browserType\.launch: /m);
    assert.ok(rendered.includes("Chromium sandboxing failed!"));
    assert.match(rendered, /^Cause: the private temporary directory could not be removed$/m);
    assert.match(rendered, /^Call log:$/m);
    assert.equal(/[\u0000-\u0009\u000b-\u001f\u007f-\u009f]/.test(rendered), false);
  });

  test("cause rendering is bounded and terminates on cyclic chains", () => {
    const inner = new Error("inner");
    const outer = new Error("outer", { cause: inner });
    inner.cause = outer;
    const cyclic = humanDiagnostic(outer);
    assert.equal(cyclic.match(/^Cause: /gm).length, 1);
    assert.match(cyclic, /^Cause: inner$/m);

    const bounded = humanDiagnostic(new Error("top", {
      cause: new Error("x".repeat(64 * 1024)),
    }));
    assert.ok(bounded.length <= 16_384, `rendered ${bounded.length} characters`);
    assert.ok(bounded.endsWith("..."));

    const starved = humanDiagnostic(new DocxodusExportError(
      "conversion_failure",
      "conversion",
      "m",
      "r",
      { detail: "d".repeat(64 * 1024), cause: new Error("the reason this failed") },
    ));
    assert.match(starved, /^Cause: the reason this failed$/m);
  });

  test("an unavailable Chromium process sandbox is recognized behind its cause chain", () => {
    assert.equal(chromiumSandboxUnavailable(new Error("Chromium sandboxing failed!\nlogs")), true);
    assert.equal(
      chromiumSandboxUnavailable(new Error(
        "[err] No usable sandbox! If you are running on Ubuntu 23.10+ or another Linux distro",
      )),
      true,
    );
    assert.equal(
      chromiumSandboxUnavailable(new AggregateError([
        new Error("[err] see https://crbug.com/638180 for more information"),
        new Error("cleanup failed"),
      ])),
      true,
    );
    assert.equal(chromiumSandboxUnavailable(new Error("ENOENT: no such file or directory")), false);
    assert.equal(chromiumSandboxUnavailable(undefined), false);
  });

  test("CLI stderr carries the reason a Chromium launch failed", async () => {
    const output = join(scratch, "launch-cause-must-not-exist.pdf");
    const missing = join(scratch, "chromium-that-does-not-exist");
    const failed = spawnSync(process.execPath, [
      cliEntry,
      "convert",
      fixture,
      "--to", "pdf",
      "--output", output,
      "--review-profile", "final",
      "--comments", "hidden",
      "--browser-executable", missing,
    ], { cwd: packageRoot });
    const stderr = failed.stderr.toString();
    assert.equal(failed.status, 1, stderr);
    assert.match(stderr, /^browser_launch_failure \(browser_launch\): /m);
    assert.match(stderr, /^Cause: [\s\S]*ENOENT/m);
    await assert.rejects(stat(output), { code: "ENOENT" });
  });
});

import assert from "node:assert/strict";
import { copyFile, mkdir, mkdtemp, readFile, rm, symlink, truncate, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import { DEFAULT_EXPORT_RESOURCE_LIMITS, renderDocxArtifacts } from "../dist/index.js";
import { createNodeFontRuntime } from "../dist/fonts/index.js";
import { localFontDirectories, LocalFontIndex } from "../dist/fonts/local.js";
import { inspectCommonFonts } from "../dist/fonts/doctor.js";
import fs from "node:fs/promises";
import { constants } from "node:fs";
import { execFileSync } from "node:child_process";
import { syncBuiltinESMExports } from "node:module";
import { readStableFontFile } from "../dist/fonts/discovery.js";

const fixtures = new URL("./fixtures/", import.meta.url);
const policy = await readFile(new URL("docxodus-policy-base.ttf", fixtures));
const metric = await readFile(new URL("docxodus-metric-test.ttf", fixtures));

test("rejects malformed local-source options before starting a browser", async () => {
  for (const fontSources of [null, true, "local", ["remote"], ["local", "local"]]) {
    await assert.rejects(renderDocxArtifacts(new Uint8Array([1]), {
      outputs: ["html"], reviewProfile: "final", commentProfile: "hidden", fontSources,
    }), error => error.code === "invalid_argument" && /fontSources/.test(error.message));
  }
});

async function folder(t) {
  const root = await mkdtemp(join(tmpdir(), "docxodus-local-fonts-"));
  t.after(() => rm(root, { recursive: true, force: true }));
  return root;
}

function runtime(directories, options = {}) {
  const warnings = [];
  const fontRuntime = createNodeFontRuntime(options.configured ?? [], [],
    { ...DEFAULT_EXPORT_RESOURCE_LIMITS, ...options.limits }, options.outputs ?? ["html", "pdf"],
    { directories, strict: options.strict ?? false, warn: warning => warnings.push(warning) });
  return { warnings, resolve: async (...families) => fontRuntime.resolver({
    schemaVersion: 1, requests: families.map((family, index) => ({
      id: `r${index}`, familyStack: [family, "serif"], familyKinds: ["named", "generic"],
      style: "normal", weight: 400, stretch: 100, sampleCodePoints: [32, 65, 90],
    })),
  }, new AbortController().signal) };
}

function fsType(bytes, value) {
  const changed = Buffer.from(bytes);
  for (let index = 0; index < changed.readUInt16BE(4); index++) {
    const record = 12 + index * 16;
    if (changed.toString("ascii", record, record + 4) === "OS/2") changed.writeUInt16BE(value, changed.readUInt32BE(record + 8) + 8);
  }
  return changed;
}

test("loads only requested local files from a large index", async t => {
  const root = await folder(t);
  await writeFile(join(root, "requested.ttf"), metric);
  await Promise.all(Array.from({ length: 1100 }, (_, index) => writeFile(join(root, `unused-${index}.ttf`), policy)));
  // Its name is readable, but full loading would exceed the per-file resource limit.
  await truncate(join(root, "unused-0.ttf"), DEFAULT_EXPORT_RESOURCE_LIMITS.fontFileBytes + 1);
  const { resolve, warnings } = runtime([root], { limits: { fontFiles: 1 } });
  const result = await resolve("Docxodus Metric Test");
  assert.equal(result.outcomes[0].status, "resolved");
  assert.equal(result.faces.length, 1);
  assert.equal(result.faces[0].source, "local");
  assert.deepEqual(warnings, []);
  assert.equal(JSON.stringify(result).includes(root), false);
});

test("explicit faces win over local candidates and local matching is exact", async t => {
  const root = await folder(t);
  const configured = join(root, "configured");
  const local = join(root, "local");
  await mkdir(configured); await mkdir(local);
  await writeFile(join(configured, "font.ttf"), policy);
  await writeFile(join(local, "font.ttf"), fsType(policy, 2));
  const { resolve, warnings } = runtime([local], { configured: [configured] });
  const result = await resolve("Docxodus Policy Test", "Docxodus Policy");
  assert.equal(result.outcomes[0].status, "resolved");
  assert.equal(result.faces[0].source, undefined);
  assert.equal(result.outcomes[1].status, "missing");
  assert.deepEqual(warnings, []);
});

test("repeated over-budget requests do not retain rejected local faces", async t => {
  const root = await folder(t);
  for (const name of ["docxodus-policy-base.ttf", "docxodus-metric-test.ttf", "synthetic-calibri.ttf", "synthetic-carlito.ttf"]) {
    await copyFile(new URL(name, fixtures), join(root, name));
  }
  const { resolve } = runtime([root], { limits: { fontFiles: 1 } });
  const first = await resolve("Docxodus Policy Test");
  for (const family of ["Docxodus Metric Test", "Calibri", "Carlito", "Docxodus Metric Test"]) {
    await assert.rejects(resolve(family), error => error.code === "resource_limit"
      && error.message === "fontFiles limit exceeded (2 > 1).");
  }
  assert.deepEqual(await resolve("Docxodus Policy Test"), first);
});

test("skips forbidden local faces with a path-free warning, and strict mode fails", async t => {
  const root = await folder(t);
  await writeFile(join(root, "private-font.ttf"), fsType(policy, 2));
  const { resolve, warnings } = runtime([root]);
  const result = await resolve("Docxodus Policy Test");
  assert.equal(result.outcomes[0].status, "missing");
  assert.match(warnings[0].message, /Docxodus Policy Test.*forbids font embedding/);
  assert.equal(JSON.stringify(warnings).includes(root), false);
  assert.equal(JSON.stringify(warnings).includes("private-font"), false);
  await assert.rejects(runtime([root], { strict: true }).resolve("Docxodus Policy Test"), /Docxodus Policy Test.*forbids font embedding/);
});

test("deduplicates identical local faces but rejects disagreeing duplicates", async t => {
  const root = await folder(t);
  await writeFile(join(root, "one.ttf"), policy);
  await copyFile(join(root, "one.ttf"), join(root, "two.ttf"));
  assert.equal((await runtime([root], { limits: { fontFiles: 1 } }).resolve("Docxodus Policy Test")).outcomes[0].status, "resolved");
  const changed = Buffer.concat([policy, Buffer.from([0, 0, 0, 0])]);
  await writeFile(join(root, "two.ttf"), changed);
  const { resolve, warnings } = runtime([root]);
  assert.equal((await resolve("Docxodus Policy Test")).outcomes[0].status, "missing");
  assert.match(warnings[0].message, /Docxodus Policy Test.*disagree/);
});

test("accepts bounded local aliases and skips escaping links", async t => {
  const root = await folder(t);
  const fonts = join(root, "fonts"); await mkdir(fonts);
  await writeFile(join(fonts, "face.ttf"), policy);
  await symlink(join(fonts, "face.ttf"), join(fonts, "alias.ttf"));
  assert.equal((await runtime([fonts]).resolve("Docxodus Policy Test")).outcomes[0].status, "resolved");
  await writeFile(join(root, "outside.ttf"), metric);
  await symlink(join(root, "outside.ttf"), join(fonts, "escape.ttf"));
  const { resolve, warnings } = runtime([fonts]);
  assert.equal((await resolve("Docxodus Metric Test")).outcomes[0].status, "missing");
  assert.match(warnings[0].message, /Docxodus Metric Test.*unsafe symlink/);
});

test("loads a single requested face from a local collection", async t => {
  const root = await folder(t);
  await copyFile(new URL("docxodus-test-collection.ttc", fixtures), join(root, "faces.ttc"));
  const result = await runtime([root], { limits: { fontFiles: 1 } }).resolve("Docxodus Metric Test");
  assert.equal(result.outcomes[0].status, "resolved");
  assert.equal(result.faces[0].resolvedFamily, "Docxodus Metric Test");
  assert.equal(result.faces[0].source, "local");
});

test("searches user, system and Office font locations on each supported platform", () => {
  assert.deepEqual(localFontDirectories("linux", "/user", {}), [
    "/user/.local/share/fonts", "/user/.fonts", "/usr/local/share/fonts", "/usr/share/fonts",
  ]);
  assert.deepEqual(localFontDirectories("win32", "C:\\User", { WINDIR: "C:\\Windows", LOCALAPPDATA: "C:\\User\\Local" }), [
    "C:\\Windows\\Fonts", "C:\\User\\Local\\Microsoft\\Windows\\Fonts", "C:\\User\\Local\\Microsoft\\FontCache\\4\\CloudFonts",
  ]);
  const mac = localFontDirectories("darwin", "/user", {});
  for (const expected of ["/user/Library/Fonts", "/Library/Fonts", "/System/Library/Fonts",
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts", "/user/Applications/Microsoft Word.app/Contents/Resources/Fonts"]) {
    assert.ok(mac.includes(expected), expected);
  }
});

test("doctor reports exact common families, known substitutions and missing faces", async t => {
  const root = await folder(t);
  await copyFile(new URL("synthetic-carlito.ttf", fixtures), join(root, "carlito.ttf"));
  const results = await inspectCommonFonts([], { directories: [root] });
  assert.equal(results.find(face => face.requestedFamily === "Calibri").status, "substituted");
  assert.equal(results.find(face => face.requestedFamily === "Calibri").resolvedFamily, "Carlito");
  assert.equal(results.find(face => face.requestedFamily === "Aptos").status, "missing");
  await copyFile(new URL("synthetic-calibri.ttf", fixtures), join(root, "calibri.ttf"));
  const exact = await inspectCommonFonts([], { directories: [root] });
  assert.equal(exact.find(face => face.requestedFamily === "Calibri").status, "resolved");
});

test("an exact local family precedes a configured substitution and forbidden faces fall through", async t => {
  const root = await folder(t);
  const explicit = join(root, "explicit"), local = join(root, "local");
  await mkdir(explicit); await mkdir(local);
  await copyFile(new URL("synthetic-carlito.ttf", fixtures), join(explicit, "carlito.ttf"));
  const calibri = await readFile(new URL("synthetic-calibri.ttf", fixtures));
  await writeFile(join(local, "calibri.ttf"), calibri);
  const exact = await runtime([local], { configured: [explicit] }).resolve("Calibri");
  assert.equal(exact.outcomes[0].status, "resolved");
  assert.equal(exact.faces[0].source, "local");
  await writeFile(join(local, "calibri.ttf"), fsType(calibri, 2));
  const fallback = runtime([local], { configured: [explicit] });
  const substitute = await fallback.resolve("Calibri");
  assert.equal(substitute.outcomes[0].status, "substituted");
  assert.equal(substitute.outcomes[0].resolvedFamily, "Carlito");
  assert.match(fallback.warnings[0].message, /Calibri.*forbids font embedding/);
});

test("decode retries accept only previously offered local identities and never return them again", async t => {
  const root = await folder(t), configured = join(root, "explicit"), local = join(root, "local");
  await mkdir(configured); await mkdir(local);
  await copyFile(new URL("synthetic-calibri.ttf", fixtures), join(local, "calibri.ttf"));
  await copyFile(new URL("synthetic-carlito.ttf", fixtures), join(configured, "carlito.ttf"));
  const warnings = [];
  const resolver = createNodeFontRuntime([configured], [], DEFAULT_EXPORT_RESOURCE_LIMITS, ["pdf"],
    { directories: [local], warn: warning => warnings.push(warning) }).resolver;
  const request = { schemaVersion: 1, requests: [{ id: "calibri", familyStack: ["Calibri"], familyKinds: ["named"],
    style: "normal", weight: 400, stretch: 100, sampleCodePoints: [65] }] };
  const signal = new AbortController().signal;
  const initial = await resolver(request, signal);
  const rejectedFaceIds = [initial.faces[0].id];
  const retry = await resolver({ ...request, rejectedFaceIds }, signal);
  assert.equal(retry.outcomes[0].resolvedFamily, "Carlito");
  assert.equal(retry.faces.some(face => rejectedFaceIds.includes(face.id)), false);
  assert.equal((await resolver({ ...request, rejectedFaceIds }, signal)).outcomes[0].resolvedFamily, "Carlito");
  assert.equal(warnings.length, 1);
  for (const rejected of [["unknown"], [retry.faces[0].id], [initial.faces[0].id, initial.faces[0].id], null, [42]]) {
    await assert.rejects(resolver({ ...request, rejectedFaceIds: rejected }, signal), /rejectedFaceIds|unknown local font/);
  }
});

test("a changed or removed indexed file is skipped without path disclosure", async t => {
  const root = await folder(t);
  const path = join(root, "private.ttf");
  await writeFile(path, metric);
  const warnings = [];
  const index = new LocalFontIndex({ directories: [root], warn: warning => warnings.push(warning) });
  await index.prepare(DEFAULT_EXPORT_RESOURCE_LIMITS);
  await rm(path);
  assert.equal(await index.load(index.candidates("Docxodus Metric Test")[0], DEFAULT_EXPORT_RESOURCE_LIMITS, ["pdf"]), undefined);
  assert.match(warnings[0].message, /Docxodus Metric Test.*read safely/);
  assert.equal(JSON.stringify(warnings).includes(path), false);
});

test("counts used files and rejects no-subsetting fonts only for PDF", async t => {
  const root = await folder(t);
  await writeFile(join(root, "policy.ttf"), policy);
  await writeFile(join(root, "metric.ttf"), metric);
  await assert.rejects(runtime([root], { limits: { fontFiles: 1 } }).resolve("Docxodus Policy Test", "Docxodus Metric Test"), /fontFiles/);
  await writeFile(join(root, "policy.ttf"), fsType(policy, 256));
  const html = await runtime([root], { outputs: ["html"] }).resolve("Docxodus Policy Test");
  assert.equal(html.outcomes[0].status, "resolved");
  assert.equal(html.faces[0].licenseEvidence.noSubsetting, true);
  const pdf = runtime([root]);
  assert.equal((await pdf.resolve("Docxodus Policy Test")).outcomes[0].status, "missing");
  assert.match(pdf.warnings[0].message, /forbid PDF subsetting/);
  await writeFile(join(root, "policy.ttf"), fsType(policy, 4));
  assert.equal((await runtime([root], { outputs: ["pdf"] }).resolve("Docxodus Policy Test")).outcomes[0].status, "resolved");
  const previewHtml = runtime([root], { outputs: ["html"] });
  assert.equal((await previewHtml.resolve("Docxodus Policy Test")).outcomes[0].status, "missing");
  assert.match(previewHtml.warnings[0].message, /preview-and-print rights.*HTML/);
});

for (const phase of ["metadata", "snapshot"]) {
  test(`a file replaced by a FIFO cannot block ${phase} loading or cancellation`, {
    skip: process.platform === "win32" || !constants.O_NONBLOCK,
  }, async t => {
    const root = await folder(t), path = join(root, "raced.ttf");
    await writeFile(path, policy);
    const original = fs.lstat;
    let exchanged = false;
    const patched = t.mock.method(fs, "lstat", async (candidate, ...args) => {
      const result = await original(candidate, ...args);
      if (candidate === path && !exchanged) {
        exchanged = true;
        await rm(path);
        execFileSync("mkfifo", [path]);
      }
      return result;
    });
    syncBuiltinESMExports();
    const controller = new AbortController();
    const cancel = setTimeout(() => controller.abort(), 50);
    let timeout;
    const pending = phase === "metadata"
      ? new LocalFontIndex({ directories: [root] }).prepare(DEFAULT_EXPORT_RESOURCE_LIMITS, controller.signal)
      : assert.rejects(readStableFontFile(path, root, "race fixture", DEFAULT_EXPORT_RESOURCE_LIMITS.fontFileBytes, controller.signal));
    try {
      await Promise.race([pending, new Promise((_, reject) => {
        timeout = setTimeout(() => reject(new Error("FIFO open blocked after cancellation")), 300);
      })]);
      assert.equal(exchanged, true);
    } finally {
      clearTimeout(cancel); clearTimeout(timeout);
      patched.mock.restore(); syncBuiltinESMExports();
      // Release an old implementation's blocked reader so a failing regression can exit.
      const writer = await fs.open(path, constants.O_WRONLY | constants.O_NONBLOCK).catch(() => undefined);
      await writer?.close();
      await pending.catch(() => {});
    }
  });
}

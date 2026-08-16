import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, readFile, rm, symlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, test } from "node:test";
import { fileURLToPath } from "node:url";
import { DEFAULT_EXPORT_RESOURCE_LIMITS } from "../dist/index.js";
import {
  createNodeFontResolver,
  discoverFontCatalog,
  pathFreeCatalogManifest,
  resolveCatalogRequests,
} from "../dist/fonts/index.js";

const fixture = fileURLToPath(new URL(
  "../../docs/demo/fonts/docxodus-canvas-mono.woff2",
  import.meta.url,
));
const policyFixtures = fileURLToPath(new URL("./fixtures", import.meta.url));
const temporaryDirectories = [];

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function withFsType(bytes, fsType) {
  const result = Buffer.from(bytes);
  const tableCount = result.readUInt16BE(4);
  for (let index = 0; index < tableCount; index++) {
    const record = 12 + index * 16;
    if (result.toString("ascii", record, record + 4) !== "OS/2") continue;
    result.writeUInt16BE(fsType, result.readUInt32BE(record + 8) + 8);
    return result;
  }
  throw new Error("Fixture has no OS/2 table");
}

function withWeightClass(bytes, weight) {
  const result = Buffer.from(bytes);
  const tableCount = result.readUInt16BE(4);
  for (let index = 0; index < tableCount; index++) {
    const record = 12 + index * 16;
    if (result.toString("ascii", record, record + 4) !== "OS/2") continue;
    result.writeUInt16BE(weight, result.readUInt32BE(record + 8) + 4);
    return result;
  }
  throw new Error("Fixture has no OS/2 table");
}

function attestation(fileSha256) {
  return {
    schemaVersion: 1,
    usage: "standalone-document-font-embedding",
    fileSha256,
    embeddingPermitted: true,
    permittedOutputs: ["html", "pdf"],
    subsettingPermitted: true,
    basis: "DejaVu-derived test fixture license",
    attester: "Docxodus test suite",
  };
}

async function fontDirectory(name = "configured-fonts") {
  const root = await import("node:fs/promises").then(({ mkdtemp }) =>
    mkdtemp(join(tmpdir(), "docxodus-font-runtime-")));
  temporaryDirectories.push(root);
  const directory = join(root, name);
  await mkdir(directory);
  return { root, directory };
}

function request(overrides = {}) {
  return {
    schemaVersion: 1,
    requests: [{
      id: "face-1",
      familyStack: ["Docxodus Canvas Mono", "monospace"],
      style: "normal",
      weight: 400,
      stretch: 100,
      sampleCodePoints: [32, 65, 90],
      ...overrides,
    }],
  };
}

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((path) =>
    rm(path, { recursive: true, force: true })));
});

describe("verified Node font runtime", () => {
  test("resolves an exact attested WOFF2 face from one immutable snapshot", async () => {
    const { directory } = await fontDirectory();
    const original = await readFile(fixture);
    const fileSha256 = digest(original);
    const target = join(directory, "face.woff2");
    await writeFile(target, original);

    const catalog = await discoverFontCatalog(
      [directory],
      [attestation(fileSha256)],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.equal(catalog.faces.length, 1);
    assert.equal(catalog.faces[0].sha256, fileSha256);
    assert.equal(catalog.faces[0].licenseEvidence.kind, "attested");

    // Mutating the path after discovery cannot alter the already parsed/injected snapshot.
    await writeFile(target, Buffer.alloc(original.byteLength, 0));
    const resolved = resolveCatalogRequests(catalog, request().requests);
    assert.deepEqual(resolved.outcomes, [{
      requestId: "face-1",
      requestedFamily: "Docxodus Canvas Mono",
      resolvedFamily: "Docxodus Canvas Mono",
      status: "resolved",
      faceId: `font-${fileSha256}`,
      metricCompatible: true,
      faceMatch: "exact",
      glyphCoverage: "complete",
    }]);
    assert.equal(digest(Buffer.from(resolved.faces[0].bytesBase64, "base64")), fileSha256);

    const manifestText = JSON.stringify(pathFreeCatalogManifest(catalog));
    assert.equal(manifestText.includes(directory), false);
    assert.equal(manifestText.includes("bytesBase64"), false);
  });

  test("requires exact WOFF2 embedding evidence and redacts filesystem paths", async () => {
    const { root, directory } = await fontDirectory("private-customer-fonts");
    const bytes = await readFile(fixture);
    await writeFile(join(directory, "confidential-name.woff2"), bytes);
    const catalog = await discoverFontCatalog(
      [directory],
      [],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.throws(
      () => resolveCatalogRequests(catalog, request().requests),
      (error) => {
        const serialized = JSON.stringify(error.toJSON());
        assert.equal(error.cause, undefined);
        assert.equal(serialized.includes(root), false);
        assert.equal(serialized.includes("confidential-name"), false);
        assert.match(serialized, /requires an exact embedding-rights attestation/);
        return true;
      },
    );
  });

  test("rejects WOFF attestation digest mismatch exactly", async () => {
    const { root, directory } = await fontDirectory("woff-digest-mismatch");
    const bytes = await readFile(join(policyFixtures, "docxodus-metric-test.woff"));
    await writeFile(join(directory, "face.woff"), bytes);
    const catalog = await discoverFontCatalog(
      [directory],
      [attestation("0".repeat(64))],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.throws(
      () => resolveCatalogRequests(catalog, [request({
        familyStack: ["Docxodus Metric Test"],
      }).requests[0]]),
      (error) => error.code === "resource_policy_failure"
        && /requires an exact embedding-rights attestation/.test(error.message)
        && !JSON.stringify(error.toJSON()).includes(root)
        && error.cause === undefined,
    );
  });

  test("enforces OS/2 restricted and bitmap-only embedding rights", async () => {
    const base = await readFile(join(policyFixtures, "docxodus-policy-base.ttf"));
    for (const [fsType, expected] of [
      [0x0002, /forbids font embedding/],
      [0x0200, /bitmap embedding only/],
    ]) {
      const { root, directory } = await fontDirectory(`fsType-${fsType}`);
      const restricted = withFsType(base, fsType);
      await writeFile(join(directory, "face.ttf"), restricted);
      const catalog = await discoverFontCatalog(
        [directory],
        [attestation(digest(restricted))],
        DEFAULT_EXPORT_RESOURCE_LIMITS,
      );
      assert.throws(
        () => resolveCatalogRequests(catalog, [request({
          familyStack: ["Docxodus Policy Test"],
        }).requests[0]]),
        (error) => error.code === "resource_policy_failure"
          && expected.test(error.message)
          && !JSON.stringify(error.toJSON()).includes(root)
          && error.cause === undefined,
      );
    }
  });

  test("enforces attested output scope and no-subsetting PDF policy", async () => {
    const first = await fontDirectory("html-only-attestation");
    const bytes = await readFile(fixture);
    const fileSha256 = digest(bytes);
    await writeFile(join(first.directory, "face.woff2"), bytes);
    const htmlOnly = {
      ...attestation(fileSha256),
      permittedOutputs: ["html"],
    };
    const htmlCatalog = await discoverFontCatalog(
      [first.directory],
      [htmlOnly],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.equal(resolveCatalogRequests(htmlCatalog, request().requests, ["html"]).faces.length, 1);
    assert.throws(
      () => resolveCatalogRequests(htmlCatalog, request().requests, ["pdf"]),
      (error) => error.code === "resource_policy_failure" && /requested output/.test(error.message),
    );

    const second = await fontDirectory("no-subsetting-attestation");
    await writeFile(join(second.directory, "face.woff2"), bytes);
    const noSubsetting = {
      ...attestation(fileSha256),
      subsettingPermitted: false,
    };
    const noSubsetCatalog = await discoverFontCatalog(
      [second.directory],
      [noSubsetting],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.equal(resolveCatalogRequests(noSubsetCatalog, request().requests, ["html"]).faces.length, 1);
    assert.throws(
      () => resolveCatalogRequests(noSubsetCatalog, request().requests, ["pdf"]),
      (error) => error.code === "resource_policy_failure" && /forbids subsetting/.test(error.message),
    );
    assert.notEqual(
      htmlCatalog.faces[0].licenseEvidence.identity,
      noSubsetCatalog.faces[0].licenseEvidence.identity,
    );
  });

  test("rejects OS/2 no-subsetting fonts for PDF but permits full-byte HTML embedding", async () => {
    const { directory } = await fontDirectory("os2-no-subsetting");
    const base = await readFile(join(policyFixtures, "docxodus-policy-base.ttf"));
    await writeFile(join(directory, "face.ttf"), withFsType(base, 0x0100));
    const catalog = await discoverFontCatalog([directory], [], DEFAULT_EXPORT_RESOURCE_LIMITS);
    const fontRequest = request({ familyStack: ["Docxodus Policy Test"] }).requests;
    assert.equal(resolveCatalogRequests(catalog, fontRequest, ["html"]).faces.length, 1);
    assert.throws(
      () => resolveCatalogRequests(catalog, fontRequest, ["pdf"]),
      (error) => error.code === "resource_policy_failure" && /forbids subsetting/.test(error.message),
    );
  });

  test("honors earlier-directory precedence for otherwise identical faces", async () => {
    const earlier = await fontDirectory("earlier-precedence");
    const later = await fontDirectory("later-precedence");
    const ttf = await readFile(join(policyFixtures, "docxodus-metric-test.ttf"));
    const woff = await readFile(join(policyFixtures, "docxodus-metric-test.woff"));
    await writeFile(join(earlier.directory, "face.woff"), woff);
    await writeFile(join(later.directory, "face.ttf"), ttf);
    const woffDigest = digest(woff);
    const firstCatalog = await discoverFontCatalog(
      [earlier.directory, later.directory],
      [attestation(woffDigest)],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    const faceRequest = request({ familyStack: ["Docxodus Metric Test"] }).requests;
    assert.equal(resolveCatalogRequests(firstCatalog, faceRequest).faces[0].format, "woff");

    const reversedCatalog = await discoverFontCatalog(
      [later.directory, earlier.directory],
      [attestation(woffDigest)],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    const reversed = resolveCatalogRequests(reversedCatalog, faceRequest);
    assert.equal(reversed.faces[0].format, "ttf");
    assert.equal(reversed.faces[0].licenseEvidence.kind, "installable");
  });

  test("keeps earlier-directory authority ahead of a closer face in a later root", async () => {
    const earlier = await fontDirectory("earlier-policy-root");
    const later = await fontDirectory("later-policy-root");
    const regular = await readFile(join(policyFixtures, "docxodus-metric-test.ttf"));
    const bold = withWeightClass(regular, 700);
    await writeFile(join(earlier.directory, "regular.ttf"), regular);
    await writeFile(join(later.directory, "bold.ttf"), bold);

    const catalog = await discoverFontCatalog(
      [earlier.directory, later.directory],
      [],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    const resolved = resolveCatalogRequests(catalog, request({
      familyStack: ["Docxodus Metric Test"],
      weight: 700,
    }).requests);

    assert.equal(resolved.faces[0].weight, 400);
    assert.equal(resolved.outcomes[0].faceMatch, "synthesized");
  });

  test("rejects symlinks and malformed webfont lengths without path disclosure", async () => {
    const first = await fontDirectory("symlink-case");
    await symlink(fixture, join(first.directory, "linked.woff2"));
    await assert.rejects(
      discoverFontCatalog([first.directory], [], DEFAULT_EXPORT_RESOURCE_LIMITS),
      (error) => {
        const serialized = JSON.stringify(error.toJSON());
        assert.equal(error.cause, undefined);
        assert.equal(serialized.includes(first.root), false);
        assert.equal(serialized.includes("linked.woff2"), false);
        assert.match(serialized, /symlink/);
        return true;
      },
    );

    const second = await fontDirectory("malformed-case");
    const bytes = await readFile(fixture);
    await writeFile(join(second.directory, "malformed.woff2"), Buffer.concat([bytes, Buffer.of(0)]));
    await assert.rejects(
      discoverFontCatalog([second.directory], [], DEFAULT_EXPORT_RESOURCE_LIMITS),
      (error) => error.code === "resource_policy_failure"
        && error.phase === "font_loading"
        && !JSON.stringify(error.toJSON()).includes(second.root),
    );
  });

  test("enforces declared expanded bytes before parsing and deduplicates identical files", async () => {
    const { directory } = await fontDirectory();
    const bytes = await readFile(fixture);
    await writeFile(join(directory, "a.woff2"), bytes);
    await writeFile(join(directory, "b.woff2"), bytes);

    const catalog = await discoverFontCatalog(
      [directory],
      [attestation(digest(bytes))],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    assert.equal(catalog.fileCount, 2);
    assert.equal(catalog.faces.length, 1);

    await assert.rejects(
      discoverFontCatalog([directory], [attestation(digest(bytes))], {
        ...DEFAULT_EXPORT_RESOURCE_LIMITS,
        fontFileBytes: bytes.byteLength + 1,
      }),
      (error) => error.code === "resource_limit"
        && error.phase === "font_loading"
        && /fontFileBytes/.test(error.message),
    );
  });

  test("rejects ambiguous same-face bytes within one directory", async () => {
    const { root, directory } = await fontDirectory("ambiguous-private-fonts");
    const first = Buffer.from(await readFile(fixture));
    const second = Buffer.from(first);
    second.writeUInt16BE((second.readUInt16BE(24) + 1) & 0xffff, 24);
    const firstDigest = digest(first);
    const secondDigest = digest(second);
    await writeFile(join(directory, "first.woff2"), first);
    await writeFile(join(directory, "second.woff2"), second);

    await assert.rejects(
      discoverFontCatalog(
        [directory],
        [attestation(firstDigest), attestation(secondDigest)],
        DEFAULT_EXPORT_RESOURCE_LIMITS,
      ),
      (error) => {
        const serialized = JSON.stringify(error.toJSON());
        assert.equal(error.cause, undefined);
        assert.equal(serialized.includes(root), false);
        assert.equal(serialized.includes("first.woff2"), false);
        assert.equal(serialized.includes("second.woff2"), false);
        assert.match(error.message, /ambiguous files for one family and face/);
        return true;
      },
    );

    const earlier = await fontDirectory("earlier-directory");
    const later = await fontDirectory("later-directory");
    await writeFile(join(earlier.directory, "only.woff2"), first);
    await writeFile(join(later.directory, "a-identical.woff2"), first);
    await writeFile(join(later.directory, "b-conflict.woff2"), second);
    await assert.rejects(
      discoverFontCatalog(
        [earlier.directory, later.directory],
        [attestation(firstDigest), attestation(secondDigest)],
        DEFAULT_EXPORT_RESOURCE_LIMITS,
      ),
      (error) => error.code === "resource_policy_failure"
        && /ambiguous files for one family and face/.test(error.message)
        && error.cause === undefined,
    );
  });

  test("applies the shared substitution contract and reports a missing family canonically", async () => {
    const { directory } = await fontDirectory();
    const bytes = await readFile(fixture);
    const fileSha256 = digest(bytes);
    await writeFile(join(directory, "face.woff2"), bytes);
    const catalog = await discoverFontCatalog(
      [directory],
      [attestation(fileSha256)],
      DEFAULT_EXPORT_RESOURCE_LIMITS,
    );
    const substituteCatalog = {
      ...catalog,
      faces: Object.freeze(catalog.faces.map((face) => Object.freeze({
        ...face,
        family: "Liberation Mono",
        familyKey: "liberation mono",
      }))),
    };
    const substituted = resolveCatalogRequests(substituteCatalog, [request({
      familyStack: ["Courier New"],
    }).requests[0]]);
    assert.equal(substituted.outcomes[0].status, "substituted");
    assert.equal(substituted.outcomes[0].resolvedFamily, "Liberation Mono");
    assert.equal(substituted.outcomes[0].metricCompatible, true);

    const missing = resolveCatalogRequests(catalog, [request({
      familyStack: ["Definitely Missing Font"],
    }).requests[0]]);
    assert.deepEqual(missing, {
      outcomes: [{
        requestId: "face-1",
        requestedFamily: "Definitely Missing Font",
        status: "missing",
      }],
      faces: [],
    });
  });

  test("validates the exposed resolver request and returns canonical response records", async () => {
    const { directory } = await fontDirectory();
    const bytes = await readFile(fixture);
    const fileSha256 = digest(bytes);
    await writeFile(join(directory, "face.woff2"), bytes);
    const directories = [directory];
    const license = attestation(fileSha256);
    const limits = { ...DEFAULT_EXPORT_RESOURCE_LIMITS, fontRequests: 1 };
    const resolver = createNodeFontResolver(
      directories,
      [license],
      limits,
    );
    directories[0] = join(directory, "caller-mutated-path");
    license.fileSha256 = "0".repeat(64);
    limits.fontRequests = 0;
    const response = await resolver(request(), new AbortController().signal);
    assert.equal(response.schemaVersion, 1);
    assert.equal(response.outcomes[0].status, "resolved");
    assert.equal(response.faces[0].sha256, fileSha256);

    await assert.rejects(
      resolver(request({ familyStack: ["Bad\u0000Family"] }), new AbortController().signal),
      (error) => error.code === "resource_policy_failure" && error.phase === "font_loading",
    );
    await assert.rejects(
      resolver({
        schemaVersion: 1,
        requests: [request().requests[0], { ...request().requests[0], id: "face-2" }],
      }, new AbortController().signal),
      (error) => error.code === "resource_limit" && error.phase === "font_loading",
    );
  });
});

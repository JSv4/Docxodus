// The Node resolver → Playwright binding → browser runtime path is the mechanism #442 turns
// on: fonts.test.mjs exercises discovery/resolver in isolation, and npm/tests/font-runtime.spec.ts
// exercises the browser runtime against a hand-written in-page resolver, but nothing previously
// connected them through the real `__docxodusResolveFonts` binding. This test does.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { mkdir, mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { after, before, describe, test } from "node:test";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright-core";
import { renderDocxArtifacts } from "../dist/index.js";

const here = dirname(fileURLToPath(import.meta.url));
const packageRoot = dirname(here);
const repositoryRoot = dirname(packageRoot);
const fontFixture = join(repositoryRoot, "docs", "demo", "fonts", "docxodus-canvas-mono.woff2");
const baseOptions = Object.freeze({
  reviewProfile: "final",
  commentProfile: "hidden",
  timeoutMs: 120_000,
});
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function digest(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function xml(value) {
  return Buffer.from(value, "utf8");
}

function storedZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  for (const entry of entries) {
    const name = Buffer.from(entry.name, "utf8");
    const checksum = crc32(entry.data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(entry.data.length, 18);
    local.writeUInt32LE(entry.data.length, 22);
    local.writeUInt16LE(name.length, 26);
    localParts.push(local, name, entry.data);

    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt32LE(checksum, 16);
    central.writeUInt32LE(entry.data.length, 20);
    central.writeUInt32LE(entry.data.length, 24);
    central.writeUInt16LE(name.length, 28);
    central.writeUInt32LE(offset, 42);
    centralParts.push(central, name);
    offset += local.length + name.length + entry.data.length;
  }
  const directory = Buffer.concat(centralParts);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(entries.length, 8);
  end.writeUInt16LE(entries.length, 10);
  end.writeUInt32LE(directory.length, 12);
  end.writeUInt32LE(offset, 16);
  return new Uint8Array(Buffer.concat([...localParts, directory, end]));
}

// Requests "Docxodus Canvas Mono" — the exact family the fixture webfont declares — so a
// correctly-wired resolver resolves it exactly rather than falling through to substitution.
function generateFontDirectoryDocx() {
  return storedZip([
    {
      name: "[Content_Types].xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`),
    },
    {
      name: "_rels/.rels",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/>
</Relationships>`),
    },
    {
      name: "word/document.xml",
      data: xml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>
  <w:p><w:r><w:rPr><w:rFonts w:ascii="Docxodus Canvas Mono" w:hAnsi="Docxodus Canvas Mono"/></w:rPr><w:t>Configured font round trip.</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>
</w:body></w:document>`),
    },
  ]);
}

let browser;

before(async () => {
  browser = await chromium.launch({ headless: true });
});

after(async () => {
  await browser?.close();
});

describe("fontDirectories end to end", { concurrency: false }, () => {
  test("resolves a configured face through the Playwright binding into the render report", async () => {
    const fontBytes = await readFile(fontFixture);
    const root = await mkdtemp(join(tmpdir(), "docxodus-font-directory-"));
    try {
      const directory = join(root, "fonts");
      await mkdir(directory);
      await writeFile(join(directory, "docxodus-canvas-mono.woff2"), fontBytes);

      const source = generateFontDirectoryDocx();
      const result = await renderDocxArtifacts(source, {
        ...baseOptions,
        browser,
        outputs: ["html"],
        fontDirectories: [directory],
        fontLicenseAttestations: [{
          schemaVersion: 1,
          usage: "standalone-document-font-embedding",
          fileSha256: digest(fontBytes),
          embeddingPermitted: true,
          permittedOutputs: ["html", "pdf"],
          subsettingPermitted: true,
          basis: "Docxodus test fixture license",
          attester: "Docxodus test suite",
        }],
      });

      assert.equal(result.renderReport.fonts.length, 1);
      const [font] = result.renderReport.fonts;
      assert.equal(font.requestedFamily, "Docxodus Canvas Mono");
      assert.equal(font.status, "resolved");
      assert.equal(font.source, "attested");
      assert.equal(font.faceMatch, "exact");
      assert.equal(font.fileSha256, digest(fontBytes));
      assert.equal(result.renderReport.fontIdentity.resolverContract,
        "https://docxodus.dev/contracts/font-resolver/v1");
      assert.equal(
        result.renderReport.warnings.some((warning) => warning.code === "font_unavailable"),
        false,
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });

  test("carries a Node-side discovery failure's real code and remediation across the binding", async () => {
    const root = await mkdtemp(join(tmpdir(), "docxodus-font-directory-"));
    try {
      const directory = join(root, "fonts");
      await mkdir(directory);
      const linkTarget = join(root, "outside.woff2");
      await writeFile(linkTarget, await readFile(fontFixture));
      const { symlink } = await import("node:fs/promises");
      await symlink(linkTarget, join(directory, "linked.woff2"));

      const source = generateFontDirectoryDocx();
      await assert.rejects(
        renderDocxArtifacts(source, {
          ...baseOptions,
          browser,
          outputs: ["html"],
          fontDirectories: [directory],
        }),
        (error) => error.code === "resource_policy_failure"
          && error.phase === "font_loading"
          && /symlink/.test(error.message)
          && /[Ss]ymlink/.test(error.remediation),
      );
    } finally {
      await rm(root, { recursive: true, force: true });
    }
  });
});

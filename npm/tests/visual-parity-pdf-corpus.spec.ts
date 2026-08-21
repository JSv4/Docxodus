import { test, expect } from "@playwright/test";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  PDF_PARITY_CORPUS,
  PDF_PARITY_CORPUS_SCHEMA_VERSION,
  REQUIRED_PDF_PARITY_CATEGORIES,
  type PdfLinkExpectation,
  type PdfParityCorpusEntry,
} from "./visual-parity/pdf-corpus.js";
import {
  assertSafeCaseId,
  resolveTrackedRegularFile,
} from "./visual-parity/benchmark-paths.js";
import { pinExecutable } from "./visual-parity/toolchain.js";

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, "../..");
const sha256Pattern = /^[0-9a-f]{64}$/;
const gitObjectPattern = /^[0-9a-f]{40}$/;
const pdfCases: readonly PdfParityCorpusEntry[] = PDF_PARITY_CORPUS.cases;
const git = pinExecutable("git", ["--version"]);

function sha256(bytes: Uint8Array): string {
  return createHash("sha256").update(bytes).digest("hex");
}

/** Minimal read-only ZIP reader for fixture-shape assertions; DOCX uses stored or deflated parts. */
function zipParts(path: string): Map<string, Buffer> {
  const bytes = readFileSync(path);
  const minimumEocd = 22;
  const earliestEocd = Math.max(0, bytes.length - 65_557);
  let eocd = -1;
  for (let offset = bytes.length - minimumEocd; offset >= earliestEocd; offset--) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) {
      eocd = offset;
      break;
    }
  }
  if (eocd < 0) throw new Error(`${path} has no ZIP end-of-central-directory record`);

  const count = bytes.readUInt16LE(eocd + 10);
  let cursor = bytes.readUInt32LE(eocd + 16);
  const parts = new Map<string, Buffer>();
  for (let index = 0; index < count; index++) {
    if (bytes.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error(`${path} has an invalid central-directory record at ${cursor}`);
    }
    const method = bytes.readUInt16LE(cursor + 10);
    const compressedSize = bytes.readUInt32LE(cursor + 20);
    const nameLength = bytes.readUInt16LE(cursor + 28);
    const extraLength = bytes.readUInt16LE(cursor + 30);
    const commentLength = bytes.readUInt16LE(cursor + 32);
    const localOffset = bytes.readUInt32LE(cursor + 42);
    const name = bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString("utf8");

    if (bytes.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`${path}:${name} has an invalid local-file record`);
    }
    const localNameLength = bytes.readUInt16LE(localOffset + 26);
    const localExtraLength = bytes.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressed = bytes.subarray(dataStart, dataStart + compressedSize);
    if (!name.endsWith("/")) {
      if (method === 0) parts.set(name, Buffer.from(compressed));
      else if (method === 8) parts.set(name, inflateRawSync(compressed));
      else throw new Error(`${path}:${name} uses unsupported ZIP compression method ${method}`);
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  return parts;
}

function xml(parts: Map<string, Buffer>, name: string): string {
  const part = parts.get(name);
  if (!part) throw new Error(`DOCX is missing ${name}`);
  return part.toString("utf8");
}

function occurrences(value: string, pattern: RegExp): number {
  return value.match(pattern)?.length ?? 0;
}

function regexEscape(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function decodedText(xmlFragment: string): string {
  return xmlFragment
    .replace(/<[^>]+>/g, "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/\s+/g, " ")
    .trim();
}

function byId(id: string): PdfParityCorpusEntry {
  const entry = pdfCases.find((candidate) => candidate.id === id);
  if (!entry) throw new Error(`Missing PDF corpus case ${id}`);
  return entry;
}

function sourceParts(id: string): Map<string, Buffer> {
  const entry = byId(id);
  return zipParts(resolveTrackedRegularFile(repoRoot, entry.source.path));
}

function packageText(parts: Map<string, Buffer>): string {
  return [...parts]
    .filter(([name]) => name.startsWith("word/") && name.endsWith(".xml"))
    .map(([, bytes]) => decodedText(bytes.toString("utf8")))
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

test.describe("versioned generated-PDF corpus", () => {
  test("is complete, compact, explicitly profiled, and semantically inspectable", () => {
    expect(PDF_PARITY_CORPUS.schemaVersion).toBe(PDF_PARITY_CORPUS_SCHEMA_VERSION);
    expect(pdfCases.length).toBeLessThanOrEqual(10);

    const ids = pdfCases.map((entry) => entry.id);
    const paths = pdfCases.map((entry) => entry.source.path);
    expect(new Set(ids).size).toBe(ids.length);
    expect(new Set(paths).size).toBe(paths.length);

    const covered = new Set(pdfCases.flatMap((entry) => entry.categories));
    expect(REQUIRED_PDF_PARITY_CATEGORIES.filter((category) => !covered.has(category))).toEqual([]);

    for (const entry of pdfCases) {
      assertSafeCaseId(entry.id);
      expect(entry.rationale.trim(), `${entry.id} corpus rationale`).not.toBe("");
      expect(entry.disposition.rationale.trim(), `${entry.id} disposition rationale`).not.toBe("");
      expect(entry.semantics.requiredText.length, `${entry.id} searchable text contract`).toBeGreaterThan(0);
      expect(entry.profiles.candidate).toEqual({ reviewProfile: "final", commentProfile: "hidden" });
      expect(entry.profiles.reference.commentProjection).toBe("hidden");
      expect(entry.profiles.rationale.trim(), `${entry.id} profile rationale`).not.toBe("");
      expect(entry.source.sha256).toMatch(sha256Pattern);
      expect(entry.source.gitBlob).toMatch(gitObjectPattern);
      expect(entry.source.provenance.introducedBy).toMatch(gitObjectPattern);
      expect(entry.source.provenance.rationale.trim(), `${entry.id} provenance rationale`).not.toBe("");
    }

    const links = pdfCases.flatMap((entry) => entry.semantics.links ?? []);
    expect(links.filter((link) => link.kind === "internal")).toHaveLength(11);
    expect(links.filter((link) => link.kind === "external")).toEqual([{
      kind: "external",
      sourceText: "EricWhite.com",
      relationshipId: "rId4",
      exactTarget: "http://www.ericwhite.com",
      expectedPdfTarget: "http://www.ericwhite.com/",
      expectedPdfAnnotations: 1,
    }]);
  });

  test("pins every tracked fixture and generator to exact repository bytes", () => {
    const checkedGenerators = new Set<string>();
    for (const entry of pdfCases) {
      const path = resolveTrackedRegularFile(repoRoot, entry.source.path);
      execFileSync(git.path, ["ls-files", "--error-unmatch", entry.source.path], {
        cwd: repoRoot,
        stdio: "pipe",
      });
      expect(sha256(readFileSync(path)), `${entry.id} fixture SHA-256`).toBe(entry.source.sha256);
      expect(execFileSync(git.path, ["hash-object", entry.source.path], {
        cwd: repoRoot,
        encoding: "utf8",
      }).trim(), `${entry.id} fixture Git blob`).toBe(entry.source.gitBlob);
      expect(execFileSync(git.path, ["rev-parse", `HEAD:${entry.source.path}`], {
        cwd: repoRoot,
        encoding: "utf8",
      }).trim(), `${entry.id} fixture blob at HEAD`).toBe(entry.source.gitBlob);

      const generator = entry.source.provenance.generator;
      if (!generator || checkedGenerators.has(generator.path)) continue;
      checkedGenerators.add(generator.path);
      const generatorPath = resolveTrackedRegularFile(repoRoot, generator.path);
      execFileSync(git.path, ["ls-files", "--error-unmatch", generator.path], {
        cwd: repoRoot,
        stdio: "pipe",
      });
      expect(sha256(readFileSync(generatorPath)), `${generator.path} SHA-256`).toBe(generator.sha256);
      expect(execFileSync(git.path, ["hash-object", generator.path], {
        cwd: repoRoot,
        encoding: "utf8",
      }).trim(), `${generator.path} Git blob`).toBe(generator.gitBlob);
      expect(execFileSync(git.path, ["rev-parse", `HEAD:${generator.path}`], {
        cwd: repoRoot,
        encoding: "utf8",
      }).trim(), `${generator.path} blob at HEAD`).toBe(generator.gitBlob);
      expect(generator.rationale.trim()).not.toBe("");
    }
  });

  test("binds every strict text token to content in its pinned source package", () => {
    for (const entry of pdfCases) {
      const text = packageText(sourceParts(entry.id));
      for (const required of entry.semantics.requiredText) {
        expect(text, `${entry.id} required text ${JSON.stringify(required)}`).toContain(required);
      }
      for (const hidden of entry.semantics.forbiddenText ?? []) {
        expect(text, `${entry.id} hidden/final-view text ${JSON.stringify(hidden)}`).toContain(hidden);
      }
    }
  });

  test("the pinned packages contain every claimed high-risk document shape", () => {
    const legal = sourceParts("pdf-legal-contract");
    expect(xml(legal, "word/document.xml")).toContain("<w:tbl>");

    const columns = sourceParts("pdf-two-column-section");
    expect(xml(columns, "word/document.xml")).toMatch(/<w:cols\b[^>]*w:num="2"/);

    const footnote = sourceParts("pdf-footnote");
    expect(xml(footnote, "word/document.xml")).toContain("<w:footnoteReference");
    expect(xml(footnote, "word/footnotes.xml")).toContain("This is a test.");

    const endnote = sourceParts("pdf-endnote-table");
    expect(xml(endnote, "word/document.xml")).toContain("<w:endnoteReference");
    expect(xml(endnote, "word/endnotes.xml")).toContain("<w:tbl>");

    const mixed = sourceParts("pdf-mixed-orientation-running-stories");
    const mixedDocument = xml(mixed, "word/document.xml");
    const pageSizes = mixedDocument.match(/<w:pgSz\b[^>]*\/>/g) ?? [];
    expect(pageSizes.some((tag) => /w:orient="landscape"/.test(tag))).toBe(true);
    expect(pageSizes.some((tag) => !/w:orient=/.test(tag))).toBe(true);
    expect(occurrences(mixedDocument, /<w:headerReference\b/g)).toBeGreaterThan(0);
    expect(occurrences(mixedDocument, /<w:footerReference\b/g)).toBeGreaterThan(0);
    expect([...mixed.keys()].some((name) => /^word\/header\d+\.xml$/.test(name))).toBe(true);
    expect([...mixed.keys()].some((name) => /^word\/footer\d+\.xml$/.test(name))).toBe(true);

    const image = sourceParts("pdf-raster-image");
    expect(xml(image, "word/document.xml")).toContain("<w:drawing>");
    expect([...image.keys()].some((name) => name.startsWith("word/media/"))).toBe(true);

    const chart = sourceParts("pdf-chart");
    expect(xml(chart, "word/document.xml")).toContain("<c:chart");
    expect([...chart.keys()].some((name) => /^word\/charts\/chart\d+\.xml$/.test(name))).toBe(true);

    const revisions = sourceParts("pdf-final-revisions");
    expect(xml(revisions, "word/document.xml")).toContain("<w:del ");
    expect(xml(revisions, "word/document.xml")).toContain("powerful");
    expect(byId("pdf-final-revisions").profiles.reference.revisionProjection).toBe("accepted");

    const comments = sourceParts("pdf-hidden-comments");
    expect(occurrences(xml(comments, "word/document.xml"), /<w:commentReference\b/g)).toBe(10);
    expect(occurrences(xml(comments, "word/comments.xml"), /<w:comment\b/g)).toBe(10);
    expect(occurrences(xml(comments, "word/commentsExtended.xml"), /<w15:commentEx\b/g)).toBe(10);
    expect(byId("pdf-hidden-comments").profiles.candidate.commentProfile).toBe("hidden");
  });

  test("pins exact internal destinations and the external URI", () => {
    const checkLink = (link: PdfLinkExpectation, document: string, relationships: string): void => {
      if (link.kind === "external") {
        expect(document).toMatch(new RegExp(
          `<w:hyperlink\\b[^>]*r:id="${regexEscape(link.relationshipId)}"[^>]*>[\\s\\S]*?${regexEscape(link.sourceText)}[\\s\\S]*?</w:hyperlink>`,
        ));
        expect(relationships).toMatch(new RegExp(
          `<Relationship\\b(?=[^>]*\\bId="${regexEscape(link.relationshipId)}")(?=[^>]*\\bTarget="${regexEscape(link.exactTarget)}")(?=[^>]*\\bTargetMode="External")[^>]*/>`,
        ));
        return;
      }

      const hyperlink = document.match(new RegExp(
        `<w:hyperlink\\b[^>]*w:anchor="${regexEscape(link.anchor)}"[^>]*>([\\s\\S]*?)</w:hyperlink>`,
      ));
      expect(hyperlink, `missing hyperlink to ${link.anchor}`).not.toBeNull();
      expect(decodedText(hyperlink![1])).toContain(link.sourceText);
      const bookmark = document.match(new RegExp(
        `<w:bookmarkStart\\b[^>]*w:name="${regexEscape(link.anchor)}"[^>]*/>([\\s\\S]*?)<w:bookmarkEnd\\b`,
      ));
      expect(bookmark, `missing bookmark ${link.anchor}`).not.toBeNull();
      expect(decodedText(bookmark![1])).toContain(link.destinationText);
    };

    for (const entry of pdfCases) {
      if (!entry.semantics.links?.length) continue;
      const parts = sourceParts(entry.id);
      const document = xml(parts, "word/document.xml");
      const relationships = xml(parts, "word/_rels/document.xml.rels");
      for (const link of entry.semantics.links) checkLink(link, document, relationships);
    }
  });
});

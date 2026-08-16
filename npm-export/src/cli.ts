#!/usr/bin/env node
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parseArgs } from "node:util";
import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  COMMENT_PROFILES,
  DocxodusExportError,
  renderDocxFile,
  REVIEW_PROFILES,
  type CommentProfile,
  type ExportResourceLimits,
  type FontLicenseAttestation,
  type NodeExportOptions,
  type RenderEnvironmentAttestation,
  type ReviewProfile,
} from "./index.js";
import { canonicalJsonBytes } from "./canonical.js";
import { writeDiagnosticNoReplace } from "./files.js";

const HELP = `Usage:
  docxodus convert <input.docx> --to <html|pdf> --output <path>
    --review-profile <final|original|markup> --comments <hidden|inline|endnotes|margin>

Options:
  --document-version <integer>       Immutable source version (default: 0)
  --expected-source-digest <sha256>  Required source digest
  --unsupported-content <warn|strict>
  --strict-fonts
  --timeout <milliseconds>
  --browser-executable <path>        Or DOCXODUS_CHROMIUM_PATH
  --font-directory <path>            Repeatable; earlier directories take precedence
  --font-license-attestations <path>
  --environment-attestation <path>
  --limit <name=integer>              Repeatable; values may only lower defaults
  --title <text>
  --report <path>
  --page-map <path>
`;

function integer(value: string | undefined, label: string): number | undefined {
  if (value === undefined) return undefined;
  if (!/^(?:0|[1-9][0-9]*)$/.test(value)) {
    throw new Error(`${label} must be a non-negative integer.`);
  }
  const parsed = Number(value);
  if (!Number.isSafeInteger(parsed)) throw new Error(`${label} exceeds JavaScript's safe range.`);
  return parsed;
}

function oneOf<T extends string>(
  value: string | undefined,
  allowed: readonly T[],
  label: string,
): T {
  if (!value || !allowed.includes(value as T)) {
    throw new Error(`${label} must be one of: ${allowed.join(", ")}.`);
  }
  return value as T;
}

function parseLimits(values: string[] | undefined): Partial<ExportResourceLimits> | undefined {
  if (!values || values.length === 0) return undefined;
  const result: Partial<ExportResourceLimits> = {};
  const keys = new Set(Object.keys(DEFAULT_EXPORT_RESOURCE_LIMITS) as Array<keyof ExportResourceLimits>);
  for (const expression of values) {
    const match = /^([^=]+)=([1-9][0-9]*)$/.exec(expression);
    if (!match) throw new Error(`--limit must use name=positive-integer: ${expression}`);
    const key = match[1] as keyof ExportResourceLimits;
    if (!keys.has(key)) throw new Error(`Unknown --limit key: ${match[1]}`);
    if (result[key] !== undefined) throw new Error(`Duplicate --limit key: ${key}`);
    const value = Number(match[2]);
    if (!Number.isSafeInteger(value)) throw new Error(`--limit ${key} exceeds the safe integer range.`);
    result[key] = value;
  }
  return result;
}

async function readJson<T>(path: string | undefined, label: string): Promise<T | undefined> {
  if (!path) return undefined;
  try {
    return JSON.parse(await readFile(resolve(path), "utf8")) as T;
  } catch (cause) {
    throw new Error(`${label} is not readable JSON: ${cause instanceof Error ? cause.message : cause}`);
  }
}

function humanError(error: unknown): string {
  if (error instanceof DocxodusExportError) {
    return `${error.code} (${error.phase}): ${error.message}\nRemediation: ${error.remediation}`
      + (error.detail ? `\nDetail: ${error.detail}` : "")
      + (error.committedDestinations.length
        ? `\nAlready committed: ${error.committedDestinations.join(", ")}`
        : "");
  }
  return error instanceof Error ? error.message : String(error);
}

export async function runCli(argv: readonly string[]): Promise<number> {
  let reportPath: string | undefined;
  try {
    const parsed = parseArgs({
      args: [...argv],
      allowPositionals: true,
      strict: true,
      options: {
        to: { type: "string" },
        output: { type: "string", short: "o" },
        "review-profile": { type: "string" },
        comments: { type: "string" },
        "document-version": { type: "string" },
        "expected-source-digest": { type: "string" },
        "unsupported-content": { type: "string" },
        "strict-fonts": { type: "boolean", default: false },
        timeout: { type: "string" },
        "browser-executable": { type: "string" },
        "font-directory": { type: "string", multiple: true },
        "font-license-attestations": { type: "string" },
        "environment-attestation": { type: "string" },
        limit: { type: "string", multiple: true },
        title: { type: "string" },
        report: { type: "string" },
        "page-map": { type: "string" },
        help: { type: "boolean", short: "h", default: false },
      },
    });
    if (parsed.values.help) {
      process.stderr.write(HELP);
      return 0;
    }
    if (parsed.positionals.length !== 2 || parsed.positionals[0] !== "convert") {
      throw new Error("Expected `docxodus convert <input.docx>`.\n\n" + HELP);
    }
    const inputPath = parsed.positionals[1];
    const target = oneOf(parsed.values.to, ["html", "pdf"] as const, "--to");
    const outputPath = parsed.values.output;
    if (!outputPath) throw new Error("--output is required.");
    const reviewProfile = oneOf<ReviewProfile>(
      parsed.values["review-profile"],
      REVIEW_PROFILES,
      "--review-profile",
    );
    const commentProfile = oneOf<CommentProfile>(
      parsed.values.comments,
      COMMENT_PROFILES,
      "--comments",
    );
    reportPath = parsed.values.report;
    const fontLicenseAttestations = await readJson<FontLicenseAttestation[]>(
      parsed.values["font-license-attestations"],
      "--font-license-attestations",
    );
    const environmentAttestation = await readJson<RenderEnvironmentAttestation>(
      parsed.values["environment-attestation"],
      "--environment-attestation",
    );
    const timeoutMs = integer(parsed.values.timeout, "--timeout");
    if (timeoutMs === 0) throw new Error("--timeout must be positive.");
    const options: NodeExportOptions = {
      reviewProfile,
      commentProfile,
      documentVersion: integer(parsed.values["document-version"], "--document-version"),
      expectedSourceDigest: parsed.values["expected-source-digest"],
      unsupportedContent: parsed.values["unsupported-content"] === undefined
        ? undefined
        : oneOf(parsed.values["unsupported-content"], ["warn", "strict"] as const,
          "--unsupported-content"),
      strictFonts: parsed.values["strict-fonts"],
      timeoutMs,
      browserExecutablePath: parsed.values["browser-executable"]
        ?? process.env.DOCXODUS_CHROMIUM_PATH,
      fontDirectories: parsed.values["font-directory"],
      fontLicenseAttestations,
      environmentAttestation,
      limits: parseLimits(parsed.values.limit),
      title: parsed.values.title,
    };
    const result = await renderDocxFile(inputPath, {
      ...(target === "pdf" ? { pdfPath: outputPath } : { htmlPath: outputPath }),
      reportPath,
      pageMapPath: parsed.values["page-map"],
    }, options);
    process.stderr.write(
      `Rendered ${result.pageCount} page${result.pageCount === 1 ? "" : "s"}; `
      + `fingerprint ${result.rendererFingerprint}.\n`,
    );
    return 0;
  } catch (error) {
    if (reportPath && error instanceof DocxodusExportError && error.report) {
      try {
        await writeDiagnosticNoReplace(reportPath, canonicalJsonBytes(error.report));
      } catch (reportError) {
        process.stderr.write(`Failed to preserve render report: ${humanError(reportError)}\n`);
      }
    }
    process.stderr.write(`${humanError(error)}\n`);
    return error instanceof DocxodusExportError ? 1 : 2;
  }
}

if (process.argv[1] && fileURLToPath(import.meta.url) === resolve(process.argv[1])) {
  process.exitCode = await runCli(process.argv.slice(2));
}

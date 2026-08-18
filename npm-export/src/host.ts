#!/usr/bin/env node
import { lstat, open } from "node:fs/promises";
import type { Stats } from "node:fs";
import { dirname, resolve } from "node:path";
import {
  DEFAULT_EXPORT_RESOURCE_LIMITS,
  DEFAULT_EXPORT_TIMEOUT_MS,
  HARD_EXPORT_TIMEOUT_MS,
  renderDocxArtifacts,
  DocxodusExportError,
  type RenderBatchOptions,
  type RenderBatchResult,
} from "./index.js";
import type { FontLicenseAttestation } from "./contracts.js";
import { openOwnedExportBrowserSession } from "./browser-session.js";
import { canonicalJson, canonicalJsonBytes, sha256 } from "./canonical.js";
import { discoverFontCatalog } from "./fonts/index.js";
import { decodeStrictUtf8, strictJsonParse } from "./strict-json.js";

const MAX_CONTROL_FRAME_BYTES = 8_388_608;
const MAX_BATCHES = 64;
const MAX_SOURCES = 64;
const MAX_ARTIFACTS = MAX_BATCHES * 4;
const MAX_AGGREGATE_INPUT_BYTES = 536_870_912;
const MAX_AGGREGATE_OUTPUT_BYTES = 1_073_741_824;
const MAX_ID_CHARACTERS = 256;
const MAX_ARTIFACT_REQUEST_IDS = 256;
const WRITE_CHUNK_BYTES = 65_536;
const MAX_ERROR_TEXT_CHARACTERS = 16_384;
const RESPONSE_WRITE_TIMEOUT_MS = 30_000;
const MAX_FONT_POLICY_BYTES = 1_048_576;
const DOCX_MEDIA_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

interface HostSourceDescriptor {
  id: string;
  byteLength: number;
  sha256: string;
  mediaType: typeof DOCX_MEDIA_TYPE;
}

interface HostBatchRequest {
  id: string;
  sourceId: string;
  artifactRequestIds: string[];
  options: Omit<
    RenderBatchOptions,
    "browser" | "browserExecutablePath" | "fontDirectories" | "fontLicenseAttestations" | "signal"
  >;
}

interface HostRequest {
  schemaVersion: 1;
  sources: HostSourceDescriptor[];
  batches: HostBatchRequest[];
}

interface HostFontPolicy {
  schemaVersion: 1;
  fontDirectories: readonly string[];
  fontLicenseAttestations: readonly FontLicenseAttestation[];
}

function sameFileIdentity(left: Stats, right: Stats): boolean {
  return left.dev === right.dev
    && left.ino === right.ino
    && left.mode === right.mode
    && left.size === right.size
    && left.mtimeMs === right.mtimeMs
    && left.ctimeMs === right.ctimeMs;
}

async function readStablePolicyFile(
  path: string,
  signal: AbortSignal,
): Promise<Buffer> {
  if (signal.aborted) throw new Error("The host request was cancelled while reading its font policy.");
  const pathBefore = await lstat(path);
  if (!pathBefore.isFile() || !Number.isSafeInteger(pathBefore.size)
    || pathBefore.size <= 0 || pathBefore.size > MAX_FONT_POLICY_BYTES) {
    throw new Error("invalid policy file");
  }
  const handle = await open(path, "r");
  try {
    const opened = await handle.stat();
    if (!opened.isFile() || !sameFileIdentity(pathBefore, opened)) {
      throw new Error("font policy changed before it was opened");
    }
    const bytes = Buffer.alloc(opened.size);
    let offset = 0;
    while (offset < bytes.byteLength) {
      if (signal.aborted) {
        throw new Error("The host request was cancelled while reading its font policy.");
      }
      const { bytesRead } = await handle.read(bytes, offset, bytes.byteLength - offset, offset);
      if (bytesRead === 0) throw new Error("font policy ended before its declared size");
      offset += bytesRead;
    }
    const trailing = Buffer.alloc(1);
    if ((await handle.read(trailing, 0, 1, bytes.byteLength)).bytesRead !== 0) {
      throw new Error("font policy grew while it was read");
    }
    const [openedAfter, pathAfter] = await Promise.all([handle.stat(), lstat(path)]);
    if (!sameFileIdentity(opened, openedAfter) || !sameFileIdentity(opened, pathAfter)) {
      throw new Error("font policy changed while it was read");
    }
    return bytes;
  } finally {
    await handle.close();
  }
}

type ArtifactKind = "html" | "pdf" | "pageMap" | "renderReport";

interface ResponseArtifact {
  id: string;
  batchId: string;
  kind: ArtifactKind;
  mediaType: string;
  byteLength: number;
  sha256: string;
  bytes: Uint8Array;
}

class FrameReader {
  readonly #iterator = process.stdin[Symbol.asyncIterator]();
  #current = Buffer.alloc(0);
  #offset = 0;

  async #nextChunk(signal: AbortSignal): Promise<void> {
    if (signal.aborted) throw new Error("The host request was cancelled while reading stdin.");
    let listener: (() => void) | undefined;
    try {
      const next = await Promise.race([
        this.#iterator.next(),
        new Promise<never>((_, reject) => {
          listener = () => reject(new Error("The host request was cancelled while reading stdin."));
          signal.addEventListener("abort", listener, { once: true });
        }),
      ]);
      if (next.done) throw new Error("The export host input ended before one complete frame arrived.");
      this.#current = Buffer.isBuffer(next.value) ? next.value : Buffer.from(next.value);
      this.#offset = 0;
    } finally {
      if (listener) signal.removeEventListener("abort", listener);
    }
  }

  async #readExact(target: Uint8Array, signal: AbortSignal): Promise<void> {
    let written = 0;
    while (written < target.byteLength) {
      if (this.#offset >= this.#current.byteLength) await this.#nextChunk(signal);
      const count = Math.min(
        target.byteLength - written,
        this.#current.byteLength - this.#offset,
      );
      target.set(this.#current.subarray(this.#offset, this.#offset + count), written);
      this.#offset += count;
      written += count;
    }
  }

  async readFrame(maximumBytes: number, signal: AbortSignal, exactBytes?: number): Promise<Buffer> {
    const header = Buffer.allocUnsafe(4);
    await this.#readExact(header, signal);
    const length = header.readUInt32BE(0);
    if (length === 0 || length > maximumBytes || (exactBytes !== undefined && length !== exactBytes)) {
      throw new Error(
        exactBytes === undefined
          ? `Invalid host frame length: ${length}`
          : `Host blob frame length ${length} does not match its declaration ${exactBytes}.`,
      );
    }
    const payload = Buffer.allocUnsafe(length);
    await this.#readExact(payload, signal);
    return payload;
  }

  async assertEnd(signal: AbortSignal): Promise<void> {
    if (this.#offset < this.#current.byteLength) {
      throw new Error("The export host request contains trailing bytes.");
    }
    while (!signal.aborted) {
      const next = await this.#iterator.next();
      if (next.done) return;
      if (Buffer.byteLength(next.value) > 0) {
        throw new Error("The export host request contains trailing bytes.");
      }
    }
    throw new Error("The host request was cancelled while checking the frame boundary.");
  }
}

function exactKeys(value: object, allowed: readonly string[], label: string): void {
  const unknown = Object.keys(value).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) throw new Error(`${label} contains unknown fields: ${unknown.join(", ")}`);
}

function validId(value: unknown, label: string): string {
  if (typeof value !== "string" || value.length === 0 || value.length > MAX_ID_CHARACTERS
    || value !== value.trim() || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new Error(`${label} must be 1-${MAX_ID_CHARACTERS} printable characters without edge whitespace.`);
  }
  return value;
}

function boundedErrorText(value: unknown): string {
  const text = value instanceof Error ? value.message : String(value);
  return text.length <= MAX_ERROR_TEXT_CHARACTERS
    ? text
    : `${text.slice(0, MAX_ERROR_TEXT_CHARACTERS - 3)}...`;
}

function boundTransportError(record: Record<string, unknown>): Record<string, unknown> {
  const result = { ...record };
  for (const key of ["message", "remediation", "detail", "partUri", "anchorId", "resource"]) {
    if (typeof result[key] === "string") result[key] = boundedErrorText(result[key]);
  }
  for (const key of ["pending", "committedDestinations"]) {
    if (Array.isArray(result[key])) {
      result[key] = result[key].slice(0, 64).map((entry) => boundedErrorText(entry));
    }
  }
  return result;
}

function positiveSafeInteger(value: unknown, label: string, maximum: number): number {
  if (!Number.isSafeInteger(value) || (value as number) <= 0 || (value as number) > maximum) {
    throw new Error(`${label} must be an integer from 1 through ${maximum}.`);
  }
  return value as number;
}

function validateRequest(value: unknown): HostRequest {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("The host request must be an object.");
  }
  const request = value as Partial<HostRequest>;
  exactKeys(request, ["schemaVersion", "sources", "batches"], "The host request");
  if (request.schemaVersion !== 1 || !Array.isArray(request.sources)
    || !Array.isArray(request.batches)) {
    throw new Error("The host request must use schemaVersion 1 with sources and batches arrays.");
  }
  if (request.sources.length > MAX_SOURCES) {
    throw new Error(`The host accepts at most ${MAX_SOURCES} unique sources.`);
  }
  if (request.batches.length > MAX_BATCHES) {
    throw new Error(`The host accepts at most ${MAX_BATCHES} render batches.`);
  }

  const sourceIds = new Set<string>();
  const sourceIdentities = new Set<string>();
  let aggregateInputBytes = 0;
  for (const [index, source] of request.sources.entries()) {
    if (!source || typeof source !== "object" || Array.isArray(source)) {
      throw new Error(`Source ${index} is malformed.`);
    }
    exactKeys(source, ["id", "byteLength", "sha256", "mediaType"], `Source ${index}`);
    const id = validId(source.id, `Source ${index} id`);
    if (sourceIds.has(id)) throw new Error(`Duplicate source id: ${id}`);
    sourceIds.add(id);
    const byteLength = positiveSafeInteger(
      source.byteLength,
      `Source ${index} byteLength`,
      DEFAULT_EXPORT_RESOURCE_LIMITS.compressedDocxBytes,
    );
    aggregateInputBytes += byteLength;
    if (aggregateInputBytes > MAX_AGGREGATE_INPUT_BYTES) {
      throw new Error(`Host source bytes exceed the ${MAX_AGGREGATE_INPUT_BYTES} byte aggregate limit.`);
    }
    if (typeof source.sha256 !== "string" || !/^[0-9a-f]{64}$/.test(source.sha256)) {
      throw new Error(`Source ${index} sha256 must be a lower-case SHA-256 digest.`);
    }
    const sourceIdentity = `${source.byteLength}:${source.sha256}`;
    if (sourceIdentities.has(sourceIdentity)) {
      throw new Error(`Source ${index} duplicates an already declared source blob.`);
    }
    sourceIdentities.add(sourceIdentity);
    if (source.mediaType !== DOCX_MEDIA_TYPE) {
      throw new Error(`Source ${index} has an unsupported mediaType.`);
    }
  }

  const batchIds = new Set<string>();
  const referencedSources = new Set<string>();
  const requestArtifactIds = new Set<string>();
  for (const [index, batch] of request.batches.entries()) {
    if (!batch || typeof batch !== "object" || Array.isArray(batch)
      || !batch.options || typeof batch.options !== "object" || Array.isArray(batch.options)) {
      throw new Error(`Batch ${index} is malformed.`);
    }
    exactKeys(batch, ["id", "sourceId", "artifactRequestIds", "options"], `Batch ${index}`);
    const id = validId(batch.id, `Batch ${index} id`);
    if (batchIds.has(id)) throw new Error(`Duplicate batch id: ${id}`);
    batchIds.add(id);
    const sourceId = validId(batch.sourceId, `Batch ${index} sourceId`);
    if (!sourceIds.has(sourceId)) throw new Error(`Batch ${index} refers to unknown source: ${sourceId}`);
    referencedSources.add(sourceId);
    if (!Array.isArray(batch.artifactRequestIds)
      || batch.artifactRequestIds.length > MAX_ARTIFACT_REQUEST_IDS) {
      throw new Error(
        `Batch ${index} artifactRequestIds must be an array of at most ${MAX_ARTIFACT_REQUEST_IDS} ids.`,
      );
    }
    let previousArtifactId: string | undefined;
    for (const [artifactIndex, artifactIdValue] of batch.artifactRequestIds.entries()) {
      const artifactId = validId(
        artifactIdValue,
        `Batch ${index} artifactRequestIds[${artifactIndex}]`,
      );
      if (previousArtifactId !== undefined && previousArtifactId >= artifactId) {
        throw new Error(`Batch ${index} artifactRequestIds must be unique and code-unit sorted.`);
      }
      if (requestArtifactIds.has(artifactId)) {
        throw new Error(`Artifact request id appears in more than one batch: ${artifactId}`);
      }
      requestArtifactIds.add(artifactId);
      if (requestArtifactIds.size > MAX_ARTIFACT_REQUEST_IDS) {
        throw new Error(`The host accepts at most ${MAX_ARTIFACT_REQUEST_IDS} artifact request ids.`);
      }
      previousArtifactId = artifactId;
    }
    exactKeys(batch.options, [
      "outputs",
      "documentVersion",
      "expectedSourceDigest",
      "reviewProfile",
      "reviewProfileAlreadyApplied",
      "commentProfile",
      "title",
      "unsupportedContent",
      "strictFonts",
      "timeoutMs",
      "limits",
      "environmentAttestation",
    ], `Batch ${index} options`);
    if (!Array.isArray(batch.options.outputs) || batch.options.outputs.length > 2
      || batch.options.outputs.some((output) => output !== "html" && output !== "pdf")
      || new Set(batch.options.outputs).size !== batch.options.outputs.length) {
      throw new Error(`Batch ${index} outputs must contain html, pdf, or both exactly once.`);
    }
    if (batch.options.timeoutMs !== undefined) {
      positiveSafeInteger(batch.options.timeoutMs, `Batch ${index} timeoutMs`, HARD_EXPORT_TIMEOUT_MS);
    }
  }
  if (referencedSources.size !== sourceIds.size) {
    throw new Error("Every declared source must be referenced by at least one render batch.");
  }
  return request as HostRequest;
}

async function waitForDrain(signal: AbortSignal): Promise<void> {
  let abortListener: (() => void) | undefined;
  let drainListener: (() => void) | undefined;
  try {
    await new Promise<void>((resolve, reject) => {
      drainListener = resolve;
      abortListener = () => reject(new Error("The host response deadline expired."));
      process.stdout.once("drain", drainListener);
      signal.addEventListener("abort", abortListener, { once: true });
    });
  } finally {
    if (drainListener) process.stdout.removeListener("drain", drainListener);
    if (abortListener) signal.removeEventListener("abort", abortListener);
  }
}

async function writeBytes(bytes: Uint8Array, signal: AbortSignal): Promise<void> {
  for (let offset = 0; offset < bytes.byteLength; offset += WRITE_CHUNK_BYTES) {
    if (signal.aborted) throw new Error("The host response deadline expired.");
    const chunk = bytes.subarray(offset, Math.min(bytes.byteLength, offset + WRITE_CHUNK_BYTES));
    if (!process.stdout.write(chunk)) await waitForDrain(signal);
  }
}

async function writeFrame(
  bytes: Uint8Array,
  maximumBytes: number,
  signal: AbortSignal,
): Promise<void> {
  if (bytes.byteLength === 0 || bytes.byteLength > maximumBytes || bytes.byteLength > 0xffff_ffff) {
    throw new Error(`The host response frame has an invalid length: ${bytes.byteLength}.`);
  }
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(bytes.byteLength);
  await writeBytes(header, signal);
  await writeBytes(bytes, signal);
}

function responseArtifact(
  id: string,
  batchId: string,
  kind: ArtifactKind,
  mediaType: string,
  bytes: Uint8Array,
): ResponseArtifact {
  return { id, batchId, kind, mediaType, byteLength: bytes.byteLength, sha256: sha256(bytes), bytes };
}

function collectArtifacts(
  batchIndex: number,
  batchId: string,
  result: RenderBatchResult,
  remainingBytes: number,
): {
  artifacts: ResponseArtifact[];
  artifactIds: Partial<Record<ArtifactKind, string>>;
  byteLength: number;
} {
  const artifacts: ResponseArtifact[] = [];
  const artifactIds: Partial<Record<ArtifactKind, string>> = {};
  let byteLength = 0;
  const reserve = (length: number): void => {
    if (!Number.isSafeInteger(length) || length < 0 || length > remainingBytes - byteLength) {
      throw new DocxodusExportError(
        "resource_limit",
        "output_verification",
        `Host artifact bytes exceed the ${MAX_AGGREGATE_OUTPUT_BYTES} byte aggregate limit.`,
        "Split the logical delivery build into smaller bounded requests.",
      );
    }
    byteLength += length;
  };
  const add = (kind: ArtifactKind, mediaType: string, bytes: Uint8Array): void => {
    if (artifacts.length >= 4) throw new Error("A render batch produced too many artifacts.");
    const id = `b${batchIndex}-${kind}`;
    artifactIds[kind] = id;
    artifacts.push(responseArtifact(id, batchId, kind, mediaType, bytes));
  };
  if (result.html !== undefined) {
    const length = Buffer.byteLength(result.html, "utf8");
    reserve(length);
    add("html", "text/html; charset=utf-8", Buffer.from(result.html, "utf8"));
  }
  if (result.pdf !== undefined) {
    reserve(result.pdf.byteLength);
    add("pdf", "application/pdf", result.pdf);
  }
  const pageMap = canonicalJson(result.pageMap);
  reserve(Buffer.byteLength(pageMap, "utf8"));
  add("pageMap", "application/json; charset=utf-8", Buffer.from(pageMap, "utf8"));
  const report = canonicalJson(result.renderReport);
  const reportLength = Buffer.byteLength(report, "utf8");
  const reportMaximum = result.renderReport.options.policy.limits.renderReportOutputBytes;
  if (reportLength > reportMaximum) {
    throw new DocxodusExportError(
      "resource_limit",
      "output_verification",
      `renderReportOutputBytes limit exceeded after host bindings (${reportLength} > ${reportMaximum}).`,
      "Use fewer artifact request ids or a smaller render report.",
    );
  }
  reserve(reportLength);
  add("renderReport", "application/json; charset=utf-8", Buffer.from(report, "utf8"));
  return { artifacts, artifactIds, byteLength };
}

async function loadHostFontPolicy(
  path: string | undefined,
  signal: AbortSignal,
): Promise<HostFontPolicy> {
  if (path === undefined) {
    return Object.freeze({
      schemaVersion: 1 as const,
      fontDirectories: Object.freeze([]),
      fontLicenseAttestations: Object.freeze([]),
    });
  }
  if (path.trim() === "" || /[\u0000-\u001f\u007f]/u.test(path)) {
    throw new Error("DOCXODUS_FONT_POLICY_PATH must name one policy file.");
  }
  const policyPath = resolve(path);
  let bytes: Buffer;
  try {
    bytes = await readStablePolicyFile(policyPath, signal);
  } catch {
    throw new Error("The configured host font policy is not a readable bounded regular file.");
  }
  if (bytes.byteLength === 0 || bytes.byteLength > MAX_FONT_POLICY_BYTES) {
    throw new Error("The configured host font policy exceeds its byte limit.");
  }
  const value = strictJsonParse(decodeStrictUtf8(bytes, "Host font policy"));
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("The configured host font policy must be an object.");
  }
  const record = value as Record<string, unknown>;
  exactKeys(record, ["schemaVersion", "fontDirectories", "fontLicenseAttestations"],
    "The host font policy");
  if (record.schemaVersion !== 1 || !Array.isArray(record.fontDirectories)
    || !Array.isArray(record.fontLicenseAttestations)) {
    throw new Error("The host font policy must use schemaVersion 1 and both documented arrays.");
  }
  if (!record.fontDirectories.every((entry) => typeof entry === "string" && entry.trim() !== "")) {
    throw new Error("The host font policy contains an invalid font directory.");
  }
  if (record.fontDirectories.length > DEFAULT_EXPORT_RESOURCE_LIMITS.fontDirectoryEntries
    || record.fontLicenseAttestations.length > DEFAULT_EXPORT_RESOURCE_LIMITS.fontFiles) {
    throw new Error("The host font policy exceeds the configured font-count limits.");
  }
  if (record.fontLicenseAttestations.length > 0 && record.fontDirectories.length === 0) {
    throw new Error("The host font policy cannot attach attestations without a font directory.");
  }
  const digests = new Set<string>();
  const attestations = record.fontLicenseAttestations.map((entry, index): FontLicenseAttestation => {
    if (!entry || typeof entry !== "object" || Array.isArray(entry)) {
      throw new Error(`The host font policy attestation ${index} is not an object.`);
    }
    const item = entry as Record<string, unknown>;
    exactKeys(item, [
      "schemaVersion", "usage", "fileSha256", "embeddingPermitted", "permittedOutputs",
      "subsettingPermitted", "basis", "attester",
    ], `The host font policy attestation ${index}`);
    if (item.schemaVersion !== 1 || item.usage !== "standalone-document-font-embedding"
      || item.embeddingPermitted !== true || typeof item.fileSha256 !== "string"
      || !/^[0-9a-f]{64}$/.test(item.fileSha256) || digests.has(item.fileSha256)
      || !Array.isArray(item.permittedOutputs) || item.permittedOutputs.length === 0
      || item.permittedOutputs.length > 2
      || item.permittedOutputs.some((output) => output !== "html" && output !== "pdf")
      || new Set(item.permittedOutputs).size !== item.permittedOutputs.length
      || typeof item.subsettingPermitted !== "boolean"
      || typeof item.basis !== "string" || item.basis.trim() === ""
      || /[\u0000-\u001f\u007f]/u.test(item.basis)
      || (item.attester !== undefined && (typeof item.attester !== "string"
        || item.attester.trim() === "" || /[\u0000-\u001f\u007f]/u.test(item.attester)))) {
      throw new Error(`The host font policy attestation ${index} is invalid or duplicated.`);
    }
    digests.add(item.fileSha256);
    return Object.freeze({
      schemaVersion: 1,
      usage: "standalone-document-font-embedding",
      fileSha256: item.fileSha256,
      embeddingPermitted: true,
      permittedOutputs: Object.freeze([
        ...(item.permittedOutputs.includes("html") ? ["html" as const] : []),
        ...(item.permittedOutputs.includes("pdf") ? ["pdf" as const] : []),
      ]),
      subsettingPermitted: item.subsettingPermitted,
      basis: item.basis,
      ...(item.attester === undefined ? {} : { attester: item.attester }),
    });
  });
  const policyDirectory = dirname(policyPath);
  return Object.freeze({
    schemaVersion: 1 as const,
    fontDirectories: Object.freeze(record.fontDirectories.map((entry) =>
      resolve(policyDirectory, entry as string))),
    fontLicenseAttestations: Object.freeze(attestations),
  });
}

async function main(): Promise<void> {
  const controller = new AbortController();
  const timer = setTimeout(() => {
    controller.abort();
    process.stdin.destroy();
  }, HARD_EXPORT_TIMEOUT_MS);
  timer.unref();
  const reader = new FrameReader();
  let response: unknown;
  let responseArtifacts: ResponseArtifact[] = [];
  let activeBatchId: string | undefined;
  let activeArtifactRequestIds: string[] | undefined;
  let session: Awaited<ReturnType<typeof openOwnedExportBrowserSession>> | undefined;
  try {
    const control = await reader.readFrame(MAX_CONTROL_FRAME_BYTES, controller.signal);
    const request = validateRequest(strictJsonParse(decodeStrictUtf8(control, "Host control frame")));
    const sourceBytes = new Map<string, Uint8Array>();
    for (const source of request.sources) {
      const bytes = await reader.readFrame(source.byteLength, controller.signal, source.byteLength);
      const digest = sha256(bytes);
      if (digest !== source.sha256) {
        throw new Error(`Source ${source.id} digest mismatch: expected ${source.sha256}; actual ${digest}.`);
      }
      sourceBytes.set(source.id, bytes);
    }
    await reader.assertEnd(controller.signal);

    const fontPolicy = await loadHostFontPolicy(
      process.env.DOCXODUS_FONT_POLICY_PATH,
      controller.signal,
    );
    if (fontPolicy.fontDirectories.length > 0) {
      await discoverFontCatalog(
        fontPolicy.fontDirectories,
        fontPolicy.fontLicenseAttestations,
        DEFAULT_EXPORT_RESOURCE_LIMITS,
        controller.signal,
      );
    }

    if (request.batches.length > 0) {
      session = await openOwnedExportBrowserSession(
        process.env.DOCXODUS_CHROMIUM_PATH,
        DEFAULT_EXPORT_TIMEOUT_MS,
        controller.signal,
      );
    }
    const batches = [];
    let aggregateOutputBytes = 0;
    for (const [index, batch] of request.batches.entries()) {
      activeBatchId = batch.id;
      activeArtifactRequestIds = batch.artifactRequestIds;
      const source = sourceBytes.get(batch.sourceId)!;
      const result = await renderDocxArtifacts(source, {
        ...batch.options,
        fontDirectories: fontPolicy.fontDirectories,
        fontLicenseAttestations: fontPolicy.fontLicenseAttestations,
        signal: controller.signal,
        browser: session?.browser,
      });
      result.renderReport.bindings.artifactRequestIds = [...batch.artifactRequestIds];
      const collected = collectArtifacts(
        index,
        batch.id,
        result,
        MAX_AGGREGATE_OUTPUT_BYTES - aggregateOutputBytes,
      );
      aggregateOutputBytes += collected.byteLength;
      responseArtifacts.push(...collected.artifacts);
      batches.push({
        id: batch.id,
        sourceId: batch.sourceId,
        pageCount: result.pageCount,
        rendererFingerprint: result.rendererFingerprint,
        artifacts: collected.artifactIds,
      });
      activeBatchId = undefined;
      activeArtifactRequestIds = undefined;
    }
    if (responseArtifacts.length > MAX_ARTIFACTS) throw new Error("The host produced too many artifacts.");
    await session?.close();
    session = undefined;
    response = {
      schemaVersion: 1,
      batches,
      artifacts: responseArtifacts.map(({ bytes: _bytes, ...descriptor }) => descriptor),
    };
  } catch (error) {
    let cleanupFailure: unknown;
    if (session) {
      await session.close().catch((cause) => { cleanupFailure = cause; });
      session = undefined;
    }
    responseArtifacts = [];
    let fatal: Record<string, unknown>;
    if (error instanceof DocxodusExportError) {
      if (error.report?.partial?.bindings && activeArtifactRequestIds) {
        error.report.partial.bindings.artifactRequestIds = [...activeArtifactRequestIds];
      }
      const { report, ...transportError } = error.toJSON();
      fatal = boundTransportError(transportError);
      if (report !== undefined && activeBatchId !== undefined) {
        const reportBytes = canonicalJsonBytes(report);
        const reportLimit = error.report?.options.policy.limits.renderReportOutputBytes
          ?? DEFAULT_EXPORT_RESOURCE_LIMITS.renderReportOutputBytes;
        if (reportBytes.byteLength <= Math.min(MAX_AGGREGATE_OUTPUT_BYTES, reportLimit)) {
          const diagnostic = responseArtifact(
            "failure-renderReport",
            activeBatchId,
            "renderReport",
            "application/json; charset=utf-8",
            reportBytes,
          );
          responseArtifacts.push(diagnostic);
          fatal.reportArtifactId = diagnostic.id;
        } else {
          fatal.reportUnavailable = "failed render report exceeds the host diagnostic-artifact limit";
        }
      }
    } else {
      fatal = {
        name: error instanceof Error ? error.name : "Error",
        severity: "error",
        message: boundedErrorText(error),
      };
    }
    if (cleanupFailure !== undefined) {
      fatal.cleanupFailure = {
        name: cleanupFailure instanceof Error ? cleanupFailure.name : "Error",
        message: boundedErrorText(cleanupFailure),
      };
    }
    response = {
      schemaVersion: 1,
      fatal,
      ...(responseArtifacts.length === 0
        ? {}
        : {
          diagnosticArtifacts: responseArtifacts.map(({ bytes: _bytes, ...descriptor }) => descriptor),
        }),
    };
  } finally {
    await session?.close().catch(() => undefined);
  }
  clearTimeout(timer);
  const responseController = new AbortController();
  const responseTimer = setTimeout(() => responseController.abort(), RESPONSE_WRITE_TIMEOUT_MS);
  responseTimer.unref();
  try {
    await writeFrame(canonicalJsonBytes(response), MAX_CONTROL_FRAME_BYTES, responseController.signal);
    for (const artifact of responseArtifacts) {
      await writeFrame(artifact.bytes, MAX_AGGREGATE_OUTPUT_BYTES, responseController.signal);
    }
  } finally {
    clearTimeout(responseTimer);
  }
}

await main();

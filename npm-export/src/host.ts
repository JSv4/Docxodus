#!/usr/bin/env node
import { once } from "node:events";
import {
  renderDocxArtifacts,
  DocxodusExportError,
  type RenderBatchOptions,
} from "./index.js";
import { openOwnedExportBrowserSession } from "./browser-session.js";
import { DEFAULT_EXPORT_TIMEOUT_MS } from "docxodus/export-browser";

const MAX_FRAME_BYTES = 536_870_912;

interface HostBatchRequest {
  id: string;
  documentBase64: string;
  options: Omit<
    RenderBatchOptions,
    "browser" | "browserExecutablePath" | "fontDirectories"
  >;
}

interface HostRequest {
  schemaVersion: 1;
  batches: HostBatchRequest[];
}

async function readFrame(): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let total = 0;
  let expected: number | undefined;
  for await (const chunk of process.stdin) {
    const bytes = Buffer.from(chunk);
    chunks.push(bytes);
    total += bytes.byteLength;
    if (total > MAX_FRAME_BYTES + 4) throw new Error("The host request exceeds its frame limit.");
    if (expected === undefined && total >= 4) {
      const prefix = Buffer.concat(chunks, total);
      expected = prefix.readUInt32BE(0);
      chunks.length = 0;
      chunks.push(prefix);
      if (expected === 0 || expected > MAX_FRAME_BYTES) {
        throw new Error(`Invalid host frame length: ${expected}`);
      }
    }
    if (expected !== undefined && total >= expected + 4) {
      if (total !== expected + 4) {
          throw new Error("The export host accepts exactly one framed request.");
      }
      return Buffer.concat(chunks, total).subarray(4);
    }
  }
  throw new Error("The export host input ended before one complete frame arrived.");
}

function exactKeys(value: object, allowed: readonly string[], label: string): void {
  const unknown = Object.keys(value).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) throw new Error(`${label} contains unknown fields: ${unknown.join(", ")}`);
}

function validateRequest(value: unknown): HostRequest {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("The host request must be an object.");
  }
  const request = value as Partial<HostRequest>;
  exactKeys(request, ["schemaVersion", "batches"], "The host request");
  if (request.schemaVersion !== 1 || !Array.isArray(request.batches)) {
    throw new Error("The host request must use schemaVersion 1 and a batches array.");
  }
  const seen = new Set<string>();
  for (const [index, batch] of request.batches.entries()) {
    if (!batch || typeof batch !== "object"
      || typeof batch.id !== "string" || batch.id.trim() === ""
      || typeof batch.documentBase64 !== "string"
      || !batch.options || typeof batch.options !== "object") {
      throw new Error(`Batch ${index} is malformed.`);
    }
    exactKeys(batch, ["id", "documentBase64", "options"], `Batch ${index}`);
    exactKeys(batch.options, [
      "outputs",
      "documentVersion",
      "expectedSourceDigest",
      "reviewProfile",
      "commentProfile",
      "title",
      "unsupportedContent",
      "strictFonts",
      "timeoutMs",
      "limits",
      "fontLicenseAttestations",
      "environmentAttestation",
    ], `Batch ${index} options`);
    if (seen.has(batch.id)) throw new Error(`Duplicate batch id: ${batch.id}`);
    seen.add(batch.id);
  }
  return request as HostRequest;
}

function decodeCanonicalBase64(value: string): Uint8Array {
  if (value.length % 4 !== 0
    || !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(value)) {
    throw new Error("documentBase64 must be canonical padded base64 without whitespace.");
  }
  const bytes = Buffer.from(value, "base64");
  if (bytes.toString("base64") !== value) {
    throw new Error("documentBase64 is not canonical base64 data.");
  }
  return new Uint8Array(bytes);
}

async function writeFrame(value: unknown): Promise<void> {
  const payload = Buffer.from(JSON.stringify(value), "utf8");
  if (payload.byteLength > MAX_FRAME_BYTES) throw new Error("The host response exceeds its frame limit.");
  const header = Buffer.allocUnsafe(4);
  header.writeUInt32BE(payload.byteLength);
  if (!process.stdout.write(Buffer.concat([header, payload]))) await once(process.stdout, "drain");
}

async function main(): Promise<void> {
  let response: unknown;
  let session: Awaited<ReturnType<typeof openOwnedExportBrowserSession>> | undefined;
  try {
    const request = validateRequest(JSON.parse((await readFrame()).toString("utf8")));
    if (request.batches.length > 0) {
      session = await openOwnedExportBrowserSession(
        process.env.DOCXODUS_CHROMIUM_PATH,
        DEFAULT_EXPORT_TIMEOUT_MS,
      );
    }
    const batches = [];
    for (const batch of request.batches) {
      try {
        const source = decodeCanonicalBase64(batch.documentBase64);
        const result = await renderDocxArtifacts(source, {
          ...batch.options,
          browser: session?.browser,
        });
        batches.push({
          id: batch.id,
          ok: true,
          result: {
            ...result,
            ...(result.pdf ? { pdfBase64: Buffer.from(result.pdf).toString("base64") } : {}),
            pdf: undefined,
          },
        });
      } catch (error) {
        batches.push({
          id: batch.id,
          ok: false,
          error: error instanceof DocxodusExportError
            ? error.toJSON()
            : { name: "Error", message: error instanceof Error ? error.message : String(error) },
        });
      }
    }
    response = { schemaVersion: 1, batches };
    await session?.close();
    session = undefined;
  } catch (error) {
    response = {
      schemaVersion: 1,
      fatal: error instanceof DocxodusExportError
        ? error.toJSON()
        : {
          name: error instanceof Error ? error.name : "Error",
          message: error instanceof Error ? error.message : String(error),
        },
    };
  } finally {
    await session?.close().catch(() => undefined);
  }
  await writeFrame(response);
}

await main();

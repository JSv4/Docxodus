import type {
  DeliverableVerificationRequest,
  DeliverableCompanionArtifactInput,
} from "./types.js";

/**
 * Serialize the full deliverable-verification request (issue #747) to the wire shape every
 * transport parses: companion bytes become `bytesB64`, everything else passes through. Kept
 * separate from the API modules so the direct API, the session, and the worker proxy share
 * one encoding.
 */
export function serializeVerificationRequest(request: DeliverableVerificationRequest): string {
  if (request === null || typeof request !== "object" || Array.isArray(request)) {
    throw new RangeError("a verification request must be an object");
  }
  const { companionArtifacts, ...rest } = request;
  const wire: Record<string, unknown> = { ...rest };
  if (companionArtifacts !== undefined) {
    wire.companionArtifacts = companionArtifacts.map(encodeCompanion);
  }
  return JSON.stringify(wire);
}

function encodeCompanion(artifact: DeliverableCompanionArtifactInput): Record<string, unknown> {
  const { bytes, ...rest } = artifact;
  const wire: Record<string, unknown> = { ...rest };
  if (bytes !== undefined) wire.bytesB64 = bytesToBase64(bytes);
  return wire;
}

function bytesToBase64(bytes: Uint8Array): string {
  if (typeof btoa === "function") {
    let binary = "";
    const chunkSize = 0x8000;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, chunk as unknown as number[]);
    }
    return btoa(binary);
  }
  return Buffer.from(bytes).toString("base64");
}

import {
  fontFamilyKey,
  FONT_RESOLVER_CONTRACT_ID,
  FONT_RESOLVER_SCHEMA_VERSION,
  FONT_SUBSTITUTION_CONTRACT,
  FONT_SUBSTITUTION_CONTRACT_MATERIAL,
  FONT_SUBSTITUTION_CONTRACT_VERSION,
  normalizeFontFamilyName,
  type ExportResourceLimits,
  type FontRequest,
  type FontResolver,
  type FontResolverFace,
  type FontResolverOutcome,
  type FontResolverRequest,
  type FontResolverResponse,
} from "docxodus/export-browser";
import { canonicalJson, sha256 } from "../canonical.js";
import type { FontLicenseAttestation, RenderOutput } from "../contracts.js";
import { exportError } from "../contracts.js";
import {
  discoverFontCatalog,
  faceKey,
  FONT_CATALOG_CODE_POINTS_MAX,
  type ConfiguredFontFace,
  type FontCatalog,
} from "./discovery.js";

import { LocalFontIndex, type LocalFontOptions } from "./local.js";

const SUBSTITUTION_CONTRACT_DIGEST = sha256(canonicalJson(FONT_SUBSTITUTION_CONTRACT_MATERIAL));

function policyError(message: string, remediation: string, detail?: string): never {
  exportError("resource_policy_failure", "font_loading", message, remediation,
    detail ? { detail } : {});
}

function limitError(name: keyof ExportResourceLimits, actual: number, maximum: number): never {
  exportError(
    "resource_limit",
    "font_loading",
    `${name} limit exceeded (${actual} > ${maximum}).`,
    "Reduce the number of distinct document font requests or sampled code points.",
  );
}

// Same shape as npm/src/font-runtime.ts's exactKeys, duplicated rather than shared across the
// package boundary; each closes over its own error type.
function exactKeys(value: object, allowed: readonly string[], label: string): void {
  const unknown = Object.keys(value).filter((key) => !allowed.includes(key));
  if (unknown.length > 0) {
    policyError(`${label} contains unknown fields.`, "Use the versioned font resolver contract.",
      unknown.sort().join(","));
  }
}

function wellFormed(value: string): boolean {
  for (let index = 0; index < value.length; index++) {
    const unit = value.charCodeAt(index);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = value.charCodeAt(++index);
      if (!(next >= 0xdc00 && next <= 0xdfff)) return false;
    } else if (unit >= 0xdc00 && unit <= 0xdfff) return false;
  }
  return true;
}

function validateRequestRecord(value: unknown, index: number): FontRequest {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    policyError(`Font request ${index} is not an object.`, "Use the versioned font resolver contract.");
  }
  const record = value as unknown as Record<string, unknown>;
  exactKeys(record, [
    "id", "familyStack", "familyKinds", "style", "weight", "stretch", "sampleCodePoints",
  ],
    `Font request ${index}`);
  if (typeof record.id !== "string" || !/^[a-zA-Z0-9._:-]{1,128}$/.test(record.id)) {
    policyError(`Font request ${index} has an invalid id.`, "Use a short opaque request identifier.");
  }
  if (!Array.isArray(record.familyStack) || record.familyStack.length === 0
    || record.familyStack.length > 64 || !record.familyStack.every((family) =>
      typeof family === "string" && wellFormed(family) && normalizeFontFamilyName(family).length > 0
      && normalizeFontFamilyName(family).length <= 256
      && !/[\u0000-\u001f\u007f]/u.test(normalizeFontFamilyName(family)))
    || (record.familyStack as string[]).reduce((total, family) =>
      total + normalizeFontFamilyName(family).length, 0) > 4096) {
    policyError(`Font request ${index} has an invalid family stack.`,
      "Provide one to 64 bounded CSS family names.");
  }
  if (!Array.isArray(record.familyKinds)
    || record.familyKinds.length !== record.familyStack.length
    || record.familyKinds.some((kind) => kind !== "named" && kind !== "generic")) {
    policyError(`Font request ${index} has invalid family syntax classifications.`,
      "Provide exactly one named or generic classification per family.");
  }
  if (record.style !== "normal" && record.style !== "italic" && record.style !== "oblique") {
    policyError(`Font request ${index} has an invalid style.`, "Use normal, italic, or oblique.");
  }
  if (!Number.isSafeInteger(record.weight) || (record.weight as number) < 1
    || (record.weight as number) > 1000) {
    policyError(`Font request ${index} has an invalid weight.`, "Use an integer from 1 through 1000.");
  }
  if (typeof record.stretch !== "number" || !Number.isFinite(record.stretch)
    || record.stretch < 50 || record.stretch > 200) {
    policyError(`Font request ${index} has an invalid stretch.`, "Use a percentage from 50 through 200.");
  }
  if (!Array.isArray(record.sampleCodePoints)) {
    policyError(`Font request ${index} has invalid sample code points.`, "Provide a sorted integer array.");
  }
  let prior = -1;
  for (const codePoint of record.sampleCodePoints) {
    if (!Number.isSafeInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff
      || (codePoint >= 0xd800 && codePoint <= 0xdfff) || codePoint <= prior) {
      policyError(`Font request ${index} sample code points are not sorted Unicode scalars.`,
        "Deduplicate and sort sampled Unicode scalar values.");
    }
    prior = codePoint;
  }
  return {
    id: record.id,
    familyStack: Object.freeze((record.familyStack as string[]).map(normalizeFontFamilyName)),
    familyKinds: Object.freeze([...(record.familyKinds as Array<"named" | "generic">)]),
    style: record.style,
    weight: record.weight as number,
    stretch: record.stretch,
    sampleCodePoints: Object.freeze([...(record.sampleCodePoints as number[])]),
  };
}

function validateResolverRequest(
  value: FontResolverRequest,
  limits: Readonly<ExportResourceLimits>,
): { requests: readonly FontRequest[]; rejectedFaceIds: readonly string[] } {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    policyError("The font resolver request is not an object.", "Use the versioned font resolver contract.");
  }
  const record = value as unknown as Record<string, unknown>;
  exactKeys(record, ["schemaVersion", "requests", "rejectedFaceIds"], "The font resolver request");
  if (record.schemaVersion !== FONT_RESOLVER_SCHEMA_VERSION || !Array.isArray(record.requests)) {
    policyError("The font resolver request has an unsupported schema.",
      "Use font resolver schema version 1.");
  }
  if (record.requests.length > limits.fontRequests) {
    limitError("fontRequests", record.requests.length, limits.fontRequests);
  }
  const rejected = record.rejectedFaceIds === undefined ? [] : record.rejectedFaceIds;
  if (!Array.isArray(rejected) || rejected.length > limits.fontFiles
    || rejected.some(id => typeof id !== "string" || !/^[a-zA-Z0-9._:-]{1,128}$/.test(id))
    || new Set(rejected).size !== rejected.length) {
    policyError("The font resolver request has invalid rejectedFaceIds.", "Provide distinct previously returned local face ids within fontFiles.");
  }
  const ids = new Set<string>();
  let sampleCount = 0;
  const requests = record.requests.map((request, index) => {
    const validated = validateRequestRecord(request, index);
    if (ids.has(validated.id)) {
      policyError("The font resolver request contains duplicate ids.", "Use one id per distinct face request.");
    }
    ids.add(validated.id);
    sampleCount += validated.sampleCodePoints.length;
    if (sampleCount > limits.fontSampleCodePoints) {
      limitError("fontSampleCodePoints", sampleCount, limits.fontSampleCodePoints);
    }
    return validated;
  });
  return { requests: Object.freeze(requests), rejectedFaceIds: Object.freeze([...rejected]) };
}

function directionalStretchScore(requested: number, candidate: number): readonly number[] {
  return requested <= 100
    ? candidate <= requested ? [0, requested - candidate] : [1, candidate - requested]
    : candidate >= requested ? [0, candidate - requested] : [1, requested - candidate];
}

function directionalWeightScore(requested: number, candidate: number): readonly number[] {
  if (requested >= 400 && requested <= 500) {
    if (candidate >= requested && candidate <= 500) return [0, candidate - requested];
    if (candidate < requested) return [1, requested - candidate];
    return [2, candidate - 500];
  }
  return requested < 400
    ? candidate <= requested ? [0, requested - candidate] : [1, candidate - requested]
    : candidate >= requested ? [0, candidate - requested] : [1, requested - candidate];
}

function styleScore(requested: FontRequest["style"], candidate: ConfiguredFontFace["style"]): number {
  if (requested === candidate) return 0;
  if ((requested === "italic" && candidate === "oblique")
    || (requested === "oblique" && candidate === "italic")) return 1;
  return 2;
}

function candidateScore(request: FontRequest, face: Pick<ConfiguredFontFace, "directoryIndex" | "discoveryIndex" | "style" | "weight" | "stretch">): readonly number[] {
  const stretch = directionalStretchScore(request.stretch, face.stretch);
  const weight = directionalWeightScore(request.weight, face.weight);
  return [
    // A font root is an ordered deployment policy boundary, not another
    // face-selection tiebreaker. Once a family exists in an earlier root,
    // select the closest face from that root before considering later roots.
    face.directoryIndex,
    ...stretch,
    styleScore(request.style, face.style),
    ...weight,
    face.discoveryIndex,
  ];
}

function compareScore(left: readonly number[], right: readonly number[]): number {
  for (let index = 0; index < Math.max(left.length, right.length); index++) {
    const delta = (left[index] ?? 0) - (right[index] ?? 0);
    if (delta !== 0) return delta;
  }
  return 0;
}

function bestFace(
  catalog: FontCatalog,
  family: string,
  request: FontRequest,
): ConfiguredFontFace | undefined {
  const key = fontFamilyKey(family);
  return catalog.faces
    .filter((face) => face.familyKey === key)
    .sort((left, right) => compareScore(candidateScore(request, left), candidateScore(request, right)))[0];
}

interface Selection {
  face: ConfiguredFontFace;
  requestedFamily: string;
  substituted: boolean;
  metricCompatible: boolean;
}

function* familyCandidates(request: FontRequest): Generator<Omit<Selection, "face"> & { family: string }> {
  const primary = request.familyStack[0];
  for (const [index, family] of request.familyStack.entries()) {
    if (request.familyKinds[index] === "generic") continue;
    yield { family, requestedFamily: primary, substituted: index !== 0, metricCompatible: index === 0 };
    const substitution = FONT_SUBSTITUTION_CONTRACT.find(entry => fontFamilyKey(entry.family) === fontFamilyKey(family));
    if (substitution) yield { family: substitution.substitute, requestedFamily: primary,
      substituted: true, metricCompatible: index === 0 && substitution.metricCompatible };
  }
}

function selectFace(catalog: FontCatalog, request: FontRequest): Selection | undefined {
  for (const candidate of familyCandidates(request)) {
    const face = bestFace(catalog, candidate.family, request);
    if (face) return { ...candidate, face };
  }
  return undefined;
}

function resolverFace(face: ConfiguredFontFace): FontResolverFace {
  if (!face.license.ok) {
    policyError("A selected configured font has no legal embedding evidence.",
      "Remove the font or provide an exact WOFF/WOFF2 embedding-rights attestation.",
      face.sha256);
  }
  const bytes = new Uint8Array(face.bytes);
  if (sha256(bytes) !== face.sha256) {
    policyError("A configured font changed after catalog discovery.",
      "Rebuild the immutable font catalog and retry.", face.sha256);
  }
  return Object.freeze({
    id: face.id,
    ...(face.source ? { source: face.source } : {}),
    resolvedFamily: face.family,
    ...(face.postscriptName ? { postscriptName: face.postscriptName } : {}),
    version: face.version,
    style: face.style,
    weight: face.weight,
    stretch: face.stretch,
    format: face.format,
    mediaType: face.mediaType,
    byteLength: face.byteLength,
    sha256: face.sha256,
    bytesBase64: Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength).toString("base64"),
    licenseEvidence: Object.freeze({ ...face.license.evidence }),
  });
}

function includesCodePoint(values: readonly number[], target: number): boolean {
  let low = 0;
  let high = values.length - 1;
  while (low <= high) {
    const middle = (low + high) >>> 1;
    const value = values[middle];
    if (value === target) return true;
    if (value < target) low = middle + 1;
    else high = middle - 1;
  }
  return false;
}

export function resolveCatalogRequests(
  catalog: FontCatalog,
  requests: readonly FontRequest[],
  outputs: readonly RenderOutput[] = ["html"],
): Pick<FontResolverResponse, "outcomes" | "faces"> {
  const selectedFaces = new Map<string, FontResolverFace>();
  const outcomes: FontResolverOutcome[] = requests.map((request) => {
    const selection = selectFace(catalog, request);
    if (!selection) {
      return Object.freeze({
        requestId: request.id,
        requestedFamily: request.familyStack[0],
        status: "missing" as const,
      });
    }
    if (!selection.face.license.ok) {
      policyError(selection.face.license.failure,
        "Remove the font or provide an exact WOFF/WOFF2 embedding-rights attestation.",
        selection.face.sha256);
    }
    const forbiddenOutputs = outputs.filter((output) =>
      !selection.face.permittedOutputs.includes(output));
    if (forbiddenOutputs.length > 0) {
      policyError(
        "A selected configured font is not attested for every requested output.",
        "Expand the exact font attestation output scope or omit the disallowed output.",
        `${selection.face.sha256}:${forbiddenOutputs.join(",")}`,
      );
    }
    if (outputs.includes("pdf") && selection.face.license.evidence.noSubsetting) {
      policyError(
        "A selected configured font forbids subsetting and cannot be verified in Chromium PDF output.",
        "Use HTML-only output or supply a font whose license permits PDF subsetting.",
        selection.face.sha256,
      );
    }
    const exactFace = selection.face.style === request.style
      && selection.face.weight === request.weight
      && selection.face.stretch === request.stretch;
    const missingCodePoints = request.sampleCodePoints.filter((codePoint) =>
      !includesCodePoint(selection.face.codePoints, codePoint));
    if (!selectedFaces.has(selection.face.id)) {
      selectedFaces.set(selection.face.id, resolverFace(selection.face));
    }
    return Object.freeze({
      requestId: request.id,
      requestedFamily: selection.requestedFamily,
      resolvedFamily: selection.face.family,
      status: selection.substituted ? "substituted" as const : "resolved" as const,
      faceId: selection.face.id,
      metricCompatible: selection.metricCompatible,
      faceMatch: exactFace ? "exact" as const : "synthesized" as const,
      glyphCoverage: missingCodePoints.length === 0 ? "complete" as const : "partial" as const,
      ...(missingCodePoints.length === 0 ? {} : { missingCodePoints: Object.freeze(missingCodePoints) }),
    });
  });
  return {
    outcomes: Object.freeze(outcomes),
    faces: Object.freeze(Array.from(selectedFaces.values()).sort((left, right) =>
      left.id < right.id ? -1 : left.id > right.id ? 1 : 0)),
  };
}

export interface NodeFontRuntime {
  readonly resolver: FontResolver;
  prepare(signal?: AbortSignal): Promise<void>;
}

export function createNodeFontRuntime(
  directories: readonly string[],
  attestations: readonly FontLicenseAttestation[],
  limits: Readonly<ExportResourceLimits>,
  outputs: readonly RenderOutput[] = ["html"],
  local?: LocalFontOptions,
): NodeFontRuntime {
  const directorySnapshot = Object.freeze([...directories]);
  const attestationSnapshot = Object.freeze(attestations.map((entry) => Object.freeze({
    ...entry,
    permittedOutputs: Object.freeze([...entry.permittedOutputs]),
  })));
  const limitSnapshot = Object.freeze({ ...limits });
  const outputSnapshot = Object.freeze([...outputs]);
  let catalogPromise: Promise<FontCatalog> | undefined;
  const prepare = (signal?: AbortSignal): Promise<FontCatalog> => {
    catalogPromise ??= discoverFontCatalog(
      directorySnapshot,
      attestationSnapshot,
      limitSnapshot,
      signal,
    ).catch((error) => {
      catalogPromise = undefined;
      throw error;
    });
    return catalogPromise;
  };
  const waitForCatalog = (signal?: AbortSignal): Promise<FontCatalog> => {
    const pending = prepare(signal);
    if (!signal) return pending;
    if (signal.aborted) return Promise.reject(signal.reason);
    return new Promise((resolvePromise, rejectPromise) => {
      const onAbort = (): void => rejectPromise(signal.reason);
      signal.addEventListener("abort", onAbort, { once: true });
      pending.then(
        (catalog) => {
          signal.removeEventListener("abort", onAbort);
          resolvePromise(catalog);
        },
        (error) => {
          signal.removeEventListener("abort", onAbort);
          rejectPromise(error);
        },
      );
    });
  };
  let localPromise: Promise<LocalFontIndex> | undefined;
  const localFaces = new Map<string, ConfiguredFontFace>();
  const matchedGroups = new Map<string, { face: ConfiguredFontFace; path: string; size: number } | undefined>();
  const withLocalFaces = async (catalog: FontCatalog, requests: readonly FontRequest[], signal: AbortSignal): Promise<FontCatalog> => {
    if (!local) return catalog;
    localFaces.clear();
    for (const request of requests) {
      for (const candidate of familyCandidates(request)) {
        if (bestFace(catalog, candidate.family, request)) break;
        localPromise ??= (async () => {
          const index = new LocalFontIndex(local);
          await index.prepare(limitSnapshot, signal);
          return index;
        })();
        const index = await localPromise;
        const entries = index.candidates(candidate.family).map((entry, discoveryIndex) => ({
          ...entry, directoryIndex: 0, discoveryIndex,
        })).sort((a, b) => compareScore(candidateScore(request, a), candidateScore(request, b)));
        const groups = new Map<string, typeof entries>();
        for (const entry of entries) {
          const key = faceKey(entry);
          if (!groups.has(key)) groups.set(key, []);
          groups.get(key)!.push(entry);
        }
        let selected = false;
        for (const group of groups.values()) {
          const groupKey = faceKey(group[0]);
          if (!matchedGroups.has(groupKey)) {
            let match: { face: ConfiguredFontFace; path: string; size: number } | undefined;
            let ambiguous = false;
            for (const entry of group) {
              const face = await index.load(entry, limitSnapshot, outputSnapshot, signal);
              if (!face) continue;
              if (match && match.face.sha256 !== face.sha256) { ambiguous = true; break; }
              match = { face, path: entry.path, size: Number(entry.identity.size) };
            }
            if (ambiguous) {
              index.warning(candidate.family, "duplicate files disagree for the same family and face");
              match = undefined;
            }
            if (match) {
              const cached = [...matchedGroups.values()].flatMap(value => value ? [value] : []).concat(match);
              const localFiles = new Map(cached.map(value => [value.path, value.size]));
              const codePoints = [...catalog.faces, ...cached.map(value => value.face)].reduce((sum, face) => sum + face.codePoints.length, 0);
              if (codePoints > FONT_CATALOG_CODE_POINTS_MAX) exportError("resource_limit", "font_loading",
                `Font character maps exceed ${FONT_CATALOG_CODE_POINTS_MAX} entries.`, "Use fewer font faces in the document.");
              const count = catalog.fileCount + localFiles.size;
              if (count > limitSnapshot.fontFiles) limitError("fontFiles", count, limitSnapshot.fontFiles);
              const bytes = catalog.totalBytes + [...localFiles.values()].reduce((sum, size) => sum + size, 0);
              const expanded = catalog.totalExpandedBytes + cached.reduce((sum, value) => sum + value.face.expandedByteLength, 0);
              if (Math.max(bytes, expanded) > limitSnapshot.fontTotalBytes) limitError("fontTotalBytes", Math.max(bytes, expanded), limitSnapshot.fontTotalBytes);
            }
            matchedGroups.set(groupKey, match);
          }
          const match = matchedGroups.get(groupKey);
          if (!match) continue;
          localFaces.set(match.face.id, match.face);
          selected = true;
          break;
        }
        if (selected) break;
      }
    }
    return { ...catalog, faces: [...catalog.faces, ...localFaces.values()] };
  };
  const rejectedFaces = new Set<string>();
  const resolver = async (
    request: FontResolverRequest,
    signal: AbortSignal,
  ): Promise<FontResolverResponse> => {
    const { requests, rejectedFaceIds } = validateResolverRequest(request, limitSnapshot);
    for (const id of rejectedFaceIds) {
      if (rejectedFaces.has(id)) continue;
      const group = [...matchedGroups].find(([, match]) => match?.face.id === id);
      if (!local || !group) policyError("The browser rejected an unknown local font face.", "Reject only a previously returned local face.");
      if (rejectedFaces.size >= limitSnapshot.fontFiles) limitError("fontFiles", rejectedFaces.size + 1, limitSnapshot.fontFiles);
      const index = await localPromise!;
      index.warning(group[1]!.face.family, "browser could not decode the selected font bytes");
      matchedGroups.set(group[0], undefined);
      rejectedFaces.add(id);
    }
    const catalog = await withLocalFaces(await waitForCatalog(signal), requests, signal);
    if (signal.aborted) throw signal.reason;
    const resolved = resolveCatalogRequests(catalog, requests, outputSnapshot);
    return Object.freeze({
      schemaVersion: FONT_RESOLVER_SCHEMA_VERSION,
      resolverContract: FONT_RESOLVER_CONTRACT_ID,
      substitutionContractVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
      substitutionContractDigest: SUBSTITUTION_CONTRACT_DIGEST,
      ...(local ? { fallback: "generic" as const } : {}),
      outcomes: resolved.outcomes,
      faces: resolved.faces,
    });
  };
  return Object.freeze({
    resolver,
    async prepare(signal?: AbortSignal): Promise<void> {
      await waitForCatalog(signal);
    },
  });
}

export function createNodeFontResolver(
  directories: readonly string[],
  attestations: readonly FontLicenseAttestation[],
  limits: Readonly<ExportResourceLimits>,
  outputs: readonly RenderOutput[] = ["html"],
): FontResolver {
  return createNodeFontRuntime(directories, attestations, limits, outputs).resolver;
}

export function pathFreeCatalogManifest(catalog: FontCatalog): Record<string, unknown> {
  return {
    schemaVersion: 1,
    directoryCount: catalog.directoryCount,
    entryCount: catalog.entryCount,
    fileCount: catalog.fileCount,
    totalBytes: catalog.totalBytes,
    totalExpandedBytes: catalog.totalExpandedBytes,
    faces: catalog.faces.map((face) => ({
      id: face.id,
      directoryIndex: face.directoryIndex,
      discoveryIndex: face.discoveryIndex,
      family: face.family,
      postscriptName: face.postscriptName,
      version: face.version,
      style: face.style,
      weight: face.weight,
      stretch: face.stretch,
      format: face.format,
      byteLength: face.byteLength,
      expandedByteLength: face.expandedByteLength,
      sha256: face.sha256,
      glyphCount: face.codePoints.length,
      permittedOutputs: face.permittedOutputs,
      licenseEvidence: face.license.ok ? face.license.evidence : undefined,
      licenseFailure: face.license.ok ? undefined : face.license.failure,
    })),
  };
}

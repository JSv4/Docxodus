/**
 * Browser-portable font substitution policy shared by production export and
 * the visual-parity harness. Keep filesystem and fontconfig probes out of this
 * module so `docxodus/export-browser` remains safe to bundle for browsers.
 */

export const FONT_SUBSTITUTION_CONTRACT_VERSION = 1 as const;
export const FONT_RESOLVER_SCHEMA_VERSION = 1 as const;
export const FONT_RESOLVER_CONTRACT_ID = "https://docxodus.dev/contracts/font-resolver/v1" as const;

export interface FontSubstitutionEntry {
  /** Family declared by the source document. */
  family: string;
  /** License-safe configured family selected when an exact family is absent. */
  substitute: string;
  /** Whether the substitute is designed to preserve the source family's metrics. */
  metricCompatible: boolean;
}

/** NFC and whitespace normalization used before deterministic family matching. */
export function normalizeFontFamilyName(value: string): string {
  return value.normalize("NFC").trim().replace(/\s+/gu, " ");
}

/** Locale-independent case-insensitive lookup key for a normalized family. */
export function fontFamilyKey(value: string): string {
  return normalizeFontFamilyName(value).toLowerCase();
}

function substitution(entry: FontSubstitutionEntry): Readonly<FontSubstitutionEntry> {
  return Object.freeze(entry);
}

export const FONT_SUBSTITUTION_CONTRACT: ReadonlyArray<Readonly<FontSubstitutionEntry>> =
  Object.freeze([
    substitution({
      family: "Calibri",
      substitute: "Carlito",
      metricCompatible: true,
    }),
    // No open metric clone of the Light cut exists. Determinism requires both
    // renderers to make the same documented approximation.
    substitution({
      family: "Calibri Light",
      substitute: "Carlito",
      metricCompatible: false,
    }),
    substitution({
      family: "Cambria",
      substitute: "Caladea",
      metricCompatible: true,
    }),
    substitution({
      family: "Times New Roman",
      substitute: "Liberation Serif",
      metricCompatible: true,
    }),
    substitution({
      family: "Arial",
      substitute: "Liberation Sans",
      metricCompatible: true,
    }),
    substitution({
      family: "Courier New",
      substitute: "Liberation Mono",
      metricCompatible: true,
    }),
  ]);

/** Canonical path-free material both resolver sides hash for contract drift detection. */
export const FONT_SUBSTITUTION_CONTRACT_MATERIAL = Object.freeze({
  schemaVersion: FONT_SUBSTITUTION_CONTRACT_VERSION,
  entries: FONT_SUBSTITUTION_CONTRACT,
});

export type FontFaceStyle = "normal" | "italic" | "oblique";
export type FontFileFormat = "ttf" | "otf" | "woff" | "woff2";
export type FontMediaType = "font/ttf" | "font/otf" | "font/woff" | "font/woff2";
export type FontResolutionStatus =
  | "resolved"
  | "substituted"
  | "missing"
  | "load_failed"
  | "unverified";
export type FontResolutionSource = "browser" | "configured" | "attested";
export type FontFaceMatch = "exact" | "synthesized";
export type FontGlyphCoverage = "complete" | "partial" | "unverified";
export type FontEmbeddingKind = "installable" | "previewPrint" | "editable" | "attested";

/** A deterministic, text-free description of one computed CSS face request. */
export interface FontRequest {
  id: string;
  familyStack: readonly string[];
  style: FontFaceStyle;
  weight: number;
  /** CSS font-stretch as a percentage, where 100 is normal width. */
  stretch: number;
  /** Sorted, distinct Unicode scalar values sampled from nodes using this face. */
  sampleCodePoints: readonly number[];
}

export interface FontResolverRequest {
  schemaVersion: typeof FONT_RESOLVER_SCHEMA_VERSION;
  requests: readonly FontRequest[];
}

export interface FontLicenseEvidence {
  kind: FontEmbeddingKind;
  /** Canonical digest of path-free OS/2 or caller-attested evidence. */
  identity: string;
  noSubsetting: boolean;
}

/** Immutable configured bytes and metadata returned to the browser coordinator. */
export interface FontResolverFace {
  id: string;
  resolvedFamily: string;
  postscriptName?: string;
  version: string;
  style: FontFaceStyle;
  weight: number;
  stretch: number;
  format: FontFileFormat;
  mediaType: FontMediaType;
  byteLength: number;
  sha256: string;
  /** Canonical RFC 4648 padded base64. The browser constructs the data URL. */
  bytesBase64: string;
  licenseEvidence: FontLicenseEvidence;
}

export interface FontResolverOutcome {
  requestId: string;
  status: Exclude<FontResolutionStatus, "load_failed">;
  faceId?: string;
  requestedFamily?: string;
  resolvedFamily?: string;
  metricCompatible?: boolean;
  faceMatch?: FontFaceMatch;
  glyphCoverage?: FontGlyphCoverage;
  missingCodePoints?: readonly number[];
}

export interface FontResolverResponse {
  schemaVersion: typeof FONT_RESOLVER_SCHEMA_VERSION;
  resolverContract: typeof FONT_RESOLVER_CONTRACT_ID;
  substitutionContractVersion: typeof FONT_SUBSTITUTION_CONTRACT_VERSION;
  substitutionContractDigest: string;
  outcomes: readonly FontResolverOutcome[];
  faces: readonly FontResolverFace[];
}

/**
 * Trusted caller policy authority for face selection, license evidence, and
 * declared glyph coverage. Docxodus validates its schema, byte digest, and
 * browser loadability, but cannot independently prove those policy claims.
 */
export type FontResolver = (
  request: FontResolverRequest,
  signal: AbortSignal,
) => Promise<FontResolverResponse>;

/** Path- and byte-free evidence retained in render reports and fingerprints. */
export interface FontResolution {
  requestId: string;
  requestedFamily: string;
  requestedFamilies: readonly string[];
  requestedStyle: FontFaceStyle;
  requestedWeight: number;
  requestedStretch: number;
  sampleCodePointCount: number;
  sampleDigest: string;
  resolvedFamily?: string;
  resolvedFace?: string;
  status: FontResolutionStatus;
  source: FontResolutionSource;
  format?: FontFileFormat;
  fileSha256?: string;
  version?: string;
  faceMatch?: FontFaceMatch;
  metricCompatible?: boolean;
  glyphCoverage?: FontGlyphCoverage;
  missingCodePointCount?: number;
  /** Whether Chromium can paint some fallback despite an authoritative resolver miss. */
  browserFallbackAvailable?: boolean;
  licenseEvidence?: FontLicenseEvidence;
}

export interface FontConfigurationIdentity {
  resolverContract: typeof FONT_RESOLVER_CONTRACT_ID;
  substitutionContractVersion: typeof FONT_SUBSTITUTION_CONTRACT_VERSION;
  substitutionContractDigest: string;
  resolutionDigest: string;
}

import type { ConversionOptions } from "./types.js";

/**
 * The conversion options only the complete WASM entry point
 * (`DocumentConverter.ConvertDocxToHtmlComplete`) accepts. Setting any of them, even to its
 * default, routes a conversion there; the plain and pagination entry points would drop it.
 *
 * The main-thread `convertDocxToHtml` and the worker's both read this list, so a new
 * complete-only option is added here once rather than to two copies that can drift.
 */
export const COMPLETE_ONLY_CONVERSION_OPTIONS = [
  "renderFootnotesAndEndnotes",
  "renderHeadersAndFooters",
  "renderTrackedChanges",
  "showFormattingChanges",
  "showInsertionsAndDeletions",
  "showDeletedContent",
  "renderMoveOperations",
  "renderUnsupportedContentPlaceholders",
  "documentLanguage",
  "stampAnchors",
  "revisionPresentation",
  "semanticLists",
  "normalizeBulletMarkers",
  "stampPageNumberFields",
  "renderHeaderFooterRegistry",
] as const satisfies readonly (keyof ConversionOptions)[];

/** Appended WASM mask: zero keeps the default display, including for older callers. */
export function hiddenRevisionKinds(options: ConversionOptions | undefined): number {
  return (options?.showFormattingChanges === false ? 1 : 0)
    | (options?.showInsertionsAndDeletions === false ? 2 : 0);
}

/**
 * Whether a conversion must use the complete entry point: any complete-only option is set,
 * or annotations are rendered.
 */
export function needsCompleteConversion(options: ConversionOptions | undefined): boolean {
  if (!options) return false;
  return options.renderAnnotations === true ||
    COMPLETE_ONLY_CONVERSION_OPTIONS.some(key => options[key] !== undefined);
}

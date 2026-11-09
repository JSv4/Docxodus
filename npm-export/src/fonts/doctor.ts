import { DEFAULT_EXPORT_RESOURCE_LIMITS, FONT_SUBSTITUTION_CONTRACT } from "docxodus/export-browser";
import { createNodeFontRuntime } from "./resolver.js";
import type { LocalFontOptions } from "./local.js";

export const COMMON_WORD_FAMILIES = Object.freeze([
  "Aptos", ...FONT_SUBSTITUTION_CONTRACT.map(entry => entry.family), "Symbol", "Wingdings",
]);

/** Diagnose regular faces for PDF embedding using the same policy as conversion. */
export async function inspectCommonFonts(directories: readonly string[] = [], local: LocalFontOptions = {}) {
  const runtime = createNodeFontRuntime(directories, [], DEFAULT_EXPORT_RESOURCE_LIMITS, ["pdf"], local);
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(new Error("Local font inspection timed out.")), 30_000);
  try {
    const result = await runtime.resolver({ schemaVersion: 1, requests: COMMON_WORD_FAMILIES.map((family, index) => ({
      id: `doctor-${index}`, familyStack: [family], familyKinds: ["named"],
      style: "normal", weight: 400, stretch: 100, sampleCodePoints: [],
    })) }, controller.signal);
    return result.outcomes;
  } finally { clearTimeout(timer); }
}

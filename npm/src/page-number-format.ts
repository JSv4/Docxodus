/**
 * Rendering a page number in an OOXML number format — the browser-side counterpart of the .NET
 * `NumberFormats.Render`, used by the paginator to fill in `[data-field]` page-number markers.
 *
 * Two vocabularies reach this module and both are handled here so callers never have to know which
 * one they hold:
 *
 * - `ST_NumberFormat` tokens (`"lowerRoman"`, …), from a section's `w:pgNumType/@w:fmt` —
 *   stamped on the section wrapper as `data-page-num-fmt`.
 * - Field `\*` general-formatting switch arguments (`"roman"`, `"ROMAN"`, `"alphabetic"`,
 *   `"ALPHABETIC"`, `"Arabic"`), from a field's own instruction — stamped as `data-field-format`.
 *   Case is load-bearing: `roman` is `i, ii, iii` and `ROMAN` is `I, II, III`.
 *
 * An unrecognized token renders as decimal, matching how Word treats a format it does not
 * implement — a page number that reads "12" is wrong in style; one that reads "" is a hole.
 */

/** Formats this module renders. Anything else falls back to decimal. */
type Renderer = (value: number) => string;

const ROMAN_ONES = ["", "i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix"];
const ROMAN_TENS = ["", "x", "xx", "xxx", "xl", "l", "lx", "lxx", "lxxx", "xc"];
const ROMAN_HUNDREDS = ["", "c", "cc", "ccc", "cd", "d", "dc", "dcc", "dccc", "cm"];
const ROMAN_THOUSANDS = ["", "m", "mm", "mmm"];

function toRoman(value: number): string {
  // Word renders anything outside the classical range as a plain number rather than nothing.
  if (value <= 0 || value >= 4000) return String(value);
  return (
    ROMAN_THOUSANDS[Math.floor(value / 1000)] +
    ROMAN_HUNDREDS[Math.floor((value % 1000) / 100)] +
    ROMAN_TENS[Math.floor((value % 100) / 10)] +
    ROMAN_ONES[value % 10]
  );
}

/** `1 → a`, `26 → z`, `27 → aa` — Word repeats the letter rather than carrying like a base-26 number. */
function toLetter(value: number): string {
  if (value <= 0) return String(value);
  const wrapped = value % 780 === 0 ? 780 : value % 780;
  const repeats = Math.floor((wrapped - 1) / 26) + 1;
  return "abcdefghijklmnopqrstuvwxyz".charAt((wrapped - 1) % 26).repeat(repeats);
}

const RENDERERS: Record<string, Renderer> = {
  // ST_NumberFormat tokens (w:pgNumType/@w:fmt, w:numFmt).
  decimal: (v) => String(v),
  lowerRoman: toRoman,
  upperRoman: (v) => toRoman(v).toUpperCase(),
  lowerLetter: toLetter,
  upperLetter: (v) => toLetter(v).toUpperCase(),
  // Field `\*` general-formatting switch arguments.
  Arabic: (v) => String(v),
  roman: toRoman,
  ROMAN: (v) => toRoman(v).toUpperCase(),
  alphabetic: toLetter,
  ALPHABETIC: (v) => toLetter(v).toUpperCase(),
};

/**
 * Render `value` in `format`. `format` may be an `ST_NumberFormat` token or a `\*` switch argument;
 * an absent or unrecognized one renders as decimal.
 */
export function formatPageNumber(value: number, format?: string | null): string {
  const renderer = format ? RENDERERS[format] : undefined;
  return (renderer ?? RENDERERS.decimal)(value);
}

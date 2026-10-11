// Test inspection preserves PDF.js's explicit line boundaries independently of
// the ordinary searchableText representation used by the other export checks.
export function pdfLineText(items) {
  return items.filter((item) => "str" in item)
    .map((item) => item.str + (item.hasEOL ? "\n" : " ")).join("");
}

export function joinHyphenatedLines(text) {
  return text.replace(/-\n(?=\S)/g, "-");
}

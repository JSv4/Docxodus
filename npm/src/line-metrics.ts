/**
 * Unrounded natural line heights for the paginated export (issue #850).
 *
 * Chromium lays a `line-height: normal` line out at the font's natural line height — ascent +
 * descent + line gap from its metrics — but first rounds that to a whole CSS pixel (0.75 pt). The
 * converter relies on `normal` for single spacing and on `calc(1lh * m)` for Word's `auto`
 * multiples, so every line inherits the rounding: 11 pt Carlito advances 18 px instead of
 * 17.90 px, and the error accumulates down the page and moves page breaks. An explicit length keeps
 * Chromium's 1/64-pixel layout precision, so this pass replaces each `normal` with the same font's
 * natural height measured where rounding is negligible.
 */

/** Font size the natural line height is probed at: rounding to a pixel there is < 0.0005 em. */
const PROBE_FONT_SIZE_PX = 1000;

/** Probe text for an element whose own text the primary font covers (Latin, through U+024F). */
const LATIN_PROBE = "Hg";

/** At most this many distinct characters of an element's own text go into its probe. */
const MAX_PROBE_CHARACTERS = 256;

/**
 * What to measure an element's natural line height with. Under `normal`, Chromium also grows a line
 * for the fallback fonts its text needs (a CJK run in a Calibri paragraph lays out taller than Latin
 * text), so an element whose own text goes beyond Latin is probed with those characters themselves.
 */
function probeText(element: Element): string {
  const characters = new Set<string>();
  for (const node of Array.from(element.childNodes)) {
    if (node.nodeType !== 3) continue;
    for (const character of node.textContent ?? "") {
      if (/\s/.test(character)) continue;
      characters.add(character);
      if (characters.size >= MAX_PROBE_CHARACTERS) break;
    }
  }
  const beyondLatin = Array.from(characters).filter((character) => character.codePointAt(0)! > 0x024f);
  return beyondLatin.length === 0 ? LATIN_PROBE : LATIN_PROBE + beyondLatin.sort().join("");
}

/**
 * Give every element under `root` whose computed `line-height` is `normal` an explicit line height
 * equal to its own font's unrounded natural line height (including any fallback font its own text
 * needs). Elements are collected before any is
 * changed, because an explicit value on a parent would otherwise be inherited by children that
 * should keep sizing their lines from their own font. Returns how many elements changed.
 */
export function applyUnroundedNormalLineHeights(root: Element): number {
  const document = root.ownerDocument;
  const view = document.defaultView;
  if (!view) return 0;
  const ratios = new Map<string, number>();
  const ratioFor = (style: CSSStyleDeclaration, text: string): number => {
    const key = `${style.fontStyle}|${style.fontWeight}|${style.fontStretch}|${style.fontFamily}|${text}`;
    let ratio = ratios.get(key);
    if (ratio === undefined) {
      const probe = document.createElement("span");
      probe.setAttribute("aria-hidden", "true");
      probe.style.cssText = "position:absolute;left:-100000px;top:0;visibility:hidden;"
        + "display:inline-block;white-space:nowrap;line-height:normal;margin:0;padding:0;border:0;";
      probe.style.fontFamily = style.fontFamily;
      probe.style.fontStyle = style.fontStyle;
      probe.style.fontWeight = style.fontWeight;
      probe.style.fontStretch = style.fontStretch;
      probe.style.fontSize = `${PROBE_FONT_SIZE_PX}px`;
      probe.textContent = text;
      (document.body ?? document.documentElement).appendChild(probe);
      ratio = probe.getBoundingClientRect().height / PROBE_FONT_SIZE_PX;
      probe.remove();
      ratios.set(key, ratio);
    }
    return ratio;
  };

  const targets: Array<{ element: HTMLElement; lineHeight: number }> = [];
  for (const element of [root, ...Array.from(root.querySelectorAll("*"))]) {
    if (!(element instanceof view.HTMLElement)) continue;
    const style = view.getComputedStyle(element);
    if (style.lineHeight !== "normal") continue;
    const fontSize = Number.parseFloat(style.fontSize);
    const ratio = ratioFor(style, probeText(element));
    if (!(fontSize > 0) || !(ratio > 0)) continue;
    targets.push({ element, lineHeight: ratio * fontSize });
  }
  for (const { element, lineHeight } of targets) {
    element.style.setProperty("line-height", `${lineHeight.toFixed(4)}px`);
  }
  return targets.length;
}

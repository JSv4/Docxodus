/**
 * In-place image patching for the editor's incremental reconcile.
 *
 * When a block's fresh render differs from its live DOM node only in the attributes
 * of `<img>` elements — a host replacing an image's media through the session, most
 * visibly the arcade's Doom cartridge writing a new framebuffer PNG every frame —
 * the reconciler must NOT swap the node. Firefox and WebKit paint a freshly inserted
 * `<img>` as an empty box until its (data-URI) source is decoded, so a per-frame swap
 * strobes white between frames. Changing `src` on the element that is already on
 * screen keeps the previous bitmap visible until the new one is ready, in every
 * engine.
 *
 * Both functions are deliberately self-contained (no module-level state, no imports)
 * so a browser test can inject them into a page with `Function.prototype.toString`.
 */

/** Attributes the editor stamps on live blocks after rendering; a fresh render never
 *  carries them, so they are not a difference. Mirrors `DocxEditor.wireBlock`. Each
 *  function repeats the list inline rather than defaulting to this constant so that
 *  its source, injected into a page on its own, has no free variable. */
export const EDITOR_OWNED_ATTRIBUTES: readonly string[] = [
  "contenteditable", "data-committed-text", "data-render-sig", "data-note-anchor",
];

/**
 * The `[live, fresh]` `<img>` pairs when `live` and `fresh` are the same tree apart
 * from `<img>` attributes; `null` when anything else differs (tag, text, any other
 * attribute, child count), in which case the caller must swap nodes as before.
 *
 * Editor-owned attributes on the live side are ignored, as is the list-marker
 * separator element the editor injects into list blocks. Everything else must match
 * exactly — the patch is only ever taken when it is provably equivalent to a swap.
 */
export function imageOnlyDelta(
  live: Element,
  fresh: Element,
  ignoredAttributes?: readonly string[],
): Array<[HTMLImageElement, HTMLImageElement]> | null {
  const ignored = new Set(
    ignoredAttributes ?? ["contenteditable", "data-committed-text", "data-render-sig", "data-note-anchor"],
  );
  const pairs: Array<[HTMLImageElement, HTMLImageElement]> = [];
  const attributeNames = (el: Element): string[] =>
    Array.from(el.attributes)
      .map((a) => a.name)
      .filter((n) => !ignored.has(n))
      .sort();
  const sameAttributes = (a: Element, b: Element): boolean => {
    const na = attributeNames(a);
    const nb = attributeNames(b);
    if (na.length !== nb.length) return false;
    for (let i = 0; i < na.length; i++) {
      if (na[i] !== nb[i] || a.getAttribute(na[i]) !== b.getAttribute(nb[i])) return false;
    }
    return true;
  };
  const liveChildren = (el: Element): ChildNode[] =>
    Array.from(el.childNodes).filter(
      (n) => !(n.nodeType === 1 && (n as Element).hasAttribute("data-editor-list-separator")),
    );
  const walk = (a: Node, b: Node): boolean => {
    if (a.nodeType !== b.nodeType) return false;
    if (a.nodeType === 3) return a.nodeValue === b.nodeValue; // text
    if (a.nodeType !== 1) return true; // comments and the like carry no rendering
    const ea = a as Element;
    const eb = b as Element;
    if (ea.tagName !== eb.tagName) return false;
    if (ea.tagName === "IMG") {
      pairs.push([ea as HTMLImageElement, eb as HTMLImageElement]);
      return true;
    }
    if (!sameAttributes(ea, eb)) return false;
    const ca = liveChildren(ea);
    const cb = Array.from(eb.childNodes);
    if (ca.length !== cb.length) return false;
    for (let i = 0; i < ca.length; i++) if (!walk(ca[i], cb[i])) return false;
    return true;
  };
  return walk(live, fresh) ? pairs : null;
}

/**
 * Make `live` carry exactly `fresh`'s attributes, touching only what changed and
 * keeping the editor-owned stamps. `src` is set last, after any new dimensions, so
 * the box is already the right size when the new source starts loading. An unchanged
 * `src` is left alone — re-setting it would restart the load for nothing.
 */
export function patchImageAttributes(
  live: HTMLImageElement,
  fresh: HTMLImageElement,
  ignoredAttributes?: readonly string[],
): void {
  const ignored = new Set(
    ignoredAttributes ?? ["contenteditable", "data-committed-text", "data-render-sig", "data-note-anchor"],
  );
  for (const attr of Array.from(live.attributes)) {
    if (!ignored.has(attr.name) && !fresh.hasAttribute(attr.name)) live.removeAttribute(attr.name);
  }
  for (const attr of Array.from(fresh.attributes)) {
    if (attr.name === "src") continue;
    if (live.getAttribute(attr.name) !== attr.value) live.setAttribute(attr.name, attr.value);
  }
  const src = fresh.getAttribute("src");
  if (src !== null && live.getAttribute("src") !== src) live.setAttribute("src", src);
}

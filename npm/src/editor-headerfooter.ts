/**
 * HeaderFooterRegion — the DocxEditor's docked header/footer editing bands.
 *
 * Header and footer stories live in their own OOXML parts (HeaderPart/FooterPart) *outside*
 * the body, so they cannot be another block in the body flow. This module renders them as two
 * docked bands — one above the body, one below — each composed PER STORY PARAGRAPH via the
 * session-attached `RenderBlockHtml`, which resolves `hdr`/`ftr` anchors natively.
 *
 * Using the same renderer for first paint and for the post-edit incremental swap means there is
 * no fidelity drift between them. It is also the only option that works in continuous mode at
 * all: the editor's full-document render only asks for headers/footers in paginated mode, and it
 * never stamps `data-anchor` inside header/footer parts (Unids are assigned to the main document
 * part only). In paginated mode the page boxes additionally clone one header node onto every
 * page, so page-margin nodes could never be uniquely addressable — bands keep exactly one DOM
 * node per story paragraph.
 *
 * Once rendered, a story paragraph is an ORDINARY editable block: DocxSession's anchor
 * resolution indexes every scope and routes by part, so ReplaceText/ApplyFormat/SetParagraphFormat
 * all accept a `p:hdr1:<unid>` anchor. The editor wires band blocks with the same `wireBlock` it
 * uses for the body, which is why the whole ribbon works inside a band with no new command code.
 */

import type { HeaderFooterKind } from "./types.js";

/** Which of the two bands an operation targets. */
export type BandWhich = "header" | "footer";

/** The bridge slice the region needs. Structurally satisfied by `DocxEditorExports.DocxSessionBridge`;
 *  declared locally so this module and `editor.ts` don't import each other. */
export interface HeaderFooterBridge {
  Project: (handle: number) => string;
  GetSectionInfo: (handle: number, anchorId: string) => string;
  SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  InsertPageNumberField: (handle: number, anchor: string, field: string) => string;
  EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
  RenderBlockHtml: (
    handle: number,
    anchorId: string,
    cssPrefix: string,
    fabricateClasses: boolean,
  ) => string;
}

export interface HeaderFooterRegionCallbacks {
  /** Wire a freshly rendered story paragraph as an editable block (DocxEditor.wireBlock). */
  wireBlock: (el: HTMLElement) => void;
  /** Rebuild the editor's unid → full-anchor map (a seed adds a part, changing the projection). */
  refreshAnchorMap: () => void;
}

export interface HeaderFooterRegionOptions {
  cssPrefix: string;
  fabricateClasses: boolean;
}

interface HeaderFooterRefLite {
  kind: HeaderFooterKind;
  partUri: string;
}

interface SectionInfoLite {
  sectionUnid: string;
  headerRefs?: HeaderFooterRefLite[];
  footerRefs?: HeaderFooterRefLite[];
}

interface AnchorEntryLite {
  partUri: string;
  unid: string;
  kind: string;
  scope: string;
}

const KIND_LABELS: Array<{ value: HeaderFooterKind; label: string }> = [
  { value: "default", label: "Default" },
  { value: "first", label: "First page" },
  { value: "even", label: "Even pages" },
];

/**
 * Verbatim from docs/architecture/docx_mutation_api.md. `w:evenAndOddHeaders` is set on the
 * SETTINGS part, so it is document-global and governs footers as well as headers: once set,
 * even pages stop inheriting the Default stories entirely. A demo user who creates only an
 * even header hits this immediately and sees their footer vanish on page 2.
 */
const EVEN_WARNING =
  "w:evenAndOddHeaders is document-global and governs footers too — even pages stop " +
  "inheriting the Default stories, so a section with only a Default footer shows no footer " +
  "at all on even pages.";

/**
 * `w:titlePg` is per-section but has the same shape of surprise: once set, page 1 uses its OWN
 * header AND footer, so an empty first-page footer means page 1 shows no footer even though the
 * Default one is populated.
 */
const FIRST_WARNING =
  "w:titlePg makes page 1 use its own first-page header AND footer instead of the Default " +
  "ones, so an empty first-page footer leaves page 1 with no footer.";

/** Kinds/scopes the markdown projection addresses as editable text blocks. */
const STORY_BLOCK_KINDS = new Set(["p", "h", "li"]);

export class HeaderFooterRegion {
  readonly headerBand: HTMLElement;
  readonly footerBand: HTMLElement;

  private readonly bridge: HeaderFooterBridge;
  private readonly handle: number;
  private readonly options: HeaderFooterRegionOptions;
  private readonly callbacks: HeaderFooterRegionCallbacks;

  /** The body anchor whose section the bands currently describe (needed to seed a story). */
  private bodyAnchorId: string | null = null;
  private sectionInfo: SectionInfoLite | null = null;
  /** Memoized body anchor → governing sectionUnid, so repeat focus costs no bridge call. */
  private readonly sectionOfAnchor = new Map<string, string>();

  private readonly kinds: Record<BandWhich, HeaderFooterKind> = {
    header: "default",
    footer: "default",
  };

  constructor(
    bridge: HeaderFooterBridge,
    handle: number,
    options: HeaderFooterRegionOptions,
    callbacks: HeaderFooterRegionCallbacks,
  ) {
    this.bridge = bridge;
    this.handle = handle;
    this.options = options;
    this.callbacks = callbacks;
    this.headerBand = this.buildBand("header");
    this.footerBand = this.buildBand("footer");
  }

  // ─── public surface ──────────────────────────────────────────────────

  /**
   * Point the bands at the section governing `bodyAnchorId` and repaint if it changed.
   * Cheap to call on every focus change: the anchor → section lookup is memoized, and an
   * unchanged `sectionUnid` returns without touching the DOM (so it never clobbers the
   * user's kind selection).
   */
  syncToBody(bodyAnchorId: string | null): void {
    if (!bodyAnchorId) return;
    const cached = this.sectionOfAnchor.get(bodyAnchorId);
    if (cached !== undefined && cached === this.sectionInfo?.sectionUnid) {
      this.bodyAnchorId = bodyAnchorId; // same section — keep it as the seed target
      return;
    }
    const info = this.readSectionInfo(bodyAnchorId);
    if (!info) return; // not a body anchor (or no sectPr) — leave the bands as they are
    this.sectionOfAnchor.set(bodyAnchorId, info.sectionUnid);
    const changed = info.sectionUnid !== this.sectionInfo?.sectionUnid;
    this.bodyAnchorId = bodyAnchorId;
    this.sectionInfo = info;
    if (changed) this.refreshAll();
  }

  /** Repaint one band from the live session. */
  refresh(which: BandWhich): void {
    this.sectionInfo = this.bodyAnchorId ? this.readSectionInfo(this.bodyAnchorId) : null;
    this.renderBand(which);
  }

  /** Repaint both bands from the live session (after a remount, undo/redo, or section change). */
  refreshAll(): void {
    this.renderBand("header");
    this.renderBand("footer");
  }

  /** Select (and, if absent, create) the story kind a band edits, and make it render. */
  setKind(which: BandWhich, kind: HeaderFooterKind): void {
    this.kinds[which] = kind;
    const select = this.bandFor(which).querySelector<HTMLSelectElement>("[data-hf-kind]");
    if (select && select.value !== kind) select.value = kind;

    if (!this.partUriFor(which, kind)) this.seedStory(which, kind);
    // Selecting first/even IS the user saying "use a different first/even page", so the section's
    // visibility flag must be set even when the part already exists — seedStory (which sets it as
    // a side effect of writing content) doesn't run then. Word leaves first/even parts behind when
    // those options are switched off, so a real document commonly has the reference WITHOUT the
    // flag; without this the typed story is saved but never rendered.
    if (this.bodyAnchorId && kind !== "default") {
      this.bridge.EnsureHeaderFooterVisible(this.handle, this.bodyAnchorId, kind);
    }
    this.refresh(which);
  }

  /** The kind a band is currently editing. */
  kindOf(which: BandWhich): HeaderFooterKind {
    return this.kinds[which];
  }

  /** Append a PAGE / NUMPAGES field to the story paragraph addressed by `anchorId`. */
  insertPageNumber(which: BandWhich, anchorId: string, field: "currentPage" | "totalPages"): boolean {
    const res = parseResult(this.bridge.InsertPageNumberField(this.handle, anchorId, field));
    if (!res.success) return false;
    this.refresh(which);
    return true;
  }

  /** Insert a page number into a band's own target paragraph (focused, else last). Seeds the
   *  story first if the band's selected kind has none, so the command is never a silent no-op. */
  insertPageNumberInBand(which: BandWhich, field: "currentPage" | "totalPages"): boolean {
    if (!this.partUriFor(which, this.kinds[which])) {
      this.seedStory(which, this.kinds[which]);
      this.renderBand(which);
    }
    const target = this.pageNumberTarget(which);
    return target ? this.insertPageNumber(which, target, field) : false;
  }

  /** The band element containing `node`, or null. */
  bandOf(node: Node | null): HTMLElement | null {
    const el = node && node.nodeType === 1 ? (node as HTMLElement) : node?.parentElement ?? null;
    const band = el?.closest<HTMLElement>("[data-hf-band]") ?? null;
    return band === this.headerBand || band === this.footerBand ? band : null;
  }

  /** The block-list root (a band's story container) owning `node`, or null. */
  blockRootOf(node: Node | null): HTMLElement | null {
    const band = this.bandOf(node);
    return band ? band.querySelector<HTMLElement>("[data-hf-body]") : null;
  }

  /** True when `node` is inside either band (including its chrome). */
  contains(node: Node | null): boolean {
    return this.bandOf(node) !== null;
  }

  /** `"header"` / `"footer"` for a band element produced by this region. */
  whichOf(band: HTMLElement): BandWhich {
    return band === this.footerBand ? "footer" : "header";
  }

  // ─── section + anchor resolution ─────────────────────────────────────

  private readSectionInfo(bodyAnchorId: string): SectionInfoLite | null {
    try {
      const raw = this.bridge.GetSectionInfo(this.handle, bodyAnchorId);
      const parsed = JSON.parse(raw) as SectionInfoLite | null;
      return parsed && typeof parsed.sectionUnid === "string" ? parsed : null;
    } catch {
      return null;
    }
  }

  private refsFor(which: BandWhich): HeaderFooterRefLite[] {
    const info = this.sectionInfo;
    if (!info) return [];
    return (which === "header" ? info.headerRefs : info.footerRefs) ?? [];
  }

  /** URI of the part supplying `kind` for this band, or null when the story doesn't exist. */
  private partUriFor(which: BandWhich, kind: HeaderFooterKind): string | null {
    return this.refsFor(which).find((r) => r.kind === kind)?.partUri ?? null;
  }

  /** Full anchor ids of the story paragraphs held in `partUri`, in document order. */
  private storyAnchors(partUri: string): string[] {
    let index: Record<string, AnchorEntryLite>;
    try {
      index = (JSON.parse(this.bridge.Project(this.handle)) as {
        anchorIndex: Record<string, AnchorEntryLite>;
      }).anchorIndex;
    } catch {
      return [];
    }
    return Object.entries(index)
      .filter(([, t]) => t.partUri === partUri && STORY_BLOCK_KINDS.has(t.kind))
      .map(([id]) => id);
  }

  /**
   * Create an empty story for `kind`. Seeding happens the moment a kind is selected rather than
   * lazily on first keystroke: an absent story renders no contenteditable element, so there
   * would be nothing to click into.
   */
  private seedStory(which: BandWhich, kind: HeaderFooterKind): void {
    if (!this.bodyAnchorId) return;
    const set = which === "header" ? this.bridge.SetHeaderText : this.bridge.SetFooterText;
    const res = parseResult(set(this.handle, this.bodyAnchorId, kind, ""));
    if (!res.success) return;
    // A new part changes the projection (fresh hdr{N}/ftr{N} scope) and the section's refs.
    this.callbacks.refreshAnchorMap();
    this.sectionInfo = this.readSectionInfo(this.bodyAnchorId);
  }

  // ─── rendering ───────────────────────────────────────────────────────

  private bandFor(which: BandWhich): HTMLElement {
    return which === "header" ? this.headerBand : this.footerBand;
  }

  private buildBand(which: BandWhich): HTMLElement {
    const band = document.createElement("div");
    band.className = "docx-hf-band";
    band.setAttribute("data-hf-band", which);

    const chrome = document.createElement("div");
    chrome.className = "docx-hf-chrome";

    const label = document.createElement("span");
    label.className = "docx-hf-label";
    label.textContent = which === "header" ? "Header" : "Footer";
    chrome.appendChild(label);

    const kindSelect = document.createElement("select");
    kindSelect.setAttribute("data-hf-kind", "");
    kindSelect.title = "Which header/footer story to edit";
    for (const k of KIND_LABELS) {
      const opt = document.createElement("option");
      opt.value = k.value;
      opt.textContent = k.label;
      kindSelect.appendChild(opt);
    }
    kindSelect.addEventListener("change", () =>
      this.setKind(which, kindSelect.value as HeaderFooterKind),
    );
    chrome.appendChild(kindSelect);

    const pageNum = document.createElement("select");
    pageNum.setAttribute("data-hf-pagenum", "");
    pageNum.title = "Insert a page-number field into the focused story paragraph";
    for (const o of [
      { value: "", label: "Page №…" },
      { value: "currentPage", label: "Page number" },
      { value: "totalPages", label: "Total pages" },
    ]) {
      const opt = document.createElement("option");
      opt.value = o.value;
      opt.textContent = o.label;
      pageNum.appendChild(opt);
    }
    // The <select> steals focus from the story paragraph, collapsing the caret, so the target
    // is resolved from the band's own blocks rather than from the editor's live selection.
    pageNum.addEventListener("change", () => {
      const field = pageNum.value;
      pageNum.value = "";
      if (field !== "currentPage" && field !== "totalPages") return;
      const target = this.pageNumberTarget(which);
      if (target) this.insertPageNumber(which, target, field);
    });
    chrome.appendChild(pageNum);

    band.appendChild(chrome);

    const warning = document.createElement("div");
    warning.className = "docx-hf-warning";
    warning.setAttribute("data-hf-warning", "");
    warning.hidden = true;
    band.appendChild(warning);

    const body = document.createElement("div");
    body.className = "docx-hf-body";
    body.setAttribute("data-hf-body", "");
    band.appendChild(body);

    return band;
  }

  /**
   * Which story paragraph a page-number insert targets: the last one the user focused in this
   * band if it is still attached, else the band's last paragraph (Word's convention — the page
   * number goes at the end of the footer line).
   */
  private pageNumberTarget(which: BandWhich): string | null {
    const body = this.bandFor(which).querySelector<HTMLElement>("[data-hf-body]");
    if (!body) return null;
    const focused = body.querySelector<HTMLElement>('[data-hf-focused="true"][data-hf-anchor]');
    const blocks = Array.from(body.querySelectorAll<HTMLElement>("[data-hf-anchor]"));
    const target = focused ?? blocks[blocks.length - 1];
    return target?.getAttribute("data-hf-anchor") ?? null;
  }

  private renderBand(which: BandWhich): void {
    const band = this.bandFor(which);
    const body = band.querySelector<HTMLElement>("[data-hf-body]");
    if (!body) return;

    const kind = this.kinds[which];
    const select = band.querySelector<HTMLSelectElement>("[data-hf-kind]");
    if (select && select.value !== kind) select.value = kind;

    this.renderKindWarning(which);

    body.innerHTML = "";
    const partUri = this.partUriFor(which, kind);
    if (!partUri) {
      body.appendChild(this.placeholder(which, kind));
      band.setAttribute("data-hf-empty", "true");
      return;
    }
    band.removeAttribute("data-hf-empty");

    const anchors = this.storyAnchors(partUri);
    if (anchors.length === 0) {
      body.appendChild(this.placeholder(which, kind));
      return;
    }
    for (const anchorId of anchors) {
      const el = this.renderStoryBlock(anchorId);
      if (!el) continue;
      body.appendChild(el);
      // Adopt BEFORE wiring: the editor resolves a block's anchor from `data-hf-anchor` first,
      // because a story paragraph's content-addressed unid can collide with another part's.
      this.adoptBlock(el, anchorId);
      this.callbacks.wireBlock(el);
    }
  }

  /**
   * Mark a story paragraph as belonging to this region. The full anchor id is stamped alongside
   * `data-anchor` (which carries only the bare unid, matching the body's convention) so band
   * chrome can address the block without going through the editor's unid map. Called on first
   * render AND after an incremental swap, which replaces the DOM node.
   */
  adoptBlock(el: HTMLElement, anchorId: string): void {
    const body = this.blockRootOf(el);
    if (!body) return;
    el.setAttribute("data-hf-anchor", anchorId);
    if (el.dataset.hfAdopted === "true") return;
    el.dataset.hfAdopted = "true";
    el.addEventListener("focus", () => {
      body
        .querySelectorAll("[data-hf-focused]")
        .forEach((sib) => sib.removeAttribute("data-hf-focused"));
      el.setAttribute("data-hf-focused", "true");
    });
  }

  private renderStoryBlock(anchorId: string): HTMLElement | null {
    const html = this.bridge.RenderBlockHtml(
      this.handle,
      anchorId,
      this.options.cssPrefix,
      this.options.fabricateClasses,
    );
    if (html.charCodeAt(0) === 0x7b /* error object */) return null;
    return new DOMParser().parseFromString(html, "text/html").body
      .firstElementChild as HTMLElement | null;
  }

  private placeholder(which: BandWhich, kind: HeaderFooterKind): HTMLElement {
    const el = document.createElement("div");
    el.className = "docx-hf-placeholder";
    el.setAttribute("data-hf-placeholder", "");
    const label = KIND_LABELS.find((k) => k.value === kind)?.label ?? kind;
    el.textContent = `No ${label.toLowerCase()} ${which} — pick a kind to create one.`;
    return el;
  }

  /**
   * Surface the caveat that comes with the selected kind. Turning on first/even means those pages
   * stop using the Default stories entirely, which bites hardest on the OTHER band: enabling an
   * even header with no even footer leaves even pages footer-less, and enabling a first-page
   * header with an empty first-page footer leaves page 1 footer-less. The note is shown whenever
   * first/even is selected (the behavior change is real either way); the fix button appears only
   * when the counterpart story is missing entirely, which is what a user almost always wants next.
   */
  private renderKindWarning(which: BandWhich): void {
    const band = this.bandFor(which);
    const warning = band.querySelector<HTMLElement>("[data-hf-warning]");
    if (!warning) return;

    const kind = this.kinds[which];
    if (kind === "default") {
      warning.hidden = true;
      warning.textContent = "";
      return;
    }
    const other: BandWhich = which === "header" ? "footer" : "header";
    warning.hidden = false;
    warning.textContent = `${kind === "even" ? EVEN_WARNING : FIRST_WARNING} `;
    if (this.partUriFor(other, kind)) return; // counterpart exists — nothing to offer

    const fix = document.createElement("button");
    fix.type = "button";
    fix.setAttribute("data-hf-fix-even-footer", "");
    fix.textContent = `Also create a ${kind === "even" ? "matching even" : "first-page"} ${other}`;
    fix.addEventListener("click", () => {
      this.setKind(other, kind);
      this.renderKindWarning(which);
    });
    warning.appendChild(fix);
  }
}

function parseResult(json: string): { success: boolean } {
  try {
    return JSON.parse(json) as { success: boolean };
  } catch {
    return { success: false };
  }
}

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

import type { HeaderFooterKind, NumberFormat } from "./types.js";

/** Which of the two bands an operation targets. */
export type BandWhich = "header" | "footer";

/** The bridge slice the region needs. Structurally satisfied by `DocxEditorExports.DocxSessionBridge`;
 *  declared locally so this module and `editor.ts` don't import each other. */
export interface HeaderFooterBridge {
  Project: (handle: number) => string;
  GetSectionInfo: (handle: number, anchorId: string) => string;
  SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  InsertPageNumberField: (handle: number, anchor: string, field: string, format: string) => string;
  EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
  SetPageNumbering: (handle: number, anchor: string, opJson: string) => string;
  ClearPageNumbering: (handle: number, anchor: string) => string;
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
  /** The story comes from an earlier section (this one declares no reference of that kind). */
  inherited?: boolean;
}

interface SectionInfoLite {
  sectionUnid: string;
  headerRefs?: HeaderFooterRefLite[];
  footerRefs?: HeaderFooterRefLite[];
  /** `w:pgNumType/@w:start` — absent when the section continues the previous one's numbering. */
  pageNumberStart?: number;
  /** `w:pgNumType/@w:fmt` — absent means Word's default `1, 2, 3`. */
  pageNumberFormat?: NumberFormat;
}

/** The page-number formats Word's *Format Page Numbers…* dialog offers, in its order. The blank
 *  entry means "leave the section's format alone" — not "decimal", which would write an attribute
 *  the document may never have had. */
const PAGE_FORMAT_LABELS: ReadonlyArray<{ value: string; label: string }> = [
  { value: "", label: "Format…" },
  { value: "decimal", label: "1, 2, 3" },
  { value: "lowerLetter", label: "a, b, c" },
  { value: "upperLetter", label: "A, B, C" },
  { value: "lowerRoman", label: "i, ii, iii" },
  { value: "upperRoman", label: "I, II, III" },
];

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

  /** Append a PAGE / NUMPAGES field to the story paragraph addressed by `anchorId`.
   *  Deliberately a PLAIN field (no `\*` switch), exactly as Word inserts one, so it follows the
   *  section's page-number format — see {@link setPageNumbering}. Stamping the section's current
   *  format as a switch here would silently WIN over any later format change. */
  insertPageNumber(which: BandWhich, anchorId: string, field: "currentPage" | "totalPages"): boolean {
    const res = parseResult(this.bridge.InsertPageNumberField(this.handle, anchorId, field, ""));
    if (!res.success) return false;
    this.refresh(which);
    return true;
  }

  /**
   * Set this section's page numbering (`w:pgNumType`) — Word's *Format Page Numbers…*. Omitted
   * fields are left alone, so the format and the start are independently settable. Both bands then
   * repaint, because the values belong to the section rather than to either story.
   *
   * The rendered page numbers in the editor do not change: a page-number field's cached result is
   * what the browser shows, and Word recomputes it on open. Paginated mode substitutes the real
   * per-page number, so it does reflect the format immediately.
   */
  setPageNumbering(op: { start?: number; format?: NumberFormat }): boolean {
    if (!this.bodyAnchorId) return false;
    const res = parseResult(
      this.bridge.SetPageNumbering(this.handle, this.bodyAnchorId, JSON.stringify(op)),
    );
    if (!res.success) return false;
    this.reloadSection();
    return true;
  }

  /** This section's page numbering as the live document states it. Fields are absent, not
   *  defaulted — "continues the previous section" is not the same claim as "starts at 1". */
  pageNumbering(): { start?: number; format?: NumberFormat } {
    return {
      start: this.sectionInfo?.pageNumberStart,
      format: this.sectionInfo?.pageNumberFormat,
    };
  }

  /** Remove this section's page-numbering start/format — it reverts to continuing the previous
   *  section's numbering in Word's default `1, 2, 3`. */
  clearPageNumbering(): boolean {
    if (!this.bodyAnchorId) return false;
    const res = parseResult(this.bridge.ClearPageNumbering(this.handle, this.bodyAnchorId));
    if (!res.success) return false;
    this.reloadSection();
    return true;
  }

  /**
   * Re-read the section from the live document and repaint BOTH bands.
   *
   * `refreshAll` only repaints — it deliberately does not re-read, because its callers (remount,
   * undo/redo) already refreshed the section. A section-property write has no such caller, and
   * repainting from the stale snapshot would leave the chrome reporting the value the document had
   * before the edit.
   */
  private reloadSection(): void {
    this.sectionInfo = this.bodyAnchorId ? this.readSectionInfo(this.bodyAnchorId) : null;
    this.refreshAll();
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

  /** The reference supplying `kind` for this band — possibly inherited from an earlier section. */
  private refFor(which: BandWhich, kind: HeaderFooterKind): HeaderFooterRefLite | null {
    return this.refsFor(which).find((r) => r.kind === kind) ?? null;
  }

  /** URI of the part supplying `kind` for this band, or null when the story doesn't exist. */
  private partUriFor(which: BandWhich, kind: HeaderFooterKind): string | null {
    return this.refFor(which, kind)?.partUri ?? null;
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

    // Word's "Format Page Numbers…": these are SECTION properties (w:pgNumType), not properties of
    // this band's story, so both bands show the same values and either can set them. An inserted
    // page-number field is plain, so it renders through whatever is chosen here.
    const pageFmt = document.createElement("select");
    pageFmt.setAttribute("data-hf-pagefmt", "");
    pageFmt.title = "Page-number format for this section";
    for (const o of PAGE_FORMAT_LABELS) {
      const opt = document.createElement("option");
      opt.value = o.value;
      opt.textContent = o.label;
      pageFmt.appendChild(opt);
    }
    pageFmt.addEventListener("change", () => {
      if (pageFmt.value === "") return;
      this.setPageNumbering({ format: pageFmt.value as NumberFormat });
    });
    chrome.appendChild(pageFmt);

    const pageStart = document.createElement("input");
    pageStart.setAttribute("data-hf-pagestart", "");
    pageStart.type = "number";
    pageStart.min = "0";
    pageStart.placeholder = "Start at";
    pageStart.title = "Restart this section's page numbering at this number";
    // Commit on blur/Enter, not on every keystroke: typing "12" would otherwise apply a start of 1
    // first, and each apply is an undo step.
    const commitStart = () => {
      const raw = pageStart.value.trim();
      if (raw === "") return;
      const start = Number(raw);
      if (!Number.isInteger(start) || start < 0) return;
      this.setPageNumbering({ start });
    };
    pageStart.addEventListener("blur", commitStart);
    pageStart.addEventListener("keydown", (e) => {
      if ((e as KeyboardEvent).key === "Enter") {
        e.preventDefault();
        commitStart();
      }
    });
    chrome.appendChild(pageStart);

    const inheritedNote = document.createElement("span");
    inheritedNote.className = "docx-hf-inherited";
    inheritedNote.setAttribute("data-hf-inherited-note", "");
    inheritedNote.hidden = true;
    chrome.appendChild(inheritedNote);

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

    // Section-level page numbering — the same values on both bands, seeded from the live sectPr so
    // the chrome reports what the document says rather than what was last clicked.
    const pageFmt = band.querySelector<HTMLSelectElement>("[data-hf-pagefmt]");
    if (pageFmt) pageFmt.value = this.sectionInfo?.pageNumberFormat ?? "";
    const pageStart = band.querySelector<HTMLInputElement>("[data-hf-pagestart]");
    if (pageStart && document.activeElement !== pageStart) {
      const start = this.sectionInfo?.pageNumberStart;
      pageStart.value = start === undefined ? "" : String(start);
    }

    // A section that declares no reference of this kind shows the story it inherits from an
    // earlier section. Say so: editing it changes BOTH sections, because they share one part.
    const inherited = this.refFor(which, kind)?.inherited === true;
    band.toggleAttribute("data-hf-inherited", inherited);
    const note = band.querySelector<HTMLElement>("[data-hf-inherited-note]");
    if (note) {
      note.hidden = !inherited;
      note.textContent = inherited ? "inherited from an earlier section" : "";
    }

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
    fix.setAttribute("data-hf-fix-counterpart", "");
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

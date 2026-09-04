/**
 * HeaderFooterRegion — the DocxEditor's header/footer editing surface.
 *
 * Header and footer stories live in their own OOXML parts (HeaderPart/FooterPart) *outside*
 * the body, so they cannot be another block in the body flow. This module presents them two
 * ways, both composed PER STORY PARAGRAPH via the editor's block renderer, which resolves
 * `hdr`/`ftr` anchors natively:
 *
 *  - **Continuous view** — two bands, one docked above the body and one below, drawn as the
 *    top and bottom margin of the sheet (a dashed rule and a small "Header"/"Footer" tag, the
 *    look Word has while a header is being edited).
 *  - **Page view** — no bands. The paginator clones each story onto every page as inert
 *    presentation; clicking a page's header or footer area swaps THAT page's clone for the
 *    real, editable story paragraphs. A commit re-renders the paragraph and re-clones the story
 *    onto every other page that shows it, so all pages update live, and page-number fields on
 *    each clone are substituted per page just as the paginator does.
 *
 * Once rendered, a story paragraph is an ORDINARY editable block: DocxSession's anchor
 * resolution indexes every scope and routes by part, so ReplaceText/ApplyFormat/SetParagraphFormat
 * all accept a `p:hdr1:<unid>` anchor. The editor wires story blocks with the same `wireBlock` it
 * uses for the body, which is why the whole ribbon works inside a story with no new command code.
 *
 * "Different first page" and "Different odd & even pages" are Word's two checkboxes, backed by
 * `w:titlePg` (per section) and `w:evenAndOddHeaders` (document-global). Enabling one seeds BOTH
 * the header and the footer story of that kind, exactly as Word does — page 1 stops using the
 * default stories the moment `w:titlePg` is set, so seeding only one side would silently leave
 * page 1 with no footer.
 */

import { formatPageNumber } from "./page-number-format.js";
import type { HeaderFooterKind, NumberFormat } from "./types.js";

/** Which of the two stories an operation targets. */
export type BandWhich = "header" | "footer";

/** The bridge slice the region needs. Structurally satisfied by `DocxEditorExports.DocxSessionBridge`;
 *  declared locally so this module and `editor.ts` don't import each other. */
export interface HeaderFooterBridge {
  Project: (handle: number) => string;
  ListAnchors?: (handle: number) => string;
  GetSectionInfo: (handle: number, anchorId: string) => string;
  SetHeaderText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  SetFooterText: (handle: number, anchor: string, kind: string, markdown: string) => string;
  InsertPageNumberField: (handle: number, anchor: string, field: string, format: string) => string;
  EnsureHeaderFooterVisible: (handle: number, anchor: string, kind: string) => string;
  /** Turn `w:titlePg` / `w:evenAndOddHeaders` on or off (optional: older bundles can only enable). */
  SetHeaderFooterKindEnabled?: (handle: number, anchor: string, kind: string, enabled: boolean) => string;
  SetPageNumbering: (handle: number, anchor: string, opJson: string) => string;
  ClearPageNumbering: (handle: number, anchor: string) => string;
}

export interface HeaderFooterRegionCallbacks {
  /** Wire a freshly rendered story paragraph as an editable block (DocxEditor.wireBlock). */
  wireBlock: (el: HTMLElement) => void;
  /** Render one story paragraph by full anchor id, with the editor's own render profile. */
  renderBlock: (anchorId: string) => HTMLElement | null;
  /** Rebuild the editor's unid → full-anchor map (a seed adds a part, changing the projection). */
  refreshAnchorMap: () => void;
  /** Full anchor id of a rendered BODY block's unid, or undefined. */
  bodyAnchorIdOf: (unid: string) => string | undefined;
  /** Re-render the whole document (page view: a story changed height, or a kind flag flipped). */
  remount: () => void;
  /** The story host that is active changed (null = the caret left every story). */
  onActiveChange?: (host: HTMLElement | null, which: BandWhich | null) => void;
}

interface HeaderFooterRefLite {
  kind: HeaderFooterKind;
  partUri: string;
  inherited?: boolean;
}

interface SectionInfoLite {
  sectionUnid: string;
  headerRefs?: HeaderFooterRefLite[];
  footerRefs?: HeaderFooterRefLite[];
  pageNumberStart?: number;
  pageNumberFormat?: NumberFormat;
  /** `w:titlePg` on the governing sectPr (older bundles omit it: absent = unknown). */
  titlePage?: boolean;
  /** `w:evenAndOddHeaders` in the settings part. */
  evenAndOddHeaders?: boolean;
}

interface AnchorEntryLite {
  partUri: string;
  unid: string;
  kind: string;
  scope: string;
}

/** Kinds/scopes the markdown projection addresses as editable text blocks. */
const STORY_BLOCK_KINDS = new Set(["p", "h", "li"]);

const KIND_LABELS: Record<HeaderFooterKind, { header: string; footer: string }> = {
  default: { header: "Header", footer: "Footer" },
  first: { header: "First Page Header", footer: "First Page Footer" },
  even: { header: "Even Page Header", footer: "Even Page Footer" },
};

/** Attributes the editor stamps on live story blocks and that a page clone must not carry. */
const LIVE_BLOCK_ATTRS = [
  "contenteditable", "data-anchor", "data-hf-anchor", "data-hf-adopted", "data-hf-focused",
  "data-committed-text", "data-render-sig", "tabindex",
];

export class HeaderFooterRegion {
  readonly headerBand: HTMLElement;
  readonly footerBand: HTMLElement;

  private readonly bridge: HeaderFooterBridge;
  private readonly handle: number;
  private readonly callbacks: HeaderFooterRegionCallbacks;

  /** The body anchor whose section the region currently describes (needed to seed a story). */
  private bodyAnchorId: string | null = null;
  private sectionInfo: SectionInfoLite | null = null;
  /** Memoized body anchor → governing sectionUnid, so repeat focus costs no bridge call. */
  private readonly sectionOfAnchor = new Map<string, string>();

  /** Which story kind each continuous band shows. */
  private readonly kinds: Record<BandWhich, HeaderFooterKind> = { header: "default", footer: "default" };

  /** Page view: the page stack being edited in place, and the hosts currently holding live blocks. */
  private pageRoot: HTMLElement | null = null;
  private readonly pageHosts = new Set<HTMLElement>();
  private activeHost: HTMLElement | null = null;
  private readonly pageListeners = new Map<HTMLElement, (event: Event) => void>();

  constructor(bridge: HeaderFooterBridge, handle: number, callbacks: HeaderFooterRegionCallbacks) {
    this.bridge = bridge;
    this.handle = handle;
    this.callbacks = callbacks;
    this.headerBand = this.buildBand("header");
    this.footerBand = this.buildBand("footer");
  }

  // ─── public surface ──────────────────────────────────────────────────

  /**
   * Point the region at the section governing `bodyAnchorId` and repaint if it changed.
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
    if (!info) return; // not a body anchor (or no sectPr) — leave things as they are
    this.sectionOfAnchor.set(bodyAnchorId, info.sectionUnid);
    const changed = info.sectionUnid !== this.sectionInfo?.sectionUnid;
    this.bodyAnchorId = bodyAnchorId;
    this.sectionInfo = info;
    if (changed && !this.pageRoot) {
      // A kind the new section cannot show falls back to its default story.
      for (const which of ["header", "footer"] as BandWhich[]) {
        if (!this.partUriFor(which, this.kinds[which])) this.kinds[which] = "default";
      }
      this.refreshAll();
    }
  }

  /** Re-read the section and repaint one story (band, or the active page host and its clones). */
  refresh(which: BandWhich): void {
    this.sectionInfo = this.bodyAnchorId ? this.readSectionInfo(this.bodyAnchorId) : null;
    if (this.pageRoot) this.repaintPageStory(which);
    else this.renderBand(which);
  }

  /** Repaint both stories from the live session (after a remount, undo/redo, or section change). */
  refreshAll(): void {
    if (this.pageRoot) {
      this.repaintPageStory("header");
      this.repaintPageStory("footer");
      return;
    }
    this.renderBand("header");
    this.renderBand("footer");
  }

  /** Select (and, if absent, create) the story kind a continuous band shows. */
  setKind(which: BandWhich, kind: HeaderFooterKind): void {
    this.kinds[which] = kind;
    if (!this.partUriFor(which, kind)) this.seedStory(which, kind);
    // Selecting first/even IS the user saying "use a different first/even page", so the
    // section's flag must be set even when the part already exists: Word leaves first/even parts
    // behind when those options are switched off, so a real document commonly has the reference
    // WITHOUT the flag; without this the typed story is saved but never rendered.
    if (this.bodyAnchorId && kind !== "default") {
      this.bridge.EnsureHeaderFooterVisible(this.handle, this.bodyAnchorId, kind);
    }
    this.refresh(which);
  }

  /** The kind a band (or the active page host) is currently editing. */
  kindOf(which: BandWhich): HeaderFooterKind {
    if (this.pageRoot && this.activeHost && this.whichOf(this.activeHost) === which) {
      return (this.activeHost.dataset.hfType as HeaderFooterKind) ?? "default";
    }
    return this.kinds[which];
  }

  /** Word's checkbox state: whether `w:titlePg` (first) / `w:evenAndOddHeaders` (even) is set. */
  kindEnabled(kind: "first" | "even"): boolean {
    const info = this.sectionInfo;
    if (!info) return false;
    if (kind === "first") return info.titlePage ?? this.refsFor("header").some((r) => r.kind === "first" && !r.inherited);
    return info.evenAndOddHeaders ?? false;
  }

  /**
   * Word's "Different first page" / "Different odd & even pages". Enabling seeds BOTH stories of
   * that kind when absent (page 1 uses its own header AND footer once `w:titlePg` is set, so a
   * missing first-page footer would leave page 1 footer-less) and sets the flag; disabling clears
   * the flag and leaves the parts in place, as Word does. Returns false when the bundle predates
   * the disable op and `enabled` is false.
   */
  setKindEnabled(kind: "first" | "even", enabled: boolean): boolean {
    if (!this.bodyAnchorId) return false;
    if (enabled) {
      for (const which of ["header", "footer"] as BandWhich[]) {
        // The odd/all-pages story too: once the option is on, its pages need somewhere to
        // type as much as the first/even ones do (Word shows an empty area for both).
        if (!this.partUriFor(which, "default")) this.seedStory(which, "default");
        if (!this.partUriFor(which, kind)) this.seedStory(which, kind);
      }
      const res = parseResult(this.bridge.EnsureHeaderFooterVisible(this.handle, this.bodyAnchorId, kind));
      if (!res.success) return false;
    } else {
      if (!this.bridge.SetHeaderFooterKindEnabled) return false;
      const res = parseResult(this.bridge.SetHeaderFooterKindEnabled(this.handle, this.bodyAnchorId, kind, false));
      if (!res.success) return false;
      for (const which of ["header", "footer"] as BandWhich[]) {
        if (this.kinds[which] === kind) this.kinds[which] = "default";
      }
    }
    this.reloadSection();
    // Page view selects stories per page from the flags, so the pages themselves must rebuild.
    if (this.pageRoot) this.callbacks.remount();
    return true;
  }

  /** Human label for the story a band (or the active host) shows — "First Page Header". */
  storyLabel(which: BandWhich): string {
    return KIND_LABELS[this.kindOf(which)][which];
  }

  /** The kinds a band could show right now (those with a story or an enabled flag). */
  availableKinds(which: BandWhich): HeaderFooterKind[] {
    const kinds: HeaderFooterKind[] = ["default"];
    if (this.kindEnabled("first") || this.partUriFor(which, "first")) kinds.push("first");
    if (this.kindEnabled("even") || this.partUriFor(which, "even")) kinds.push("even");
    return kinds;
  }

  /** Append a PAGE / NUMPAGES field to the story paragraph addressed by `anchorId`.
   *  Deliberately a PLAIN field (no `\*` switch), exactly as Word inserts one, so it follows the
   *  section's page-number format — see {@link setPageNumbering}. */
  insertPageNumber(which: BandWhich, anchorId: string, field: "currentPage" | "totalPages" | "pageOfTotal"): boolean {
    const res = parseResult(this.bridge.InsertPageNumberField(this.handle, anchorId, field, ""));
    if (!res.success) return false;
    this.refresh(which);
    return true;
  }

  /** Insert a page number into a story's own target paragraph (focused, else last). Seeds the
   *  story first if the selected kind has none, so the command is never a silent no-op. */
  insertPageNumberInBand(which: BandWhich, field: "currentPage" | "totalPages" | "pageOfTotal"): boolean {
    // Page view: the field goes into a live story, so open the story on the current page
    // first (Word does the same — inserting a page number takes you into the footer).
    if (this.pageRoot && !(this.activeHost && this.whichOf(this.activeHost) === which)) {
      const near = this.activeHost ?? null;
      if (!this.focusStory(which, near)) return false;
    }
    const kind = this.kindOf(which);
    if (!this.partUriFor(which, kind)) {
      this.seedStory(which, kind);
      this.refresh(which);
    }
    const target = this.pageNumberTarget(which);
    return target ? this.insertPageNumber(which, target, field) : false;
  }

  /** Set this section's page numbering (`w:pgNumType`) — Word's *Format Page Numbers…*. */
  setPageNumbering(op: { start?: number; format?: NumberFormat }): boolean {
    if (!this.bodyAnchorId) return false;
    const res = parseResult(this.bridge.SetPageNumbering(this.handle, this.bodyAnchorId, JSON.stringify(op)));
    if (!res.success) return false;
    this.reloadSection();
    return true;
  }

  /** This section's page numbering as the live document states it (fields absent, not defaulted). */
  pageNumbering(): { start?: number; format?: NumberFormat } {
    return { start: this.sectionInfo?.pageNumberStart, format: this.sectionInfo?.pageNumberFormat };
  }

  /** Remove this section's page-numbering start/format. */
  clearPageNumbering(): boolean {
    if (!this.bodyAnchorId) return false;
    const res = parseResult(this.bridge.ClearPageNumbering(this.handle, this.bodyAnchorId));
    if (!res.success) return false;
    this.reloadSection();
    return true;
  }

  /** The story host (continuous band or active page area) containing `node`, or null. */
  bandOf(node: Node | null): HTMLElement | null {
    const el = node && node.nodeType === 1 ? (node as HTMLElement) : node?.parentElement ?? null;
    const band = el?.closest<HTMLElement>("[data-hf-band]") ?? null;
    if (!band) return null;
    return band === this.headerBand || band === this.footerBand || this.pageHosts.has(band) ? band : null;
  }

  /** The block-list root (a story's container) owning `node`, or null. */
  blockRootOf(node: Node | null): HTMLElement | null {
    const band = this.bandOf(node);
    return band ? band.querySelector<HTMLElement>("[data-hf-body]") : null;
  }

  /** True when `node` is inside a story host (band or page area). */
  contains(node: Node | null): boolean {
    return this.bandOf(node) !== null;
  }

  /** `"header"` / `"footer"` for a story host produced by this region. */
  whichOf(band: HTMLElement): BandWhich {
    return band.getAttribute("data-hf-band") === "footer" ? "footer" : "header";
  }

  /** The story host currently being edited (page view), or null. */
  get active(): HTMLElement | null {
    return this.activeHost;
  }

  /** True while the caret is in a story (either view) — drives the contextual ribbon tab. */
  isStoryActive(): boolean {
    return this.activeHost !== null;
  }

  /**
   * Track focus: the editor calls this whenever any block takes focus. A body block leaving a
   * page story deactivates it (and re-paginates if the story grew past its band); a story block
   * activates its host.
   */
  noteFocus(el: HTMLElement | null): void {
    const host = el ? this.bandOf(el) : null;
    if (host === this.activeHost) return;
    const previous = this.activeHost;
    this.activeHost = host;
    if (previous) this.deactivateHost(previous);
    if (host) {
      host.setAttribute("data-hf-active", "");
      if (this.pageHosts.has(host)) host.style.overflow = "visible";
    }
    this.callbacks.onActiveChange?.(host, host ? this.whichOf(host) : null);
  }

  /** Leave story editing: the caret goes back to the body (Word's "Close Header and Footer"). */
  close(): void {
    const previous = this.activeHost;
    this.activeHost = null;
    if (previous) this.deactivateHost(previous);
    this.callbacks.onActiveChange?.(null, null);
  }

  /**
   * Move the caret into a story (Word's "Go to Header / Go to Footer"). Page view activates the
   * area on the page the caret is on (else the first page); continuous view focuses the band.
   */
  focusStory(which: BandWhich, near?: HTMLElement | null): boolean {
    if (this.pageRoot) {
      const pages = Array.from(this.pageRoot.querySelectorAll<HTMLElement>(".page-box"));
      let page = near?.closest<HTMLElement>(".page-box") ?? pages[0] ?? null;
      if (!page) return false;
      let area = page.querySelector<HTMLElement>(`.page-${which}`);
      if (!area) {
        // No story of this kind exists, so the paginator drew no area to click. Seed it and
        // re-paginate; the page then carries an (empty) area for the caret to land in.
        const pageIndex = pages.indexOf(page);
        const bodyAnchor = this.bodyAnchorForPage(page);
        if (!bodyAnchor) return false;
        this.syncToBody(bodyAnchor);
        const kind = (page.querySelector<HTMLElement>(`.page-${which === "header" ? "footer" : "header"}`)?.dataset.hfType as HeaderFooterKind | undefined) ?? "default";
        if (!this.partUriFor(which, kind)) this.seedStory(which, kind);
        this.callbacks.remount();
        if (!this.pageRoot) return false;
        page = this.pageRoot.querySelectorAll<HTMLElement>(".page-box")[Math.max(0, pageIndex)] ?? null;
        area = page?.querySelector<HTMLElement>(`.page-${which}`) ?? null;
        if (!area) return false;
      }
      this.editInPage(area, "end");
      return true;
    }
    const band = which === "header" ? this.headerBand : this.footerBand;
    const block = band.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]');
    if (!block) return false;
    placeCaretAtEnd(block);
    return true;
  }

  /** After the editor swapped/split/merged a story block in place: show this page's own numbers
   *  in the fresh render, then mirror it to the page clones. */
  afterStoryEdit(el: HTMLElement): void {
    const host = this.bandOf(el);
    if (!host || !this.pageRoot) return;
    this.substituteFields(host);
    this.propagate(host);
  }

  // ─── page view: edit in place ────────────────────────────────────────

  /**
   * Adopt a page stack: every page's header/footer area becomes click-to-edit. The paginator's
   * clones stay as they are until a click swaps one for the live story.
   */
  attachPages(pageRoot: HTMLElement): void {
    this.detachPages();
    this.pageRoot = pageRoot;
    for (const area of Array.from(pageRoot.querySelectorAll<HTMLElement>(".page-header, .page-footer"))) {
      area.setAttribute("data-hf-page", area.classList.contains("page-footer") ? "footer" : "header");
      area.setAttribute("data-hf-inert", "");
      area.title = `Click to edit the ${area.classList.contains("page-footer") ? "footer" : "header"}`;
      const listener = (event: Event) => {
        if (this.pageHosts.has(area)) return; // already live — the click lands in a block
        event.preventDefault();
        this.editInPage(area, "point", event as MouseEvent);
      };
      area.addEventListener("mousedown", listener);
      this.pageListeners.set(area, listener);
    }
  }

  /** Release a page stack (before a remount replaces it). */
  detachPages(): void {
    for (const [area, listener] of this.pageListeners) area.removeEventListener("mousedown", listener);
    this.pageListeners.clear();
    this.pageHosts.clear();
    const hadActive = this.activeHost !== null;
    this.activeHost = null;
    this.pageRoot = null;
    // The pages are being thrown away, so there is no host to deactivate — but the caret WAS
    // in a story, and the ribbon's contextual tab and story label follow this callback. Without
    // it, the next body focus compared null to null and never published "back in the body".
    if (hadActive) this.callbacks.onActiveChange?.(null, null);
  }

  /** True when the region is presenting stories inside page boxes rather than as bands. */
  get inPageMode(): boolean {
    return this.pageRoot !== null;
  }

  /**
   * Swap a page's cloned story for the live, editable one and put the caret in it. The page
   * advertises which story it shows (`data-hf-type`, stamped by the paginator) and which
   * section it belongs to (`data-section-index`); the section's body anchor comes from the
   * first anchored block on the page (or an earlier page of the same section).
   */
  private editInPage(area: HTMLElement, caret: "point" | "end", event?: MouseEvent): void {
    if (!this.pageRoot) return;
    const which: BandWhich = area.getAttribute("data-hf-page") === "footer" ? "footer" : "header";
    const kind = (area.dataset.hfType as HeaderFooterKind | undefined) ?? "default";
    const page = area.closest<HTMLElement>(".page-box");
    const bodyAnchor = page ? this.bodyAnchorForPage(page) : null;
    if (!bodyAnchor) return;
    this.syncToBody(bodyAnchor);
    this.kinds[which] = kind;

    if (!this.pageHosts.has(area)) {
      area.setAttribute("data-hf-band", which);
      area.removeAttribute("data-hf-inert");
      area.dataset.hfLabel = KIND_LABELS[kind][which];
      this.pageHosts.add(area);
      this.renderHostStory(area, which, kind);
    }
    const blocks = Array.from(area.querySelectorAll<HTMLElement>('[data-anchor][contenteditable="true"]'));
    if (blocks.length === 0) return;
    let target: HTMLElement = blocks[blocks.length - 1];
    if (caret === "point" && event) {
      const hit = (event.target as HTMLElement | null)?.closest<HTMLElement>('[data-anchor][contenteditable="true"]');
      if (hit && area.contains(hit)) target = hit;
      const doc = area.ownerDocument;
      const point = caretFromPoint(doc, event.clientX, event.clientY);
      if (point && target.contains(point.node)) {
        const sel = doc.getSelection();
        const range = doc.createRange();
        range.setStart(point.node, point.offset);
        range.collapse(true);
        sel?.removeAllRanges();
        sel?.addRange(range);
        target.focus({ preventScroll: true });
        return;
      }
    }
    placeCaretAtEnd(target);
  }

  /** The full body anchor of the first anchored block on `page`, walking back to earlier pages. */
  private bodyAnchorForPage(page: HTMLElement): string | null {
    let box: HTMLElement | null = page;
    while (box) {
      const block = box.querySelector<HTMLElement>(".page-content [data-anchor]");
      const unid = block?.getAttribute("data-anchor");
      const id = unid ? this.callbacks.bodyAnchorIdOf(unid) : undefined;
      if (id) return id;
      box = box.previousElementSibling as HTMLElement | null;
    }
    return null;
  }

  /** Render the live story into a page host (replacing the paginator's clone). */
  private renderHostStory(host: HTMLElement, which: BandWhich, kind: HeaderFooterKind): void {
    let body = host.querySelector<HTMLElement>(":scope > [data-hf-body]");
    if (!body) {
      host.replaceChildren();
      body = host.ownerDocument.createElement("div");
      body.className = "docx-hf-body";
      body.setAttribute("data-hf-body", "");
      host.appendChild(body);
    }
    const partUri = this.partUriFor(which, kind);
    if (!partUri) {
      // The page shows nothing because the story does not exist; Word still lets you click into
      // the margin and type. Seed it and render the empty paragraph.
      this.seedStory(which, kind);
    }
    this.fillStoryBody(body, which, kind);
    this.substituteFields(host);
  }

  /** Repaint the page story `which` (active host first, then every clone of the same story). */
  private repaintPageStory(which: BandWhich): void {
    const host = this.activeHost && this.whichOf(this.activeHost) === which ? this.activeHost : null;
    if (host) {
      const body = host.querySelector<HTMLElement>(":scope > [data-hf-body]");
      if (body) this.fillStoryBody(body, which, this.kindOf(which));
      this.substituteFields(host);
      this.propagate(host);
      return;
    }
    // No live host for this story — page clones carry the paginator's render; nothing to do
    // until the next remount picks the change up.
  }

  /** Mirror a live host's story onto every other page area showing the same story. */
  private propagate(host: HTMLElement): void {
    if (!this.pageRoot) return;
    const which = this.whichOf(host);
    const kind = (host.dataset.hfType as HeaderFooterKind | undefined) ?? "default";
    const partUri = this.partUriFor(which, kind);
    const sourceBody = host.querySelector<HTMLElement>(":scope > [data-hf-body]");
    if (!sourceBody) return;
    const sectionParts = new Map<string, string | null>();
    for (const area of Array.from(this.pageRoot.querySelectorAll<HTMLElement>(`.page-${which}`))) {
      if (area === host) continue;
      if ((area.dataset.hfType ?? "default") !== kind) continue;
      const page = area.closest<HTMLElement>(".page-box");
      if (!page) continue;
      const sectionIndex = page.dataset.sectionIndex ?? "0";
      let part = sectionParts.get(sectionIndex);
      if (part === undefined) {
        const anchor = this.bodyAnchorForPage(page);
        const info = anchor ? this.readSectionInfo(anchor) : null;
        const refs = (which === "header" ? info?.headerRefs : info?.footerRefs) ?? [];
        part = refs.find((r) => r.kind === kind)?.partUri ?? null;
        sectionParts.set(sectionIndex, part);
      }
      if (!partUri || part !== partUri) continue;
      if (this.pageHosts.has(area)) {
        // Another page already holds a live copy of this story: re-render it from truth too.
        const body = area.querySelector<HTMLElement>(":scope > [data-hf-body]");
        if (body) this.fillStoryBody(body, which, kind);
      } else {
        area.replaceChildren();
        for (const child of Array.from(sourceBody.children)) area.appendChild(inertClone(child as HTMLElement));
      }
      this.substituteFields(area);
    }
  }

  /** Substitute PAGE / NUMPAGES markers in a page area from the page's own numbering. */
  private substituteFields(area: HTMLElement): void {
    const page = area.closest<HTMLElement>(".page-box");
    if (!page || !this.pageRoot) return;
    const total = this.pageRoot.querySelectorAll(".page-box:not([data-section-filler])").length;
    const displayed = parseInt(page.dataset.displayedPageNumber || page.dataset.pageNumber || "1", 10);
    const format = this.sectionInfo?.pageNumberFormat;
    for (const marker of Array.from(area.querySelectorAll<HTMLElement>("[data-field]"))) {
      const kind = marker.dataset.field;
      if (kind !== "PAGE" && kind !== "NUMPAGES") continue;
      // Keep the session's cached result on the element: the editor's offset space counts THAT,
      // not the per-page number shown here, so a later edit diffs against the right text.
      if (marker.dataset.fieldCached === undefined) marker.dataset.fieldCached = marker.textContent ?? "";
      marker.textContent = formatPageNumber(kind === "PAGE" ? displayed : total, marker.dataset.fieldFormat ?? format);
    }
  }

  private deactivateHost(host: HTMLElement): void {
    host.removeAttribute("data-hf-active");
    if (!this.pageHosts.has(host)) return;
    host.style.overflow = "hidden";
    // The story may have grown past the band the paginator reserved for it; re-flow the pages
    // rather than clip it (Word pushes the body down).
    const body = host.querySelector<HTMLElement>(":scope > [data-hf-body]");
    // A couple of pixels of slack: the paginator measured the story in a detached box, and
    // sub-pixel rounding between that and the live host must not read as growth.
    if (body && body.scrollHeight > host.clientHeight + 3) {
      this.callbacks.remount();
    }
  }

  // ─── section + anchor resolution ─────────────────────────────────────

  private readSectionInfo(bodyAnchorId: string): SectionInfoLite | null {
    try {
      const parsed = JSON.parse(this.bridge.GetSectionInfo(this.handle, bodyAnchorId)) as SectionInfoLite | null;
      return parsed && typeof parsed.sectionUnid === "string" ? parsed : null;
    } catch {
      return null;
    }
  }

  private reloadSection(): void {
    this.sectionInfo = this.bodyAnchorId ? this.readSectionInfo(this.bodyAnchorId) : null;
    this.refreshAll();
  }

  private refsFor(which: BandWhich): HeaderFooterRefLite[] {
    const info = this.sectionInfo;
    if (!info) return [];
    return (which === "header" ? info.headerRefs : info.footerRefs) ?? [];
  }

  private refFor(which: BandWhich, kind: HeaderFooterKind): HeaderFooterRefLite | null {
    return this.refsFor(which).find((r) => r.kind === kind) ?? null;
  }

  private partUriFor(which: BandWhich, kind: HeaderFooterKind): string | null {
    return this.refFor(which, kind)?.partUri ?? null;
  }

  /** Full anchor ids of the story paragraphs held in `partUri`, in document order. */
  private storyAnchors(partUri: string): string[] {
    let index: Record<string, AnchorEntryLite>;
    try {
      const raw = this.bridge.ListAnchors ? this.bridge.ListAnchors(this.handle) : this.bridge.Project(this.handle);
      index = (JSON.parse(raw) as { anchorIndex: Record<string, AnchorEntryLite> }).anchorIndex;
    } catch {
      return [];
    }
    return Object.entries(index)
      .filter(([, t]) => t.partUri === partUri && STORY_BLOCK_KINDS.has(t.kind))
      .map(([id]) => id);
  }

  /** Create an empty story for `kind` so there is a paragraph to click into. */
  private seedStory(which: BandWhich, kind: HeaderFooterKind): void {
    if (!this.bodyAnchorId) return;
    const set = which === "header" ? this.bridge.SetHeaderText : this.bridge.SetFooterText;
    const res = parseResult(set(this.handle, this.bodyAnchorId, kind, ""));
    if (!res.success) return;
    this.callbacks.refreshAnchorMap();
    this.sectionInfo = this.readSectionInfo(this.bodyAnchorId);
  }

  // ─── continuous view: bands ──────────────────────────────────────────

  private buildBand(which: BandWhich): HTMLElement {
    const band = document.createElement("div");
    band.className = "docx-hf-band";
    band.setAttribute("data-hf-band", which);

    const tag = document.createElement("div");
    tag.className = "docx-hf-tag";
    const label = document.createElement("span");
    label.className = "docx-hf-label";
    label.setAttribute("data-hf-label", "");
    label.textContent = KIND_LABELS.default[which];
    tag.appendChild(label);
    const kinds = document.createElement("span");
    kinds.className = "docx-hf-kinds";
    kinds.setAttribute("data-hf-kinds", "");
    kinds.hidden = true;
    tag.appendChild(kinds);
    const note = document.createElement("span");
    note.className = "docx-hf-inherited";
    note.setAttribute("data-hf-inherited-note", "");
    note.hidden = true;
    tag.appendChild(note);
    band.appendChild(tag);

    const body = document.createElement("div");
    body.className = "docx-hf-body";
    body.setAttribute("data-hf-body", "");
    band.appendChild(body);
    return band;
  }

  /** Which story paragraph a page-number insert targets: the focused one, else the last. */
  private pageNumberTarget(which: BandWhich): string | null {
    const host = this.pageRoot
      ? (this.activeHost && this.whichOf(this.activeHost) === which ? this.activeHost : null)
      : (which === "header" ? this.headerBand : this.footerBand);
    const body = host?.querySelector<HTMLElement>("[data-hf-body]");
    if (!body) return null;
    const focused = body.querySelector<HTMLElement>('[data-hf-focused="true"][data-hf-anchor]');
    const blocks = Array.from(body.querySelectorAll<HTMLElement>("[data-hf-anchor]"));
    return (focused ?? blocks[blocks.length - 1])?.getAttribute("data-hf-anchor") ?? null;
  }

  private renderBand(which: BandWhich): void {
    const band = which === "header" ? this.headerBand : this.footerBand;
    const body = band.querySelector<HTMLElement>("[data-hf-body]");
    if (!body) return;
    const kind = this.kinds[which];

    const label = band.querySelector<HTMLElement>("[data-hf-label]");
    if (label) label.textContent = KIND_LABELS[kind][which];

    // The story switcher shows only when there is more than one story to switch between.
    const kinds = band.querySelector<HTMLElement>("[data-hf-kinds]");
    if (kinds) {
      const available = this.availableKinds(which);
      kinds.hidden = available.length < 2;
      kinds.replaceChildren();
      for (const k of available) {
        const button = document.createElement("button");
        button.type = "button";
        button.setAttribute("data-hf-kind", k);
        button.textContent =
          k === "default"
            ? (this.kindEnabled("even") ? "Odd pages" : "All pages")
            : k === "first" ? "First page" : "Even pages";
        button.toggleAttribute("data-on", k === kind);
        button.addEventListener("mousedown", (e) => e.preventDefault());
        button.addEventListener("click", () => this.setKind(which, k));
        kinds.appendChild(button);
      }
    }

    const inherited = this.refFor(which, kind)?.inherited === true;
    band.toggleAttribute("data-hf-inherited", inherited);
    const note = band.querySelector<HTMLElement>("[data-hf-inherited-note]");
    if (note) {
      note.hidden = !inherited;
      note.textContent = inherited ? "Same as previous section" : "";
    }

    this.fillStoryBody(body, which, kind);
  }

  /** Render the story's paragraphs into `body` (a placeholder when there is no story yet). */
  private fillStoryBody(body: HTMLElement, which: BandWhich, kind: HeaderFooterKind): void {
    // A repaint replaces the story's nodes; if the caret was in one of them, put it back at the
    // end of the story (where a page-number field just landed) instead of dropping it.
    const hadFocus = body.contains(body.ownerDocument.activeElement);
    body.innerHTML = "";
    const band = body.closest<HTMLElement>("[data-hf-band]");
    const partUri = this.partUriFor(which, kind);
    const anchors = partUri ? this.storyAnchors(partUri) : [];
    if (anchors.length === 0) {
      body.appendChild(this.placeholder(which, kind));
      band?.setAttribute("data-hf-empty", "true");
      return;
    }
    band?.removeAttribute("data-hf-empty");
    let last: HTMLElement | null = null;
    for (const anchorId of anchors) {
      const el = this.callbacks.renderBlock(anchorId);
      if (!el) continue;
      body.appendChild(el);
      // Adopt BEFORE wiring: the editor resolves a block's anchor from `data-hf-anchor` first,
      // because a story paragraph's content-addressed unid can collide with another part's.
      this.adoptBlock(el, anchorId);
      this.callbacks.wireBlock(el);
      last = el;
    }
    if (hadFocus && last) placeCaretAtEnd(last);
  }

  /**
   * Mark a story paragraph as belonging to this region. The full anchor id is stamped alongside
   * `data-anchor` (which carries only the bare unid) so chrome can address the block without
   * going through the editor's unid map. Called on first render AND after an incremental swap.
   */
  adoptBlock(el: HTMLElement, anchorId: string): void {
    const body = this.blockRootOf(el);
    if (!body) return;
    el.setAttribute("data-hf-anchor", anchorId);
    if (el.dataset.hfAdopted === "true") return;
    el.dataset.hfAdopted = "true";
    el.addEventListener("focus", () => {
      body.querySelectorAll("[data-hf-focused]").forEach((sib) => sib.removeAttribute("data-hf-focused"));
      el.setAttribute("data-hf-focused", "true");
    });
  }

  private placeholder(which: BandWhich, kind: HeaderFooterKind): HTMLElement {
    const el = document.createElement("div");
    el.className = "docx-hf-placeholder";
    el.setAttribute("data-hf-placeholder", "");
    el.textContent = `No ${KIND_LABELS[kind][which].toLowerCase()} yet — click here to add one.`;
    el.addEventListener("mousedown", (event) => {
      event.preventDefault();
      if (!this.bodyAnchorId) return;
      this.seedStory(which, kind);
      this.refresh(which);
      const block = el.parentElement?.querySelector<HTMLElement>('[data-anchor][contenteditable="true"]');
      if (block) placeCaretAtEnd(block);
    });
    return el;
  }
}

/** A presentation-only copy of a live story block for another page. */
function inertClone(el: HTMLElement): HTMLElement {
  const clone = el.cloneNode(true) as HTMLElement;
  for (const node of [clone, ...Array.from(clone.querySelectorAll<HTMLElement>("*"))]) {
    for (const attr of LIVE_BLOCK_ATTRS) node.removeAttribute(attr);
  }
  return clone;
}

function placeCaretAtEnd(el: HTMLElement): void {
  const doc = el.ownerDocument;
  el.focus({ preventScroll: true });
  const sel = doc.getSelection();
  if (!sel) return;
  const range = doc.createRange();
  range.selectNodeContents(el);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);
}

function caretFromPoint(doc: Document, x: number, y: number): { node: Node; offset: number } | null {
  const d = doc as Document & {
    caretPositionFromPoint?: (x: number, y: number) => { offsetNode: Node; offset: number } | null;
    caretRangeFromPoint?: (x: number, y: number) => Range | null;
  };
  const position = d.caretPositionFromPoint?.(x, y);
  if (position) return { node: position.offsetNode, offset: position.offset };
  const range = d.caretRangeFromPoint?.(x, y);
  return range ? { node: range.startContainer, offset: range.startOffset } : null;
}

function parseResult(json: string): { success: boolean } {
  try {
    return JSON.parse(json) as { success: boolean };
  } catch {
    return { success: false };
  }
}

/**
 * CommentGutter — Word-style comment bubbles beside the page.
 *
 * The converter renders a commented range inline (`span.comment-highlight[data-comment-id]`
 * around the runs, `a.comment-marker` at the reference), and `DocxSession.ListComments` is the
 * truth about the threads themselves. This module joins the two: for every thread root it finds
 * the highlight in the live DOM, places a bubble in a column beside the sheet at the highlight's
 * vertical position (stacking bubbles downward when they would overlap, the way Word's markup
 * area does), and draws a leader line from the highlight to the bubble.
 *
 * Nothing here is a second source of state. The bubbles are re-derived from `listComments()`
 * plus the DOM on every layout pass, and every action (reply, resolve, edit, delete, post a new
 * comment) goes through the editor's session commands, then re-lays out.
 *
 * The gutter lives INSIDE the editor's scrolling container so bubbles scroll with the page; the
 * host is responsible for reserving horizontal room for it (the ribbon pads the surface).
 */

import type { CommentListEntry } from "./types.js";

/** What the gutter needs from its editor — the comment commands plus the DOM it overlays. */
export interface CommentGutterHost {
  /** The element the document renders into; the gutter is appended to it. */
  readonly container: HTMLElement;
  /** Author for comments and replies posted from the gutter. */
  readonly commentAuthor: string;
  listComments(): CommentListEntry[];
  /** Post a comment on `target` (captured at draft time — the live selection is long gone once
   *  the user has typed into the bubble). Returns the created thread root, or null. */
  addComment(
    markdown: string,
    author: string,
    target?: CommentTarget,
  ): CommentListEntry | null;
  addCommentReply(parentAnchorId: string, markdown: string, author: string): boolean;
  updateComment(anchorId: string, markdown: string): boolean;
  removeComment(anchorId: string): boolean;
  setCommentResolved(anchorId: string, resolved: boolean): boolean;
  /** The block that would receive a new comment, and the selection span inside it. */
  commentTarget(): CommentTarget | null;
}

export interface CommentGutterOptions {
  /** Column width in px. Default 248. */
  width?: number;
  /** Vertical gap between stacked bubbles. Default 8. */
  gap?: number;
  /** Called after every layout with the number of visible (unresolved) threads. */
  onChange?: (info: { threads: number; open: number; active: string | null }) => void;
}

interface ThreadView {
  root: CommentListEntry;
  replies: CommentListEntry[];
  anchor: HTMLElement | null;
  bubble: HTMLElement;
  desiredTop: number;
  /** Horizontal position of the anchor, so threads on one line keep reading order. */
  desiredLeft: number;
}

const HIGHLIGHT_SELECTOR = "span.comment-highlight[data-comment-id]";
const MARKER_SELECTOR = "a.comment-marker[data-comment-id]";

/** Whether a `data-comment-id` value (possibly "3,5" for overlapping ranges) names `id`. */
function idsOf(el: Element): string[] {
  return (el.getAttribute("data-comment-id") ?? "").split(",").map((s) => s.trim()).filter(Boolean);
}

function formatDate(iso: string | undefined): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  try {
    return new Intl.DateTimeFormat(undefined, {
      month: "short", day: "numeric", year: "numeric", hour: "numeric", minute: "2-digit",
    }).format(d);
  } catch {
    return d.toLocaleString();
  }
}

function initialsOf(author: string, initials?: string): string {
  if (initials && initials.trim()) return initials.trim().slice(0, 3).toUpperCase();
  const parts = author.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  return parts.slice(0, 2).map((p) => p[0]!.toUpperCase()).join("");
}

/** A stable per-author hue so every bubble and highlight of one reviewer share a colour. */
function authorHue(author: string): number {
  let h = 0;
  for (let i = 0; i < author.length; i++) h = (h * 31 + author.charCodeAt(i)) >>> 0;
  return h % 360;
}

/**
 * What a comment is about to be attached to: a block and the selection span in it (null = the
 * whole paragraph). `anchor` and `text` record which paragraph, with what content, the target
 * was captured on, so a post that comes after the block was re-rendered or edited can find the
 * live node and tell whether the span still means the same characters.
 */
export interface CommentTarget {
  block: HTMLElement;
  span: { start: number; length: number } | null;
  anchor?: string;
  text?: string;
}

/**
 * Every comment hanging off `rootAnchorId`, transitively — replies to replies included — deepest
 * first, so removing them in order never leaves an orphan behind. The root itself is not listed.
 */
export function threadMembers(comments: readonly CommentListEntry[], rootAnchorId: string): string[] {
  const children = new Map<string, string[]>();
  for (const c of comments) {
    if (!c.parentAnchorId) continue;
    const list = children.get(c.parentAnchorId) ?? [];
    list.push(c.anchorId);
    children.set(c.parentAnchorId, list);
  }
  const out: string[] = [];
  const seen = new Set<string>([rootAnchorId]);
  const walk = (id: string) => {
    for (const child of children.get(id) ?? []) {
      if (seen.has(child)) continue;
      seen.add(child);
      walk(child);
      out.push(child);
    }
  };
  walk(rootAnchorId);
  return out;
}

export class CommentGutter {
  readonly element: HTMLElement;

  private readonly host: CommentGutterHost;
  private readonly options: Required<Omit<CommentGutterOptions, "onChange">> &
    Pick<CommentGutterOptions, "onChange">;
  private readonly leaders: SVGSVGElement;
  private frame: number | null = null;
  private mutations: MutationObserver | null = null;
  private resize: ResizeObserver | null = null;
  private activeId: string | null = null;
  private draft: HTMLElement | null = null;
  private draftTarget: CommentTarget | null = null;
  private draftAnchorTop = 0;
  private editing: string | null = null;
  private replying: string | null = null;
  private expanded = new Set<string>();
  private visible = true;
  private disposed = false;
  private readonly onContainerClick: (event: MouseEvent) => void;

  constructor(host: CommentGutterHost, options: CommentGutterOptions = {}) {
    this.host = host;
    this.options = {
      width: options.width ?? 248,
      gap: options.gap ?? 8,
      onChange: options.onChange,
    };
    const doc = host.container.ownerDocument;
    this.element = doc.createElement("aside");
    this.element.className = "docx-comment-gutter";
    this.element.setAttribute("data-docx-comments", "");
    this.element.style.width = `${this.options.width}px`;
    this.leaders = doc.createElementNS("http://www.w3.org/2000/svg", "svg");
    this.leaders.setAttribute("class", "docx-comment-leaders");
    this.leaders.setAttribute("aria-hidden", "true");
    if (getComputedStyle(host.container).position === "static") {
      host.container.style.position = "relative";
    }
    host.container.appendChild(this.leaders);
    host.container.appendChild(this.element);

    this.onContainerClick = (event) => {
      const target = event.target as HTMLElement | null;
      if (!target || this.element.contains(target)) return;
      const hit = target.closest<HTMLElement>(`${HIGHLIGHT_SELECTOR}, ${MARKER_SELECTOR}`);
      if (!hit) return;
      const id = idsOf(hit)[0];
      if (id) this.setActive(id, { scrollBubble: true });
    };
    host.container.addEventListener("click", this.onContainerClick);

    if (typeof MutationObserver !== "undefined") {
      this.mutations = new MutationObserver((records) => {
        for (const record of records) {
          const node = record.target as Node;
          if (this.element.contains(node) || this.leaders.contains(node)) continue;
          this.schedule();
          return;
        }
      });
      this.mutations.observe(host.container, {
        childList: true, subtree: true, characterData: true, attributes: true,
        attributeFilter: ["style", "class", "data-anchor"],
      });
    }
    if (typeof ResizeObserver !== "undefined") {
      this.resize = new ResizeObserver(() => this.schedule());
      this.resize.observe(host.container);
    }
    this.schedule();
  }

  // ── public surface ──────────────────────────────────────────────────────────

  /** Re-derive and re-position every bubble on the next animation frame. */
  schedule(): void {
    if (this.disposed || this.frame != null) return;
    const view = this.host.container.ownerDocument.defaultView;
    if (!view) return;
    this.frame = view.requestAnimationFrame(() => {
      this.frame = null;
      this.layout();
    });
  }

  /** Synchronous layout (tests, and callers that need the DOM settled now). */
  layout(): void {
    if (this.disposed) return;
    const doc = this.host.container.ownerDocument;
    const comments = this.host.listComments();
    const byId = new Map<string, CommentListEntry>();
    for (const entry of comments) byId.set(entry.anchorId, entry);
    const roots = comments.filter((c) => !c.parentAnchorId || !byId.has(c.parentAnchorId));
    const repliesOf = new Map<string, CommentListEntry[]>();
    for (const entry of comments) {
      if (!entry.parentAnchorId || !byId.has(entry.parentAnchorId)) continue;
      let root = byId.get(entry.parentAnchorId)!;
      // Reply-to-reply chains hang off the thread root, as Word draws them.
      const seen = new Set<string>();
      while (root.parentAnchorId && byId.has(root.parentAnchorId) && !seen.has(root.anchorId)) {
        seen.add(root.anchorId);
        root = byId.get(root.parentAnchorId)!;
      }
      const list = repliesOf.get(root.anchorId) ?? [];
      list.push(entry);
      repliesOf.set(root.anchorId, list);
    }

    // Reuse bubble nodes across layouts so an open textarea keeps its focus and text.
    const existing = new Map<string, HTMLElement>();
    for (const el of Array.from(this.element.querySelectorAll<HTMLElement>("[data-thread]"))) {
      existing.set(el.dataset.thread!, el);
    }

    // Measure with the gutter shown: an empty gutter is display:none, whose rect is all zeros,
    // and every bubble top below is relative to it. The toggle at the end hides it again when
    // nothing is left to show.
    this.element.toggleAttribute("data-empty", roots.length === 0 && !this.draft);
    const containerRect = this.host.container.getBoundingClientRect();
    const gutterRect = this.element.getBoundingClientRect();
    const views: ThreadView[] = [];
    for (const root of roots) {
      const anchor = this.anchorFor(root);
      const replies = repliesOf.get(root.anchorId) ?? [];
      let bubble = existing.get(root.anchorId);
      existing.delete(root.anchorId);
      const key = this.bubbleKey(root, replies);
      if (!bubble || bubble.dataset.key !== key) {
        const fresh = this.buildBubble(root, replies, doc);
        fresh.dataset.key = key;
        if (bubble) bubble.replaceWith(fresh);
        else this.element.appendChild(fresh);
        bubble = fresh;
      }
      const rect = anchor?.getBoundingClientRect();
      const desiredTop = rect ? rect.top - gutterRect.top : Number.POSITIVE_INFINITY;
      const desiredLeft = rect ? rect.left : Number.POSITIVE_INFINITY;
      views.push({ root, replies, anchor, bubble, desiredTop, desiredLeft });
      this.markHighlights(root, replies, anchor);
    }
    for (const stale of existing.values()) stale.remove();

    // Word orders the markup area by document position, not by comment id.
    views.sort(
      (a, b) =>
        a.desiredTop - b.desiredTop ||
        a.desiredLeft - b.desiredLeft ||
        a.root.anchorId.localeCompare(b.root.anchorId),
    );

    const gap = this.options.gap;
    let cursor = 0;
    const placed: Array<{ view: ThreadView | null; top: number; el: HTMLElement }> = [];
    let draftPlaced = false;
    const placeDraft = () => {
      if (!this.draft || draftPlaced) return;
      const top = Math.max(this.draftAnchorTop, cursor);
      this.draft.style.top = `${top}px`;
      cursor = top + this.draft.offsetHeight + gap;
      placed.push({ view: null, top, el: this.draft });
      draftPlaced = true;
    };
    for (const view of views) {
      if (this.draft && !draftPlaced && this.draftAnchorTop <= view.desiredTop) placeDraft();
      const orphan = !Number.isFinite(view.desiredTop);
      view.bubble.toggleAttribute("data-orphan", orphan);
      const wanted = orphan ? cursor : Math.max(0, view.desiredTop);
      const top = Math.max(wanted, cursor);
      view.bubble.style.top = `${top}px`;
      cursor = top + view.bubble.offsetHeight + gap;
      placed.push({ view, top, el: view.bubble });
    }
    placeDraft();

    // Leader lines: from the highlight's right edge to the bubble's top-left, in the
    // container's coordinate space (the SVG covers the whole scrollable content).
    const scrollW = this.host.container.scrollWidth;
    const scrollH = this.host.container.scrollHeight;
    this.leaders.setAttribute("width", String(scrollW));
    this.leaders.setAttribute("height", String(scrollH));
    this.leaders.style.width = `${scrollW}px`;
    this.leaders.style.height = `${scrollH}px`;
    const paths: string[] = [];
    const offX = this.host.container.scrollLeft - containerRect.left;
    const offY = this.host.container.scrollTop - containerRect.top;
    const gutterX = gutterRect.left + offX;
    for (const { view, top } of placed) {
      const rect = view?.anchor?.getBoundingClientRect();
      const active = view ? view.root.anchorId === this.activeId : true;
      const source = rect ?? (view === null && this.draftTarget
        ? this.draftTarget.block.getBoundingClientRect()
        : null);
      if (!source) continue;
      // Leave the highlight along its baseline, the way Word's connector does: a line at the
      // underline position reads as an extension of the highlight rather than a strike through
      // the rest of the line.
      const x1 = source.right + offX;
      const y1 = source.bottom + offY - 1;
      const y2 = gutterRect.top + offY + top + 14;
      const x2 = gutterX;
      const xm = x2 - 14;
      paths.push(
        `<path d="M${x1.toFixed(1)} ${y1.toFixed(1)} H${xm.toFixed(1)} L${x2.toFixed(1)} ${y2.toFixed(1)}"` +
          ` class="docx-comment-leader${active ? " docx-comment-leader-active" : ""}"/>`,
      );
    }
    this.leaders.innerHTML = paths.join("");
    this.element.toggleAttribute("data-empty", views.length === 0 && !this.draft);
    this.element.hidden = !this.visible;
    this.leaders.style.display = this.visible ? "" : "none";
    this.options.onChange?.({
      threads: views.length,
      open: views.filter((v) => !v.root.resolved).length,
      active: this.activeId,
    });
  }

  /** Re-append the gutter's elements after the container was emptied by a mount. */
  reattach(): void {
    if (this.disposed) return;
    this.host.container.appendChild(this.leaders);
    this.host.container.appendChild(this.element);
    this.schedule();
  }

  /** Show or hide the markup area (Word's "Show Comments"). */
  setVisible(visible: boolean): void {
    const wasVisible = this.visible;
    this.visible = visible;
    this.element.hidden = !visible;
    this.leaders.style.display = visible ? "" : "none";
    this.host.container.toggleAttribute("data-comments-hidden", !visible);
    // Layouts that ran while hidden measured a zero-height gutter and zero-height bubbles;
    // place everything again now that it can be measured.
    if (visible && !wasVisible) this.schedule();
  }

  get isVisible(): boolean {
    return this.visible;
  }

  /** The active thread root's anchor id, if any. */
  get active(): string | null {
    return this.activeId;
  }

  /** Activate a thread: its bubble lifts, its highlight darkens, and (optionally) it scrolls
   *  into view. Pass null to clear. */
  setActive(id: string | null, opts: { scrollBubble?: boolean; scrollAnchor?: boolean } = {}): void {
    const comments = this.host.listComments();
    // A numeric id (from a highlight) or a cmt anchor id both resolve to the thread root.
    let root: CommentListEntry | undefined;
    if (id != null) {
      const byId = new Map(comments.map((c) => [c.anchorId, c] as const));
      root = comments.find((c) => c.anchorId === id || String(c.id) === id);
      const seen = new Set<string>();
      while (root?.parentAnchorId && byId.has(root.parentAnchorId) && !seen.has(root.anchorId)) {
        seen.add(root.anchorId);
        root = byId.get(root.parentAnchorId);
      }
    }
    this.activeId = root?.anchorId ?? null;
    for (const el of Array.from(this.element.querySelectorAll<HTMLElement>("[data-thread]"))) {
      el.toggleAttribute("data-active", el.dataset.thread === this.activeId);
    }
    for (const el of Array.from(this.host.container.querySelectorAll<HTMLElement>(HIGHLIGHT_SELECTOR))) {
      el.classList.toggle("docx-comment-active", !!root && idsOf(el).includes(String(root.id)));
    }
    if (root && opts.scrollBubble) {
      this.element.querySelector<HTMLElement>(`[data-thread="${CSS.escape(root.anchorId)}"]`)
        ?.scrollIntoView({ block: "nearest", behavior: "smooth" });
    }
    if (root && opts.scrollAnchor) {
      this.anchorFor(root)?.scrollIntoView({ block: "center", behavior: "smooth" });
    }
    this.schedule();
  }

  /** Thread roots in document order — what Previous/Next step through. */
  threadsInOrder(): CommentListEntry[] {
    const comments = this.host.listComments();
    const byId = new Set(comments.map((c) => c.anchorId));
    const roots = comments.filter((c) => !c.parentAnchorId || !byId.has(c.parentAnchorId));
    return roots
      .map((root) => {
        const rect = this.anchorFor(root)?.getBoundingClientRect();
        return { root, top: rect?.top ?? Number.POSITIVE_INFINITY, left: rect?.left ?? Number.POSITIVE_INFINITY };
      })
      .sort((a, b) => a.top - b.top || a.left - b.left || a.root.anchorId.localeCompare(b.root.anchorId))
      .map((x) => x.root);
  }

  /** Activate the thread after (or before) the active one, wrapping around. */
  step(direction: 1 | -1): CommentListEntry | null {
    const order = this.threadsInOrder();
    if (order.length === 0) return null;
    const index = order.findIndex((c) => c.anchorId === this.activeId);
    const next = index < 0
      ? (direction > 0 ? order[0] : order[order.length - 1])
      : order[(index + direction + order.length) % order.length];
    this.setActive(next.anchorId, { scrollBubble: true, scrollAnchor: true });
    return next;
  }

  /**
   * Start a new comment on the current selection: a draft bubble appears beside it with a
   * focused textarea. Posting goes through the editor; cancelling removes the draft. Returns
   * false when nothing is selected/focused to comment on.
   */
  beginDraft(): boolean {
    const target = this.host.commentTarget();
    if (!target) return false;
    this.cancelDraft();
    this.draftTarget = target;
    const doc = this.host.container.ownerDocument;
    // The first comment in a document drafts against a gutter that is still display:none
    // (no threads); its rect is zeros until shown, and the draft's top is relative to it.
    this.element.removeAttribute("data-empty");
    const gutterRect = this.element.getBoundingClientRect();
    const sel = doc.defaultView?.getSelection();
    let rect: DOMRect | null = null;
    if (sel && sel.rangeCount > 0 && !sel.isCollapsed) {
      const r = sel.getRangeAt(0).getBoundingClientRect();
      if (r.height > 0) rect = r;
    }
    rect ??= target.block.getBoundingClientRect();
    this.draftAnchorTop = Math.max(0, rect.top - gutterRect.top);

    const bubble = doc.createElement("div");
    bubble.className = "docx-comment-bubble";
    bubble.setAttribute("data-draft", "");
    bubble.setAttribute("data-active", "");
    bubble.style.setProperty("--docx-comment-hue", String(authorHue(this.host.commentAuthor)));
    bubble.innerHTML =
      `<div class="docx-comment-head"><span class="docx-comment-avatar">${initialsOf(this.host.commentAuthor)}</span>` +
      `<span class="docx-comment-author"></span></div>` +
      `<textarea class="docx-comment-input" rows="3" placeholder="Type a comment…" data-comment-draft-text></textarea>` +
      `<div class="docx-comment-actions">` +
      `<button type="button" class="docx-comment-primary" data-comment-action="post">Post</button>` +
      `<button type="button" data-comment-action="cancel">Cancel</button></div>`;
    bubble.querySelector(".docx-comment-author")!.textContent = this.host.commentAuthor;
    const input = bubble.querySelector<HTMLTextAreaElement>("textarea")!;
    const post = () => {
      const text = input.value.trim();
      if (!text) { input.focus(); return; }
      const created = this.host.addComment(text, this.host.commentAuthor, target);
      this.cancelDraft();
      if (created) this.setActive(created.anchorId, { scrollBubble: true });
      else this.schedule();
    };
    bubble.querySelector('[data-comment-action="post"]')!.addEventListener("click", post);
    bubble.querySelector('[data-comment-action="cancel"]')!.addEventListener("click", () => this.cancelDraft());
    input.addEventListener("keydown", (event) => {
      const key = event as KeyboardEvent;
      if (key.key === "Escape") { key.preventDefault(); this.cancelDraft(); }
      if (key.key === "Enter" && (key.ctrlKey || key.metaKey)) { key.preventDefault(); post(); }
    });
    // Mousedown on the draft must not collapse the document selection until the target is
    // captured — it already is, so let the textarea take focus normally.
    this.draft = bubble;
    this.element.appendChild(bubble);
    this.activeId = null;
    this.layout();
    input.focus();
    return true;
  }

  /** Remove the draft bubble, if any. */
  cancelDraft(): void {
    if (!this.draft) return;
    this.draft.remove();
    this.draft = null;
    this.draftTarget = null;
    this.schedule();
  }

  get hasDraft(): boolean {
    return this.draft !== null;
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    const view = this.host.container.ownerDocument.defaultView;
    if (this.frame != null) view?.cancelAnimationFrame(this.frame);
    this.frame = null;
    this.mutations?.disconnect();
    this.resize?.disconnect();
    this.host.container.removeEventListener("click", this.onContainerClick);
    this.element.remove();
    this.leaders.remove();
  }

  // ── internals ───────────────────────────────────────────────────────────────

  /** The first highlight (else marker) carrying the thread's numeric id, in the live DOM. */
  private anchorFor(root: CommentListEntry): HTMLElement | null {
    const id = String(root.id);
    const candidates = this.host.container.querySelectorAll<HTMLElement>(
      `${HIGHLIGHT_SELECTOR}, ${MARKER_SELECTOR}`,
    );
    let marker: HTMLElement | null = null;
    for (const el of Array.from(candidates)) {
      if (this.element.contains(el)) continue;
      if (!idsOf(el).includes(id)) continue;
      // Hidden staging/registry copies must never win over the visible page.
      if (el.closest("#pagination-staging, [id^='pagination-'][style*='display:none'], [data-hf-inert]")) continue;
      if (el.matches(HIGHLIGHT_SELECTOR)) return el;
      marker ??= el;
    }
    return marker;
  }

  private markHighlights(root: CommentListEntry, replies: CommentListEntry[], anchor: HTMLElement | null): void {
    const ids = new Set([String(root.id), ...replies.map((r) => String(r.id))]);
    const hue = String(authorHue(root.author || ""));
    for (const el of Array.from(this.host.container.querySelectorAll<HTMLElement>(HIGHLIGHT_SELECTOR))) {
      if (this.element.contains(el)) continue;
      if (!idsOf(el).some((x) => ids.has(x))) continue;
      // Only write what changed: the document's MutationObserver schedules a layout on every
      // style/class mutation, so an unconditional write here would lay out forever.
      if (el.style.getPropertyValue("--docx-comment-hue") !== hue) el.style.setProperty("--docx-comment-hue", hue);
      el.toggleAttribute("data-comment-resolved", !!root.resolved);
      el.classList.toggle("docx-comment-active", root.anchorId === this.activeId);
    }
    void anchor;
  }

  private bubbleKey(root: CommentListEntry, replies: CommentListEntry[]): string {
    const parts = [root.anchorId, root.author, root.date ?? "", root.text, root.resolved ? "r" : "o"];
    for (const r of replies) parts.push(r.anchorId, r.author, r.date ?? "", r.text);
    parts.push(this.editing === root.anchorId ? "edit" : "", this.replying === root.anchorId ? "reply" : "");
    parts.push(this.expanded.has(root.anchorId) ? "x" : "");
    for (const r of replies) if (this.editing === r.anchorId) parts.push(`edit:${r.anchorId}`);
    return parts.join("");
  }

  private buildBubble(root: CommentListEntry, replies: CommentListEntry[], doc: Document): HTMLElement {
    const bubble = doc.createElement("div");
    bubble.className = "docx-comment-bubble";
    bubble.dataset.thread = root.anchorId;
    bubble.dataset.commentId = String(root.id);
    bubble.style.setProperty("--docx-comment-hue", String(authorHue(root.author || "")));
    if (root.resolved) bubble.setAttribute("data-resolved", "");
    if (root.anchorId === this.activeId) bubble.setAttribute("data-active", "");
    const collapsed = !!root.resolved && !this.expanded.has(root.anchorId);
    if (collapsed) bubble.setAttribute("data-collapsed", "");

    bubble.appendChild(this.buildEntry(root, doc, true));
    for (const reply of replies) bubble.appendChild(this.buildEntry(reply, doc, false));

    if (!collapsed) {
      if (this.replying === root.anchorId) {
        const input = doc.createElement("textarea");
        input.className = "docx-comment-input";
        input.rows = 2;
        input.placeholder = "Reply…";
        input.setAttribute("data-comment-reply-text", "");
        const actions = doc.createElement("div");
        actions.className = "docx-comment-actions";
        const post = doc.createElement("button");
        post.type = "button";
        post.className = "docx-comment-primary";
        post.textContent = "Reply";
        post.setAttribute("data-comment-action", "post-reply");
        const cancel = doc.createElement("button");
        cancel.type = "button";
        cancel.textContent = "Cancel";
        cancel.setAttribute("data-comment-action", "cancel-reply");
        const submit = () => {
          const text = input.value.trim();
          if (!text) { input.focus(); return; }
          this.replying = null;
          this.host.addCommentReply(root.anchorId, text, this.host.commentAuthor);
          this.setActive(root.anchorId);
        };
        post.addEventListener("click", submit);
        cancel.addEventListener("click", () => { this.replying = null; this.schedule(); });
        input.addEventListener("keydown", (event) => {
          const key = event as KeyboardEvent;
          if (key.key === "Escape") { key.preventDefault(); this.replying = null; this.schedule(); }
          if (key.key === "Enter" && (key.ctrlKey || key.metaKey)) { key.preventDefault(); submit(); }
        });
        actions.append(post, cancel);
        bubble.append(input, actions);
        queueMicrotask(() => input.focus());
      } else {
        const bar = doc.createElement("div");
        bar.className = "docx-comment-bar";
        const reply = doc.createElement("button");
        reply.type = "button";
        reply.textContent = "Reply";
        reply.setAttribute("data-comment-action", "reply");
        reply.addEventListener("click", () => {
          this.replying = root.anchorId;
          this.editing = null;
          this.setActive(root.anchorId);
          // Build the reply box now, not on the next frame: the click handler returns before a
          // scheduled layout would run, and a keystroke that lands in between goes to the page
          // instead of the textarea — the first letter of a quickly typed reply vanished.
          this.layout();
        });
        const resolve = doc.createElement("button");
        resolve.type = "button";
        resolve.textContent = root.resolved ? "Reopen" : "Resolve";
        resolve.setAttribute("data-comment-action", root.resolved ? "reopen" : "resolve");
        resolve.addEventListener("click", () => {
          this.host.setCommentResolved(root.anchorId, !root.resolved);
          if (!root.resolved) this.expanded.delete(root.anchorId);
          this.setActive(root.anchorId);
        });
        bar.append(reply, resolve);
        bubble.appendChild(bar);
      }
    }

    bubble.addEventListener("mousedown", (event) => {
      // Buttons and inputs must keep working; a click on the bubble body just activates it.
      const t = event.target as HTMLElement;
      if (t.closest("button, textarea, input, a")) return;
      if (collapsed) { this.expanded.add(root.anchorId); }
      this.setActive(root.anchorId, { scrollAnchor: true });
    });
    return bubble;
  }

  private buildEntry(entry: CommentListEntry, doc: Document, isRoot: boolean): HTMLElement {
    const el = doc.createElement("div");
    el.className = isRoot ? "docx-comment-entry docx-comment-root" : "docx-comment-entry docx-comment-reply";
    el.dataset.comment = entry.anchorId;
    const head = doc.createElement("div");
    head.className = "docx-comment-head";
    const avatar = doc.createElement("span");
    avatar.className = "docx-comment-avatar";
    avatar.textContent = initialsOf(entry.author || "", entry.initials);
    avatar.style.setProperty("--docx-comment-hue", String(authorHue(entry.author || "")));
    const author = doc.createElement("span");
    author.className = "docx-comment-author";
    author.textContent = entry.author || "Unknown";
    const date = doc.createElement("span");
    date.className = "docx-comment-date";
    date.textContent = formatDate(entry.date);
    head.append(avatar, author, date);
    if (isRoot && entry.resolved) {
      const badge = doc.createElement("span");
      badge.className = "docx-comment-badge";
      badge.textContent = "Resolved";
      head.appendChild(badge);
    }
    const menu = doc.createElement("span");
    menu.className = "docx-comment-menu";
    const edit = doc.createElement("button");
    edit.type = "button";
    edit.title = "Edit";
    edit.setAttribute("aria-label", "Edit comment");
    edit.setAttribute("data-comment-action", "edit");
    edit.innerHTML = "&#9998;";
    edit.addEventListener("click", () => { this.editing = entry.anchorId; this.replying = null; this.layout(); });
    const del = doc.createElement("button");
    del.type = "button";
    del.title = isRoot ? "Delete thread" : "Delete reply";
    del.setAttribute("aria-label", del.title);
    del.setAttribute("data-comment-action", "delete");
    del.innerHTML = "&#128465;";
    del.addEventListener("click", () => {
      if (isRoot) {
        // A deleted root would orphan its replies into top-level comments; take the thread
        // down, replies to replies included.
        for (const id of threadMembers(this.host.listComments(), entry.anchorId)) this.host.removeComment(id);
      }
      this.host.removeComment(entry.anchorId);
      if (this.activeId === entry.anchorId) this.activeId = null;
      this.schedule();
    });
    menu.append(edit, del);
    head.appendChild(menu);
    el.appendChild(head);

    if (this.editing === entry.anchorId) {
      const input = doc.createElement("textarea");
      input.className = "docx-comment-input";
      input.rows = 3;
      input.value = entry.text;
      input.setAttribute("data-comment-edit-text", "");
      const actions = doc.createElement("div");
      actions.className = "docx-comment-actions";
      const save = doc.createElement("button");
      save.type = "button";
      save.className = "docx-comment-primary";
      save.textContent = "Save";
      save.setAttribute("data-comment-action", "save");
      const cancel = doc.createElement("button");
      cancel.type = "button";
      cancel.textContent = "Cancel";
      cancel.setAttribute("data-comment-action", "cancel-edit");
      const submit = () => {
        const text = input.value.trim();
        this.editing = null;
        if (text && text !== entry.text) this.host.updateComment(entry.anchorId, text);
        this.schedule();
      };
      save.addEventListener("click", submit);
      cancel.addEventListener("click", () => { this.editing = null; this.schedule(); });
      input.addEventListener("keydown", (event) => {
        const key = event as KeyboardEvent;
        if (key.key === "Escape") { key.preventDefault(); this.editing = null; this.schedule(); }
        if (key.key === "Enter" && (key.ctrlKey || key.metaKey)) { key.preventDefault(); submit(); }
      });
      actions.append(save, cancel);
      el.append(input, actions);
      queueMicrotask(() => { input.focus(); input.setSelectionRange(input.value.length, input.value.length); });
    } else {
      const body = doc.createElement("div");
      body.className = "docx-comment-text";
      body.textContent = entry.text;
      el.appendChild(body);
    }
    return el;
  }
}

/** The gutter's stylesheet — injected once per document by the editor. Scoped to the classes
 *  this module emits, so it is safe on a host page. Colours derive from a per-author hue so a
 *  reviewer's highlight and bubble read as one. */
export const COMMENT_GUTTER_CSS = `
.docx-comment-gutter {
  position: absolute; top: 0; right: 0; bottom: 0; z-index: 4;
  pointer-events: none; font: 12.5px/1.4 "Inter", system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
}
.docx-comment-gutter[data-empty] { display: none; }
.docx-comment-leaders { position: absolute; top: 0; left: 0; z-index: 3; pointer-events: none; overflow: visible; }
.docx-comment-leader { fill: none; stroke: hsl(var(--docx-comment-hue, 200) 45% 62%); stroke-width: 1; stroke-dasharray: 3 3; opacity: .55; }
.docx-comment-leader-active { stroke-dasharray: none; opacity: .95; stroke-width: 1.5; }
.docx-comment-bubble {
  position: absolute; left: 12px; right: 8px; pointer-events: auto;
  padding: 9px 10px 8px; border: 1px solid hsl(var(--docx-comment-hue, 200) 40% 78%);
  border-left: 3px solid hsl(var(--docx-comment-hue, 200) 55% 55%);
  border-radius: 8px; background: #fff; color: #1e293b;
  box-shadow: 0 1px 2px rgba(15, 23, 42, .06), 0 3px 8px rgba(15, 23, 42, .05);
  transition: top .18s cubic-bezier(.4,0,.2,1), box-shadow .15s ease, transform .15s ease;
}
.docx-comment-bubble[data-active] {
  border-color: hsl(var(--docx-comment-hue, 200) 55% 55%);
  box-shadow: 0 2px 6px rgba(15, 23, 42, .08), 0 10px 22px rgba(15, 23, 42, .1);
  transform: translateX(-4px); z-index: 2;
}
.docx-comment-bubble[data-resolved] { opacity: .72; background: #f8fafc; }
.docx-comment-bubble[data-collapsed] .docx-comment-text { display: -webkit-box; -webkit-line-clamp: 1; -webkit-box-orient: vertical; overflow: hidden; color: #64748b; }
.docx-comment-bubble[data-collapsed] .docx-comment-reply, .docx-comment-bubble[data-collapsed] .docx-comment-bar { display: none; }
.docx-comment-bubble[data-orphan] { border-style: dashed; }
.docx-comment-bubble[data-draft] { border-style: solid; }
.docx-comment-entry + .docx-comment-entry { margin-top: 8px; padding-top: 8px; border-top: 1px solid #f1f5f9; }
.docx-comment-reply { margin-left: 10px; }
.docx-comment-head { display: flex; align-items: center; gap: 6px; min-height: 20px; }
.docx-comment-avatar {
  display: inline-grid; place-items: center; flex: 0 0 auto; width: 20px; height: 20px; border-radius: 50%;
  background: hsl(var(--docx-comment-hue, 200) 55% 45%); color: #fff; font-size: 9px; font-weight: 700; letter-spacing: .02em;
}
.docx-comment-author { font-weight: 600; font-size: 12px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.docx-comment-date { color: #94a3b8; font-size: 10.5px; white-space: nowrap; margin-left: auto; }
.docx-comment-badge { padding: 1px 6px; border-radius: 999px; background: #dcfce7; color: #166534; font-size: 9.5px; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; }
.docx-comment-menu { display: none; gap: 2px; margin-left: 4px; }
.docx-comment-bubble:hover .docx-comment-menu, .docx-comment-bubble[data-active] .docx-comment-menu { display: inline-flex; }
.docx-comment-menu button {
  padding: 0 4px; min-height: 18px; border: 0; border-radius: 4px; background: none; color: #94a3b8; font-size: 12px; line-height: 1; cursor: pointer;
}
.docx-comment-menu button:hover { background: #f1f5f9; color: #1e293b; }
.docx-comment-text { margin: 5px 0 0; white-space: pre-wrap; word-wrap: break-word; color: #334155; }
.docx-comment-bar { display: flex; gap: 6px; margin-top: 8px; }
.docx-comment-bar button, .docx-comment-actions button {
  min-height: 24px; padding: 2px 9px; border: 1px solid #e2e8f0; border-radius: 6px; background: #fff; color: #1e293b;
  font: inherit; font-size: 11.5px; font-weight: 500; cursor: pointer;
}
.docx-comment-bar button:hover, .docx-comment-actions button:hover { background: #f1f5f9; }
.docx-comment-actions { display: flex; gap: 6px; margin-top: 6px; }
.docx-comment-actions .docx-comment-primary { background: #0f766e; border-color: #0f766e; color: #fff; }
.docx-comment-actions .docx-comment-primary:hover { background: #0d9488; }
.docx-comment-input {
  display: block; width: 100%; margin-top: 6px; padding: 6px 8px; border: 1px solid #cbd5e1; border-radius: 6px;
  background: #fff; color: #1e293b; font: inherit; font-size: 12.5px; line-height: 1.4; resize: vertical; box-sizing: border-box;
}
.docx-comment-input:focus { outline: 2px solid #0f766e; outline-offset: 1px; border-color: #0f766e; }
/* The commented range itself. The converter's stylesheet paints a flat yellow; the editor
   tints by author and lifts the active thread the way Word does. */
span.comment-highlight[data-comment-id] {
  background: hsl(var(--docx-comment-hue, 45) 90% 88%);
  border-bottom: 2px solid hsl(var(--docx-comment-hue, 45) 70% 55%);
  cursor: pointer;
}
span.comment-highlight[data-comment-id][data-comment-resolved] { background: transparent; border-bottom-color: #cbd5e1; }
span.comment-highlight.docx-comment-active { background: hsl(var(--docx-comment-hue, 45) 90% 78%); }
/* The [n] reference marker is engine chrome, not document text: hide it, keep the anchor.
   (Qualified by the editor's roots to outrank the converter's own a.comment-marker rule,
   which is injected later in the document.) */
.docx-body-flow a.comment-marker, [data-hf-band] a.comment-marker, .page-box a.comment-marker { display: none; }
[data-comments-hidden] span.comment-highlight[data-comment-id] { background: transparent; border-bottom-color: transparent; }
`;

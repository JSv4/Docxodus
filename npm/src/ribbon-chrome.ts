/**
 * The ribbon surface's markup and stylesheet — the two static assets `ribbon.ts`
 * wires behaviour onto.
 *
 * They live in their own module because they are *data*: one HTML template and
 * one stylesheet, both authored to be read as a unit. Keeping them out of
 * `ribbon.ts` leaves that file as pure wiring.
 *
 * ── Addressing ──────────────────────────────────────────────────────────────
 * Every control carries `data-dxr="<name>"`. `ribbon.ts` walks those on mount
 * and stamps `id = idPrefix + name`, so the historical element ids from the
 * standalone demo page survive (specs and bookmarklets keep working) while a
 * second ribbon on the same page can take a prefix and stay collision-free.
 * `data-dxr-list` does the same job for the `<input list>` → `<datalist id>`
 * link, which is by-id and cannot use a data attribute directly.
 *
 * ── Responsiveness ──────────────────────────────────────────────────────────
 * Layout keys off `data-chrome` on the root ("full" | "compact"), which
 * `ribbon.ts` sets from a ResizeObserver on the ROOT rather than from viewport
 * media queries. A narrow embed inside a wide desktop page is narrow, and only
 * container measurement gets that right. Compact is the base (mobile-first);
 * `[data-chrome="full"]` adds back the desktop affordances.
 */

/** Bumped whenever RIBBON_CSS changes, so a stale injected stylesheet is replaced. */
export const RIBBON_STYLE_VERSION = "8";
export const RIBBON_STYLE_ATTR = "data-docxodus-ribbon-styles";

/**
 * The ribbon stylesheet. Every rule is scoped under `.dxr` so injecting it is
 * safe on a host page we do not control — the same guarantee `embed.ts` gives
 * for the converter's document CSS, achieved here by authoring rather than by
 * rewriting.
 */
export const RIBBON_CSS = `
.dxr {
  /* OS-Legal design tokens (github.com/Open-Source-Legal/OS-Legal-Style):
     deep-teal accent, slate foreground scale, warm neutral surfaces, and
     "light touch" shadows at 0.03-0.06 opacity — structure over decoration. */
  --dxr-ink: #1e293b;
  --dxr-sheet: #ffffff;
  --dxr-chrome: #ffffff;
  --dxr-chrome-sunk: #f8fafc;
  --dxr-rule: #e2e8f0;
  --dxr-rule-strong: #cbd5e1;
  --dxr-accent: #0f766e;
  --dxr-accent-hover: #0d9488;
  --dxr-accent-active: #115e59;
  --dxr-wash: #f0fdfa;
  --dxr-wash-on: #ccfbf1;
  --dxr-muted: #475569;
  --dxr-faint: #94a3b8;
  --dxr-data: #0f766e;
  --dxr-danger: #dc2626;
  --dxr-danger-wash: #fef2f2;
  --dxr-desk: #f1f5f9;
  --dxr-ui: "Inter", system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  --dxr-serif: Georgia, "Times New Roman", serif;
  --dxr-mono: ui-monospace, SFMono-Regular, "Cascadia Mono", Consolas, monospace;
  --dxr-shadow-sm: 0 1px 2px rgba(15, 23, 42, .04);
  --dxr-shadow-md: 0 1px 3px rgba(15, 23, 42, .06), 0 4px 6px rgba(15, 23, 42, .03);
  --dxr-shadow-lg: 0 4px 6px rgba(15, 23, 42, .05), 0 10px 15px rgba(15, 23, 42, .04);
  --dxr-shadow-xl: 0 8px 16px rgba(15, 23, 42, .06), 0 20px 25px rgba(15, 23, 42, .05);
  --dxr-ease: cubic-bezier(.4, 0, .2, 1);
  --dxr-tap: 30px;

  position: relative;
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
  isolation: isolate;
  font-family: var(--dxr-ui);
  color: var(--dxr-ink);
  background: var(--dxr-desk);
  -webkit-text-size-adjust: 100%;
}
.dxr *, .dxr *::before, .dxr *::after { box-sizing: border-box; }

/* ── Card frame ────────────────────────────────────────────────────────────────
   A drop-in embed carries its OWN boundary instead of hoping the host clips it:
   rounded corners, a hairline, and the house elevation. overflow: clip keeps
   the sticky chrome, scrollbars, and the loading overlay inside the curve.
   Full-bleed hosts mount with data-frame="flush" and stay edge-to-edge. */
.dxr[data-frame="card"] {
  border: 1px solid var(--dxr-rule);
  border-radius: 14px;
  box-shadow: var(--dxr-shadow-xl);
  overflow: clip;
}

/* Touch devices get bigger hit targets everywhere the token is used. */
@media (pointer: coarse) { .dxr { --dxr-tap: 40px; } }

/* ── Chrome shell ─────────────────────────────────────────────────────────── */
.dxr-chrome {
  position: sticky;
  top: 0;
  z-index: 10;
  flex: 0 0 auto;
  background: var(--dxr-chrome);
  border-bottom: 1px solid var(--dxr-rule);
}

.dxr-titlebar {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 6px 10px 0;
}
.dxr-brand {
  display: flex;
  align-items: baseline;
  gap: 7px;
  min-width: 0;
  font-size: 13px;
  font-weight: 600;
  letter-spacing: -.01em;
}
/* A filled circle, echoing the OS-Legal "Node" motif rather than a diamond. */
.dxr-brand .dxr-mark {
  flex: 0 0 auto;
  width: 9px;
  height: 9px;
  border-radius: 50%;
  background: var(--dxr-accent);
}
.dxr-brand .dxr-docname {
  overflow: hidden;
  max-width: 22ch;
  color: var(--dxr-muted);
  font-size: 12px;
  font-weight: 400;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.dxr-quick { flex: 0 0 auto; display: flex; align-items: center; gap: 3px; }
.dxr-titlebar .dxr-spacer { flex: 1; }
.dxr-status {
  overflow: hidden;
  max-width: 42ch;
  color: var(--dxr-muted);
  font-family: var(--dxr-mono);
  font-size: 11.5px;
  text-overflow: ellipsis;
  white-space: nowrap;
}

/* ── Tab strip ────────────────────────────────────────────────────────────── */
.dxr-tabs {
  display: flex;
  gap: 1px;
  padding: 6px 10px 0;
  overflow-x: auto;
  scrollbar-width: none;
}
.dxr-tabs::-webkit-scrollbar { display: none; }
/* Line-variant tabs (the house Tabs default): no boxes, no fills — the selected
   tab is a 2px accent underline sitting on the panel's own hairline. */
.dxr-tab {
  flex: 0 0 auto;
  padding: 7px 14px 8px;
  border: none;
  background: none;
  color: var(--dxr-muted);
  font: inherit;
  font-size: 12.5px;
  font-weight: 500;
  cursor: pointer;
  transition: color .15s var(--dxr-ease), box-shadow .15s var(--dxr-ease);
}
.dxr-tab:hover { color: var(--dxr-ink); }
.dxr-tab[aria-selected="true"] {
  color: var(--dxr-ink);
  font-weight: 600;
  box-shadow: inset 0 -2px 0 var(--dxr-accent);
}
/* Contextual tab — present only while the caret is inside a table. */
.dxr-tab[data-contextual] { color: var(--dxr-accent); }
.dxr-tab[hidden] { display: none; }

/* ── Ribbon panels ────────────────────────────────────────────────────────── */
.dxr-ribbon { background: var(--dxr-chrome-sunk); border-top: 1px solid var(--dxr-rule); }
.dxr-ribbon[aria-disabled="true"] { opacity: .45; pointer-events: none; }
.dxr-panel {
  display: none;
  align-items: stretch;
  padding: 7px 10px 8px;
  overflow-x: auto;
  overscroll-behavior-x: contain;
}
.dxr-panel[data-active] { display: flex; }
/* Group label sits ABOVE its controls (a spec-sheet reading), with hairline
   dividers rather than boxes — lighter than the boxed, label-under convention. */
.dxr-group {
  flex: 0 0 auto;
  display: flex;
  flex-direction: column;
  gap: 5px;
  padding: 0 13px;
  border-right: 1px solid var(--dxr-rule);
}
.dxr-group:last-child { border-right: none; }
.dxr-group:first-child { padding-left: 0; }
.dxr-glabel {
  color: var(--dxr-faint);
  font-size: 9.5px;
  font-weight: 600;
  letter-spacing: .11em;
  text-transform: uppercase;
}
.dxr-row { display: flex; gap: 3px; align-items: center; }

/* ── Controls ─────────────────────────────────────────────────────────────── */
/* One button family, house-style: every control is a ghost — no borders, no
   fills, no per-button boxes. Hover paints a quiet slate wash; the active
   (toggled) state paints a teal one. Save alone is the flat primary, so the
   surface has exactly one accent-colored action. */
.dxr button, .dxr label.dxr-btn {
  min-height: var(--dxr-tap);
  padding: 5px 9px;
  border: none;
  border-radius: 6px;
  background: none;
  color: var(--dxr-ink);
  font: inherit;
  font-size: 13px;
  font-weight: 500;
  line-height: 1.15;
  cursor: pointer;
  transition: background-color .15s var(--dxr-ease), color .15s var(--dxr-ease),
    box-shadow .15s var(--dxr-ease), transform .15s var(--dxr-ease);
}
.dxr label.dxr-btn { display: inline-flex; align-items: center; }
.dxr button:hover, .dxr label.dxr-btn:hover { background: #f1f5f9; }
.dxr button:active { background: #e2e8f0; }
.dxr button:disabled { opacity: .38; cursor: default; background: none; }
.dxr button.dxr-on { color: var(--dxr-accent); background: var(--dxr-wash-on); }
.dxr-quick button, .dxr-quick label.dxr-btn { padding: 4px 10px; font-size: 12.5px; }
/* Save — the house primary: flat accent, sm shadow, the 1px lift on hover. */
.dxr-quick button[data-dxr="save"]:not(:disabled) {
  background: var(--dxr-accent);
  color: #fff;
  font-weight: 600;
  box-shadow: var(--dxr-shadow-sm);
}
.dxr-quick button[data-dxr="save"]:not(:disabled):hover {
  background: var(--dxr-accent-hover);
  box-shadow: var(--dxr-shadow-md);
  transform: translateY(-1px);
}
.dxr-quick button[data-dxr="save"]:not(:disabled):active {
  background: var(--dxr-accent-active);
  box-shadow: var(--dxr-shadow-sm);
  transform: translateY(0);
}
/* Icon controls sit quieter than text ones until touched (house IconButton). */
.dxr button.dxr-icon { min-width: var(--dxr-tap); text-align: center; color: var(--dxr-muted); }
.dxr button.dxr-icon:hover:not(:disabled) { color: var(--dxr-ink); }
.dxr button.dxr-icon.dxr-on { color: var(--dxr-accent); }
/* Icons are inline SVG drawn from the same shapes the control acts on (text lines,
   rules), so alignment and indent read at a glance instead of relying on arrow
   glyphs that collide with undo/redo. currentColor keeps them in step with state. */
.dxr button svg { display: block; margin: 0 auto; fill: currentColor; }
.dxr button.dxr-wide { display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
.dxr button.dxr-danger:hover { color: var(--dxr-danger); background: var(--dxr-danger-wash); }
.dxr select, .dxr input[type="number"], .dxr input[type="text"] {
  min-height: var(--dxr-tap);
  padding: 4px 6px;
  border: 1px solid var(--dxr-rule);
  border-radius: 6px;
  background: var(--dxr-sheet);
  color: var(--dxr-ink);
  font: inherit;
  font-size: 12.5px;
}
.dxr select:hover, .dxr input:hover { border-color: var(--dxr-rule-strong); }
.dxr select:focus-visible, .dxr input:focus-visible, .dxr button:focus-visible,
.dxr .dxr-tab:focus-visible, .dxr label.dxr-btn:focus-within {
  outline: 2px solid var(--dxr-accent);
  outline-offset: 1px;
}
.dxr [data-dxr="fontsize"] { width: 62px; }
.dxr [data-dxr="fontfamily"] { max-width: 132px; }
.dxr [data-dxr="pgstart"] { width: 64px; }
.dxr-toggle {
  display: inline-flex;
  gap: 5px;
  align-items: center;
  color: var(--dxr-ink);
  font-size: 12.5px;
  white-space: nowrap;
  cursor: pointer;
}
.dxr-note { color: var(--dxr-muted); font-size: 11.5px; }

/* ── Anchor rail ──────────────────────────────────────────────────────────────
   The addressing spine made permanent chrome. Every block in this editor is
   addressable as kind:scope:unid, and the model of record is a live WASM session —
   so the surface states both, live, instead of hiding them behind devtools. It also
   reports each command's real cost, which is what a smoke test needs to see. */
.dxr-rail {
  display: flex;
  align-items: center;
  height: 25px;
  padding: 0 10px;
  overflow-x: auto;
  background: #fafafa;
  border-top: 1px solid var(--dxr-rule);
  color: var(--dxr-muted);
  font-family: var(--dxr-mono);
  font-size: 11.5px;
  scrollbar-width: none;
}
.dxr-rail::-webkit-scrollbar { display: none; }
.dxr-rail .dxr-cell {
  display: flex;
  align-items: baseline;
  gap: 6px;
  padding: 0 13px;
  border-right: 1px solid var(--dxr-rule);
  white-space: nowrap;
}
.dxr-rail .dxr-cell:first-child { padding-left: 0; }
.dxr-rail .dxr-cell:last-child { border-right: none; }
.dxr-rail .dxr-k {
  color: var(--dxr-faint);
  font-family: var(--dxr-ui);
  font-size: 9.5px;
  letter-spacing: .1em;
  text-transform: uppercase;
}
.dxr-rail .dxr-v { color: var(--dxr-data); }
.dxr-rail .dxr-v.dxr-flash { animation: dxr-railflash .45s ease-out; }
@keyframes dxr-railflash { from { background: #ccfbf1; } to { background: transparent; } }

/* ── Hint ─────────────────────────────────────────────────────────────────── */
.dxr-hint {
  flex: 0 0 auto;
  max-width: 920px;
  margin: 0 auto;
  padding: 9px 16px 0;
  color: var(--dxr-muted);
  font-size: 12px;
  line-height: 1.5;
}
.dxr-hint kbd {
  padding: 0 4px;
  border: 1px solid var(--dxr-rule);
  border-radius: 4px;
  background: #f1f5f9;
  font-family: var(--dxr-mono);
  font-size: 11px;
}

/* ── Document surface ──────────────────────────────────────────────────────────
   The surface is also the element the converter's own stylesheet treats as the
   document body: in an embed, its "body { margin: 20px }" is rewritten to
   [data-docxodus-embed-root="dN"] [data-dxr-surface] — (0,2,0), and inserted AFTER
   this sheet, so it wins a tie. Rules that own the SHEET BOX (centering, page
   padding, the paper itself) therefore qualify with the root's data-chrome to reach
   (0,3,0). Rules that only style CONTENT stay unqualified — there the converter
   SHOULD win.
   (No backticks in this file's comments: the stylesheet is a template literal.) */
.dxr-scroll { flex: 1 1 auto; min-height: 0; overflow: auto; -webkit-overflow-scrolling: touch; }
/* Soft scroll boundaries: content dissolves at the chrome and at the bottom edge
   instead of hard-clipping. Sticky gradient veils cost one paint layer each and
   never intercept input. */
.dxr-scroll::before, .dxr-scroll::after {
  content: "";
  position: sticky;
  z-index: 3;
  display: block;
  height: 26px;
  pointer-events: none;
}
.dxr-scroll::before {
  top: 0;
  margin-bottom: -26px;
  background: linear-gradient(to bottom, var(--dxr-desk), transparent);
}
.dxr-scroll::after {
  bottom: 0;
  margin-top: -26px;
  background: linear-gradient(to top, var(--dxr-desk), transparent);
}
.dxr[data-chrome] .dxr-surface { margin: 26px auto; padding: 0 16px 96px; }
.dxr-surface [contenteditable="true"]:focus {
  outline: 2px solid var(--dxr-accent);
  outline-offset: 2px;
  border-radius: 2px;
}
@media (hover: hover) { .dxr-surface [contenteditable="true"]:hover { background: var(--dxr-wash); } }

/* Header/footer bands: stories live in their own OOXML parts outside the body,
   so they dock as their own regions. Styled from the same tokens as the ribbon
   so the surface reads as one instrument rather than two apps. */
.dxr-surface .docx-hf-band {
  /* Docked outside the zoomed sheet, so it takes the page's on-screen width from the
     custom property the viewport publishes rather than stretching to the whole surface. */
  width: min(100%, var(--docx-sheet-width, 100%));
  margin: 0 auto 14px;
  padding: 9px 72px 13px;
  border: 1px solid var(--dxr-rule);
  border-left: 2px solid #5eead4;
  border-radius: 6px;
  background: var(--dxr-sheet);
  box-shadow: var(--dxr-shadow-sm);
}
.dxr-surface .docx-hf-band + .docx-body-flow { margin-top: 0; }
.dxr-surface .docx-hf-band[data-hf-band="footer"] { margin: 14px auto 0; }
.dxr-surface .docx-hf-chrome {
  display: flex;
  gap: 8px;
  align-items: center;
  flex-wrap: wrap;
  margin: 0 0 7px -60px;
  color: var(--dxr-muted);
  font-size: 9.5px;
  font-weight: 600;
  letter-spacing: .11em;
  text-transform: uppercase;
}
.dxr-surface .docx-hf-chrome select, .dxr-surface .docx-hf-chrome input {
  padding: 3px 5px;
  border: 1px solid var(--dxr-rule);
  border-radius: 6px;
  background: var(--dxr-sheet);
  color: var(--dxr-ink);
  font: inherit;
  font-family: var(--dxr-ui);
  font-size: 12.5px;
  font-weight: 400;
  letter-spacing: normal;
  text-transform: none;
}
.dxr-surface .docx-hf-chrome input[data-hf-pagestart] { width: 64px; }
.dxr-surface .docx-hf-label { font-weight: 600; }
.dxr-surface .docx-hf-warning {
  margin: 0 0 8px -60px;
  padding: 7px 9px;
  border: 1px solid #fde68a;
  border-left: 2px solid #d97706;
  border-radius: 6px;
  background: #fffbeb;
  color: #92400e;
  font-size: 12px;
  line-height: 1.45;
}
.dxr-surface .docx-hf-warning button {
  margin-left: 6px;
  padding: 2px 8px;
  border-color: #fde68a;
  background: #fff;
  font-size: 12px;
}
.dxr-surface .docx-hf-placeholder { color: var(--dxr-faint); font-style: italic; }
.dxr-surface .docx-hf-inherited {
  color: var(--dxr-muted);
  font-style: italic;
  font-weight: 400;
  letter-spacing: normal;
  text-transform: none;
}
.dxr-surface .docx-hf-band[data-hf-inherited] { border-style: dashed; }
/* The sheet IS the page. The editor's viewport sizes .docx-body-flow to the section's page
   width and its section wrappers to the authored text column, so the horizontal gutters here
   are the document's own w:sectPr margins, not a padding this chrome invents; only the
   vertical breathing room is ours. Centering is left to margin:auto so a page the viewport
   has zoomed to fit stays centered at its scaled width. */
.dxr[data-chrome] .dxr-surface[data-view="continuous"] .docx-body-flow {
  max-width: 100%;
  margin: 0 auto;
  padding: 56px 0;
  border-radius: 3px;
  background: var(--dxr-sheet);
  box-shadow: 0 1px 3px rgba(15, 23, 42, .08), 0 12px 24px rgba(15, 23, 42, .06);
}

/* ── Table size picker ────────────────────────────────────────────────────── */
.dxr-pop {
  display: none;
  position: absolute;
  z-index: 30;
  padding: 9px;
  border: 1px solid var(--dxr-rule);
  border-radius: 12px;
  background: #fff;
  box-shadow: var(--dxr-shadow-xl);
}
.dxr-pop[data-open] { display: block; }
.dxr-gridcells { display: grid; grid-template-columns: repeat(10, 16px); gap: 2px; }
.dxr-gridcells div {
  width: 16px;
  height: 16px;
  border: 1px solid var(--dxr-rule);
  border-radius: 3px;
  background: var(--dxr-chrome-sunk);
  cursor: pointer;
}
.dxr-gridcells div[data-on] { background: #ccfbf1; border-color: var(--dxr-accent); }
.dxr-popfoot {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 8px;
  font-size: 12px;
}

/* ── Compact chrome (the mobile-first base) ───────────────────────────────────
   One scrolling strip per tab instead of a multi-row ribbon: group labels turn
   into inline dividers, the rail and hint step aside, and the picker becomes a
   bottom sheet because there is no room to hang a popover off a button. */
/* A 320px phone cannot hold the product name AND the file actions, and the file
   actions are what a user came for. The strip also scrolls, so nothing is stranded. */
.dxr[data-chrome="compact"] .dxr-titlebar {
  gap: 6px;
  padding: 5px 8px 0;
  overflow-x: auto;
  scrollbar-width: none;
}
.dxr[data-chrome="compact"] .dxr-titlebar::-webkit-scrollbar { display: none; }
.dxr[data-chrome="compact"] .dxr-brandname { display: none; }
.dxr[data-chrome="compact"] .dxr-brand { flex: 0 1 auto; }
/* A filename clipped to "do..." tells you less than no filename at all, so it keeps a
   readable floor and the strip scrolls instead. The tighter file-action padding is what
   buys that floor back on a 390px screen. */
.dxr[data-chrome="compact"] .dxr-brand .dxr-docname { min-width: 6ch; max-width: 14ch; }
.dxr[data-chrome="compact"] .dxr-quick button,
.dxr[data-chrome="compact"] .dxr-quick label.dxr-btn { padding: 4px 8px; }
.dxr[data-chrome="compact"] .dxr-status { display: none; }
.dxr[data-chrome="compact"] .dxr-tabs { padding: 5px 8px 0; }
.dxr[data-chrome="compact"] .dxr-tab { padding: 6px 12px 7px; font-size: 12px; }
.dxr[data-chrome="compact"] .dxr-panel {
  align-items: center;
  gap: 4px;
  padding: 5px 8px;
  scroll-snap-type: x proximity;
}
.dxr[data-chrome="compact"] .dxr-group {
  flex-direction: row;
  align-items: center;
  gap: 4px;
  padding: 0 8px;
  scroll-snap-align: start;
}
.dxr[data-chrome="compact"] .dxr-glabel { display: none; }
.dxr[data-chrome="compact"] .dxr-note { display: none; }
.dxr[data-chrome="compact"] .dxr-rail { display: none; }
.dxr[data-chrome="compact"] .dxr-hint { display: none; }
/* Compact trims the chrome around the page, never the page: the document's own column width
   is what the viewport's fit-to-width zoom scales, so a phone shows a whole smaller page
   instead of a narrower one that breaks its lines somewhere Word never would. */
.dxr[data-chrome="compact"] .dxr-surface { margin: 12px auto; padding: 0 10px 64px; }
.dxr[data-chrome="compact"] .dxr-surface[data-view="continuous"] .docx-body-flow { padding: 22px 0; }
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-band { padding: 9px 18px 13px; }
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-chrome,
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-warning { margin-left: 0; }
/* A popover anchored to a button has nowhere to go on a narrow surface, so the
   picker docks to the bottom edge where a thumb already is. */
.dxr[data-chrome="compact"] .dxr-pop[data-open] {
  position: fixed;
  left: 50%;
  right: auto;
  bottom: 12px;
  top: auto !important;
  transform: translateX(-50%);
  max-width: calc(100vw - 20px);
}

/* ── Loading overlay ──────────────────────────────────────────────────────────
   Kept from the shipped demo: the wait is real (a .NET runtime is streaming), so
   the surface spends it explaining what is being built rather than showing a
   spinner. It covers the whole instrument so half-built chrome never flashes. */
.dxr-loader {
  position: absolute;
  inset: 0;
  z-index: 40;
  display: grid;
  place-items: center;
  padding: 24px;
  color: #f8fafc;
  background:
    radial-gradient(circle at 20% 20%, rgba(13, 148, 136, .2), transparent 22rem),
    radial-gradient(circle at 82% 80%, rgba(45, 212, 191, .1), transparent 24rem),
    linear-gradient(150deg, #0b1220 0%, #0f172a 55%, #0f2723 100%);
  transition: opacity .55s ease, visibility .55s ease;
}
.dxr-loader[hidden] { display: none; }
/* pointer-events drops the instant the fade starts: the surface underneath is already
   live, so a click during the half-second fade should reach it rather than be eaten. */
.dxr-loader[data-done] { opacity: 0; visibility: hidden; pointer-events: none; }
.dxr-loader-grid {
  width: min(850px, 100%);
  display: grid;
  grid-template-columns: 1fr;
  align-items: center;
  gap: 22px;
  text-align: center;
}
.dxr-visual { position: relative; width: min(170px, 52vw); aspect-ratio: 1; margin: 0 auto; }
.dxr-orbit {
  position: absolute;
  inset: 7%;
  border: 1px solid rgba(45, 212, 191, .24);
  border-radius: 50%;
  animation: dxr-spin 9s linear infinite;
}
.dxr-orbit.dxr-two {
  inset: 20%;
  border-style: dashed;
  border-color: rgba(94, 234, 212, .3);
  animation-duration: 6s;
  animation-direction: reverse;
}
.dxr-orbit::before, .dxr-orbit::after {
  position: absolute;
  width: 10px;
  height: 10px;
  border-radius: 50%;
  content: "";
  background: #2dd4bf;
  box-shadow: 0 0 18px rgba(45, 212, 191, .8);
}
.dxr-orbit::before { top: -5px; left: 50%; }
.dxr-orbit::after { right: 7%; bottom: 12%; background: #5eead4; box-shadow: 0 0 18px rgba(94, 234, 212, .7); }
.dxr-card {
  position: absolute;
  inset: 24% 28%;
  padding: 20px 15px;
  border: 1px solid rgba(255, 255, 255, .3);
  border-radius: 14px;
  background: linear-gradient(150deg, rgba(255, 255, 255, .16), rgba(255, 255, 255, .05));
  box-shadow: 0 22px 50px rgba(0, 0, 0, .35), 0 0 45px rgba(13, 148, 136, .15);
  backdrop-filter: blur(10px);
  animation: dxr-float 3.6s ease-in-out infinite;
}
.dxr-card::before {
  position: absolute;
  top: -9px;
  right: -9px;
  padding: 4px 7px;
  border-radius: 6px;
  background: #2dd4bf;
  color: #042f2e;
  content: "DOCX";
  font-size: 8px;
  font-weight: 900;
  letter-spacing: .12em;
}
.dxr-card i {
  display: block;
  height: 4px;
  margin-bottom: 10px;
  border-radius: 4px;
  background: rgba(255, 255, 255, .56);
  transform-origin: left;
  animation: dxr-pulse 2.2s ease-in-out infinite;
}
.dxr-card i:nth-child(2) { width: 76%; animation-delay: -.45s; }
.dxr-card i:nth-child(3) { width: 88%; animation-delay: -.85s; }
.dxr-card i:nth-child(4) { width: 58%; background: rgba(248, 113, 113, .75); animation-delay: -1.2s; }
.dxr-chip {
  position: absolute;
  display: grid;
  width: 34px;
  height: 34px;
  place-items: center;
  border: 1px solid rgba(255, 255, 255, .16);
  border-radius: 11px;
  background: rgba(15, 30, 46, .88);
  box-shadow: 0 12px 30px rgba(0, 0, 0, .28);
  font-size: 11px;
  font-weight: 850;
}
.dxr-chip.dxr-b { left: 2%; top: 30%; color: #5eead4; animation: dxr-chip 3s ease-in-out infinite; }
.dxr-chip.dxr-r { right: -2%; bottom: 24%; color: #fca5a5; animation: dxr-chip 3s -1.4s ease-in-out infinite; }
.dxr-chip.dxr-s { left: 20%; bottom: 0; color: #6ee7b7; animation: dxr-chip 3s -.7s ease-in-out infinite; }
.dxr-eyebrow { color: #5eead4; font-size: 10px; font-weight: 850; letter-spacing: .16em; text-transform: uppercase; }
/* Georgia headline — the house pairing of a serif headline over Inter UI copy. */
.dxr-loader h2 {
  margin: 10px auto 10px;
  max-width: 520px;
  font-family: var(--dxr-serif);
  font-size: clamp(23px, 5vw, 38px);
  font-weight: 500;
  line-height: 1.08;
  letter-spacing: -.02em;
}
.dxr-loader-copy > p { margin: 0 auto; max-width: 46ch; color: #94a3b8; font-size: 13.5px; line-height: 1.6; }
.dxr-ad {
  display: flex;
  gap: 12px;
  margin: 20px auto 0;
  max-width: 420px;
  padding: 13px;
  border: 1px solid rgba(148, 163, 184, .16);
  border-radius: 12px;
  background: rgba(255, 255, 255, .04);
  text-align: left;
}
.dxr-ad .dxr-num { color: #2dd4bf; font: 800 10px/1.4 var(--dxr-mono); }
.dxr-ad strong { display: block; font-size: 12.5px; }
.dxr-ad .dxr-adcopy { display: block; margin-top: 4px; color: #8ba0b9; font-size: 11.5px; line-height: 1.45; }
.dxr-ad[data-swap] { animation: dxr-swap .4s ease; }
.dxr-track {
  height: 3px;
  margin: 22px auto 0;
  max-width: 420px;
  overflow: hidden;
  border-radius: 999px;
  background: rgba(255, 255, 255, .09);
}
.dxr-bar {
  width: 12%;
  height: 100%;
  border-radius: inherit;
  background: linear-gradient(90deg, #14b8a6, #2dd4bf, #5eead4);
  box-shadow: 0 0 16px rgba(45, 212, 191, .5);
  transition: width .65s cubic-bezier(.22, .8, .28, 1);
}
.dxr-meta {
  display: flex;
  justify-content: space-between;
  gap: 14px;
  margin: 9px auto 0;
  max-width: 420px;
  color: #64748b;
  font: 700 9px/1 var(--dxr-mono);
  letter-spacing: .09em;
  text-transform: uppercase;
}
.dxr-retry {
  display: none;
  margin-top: 18px;
  padding: 9px 14px;
  border: 1px solid rgba(248, 113, 113, .4);
  border-radius: 9px;
  background: rgba(248, 113, 113, .12);
  color: #fff;
  cursor: pointer;
}
.dxr-loader[data-error] .dxr-retry { display: inline-flex; }
.dxr-loader[data-error] .dxr-visual { opacity: .35; }

/* Two columns once there is room — the visual earns its space beside the copy. */
@media (min-width: 760px) {
  .dxr-loader-grid {
    grid-template-columns: minmax(220px, .8fr) minmax(300px, 1.2fr);
    gap: clamp(28px, 6vw, 72px);
    text-align: left;
  }
  .dxr-visual { width: min(280px, 30vw); }
  .dxr-loader h2 { margin-left: 0; }
  .dxr-loader-copy > p { margin-left: 0; min-height: 44px; }
  .dxr-ad, .dxr-track, .dxr-meta { margin-left: 0; }
}

@keyframes dxr-spin { to { transform: rotate(360deg); } }
@keyframes dxr-float { 0%, 100% { transform: translateY(-5px) rotate(-2deg); } 50% { transform: translateY(7px) rotate(1deg); } }
@keyframes dxr-chip { 0%, 100% { transform: translateY(-4px); } 50% { transform: translateY(5px); } }
@keyframes dxr-pulse { 0%, 100% { transform: scaleX(.65); opacity: .45; } 50% { transform: scaleX(1); opacity: .95; } }
@keyframes dxr-swap { from { opacity: .1; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }

@media (prefers-reduced-motion: reduce) {
  .dxr *, .dxr *::before, .dxr *::after {
    animation-duration: .01ms !important;
    animation-iteration-count: 1 !important;
    transition-duration: .01ms !important;
  }
}
`;

const ICON_ALIGN = (bars: string) =>
  `<svg width="15" height="13" viewBox="0 0 15 13" aria-hidden="true">${bars}</svg>`;

/**
 * The ribbon's DOM, as one template.
 *
 * `data-dxr` names are the addressing contract (see the module header); the
 * behavioural attributes (`data-cmd`, `data-align`, `data-indent`, `data-list`,
 * `data-tt`) are what `ribbon.ts` delegates on, so adding a control here that
 * follows an existing convention needs no wiring change.
 */
export const RIBBON_HTML = `
<div class="dxr-chrome">
  <div class="dxr-titlebar">
    <span class="dxr-brand"><span class="dxr-mark"></span><span class="dxr-brandname">Docxodus</span>
      <span class="dxr-docname" data-dxr="docname">no document</span></span>
    <div class="dxr-quick" data-dxr-files>
      <button type="button" data-dxr="new" title="Start a new blank document">New</button>
      <label class="dxr-btn" tabindex="0">Open<input data-dxr="file" type="file" accept=".docx" hidden /></label>
      <button type="button" data-dxr="save" disabled>Save</button>
    </div>
    <div class="dxr-quick">
      <button type="button" class="dxr-icon" data-dxr="undo" title="Undo (Ctrl+Z)" aria-label="Undo">&#8630;</button>
      <button type="button" class="dxr-icon" data-dxr="redo" title="Redo (Ctrl+Shift+Z)" aria-label="Redo">&#8631;</button>
    </div>
    <span class="dxr-spacer"></span>
    <span class="dxr-status" data-dxr="status" role="status" aria-live="polite">Booting WASM&#8230;</span>
  </div>

  <div class="dxr-tabs" role="tablist">
    <button type="button" class="dxr-tab" role="tab" data-tab="home" aria-selected="true">Home</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="insert" aria-selected="false">Insert</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="layout" aria-selected="false">Layout</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="table" data-contextual aria-selected="false" hidden>Table</button>
  </div>

  <div class="dxr-ribbon" data-dxr="ribbon" aria-disabled="true">
    <!-- HOME ──────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="home" data-active>
      <div class="dxr-group">
        <span class="dxr-glabel">Text</span>
        <div class="dxr-row">
          <button type="button" class="dxr-icon" data-cmd="bold" title="Bold (Ctrl+B)"><b>B</b></button>
          <button type="button" class="dxr-icon" data-cmd="italic" title="Italic (Ctrl+I)"><i>I</i></button>
          <button type="button" class="dxr-icon" data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
          <button type="button" class="dxr-icon" data-cmd="strike" title="Strikethrough"><s>S</s></button>
          <button type="button" class="dxr-icon" data-cmd="code" title="Inline code">&lt;/&gt;</button>
          <button type="button" class="dxr-icon" data-cmd="superscript" title="Superscript">x&#178;</button>
          <button type="button" class="dxr-icon" data-cmd="subscript" title="Subscript">x&#8322;</button>
        </div>
        <div class="dxr-row">
          <input data-dxr="fontsize" data-dxr-list="fontsizes" type="number" min="1" max="1638" step="0.5"
                 placeholder="pt" title="Font size in points — type any value or pick a preset" />
          <datalist data-dxr="fontsizes">
            <option>8</option><option>9</option><option>10</option><option>11</option>
            <option>12</option><option>14</option><option>16</option><option>18</option>
            <option>20</option><option>24</option><option>28</option><option>36</option>
            <option>48</option><option>72</option><option>96</option>
          </datalist>
          <select data-dxr="fontfamily" title="Font family — applies to the selection">
            <option value="">Font&#8230;</option>
            <option>Calibri</option><option>Times New Roman</option><option>Arial</option>
            <option>Georgia</option><option>Cambria</option><option>Courier New</option>
            <option>Verdana</option><option>Garamond</option>
          </select>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Paragraph</span>
        <div class="dxr-row">
          <button type="button" class="dxr-icon" data-align="left" title="Align left">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="0" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="0" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="center" title="Align center">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="3" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="3" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="right" title="Align right">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="6" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="justify" title="Justify">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="0" y="3.8" width="15" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-indent="-720" title="Decrease indent">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="6" y="7.6" width="9" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/><path d="M4.6 4.2v4.6L0.6 6.5z"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-indent="720" title="Increase indent">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="6" y="7.6" width="9" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/><path d="M0.6 4.2v4.6l4-2.3z"/>',
          )}</button>
        </div>
        <div class="dxr-row">
          <button type="button" data-list="bullet" title="Bullet list">Bullets</button>
          <button type="button" data-list="decimal" title="Numbered list">Numbered</button>
          <button type="button" data-pagebreak title="Start this block on a new page">Page break</button>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Block</span>
        <div class="dxr-row">
          <select data-dxr="style" title="Paragraph style">
            <option value="">Style&#8230;</option>
            <option value="Normal">Normal</option>
            <option value="Heading1">Heading 1</option>
            <option value="Heading2">Heading 2</option>
            <option value="Heading3">Heading 3</option>
            <option value="Title">Title</option>
          </select>
        </div>
        <div class="dxr-row">
          <button type="button" data-dxr="delblock" class="dxr-danger" title="Delete the block the caret is in">Delete block</button>
        </div>
      </div>
    </div>

    <!-- INSERT ────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="insert">
      <div class="dxr-group">
        <span class="dxr-glabel">Table</span>
        <div class="dxr-row">
          <button type="button" data-dxr="table" title="Insert a table — pick its size on the grid">&#9638; Table</button>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Rules</span>
        <div class="dxr-row">
          <button type="button" data-dxr="hr" class="dxr-wide" title="Insert a single rule"><svg width="22" height="8" viewBox="0 0 22 8" aria-hidden="true"><rect x="0" y="3.2" width="22" height="1.4"/></svg>Single</button>
          <button type="button" data-dxr="hrThick" class="dxr-wide" title="Insert a thick rule"><svg width="22" height="8" viewBox="0 0 22 8" aria-hidden="true"><rect x="0" y="2.4" width="22" height="3.2"/></svg>Thick</button>
          <button type="button" data-dxr="hrDouble" class="dxr-wide" title="Insert a double rule"><svg width="22" height="8" viewBox="0 0 22 8" aria-hidden="true"><rect x="0" y="1.6" width="22" height="1.3"/><rect x="0" y="5" width="22" height="1.3"/></svg>Double</button>
        </div>
        <div class="dxr-row">
          <select data-dxr="rulepos" title="Where the rule lands relative to the current block">
            <option value="below">Below block</option>
            <option value="above">Above block</option>
          </select>
          <button type="button" data-dxr="hrClear" title="Remove the rule or paragraph border">Clear</button>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">References</span>
        <div class="dxr-row">
          <button type="button" data-dxr="footnote" title="Cite a new footnote at the caret">Footnote</button>
          <button type="button" data-dxr="endnote" title="Cite a new endnote at the caret">Endnote</button>
        </div>
      </div>
    </div>

    <!-- LAYOUT ────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="layout">
      <div class="dxr-group">
        <span class="dxr-glabel">View</span>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="paginated" type="checkbox" /> Page view</label>
        </div>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="headerfooter" type="checkbox" /> Header &amp; footer bands</label>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Page numbers</span>
        <div class="dxr-row">
          <select data-dxr="pgfmt" title="Number format for this section">
            <option value="">Format&#8230;</option>
            <option value="decimal">1, 2, 3</option>
            <option value="lowerLetter">a, b, c</option>
            <option value="upperLetter">A, B, C</option>
            <option value="lowerRoman">i, ii, iii</option>
            <option value="upperRoman">I, II, III</option>
          </select>
          <label class="dxr-toggle">Start at
            <input data-dxr="pgstart" type="number" min="1" step="1"
                   title="Restart this section's numbering at this value" />
          </label>
          <button type="button" data-dxr="pgclear" title="Continue the previous section's numbering">Clear</button>
        </div>
        <div class="dxr-row">
          <span class="dxr-note">Applies to the section holding the caret.</span>
        </div>
      </div>

      <!-- The field lands in the running footer (Word's convention, and where the band
           puts it), so it belongs with the section's numbering rather than under Insert,
           whose controls all act at the caret. -->
      <div class="dxr-group">
        <span class="dxr-glabel">Footer fields</span>
        <div class="dxr-row">
          <button type="button" data-dxr="pagenum" title="Add a page-number field to the footer">Page number</button>
          <button type="button" data-dxr="totalpages" title="Add a total-pages field to the footer">Total pages</button>
        </div>
        <div class="dxr-row">
          <span class="dxr-note">Added to the footer story.</span>
        </div>
      </div>
    </div>

    <!-- TABLE (contextual) ──────────────────────────────────────────────────
         Row/column editing lives here rather than in a floating toolbar: a docked
         contextual tab cannot overlap the cell you are editing. -->
    <div class="dxr-panel" data-panel="table">
      <div class="dxr-group">
        <span class="dxr-glabel">Rows</span>
        <div class="dxr-row">
          <button type="button" data-tt="rowAbove" title="Insert a row above this one">Insert above</button>
          <button type="button" data-tt="rowBelow" title="Insert a row below this one">Insert below</button>
          <button type="button" data-tt="delRow" class="dxr-danger" title="Delete this row">Delete row</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Columns</span>
        <div class="dxr-row">
          <button type="button" data-tt="colLeft" title="Insert a column to the left">Insert left</button>
          <button type="button" data-tt="colRight" title="Insert a column to the right">Insert right</button>
          <button type="button" data-tt="delCol" class="dxr-danger" title="Delete this column">Delete column</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Anchor rail — live engine state, not decoration. -->
  <div class="dxr-rail" data-dxr-rail>
    <span class="dxr-cell"><span class="dxr-k">anchor</span><span class="dxr-v" data-dxr="railAnchor">&#8212;</span></span>
    <span class="dxr-cell"><span class="dxr-k">blocks</span><span class="dxr-v" data-dxr="railBlocks">&#8212;</span></span>
    <span class="dxr-cell"><span class="dxr-k">session</span><span class="dxr-v" data-dxr="railSession">&#8212;</span></span>
    <span class="dxr-cell"><span class="dxr-k">last op</span><span class="dxr-v" data-dxr="railOp">&#8212;</span></span>
  </div>
</div>

<div class="dxr-scroll" data-dxr-scroll>
  <p class="dxr-hint" data-dxr-hint></p>
  <div class="dxr-surface" data-dxr="editor" data-dxr-surface data-view="continuous"></div>
</div>

<!-- Table size picker, anchored to the Insert tab's Table button. -->
<div class="dxr-pop" data-dxr="gridpicker">
  <div class="dxr-gridcells" data-dxr="gridcells"></div>
  <div class="dxr-popfoot">
    <span data-dxr="gridlabel">0 &#215; 0</span>
    <span style="display:inline-flex; align-items:center; gap:10px;">
      <label class="dxr-toggle">Align
        <select data-dxr="gridalign">
          <option value="left">Left</option>
          <option value="center">Center</option>
          <option value="right">Right</option>
        </select>
      </label>
      <label class="dxr-toggle"><input data-dxr="gridborderless" type="checkbox" checked /> Borderless</label>
    </span>
  </div>
</div>

<div class="dxr-loader" data-dxr="loader" aria-live="polite" hidden>
  <div class="dxr-loader-grid">
    <div class="dxr-visual" aria-hidden="true">
      <div class="dxr-orbit"></div><div class="dxr-orbit dxr-two"></div>
      <div class="dxr-card"><i></i><i></i><i></i><i></i></div>
      <span class="dxr-chip dxr-b">B</span><span class="dxr-chip dxr-r">&#177;</span><span class="dxr-chip dxr-s">&#8595;</span>
    </div>
    <div class="dxr-loader-copy">
      <div class="dxr-eyebrow" data-dxr="loaderEyebrow">Running locally in this tab</div>
      <h2 data-dxr="loaderTitle">Booting .NET inside your browser</h2>
      <p data-dxr="loaderCopy">Streaming the trimmed WebAssembly runtime.</p>
      <div class="dxr-ad" data-dxr="loaderAd">
        <span class="dxr-num" data-dxr="loaderNumber">01</span>
        <div><strong data-dxr="loaderAdTitle"></strong><span class="dxr-adcopy" data-dxr="loaderAdCopy"></span></div>
      </div>
      <div class="dxr-track" aria-hidden="true"><div class="dxr-bar" data-dxr="loaderBar"></div></div>
      <div class="dxr-meta"><span data-dxr="loaderLabel">Loading engine</span><span data-dxr="loaderMeta">DOCX &#8594; WASM &#8594; DOCX</span></div>
      <button type="button" class="dxr-retry" data-dxr="loaderRetry">Retry loading</button>
    </div>
  </div>
</div>
`;

/** The default hint copy, shown above the document in full chrome. */
export const RIBBON_HINT_HTML =
  "Click any paragraph, heading, table cell, footnote or header/footer line to edit it. " +
  "<kbd>Enter</kbd> splits a block, <kbd>Backspace</kbd> at the start merges it into the previous one, " +
  "<kbd>Ctrl</kbd>+<kbd>Z</kbd> undoes. Only the block you changed re-renders — everything else keeps " +
  "full fidelity, and <b>Save</b> writes a lossless .docx.";

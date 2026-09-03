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
 * ── Shape ───────────────────────────────────────────────────────────────────
 * Word's: a title bar, a tab strip (Home · Insert · Layout · References ·
 * Review · View, plus the contextual Table and Header & Footer tabs), the panel
 * for the active tab, the document with its comment gutter, and a status bar
 * that carries the anchor rail, the page/word counts and the zoom control.
 *
 * ── Responsiveness ──────────────────────────────────────────────────────────
 * Layout keys off `data-chrome` on the root ("full" | "compact"), which
 * `ribbon.ts` sets from a ResizeObserver on the ROOT rather than from viewport
 * media queries. A narrow embed inside a wide desktop page is narrow, and only
 * container measurement gets that right. Compact is the base (mobile-first);
 * `[data-chrome="full"]` adds back the desktop affordances.
 */

/** Bumped whenever RIBBON_CSS changes, so a stale injected stylesheet is replaced. */
export const RIBBON_STYLE_VERSION = "10";
export const RIBBON_STYLE_ATTR = "data-docxodus-ribbon-styles";

/** Width of the comment gutter the surface reserves beside the sheet, in px. */
export const COMMENT_GUTTER_WIDTH = 264;

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
  --dxr-desk: #eef2f6;
  --dxr-ui: "Inter", system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  --dxr-serif: Georgia, "Times New Roman", serif;
  --dxr-mono: ui-monospace, SFMono-Regular, "Cascadia Mono", Consolas, monospace;
  --dxr-shadow-sm: 0 1px 2px rgba(15, 23, 42, .04);
  --dxr-shadow-md: 0 1px 3px rgba(15, 23, 42, .06), 0 4px 6px rgba(15, 23, 42, .03);
  --dxr-shadow-lg: 0 4px 6px rgba(15, 23, 42, .05), 0 10px 15px rgba(15, 23, 42, .04);
  --dxr-shadow-xl: 0 8px 16px rgba(15, 23, 42, .06), 0 20px 25px rgba(15, 23, 42, .05);
  --dxr-ease: cubic-bezier(.4, 0, .2, 1);
  --dxr-tap: 28px;
  --dxr-gutter: ${COMMENT_GUTTER_WIDTH}px;

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
  text-size-adjust: 100%;
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

/* ── Scroll-edge affordances ──────────────────────────────────────────────────
   The chrome's strips (title bar, tabs, panels, rail) scroll horizontally
   rather than squashing or wrapping — the compact rule is "no command is
   removed, the strip scrolls". But a hard clip mid-control reads as a broken
   layout, especially on a phone, so ribbon.ts measures every strip and stamps
   data-fade with the edges that hide more content ("l", "r", or both). The
   mask dissolves content at exactly those edges — the horizontal twin of
   .dxr-scroll's gradient veils — and lifts entirely once the strip fits, so
   nothing is ever faded unless there is more behind it. */
.dxr-titlebar, .dxr-tabs, .dxr-panel, .dxr-rail, .dxr-findbar { --dxr-fade-l: 0px; --dxr-fade-r: 0px; }
.dxr [data-fade] {
  -webkit-mask-image: linear-gradient(to right,
    transparent, #000 var(--dxr-fade-l),
    #000 calc(100% - var(--dxr-fade-r)), transparent);
  mask-image: linear-gradient(to right,
    transparent, #000 var(--dxr-fade-l),
    #000 calc(100% - var(--dxr-fade-r)), transparent);
}
.dxr [data-fade~="l"] { --dxr-fade-l: 28px; }
.dxr [data-fade~="r"] { --dxr-fade-r: 28px; }

/* ── Chrome shell ─────────────────────────────────────────────────────────── */
.dxr-chrome {
  position: sticky;
  top: 0;
  z-index: 10;
  flex: 0 0 auto;
  background: var(--dxr-chrome);
  border-bottom: 1px solid var(--dxr-rule);
}

/* Scrolls in EVERY density (compact just tightens it): the fade affordance
   marks any overflowing strip as scrollable, so it must actually be. */
.dxr-titlebar {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 6px 10px 0;
  overflow-x: auto;
  scrollbar-width: none;
}
.dxr-titlebar::-webkit-scrollbar { display: none; }
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
  padding: 4px 10px 0;
  overflow-x: auto;
  scrollbar-width: none;
}
.dxr-tabs::-webkit-scrollbar { display: none; }
/* Line-variant tabs (the house Tabs default): no boxes, no fills — the selected
   tab is a 2px accent underline sitting on the panel's own hairline. */
.dxr-tab {
  flex: 0 0 auto;
  padding: 6px 13px 7px;
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
/* Contextual tabs — present only while the caret is in a table / a header or footer story. */
.dxr-tab[data-contextual] { color: var(--dxr-accent); }
.dxr-tab[data-contextual]::before { content: "\\2022"; margin-right: 6px; opacity: .6; }
.dxr-tab[hidden] { display: none; }

/* ── Ribbon panels ────────────────────────────────────────────────────────── */
.dxr-ribbon { background: var(--dxr-chrome-sunk); border-top: 1px solid var(--dxr-rule); }
.dxr-ribbon[aria-disabled="true"] { opacity: .45; pointer-events: none; }
.dxr-panel {
  display: none;
  align-items: stretch;
  padding: 6px 10px 7px;
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
  gap: 4px;
  padding: 0 12px;
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
.dxr-row + .dxr-row { margin-top: 1px; }

/* ── Controls ─────────────────────────────────────────────────────────────── */
/* One button family, house-style: every control is a ghost — no borders, no
   fills, no per-button boxes. Hover paints a quiet slate wash; the active
   (toggled) state paints a teal one. Save alone is the flat primary, so the
   surface has exactly one accent-colored action. */
.dxr button, .dxr label.dxr-btn {
  min-height: var(--dxr-tap);
  padding: 4px 8px;
  border: none;
  border-radius: 6px;
  background: none;
  color: var(--dxr-ink);
  font: inherit;
  font-size: 12.5px;
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
.dxr button.dxr-on, .dxr button[aria-pressed="true"] { color: var(--dxr-accent); background: var(--dxr-wash-on); }
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
.dxr button.dxr-icon { min-width: var(--dxr-tap); padding: 4px 6px; text-align: center; color: var(--dxr-muted); }
.dxr button.dxr-icon:hover:not(:disabled) { color: var(--dxr-ink); }
.dxr button.dxr-icon.dxr-on { color: var(--dxr-accent); }
/* Icons are inline SVG drawn from the same shapes the control acts on (text lines,
   rules), so alignment and indent read at a glance instead of relying on arrow
   glyphs that collide with undo/redo. currentColor keeps them in step with state. */
.dxr button svg { display: block; margin: 0 auto; fill: currentColor; }
.dxr button.dxr-wide { display: inline-flex; align-items: center; justify-content: center; gap: 6px; }
.dxr button.dxr-danger:hover { color: var(--dxr-danger); background: var(--dxr-danger-wash); }
.dxr select, .dxr input[type="number"], .dxr input[type="text"], .dxr input[type="search"], .dxr input[type="url"] {
  min-height: var(--dxr-tap);
  padding: 3px 6px;
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
.dxr [data-dxr="fontsize"] { width: 58px; }
.dxr [data-dxr="fontfamily"] { width: 138px; }
.dxr [data-dxr="style"] { width: 148px; }
.dxr [data-dxr="pgstart"], .dxr [data-dxr="spacebefore"], .dxr [data-dxr="spaceafter"],
.dxr [data-dxr="specialindentval"] { width: 62px; }
.dxr [data-dxr="linespacing"], .dxr [data-dxr="listmenu"], .dxr [data-dxr="highlight"],
.dxr [data-dxr="markup"], .dxr [data-dxr="zoom"], .dxr [data-dxr="margins"],
.dxr [data-dxr="orientation"], .dxr [data-dxr="pagesize"], .dxr [data-dxr="specialindent"] { max-width: 150px; }
.dxr-toggle {
  display: inline-flex;
  gap: 5px;
  align-items: center;
  min-height: var(--dxr-tap);
  color: var(--dxr-ink);
  font-size: 12.5px;
  white-space: nowrap;
  cursor: pointer;
}
.dxr-toggle input[type="checkbox"] { accent-color: var(--dxr-accent); margin: 0; }
.dxr-note { color: var(--dxr-muted); font-size: 11.5px; }
.dxr-kbd { color: var(--dxr-faint); font-family: var(--dxr-mono); font-size: 10px; }

/* Colour swatch buttons: the button shows the current colour as a bar under its glyph and
   opens the native picker that sits (visually hidden) inside it. */
.dxr-swatch { position: relative; display: inline-flex; flex-direction: column; align-items: center; gap: 2px; }
.dxr-swatch > button { padding-bottom: 6px; }
.dxr-swatch .dxr-swatch-bar {
  position: absolute; left: 7px; right: 7px; bottom: 4px; height: 3px; border-radius: 2px;
  background: var(--dxr-swatch, #1e293b); pointer-events: none;
}
.dxr-swatch input[type="color"] {
  position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; cursor: pointer; border: 0; padding: 0;
}
.dxr-swatch [data-dxr="highlight"] {
  min-height: var(--dxr-tap);
}
.dxr-hl-swatch { display: inline-block; width: 12px; height: 12px; border-radius: 2px; vertical-align: -2px; margin-right: 4px; border: 1px solid rgba(15,23,42,.15); }

/* ── Find & replace bar ───────────────────────────────────────────────────── */
.dxr-findbar {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 5px 10px;
  overflow-x: auto;
  border-top: 1px solid var(--dxr-rule);
  background: var(--dxr-chrome);
  scrollbar-width: none;
}
.dxr-findbar[hidden] { display: none; }
.dxr-findbar > *, .dxr-findbar [data-dxr="replacegroup"] { flex: 0 0 auto; white-space: nowrap; }
.dxr-findbar [data-dxr="replacegroup"] { display: inline-flex; align-items: center; gap: 6px; }
.dxr-findbar input[type="search"], .dxr-findbar input[type="text"] { width: 180px; }
.dxr-findbar .dxr-findcount { min-width: 5ch; color: var(--dxr-muted); font-family: var(--dxr-mono); font-size: 11.5px; }

/* ── Status bar (carries the anchor rail) ─────────────────────────────────────
   The addressing spine made permanent chrome. Every block in this editor is
   addressable as kind:scope:unid, and the model of record is a live WASM session —
   so the surface states both, live, instead of hiding them behind devtools. It also
   reports each command's real cost, which is what a smoke test needs to see. Word's
   own status-bar cells (page, words, zoom) sit on the same strip. */
.dxr-rail {
  display: flex;
  align-items: center;
  flex: 0 0 auto;
  height: 26px;
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
  padding: 0 12px;
  border-right: 1px solid var(--dxr-rule);
  white-space: nowrap;
}
.dxr-rail .dxr-cell:first-child { padding-left: 0; }
.dxr-rail .dxr-cell.dxr-cell-end { margin-left: auto; border-right: none; padding-right: 0; }
.dxr-rail .dxr-k {
  color: var(--dxr-faint);
  font-family: var(--dxr-ui);
  font-size: 9.5px;
  letter-spacing: .1em;
  text-transform: uppercase;
}
.dxr-rail .dxr-v { color: var(--dxr-data); }
.dxr-rail .dxr-v.dxr-flash { animation: dxr-railflash .45s ease-out; }
.dxr-rail .dxr-v.dxr-plain { color: var(--dxr-ink); font-family: var(--dxr-ui); font-size: 11.5px; }
@keyframes dxr-railflash { from { background: #ccfbf1; } to { background: transparent; } }
.dxr-zoom { display: inline-flex; align-items: center; gap: 2px; }
.dxr-zoom button { min-height: 22px; min-width: 22px; padding: 0 5px; font-size: 13px; color: var(--dxr-muted); }
.dxr-zoom select { min-height: 22px; padding: 0 4px; font-size: 11.5px; }

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
.dxr[data-chrome] .dxr-surface { position: relative; margin: 26px auto; padding: 0 16px 96px; }
/* The comment gutter sits to the right of the sheet; the surface reserves its width so the
   page shifts left rather than the bubbles covering it — Word's markup area. */
.dxr[data-chrome="full"] .dxr-surface[data-comments="on"] { padding-right: calc(var(--dxr-gutter) + 8px); }
.dxr-surface [contenteditable="true"]:focus {
  outline: 2px solid var(--dxr-accent);
  outline-offset: 2px;
  border-radius: 2px;
}
@media (hover: hover) { .dxr-surface [contenteditable="true"]:hover { background: var(--dxr-wash); } }

/* Header/footer bands (continuous view): the top and bottom margin of the sheet, drawn the way
   Word draws a header that is being edited — a dashed rule with a small tag in the margin. */
.dxr-surface .docx-hf-band {
  /* Docked outside the zoomed sheet, so it takes the page's on-screen width from the
     custom property the viewport publishes rather than stretching to the whole surface. */
  position: relative;
  width: min(100%, var(--docx-sheet-width, 100%));
  margin: 0 auto;
  padding: 34px 72px 12px;
  background: var(--dxr-sheet);
  border-radius: 3px 3px 0 0;
  box-shadow: 0 1px 3px rgba(15, 23, 42, .08), 0 12px 24px rgba(15, 23, 42, .06);
}
.dxr-surface .docx-hf-band[data-hf-band="footer"] { padding: 12px 72px 34px; border-radius: 0 0 3px 3px; }
/* The band and the sheet are one piece of paper: kill the sheet's own top/bottom corners and
   let the dashed rule separate the stories from the body. */
.dxr-surface .docx-hf-band + .docx-body-flow { margin-top: 0; border-radius: 0; box-shadow: none; }
.dxr-surface .docx-hf-band[data-hf-band="footer"] { margin-top: 0; }
.dxr[data-chrome] .dxr-surface[data-view="continuous"] .docx-hf-band ~ .docx-body-flow { box-shadow: none; }
.dxr[data-chrome] .dxr-surface[data-view="continuous"]:has(.docx-hf-band) .docx-body-flow { padding: 18px 0; }
.dxr-surface .docx-hf-band::after {
  content: ""; position: absolute; left: 0; right: 0; bottom: 0; border-bottom: 1px dashed var(--dxr-rule-strong);
}
.dxr-surface .docx-hf-band[data-hf-band="footer"]::after { top: 0; bottom: auto; border-bottom: 0; border-top: 1px dashed var(--dxr-rule-strong); }
.dxr-surface .docx-hf-band:hover::after, .dxr-surface .docx-hf-band[data-hf-active]::after { border-color: var(--dxr-accent); }
.dxr-surface .docx-hf-tag {
  position: absolute; left: 72px; bottom: -9px; z-index: 2;
  display: flex; align-items: center; gap: 8px;
  padding: 1px 7px; border: 1px solid var(--dxr-rule-strong); border-radius: 4px;
  background: #fff; color: var(--dxr-muted);
  font-family: var(--dxr-ui); font-size: 9.5px; font-weight: 600; letter-spacing: .1em; text-transform: uppercase;
}
.dxr-surface .docx-hf-band[data-hf-band="footer"] .docx-hf-tag { bottom: auto; top: -9px; }
.dxr-surface .docx-hf-band[data-hf-active] .docx-hf-tag { border-color: var(--dxr-accent); color: var(--dxr-accent); }
.dxr-surface .docx-hf-kinds { display: inline-flex; gap: 2px; }
.dxr-surface .docx-hf-kinds button {
  min-height: 0; padding: 1px 6px; border-radius: 3px; font-family: var(--dxr-ui); font-size: 9.5px; font-weight: 600;
  letter-spacing: .06em; text-transform: uppercase; color: var(--dxr-muted); background: none;
}
.dxr-surface .docx-hf-kinds button[data-on] { background: var(--dxr-wash-on); color: var(--dxr-accent); }
.dxr-surface .docx-hf-inherited { font-weight: 400; letter-spacing: normal; text-transform: none; font-style: italic; }
.dxr-surface .docx-hf-placeholder { color: var(--dxr-faint); font-style: italic; cursor: text; }
/* Bands only: a page host's height is the band the paginator reserved, and a floor here would
   read as "the story grew" and trigger a needless re-paginate. */
.dxr-surface .docx-hf-band .docx-hf-body { min-height: 1.4em; }
/* Page view: the page's own header/footer area is click-to-edit. */
.dxr-surface .page-header[data-hf-page], .dxr-surface .page-footer[data-hf-page] { cursor: text; }
.dxr-surface .page-header[data-hf-page]:hover, .dxr-surface .page-footer[data-hf-page]:hover {
  outline: 1px dashed var(--dxr-rule-strong); outline-offset: 2px;
}
.dxr-surface .page-header[data-hf-active], .dxr-surface .page-footer[data-hf-active] {
  outline: 1px dashed var(--dxr-accent); outline-offset: 2px; z-index: 5; background: #fff;
}
.dxr-surface [data-hf-active][data-hf-label]::before {
  content: attr(data-hf-label); position: absolute; left: -2px; top: -18px;
  padding: 1px 6px; border: 1px solid var(--dxr-accent); border-radius: 4px; background: #fff; color: var(--dxr-accent);
  font-family: var(--dxr-ui); font-size: 9px; font-weight: 600; letter-spacing: .1em; text-transform: uppercase; line-height: 1.4;
}
.dxr-surface .page-footer[data-hf-active][data-hf-label]::before { top: auto; bottom: -18px; }
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
/* Small prompt popovers (link URL, page setup), anchored under their button. */
.dxr-pop form { display: flex; flex-direction: column; gap: 8px; min-width: 260px; font-size: 12.5px; }
.dxr-pop form label { display: flex; flex-direction: column; gap: 3px; color: var(--dxr-muted); font-size: 11.5px; }
.dxr-pop form .dxr-row { justify-content: flex-end; }
.dxr-pop form button[type="submit"] { background: var(--dxr-accent); color: #fff; }
.dxr-pop form button[type="submit"]:hover { background: var(--dxr-accent-hover); }

/* ── Compact chrome (the mobile-first base) ───────────────────────────────────
   One scrolling strip per tab instead of a multi-row ribbon: group labels turn
   into inline dividers, the rail and hint step aside, and the picker becomes a
   bottom sheet because there is no room to hang a popover off a button. */
/* A 320px phone cannot hold the product name AND the file actions, and the file
   actions are what a user came for. The strip also scrolls, so nothing is stranded. */
/* Touch targets grow with the LAYOUT, not the pointer alone: a phone-width chrome is driven
   by a thumb whether or not the browser reports a coarse pointer (an emulated narrow viewport
   does not), so the compact strip adopts the same 40px control floor the coarse-pointer media
   query sets, keeping every icon button at a real tap size. */
.dxr[data-chrome="compact"] { --dxr-tap: 40px; }
/* The floor buys tap SIZE; the rows must not spend it again on padding. A phone stacks three
   chrome rows — title, tabs, commands — above the page, so every pixel of row padding is paid
   three times before the document starts, and the arcade demo's card (which sizes itself to the
   game screen plus its controls) runs out of viewport when it is. */
.dxr[data-chrome="compact"] .dxr-titlebar {
  gap: 6px;
  padding: 0 8px;
}
.dxr[data-chrome="compact"] .dxr-brandname { display: none; }
.dxr[data-chrome="compact"] .dxr-brand { flex: 0 1 auto; }
/* A filename clipped to "do..." tells you less than no filename at all, so it keeps a
   readable floor and the strip scrolls instead. The tighter file-action padding is what
   buys that floor back on a 390px screen. */
.dxr[data-chrome="compact"] .dxr-brand .dxr-docname { min-width: 6ch; max-width: 14ch; }
.dxr[data-chrome="compact"] .dxr-quick button,
.dxr[data-chrome="compact"] .dxr-quick label.dxr-btn { padding: 4px 8px; }
.dxr[data-chrome="compact"] .dxr-status { display: none; }
.dxr[data-chrome="compact"] .dxr-tabs { padding: 0 8px; }
/* A tab is a WIDE target — its label plus 11px of side padding already clears the tap floor
   horizontally — and it is navigation rather than a command, so it does not need the command
   row's full height. Word's own phone ribbon keeps its tab strip slim for the same reason: the
   page starts higher. The commands below keep the full 40px floor. */
.dxr[data-chrome="compact"] .dxr-tab { min-height: 34px; padding: 6px 11px 7px; font-size: 12px; }
.dxr[data-chrome="compact"] .dxr-panel {
  align-items: center;
  gap: 4px;
  padding: 2px 8px;
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
.dxr[data-chrome="compact"] .dxr-row + .dxr-row { margin-top: 0; }
/* Compact trims the chrome around the page, never the page: the document's own column width
   is what the viewport's fit-to-width zoom scales, so a phone shows a whole smaller page
   instead of a narrower one that breaks its lines somewhere Word never would. */
.dxr[data-chrome="compact"] .dxr-surface { margin: 12px auto; padding: 0 10px 64px; }
.dxr[data-chrome="compact"] .dxr-surface[data-view="continuous"] .docx-body-flow { padding: 22px 0; }
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-band { padding: 26px 18px 10px; }
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-band[data-hf-band="footer"] { padding: 10px 18px 26px; }
.dxr[data-chrome="compact"] .dxr-surface .docx-hf-tag { left: 18px; }
/* No room for a markup column on a phone: comments stay highlighted, bubbles step aside. */
.dxr[data-chrome="compact"] .docx-comment-gutter, .dxr[data-chrome="compact"] .docx-comment-leaders { display: none; }
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

/** Word's highlighter palette (ST_HighlightColor) with its screen colours, in Word's order. */
export const HIGHLIGHT_COLORS: ReadonlyArray<{ value: string; label: string; css: string }> = [
  { value: "yellow", label: "Yellow", css: "#ffff00" },
  { value: "green", label: "Bright green", css: "#00ff00" },
  { value: "cyan", label: "Turquoise", css: "#00ffff" },
  { value: "magenta", label: "Pink", css: "#ff00ff" },
  { value: "blue", label: "Blue", css: "#0000ff" },
  { value: "red", label: "Red", css: "#ff0000" },
  { value: "darkBlue", label: "Dark blue", css: "#000080" },
  { value: "darkCyan", label: "Teal", css: "#008080" },
  { value: "darkGreen", label: "Green", css: "#008000" },
  { value: "darkMagenta", label: "Violet", css: "#800080" },
  { value: "darkRed", label: "Dark red", css: "#800000" },
  { value: "darkYellow", label: "Dark yellow", css: "#808000" },
  { value: "darkGray", label: "Gray 50%", css: "#808080" },
  { value: "lightGray", label: "Gray 25%", css: "#c0c0c0" },
  { value: "black", label: "Black", css: "#000000" },
];

/** The fonts every Word install has; the document's own fonts are added on top at open. */
export const COMMON_FONTS: readonly string[] = [
  "Calibri", "Cambria", "Arial", "Times New Roman", "Georgia", "Garamond", "Verdana",
  "Tahoma", "Trebuchet MS", "Book Antiqua", "Century Gothic", "Courier New", "Consolas",
  "Segoe UI", "Helvetica",
];

/**
 * The ribbon's DOM, as one template.
 *
 * `data-dxr` names are the addressing contract (see the module header); the
 * behavioural attributes (`data-cmd`, `data-align`, `data-indent`, `data-list`,
 * `data-tt`, `data-hf`, `data-rev`, `data-cmt`) are what `ribbon.ts` delegates on,
 * so adding a control here that follows an existing convention needs no wiring change.
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
    <button type="button" class="dxr-tab" role="tab" data-tab="references" aria-selected="false">References</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="review" aria-selected="false">Review</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="view" aria-selected="false">View</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="table" data-contextual aria-selected="false" hidden>Table</button>
    <button type="button" class="dxr-tab" role="tab" data-tab="headerfooter" data-contextual aria-selected="false" hidden>Header &amp; Footer</button>
  </div>

  <div class="dxr-ribbon" data-dxr="ribbon" aria-disabled="true">
    <!-- HOME ──────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="home" data-active>
      <div class="dxr-group">
        <span class="dxr-glabel">Font</span>
        <div class="dxr-row">
          <select data-dxr="fontfamily" title="Font family — applies to the selection">
            <option value="">Font&#8230;</option>
          </select>
          <input data-dxr="fontsize" data-dxr-list="fontsizes" type="number" min="1" max="1638" step="0.5"
                 placeholder="pt" title="Font size in points — type any value or pick a preset" />
          <datalist data-dxr="fontsizes">
            <option>8</option><option>9</option><option>10</option><option>11</option>
            <option>12</option><option>14</option><option>16</option><option>18</option>
            <option>20</option><option>24</option><option>28</option><option>36</option>
            <option>48</option><option>72</option><option>96</option>
          </datalist>
          <button type="button" class="dxr-icon" data-dxr="fontgrow" title="Increase font size (Ctrl+])">A<sup>&#9650;</sup></button>
          <button type="button" class="dxr-icon" data-dxr="fontshrink" title="Decrease font size (Ctrl+[)">A<sub>&#9660;</sub></button>
          <button type="button" class="dxr-icon" data-dxr="clearformat" title="Clear all formatting">A&#10008;</button>
        </div>
        <div class="dxr-row">
          <button type="button" class="dxr-icon" data-cmd="bold" title="Bold (Ctrl+B)"><b>B</b></button>
          <button type="button" class="dxr-icon" data-cmd="italic" title="Italic (Ctrl+I)"><i>I</i></button>
          <button type="button" class="dxr-icon" data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button>
          <button type="button" class="dxr-icon" data-cmd="strike" title="Strikethrough"><s>S</s></button>
          <button type="button" class="dxr-icon" data-cmd="superscript" title="Superscript">x&#178;</button>
          <button type="button" class="dxr-icon" data-cmd="subscript" title="Subscript">x&#8322;</button>
          <button type="button" class="dxr-icon" data-dxr="smallcaps" title="Small caps" aria-pressed="false"
                  style="font-variant: small-caps;">Aa</button>
          <button type="button" class="dxr-icon" data-cmd="code" title="Inline code">&lt;/&gt;</button>
          <span class="dxr-swatch" title="Font color">
            <button type="button" class="dxr-icon" data-dxr="fontcolorbtn" aria-label="Font color"><b>A</b></button>
            <span class="dxr-swatch-bar" data-dxr="fontcolorbar"></span>
            <input type="color" data-dxr="fontcolor" value="#1e293b" aria-label="Pick a font color" />
          </span>
          <span class="dxr-swatch" title="Text highlight color">
            <select data-dxr="highlight" aria-label="Text highlight color">
              <option value="">&#9998; Highlight</option>
              <option value="none">No color</option>
            </select>
          </span>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Paragraph</span>
        <div class="dxr-row">
          <button type="button" class="dxr-icon" data-list="bullet" title="Bullets">&#8226;&#8801;</button>
          <button type="button" class="dxr-icon" data-list="decimal" title="Numbering">1.&#8801;</button>
          <select data-dxr="listmenu" title="Numbering format">
            <option value="">List style&#8230;</option>
            <option value="bullet">&#8226; Bullet</option>
            <option value="decimal">1. 2. 3.</option>
            <option value="decimalParenthesis">(1) (2) (3)</option>
            <option value="lowerLetter">a. b. c.</option>
            <option value="lowerLetterParenthesis">(a) (b) (c)</option>
            <option value="upperLetter">A. B. C.</option>
            <option value="lowerRoman">i. ii. iii.</option>
            <option value="lowerRomanParenthesis">(i) (ii) (iii)</option>
            <option value="upperRoman">I. II. III.</option>
            <option value="none">No list</option>
          </select>
          <button type="button" class="dxr-icon" data-indent="-720" title="Decrease indent">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="6" y="7.6" width="9" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/><path d="M4.6 4.2v4.6L0.6 6.5z"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-indent="720" title="Increase indent">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="6" y="7.6" width="9" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/><path d="M0.6 4.2v4.6l4-2.3z"/>',
          )}</button>
        </div>
        <div class="dxr-row">
          <button type="button" class="dxr-icon" data-align="left" title="Align left (Ctrl+L)">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="0" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="0" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="center" title="Center (Ctrl+E)">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="3" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="3" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="right" title="Align right (Ctrl+R)">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="6" y="3.8" width="9" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="6" y="11.4" width="9" height="1.6"/>',
          )}</button>
          <button type="button" class="dxr-icon" data-align="justify" title="Justify (Ctrl+J)">${ICON_ALIGN(
            '<rect x="0" y="0" width="15" height="1.6"/><rect x="0" y="3.8" width="15" height="1.6"/><rect x="0" y="7.6" width="15" height="1.6"/><rect x="0" y="11.4" width="15" height="1.6"/>',
          )}</button>
          <select data-dxr="linespacing" title="Line spacing">
            <option value="">&#8597; Spacing&#8230;</option>
            <option value="1">1.0</option>
            <option value="1.15">1.15</option>
            <option value="1.5">1.5</option>
            <option value="2">2.0</option>
            <option value="2.5">2.5</option>
            <option value="3">3.0</option>
          </select>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Styles</span>
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

      <div class="dxr-group">
        <span class="dxr-glabel">Editing</span>
        <div class="dxr-row">
          <button type="button" data-dxr="findtoggle" title="Find (Ctrl+F)">&#128269; Find</button>
        </div>
        <div class="dxr-row">
          <button type="button" data-dxr="replacetoggle" title="Find and replace (Ctrl+H)">Replace</button>
        </div>
      </div>
    </div>

    <!-- INSERT ────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="insert">
      <div class="dxr-group">
        <span class="dxr-glabel">Pages</span>
        <div class="dxr-row">
          <button type="button" data-pagebreak title="Start this block on a new page (Ctrl+Enter)">Page break</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Tables</span>
        <div class="dxr-row">
          <button type="button" data-dxr="table" title="Insert a table — pick its size on the grid">&#9638; Table</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Illustrations</span>
        <div class="dxr-row">
          <label class="dxr-btn" tabindex="0" title="Insert a picture from a file">&#128444; Picture<input data-dxr="picturefile" type="file" accept="image/png,image/jpeg,image/gif,image/bmp,image/webp,image/tiff" hidden /></label>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Links</span>
        <div class="dxr-row">
          <button type="button" data-dxr="link" title="Link the selected text to a URL (Ctrl+K)">&#128279; Link</button>
          <button type="button" data-dxr="unlink" title="Remove the link at the caret">Unlink</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Comments</span>
        <div class="dxr-row">
          <button type="button" data-dxr="insertcomment" title="Comment on the selection">&#128172; Comment</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Header &amp; Footer</span>
        <div class="dxr-row">
          <button type="button" data-dxr="gotoheader" title="Edit the header">Header</button>
          <button type="button" data-dxr="gotofooter" title="Edit the footer">Footer</button>
          <select data-dxr="pagenummenu" title="Add a page number to the footer">
            <option value=""># Page number&#8230;</option>
            <option value="currentPage">Page number</option>
            <option value="totalPages">Total pages</option>
            <option value="pageOf">Page X of Y</option>
          </select>
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
    </div>

    <!-- LAYOUT ────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="layout">
      <div class="dxr-group">
        <span class="dxr-glabel">Page setup</span>
        <div class="dxr-row">
          <select data-dxr="margins" title="Page margins for this section">
            <option value="">Margins&#8230;</option>
            <option value="1440,1440,1440,1440">Normal — 1&quot; all round</option>
            <option value="720,720,720,720">Narrow — 0.5&quot; all round</option>
            <option value="1440,1440,1080,1080">Moderate — 1&quot; top/bottom, 0.75&quot; sides</option>
            <option value="1440,1440,2880,2880">Wide — 1&quot; top/bottom, 2&quot; sides</option>
          </select>
          <select data-dxr="orientation" title="Page orientation for this section">
            <option value="">Orientation&#8230;</option>
            <option value="portrait">Portrait</option>
            <option value="landscape">Landscape</option>
          </select>
        </div>
        <div class="dxr-row">
          <select data-dxr="pagesize" title="Paper size for this section">
            <option value="">Size&#8230;</option>
            <option value="12240,15840">Letter — 8.5 &times; 11&quot;</option>
            <option value="12240,20160">Legal — 8.5 &times; 14&quot;</option>
            <option value="11906,16838">A4 — 210 &times; 297 mm</option>
            <option value="15840,24480">Tabloid — 11 &times; 17&quot;</option>
          </select>
          <span class="dxr-note" data-dxr="pagesetupnote"></span>
        </div>
      </div>

      <div class="dxr-group">
        <span class="dxr-glabel">Paragraph</span>
        <div class="dxr-row">
          <label class="dxr-toggle">Before <input data-dxr="spacebefore" type="number" min="0" step="1" placeholder="pt" title="Space before the paragraph, in points" /></label>
          <label class="dxr-toggle">After <input data-dxr="spaceafter" type="number" min="0" step="1" placeholder="pt" title="Space after the paragraph, in points" /></label>
        </div>
        <div class="dxr-row">
          <select data-dxr="specialindent" title="First-line or hanging indent">
            <option value="">Special&#8230;</option>
            <option value="none">(none)</option>
            <option value="firstLine">First line</option>
            <option value="hanging">Hanging</option>
          </select>
          <label class="dxr-toggle">by <input data-dxr="specialindentval" type="number" min="0" step="0.1" value="0.5" title="Indent amount, in inches" /> in</label>
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

      <div class="dxr-group">
        <span class="dxr-glabel">View</span>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="paginated" type="checkbox" /> Page view</label>
        </div>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="headerfooter" type="checkbox" checked /> Headers &amp; footers</label>
        </div>
      </div>
    </div>

    <!-- REFERENCES ───────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="references">
      <div class="dxr-group">
        <span class="dxr-glabel">Table of contents</span>
        <div class="dxr-row">
          <button type="button" data-dxr="toc" title="Insert a table of contents built from the document's headings">&#9776; Table of Contents</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Footnotes</span>
        <div class="dxr-row">
          <button type="button" data-dxr="footnote" title="Cite a new footnote at the caret">Insert Footnote</button>
          <button type="button" data-dxr="endnote" title="Cite a new endnote at the caret">Insert Endnote</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Footer fields</span>
        <div class="dxr-row">
          <button type="button" data-dxr="pagenum" title="Add a page-number field to the footer">Page number</button>
          <button type="button" data-dxr="totalpages" title="Add a total-pages field to the footer">Total pages</button>
        </div>
      </div>
    </div>

    <!-- REVIEW ────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="review">
      <div class="dxr-group">
        <span class="dxr-glabel">Comments</span>
        <div class="dxr-row">
          <button type="button" data-dxr="comment" title="Comment on the selection (Ctrl+Alt+M)">&#128172; New Comment</button>
          <button type="button" data-dxr="commentdelete" class="dxr-danger" title="Delete the active comment thread">Delete</button>
          <button type="button" data-dxr="commentprev" class="dxr-icon" title="Previous comment">&#9650;</button>
          <button type="button" data-dxr="commentnext" class="dxr-icon" title="Next comment">&#9660;</button>
        </div>
        <div class="dxr-row">
          <button type="button" data-dxr="commentresolve" title="Resolve or reopen the active thread">Resolve</button>
          <label class="dxr-toggle"><input data-dxr="showcomments" type="checkbox" checked /> Show comments</label>
          <span class="dxr-note" data-dxr="commentcount"></span>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Tracking</span>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="trackchanges" type="checkbox" /> Track changes</label>
        </div>
        <div class="dxr-row">
          <select data-dxr="markup" title="How revisions are shown">
            <option value="all">All markup</option>
            <option value="none">No markup</option>
          </select>
          <input data-dxr="author" type="text" placeholder="Author" title="Name stamped on your comments and revisions" style="width: 110px;" />
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Changes</span>
        <div class="dxr-row">
          <button type="button" data-rev="accept" title="Accept the change at the caret">&#10003; Accept</button>
          <button type="button" data-rev="reject" title="Reject the change at the caret">&#10007; Reject</button>
          <button type="button" data-rev="prev" class="dxr-icon" title="Previous change">&#9650;</button>
          <button type="button" data-rev="next" class="dxr-icon" title="Next change">&#9660;</button>
        </div>
        <div class="dxr-row">
          <button type="button" data-rev="acceptall" title="Accept every tracked change">Accept all</button>
          <button type="button" data-rev="rejectall" title="Reject every tracked change">Reject all</button>
          <span class="dxr-note" data-dxr="revcount"></span>
        </div>
      </div>
    </div>

    <!-- VIEW ──────────────────────────────────────────────────────────────── -->
    <div class="dxr-panel" data-panel="view">
      <div class="dxr-group">
        <span class="dxr-glabel">Views</span>
        <div class="dxr-row">
          <button type="button" data-dxr="viewpage" title="Page view — real page boxes, headers and footers in the margins">&#128441; Page view</button>
          <button type="button" data-dxr="viewweb" title="Continuous view — one sheet, no page breaks">&#9776; Continuous</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Zoom</span>
        <div class="dxr-row">
          <select data-dxr="zoom" title="Zoom">
            <option value="0.5">50%</option>
            <option value="0.75">75%</option>
            <option value="0.9">90%</option>
            <option value="1" selected>100%</option>
            <option value="1.25">125%</option>
            <option value="1.5">150%</option>
            <option value="2">200%</option>
          </select>
          <button type="button" data-dxr="zoomfit" title="Fit the page to the window">Fit width</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Show</span>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="showrail" type="checkbox" checked /> Status bar</label>
          <label class="dxr-toggle"><input data-dxr="showhint" type="checkbox" checked /> Editing hint</label>
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
      <div class="dxr-group">
        <span class="dxr-glabel">Merge</span>
        <div class="dxr-row">
          <button type="button" data-tt="mergeRight" title="Merge this cell with the one to its right">Merge right</button>
          <button type="button" data-tt="mergeDown" title="Merge this cell with the one below">Merge down</button>
          <button type="button" data-tt="unmerge" title="Split a merged cell back into its cells">Split</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Design</span>
        <div class="dxr-row">
          <button type="button" data-tt="borders" title="Draw all borders">Borders</button>
          <button type="button" data-tt="noborders" title="Remove all borders">No borders</button>
          <span class="dxr-swatch" title="Cell shading">
            <button type="button" class="dxr-icon" data-dxr="shadebtn" aria-label="Cell shading">&#9638;</button>
            <span class="dxr-swatch-bar" data-dxr="shadebar" style="--dxr-swatch: #fef08a;"></span>
            <input type="color" data-dxr="shade" value="#fef08a" aria-label="Pick a cell shading" />
          </span>
          <button type="button" data-tt="noshade" title="Remove cell shading">No shading</button>
        </div>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="repeatheader" type="checkbox" /> Repeat header row</label>
          <button type="button" data-tt="delTable" class="dxr-danger" title="Delete the whole table">Delete table</button>
        </div>
      </div>
    </div>

    <!-- HEADER & FOOTER (contextual) ────────────────────────────────────────
         Word's Header & Footer Tools: the two option checkboxes, page-number fields,
         navigation between the stories, and Close. -->
    <div class="dxr-panel" data-panel="headerfooter">
      <div class="dxr-group">
        <span class="dxr-glabel">Options</span>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="hfFirst" type="checkbox" /> Different first page</label>
        </div>
        <div class="dxr-row">
          <label class="dxr-toggle"><input data-dxr="hfEven" type="checkbox" /> Different odd &amp; even pages</label>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Insert</span>
        <div class="dxr-row">
          <button type="button" data-hf="pagenum" title="Insert the page number at the end of this line">Page number</button>
          <button type="button" data-hf="totalpages" title="Insert the total page count">Total pages</button>
          <button type="button" data-hf="pageof" title="Insert &quot;Page X of Y&quot;">Page X of Y</button>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Navigation</span>
        <div class="dxr-row">
          <button type="button" data-hf="header" title="Go to the header">Go to Header</button>
          <button type="button" data-hf="footer" title="Go to the footer">Go to Footer</button>
        </div>
        <div class="dxr-row">
          <span class="dxr-note" data-dxr="hfstory">Header</span>
        </div>
      </div>
      <div class="dxr-group">
        <span class="dxr-glabel">Close</span>
        <div class="dxr-row">
          <button type="button" data-hf="close" title="Return to the document body (Esc)">&#10005; Close Header and Footer</button>
        </div>
      </div>
    </div>
  </div>

  <!-- Find & replace strip, under the ribbon. -->
  <div class="dxr-findbar" data-dxr="findbar" hidden>
    <input data-dxr="findtext" type="search" placeholder="Find" aria-label="Find" />
    <button type="button" class="dxr-icon" data-dxr="findprev" title="Previous match (Shift+Enter)">&#9650;</button>
    <button type="button" class="dxr-icon" data-dxr="findnext" title="Next match (Enter)">&#9660;</button>
    <span class="dxr-findcount" data-dxr="findcount"></span>
    <label class="dxr-toggle"><input data-dxr="findcase" type="checkbox" /> Match case</label>
    <span data-dxr="replacegroup">
      <input data-dxr="replacetext" type="text" placeholder="Replace with" aria-label="Replace with" />
      <button type="button" data-dxr="replaceone" title="Replace this match">Replace</button>
      <button type="button" data-dxr="replaceall" title="Replace every match">Replace all</button>
    </span>
    <span class="dxr-spacer" style="flex:1"></span>
    <button type="button" class="dxr-icon" data-dxr="findclose" title="Close (Esc)">&#10005;</button>
  </div>
</div>

<div class="dxr-scroll" data-dxr-scroll>
  <p class="dxr-hint" data-dxr-hint></p>
  <div class="dxr-surface" data-dxr="editor" data-dxr-surface data-view="continuous" data-comments="on"></div>
</div>

<!-- Status bar — live engine state (the anchor rail), Word's page/word cells, and zoom. -->
<div class="dxr-rail" data-dxr-rail>
  <span class="dxr-cell"><span class="dxr-k">anchor</span><span class="dxr-v" data-dxr="railAnchor">&#8212;</span></span>
  <span class="dxr-cell"><span class="dxr-k">blocks</span><span class="dxr-v" data-dxr="railBlocks">&#8212;</span></span>
  <span class="dxr-cell"><span class="dxr-k">session</span><span class="dxr-v" data-dxr="railSession">&#8212;</span></span>
  <span class="dxr-cell"><span class="dxr-k">last op</span><span class="dxr-v" data-dxr="railOp">&#8212;</span></span>
  <span class="dxr-cell"><span class="dxr-v dxr-plain" data-dxr="pageinfo">&#8212;</span></span>
  <span class="dxr-cell"><span class="dxr-v dxr-plain" data-dxr="wordcount">&#8212;</span></span>
  <span class="dxr-cell dxr-cell-end">
    <span class="dxr-zoom">
      <button type="button" data-dxr="zoomout" title="Zoom out" aria-label="Zoom out">&#8722;</button>
      <select data-dxr="zoomlevel" title="Zoom" aria-label="Zoom level">
        <option value="0.5">50%</option>
        <option value="0.75">75%</option>
        <option value="0.9">90%</option>
        <option value="1" selected>100%</option>
        <option value="1.25">125%</option>
        <option value="1.5">150%</option>
        <option value="2">200%</option>
      </select>
      <button type="button" data-dxr="zoomin" title="Zoom in" aria-label="Zoom in">+</button>
    </span>
  </span>
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

<!-- Link popover, anchored to the Insert tab's Link button. -->
<div class="dxr-pop" data-dxr="linkpop">
  <form data-dxr="linkform">
    <label>Address <input data-dxr="linkurl" type="url" placeholder="https://" required /></label>
    <div class="dxr-row">
      <button type="button" data-dxr="linkcancel">Cancel</button>
      <button type="submit">Insert link</button>
    </div>
  </form>
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
  "Click any paragraph, heading, table cell, footnote, header or footer to edit it. " +
  "<kbd>Enter</kbd> splits a block, <kbd>Backspace</kbd> at the start merges it into the previous one, " +
  "<kbd>Ctrl</kbd>+<kbd>Z</kbd> undoes. Select text and press <b>New Comment</b> to annotate it. Only the block you " +
  "changed re-renders — everything else keeps full fidelity, and <b>Save</b> writes a lossless .docx.";

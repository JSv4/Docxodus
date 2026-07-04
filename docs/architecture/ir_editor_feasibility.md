# IR-Powered DOCX Editor — Feasibility & PoC Design

Status: **Design / feasibility** (pre-implementation). Target deliverable of the
first effort is a written feasibility spec (this document) plus a **focused,
runnable proof-of-concept** that turns "could we?" into a measured yes/no.

Branch: `feat/ir-editor-feasibility-poc`.

---

## 1. Summary / verdict

**Can the Docxodus IR power a performant, format-faithful browser DOCX editor
that "renders pages and populates with editable blocks"? Yes — but not in the
literal framing.** A source-level scan (verified with file/line citations)
establishes that the IR (`Docxodus/Ir/`) is the **wrong layer to be the
editor's model-of-record**, while the surrounding stack already provides ~80% of
an editor:

- The IR has **no IR→OOXML writer** anywhere (`IrWriter`/`IrToWml`/`EmitOoxml`
  grep-empty). The only emitter is `IrMarkdownEmitter`. You cannot edit an
  `IrDocument` and serialize a valid `.docx` from it.
- The diff engine proves the model: `IrMarkupRenderer`
  (`Docxodus/Ir/Diff/IrMarkupRenderer.cs:108-163`) builds tracked-changes output
  by **re-reading the source documents and cloning the original `w:p`/`w:tbl`
  XML via `IrProvenance`** — never by serializing the IR. The IR carries
  *alignment/anchors*; the original OOXML carries *fidelity*.
- The IR is `internal`/experimental, **immutable** (sealed records, no mutation
  API), and deliberately **lossy**: unmodeled run/paragraph/section properties
  survive only as one-way `UnmodeledDigest` hashes; math/charts/SmartArt/VML are
  `IrOpaqueInline` (hash only); `IrOpaqueBlock` is `ElementName` + hash; images
  keep only bytes-hash + EMU extent + alt.

The right architecture is therefore **Option B** (see §11): the live
`DocxSession` (real OOXML, mutated in place, lossless `Save()`) is the
model-of-record; `WmlToHtmlConverter` + the existing `npm/src/pagination.ts`
engine are the render substrate; and the `{#kind:scope:unid}` anchor system is
the addressing/diff overlay only.

The **two real gaps** (not the IR) that stand between this stack and a
responsive editor:

1. **Incremental rendering.** Every `DocxSession` mutation re-projects the whole
   document (`DocxSession.cs:4184` `ProjectScope` runs
   `WmlToMarkdownConverter.Convert` over the entire doc; `MarkdownPatch` carries
   whole-doc markdown), and getting an edit *on screen* requires a full
   `convertDocxToHtml` re-conversion (~0.7–2.4s). There is **no per-block HTML
   patch path** for content edits today.
2. **Worker offload.** The full editing surface is **main-thread-only**; the Web
   Worker exposes only convert/compare/metadata + annotation ops.

The **page-fidelity ceiling** is honest and inherent: DOCX is *reflowable*
(Word computes layout at render time; nothing is stored), and there is **no
layout engine** in this stack. `pagination.ts` delivers *block-flow* pages
(paragraphs/tables flow whole; cannot split mid-line; page counts ≈ Word, not
exact; font-dependent). True line-exact WYSIWYG would require integrating an
external renderer (e.g. LibreOffice headless) and is explicitly out of scope.

## 1a. Proof status — PoC results (the gate is cleared)

The single make-or-break unknown (§5 problem 1 / §6.1: *can a single block render
faithfully out of whole-document context?*) is now **proven yes**, at two levels:

- **C# unit (`Docxodus.Tests/HtmlConversionOpsTests.cs` `HCO050`)** — for HC006 and
  HC001, `HtmlConversionOps.RenderBlockHtml(anchor)` produces an element whose tag
  and visible text match the corresponding `data-anchor` element from a full render
  (the oracle), across up to 12 paragraphs/headings per doc.
- **Browser / WASM (`npm/tests/render-block.spec.ts`)** — the same equivalence holds
  across the real WASM boundary in Chromium, using the actual DOM — i.e. the editor's
  incremental per-block re-render path.
- **Full editor loop (browser, same spec)** — open `DocxSession` → `Project` → edit a
  body paragraph via `ReplaceText` → re-render **only that block** from the live session
  (`DocxSessionBridge.RenderBlockHtml`) → the edit is visible in the re-rendered block.
  The complete incremental round-trip, proven in Chromium.

**One Unid scheme, confirmed.** `convertDocxToHtml`(stampAnchors) ↔ `DocxSession` ↔
`RenderBlock` all derive anchors from `AssignToAllElementsDeterministic` over the raw
main-document root, so a full-render `data-anchor` resolves unchanged on the live session
path (`HCO052`). A DOM block's anchor is a valid session/render anchor.

**Latency (measured, `HCO052`, HC031 — a complex 42 KB doc, 20 blocks):**
- stateless `RenderBlockHtml(bytes,…)`: **26.5 ms/block**
- session-attached `RenderBlockHtml(session,…)`: **10.4 ms/block (2.55×)** — resolves
  against the live document, no byte re-open / whole-doc Unid pass.
- Both are **~70–230× faster** than the ~0.7–2.4 s full `convertDocxToHtml` baseline, and
  well under the <150 ms/edit target. The session-attached path is committed and exposed
  through WASM/npm (`DocxSession.renderBlock`).

**Editor MVP built & proven (`npm/src/editor.ts`, `npm/tests/editor.spec.ts`).** A
framework-agnostic `DocxEditor` (pure TS) renders a faithful document with `data-anchor`
blocks, makes projection-addressable paragraphs/headings `contenteditable`, and on blur
commits via `DocxSession` then re-renders **only that block** from the live session.
Browser test on HC031: render **90 editable blocks** → edit one → only that block
re-renders → save (39 KB) reopens with the edit persisted **and** untouched content intact
(lossless). Two findings that refine the design:
- **Anchors are stable within a live session.** `ReplaceText` mutates runs in place and
  does NOT re-derive the paragraph's Unid (deterministic assignment only fills *absent*
  Unids), so the edited block keeps its anchor mid-session — the §6.3 stable-key concern is
  milder than assumed (a client-side key registry is not needed for within-session edits;
  it only matters across save/reopen, where the content-hash Unid legitimately changes).
- **HTML stamps more anchors than the projection indexes.** Paragraphs inside opaque
  tables get a `data-anchor` but are not individually addressable via the markdown/
  `ReplaceText` path, so the editor only makes projection-addressable blocks editable;
  table-cell editing (via `ReplaceCellContent`) is future work.

**Block-flow pagination DONE.** `DocxEditor { paginated: true }` renders via the
converter's `PaginationMode` + `pagination.ts` so blocks flow into real `.page-box` page
boxes (margins/headers), and those page blocks stay editable with the same incremental
re-render loop (browser test confirms page boxes render + an in-page edit re-renders).
This satisfies the original "render pages and populate with editable blocks" goal — the
editor wires only the visible page container (pagination clones blocks; hidden originals
stay in `#pagination-staging`).

**Remaining work is tracked, prioritized by impact, in `ir_editor_roadmap.md`** (M1 rich
in-block editing → M2 structural editing → M3 worker offload → M4 re-paginate-on-edit →
M5 toolbar/undo → M6 tracked-changes → M7 tables → M8 React → M9 render fidelity).

What was built to clear it (committed):
- `WmlToHtmlConverterSettings.StampAnchors` → stamps `data-anchor=Unid` on
  `p`/`h*`/`li`/`table` (`WmlToHtmlConverter.cs`).
- `HtmlConversionOps.RenderBlockHtml(bytes, anchor, options)` → renders one block via a
  **throwaway document** that copies the source's styles / numbering / theme / font /
  **settings** parts (the converter requires `DocumentSettingsPart`), then runs the
  standard converter and extracts the stamped block.
- WASM `DocumentConverter.RenderBlockHtml` + `stampAnchors` on `ConvertDocxToHtmlComplete`;
  npm `renderBlockHtml()` + `ConversionOptions.stampAnchors`.

Findings / refinements vs the original design:
- **Unid authority must be singular.** The first attempt resolved the anchor via the
  markdown projector while the full render stamped via `AssignToAllElementsDeterministic`
  — the two unid schemes diverged. Fix: both the stamp and `RenderBlockHtml` assign Unids
  with the **identical call** (`AssignToAllElementsDeterministic` over the main-document
  root); `RenderBlockHtml` then resolves by the `Unid` attribute directly (keying on the
  anchor's unid tail, so a bare unid or a full `kind:scope:unid` both work).
- **`data-anchor` carries the bare unid** today. `renderBlock`/`renderBlockHtml` accept it
  (they key on the unid tail), but DocxSession *edit* ops (`ReplaceText`, etc.) require the
  full `kind:scope:unid`, so the editor maps bare unid → full id via the session projection
  index (cheap — same Unid scheme, suffix match; exercised in the editor-loop test).
  **Plan 2 refinement:** stamp the full `kind:scope:unid` into `data-anchor` so DOM blocks
  are directly DocxSession-addressable with no mapping step.
- **Confirmed degradations (acceptable PoC limits, see §8):** a list item rendered in
  isolation loses numbering *continuation*; an inline image loses its (uncopied) image
  part. The oracle test targets text paragraphs/headings and skips image blocks.
- **Latency:** not yet profiled on a large doc. `RenderBlockHtml` currently re-opens the
  bytes and re-assigns Unids over the whole document per call — fine on small docs, but
  Plan 2 should offer a **session-attached** render (reuse the live `DocxSession` doc) to
  avoid the per-call whole-doc Unid pass.

## 1b. Public API & usage

**Run the demo** (a standalone in-browser editor over a sample doc):

```bash
cd npm
npm run demo       # builds + serves; open http://localhost:8088/editor.html
```

Click a paragraph/heading to edit; Enter splits a block, Backspace at the start merges;
toggle "Paginated" for page boxes; "Save .docx" downloads a lossless document. Source:
`npm/examples/editor.html`.

**Browser editor (npm).**

```ts
import { initialize, DocxEditor } from "docxodus";

await initialize();
const bytes = new Uint8Array(await file.arrayBuffer());

// `exports` is the WASM bridge object (DocxSessionBridge + DocumentConverter).
// In the npm runtime it is the initialized module; in the test harness it is
// `window.Docxodus`.
const editor = DocxEditor.open(container, bytes, exports, {
  paginated: true,   // real page boxes via pagination.ts (false = continuous)
  editable: true,    // contenteditable projection-addressable blocks
  onEdit: ({ anchorId }) => console.log("edited", anchorId),
});

// ...user edits blocks in the DOM; each blur commits + re-renders that block...

const saved: Uint8Array = editor.save();  // lossless DOCX bytes
editor.close();                           // release the WASM session
```

**Single-block render (npm), e.g. to drive your own DOM:**

```ts
import { renderBlockHtml, openDocxSession } from "docxodus";

// Stateless (re-derives anchors over the bytes):
const html = await renderBlockHtml(bytes, "p:body:<unid>");

// Session-attached (faster — resolves against the live edited doc):
const session = await openDocxSession(bytes);
session.replaceText("p:body:<unid>", "new text");
const fresh = session.renderBlock("p:body:<unid>"); // or the bare unid
```

**Full-document render with addressable blocks:**

```ts
const docHtml = await convertDocxToHtml(bytes, { stampAnchors: true, paginationMode: 1 });
// every p/h*/li/table carries data-anchor="<unid>"
```

**.NET.**

```csharp
using Docxodus.Internal;

var opts = new HtmlConversionOptions { StampAnchors = true };
string fullHtml = HtmlConversionOps.ConvertToHtml(bytes, opts);          // data-anchor on blocks

using var session = new DocxSession(bytes);
session.ReplaceText("p:body:<unid>", "new text");
string blockHtml = HtmlConversionOps.RenderBlockHtml(session, "p:body:<unid>",
    new HtmlConversionOptions());                                        // session-attached
```

> The `data-anchor` value is the bare unid; DocxSession *edit* ops want the full
> `kind:scope:unid`. The editor maps bare unid → full id via the session projection index
> (same Unid scheme — suffix match). `RenderBlockHtml`/`renderBlock` accept either form.

## 2. Locked decisions

| Decision | Choice | Consequence |
|---|---|---|
| Page model | **Block-flow paginated** | Real page boxes via `pagination.ts`; no external renderer needed. |
| Save fidelity | **Lossless round-trip** | `DocxSession`'s live OOXML is the *only* viable model-of-record (the lossy IR/markdown cannot round-trip arbitrary docs). |
| Primary use | **Both** (review/redline + authoring) | Tracked-changes (`RenderInline`) available but optional; edit model is **per-block, debounce-committed** (serves both; avoids per-keystroke-through-WASM). |
| Collaboration | **Single-user now, don't foreclose collab** | Build single-user; design the edit-event model so a CRDT/OT layer could sit on the anchor ops later. |
| Deliverable | **Feasibility spec + focused PoC** | This doc + a runnable harness that de-risks the hard parts. |

## 3. Key findings (source-verified)

### 3.1 The IR is addressing/diff, not a save model
See §1. The lossless fidelity source is the *retained original XML*
(`IrProvenance.Element`, only when `RetainSources=true`) — i.e. the OOXML, not
the IR. An editor that mutated the IR alone would have no save path.

### 3.2 `DocxSession` is the real edit backend (and already WASM-surfaced)
`Docxodus/DocxSession.cs` is a stateful, transactional, anchor-addressed
mutation API over a live `WordprocessingDocument`:

- Tiers A–E: text CRUD, structural (`SplitParagraph`/`MergeParagraphs`/
  `InsertParagraph`), formatting (`ApplyFormat`/`SetParagraphStyle`/
  `SetListLevel`), cell content (`ReplaceCellContent`), annotations.
- Each op mutates the in-memory XML **in place** (single-digit ms), takes a
  pre-op snapshot for atomic rollback, and returns a typed `EditResult`
  (`Success`, `EditError(code,…)`, `Created`/`Removed`/`Modified` anchor lists,
  `MarkdownPatch`).
- Bounded undo/redo (`UndoRing`, depth 50), tracked-changes `RenderInline` mode
  (edits land as `w:ins`/`w:del` with author attribution), and a `Raw.*`
  OOXML escape hatch.
- Surfaced over WASM via `DocxSessionBridge` (int handle pool) and consumed by
  `npm/src/session.ts`.

**Granularity that works: per-block transactions** (paragraph/cell), committed
on blur or a debounced typing burst coalesced into one `ReplaceTextAtSpan`.
Per-keystroke is not viable through the WASM/projection path.

### 3.3 The anchor/Unid spine
`{#kind:scope:unid}` (`IrAnchor.cs`, public `Anchor` in
`WmlToMarkdownConverter.cs`) is shared across read (markdown), write
(`DocxSession`), and diff (`DocxDiff` `leftAnchor`/`rightAnchor`). Unids are
content-addressable (SHA over content+position, `UnidHelper.cs`), so **sibling
blocks stay stable when one block is edited** — but the **edited block's own
Unid changes** (it hashes its text), freshly-inserted/split blocks get **random
Guids**, and `Save()` **strips Unids by default** (`PersistAnchorIds=false`).
Implication: an editor must key React on a **client-side stable id**, not the
raw Unid (see §6.3).

### 3.4 Render assets that already exist
- `WmlToHtmlConverter` (~8k lines): style-cascade→CSS, Word-faithful for the
  common ~80% (runs, tables with border resolution, lists, headings, theme
  colors, `@page` CSS). **Does not stamp `data-anchor` today** (references Unid
  0 times) — an additive change is needed.
- `npm/src/pagination.ts` (1412 lines): a working client-side pagination engine
  — measures rendered blocks via `getBoundingClientRect`, flows them into fixed
  page boxes with margin-collapsing, keep-with-next, header/footer registry, and
  footnote splitting/continuation. `<PaginatedDocument>` wires it up.
- `DocxSession.GetSectionInfo(anchor)` → true page setup (size, margins,
  landscape, columns, header/footer part URIs) — the page-box frame to draw.
- `ExternalAnnotationProjector` proves an *incremental DOM-patch loop* is
  possible (0.3ms single-add) — though it addresses by fragile text-search and
  re-parses the whole HTML string per op.

### 3.5 The page-fidelity ceiling
No glyph metrics / line-box / line-breaking / page-breaking in C#/IR. The only
place a block acquires a real position is `pagination.ts` (the browser does the
line breaking). PAWLS token X/Y/W/H from `OpenContractExporter` are **fabricated**
(6pt/char, 12pt/line) — decorative, never to be aligned to a rendered view. Page
counts from every C# API are heuristic estimates. Computed layout is
browser-only and one-way; nothing persists page/line geometry back to the doc.

## 4. Target architecture (Option B)

**Hard invariant: the DOM is a projection; the in-WASM `DocxSession` is the only
source of truth.** Optimistic local edits are always reconciled against the
authoritative re-render; the worker wins.

```
┌─────────────────────────── Main thread (React) ───────────────────────────┐
│  Editor shell ──▶ Page list ──▶ Block views (contenteditable, data-anchor) │
│      ▲                                   │ edit-intent (debounced/on blur)  │
│      │ {anchorId, html, editResult}      ▼                                  │
│  Anchor↔stable-key registry        postMessage                             │
│  Paginator (pagination.ts)              │                                   │
└─────────────────────────────────────────┼──────────────────────────────────┘
                                           ▼
┌──────────────────────── Web Worker (WASM) ─────────────────────────────────┐
│  Editor worker service                                                      │
│    DocxSession (live OOXML = model-of-record, lossless Save)                │
│    apply op → EditResult (Created/Removed/Modified)                         │
│    RenderBlockHtml(anchor) → faithful HTML for one block   ◀── NEW          │
└─────────────────────────────────────────────────────────────────────────────┘
```

**Layers**

| Layer | What | Where |
|---|---|---|
| Model-of-record | Live `DocxSession`; real bytes mutated in place; lossless `Save()` | WASM, in Worker |
| Addressing spine | `{#kind:scope:unid}`; every editable block carries `data-anchor` | shared |
| Render substrate | `WmlToHtmlConverter` → faithful HTML; **new `data-anchor` stamp** | C# (additive) |
| Pagination | `pagination.ts` flows blocks into page boxes | TS, main thread |
| Incremental render | **new `RenderBlockHtml(anchor)`**; patch only that DOM node | C# + TS |
| Edit surface | React; per-block `contenteditable`; debounce-committed | TS, main thread |

**Data flow for one edit**

1. User types in a block → optimistic local DOM update (caret preserved).
2. On commit (debounce ~300ms / blur) → main thread posts an edit-intent to the
   worker (anchor + op + args).
3. Worker applies the `DocxSession` op → `EditResult`
   (`Created`/`Removed`/`Modified`).
4. Worker calls `RenderBlockHtml` on the affected block(s) → posts
   `{anchorId, html, editResult}` back.
5. Main thread reconciles: swap the block's subtree, update the
   stable-key↔unid map from `EditResult`, re-paginate from the affected page
   forward.

Reconciliation uses `EditResult` deltas + the render-block path; it **does not
depend on `MarkdownPatch`** (whole-doc re-projection).

## 5. The three hard problems & de-risking

1. **Incremental per-block render** (gap #1, biggest unknown). Build
   `RenderBlockHtml(anchor)`: run the converter pipeline over a single block's
   XML subtree with styles/numbering context resolved once. **Risk:** the
   converter assumes whole-doc context (styles, numbering, sectPr). Mitigation:
   resolve styles/numbering once at session/render-context construction and
   render the block element against the already-fabricated CSS classes.
   **Fallback baseline to beat:** full `convertDocxToHtml` + DOM-diff by
   `data-anchor` (cheaper to build, ~1s/edit — too slow, used only as the
   measurement baseline).
2. **Worker offload** (gap #2). Extend `docxodus.worker.ts` + `worker-proxy.ts`
   (or a new `editor.worker.ts`) to carry the content-editing `DocxSession` ops
   (`replaceText`/`applyFormat`/`split`/`merge`/`replaceCellContent`/`save`/
   `undo`/`redo`) plus the new render-block op, so editing never blocks the UI.
3. **Anchor stability / React keys** (§3.3). Maintain a **client-side stable
   block key** (a GUID per block assigned at first render) mapped to the
   *current* Unid, updated from `EditResult.Created`/`Removed`/`Modified`. React
   keys = the stable key, never the Unid → the actively-typed block never
   remounts (caret preserved). Reconciliation must be correct across
   split/merge (a split produces one `Modified` + one `Created`; a merge
   produces one `Modified` + one `Removed`).

Plus **incremental re-pagination**: after a block changes height, re-paginate
from the affected page forward. `pagination.ts` measures via the live DOM (no
WASM); for the PoC a full re-paginate is acceptable to measure, with
forward-only incremental reflow as a follow-up.

## 6. PoC scope — the runnable yes/no

A minimal React harness (served from the npm build / Playwright-driveable) that:

1. Loads a real, non-trivial `.docx` from `TestFiles/` — one with headings, a
   list, a table, headers/footers, and ideally an **unmodeled construct** (text
   box / equation / SmartArt) to test lossless save.
2. Renders it **block-flow paginated** (`WmlToHtmlConverter` + `pagination.ts`)
   with `data-anchor` on every block.
3. Makes 3 representative edits via `DocxSession` in the worker:
   - **text** edit in a paragraph (`ReplaceTextAtSpan`),
   - **format** toggle (bold on a span, `ApplyFormat`),
   - **structural** edit (`SplitParagraph` or `InsertParagraph`).
4. **Incrementally re-renders only the changed block(s)** and re-paginates —
   measuring per-edit visible-update latency.
5. **Verifies lossless save**: reopen `Save()` bytes, assert untouched/unmodeled
   content is byte-preserved and the edits landed.
6. (Optional) toggles tracked-changes (`RenderInline`) and shows `w:ins`/`w:del`.

### 6.1 Success criteria (= the feasibility verdict, made runnable)

- Per-edit visible update **materially beats** full re-convert (target
  **<150ms** vs ~1s baseline).
- **Lossless save verified** on a doc containing unmodeled content.
- **Caret preserved** across an edit (stable-key reconciliation works).
- Block-flow pages render recognizably (margins / headers & footers / page
  boxes).

If these hold → "yes, and here is the architecture + measured costs." If
incremental render cannot be made fast *and* faithful → that is the documented,
honest "not yet."

### 6.2 Non-goals (YAGNI for this pass)

Line-exact WYSIWYG / mid-paragraph page splits; table structure editing
(insert/delete rows/cols, cell merge), image insert, footnote/comment creation
(no write API — viewable only); collaboration/CRDT; a full formatting toolbar or
style picker; optimizing `ProjectScope`. The edit-event model is designed so a
CRDT/OT layer *could* sit on the anchor ops later, but none is built now.

### 6.3 Anchor/key handling (explicit, because it is a known trap)

- React `key` = client-side stable GUID, assigned per block at first render.
- A `Map<stableKey, currentUnid>` and reverse index are updated on every
  `EditResult`.
- The actively-edited block keeps its stable key even though its Unid changes —
  no remount, caret intact.
- `Save()` is called with default `PersistAnchorIds=false` (no bloat); on reopen,
  anchors are re-derived deterministically and the editor re-attaches by
  re-projection. (Persisting ids is an opt-in follow-up, not in the PoC.)

## 7. Units & boundaries (isolated, testable)

- **`RenderBlockHtml` bridge** — *new* C#: a method on the HTML-conversion facade
  (`Docxodus/Internal/HtmlConversionOps.cs`) + a `DocumentConverter`/session WASM
  `[JSExport]`: `(sessionHandle | bytes, anchor) → html`. The *one* substantive
  new C# surface. Plus an additive `data-anchor=Unid` stamp on block elements in
  `WmlToHtmlConverter`.
- **Editor worker service** (`editor.worker.ts`) — owns the session; message
  contract for open / edit-ops / render-block / save / undo. Testable via its
  message protocol.
- **Anchor↔stable-key registry** (TS, pure) — consumes `EditResult` deltas;
  unit-testable in isolation.
- **Block view** (React) — renders one block's HTML, `contenteditable`, emits
  edit intents.
- **Paginator** — existing `pagination.ts`, lightly extended for incremental
  reflow.
- **Editor shell** (React) — orchestrates the page/block list and the worker
  round-trip.

Per the repo's single-owner-facade convention, any new cross-boundary surface
(`RenderBlockHtml`) lands in its `*Ops` facade first, then ripples to the WASM
bridge + TS client.

## 8. Risks & unknowns the PoC will settle

- **Single-block faithful render out of whole-doc context** (the big one —
  styles, numbering continuation, list markers, sectPr).
- **Incremental re-pagination cost** and correctness across page boundaries.
- **Worker round-trip overhead** per edit (postMessage + JSON + WASM).
- **Stable-key reconciliation correctness** across split/merge.
- Converter fidelity holes (text boxes dropped, math/charts/SmartArt opaque,
  most field codes) → these render as holes; the editor must treat them as
  viewable-not-editable and must never silently drop them on save (lossless save
  via `DocxSession` guarantees this, since the bytes are never round-tripped
  through the lossy projection).

## 9. Appendix — architecture options considered

| Option | Verdict | Why |
|---|---|---|
| **A. IR as model-of-record** | ❌ Fatal | No IR→OOXML writer; IR is immutable + lossy + internal. No save path. |
| **B. HTML-render + anchor-addressed edit overlay** | ✅ **Chosen** | Reuses the converter (fidelity) + `pagination.ts` (pages) + `DocxSession` (lossless transactional edits); IR/anchors as addressing only. |
| **C. DocxSession-backed galley/markdown blocks** | Stepping stone | Lowest-friction (public + WASM today) but lossy text view, no pages — contradicts the block-flow-paginated requirement. Useful as an even-smaller fallback if §6 stalls. |
| **D. True paginated WYSIWYG** | ❌ Out of scope | Needs a layout engine that exists nowhere here; realistically an external-renderer integration. |

## 10. References (cited source)

- IR core & no-writer: `Docxodus/Ir/IrDocument.cs`, `IrBlocks.cs`,
  `IrInlines.cs`, `IrFormats.cs`, `IrReader.cs`,
  `Docxodus/Ir/Diff/IrMarkupRenderer.cs:108-163`, `docs/architecture/document_ir.md`.
- Anchors/Unid: `Docxodus/Ir/IrAnchor.cs`, `Docxodus/UnidHelper.cs`,
  `Docxodus/WmlToMarkdownConverter.cs`, `docs/architecture/markdown_projection.md`.
- Edit backend: `Docxodus/DocxSession.cs` (esp. `:4184` `ProjectScope`),
  `Docxodus/Internal/DocxSessionOps.cs`, `docs/architecture/docx_mutation_api.md`.
- TS/WASM surface: `npm/src/{index,session,pagination,react,docxodus.worker,worker-proxy}.ts`,
  `wasm/DocxodusWasm/{DocxSessionBridge,DocumentConverter}.cs`,
  `docs/architecture/{ui_responsiveness,wasm-optimization-plan,profiling-results}.md`.
- Rendering/pagination: `Docxodus/WmlToHtmlConverter.cs`,
  `Docxodus/ExternalAnnotationProjector.cs`, `Docxodus/OpenContractExporter.cs`,
  `docs/architecture/{docx_converter,wml_to_html_converter_gaps,incremental_annotation_overlay,paginated_headers_footers}.md`.

# GitHub Pages demo

Seven static pages, no application server. All host the **same** editor
surface — `createRibbonEditor` from the pinned `docxodus@12.0.0` embed bundle on
jsDelivr — and differ only in how much of the page belongs to the editor:

| Page | What it is |
|------|------------|
| `index.html` | The landing page. Hero, capability cards, the embed dialog, and a live demo inside its frame. This is the URL to share; its Open Graph metadata is what social platforms read. **Below 620px the frame mounts THE DOCX ARCADE instead of the plain editor** — same shipped surface, a game running on one of its paragraphs. A phone visitor is not going to draft a contract on a 390px screen, but they will play the thing, and it demonstrates the engine harder: a document rewritten and re-rendered ten times a second, still saving as a real `.docx`. The choice is taken in a `<head>` script, before first paint, so the copy framing the frame never advertises the demo the page did not mount; `?demo=editor` / `?demo=arcade` pin either on any screen, and the page links to the other one both ways. The nav links are a horizontal scroll strip on a phone rather than `display: none` — they point at the other demo pages, which is what a phone visitor most wants. |
| `app.html` | The editor full-bleed, nothing around it. The useful thing to open on a phone. |
| `player.html` | The compact iframe target, sized for ~480 × 480. Boots on tap so a feed iframe never streams a .NET runtime unasked, and pins the surface's compact layout. |
| `observatory.html` | The DOCX Observatory inside the live editor: procedural ASCII phenomena animated onto a Word paragraph in the editor's own session (`raw.replaceXml` + `editor.refresh()`). Pause — or click the water — and it is only a document: edit with the ribbon, Undo rewinds frame by frame, Save downloads the caught wave. The phenomena and frame loop are demo content, not library machinery: they live in `ascii-scenes.js` **in this directory** (also imported, via the test-webroot copy, by the two `npm/examples/ascii-animation*` pages), so `?engine=` pins the library alone and the scenes version with the site. Needs `DocxEditor.refresh()`, which 9.5.0 predates — it was pinned to `docxodus@9.6.0` ahead of that release and healed on its own when it published; it now shares the pin with its siblings. |
| `arcade.html` | THE DOCX ARCADE — four playable games rendered through the editor's ordinary incremental document path. The text cartridges replace colored runs in one Word paragraph; **DOOM** runs id Software's GPL engine on Freedoom's BSD IWAD and replaces the media payload of one native **320 × 200 full-color inline DOCX image** through the public session API. It is guarded at the ten-visible-FPS design point, including browser decode and presentation. Four separate static **18pt document paragraphs** keep the complete controls legible. Doom's world remains in its WebAssembly heap, so its honest document round trip is pause, copy/paste, undo and save; the other cartridges additionally re-parse edited terrain on resume. The games live in `ascii-arcade.js`; Doom lives in the GPL `doom-cart.js`, with pinned engine/IWAD URLs documented below. Specs in `npm/tests/demo-arcade-doom.spec.ts` prove exact framebuffer pixels, readable displayed HUD dimensions, real play, sustained visible throughput, input, copy/paste, and save/reopen fidelity. **DOOM's image now renders on this page** — the pin is `docxodus@12.0.0`, which fixes the gap that blocked it through 10.0.0: up to and including that version the single-block renderer cloned a block's XML into a throwaway shell without copying the referenced media part, and its converter settings carried no image handler, so `WmlToHtmlConverter` omitted the `w:drawing` and the frame paragraph refreshed to blank even though the image was genuinely in the package the whole time (Save and reopen showed it). `paintImage`'s three-frame capability probe — not a version check — is what caught this originally, and it stays in place as a guard against a future regression rather than being removed now that the surface proves itself. `?cart=e1m1` / `?cart=dungeon` / `?cart=platformer` (run-based, no image) always played regardless. |
| `golf.html` | DOCX GOLF — course play on the editing surface itself, the inverse bet from the arcade: where the arcade painted frames INTO one paragraph through `raw.replaceXml` (the escape hatch), golf makes the real clubs the game. Six holes, each a start document loaded into the live ribbon editor and a target document built beside it; the referee is the comparison engine — a hole is CLEARED when `docxDiffGetRevisions` between your document and the target returns **zero revisions**, and the caddie panel phrases whatever revisions remain as the work left (`remove "Purchasr"`, `move "Governing Law"`, `reformat "Duties" (style)`). Holes escalate across the surface: a one-word fix, clause reordering, scoped defined-term conformance, heading styles (the Style dropdown is the club — the diff cares about `w:pStyle`, not just words), and a table hole (fix a cell, delete a duplicated row from the table toolbar), plus a footnote hole played with Insert → Footnote (the referee reads note parts too). On a phone the caddie collapses to its head strip behind a toggle, and a stuck player can concede with "Show me" — the caddie plays the content-addressed reference line and the scorecard marks the hole assisted. Strokes are counted from the document, not the toolbar: the driver fingerprints `session.save()` on a poll, so one committed burst of editing is one swing, and undo counts. Par/birdie/bogey scoring, a Target view, and a live redline view (`docxDiffCompare` → `convertDocxToHtml`) round out the caddie. The game lives in `docx-golf.js` in this directory (same `?engine=` split as its siblings); its pure logic is tested by `tools/docx-golf.test.mjs`, and `npm/tests/demo-golf.spec.ts` keeps the course honest the way the engine's own evals are kept honest — no hole starts solved, and every hole's content-addressed reference solution reaches zero revisions within par. Boots on tap when iframed. |
| `redline.html` | REDLINE THEATER — the agent-facing surface as the show. Three counsel negotiate a Master Services Agreement, and every edit you watch land is dispatched from a real JSON-RPC 2.0 `tools/call` frame in the shape `docxodus-mcp` accepts over stdio, streamed on a wire console beside the document. Nothing here renders a diff: the session runs in `render_inline` recording mode, so each call writes native `w:ins`/`w:del` into the live package and the editor repaints only the block that changed — what you are watching IS the file that downloads, and the three counsel are three values of the session's revision author, so it is genuinely per-reviewer. Ends by proving itself: `proveRedlineReversibility` plus a `docxDiffGetRevisions` pass confirm accept-all reaches the negotiated final and reject-all restores the baseline with **zero content differences**. One scripted call is *expected to be refused* — converting the carve-outs to a Word-native list has no reversible tracked-change encoding, so the engine declines it rather than writing a mark reject-all could not undo, and the wire shows that in amber. Speed is a display choice the HUD is honest about (1× / 4× / MAX, with measured p50 and calls/s); the run is ~3 s of wall clock at 4×, and the median tool call is under 6 ms. The show lives in `redline-theater.js` and the browser MCP endpoint in `mcp-wire.js`, both in this directory (same `?engine=` split as its siblings), plus `?speed=read\|brisk\|max` and `?autorun=1`. Specs: `npm/tests/demo-redline.spec.ts` (fourteen browser assertions across both modes, including the reversibility proof, the latency budget, and that a deeper diff pipeline actually costs more) and `docs/demo/tools/redline-theater.test.mjs`, whose contract test parses the real `tools/mcp-server/ToolCatalog.cs` so the demo's claim to speak the shipped tool contract is checked rather than asserted. |

![REDLINE THEATER](../images/demo/redline-theater.gif)

### Diff stress: the other pipeline, and what it costs

The page has a second mode, and it exists because "redline" means two different
things in this codebase and only one of them was on screen. The negotiation
above **records** its redline — each MCP call writes `w:ins`/`w:del` into the
live package as it lands. **Diff stress** does the other thing: it recomputes
the redline from scratch after every single edit, baseline against current, and
times it, while each frame appends a clause so the input grows under the engine
instead of sitting at one fixed size that would measure nothing.

![DIFF STRESS](../images/demo/diff-stress.gif)

The meter in that GIF reads slower than the table below — 132 ms where the table
says 74 — and should: capturing it runs a screenshot recorder at 10fps against the
same CPU, and the clip starts after a full negotiation has already grown the
document. Both are the confound described further down; the stress p50 measures the
machine and the input as much as the engine. The GIF is there to show the loop
running, not to be read off.

Three depths, because "how fast is the diff engine" has three honest answers:

| depth | engine calls | measured |
|---|---|---|
| `revisions` | `docxDiffGetRevisions` | ~19 diffs/s, p50 **52 ms** |
| `redline` | `docxDiffCompareProducts` → redline package + revisions | ~13 diffs/s, p50 **74 ms** |
| `full + HTML` | the above, then `convertDocxToHtml` with tracked changes | ~7 diffs/s, p50 **141 ms** |

Against a mutation costing ~2 ms through the same MCP endpoint, that is **33× to
78×**, and the panel reports the ratio it just measured rather than the one
written here. So the honest answer to "can we animate a diff per edit at 60fps"
is still no — but it is a much closer no than it was. A redline-per-edit loop now
lives between 7 and 19 frames per second on a document this size, where a month of
engine work ago it was 2 to 6. The shallow depth is inside the range a viewer
reads as motion rather than as a series of updates. That gap is the entire argument for
recording a redline when you are the one making the edits, and for computing one
when somebody hands you a document they changed. Both emit the same native
markup; the meter shows what each costs to get there.

The `redline` depth uses `docxDiffCompareProducts` rather than `docxDiffCompare`
followed by `docxDiffGetRevisions`, because one memoized alignment pass yielding
both products measured **51 ms against 85 ms** for the two calls separately —
about a third saved, which is what that API is for. That pair is measured the
controlled way described below, not read off the panel.

These figures track the engine, and have moved repeatedly without a line of demo
code changing — which is the argument for a panel that measures rather than a page
that quotes. Three of those movements have a cause; the most recent refresh does
not, and saying which is which is the whole point of keeping the list:

- **#616** roughly halved `docxDiffCompare` on this document (278 ms → 153 ms) by
  removing a read amplification.
- **#626** gated the compatibility normalizer on a streaming scan instead of a full
  parse. The two deeper rows came down 10–18% (`redline` 206 → 181, `full`
  422 → 352) while `revisions` held inside the noise band — the saving growing with
  pipeline depth is what a load-path change looks like, since the deeper the
  pipeline the more packages it opens.
- **#627** stopped giving the diff engine's reads an identity nothing asks for, and
  this time all three rows moved together, 13–18% (152 → 124, 181 → 157,
  352 → 307). A saving that is uniform across depths is the signature of a change
  on the path every depth shares. (#629 landed alongside it, but its snapshot reuse
  is across comparisons and this loop makes one comparison per frame, so it has
  nothing to reuse — the panel would not see it.)
- **A refresh with no attribution, which is sometimes the honest answer.** The table
  was re-read after #620 merged its `WmlToHtmlConverter` / `MarkupSimplifier`
  speedup, and every row came out 5–15% *slower*. Nothing in that merge could do it:
  `revisions` never calls the converter at all. Decomposing said the opposite — the
  conversion stage on its own (`full` − `redline`, controlled) went 154 ms → 143 ms,
  cheaper, exactly as #620 intends, while the compare stage read dearer with no cause
  in the diff. The reading that changed was the machine, on a different day. Kept as
  a worked example of a movement that fits nothing and should be labelled as such.
- **#653 compiles the browser build's hot paths ahead of time from a recorded
  profile, and it is the largest movement this page has recorded: everything roughly
  halved.** `revisions` 143 → 52 ms, `redline` 170 → 74, `full + HTML` 324 → 141;
  `docxDiffCompareProducts` on fixed inputs 125 → 51 ms, and the conversion stage in
  isolation 144 → 60. Both methods agreed and the controlled figures were pooled over
  two sessions, because a claimed 2× deserves more than the one session that made an
  earlier baseline look better than it was. Uniform across depths again, but this
  time the shared thing is not a code path — it is the .NET execution underneath all
  of them, which is what ahead-of-time compilation buys and why even the ratio
  against the recording path moved (68–157× down to 33–78×: recording got faster
  too).

Those attributions are what the shape of the movement suggests, not what the panel
proves; it measures the total, and the commit messages are the authority on cause.
Where the shape fits nothing in the diff, as above, the honest label is "no cause
found" rather than the nearest PR.

Read them as one machine's order of magnitude, not as a benchmark. Two things
make the last digit meaningless. Repeated medians on the same container spread
5–20% run to run. And the p50 the meter reports is taken over a run during which
the document is *growing*, so it depends on how far the loop got — which makes
stress p50 unsuitable for comparing one engine build against another, however
tempting the number looks sitting there. (Checked after #623: a controlled
fixed-input measurement showed no regression, while the stress readout appeared
to move by a third. The controlled one was right.) The rows above are therefore
each a median of three full 40-frame runs on one container, refreshed together so
they describe one machine on one build rather than an accumulation of readings. To compare builds, use a fixed
input and medians of many — which is what the harness below does properly.

**The ratio column survives what the millisecond columns do not, and that is the most
useful thing measured here.** Two pooled measurements of the same build family — three
40-frame stress runs plus two controlled sessions each — came out 21–28% apart on every
absolute figure (`revisions` 52 → 41 ms, `redline` 74 → 55, `full + HTML` 141 → 101) while
the ratios against the recording path barely moved: 33 → 34×, 45 → 45×, 78 → 80×. Nothing
in the engine changed between them to explain a quarter; what changed is how fast the
container was that hour. The mutation path scaled by the same factor as the diff path, so
the ratio held. That gives a sharper test than "is the movement uniform across depths":
**when the absolutes move and the ratio holds, it is the machine; when the ratio itself
moves, the two paths changed by different amounts and something real happened** — which is
exactly how #653 announced itself, dragging 68–157× down to 33–78×. The table keeps its
original pooled figures rather than being refreshed to the newer ones, because chasing the
faster reading would just be chasing weather, and the panel recomputes the ratio live
anyway.

That test needs a threshold, or it will fire on noise. Two consecutive `full + HTML` runs on
one build within one hour read **77× and 84×** — so the ratio itself carries roughly ±10%,
and a movement inside that band means nothing. #653 is the calibration for the other end: it
halved the ratio, and there was no ambiguity about reading it as real. The rule is worth
stating with the number attached — **a ratio movement under about 10% is noise; one that
approaches a halving or doubling is the engine.** In between, measure again before believing
it.

The rigorous headless counterpart — the same question asked of a 147 KB certificate of
incorporation, with stage attribution and allocation figures — lives in
`benchmarks/docxdiff-stress/FINDINGS.md`. That harness is the authority on engine
performance; this mode is the one you can watch.

The meter resets on every run: carrying frames across a depth change would report
a p50 describing neither pipeline. Switching back to the wire stops the loop, so
it cannot keep burning the engine behind a hidden pane.

### Why the frames are real, and where the resemblance stops

`tools/mcp-server` is a .NET process on stdio; a browser cannot subprocess it. So
`mcp-wire.js` reimplements the server's **front half** — envelope parsing, the
(tool, action) routing table, the MCP result shape (`content[].text` +
`isError`), and the convention that a business-level failure is a tool *result*
rather than a JSON-RPC error — over `DocxSession` via WASM instead of over a
pipe. The frames are the frames; the transport under them differs.

That claim is checked, not asserted:
`docs/demo/tools/redline-theater.test.mjs` parses the shipped `ToolCatalog.cs`
and fails if any (tool, action) pair the endpoint routes, any search `mode`, any
`set_mode` value, or any argument the script puts on the wire is absent from the
real catalog. It caught two mismatches while the demo was being written (a batch
argument named `atomic` that the catalog calls `mode`, and the `set_mode` wire
names needing to map to the `TrackedChangeMode` enum rather than being passed
through as strings).

What the endpoint deliberately does **not** reimplement: the document store and
its containment checks (a tab has no filesystem, so sessions are opened by the
host page rather than by `docxodus_open` with a path), the transaction/retry
journal, the delivery-bundle host, and the MCP Apps UI extension.

One thing the demo surfaces about the tool surface itself: attribution is a
session setting (`revisionAuthor`, taken at `docxodus_open`), so three counsel in
one session is a host-side switch with no tool behind it. The wire shows the two
`docxodus_track_changes/set_mode` calls that bracket the baseline build, but the
author changes between acts happen off-wire.

### Playable Doom walkthrough, in the real editor

This is a capture from the shipped editor, not a mockup. It starts on the document-hosted opener,
drops the coin, boots the real engine, enters Freedoom Episode 1, shows controlled movement and fire,
then pauses on an editable frame. The lower caption strip labels each step outside the native-size
document crop, so it covers neither the controls nor the game image.

<p align="center">
  <img src="../images/arcade-doom.gif" alt="Captioned walkthrough from the live document opener through playable full-color Doom and a paused editable frame" width="656">
</p>

The playable frame is Doom's complete, lossless **320 × 200 color framebuffer**: view, weapon,
ammo, health, armor and lives/status face. It is enlarged to the document column so the 11-pixel
HUD numerals display at least 24 CSS pixels high. Four fixed 18pt control paragraphs remain
legible above it. Pause and **Ctrl+C / Ctrl+V** duplicates the frame as another native document
image; **Undo / Redo** becomes a tiny backward/forward frame scrubber. The browser regression
measures completed image mutations and decoded, animation-frame-presented sources over five
seconds; both must remain at or above the ten-FPS design point. **Esc** freezes the native inline
image as ordinary copyable, undoable, saveable document content.

### The arcade's controls (`arcade-dock.js`)

Two pages host the arcade — the cabinet and, on a phone, the landing page — so
its controls live in one module rather than being hand-written twice, for the
same reason the editor surface itself ships from `npm/src/ribbon.ts`. Density is
measured from the controls' own host, not from a viewport media query: the
landing page frames the arcade inside a card, and a narrow card in a wide page
is narrow.

| | |
|---|---|
| **wide** | One bar under the document: cartridges, transport, pacing, embed, telemetry, hint. Unchanged from the cabinet's original dock. |
| **compact** | A slim HUD strip keeps the two controls you touch mid-game (play/pause and pacing); cartridges, restart, embed, telemetry and the hint move behind a `⋯` sheet. A thumb D-pad and a round **FIRE** button float over the bottom corners of the game. Nothing is dropped, only re-placed. |

`FIRE` sends `Space` — jump in the platformer, the weapon in the raycasters and
in Doom, and the coin drop on the attract screen. The old touch row had no
Space at all, so the shooters could be walked but never fought on a phone.

Doom's complete keyboard map also appears immediately above the framebuffer as
four centered 18pt document paragraphs. Each line is deliberately short enough
for the fixed document column; one oversized run can clip instead of reflowing
in a fixed-layout viewer. At the fitted desktop page the controls compute to
24px, and the browser guard still requires at least 14px after a 60% embed
scale. The context fence keeps all four static paragraphs out of every frame
conversion.

The pad is deliberately not a descendant of the editor root: the driver pauses
the game on any `pointerdown` inside the document ("the frame you clicked is now
an ordinary paragraph"), so a control living in there would pause on every tap.
Compact layout moves the nodes themselves rather than duplicating them into two
hidden layouts, so `startArcade` wires its listeners once and never learns that
layout exists. Specs: the phone-shaped controls in
`npm/tests/demo-arcade-mobile.spec.ts`, and the landing page's mobile default,
nav strip and floating controls in `npm/tests/social-demo.spec.ts`.

### The canvas font, and why the ASCII pages ship one

`fonts/docxodus-canvas-mono.woff2` (17 KB) is what keeps the Observatory's
phenomena and the Arcade's game screen on their grid.

The Observatory, opener, and first three cartridges draw into one Word paragraph
as a 92 × 26 character grid authored for Courier New. Doom uses a native image
and does not depend on glyph metrics. A text grid holds only while
every cell advances the same width
— and that is a property of the font the DEVICE resolves, which the document
has no way to state. The art draws with box drawing (`─ │ ┌`), block elements
(`█ ░ ▓`) and geometric shapes (`▶ ◀ ▲ ▼`); Android has no Courier New, and the
monospace face Chrome substitutes covers none of those, so each one lands in a
proportional fallback whose advance is not the cell. Every cell after it is
displaced — by a different amount on each row, because rows hold different
numbers of them. That is the tilt reported from the field: the title card's `X`
reading as a `K`, the logo smearing off the right edge, worst exactly where the
art is densest.

| Android's font coverage reproduced, canvas unpinned | the same frame, pinned |
|---|---|
| ![tilted attract screen](../images/demo/arcade-canvas-tilt-before.png) | ![aligned attract screen](../images/demo/arcade-canvas-tilt-after.png) |
| ![tilted dungeon](../images/demo/arcade-canvas-tilt-dungeon-before.png) | ![aligned dungeon](../images/demo/arcade-canvas-tilt-dungeon-after.png) |

Measured on a Pixel 5 rig with that font situation reproduced, as the spread
between the widest and narrowest row of the 92-cell grid: **12.1 cells** on the
attract screen and up to **23.0** in the text cartridges, against **0.07 cells**
worst case with the pin — across both viewports, the text cartridges and all
four phenomena. The mobile Doom check instead pins its exact 320 × 200 image geometry.

So `createCanvasPin()` in `ascii-scenes.js` pins the canvas paragraph to a font
we ship, rather than hoping the platform's fallback happens to match: a subset
of DejaVu Sans Mono whose every glyph advances 1233/2048 em, identical on every
device. It pins `white-space: pre` in the same rule (a row can never wrap —
see the CHANGELOG entry for that earlier fix) and neutralizes kerning,
ligatures and inherited letter/word spacing, which are per-glyph adjustments a
grid cannot survive either. The saved `.docx` is untouched and still says
Courier New: this is a display pin, and Word has the real font.

Rebuild it with `tools/build-canvas-font.sh`, which pins the source font by
SHA-256 and refuses any other — a font that is not single-advance would not
provide the guarantee. It writes the `.woff2` and the `.json` manifest that
`tools/canvas-font.test.mjs` reads: that test drives every scene, the whole
attract-screen sweep and every cartridge, and fails if they can draw a
character the subset does not cover. `npm/tests/demo-arcade-mobile.spec.ts`
proves the rest on a phone-shaped rig with Android's font coverage reproduced.

Docxodus stays MIT-licensed. Only the font file carries the Bitstream Vera
Fonts license (`fonts/LICENSE.txt`), which permits the subset provided it is
renamed away from "Bitstream"/"Vera" — hence "Docxodus Canvas Mono". Like the
rest of this directory it is documentation-site content and is not part of the
`docxodus` npm package.

### Doom, and what it does to the licensing

Cartridge 4 is the actual game. `doom-cart.js` drives
[doomgeneric](https://github.com/grubbyplaya/doomgenericjs) — id Software's
Doom source, GPL-2.0, compiled to JavaScript — on Freedoom's BSD-licensed
IWAD, and writes its complete 320×200 framebuffer into the media part of one
inline image in the screen paragraph every frame.

**Neither the engine nor the IWAD is in this repository.** They are 13 MB of
binary that would show up in every clone forever and never diff usefully, so
both are pinned jsDelivr URLs instead — by 40-hex commit for the engine, by tag
for the IWAD, both served with `access-control-allow-origin: *` and cached
immutably. Upstream commits, SHA-256s, license texts and verification commands
are in [`vendor/NOTICE.md`](vendor/NOTICE.md), along with what the IWAD's
sibling repository has to contain and why it needs to exist at all (GitHub
serves release assets without CORS, so a browser cannot fetch Freedoom's
release directly).

The Playwright specs do not touch the CDN for game data:
`npm/scripts/fetch-doom-iwad.mjs` pulls the release asset server-side, verifies
its digest, and drops a gzipped copy into the test webroot for the specs to
load same-origin.

**This is the one place the repository is not MIT.** `doom-cart.js` is combined
with a GPL engine at runtime, so that single file is offered under
GPL-2.0-or-later and says so in its header. Everything else — `ascii-arcade.js`
included, which imports it — stays MIT, which is GPL-compatible: each file
remains separately available under its own terms, and only the combination a
browser assembles when the Doom cartridge is selected is GPL-2.0.

The engine is behind a dynamic `import()` inside `doom-cart.js`, so a visitor
who plays the platformer or the dungeon fetches neither. `?wad=` points the
cartridge at a **same-origin** IWAD you host yourself and are licensed to play
(a retail `doom.wad` works), and `?sound=0` boots it mute.

### One native image, one playable contract

The screen is not a canvas overlay and it is not a character approximation. Each engine frame is
encoded losslessly as PNG and passed to `DocxSession.replaceImage`; that replaces the package media
payload referenced by the existing `w:drawing`. `DocxEditor.refresh()` then performs the same
single-block OOXML→HTML conversion and DOM reconciliation used after any other session edit. The
editor contains one `<img>` and no game canvas.

The acceptance test is deliberately about the displayed document, not optimistic source data:
the image must decode at exactly 320 × 200, sample pixel-for-pixel against doomgeneric's BGRA
framebuffer, retain a chromatic palette, fill at least 590 × 440 CSS pixels in the fitted page,
and make the HUD and its numerals at least 70 and 24 CSS pixels high respectively. Those guards
make unreadable grayscale, quantized and tiny-screen regressions fail even if they are fast.

### Where the ten visible FPS comes from

The expensive unit is a completed image mutation plus OOXML→HTML refresh, image decode, and
browser presentation—not a Doom tic. The driver therefore does not schedule the next replacement
until the current data-URI image is decoded and at least one `requestAnimationFrame` boundary has
passed. This prevents a fast producer from claiming frames the player never saw.

On the live browser path, the integrated five-second turning sample completed **78 document
frames in 5.02 seconds (15.53 FPS)** and observed **79 distinct decoded image sources on animation
frames (15.73 visible FPS)**, with approximately 17.7ms mutation and 28.8ms refresh time. The
browser spec repeats the measurement while holding a real turn key and requires both rates to
remain at or above 9.5 FPS. It separately proves that the view moves, the HUD stays fixed, and
keyboard input reaches the real engine.

Incremental image rendering needed one general library repair: the isolated render shell now
copies each embedded image relationship and supplies the normal base64 image handler. A focused
image-session regression calls `renderBlock()` after insertion and proves it returns a valid data
PNG and alt text. The full .NET and browser suites remain the guard against collateral library
regressions.

### The screen is fenced away from the caption

A single-block re-render does not render the block alone. The engine pads each
target with **one real neighbour on each side** before converting it, so that
`w:contextualSpacing` resolves exactly as it would in a full render; those
context clones are discarded once the target's HTML is extracted. That is
correct, and cheap — unless a neighbour is large.

The screen's lower neighbour was the caption: a formatted prose paragraph of 36
runs carrying a footnote reference, converted in full on every frame of every
game purely as context, then thrown away. Fencing the screen with two
one-character paragraphs moves it out of that slot, measured **7.4 → 8.8
repaints a second** and A/B'd inside one browser process so the container's
drift could not be mistaken for the change.

The fence has to be the boundary of the litter sweep too. `syncFromDocument`
deletes stray paragraphs between the screen and the caption, so that pausing to
edit cannot slowly fill the document with Enter-splits — and on the first
attempt it dutifully deleted the fence on the first pause. It now sweeps up to
the fence rather than past it.

The native image keeps that contextual saving while giving the framebuffer the
whole document column. Controls stay outside the hot block as four real 18pt
paragraphs, so neither their conversion cost nor tiny framebuffer text can
compromise play. The one full-color path is the playable path and saves as a
`.docx` without a mode switch.

There is deliberately no engine override. `import()` executes whatever it
fetches, on this page's origin and with its privileges, so a URL parameter
naming the module would be remote code execution from a crafted link rather
than a knob. The cartridge's URL gate allows exactly its own pin plus
same-origin, and nothing else.

The demo stays a documentation-site concern. Neither its runtime nor the Doom
cartridge's third-party dependencies are included in the `docxodus` npm
tarball: the package
publishes an explicit allowlist of library JavaScript, declarations, and WASM
runtime files, and `npm run test:package-boundary` audits the actual packed
manifest after the Playwright webroot has been populated.

The surface itself is **not** written here — it ships in the npm package
(`npm/src/ribbon.ts`). These pages used to carry a hand-written toolbar each,
which drifted until the demo advertised a smaller editor than the one that
shipped. Changing the editor's UI means changing the module, not these files.

The sample document is the branded product guide in this directory; regenerate
it with `python tools/generate-demo-guide.py` from the repository root.

## Publish

The pages load the library from jsDelivr at an exact version, so the pin can
only move *after* npm publishes and the CDN serves it. Move every pin in one
change — the pages, this README, `docs/npm-package.md`, `npm/README.md`,
`npm/examples/embed.html`, and `RELEASE_ENGINE` in
`npm/tests/social-demo.spec.ts` — and `tools/engine-pin.test.mjs` (run by
`npm run test:demo-logic`, and so by every Playwright run) fails if one of them
is left behind, if the version drops below the arcade's
`IMAGE_ENGINE_MINIMUM`, or, under `DOCXODUS_CHECK_CDN=1`, if the CDN does not
serve it yet. That guard exists because a stale pin is invisible to the browser
specs: every one of them overrides `?engine=` to the locally built bundle.

1. Publish `docxodus@12.0.0` and confirm
   `https://cdn.jsdelivr.net/npm/docxodus@12.0.0/dist/embed.bundle.js` returns JavaScript.
2. In GitHub **Settings → Pages**, choose **GitHub Actions** as the publishing source.
   The `Deploy static demo to GitHub Pages` workflow uploads `/docs` without
   running Jekyll.
3. Open `https://jsv4.github.io/Docxodus/demo/` and verify the status chip becomes `Live`.
4. If the site is hosted anywhere else, update the absolute `og:url`, `og:image`,
   and `twitter:image` values in `index.html` and `app.html` before sharing.

All the pages accept query overrides for local or preview testing:

```text
?engine=<module URL of embed.bundle.js>&doc=<CORS-readable DOCX URL>
```

`app.html` also takes `?blank=1` to start from a new empty document.
`index.html` takes `?demo=arcade|editor` to pin which demo the frame mounts
(otherwise a phone gets the arcade and a desktop the editor), plus the arcade's
own `?cart=` and `?intro=0` when it is the one mounted.

## Social-platform expectations

- LinkedIn uses the landing page's Open Graph title, description, and image. It
  does not run the editor inside the feed; the user clicks through to the live page.
- X receives a `summary_large_image` link card. Its current developer docs no
  longer document historical Player Cards, so the page does not promise an
  interactive editor inside a post. `player.html` is for iframe-capable sites.
- The Playwright specs (`npm/tests/social-demo.spec.ts`) validate the static
  metadata, boot-on-tap behaviour, the live editor, and the responsive collapse —
  locally. They cannot prove how an external social client will render a post.

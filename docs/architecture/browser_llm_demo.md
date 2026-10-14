# Browser-LLM editing demo — design stub

> **Status: stub / not scheduled.** Captured while building the CDN embed packaging
> (`docs/demo/`, `npm/src/embed.ts`) so the idea isn't lost. Nothing here is
> implemented; no engine changes are expected to be needed.

## The idea

A demo page where an LLM redlines a contract **entirely in the browser**: the
document loads into a `DocxSession`, the model proposes edits against the
markdown projection's stable anchors, and the edits land as **native tracked
changes** the user can accept/reject per revision — then `save()` produces a
.docx whose redlines open in Word. No server, no upload.

This closes a loop the library was explicitly designed for: the markdown
projection (`WmlToMarkdownConverter`) is the LLM's *read* surface, `DocxSession`
is the *write* surface, and the `{#kind:scope:unid}` anchors are the shared
addressing system between them.

## The loop (all existing API — no new engine work)

1. `openDocxSession(bytes)` → `session.project()` → markdown where every block
   carries a stable anchor (`{#p:body:a1b2c3d4}`)
2. Prompt the LLM: *"suggest edits as (anchor, replacement-markdown) pairs"*
   (structured output; the anchor ids are the contract)
3. `session.setTrackedChanges("render_inline")` + `setRevisionAuthor("AI Reviewer")`,
   then `replaceText(anchor, newText)` per suggested edit
4. Re-render (or run inside `DocxEditor`, which repaints incrementally) — the
   AI's edits appear as `w:ins`/`w:del` markup
5. Review UI via `listRevisions()` / `acceptRevision(id)` / `rejectRevision(id)`
6. `save()` → lossless .docx with surviving revisions intact

## LLM transport options (feature-detect, offer both)

| Route | Pros | Cons |
|---|---|---|
| Chrome built-in AI (Prompt API, `LanguageModel.create()`, Gemini Nano) | Zero server, zero key, fully on-device — matches the "document never leaves the page" story | Chrome-only; availability gated (verify current rollout status when building); multi-GB model download + hardware requirements; small-model quality |
| Claude API direct from browser | Best quality by far; CORS-enabled endpoint; TS SDK supports it behind `dangerouslyAllowBrowser: true` | Never embed a key in a public page — BYO-key field (localStorage) or a tiny rate-limited proxy (e.g. Cloudflare Worker) |
| WebLLM / transformers.js (WebGPU) | Open models, cross-browser where WebGPU exists, on-device | Large weight downloads; similar quality ceiling to Nano |

Recommended demo shape: feature-detect `LanguageModel` → on-device path;
otherwise show a paste-your-key field. Same edit-loop code either way.

## Open questions (for when this is picked up)

- Edit contract shape: plain `(anchor, markdown)` pairs vs. richer ops
  (`delete`, `insertAfter`, formatting) — start with replace-only
- Prompt design: whole-document projection vs. per-section chunks for long
  contracts (projection is token-cheap, but 1M-context models make whole-doc viable)
- Anchor drift: edits invalidate downstream anchors? (They don't — anchors are
  Unid-stable across edits; verify under tracked-changes mode specifically)
- Where it lives: `docs/demo/redline.html` beside the player pages, sharing the
  embed bundle
- Guardrails for the BYO-key path (spend caps, model pinned, key never sent
  anywhere but api.anthropic.com)

## Non-goals

- No engine/API changes — if the demo needs one, that's a finding to file, not
  scope to absorb here
- Not a product feature of the npm package; a `docs/demo/` page only

# Agent-Facing DOCX Editing Server (MCP)

**Status:** Implemented. Source: `tools/mcp-server/` (`Program.cs`, `Dispatcher.cs`, `SessionStore.cs`,
`ToolCatalog.cs`, `JsonRpc.cs`, plus the storage layer `IDocumentStore.cs`,
`LocalFileDocumentStore.cs`, `DocumentStores.cs`). Tests:
`Docxodus.Tests/McpServerDispatcherTests.cs` (`MCP###`). Packaged as the `dotnet tool`
`docxodus-mcp`.

## What this is

A [Model Context Protocol](https://modelcontextprotocol.io) (MCP) server that lets an AI agent
open a `.docx` file, read it, edit it through a small set of high-level tools, and save it back
— entirely locally over stdio, no network calls, no telemetry. It is the agent-facing counterpart
to `tools/python-host/` (a stdio host for a *library-shaped* API, one request per method call):
this server groups the same underlying `DocxSession` surface into a smaller number of
**intent-shaped tools**, because that is the granularity an LLM tool-calling loop wants — a model
picks from a short, memorable list of verbs (`docxodus_edit`, `docxodus_format`,
`docxodus_table`, …) with an `action` discriminator, rather than one MCP tool per one of
`DocxSession`'s ~40 public methods.

Every tool call ultimately routes through `Docxodus.Internal.DocxSessionOps` (and, for tracked
changes, the public `RevisionProcessor` plus `Docxodus.Internal.DocxDiffOps`) — the exact same
facade the WASM bridge and the Python stdio host use. No new editing logic lives in this server;
`Dispatcher.cs` is argument-shape translation only.

### Why this shape

Document-editing MCP servers built around "open a file into a stateful in-memory session,
address every subsequent edit by a stable anchor id, group many operations under a handful of
grouped-intent tools (read / preview / pagination / search / edit / format / create / list /
comment / annotate / track-changes / batch-mutate / table), save on request" are a known-good shape for this problem — it matches how
this class of tool is used in practice: an agent reads a projection once, holds anchor ids in its
context, and issues a sequence of small, anchor-addressed mutations before saving. This server
adopts that shape but is a clean-room implementation against Docxodus's own `DocxSession` engine
— no code or fixtures from any other project were copied into this repository. Where Docxodus
doesn't yet have a matching capability, the corresponding tool/action is either omitted or
explicitly documented as a gap below, rather than faked.

## Architecture

```
MCP client (stdin/stdout, newline-delimited JSON-RPC 2.0)
     │
     ▼
tools/mcp-server/Program.cs        — JSON-RPC transport: initialize, tools/list, tools/call
     │
     ▼
tools/mcp-server/Dispatcher.cs     — (tool, action) → Docxodus API call; arg parsing only
     │                                also: tools/mcp-server/SessionStore.cs (external session_id
     │                                → DocxSessionOps handle + opened-from location + settings)
     │
     ├──▶ IDocumentStore  ─────────  where bytes come from and go to (see Document storage)
     │      LocalFileDocumentStore    scope-rooted local filesystem — the only backend today
     ▼
Docxodus.Internal.DocxSessionOps / DocxDiffOps   (the same facade the WASM bridge and
     │                                             tools/python-host route through)
     ▼
Docxodus.DocxSession   (the real work — see docs/architecture/docx_mutation_api.md)
```

`SessionStore` is the one piece of state this server owns that `DocxSessionOps` doesn't: a
string `session_id` → `{ handle, location, settings }` map. It exists for two reasons:

1. `docxodus_save` needs to remember the location a session was opened from so "save" can mean
   "write back to the same document" without the caller repeating it.
2. `docxodus_track_changes`'s `accept_all`/`reject_all` need to swap a session's entire
   underlying document for a whole-document byte transform (`RevisionProcessor.AcceptRevisions`/
   `RejectRevisions`, which operate on bytes, not a live session) while the caller keeps
   addressing the same `session_id` — see `SessionStore.Rebind`.

Session ids are 16 random bytes, not a counter. The id **is** the capability — holding one is
what lets a caller act on that document — so a guessable id would let anything able to make tool
calls address a session it was never handed.

## Wire protocol

Newline-delimited JSON-RPC 2.0 on stdin/stdout — the MCP stdio transport. Diagnostics go to
stderr only. Alternatively, `docxodus-mcp --http PORT` serves the same protocol over a minimal
streamable-HTTP binding (each POST carries one message and gets one `application/json` response —
no SSE, no session-id handshake, no TLS), so the server can sit behind a tunnel (`ngrok http PORT`)
for remote-MCP or ChatGPT Apps development. Requests are serialized under a lock either way; the
session registry assumes single-threaded access.

- `initialize` → `{ protocolVersion, capabilities: { tools: {}, resources: {}, extensions: { "io.modelcontextprotocol/ui": ... } }, serverInfo }`
  — the requested `protocolVersion` is echoed when present (every implemented method is
  shape-stable across published revisions; the UI extension negotiates via `capabilities.extensions`)
- `notifications/initialized` → no response (notification)
- `tools/list` → `{ tools: [ { name, description, inputSchema, _meta? }, ... ] }` — the 16 tools below
- `tools/call` params `{ name, arguments }` → `{ content: [ { type: "text", text: <JSON string> } ], isError }`
  (plus `structuredContent`/`_meta` on the two preview-related tools — see "Inline preview" below)
- `resources/list` / `resources/read` / `resources/templates/list` — serve the `ui://` viewer
  template (see "Inline preview" below)

**`isError` semantics:** a tool result is marked `isError: true` when either (a) the JSON text is
a Docxodus `EditResult`-shaped object with `"success": false` (anchor not found, malformed
markdown, etc. — a normal business outcome the agent should see and can often recover from), or
(b) the dispatcher itself threw (bad arguments, unknown session id, unknown tool/action). Neither
case is a JSON-RPC protocol error — those are reserved for malformed JSON-RPC envelopes and
unknown methods, which a well-behaved client should never produce.

## Session lifecycle

```
docxodus_open(path) → session_id
   ↓ (any number of docxodus_edit / docxodus_format / docxodus_create / docxodus_table /
   ↓  docxodus_list / docxodus_comment / docxodus_annotate / docxodus_track_changes /
   ↓  docxodus_mutations calls)
docxodus_save(session_id, path?)
   ↓
docxodus_close(session_id)
```

Sessions are in-memory only; `docxodus_close` (or process exit) discards unsaved changes.
Anchor ids returned by `docxodus_open` → `docxodus_get_content`/`docxodus_search` remain valid
across mutations per the anchor-lifecycle contract in `docs/architecture/docx_mutation_api.md`
(the `created`/`removed`/`modified` lists on every `EditResult` are the ground truth — don't
assume an anchor you haven't seen invalidated is still resolvable forever, but do trust that
list).

### Anchor stability across save→reopen

Within one session anchors are stable, but a plain `docxodus_save` deliberately strips the
`PtOpenXml:Unid` bookkeeping they are derived from (clean OOXML out by default — persisting it
bloats a file by hundreds of KB; see `docs/architecture/docx_mutation_api.md`). After a
save→`docxodus_close`→`docxodus_open` round-trip, anchors for content that was inserted or
edited in the previous session therefore 404 with `anchor_not_found` (unchanged content gets
the same deterministic ids back; changed content does not).

When a workflow knows it will need a close+reopen — the concrete case is switching
`trackedChanges` mode mid-workflow, which is also an open-time-only setting — opt in:

- `docxodus_open(persistAnchorIds: true)` makes every save of that session keep the
  bookkeeping, so a session reopened over the saved file resolves the same anchor ids.
- `docxodus_save(persistAnchorIds: true|false)` overrides the open-time choice for one save:
  `true` makes a single anchor-stable checkpoint from a default session; `false` produces a
  clean deliverable from a session that was opened anchor-stable.

Both default to the stripping behavior; nothing changes for callers that pass neither.

## Document storage

Everything this server reads or writes goes through one small interface, `IDocumentStore`:
`Resolve` (turn a caller-supplied location into the canonical in-scope identifier, or reject it),
then byte-level `Read`/`Write` of an already-resolved location. `DocumentStores` is the single
owner of "which backend, rooted where," built once at startup from the environment — the same
single-owner-facade shape the core library uses for `DocxSessionOps`/`DocxDiffOps`.

Today there is exactly one backend, `LocalFileDocumentStore`. The seam exists so that adding
another — object storage, a content repository — is a new `IDocumentStore` implementation plus a
case in `DocumentStores.Create`, with no change to the dispatcher, the tool schemas, or any
session logic. The interface is deliberately three methods over opaque location *strings* rather
than a filesystem-shaped API, so a backend without directories doesn't have to pretend to have
them.

### Isolation is a property, not a check

A store is constructed **already rooted at its scope**, and every location a tool call names
passes through `Resolve` before anything touches it. There is no read or write path that skips
resolution, so a session physically cannot name a document outside its scope — the guarantee
doesn't depend on remembering to write an `if` in the dispatcher.

The rules `LocalFileDocumentStore` enforces:

- A **relative** location resolves under the scope root.
- An **absolute** location is accepted only if it canonicalizes to something inside that root.
  This is what keeps ordinary local use working — with the default root, an agent can open
  `~/Downloads/contract.docx` by its natural path — while a narrower root confines the exact same
  unmodified tool surface completely.
- Containment is checked against **symlink-resolved** paths, both root and candidate, resolving
  every component from the volume root down rather than just the leaf. The case that motivates
  it: `{root}/link → /elsewhere`, then naming `{root}/link/secret.docx`, whose own leaf is not a
  link at all — leaf-only resolution (and deepest-existing-ancestor resolution) both miss it.
  Dangling links are detected by reading the link rather than following it, so a link to a
  not-yet-existing directory can't be used as a write escape.
- Segment boundaries are respected, so `/srv/base-2` is not inside `/srv/base`. The check uses
  `Path.GetRelativePath` rather than a string prefix test, which also gets the platform's case
  rules right.

### Configuration

Backend and root are **operator configuration, never tool arguments** — if the agent could name
a backend or a root per call, the scope would be chosen by the very thing it exists to contain.
The only storage input a tool call carries is a location *within* the configured store.

| Variable | Meaning |
|---|---|
| `DOCXODUS_STORAGE_BACKEND` | Backend id. Only `file` is implemented; defaults to `file`. |
| `DOCXODUS_STORAGE_ROOT` | Directory every location must resolve inside. Defaults to the user's home directory — permissive enough for real local work, still a genuine boundary. Set it narrower to confine the server; set it to the filesystem root to opt out of confinement entirely. |
| `DOCXODUS_STORAGE_SCOPE` | Optional single path segment appended to the root, making the effective root `{root}/{scope}`. Two processes launched with different scopes cannot see each other's documents. |

Misconfigured storage is fatal at startup rather than surfaced per call: a server that answered
`tools/list` and then failed every open would be harder to diagnose than one that refuses to
start with the reason on stderr.

### How scope reaches the server

`DOCXODUS_STORAGE_SCOPE` is supplied by **whoever launches the server**, not by the agent, which
never learns it and has no argument through which to name one. That is deliberate: stdio MCP has
no in-band authentication — there is no channel on which a caller proves who it is, because the
process was spawned by the client and whatever env and credentials it was spawned with *is* the
authorization. Server-side ACL logic would only ever be checking a self-asserted identity, so the
scope belongs at the trust boundary that actually exists: process spawn.

Two consequences worth stating plainly:

- **Persistence is free.** The scope is a stable path segment, so passing the same value next
  session reaches the same documents. There is no scope registry, no token format, and no
  revocation list — persistence is the filesystem's.
- **One process spans exactly one scope.** A caller needing documents from two scopes runs two
  server processes. That is a real constraint, and the right one: it keeps the isolation boundary
  identical to the process boundary rather than introducing an in-process notion of "current
  tenant" that a confused agent could be talked across.

This design assumes the launching client is trusted to pick its own scope. Handing a scope to
something *not* trusted to choose one would need a verifiable capability token (signed, so the
server can confirm it issued it) — deliberately not built, since it also brings a revocation
problem that has no good answer at this layer.

## Tool reference

Three lifecycle tools, thirteen grouped-intent tools. Every grouped tool takes `sessionId` plus an
`action` string; see `tools/mcp-server/ToolCatalog.cs` for the exact JSON Schema advertised over
`tools/list` (this section is the narrative version).

### Lifecycle

| Tool | Arguments | Result |
|---|---|---|
| `docxodus_open` | `path` (a location within the configured scope — see Document storage), `trackedChanges?` (`accept`\|`render_inline`\|`strip_deletions`), `revisionAuthor?`, `undoDepth?`, `persistAnchorIds?` (default false — see Anchor stability below) | `{ sessionId, path }` — `path` is the **resolved** location |
| `docxodus_save` | `sessionId`, `path?` (resolved the same way; defaults to the location the session was opened from), `persistAnchorIds?` (per-call override of the session's open-time setting; absent = use it) | `{ path, bytesWritten }` |
| `docxodus_close` | `sessionId` | `{ closed: true }` |

### `docxodus_get_content` — read

`format`: `markdown` (anchor-addressed projection — `DocxSessionOps.Project`/`ProjectAnchor`),
`html` (`DocxSessionOps.RenderHtml`/`RenderBlockHtml`), `text` (markdown with a best-effort
regex-based syntax strip — an approximation, not a real markdown parser; use `markdown` for
anything that needs to survive a write-back), `blocks` (every addressable block's
`BlockMetadata` — style id/name, outline level, list facts), `info` (`GetEditSummary` plus the
`SectionInfo` of the first body block found). Optional `anchorId` scopes
`markdown`/`text`/`html` to one block's subtree via
`ProjectionDepth.SubtreeAndFollowingSiblings`. The full markdown/text/blocks
reads include every projected package story, including `hdr*`/`ftr*`; an `anchorId` returned by
`set_header_text` or `set_footer_text` can also be handed straight back to markdown, text, or HTML
read-back. (The unscoped continuous HTML render is body-oriented; use the story anchor for
header/footer HTML.)

### `docxodus_preview` — render for the inline widget

`{ sessionId, anchorId? }` → the same converter profile as `docxodus_get_content format:"html"`
(`DocxSessionOps.RenderHtml`, or `RenderBlockHtml` when `anchorId` is given), but shaped for a UI
host instead of the model: the markup rides in the result's `_meta["docxodus/html"]`, which MCP
Apps hosts deliver to the widget (`ui/notifications/tool-result`) and ChatGPT exposes as
`window.openai.toolResponseMetadata` — while `content`/`structuredContent` carry only
`{ sessionId, anchorId?, htmlLength }`. A multi-hundred-KB render therefore costs the model's
context nothing. Call it again after edits to refresh the view; the widget's Refresh button does
exactly that via widget-initiated `tools/call`. See "Inline preview" below.

This preview is continuous until #434 provides a server-side paginated HTML substrate. With a
valid citation token it displays the exact cited page label and highlights the source anchor, but
returns `pageNavigation: "unavailable_continuous_preview"` rather than pretending a physical page
box exists.

### `docxodus_pagination` — register or consume an exact PageMap

`action: register` validates a renderer-materialized `PageMap`; `status` and `cite` consume it
with an exact `{ documentVersion, rendererFingerprint }` token. Mutations stale the map
automatically. Continuous/no-map, stale, fingerprint-mismatched, and unmapped-anchor results are
explicitly unavailable. The server does not estimate pages or bundle a browser; see
[`page_map.md`](page_map.md).

### `docxodus_search` — find text or blocks, get reusable anchor ids back

`mode: text` and `regex` use `DocxSessionOps.Grep` (returns span + context, not just an anchor —
useful for a follow-up `apply_format`/`replace_text_at_span`-style edit, though this server only
exposes the anchor-level ops from `docxodus_edit`/`docxodus_format` today). `mode: kind`,
`annotation`, `bookmark` use `FindByKind`/`FindByAnnotation`/`FindByBookmark` (anchor-only
results). Text/regex searches remain body-only by default for backward compatibility; optional
`scope: body|headers|footers|header_footer|all` widens them to every part in those story
categories (`headers` covers every `hdr*`, not merely `hdr1`). Text/regex results carry the
reusable id at `enclosingAnchor.id`; anchor-only results carry it at `id`. Either is the same
anchor every other tool's `anchorId`/`cellAnchorId` argument expects — this server doesn't invent
a separate "search result handle" concept; Docxodus's anchors already are one.
All search modes accept the exact citation token and attach a citation envelope to each result.

### `docxodus_edit` — text/block CRUD + undo/redo

`insert_paragraph`, `replace_text`, `replace_text_range`, `delete_block`, `delete_range`,
`delete_section`, `split_paragraph`, `merge_paragraphs`, `undo`, `redo` → the identically-named
(camelCase) `DocxSession` methods via `DocxSessionOps`. `markdown` payloads use the supported
subset documented in `docx_mutation_api.md` (ATX headings, bulleted/ordered lists, bold/italic/
code/strike, links, hard breaks).

### `docxodus_format` — formatting

`apply_format`/`apply_format_by_substring` (`FormatOp`: bold/italic/underline/strike/code/color/
vertAlign/fontSizePts/fontFamily), `set_paragraph_style`, `set_paragraph_format`
(`ParagraphFormatOp`: alignment/indentDelta/firstLineIndent/hangingIndent/spacingBefore/
spacingAfter/lineSpacing/lineSpacingRule/pageBreakBefore/topBorder/bottomBorder/clearBorders —
indent/spacing values are twips (1440 = 1in, 20 = 1pt); `firstLineIndent`/`hangingIndent` are
one either/or `w:ind` slot (setting one evicts the other, both in one call is rejected);
`lineSpacing` is measured per `lineSpacingRule` (`auto` default = 240ths of a line, so 240 =
single/480 = double; `exact`/`atLeast` = twips); `topBorder`/`bottomBorder` add or replace a
`w:pBdr` edge, `clearBorders` removes the whole `w:pBdr` before either is applied in the same
call), `set_list_level`, `remove_list_membership`, `apply_list_format`.

### `docxodus_create` — new structural content

`insert_paragraph`, `insert_heading` (sugar: builds a `"#".."######" + " " + text` markdown
payload and calls `InsertParagraph` — Docxodus has no separate heading-insert primitive, so this
server composes one from the paragraph op and the markdown subset's ATX-heading support),
`insert_table`, `insert_horizontal_rule`, `insert_footnote`, `insert_endnote`,
`insert_page_number_field`, plus the header/footer story actions from the existing session API:

- `set_header_text` / `set_footer_text`: `bodyAnchorId` selects the governing section,
  `kind: default|first|even` selects its running story, and `markdown` replaces that story's
  content. The result's `created` list contains the new `p:hdr*`/`p:ftr*` paragraph anchor.
- `ensure_header_footer_visible`: `bodyAnchorId` + `kind` enables Word's first-page or even-page
  visibility flag for a story already referenced by that section (default is a successful no-op).

The existing `insert_page_number_field` composes with the returned story paragraph: pass its id
as `anchorId` to append `PAGE`/`NUMPAGES` inside the header/footer, exactly as it already does for
any other paragraph. No MCP-only editing logic is involved; all three routes are thin calls into
`DocxSessionOps`.

### `docxodus_list` — list membership

`apply_format` (promotes/demotes a paragraph to a real, auto-numbered `w:numPr` list via
`ApplyListFormat` — this is the one that actually creates Word-native numbering, unlike a bare
markdown `"- item"` payload, see Known gaps), `apply_format_range` (the same conversion across a
contiguous sibling run via `ApplyListFormatRange(firstAnchorId, lastAnchorId, format)` — one call
instead of one per item, and the members are *guaranteed* to share one `w:num` instance so the
sequence stays intact), `set_level`, `remove`, `get_membership`. `listFormat` accepts the full
`ListFormat` vocabulary: `bullet`, `decimal`, `lowerLetter`, `upperLetter`, `lowerRoman`,
`upperRoman`, plus the `*Parenthesis` variants of the numbered formats (`decimalParenthesis` →
`(1)`, `lowerRomanParenthesis` → `(i)` — the legal-drafting presets) and `none` (issue #313).

### `docxodus_comment` — native Word review comments (issues #300, #317, and #341)

`add`/`reply`/`resolve`/`update`/`remove`/`list` over `DocxSession`'s comment API — real
`w:comment` markup with `w:commentRangeStart`/`End` + `w:commentReference` body plumbing,
visible in Word/Google Docs/LibreOffice's Reviewing pane. `add` requires exactly one target:
a body paragraph (`anchorId` + optional `span`) or a tracked change from
`docxodus_track_changes list` (`revisionId`; `span` is not accepted). Revision targeting brackets
the change's live extent, preserving the comment as an anchored range or collapsed point after
accept/reject. Both forms require `author`; `initials`/`date` are optional and `w:date` is written
only when provided, keeping output deterministic. `reply` takes the parent
definition's `commentAnchorId`, gives the reply its own definition/id plus an adjacent reference,
and links it with Word's `w15:paraIdParent` metadata; only the thread root owns range markers,
so nested replies inherit that range through reference-only parents. `resolve` addresses
one comment by `commentAnchorId`; `resolved` defaults true and false reopens it without losing
parentage. Flat comments are upgraded with find-or-created `commentsExtended.xml` and
`commentsIds.xml` parts when first replied to or resolved.

`list` returns part-order entries with additive `parentAnchorId` and `resolved` fields when a
Word extension entry exists; an absent field means legacy/flat metadata rather than reopened.
`update`/`remove` use the same definition anchor. Removing a comment also prunes the extension
entries it owned and clears child links that would otherwise dangle. Documented at
`docs/architecture/docx_mutation_api.md` (Comments section).

### `docxodus_annotate` — annotation overlay

`add`/`update`/`remove`/`move`/`list`/`find` over `DocxSession`'s Tier E annotation API
(`AddAnnotation`/`UpdateAnnotation`/`RemoveAnnotation`/`MoveAnnotation`/`ListAnnotations`/
`FindByAnnotation`) — a highlight + label overlay stored in a bookmark and a custom-XML part
(`Docxodus/DocumentAnnotation.cs`), documented at `docs/architecture/custom_annotations.md`.
Deliberately distinct from `docxodus_comment`: the overlay semantically tags regions for
external tools (e.g. OpenContracts) and never appears in Word's Reviewing UI.

### `docxodus_track_changes` — list/accept/reject tracked changes, switch recording mode

`set_mode` (issue #304) switches how the session records its *own subsequent* edits —
`mode: "accept" | "render_inline" | "strip_deletions"` (the same values `docxodus_open`'s
`trackedChanges` takes), plus optional `revisionAuthor` (absent = leave the current author
unchanged; empty string = reset to the `"docxodus"` default). Backed by
`DocxSession.SetTrackedChanges`/`SetRevisionAuthor`: session configuration, not a document
mutation — not undoable, and already-applied markup is never touched (switching to `accept`
does not resolve existing revisions — that's `accept_all`; switching to `render_inline` does
not retroactively track prior direct edits). The response echoes the now-current state:
`{"success":true,"trackedChanges":"render_inline","revisionAuthor":"Reviewer A"}`. Before this
action existed, flipping the mode meant `docxodus_save` → `docxodus_close` → `docxodus_open`
(and, without `persistAnchorIds`, losing every anchor id at that boundary); that dance is no
longer needed for mode switching.

Once a mutation actually emits native revision markup, the session also enables
`w:trackRevisions` in `settings.xml` (creating the part when absent), so Word keeps tracking later
interactive edits. This setting is distinct from display: `docxodus_get_content(format: "html")`
always renders pending markup as `<ins>`/`<del>`; accepting or rejecting it requires an explicit
track-changes action.

`list` (issue #318) reads the revision set directly off the live session's markup —
`DocxSession.ListRevisions` enumerates `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`, paragraph-mark
and table-row markers, and the `*PrChange` format-change family across body, headers, footers,
footnotes, and endnotes, grouping physically contiguous same-kind/same-author markup into one
entry per user-visible change. Each entry carries a stable `id` (derived from the markup's own
`w:id` attributes, so resolving other revisions never renames it), `type`
(`insert`/`delete`/`move`/`format` — a `move` is a linked pair covering both sides), the
markup's true `author`/`date`, its visible `text`, and the containing block's `anchorId`.
This replaced the original listing, which re-diffed `RevisionProcessor.RejectRevisions` vs
`.AcceptRevisions` output through `DocxDiffOps.GetRevisionsJson` — that shape had no stable
identity to address, substituted engine-default authors/dates for the markup's real ones, and
cost ~3s on a 49-page document. `author`/`changeType` are display-only filters applied after
the fact.

`accept`/`reject` resolve ONE revision by `revisionId` as an ordinary undoable session
mutation (`DocxSession.AcceptRevision`/`RejectRevision` — no whole-document
`RevisionProcessor` round-trip, no session rebind, anchors stay live), returning the standard
EditResult envelope with the affected blocks in `modified`/`removed`. An unknown or
already-resolved id fails with `revision_not_found`. `accept_all`/`reject_all` remain for
whole-document resolution: they transform via `RevisionProcessor` and swap the session's
underlying handle in place (`SessionStore.Rebind`), which also covers the exotic families the
per-revision listing does not enumerate (see Known gaps).

### `docxodus_mutations` — atomic batches, explicit partial apply, or legacy preview

`steps: [{ tool, args }]` where `tool` is one of `docxodus_edit`/`docxodus_format`/
`docxodus_create`/`docxodus_table`/`docxodus_list`/`docxodus_comment` (their `undo`/`redo` and
read-only actions — e.g. `get_membership`, comment `list` — are rejected as steps; a batch is a
sequence of *mutations*).

`mode: atomic` is the recommended default. The server preflights every supported
action, required argument/enum, and step precondition against the batch-start state
before step zero. Success creates one undo entry and advances the version once. Any
failed or thrown step restores the complete DOCX package, relationships, session
state, version, and undo/redo cursors; the receipt identifies the failing
`index`/`tool`/`action`/`error` and reports `rolledBack: true`.

`mode: best_effort` is the explicit partial-success mode. It runs every step in
order and evaluates a step preflight immediately before that step, returning a
`{ status, editsApplied, results, errors }`-compatible receipt (`status` is
`ok`/`partial`/`failed`). `mode: apply` is a deprecated compatibility alias for
`best_effort`; new clients should use the risk-signaling spelling. `mode: preview` runs every step exactly the same way, then calls
`DocxSessionOps.Undo` once per step that actually mutated before returning — see Known gaps for
why this is "apply-then-undo" rather than a true no-op dry run.

The batch itself and each step's `args` may carry `preconditions`, using the same
camel-case guard object as the core API (`expectedVersion`, `anchorId`,
`expectedContentHash`, exact text/range/kind/scope, and `expectedMatchCount`). A
failure is the standard structured `precondition_failed` result. Atomic mode
evaluates all step guards at the common batch-start boundary; best-effort mode
evaluates them sequentially. `docxodus_get_content` with `format: "version"` reads the current
monotonic document version; `format: "check_preconditions"` evaluates guards
without mutating. Preview restores its starting version after undoing speculative
steps, so a dry-run does not make an otherwise-current plan stale.

### `docxodus_table` — tables

`insert`, `insert_row`, `insert_column`, `delete_row`, `delete_column`, `replace_cell_content`,
the read actions `get_metadata`, `resolve_cell_anchor`, `resolve_cell_coordinate`, plus the
post-insert styling actions (issue #315 Stage A): `set_column_widths` (`widths`, one
positive twip value per column), `set_borders` (`borderScope` `all`/`outside`/`inside`,
`borderStyle` — `none` removes the targeted edges — `borderSize`, `borderColor`), `set_shading`
(`fill` hex/`auto`, omit to clear; `shadingScope` `cell`/`row` — row is header-row banding), and
`set_repeat_header_row` (`repeat`, default true), `set_row_options`, `merge_cells`, and
`unmerge_cells`.

The input schema does not overload anchor fields: `insert` uses `anchorId` for the neighboring
block, metadata/coordinate reads use `tableAnchorId` (`tbl`), and every cell operation uses
`cellAnchorId` (`tc`). `insert` returns canonical `tc` anchors in `created`; `get_metadata`
enumerates the explicit `tbl`/`tr`/`col`/`tc` identities and coordinates. A legacy paragraph
inside a cell is still translated during the compatibility window, but new tool calls should use
only the canonical fields and identities. Shape mutations include a deterministic `tableAnchors`
retained/added/invalidated mapping in their result.

## Inline preview (MCP Apps / ChatGPT Apps)

The server implements the MCP Apps extension (`io.modelcontextprotocol/ui`, spec 2026-01-26 —
the joint Anthropic/OpenAI standard for interactive UI in chat hosts) so a compliant host
(Claude, ChatGPT, VS Code, Goose) renders the document inline next to tool results instead of
making the user imagine it from markdown. Implementation lives in `UiResources.cs`; everything
else routes through it.

**The moving parts:**

- **Capability**: `initialize` declares `capabilities.extensions["io.modelcontextprotocol/ui"]`
  (mimeType `text/html;profile=mcp-app`) plus `resources: {}`.
- **Template**: `resources/list`/`resources/read` serve `ui://docxodus/viewer.html` — a fully
  self-contained HTML widget (no external fetches, so it renders under the spec's **default** CSP:
  `script-src 'self' 'unsafe-inline'`, no network; `_meta.ui.csp` declares empty domain lists).
- **Tool linkage**: `docxodus_open` and `docxodus_preview` carry
  `_meta.ui.resourceUri = "ui://docxodus/viewer.html"` in `tools/list`, plus ChatGPT Apps SDK
  compatibility aliases (`openai/outputTemplate`, `openai/widgetAccessible: true`,
  `openai/toolInvocation/*`). OpenAI's docs designate `_meta.ui.resourceUri` as the standard field
  and the `openai/*` keys as optional aliases, so one stamping serves both hosts.
- **Data flow**: `docxodus_open` mirrors `{sessionId, path}` as `structuredContent`; the widget
  instance attached to it reads the session id and fetches its first render by calling
  `docxodus_preview` itself (widget-initiated `tools/call` — MCP Apps default tool visibility is
  `["model", "app"]`). `docxodus_preview` returns the markup only in `_meta["docxodus/html"]`,
  keeping the model's context clean.
- **The viewer** detects its host at runtime: `window.openai` present → ChatGPT bridge
  (`toolOutput`/`toolResponseMetadata`/`callTool`, `openai:set_globals` event); otherwise MCP Apps
  JSON-RPC over `postMessage` (`ui/initialize` handshake, `ui/notifications/tool-input`/
  `tool-result` notifications, `tools/call` requests). Received document HTML is parsed with
  `DOMParser`; the converter's `<style>` blocks move into the widget head (safe — the converter
  prefixes classes `docx-`) and the body is injected. A Refresh button re-invokes
  `docxodus_preview`, which is how an agent-driven edit becomes visible mid-conversation.

**Transports**: stdio works as always (Claude Desktop/Code config unchanged). `--http PORT`
starts the minimal streamable-HTTP binding for hosts that require a hosted server (ChatGPT Apps
developer mode behind a tunnel). It is a development convenience, not a production deployment
story — no TLS, no auth, no SSE.

**Verification**: `smoke/apps_probe.py --server bin/Debug/net10.0/docxodus-mcp.dll --docx <file>`
covers both transports (capability/resource/meta shapes, open→preview→block-preview flow, the
no-markup-in-content invariant). The widget itself was validated in Chromium against a
spec-faithful fake host (handshake → notifications → widget `tools/call` → style-applied render →
Refresh); the harness pattern is documented in the PR that introduced this feature.

**Deliberate v1 limits**: one widget template for both whole-document and single-block renders;
no push — the widget refreshes by polling `docxodus_preview` on demand rather than the server
emitting `notifications/resources/updated`; `structuredContent` is stamped only on the two
widget-bearing tools (a host that wants live-updating previews after every `docxodus_edit` should
have the widget re-call `docxodus_preview` after the model announces an edit, or we add
`_meta.ui.resourceUri` to more tools later).

## Known gaps

Capabilities a full-featured document-editing agent surface might have, that Docxodus's engine
doesn't yet support — called out explicitly rather than faked, per this server's design goal of
never claiming a capability it doesn't have:

- **Exotic revision families aren't individually resolvable.** Issue #318 closed the
  selective-resolution gap for the common families — `docxodus_track_changes` `accept`/`reject`
  resolve one insert/delete/move/format revision by `revisionId` — but
  `w:cellIns`/`w:cellDel`/`w:cellMerge`, content-control ins/del ranges, and `w:numPr`
  numbering-ins markers are not enumerated by `list` and have no per-revision resolution;
  `accept_all`/`reject_all` (whole-document `RevisionProcessor`) still handle them.
- **New lists inserted via a bare markdown payload don't get real Word numbering.** A `"- item"`
  block parses to a `kind: "li"` anchor with no `w:numPr` (documented in
  `docx_mutation_api.md`). This server's `docxodus_list`/`docxodus_create` route around it by
  composing `InsertParagraph` (plain text) + `ApplyListFormat` (which *does* write real
  `w:numPr` via `NumberingFactory.EnsureNumbering`) — two calls, not a gap in what's reachable,
  just not a single one-shot "insert a numbered list" primitive.
- **`docxodus_mutations`'s `preview` mode is apply-then-undo, not a true dry run.** It runs every
  step for real, then calls the session's `Undo()` once per step that mutated. This composes
  correctly with everything else (bounded undo ring, anchor lifecycle) but consumes undo-ring
  depth like any other edit sequence, and a crash between "apply" and "undo" would leave the
  session mutated — acceptable for a local, single-process tool server, worth knowing if this
  surface is ever exposed somewhere more failure-sensitive.

## Testing

`Docxodus.Tests/McpServerDispatcherTests.cs` exercises `Dispatcher.Call` directly (no stdio
transport in the loop — `Program.cs` is a thin JSON-RPC wrapper around it) against a blank
document from `DocxSession.CreateBlankDocxBytes()`, covering the full lifecycle, every grouped
tool's primary actions, the tracked-changes accept/reject round trip, and the mutations
batch/preview paths. Each test class instance is given its own scope root, so the dispatcher runs
against a realistically-confined store rather than an unbounded filesystem.
`tools/mcp-server` also has an `InternalsVisibleTo` grant to `Docxodus.Tests` (mirroring
`Docxodus.csproj`'s existing grants to `docxodus-pyhost` and `DocxodusWasm`) so the test project
can call the dispatcher's internals directly.

The storage layer's isolation rules are covered as their own group (`MCP120`–`MCP131`): relative
and in-scope-absolute locations resolve, out-of-scope absolutes and `..` traversal are rejected, a
sibling directory sharing a name prefix is not treated as inside, a **symlink escaping the root is
rejected**, read/write round-trips (creating intermediate directories), `docxodus_open` and
`docxodus_save` both refuse out-of-scope locations, two scopes under one root cannot reach each
other, and unsafe `DOCXODUS_STORAGE_SCOPE` values and unknown backends are rejected at
construction. The symlink case is worth keeping: it failed against the first implementation, which
resolved only the leaf component and so missed a link in the middle of the path.

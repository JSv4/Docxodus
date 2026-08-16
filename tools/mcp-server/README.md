# docxodus-mcp

A Model Context Protocol (MCP) server that lets AI agents open, read, and edit `.docx`
files through Docxodus's `DocxSession` editing engine. Runs entirely locally over stdio —
no network calls, no telemetry, nothing leaves the machine it runs on.

See `docs/architecture/docx_agent_server.md` for the full tool contract (every tool's
arguments and result shape), the mapping of each action onto the underlying Docxodus API,
and a list of documented capability gaps.

## Installation

```bash
dotnet tool install -g Docxodus.McpServer
```

## Usage

Add it to your MCP client's server configuration. For example, in a client that reads a
`mcpServers` map:

```json
{
  "mcpServers": {
    "docxodus": {
      "command": "docxodus-mcp"
    }
  }
}
```

Or run it directly from source during development:

```bash
dotnet run --project tools/mcp-server
```

The server speaks newline-delimited JSON-RPC 2.0 on stdin/stdout (the MCP stdio transport).
All diagnostic logging goes to stderr, never stdout.

## Document storage and scoping

Every document the server reads or writes goes through a scoped store. A session cannot name a
document outside its scope — that's a property of how locations are resolved, not a check the
tool layer performs, so it holds for every operation.

| Variable | Meaning |
|---|---|
| `DOCXODUS_STORAGE_BACKEND` | Backend id. Only `file` is implemented; defaults to `file`. |
| `DOCXODUS_STORAGE_ROOT` | Directory every location must resolve inside. Defaults to your home directory, so an agent can open `~/Downloads/contract.docx` naturally while the rest of the machine stays out of reach. Narrow it to confine the server; set it to `/` to opt out. |
| `DOCXODUS_STORAGE_SCOPE` | Optional single path segment appended to the root (`{root}/{scope}`). Two servers launched with different scopes cannot see each other's documents. |

Relative locations resolve under the root; an absolute path is accepted only if the root contains
it. Symlinks are followed before the containment check, so a link inside the root pointing outside
it is rejected rather than followed.

Scope is chosen by whoever launches the server, never by the agent — it has no argument through
which to name one. Because the scope is just a stable path segment, passing the same value next
time reaches the same documents; persistence is the filesystem's, with no registry or token to
manage. One server process serves exactly one scope.

To confine a server to a project directory:

```json
{
  "mcpServers": {
    "docxodus": {
      "command": "docxodus-mcp",
      "env": { "DOCXODUS_STORAGE_ROOT": "/srv/documents", "DOCXODUS_STORAGE_SCOPE": "acme-corp" }
    }
  }
}
```

Adding a different backend (object storage, a content repository) means implementing
`IDocumentStore` and adding a case to `DocumentStores` — the tool surface and every session code
path stay as they are.

## Tool surface

Three lifecycle tools plus seventeen document tools, with editing tools addressed by the anchor ids the
markdown projection and search tools return:

| Tool | Purpose |
|------|---------|
| `docxodus_open` / `docxodus_save` / `docxodus_close` | Session lifecycle |
| `docxodus_pagination` | Register, inspect, or query an externally materialized PageMap |
| `docxodus_get_content` | Read markdown/HTML/text, block or section facts, styles, direct/effective formatting, mutation-ready inline spans, a complete deterministic package manifest, the opening-to-current semantic change set, or a default deliverable-verification report (`format: "verification"`) |
| `docxodus_search` | Find text (literal/regex), or blocks by kind/annotation/bookmark |
| `docxodus_edit` | Insert/replace/delete text and blocks, split/merge paragraphs, undo/redo |
| `docxodus_format` | Character and paragraph formatting, list level |
| `docxodus_create` | New paragraphs, headings, tables, horizontal rules, footnotes/endnotes, running headers/footers, page-number fields |
| `docxodus_list` | Promote/demote/renumber list membership; restart numbering (Word's *Set Numbering Value…*) |
| `docxodus_comment` | Native Word review comments (real `w:comment` markup): add on an anchor/span or tracked revision id, reply in-thread, resolve/reopen, update, remove, list |
| `docxodus_annotate` | Anchor-addressed highlight/label annotations (a custom-XML overlay for external tools, distinct from comments) |
| `docxodus_track_changes` | List tracked changes; accept/reject one by id, or all |
| `docxodus_mutations` | Apply or safely preview a batch atomically by default; opt explicitly into best-effort |
| `docxodus_deliver` | Build a verified delivery bundle from a named baseline and the current session; return its manifest and available artifact bytes |
| `docxodus_table` | Create/read tables; resolve canonical cell anchors ↔ grid coordinates; edit rows/columns/cell content/style |

`docxodus_deliver` uses the same `DeliveryBundleService` as the .NET API and
`docxodus-deliver` CLI. The MCP response returns canonical manifest bytes plus available artifacts
as base64, with a 64 MiB pre-base64 byte limit. In this server configuration, HTML/PDF artifacts
and authoritative change receipts are explicitly unavailable: rendering awaits the adapter tracked
by #434, while receipt issuance requires exact transaction snapshots/contributions rather than the
MCP retry journal's response cache. Mark either artifact optional to retain a complete bundle with
truthful unavailability metadata, or use `returnIncompleteBundle` for diagnostic required outputs.

Applying `docxodus_mutations` batches can include a caller-chosen, non-blank root `transactionId` of
at most 256 Unicode scalar values. During the open session, retrying the same canonical request
returns the exact original serialized batch result without applying again or rechecking guards; a
different request with that id fails with
`transaction_conflict`. Results expose
`transaction: { schemaVersion: 1, transactionId, requestFingerprint }`. Preview/dry-run batches and
direct or nested step calls reject transaction ids — a client that blanket-attaches an idempotency
key to *every* tool call gets a hard error on the other tools, so attach it only to applying
`docxodus_mutations` batches. Save preserves this journal; close clears it, and reopen starts a new
identity namespace. Replay after undo/redo returns the historical response without changing the
document or either history cursor; use ordinary redo to restore an undone mutation.

### Retention bound and its memory cost

Retention is bounded per session by **both** a count and a byte budget: at most 128 complete
responses, at most 32 MiB of retained response text, whichever binds first, followed by 1,024
response-less FIFO tombstones that keep an id bound without holding its payload. A single response
larger than the whole budget evicts itself rather than raising the ceiling.

The byte budget exists because a count is not a memory bound. A retained entry is the complete
serialized `MutationBatchResult`, which carries every step's `results` **twice** (once under
`steps[].results` and once in the duplicate top-level `results`) plus `patch.markdown` and the
revision/comment/annotation delta sets. Measured: a single-step `insert_paragraph` batch against a
blank document already retains **~3.2 KB**, so 128 of those is ~400 KB — and a large scoped batch
over a real document is orders of magnitude bigger, which is what the 32 MiB cap is there for. The
worst case is therefore **~32 MiB per open session**; the number of open sessions is *not* bounded,
and there is no TTL or idle-session eviction, so a long-lived server that is never sent
`docxodus_close` still grows with the number of sessions.

### Lifecycle hazards

- **A validation failure burns the id.** A structured rejection (say `mode: "sideways"`) is itself
  a terminal response and is cached. Fixing the typo and resending under the same `transactionId`
  gets `transaction_conflict`, not a retry. Always use a **fresh** id after any failure you intend
  to correct; reuse an id only for a byte-identical resend of the same request.
- **Tombstone expiry lets a stale retry re-apply.** After roughly 128 further transactions plus
  1,024 more tombstones, an id is genuinely forgotten and a retry that arrives after that becomes a
  **fresh mutation that applies again**. Idempotency is a bounded-window guarantee, not a permanent
  one; a client holding a request for a long time must not assume the window is still open.
- **`transaction_incomplete`** means the id is bound to this exact request but no terminal response
  was ever recorded, so whether the mutation applied is *unknown*. Inspect the document and retry
  under a new id.
- **Idempotency is MCP-only.** `execute_batch` through WASM/npm and through the stdio host /
  `docx-scalpel` has no transaction identity and no replay: a retry there re-applies. Do not
  assume the MCP guarantee from another transport.

## Known gaps

A few capabilities a full-featured document-editing agent surface might want are not yet
exposed, because the underlying Docxodus engine doesn't have them (rather than fake them,
these are called out so agents/tooling built against this server know to route around
them):

- **Unsafe revision topology fails closed.** `docxodus_track_changes` lists cell,
  content-control, numbering, text, move, row, and property revisions with stable ids.
  Unsupported, malformed, or ambiguous native markup remains visible with a diagnostic;
  individual and bulk resolution return a typed error without changing the session. This
  applies to `accept_all`/`reject_all` too, which used to be an always-succeeding
  whole-document transform — there is no `force` mode. See "Known gaps" in
  `docs/architecture/docx_agent_server.md` for the list of refusing shapes.
- **New lists inserted via markdown don't get real numbering** unless promoted afterward
  with `docxodus_list`'s `apply_format` action (which does write real `w:numPr`).
- **Generated-id previews are semantic rather than necessarily byte-identical.** Preview runs
  the same batch path on a complete isolated package clone and never touches live bytes,
  version, caches, configuration, or undo/redo history. Create/comment/note/image operations
  may allocate different anchors or OOXML ids when later applied, and tracked changes may stamp
  a different execution time, so receipts warn when exact anchor or package-hash equality is
  unsafe. Deterministic mutation-only batches retain exact result/hash equivalence.
- **Optimistic guards are common to every mutation tool.** Pass `preconditions` with
  `expectedVersion` and/or an anchor hash/exact text/range/kind/scope; replacement may
  also require `expectedMatchCount`. `docxodus_get_content` formats `version` and
  `check_preconditions` expose the read side. Batch-level and per-step guards use the same shape.
- **The inline MCP preview is continuous, not physically paginated.** It reports
  `pageNavigation: "unavailable_continuous_preview"`; a registered PageMap can still supply an
  exact page label, but navigation to a page box awaits the server-side paginated HTML substrate
  tracked by #434. The server does not estimate or bundle a browser. See the
  [PageMap contract](../../docs/architecture/page_map.md).

## License

MIT License

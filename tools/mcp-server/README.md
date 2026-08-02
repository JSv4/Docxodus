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

Three lifecycle tools plus ten grouped-intent tools, each addressed by the anchor ids the
markdown projection and search tools return:

| Tool | Purpose |
|------|---------|
| `docxodus_open` / `docxodus_save` / `docxodus_close` | Session lifecycle |
| `docxodus_get_content` | Read as markdown, HTML, plain text, block metadata, or document info |
| `docxodus_search` | Find text (literal/regex), or blocks by kind/annotation/bookmark |
| `docxodus_edit` | Insert/replace/delete text and blocks, split/merge paragraphs, undo/redo |
| `docxodus_format` | Character and paragraph formatting, list level |
| `docxodus_create` | New paragraphs, headings, tables, horizontal rules, footnotes/endnotes, page-number fields |
| `docxodus_list` | Promote/demote/renumber list membership; restart numbering (Word's *Set Numbering Value…*) |
| `docxodus_comment` | Native Word review comments (real `w:comment` markup): add on a span, update body, remove, list |
| `docxodus_annotate` | Anchor-addressed highlight/label annotations (a custom-XML overlay for external tools, distinct from comments) |
| `docxodus_track_changes` | List tracked changes; accept/reject one by id, or all |
| `docxodus_mutations` | Apply or dry-run-preview a batch of the above as one call |
| `docxodus_table` | Create tables; edit rows/columns/cell content |

## Known gaps

A few capabilities a full-featured document-editing agent surface might want are not yet
exposed, because the underlying Docxodus engine doesn't have them (rather than fake them,
these are called out so agents/tooling built against this server know to route around
them):

- **No comment reply-threading or resolve state.** `docxodus_comment` authors real
  `w:comment` markup, but replies and Word's resolve flag (`commentsExtended.xml`) are not
  yet authorable; existing threading metadata is preserved on update and pruned on remove.
- **Exotic revision families aren't individually resolvable.** `docxodus_track_changes`
  lists and selectively resolves inserts/deletes/moves/format changes by `revisionId`
  (issue #318), but `w:cellIns`/`w:cellDel`/`w:cellMerge`, content-control ins/del
  ranges, and `w:numPr` numbering-ins markers are not enumerated — `accept_all`/
  `reject_all` still resolve those.
- **New lists inserted via markdown don't get real numbering** unless promoted afterward
  with `docxodus_list`'s `apply_format` action (which does write real `w:numPr`).
- **`docxodus_mutations`'s `preview` mode is apply-then-undo**, not a true no-op dry run;
  it uses the session's bounded undo ring, so it composes with everything else but is not
  free of history-depth pressure.

## License

MIT License

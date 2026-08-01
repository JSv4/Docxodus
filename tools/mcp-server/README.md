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
| `docxodus_list` | Promote/demote/renumber list membership |
| `docxodus_comment` | Anchor-addressed highlight/label annotations (not native Word comment threads — see the docs) |
| `docxodus_track_changes` | List tracked changes; accept/reject them all |
| `docxodus_mutations` | Apply or dry-run-preview a batch of the above as one call |
| `docxodus_table` | Create tables; edit rows/columns/cell content |

## Known gaps

A few capabilities a full-featured document-editing agent surface might want are not yet
exposed, because the underlying Docxodus engine doesn't have them (rather than fake them,
these are called out so agents/tooling built against this server know to route around
them):

- **No native Word review-comment threads.** `docxodus_comment` creates a bookmark +
  custom-XML highlight overlay, not `w:comment` elements with Word-native reply/resolve
  semantics.
- **No selective tracked-change resolution.** `docxodus_track_changes` can list revisions
  filtered by author/type for display, but `accept_all`/`reject_all` apply to the whole
  document — there is no "accept only this author's inserts" primitive.
- **New lists inserted via markdown don't get real numbering** unless promoted afterward
  with `docxodus_list`'s `apply_format` action (which does write real `w:numPr`).
- **`docxodus_mutations`'s `preview` mode is apply-then-undo**, not a true no-op dry run;
  it uses the session's bounded undo ring, so it composes with everything else but is not
  free of history-depth pressure.

## License

MIT License

# Round-three DOCX MCP smoke workflow

This directory contains a reproducible replacement-server workflow for the October 2025 model certificate of incorporation. It is intentionally heterogeneous: the sequence reads a large existing package, previews and rolls back a mutation batch, accepts and rejects tracked changes, performs undo/redo, formats runs and paragraphs, authors a resolved comment thread, creates and restarts a nested Roman list, and creates a styled table with explicit row-layout properties.

The workflow does not assume stable anchor IDs. It discovers anchors by content, captures IDs from tool responses, and substitutes captured values into later calls. The full sequence has 56 tool calls and saves back to `local.docx` under the configured storage root.

## Files

- `round-three-workflow.json` — replacement-server tool sequence.
- `round-three-validation.json` — independent close/reopen assertions through the MCP surface.
- `mcp_probe.py` — generic stdio JSON-RPC runner with variable capture and a full trace.
- `analyze_docx.py` — package and semantic-observable comparator for source, reference, and replacement DOCX files.
- `RESULTS.md` — the verified comparison snapshot and material differences.

## Run

Build the server, place a pristine source document at `$DOCXODUS_STORAGE_ROOT/local.docx`, and run:

```bash
dotnet build tools/mcp-server/mcpserver.csproj

python3 tools/mcp-server/smoke/mcp_probe.py \
  --calls tools/mcp-server/smoke/round-three-workflow.json \
  --trace /tmp/round-three-local-trace.json \
  -- tools/mcp-server/bin/Debug/net10.0/docxodus-mcp

python3 tools/mcp-server/smoke/mcp_probe.py \
  --calls tools/mcp-server/smoke/round-three-validation.json \
  --trace /tmp/round-three-local-validation-trace.json \
  -- tools/mcp-server/bin/Debug/net10.0/docxodus-mcp
```

The runner exits nonzero for a transport error, an MCP tool error, an edit result with `success: false`, a mutation batch with `failed`/`partial` status, or a failed workflow `expect` assertion.

Given the pristine source, a separately produced reference output, and the replacement output, compare their packages with:

```bash
python3 tools/mcp-server/smoke/analyze_docx.py \
  source.docx reference.docx local.docx
```

The comparator reports exact package differences and a workflow-equivalence verdict. Equivalence requires identical expected/rejected markers, comment-thread state, revision state, and table properties; column widths permit a 5-twip rounding tolerance.

## Coverage

The edit path exercises:

- paragraph, heading, list-item, table, cell, comment, and revision anchors;
- atomic preview rollback and apply batches;
- direct and tracked text replacement, revision enumeration, accept, and reject;
- undo and redo;
- bold, italic, underline, strikeout, superscript, subscript, color, font size, and font family;
- alignment, indentation, first-line indentation, before/after spacing, and line spacing;
- root/reply comment creation, update, resolve, persistence, and reopen;
- upper-Roman numbering, nesting, and start override;
- table creation, unequal fixed widths, borders, row shading, repeat header, minimum row height, no page split, header text color/bold/alignment;
- preservation checks for bookmarks, fields, four sections, 94 footnote references, eight header parts, and ten footer parts.

The source is intentionally not checked into this directory. Retrieve it from <https://nvca.org/wp-content/uploads/2025/10/NVCA-Model-COI-10-1-2025.docx>.

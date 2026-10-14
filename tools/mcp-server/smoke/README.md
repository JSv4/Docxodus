# DOCX MCP smoke workflows

Two workflows live here. The **epic #435 acceptance run** is the one to reach for when
verifying the agent-editing surface end to end; the **round-three run** below predates it
and exercises a broader but less guarded operation mix.

## Epic #435 acceptance run

`epic-435-workflow.json` walks the acceptance gate for *Agent-safe headless document
editing* — inspect, isolated preview, atomic apply under a transaction id, a retry proving
byte-exact replay, an intentional stale request, an intentional atomic failure, an audit
that nothing moved, a page citation, then save. `epic-435-validation.json` reopens the
saved package and checks what persisted.

Neither file is edited by hand. Both are emitted by `build_epic_435_fixtures.py`, for two
reasons: the preview, the apply and the retry must send a byte-identical step array or the
retry is a `transaction_conflict` rather than a replay; and the two files assert the *same*
revision list, which drifted out of both when each was maintained separately (#687). That
list now has one declaration, `REVISION_MEMBERS`.

`scripts/mcp-smoke.sh` runs the whole gate — regenerate, check the committed fixtures are
what the generator emits, run both workflows, and re-check the refusal and replay counts
the runner reports but does not enforce. CI runs it on every pull request; run it locally
the same way:

```bash
dotnet build tools/mcp-server/mcpserver.csproj -c Release
./scripts/mcp-smoke.sh
```

To drive the two runs by hand instead:

```bash
dotnet build tools/mcp-server/mcpserver.csproj
python3 tools/mcp-server/smoke/build_epic_435_fixtures.py

# Copy the source document into a scratch storage root — the run SAVES, and pointing
# the root at TestFiles/ would overwrite a committed fixture.
export DOCXODUS_STORAGE_ROOT=$(mktemp -d)
cp TestFiles/NVCA-Model-COI.docx "$DOCXODUS_STORAGE_ROOT/local.docx"

python3 tools/mcp-server/smoke/mcp_probe.py \
  --calls tools/mcp-server/smoke/epic-435-workflow.json \
  --trace /tmp/epic435-trace.json \
  --quiet-server -- tools/mcp-server/bin/Debug/net10.0/docxodus-mcp

python3 tools/mcp-server/smoke/mcp_probe.py \
  --calls tools/mcp-server/smoke/epic-435-validation.json \
  --trace /tmp/epic435-validation-trace.json \
  --quiet-server -- tools/mcp-server/bin/Debug/net10.0/docxodus-mcp
```

Expected: 64 calls / 211 assertions and 15 calls / 24 assertions, both with zero failures,
five expected failures in the first, and `replayMismatches: 0`.

### The six revisions

The five tracked steps settle into six revisions, and the count surprises people, so it is
worth stating: `insert_footnote` accounts for two of them. The note body lands in
`word/footnotes.xml`, and the reference run lands in the body — a real tracked insertion
whose `text` is empty, because a `w:footnoteReference` carries no text to report. Rejecting
that revision is what takes the reference back out, so suppressing it from the list would
hide a change an agent is entitled to see.

The order the engine enumerates them in is not part of the contract, so the fixtures match
them as members (see `expectMembers` below) keyed on type, text, part and scope.

Five calls are *supposed* to fail, and the runner treats a success there as the defect
(`unexpectedSuccesses`): a bookmark and a table insert refused under tracked recording, a
bookmark endpoint refused inside a tracked insertion under direct recording, a stale
precondition, and the batch that rolls back. The engine failing those closed is the property
under test.

### Workflow call fields

Beyond `name`/`arguments`/`capture`/`expect`, honored by `mcp_probe.py`:

| Field | Meaning |
|---|---|
| `expectFailure` | This call must fail. Its failure stops counting against the run, and a *success* becomes an `unexpectedSuccess`. |
| `expectSameAs` | Compare this call's raw response text to that of the named earlier call. Byte-exact, before JSON parsing normalizes key order — which is what transaction replay actually promises. |
| `expectNonEmpty` | List of result paths that must resolve to a non-empty string, array, or object. Used for generated payloads such as preview HTML and package hashes where pinning fixture-specific bytes would be brittle. |
| `expectMembers` | Map of list path → expected member objects. Each member must match exactly one not-yet-claimed entry on that list, in any order, comparing only the keys the member names. Use it where a collection's contents are the contract but its enumeration order is not. |

`expect` values substitute `$variables` too, so one call can be asserted against another's
captured value (the apply's `resultVersion` against the preview's prediction).

## Round-three DOCX MCP smoke workflow

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

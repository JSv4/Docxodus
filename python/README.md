# docx-scalpel

**Anchor-addressed DOCX editing for LLM agents — a thin client over [Docxodus](https://github.com/JSv4/Docxodus)' `DocxSession`.**

`docx-scalpel` exposes Docxodus' stateful DOCX editor over a long-running .NET subprocess (`docxodus-pyhost`). The session lives in the host's memory until you explicitly release it, so an LLM agent can issue dozens of small edits against one document without paying the OOXML parse + Unid annotation + projection cost on every call.

> **Status:** Beta. Wheels ship a bundled `docxodus-pyhost` for linux-x64, linux-arm64, osx-arm64, and win-x64; any other platform installs from the sdist and needs a host of its own (see below).

## Installation

```bash
pip install docx-scalpel
```

That resolves a wheel on linux-x64, linux-arm64, osx-arm64, or win-x64 — each carrying a self-contained `docxodus-pyhost` built from the same commit as the release, so there's no .NET runtime to install and no version to pin.

Source installs (`pip install` of the sdist, or `pip install -e .` from a dev clone) don't include a bundled host. Set `DOCXODUS_HOST=/path/to/docxodus-pyhost` to point at one you built, or run `dotnet build tools/python-host/pyhost.csproj` inside a Docxodus monorepo clone — the locator auto-discovers it.

## Quick start

```python
from docx_scalpel import open_session, FormatOp, Position

with open("contract.docx", "rb") as f:
    docx_bytes = f.read()

with open_session(docx_bytes) as session:
    # Walk template placeholders and fill them. The picker returns a string to
    # replace, or None to skip. fill_placeholders handles reverse-offset
    # ordering, $-prefix preservation, and multi-pass nested-bracket convergence
    # in one call.
    result = session.fill_placeholders(lambda p: "filled value")
    print(f"filled {result.filled} placeholders in {result.passes} passes")

    # Add a heading after the first body paragraph.
    proj = session.project()
    first_p = next(
        t for t in proj.anchor_index.values()
        if t.kind in ("p", "h") and t.scope == "body"
    )
    session.insert_paragraph(first_p.id, Position.AFTER, "## Reviewed by counsel")

    # Bold the first 8 characters of that paragraph.
    session.apply_format_by_substring(first_p.id, "Reviewed", FormatOp(bold=True))

    new_bytes = session.save()

with open("filled.docx", "wb") as f:
    f.write(new_bytes)
```

### Atomic mutation batches

Use `execute_batch` when a plan spans several edits that must either all land or
all disappear. Atomic is the default; success is one version/undo unit and failure
returns the indexed operation error after restoring the complete package and
history state:

```python
from docx_scalpel import MutationBatchStep

result = session.execute_batch([
    MutationBatchStep("replace_text", {
        "anchorId": first_p.id,
        "markdown": "Replacement text",
    }),
    MutationBatchStep("set_header_text", {
        "anchorId": first_p.id,
        "kind": "default",
        "markdown": "Confidential",
    }),
])
if not result.success:
    print(result.failure.index, result.failure.action, result.failure.error)
```

Select `MutationBatchMode.BEST_EFFORT` explicitly only when retaining successful
steps after another step fails is intended.

### Isolated previews

`preview_batch` takes the same steps and predicts their outcome on a complete clone of the
package. The live session is never a mutation target, so its bytes, version and undo/redo
history are untouched whatever the steps do:

```python
from docx_scalpel import MutationPreviewHtmlMode

preview = session.preview_batch(
    [
        MutationBatchStep("replace_text", {
            "anchorId": first_p.id,
            "markdown": "Proposed replacement",
        }),
    ],
    html_mode=MutationPreviewHtmlMode.FULL,
)
print(preview.html)
print(preview.revision_changes.added, preview.comment_changes.added, preview.warnings)
```

Both `preview_batch` and `execute_batch` return the enriched receipt: `base_version`,
`result_version`, `package_hash`, `{added, removed, modified}` change sets for revisions,
comments and annotations, and `warnings`. `MutationPreviewHtmlMode.SCOPED` renders one block
and requires `html_anchor_id`. `package_hash` is `None` — never `""` — when it could not be
computed, so check it before using it as a replay assertion.

A step's `operation` is any mutating session operation, including the structural
table ops (`insert_table`, `insert_table_row`, `merge_cells`, …); read-only
operations, `undo`/`redo`, and session configuration are rejected as
`invalid_batch_step`.

The `with` block is the documented lifecycle path — it calls `session.close()` on the way out, which releases the session from the host's `SessionRegistry`. A `__del__` finalizer is a fallback for forgotten sessions but should not be relied on; interpreter shutdown may skip it.

## Why a subprocess?

`DocxSession` holds a parsed `WordprocessingDocument`, an `AnchorIndex` of Unid-stamped block-level targets, a cached `MarkdownProjection`, and a bounded `UndoRing` of per-part XDocument snapshots. Recreating it costs tens of ms on small docs and seconds on large ones. The subprocess model lets one Python process drive many sessions across many calls, all in one host's memory, until you decide to close them.

Architecture:

```
Python process                 docxodus-pyhost (.NET 10)
─────────────                  ──────────────────────────
DocxSession  ──NDJSON──>       Dispatcher
                               │
                               ▼
                               DocxSessionOps
                               │
                               ▼
                               SessionRegistry (handle → DocxSession)
```

One host per Python process. Many sessions inside the host. `atexit` sends `shutdown` and (if the host doesn't comply) terminates / kills.

Full design + wire-protocol spec: [`docs/architecture/python_docxodus.md`](../docs/architecture/python_docxodus.md).
Delta-spec for the `docx-scalpel` rebrand: [`docs/superpowers/specs/2026-05-26-docx-scalpel-design.md`](../docs/superpowers/specs/2026-05-26-docx-scalpel-design.md).

## Development

### Build the host binary (one-time)

```bash
# From the Docxodus repo root:
dotnet build tools/python-host/pyhost.csproj -c Release
```

This produces `tools/python-host/bin/Release/net10.0/docxodus-pyhost`. `_host_locator.py` discovers it automatically when you `pip install -e .` from a monorepo clone.

For non-monorepo development, set `DOCXODUS_HOST=/path/to/docxodus-pyhost` to override the discovery path.

A `dotnet build` host is framework-dependent, so it needs the .NET 10 runtime at launch. If your *system* `dotnet` is older and .NET 10 lives elsewhere (e.g. `~/.dotnet`), the host will exit with `You must install or update .NET to run this application`; export `DOTNET_ROOT` to point at the newer install. Released wheels are unaffected — they bundle a self-contained host with no runtime lookup.

```bash
export DOTNET_ROOT="$HOME/.dotnet"
```

### Editable install + tests

```bash
cd python
python -m venv .venv
.venv/bin/pip install -e .[test]
.venv/bin/pytest -v
```

### Test layout

- `tests/test_smoke.py` — end-to-end mirror of `Docxodus.Tests/DocxSessionSmokeTest.cs`. v1 acceptance gate.
- `tests/test_lifecycle.py` — proves session persistence, idempotent close, singleton host, finalizer fallback.
- `tests/test_table_addressing.py` — canonical table identities, coordinate resolution, every table mutation, mappings, and anchor-stable reopen.

Tests share the Docxodus monorepo's `TestFiles/` corpus so divergence between Python and .NET on identical inputs is detectable.

## API surface

The `DocxSession` class exposes every op in `Docxodus.Internal.DocxSessionOps` as a snake-case method:

| Tier | Methods |
|---|---|
| **Lifecycle** | `save`, `close`, `undo`, `redo`, `get_version`, `execute_batch`, `preview_batch`, `to_html`, `register_page_map`, `get_page_map_status`, `get_page_citation` |
| **Projection** | `project`, `project_anchor` |
| **Discovery** | `grep`, `grep_cross_block`, `find_placeholders`, `find_by_text`, `find_all_by_text`, `find_by_regex`, `find_by_kind`, `find_by_annotation`, `find_by_label`, `find_by_bookmark`, `list_annotations`, `exists`, `get_anchor_info`, `get_anchor_infos`, `get_edit_summary`, `remaining_placeholders`, `get_diff` |
| **Inspection** | `get_block_metadata`, `get_block_metadatas`, `get_list_membership`, `get_section_info` |
| **A: text mutations** | `replace_text`, `replace_text_range`, `replace_text_at_span`, `replace_inner`, `replace_match`, `delete_block`, `move_block`, `delete_range`, `delete_section` |
| **B: structural** | `insert_paragraph`, `split_paragraph`, `merge_paragraphs` |
| **B: headers/footers/page numbers** | `set_header_text`, `set_footer_text`, `ensure_header_footer_visible`, `insert_page_number_field`, `set_page_numbering`, `clear_page_numbering` |
| **B: footnotes/endnotes** | `insert_footnote`, `insert_endnote` |
| **B: native comments** | `add_comment`, `add_comment_to_revision`, `add_comment_reply`, `update_comment`, `set_comment_resolved`, `remove_comment`, `list_comments` |
| **C: formatting** | `apply_format`, `apply_format_by_substring`, `set_paragraph_style`, `set_paragraph_format`, `set_list_level`, `remove_list_membership`, `apply_list_format`, `apply_list_format_range`, `set_list_start_override`, `clear_list_start_override` |
| **D: tables** | `get_table_metadata`, `resolve_table_cell_anchor`, `resolve_table_cell_coordinate`, `insert_table`, `insert_table_row`, `insert_table_column`, `delete_table_row`, `delete_table_column`, `merge_cells`, `unmerge_cells`, `set_column_widths`, `set_table_borders`, `set_cell_shading`, `set_repeat_header_row`, `set_table_row_options`, `replace_cell_content` |
| **D: tracked changes** | `set_tracked_changes`, `set_revision_author`, `list_revisions`, `accept_revision`, `reject_revision` |
| **E: annotations** | `add_annotation`, `remove_annotation`, `update_annotation`, `move_annotation` |
| **Raw XML** | `session.raw.get_xml`, `session.raw.insert_xml`, `session.raw.replace_xml` |

Every mutation method returns an `EditResult` envelope — transport-level failures raise `DocxodusTransportError`, but a business outcome (`anchor_not_found`, `malformed_markdown`, etc.) returns `EditResult(success=False, error=EditError(...))`. **Never** an exception across the API boundary.

`PageMap` accepts physical pagination materialized by an external renderer. Registration requires
the session's exact document version and validates the renderer fingerprint, page/section order,
canonical anchors, geometry, story, table ownership, and fragment order. Pass the same
`PageCitationRequest` to search/scoped reads to attach citations. Continuous/no-map and stale
layouts return typed unavailable results; the client never guesses page numbers. See the
[portable PageMap contract](../docs/architecture/page_map.md).

For optimistic concurrency, build a `MutationPreconditions` object and use
`session.check_preconditions(...)` for a read-only probe or
`with session.preconditioned(guards): ...` to attach it to each mutation request in
the block. The guard can require the document version, anchor hash/exact visible
text or range/kind/scope, and an exact replacement match count. A mismatch returns
`EditErrorCode.PRECONDITION_FAILED` with structured expected/actual/current target
metadata and leaves bytes, version, and undo history unchanged.

### Stateless functions

Alongside the session API, the package exposes stateless one-shot functions at the module root — no session handle, they take DOCX bytes in and return bytes / data out:

| Function | Signature | Returns |
|---|---|---|
| `convert_docx_to_html` | `(data, options=None)` | HTML `str` |
| `docx_diff_compare` | `(left, right, settings=None)` | redlined DOCX `bytes` (native `w:ins`/`w:del`/`w:moveFrom`/`w:moveTo`/`w:rPrChange` markup) |
| `docx_diff_get_revisions` | `(left, right, settings=None)` | `tuple[DocxDiffRevision, ...]` |
| `docx_diff_get_edit_script` | `(left, right, settings=None)` | edit-script JSON `str` |
| `docx_diff_accept_revisions` | `(redline)` | `bytes` — accept every tracked change (≡ the right side of the diff) |
| `docx_diff_reject_revisions` | `(redline)` | `bytes` — reject every tracked change (≡ the left side) |
| `docx_diff_consolidate` | `(base, reviewers, settings=None)` | multi-author redlined DOCX `bytes` — merge N `DocxDiffReviewer` diffs against one shared base |
| `docx_diff_get_conflicts` | `(base, reviewers, settings=None)` | `tuple[DocxDiffConflict, ...]` |
| `docx_diff_get_consolidated_revisions` | `(base, reviewers, settings=None)` | `tuple[DocxDiffConsolidatedRevision, ...]` |
| `docx_diff_get_consolidated_edit_script` | `(base, reviewers, settings=None)` | edit-script JSON `str` |

The `docx_diff_*` family is a thin client over Docxodus' `DocxDiff` IR diff engine. Tune pairwise comparisons with `DocxDiffSettings` and N-way consolidation with `DocxDiffConsolidateSettings` (whose `conflict_resolution` takes a `ConflictResolution` value). `DetectMoves`/format-change tracking, header/footer comparison, and per-reviewer attribution all round-trip through these calls.

```python
from docx_scalpel import docx_diff_compare, docx_diff_get_revisions, DocxDiffSettings

with open("v1.docx", "rb") as f: left = f.read()
with open("v2.docx", "rb") as f: right = f.read()

redline = docx_diff_compare(left, right, DocxDiffSettings(author_for_revisions="Reviewer"))
for rev in docx_diff_get_revisions(left, right):
    print(rev.type, rev.text)
```

## License

MIT. Built on top of [Docxodus](https://github.com/JSv4/Docxodus), which is itself a fork of Open-Xml-PowerTools.

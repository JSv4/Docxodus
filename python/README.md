# docx-scalpel

**Anchor-addressed DOCX editing for LLM agents — a thin client over [Docxodus](https://github.com/JSv4/Docxodus)' `DocxSession`.**

`docx-scalpel` exposes Docxodus' stateful DOCX editor over a long-running .NET subprocess (`docxodus-pyhost`). The session lives in the host's memory until you explicitly release it, so an LLM agent can issue dozens of small edits against one document without paying the OOXML parse + Unid annotation + projection cost on every call.

> **Status:** Alpha. linux-x64 wheels ship with a bundled `docxodus-pyhost`; other RIDs require a dev clone of Docxodus until the wheel matrix is extended (tracked in `RELEASING.md`).

## Installation

```bash
pip install docx-scalpel
```

linux-x64 today; pre-release tags ship as `0.1.0a*`, so use `--pre` if you want to opt in:

```bash
pip install --pre docx-scalpel
```

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

Tests share the Docxodus monorepo's `TestFiles/` corpus so divergence between Python and .NET on identical inputs is detectable.

## API surface

The `DocxSession` class exposes every op in `Docxodus.Internal.DocxSessionOps` as a snake-case method:

| Tier | Methods |
|---|---|
| **Lifecycle** | `save`, `close`, `undo`, `redo` |
| **Projection** | `project`, `project_anchor` |
| **Discovery** | `grep`, `grep_cross_block`, `find_placeholders`, `find_by_text`, `find_all_by_text`, `find_by_regex`, `find_by_kind`, `find_by_annotation`, `find_by_label`, `find_by_bookmark`, `list_annotations`, `exists`, `get_anchor_info`, `get_anchor_infos`, `get_edit_summary`, `remaining_placeholders`, `get_diff` |
| **A: text mutations** | `replace_text`, `replace_text_range`, `replace_text_at_span`, `replace_inner`, `replace_match`, `delete_block`, `delete_range`, `delete_section` |
| **B: structural** | `insert_paragraph`, `split_paragraph`, `merge_paragraphs` |
| **C: formatting** | `apply_format`, `apply_format_by_substring`, `set_paragraph_style`, `set_list_level`, `remove_list_membership` |
| **D: tables** | `replace_cell_content` |
| **Raw XML** | `session.raw.get_xml`, `session.raw.insert_xml`, `session.raw.replace_xml` |

Every mutation method returns an `EditResult` envelope — transport-level failures raise `DocxodusTransportError`, but a business outcome (`anchor_not_found`, `malformed_markdown`, etc.) returns `EditResult(success=False, error=EditError(...))`. **Never** an exception across the API boundary.

## License

MIT. Built on top of [Docxodus](https://github.com/JSv4/Docxodus), which is itself a fork of Open-Xml-PowerTools.

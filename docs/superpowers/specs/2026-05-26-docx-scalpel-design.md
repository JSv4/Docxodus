# docx-scalpel — Python Wrapper for `DocxSession`

**Date:** 2026-05-26
**Branch:** `feat/python-docx-scalpel`
**Status:** Scaffold in progress.

## Source of truth

The full architecture is specified in [`docs/architecture/python_docxodus.md`](../../architecture/python_docxodus.md): wire protocol (NDJSON over stdio), subprocess model (one host per Python process, many sessions inside it), lifecycle, type mapping, distribution (per-RID wheels with bundled self-contained `docxodus-pyhost`), and testing strategy. **This document captures only the deltas** from that spec.

## Deltas from `python_docxodus.md`

### 1. Package name: `docx-scalpel` (not `docxodus`)

| Surface | Was (design doc) | Now |
|---|---|---|
| PyPI distribution name | `docxodus` | `docx-scalpel` |
| Python import name | `import docxodus` | `import docx_scalpel` |
| Public entrypoint | `from docxodus import open_docx_session, DocxSession` | `from docx_scalpel import open_session, DocxSession` |
| Repo subdir | `python/` | `python/` (unchanged) |
| Host binary name | `docxodus-pyhost` | `docxodus-pyhost` (unchanged — it ships with Docxodus core) |
| Env var for binary override | `DOCXODUS_HOST` | `DOCXODUS_HOST` (unchanged — it's a Docxodus-side concept) |

**Rationale.** `DocxSession` is the LLM-friendly editing slice of Docxodus, not the whole library. A distinct PyPI name signals that and avoids users confusing it with the (possible-future) full `.NET`-parity Python package. The Python source still lives inside the Docxodus monorepo because (a) it tracks `DocxSessionOps` 1:1 and any change there must propagate here in the same PR, (b) `TestFiles/` is the canonical fixture corpus, and (c) the host binary is built from `tools/python-host/` in this repo.

### 2. Public symbol naming

The design doc says `open_docx_session(...)`; `docx-scalpel` exposes it as `open_session(...)` since the package name already supplies the "docx" context. The class stays `DocxSession` (matches the C# type — there is no ambiguity to resolve).

### 3. Everything else: per the design doc

- Wire protocol, op names (snake_case), arg keys (camelCase) — unchanged.
- Subprocess model (singleton lazy-spawned, `threading.Lock`, `atexit` shutdown, stderr drained to `logging`) — unchanged.
- Session lifecycle — unchanged. **Sessions persist in the host's `SessionRegistry` until `session.close()` is called or the host process exits.** This is the load-bearing requirement: an LLM agent issues dozens of small edits to one document; recreating the session would pay tens-of-ms-to-seconds of OOXML parse + Unid annotation + projection cost each time.
- Type mapping (frozen-slots dataclasses, camelCase→snake_case on decode) — unchanged.
- Test strategy (mirror `DocxSessionSmokeTest.cs`, share `TestFiles/` fixture corpus) — unchanged.
- Distribution (per-RID wheels, `cibuildwheel`, bundled self-contained `docxodus-pyhost` extracts .NET runtime on first launch) — unchanged.

## Scaffold scope (this branch)

The "Implementation sequencing" section of the design doc lists 6 steps. This branch delivers steps 1–3 + part of 6 (skeleton README); steps 4–5 (per-RID build script, `cibuildwheel`) follow in a separate branch when we're ready to cut a wheel.

1. `python/pyproject.toml` (hatchling, stdlib-only runtime deps, `pytest` test extra).
2. `python/src/docx_scalpel/` package: `__init__.py`, `_transport.py`, `_host_locator.py`, `errors.py`, `enums.py`, `types.py`, `session.py`, `py.typed`.
3. `python/tests/`: `conftest.py`, `test_smoke.py` (mirrors `DocxSessionSmokeTest.cs`), `test_lifecycle.py` (session persists across calls; `atexit` releases; many-sessions-in-one-host).
4. `python/scripts/build_host.sh` (per-RID `dotnet publish` to `python/vendor/<rid>/`).
5. `python/README.md` (quick start + lifecycle contract).

**Out of scope on this branch:** `cibuildwheel` matrix, GitHub Actions wheel-build workflow, PyPI publishing, the async API, schema validation, the HTML/comparison/chart surface.

## Verification

The scaffold is "done" when:

1. `cd python && pip install -e .[test]` succeeds with no .NET runtime install required for development.
2. `pytest tests/test_lifecycle.py -v` passes, exercising open → mutate → close on 50 sessions in one host process with zero orphans and confirming the host process count stays at 1.
3. `pytest tests/test_smoke.py -v` passes the full Tier-A/B/C/Raw/Undo workflow against a real `TestFiles/` fixture, with byte-identical save-reopen round-trip.

Tests run against the existing `tools/python-host/bin/Debug/net8.0/docxodus-pyhost` binary (located via the dev fallback in `_host_locator.py`) — no wheel build required for the dev loop.

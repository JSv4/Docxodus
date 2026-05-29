# python-docxodus — Planned Python Wrapper for `DocxSession`

**Status:** .NET-side foundation landed (`Docxodus.Internal.{SessionRegistry, DocxSessionOps, DocxSessionJson}` + `tools/python-host/` stdio NDJSON host). Python package itself is a follow-up PR.

## Context

`DocxSession` is a stateful, in-memory DOCX editing API. The session holds a parsed `WordprocessingDocument`, an `AnchorIndex` of Unid-stamped block-level targets, a cached `MarkdownProjection`, and a bounded `UndoRing` of per-part XDocument snapshots. Recreating it pays the OOXML parse + Unid annotation + projection cost (tens of ms on small docs, seconds on large ones), so the wrapper has to preserve the in-memory session object across many Python calls — critical for agentic LLM pipelines that issue dozens of small edits to one document.

A WASM bridge for the same surface already exists (`wasm/DocxodusWasm/DocxSessionBridge.cs`). The Python wrapper consumes the **same JSON wire shapes** as the WASM bridge: adding a new `DocxSession` op is one edit in `DocxSessionOps`, and both transports pick it up automatically. We never want a Python-specific serializer drifting from the TypeScript one.

## Architecture (already in place)

```
Python client                Browser JS client
     │                              │
     │ NDJSON stdio                 │ JSExport calls
     ▼                              ▼
docxodus-pyhost              DocxodusWasm.DocxSessionBridge
     │                              │
     └──────────────┬───────────────┘
                    ▼
   Docxodus.Internal.{SessionRegistry, DocxSessionOps, DocxSessionJson}
                    ▼
              DocxSession
```

- **`SessionRegistry`** — `ConcurrentDictionary<int, DocxSession>` handle pool, one process can host many sessions.
- **`DocxSessionOps`** — per-op facade combining registry lookup + `DocxSession` call + JSON serialize.
- **`DocxSessionJson`** — StringBuilder JSON helpers (parsers and serializers), camelCase keys.
- **`tools/python-host/`** — .NET 8 console exe `docxodus-pyhost`. Reads NDJSON on stdin, writes NDJSON on stdout, diagnostics on stderr.

## Wire protocol

NDJSON over stdio — one JSON object per line, `\n` terminated.

**Request:** `{"id": <int>, "op": <string>, "args": <object>}`

**Success:** `{"id": <int>, "ok": true, "result": <any>}`

**Failure (transport):** `{"id": <int>, "ok": false, "error": {"code": <str>, "message": <str>, "trace"?: <str>}}`

**Critical distinction:** an `EditResult` with `"success": false` is a *successful* RPC carrying a normal business outcome (`anchor_not_found`, `malformed_markdown`, etc.) — `{"ok": true, "result": {"success": false, "error": {"code": "anchor_not_found", ...}}}`. Python raises on transport failures, returns an `EditResult` dataclass for business outcomes.

**Bytes:** `open_session.args = {"docxB64": "..."}`; `save` response = `{"docxB64": "..."}`. Base64 inside JSON. ~33% overhead, acceptable for the agentic edit-loop pattern.

**Op names** (~38, all snake_case): `open_session`, `close_session`, `save`, `ping`, `shutdown`, `project`, `replace_text`, `delete_block`, `insert_paragraph`, `split_paragraph`, `merge_paragraphs`, `apply_format`, `apply_format_by_substring`, `set_paragraph_style`, `set_list_level`, `remove_list_membership`, `replace_cell_content`, `raw_get_xml`, `raw_insert_xml`, `raw_replace_xml`, `grep`, `grep_cross_block`, `replace_text_range`, `replace_text_at_span`, `find_placeholders`, `find_by_annotation`, `find_by_label`, `find_by_bookmark`, `list_annotations`, `add_annotation`, `remove_annotation`, `update_annotation`, `move_annotation`, `exists`, `get_anchor_info`, `get_anchor_infos`, `find_by_text`, `find_all_by_text`, `find_by_regex`, `find_by_kind`, `undo`, `redo`, `convert_to_html`, `session_to_html`.

## Planned Python package layout

```
python/
  pyproject.toml          # hatchling backend; stdlib-only deps
  src/docxodus/
    __init__.py           # re-export DocxSession, open_session, dataclasses
    session.py            # public DocxSession class — snake_case methods, 1:1 with C# surface
    types.py              # @dataclass(frozen=True, slots=True) value types
    enums.py              # Position, EditErrorCode, TrackedChangeMode, PlaceholderKind/Kinds, ProjectionScopes, etc.
    errors.py             # DocxodusTransportError, DocxodusInstallError
    _transport.py         # singleton subprocess + NDJSON request/response loop, threading.Lock
    _host_locator.py      # find bundled binary (DOCXODUS_HOST env var / vendored / dev fallback)
    _bin/                 # populated by wheel-build with docxodus-pyhost binary for this RID
    py.typed
  tests/
    conftest.py           # TEST_FILES = ../../TestFiles  (shares C# fixture corpus)
    test_smoke.py         # mirror Docxodus.Tests/DocxSessionSmokeTest.cs line-for-line
    test_grep.py
    test_format.py
    test_round_trip.py
    test_errors.py        # EditError surfaces as EditResult, not exception
    test_lifecycle.py     # 50 sessions in one host process; no orphans on exit
    test_transport.py     # malformed reply → DocxodusTransportError
  scripts/
    build_host.sh         # dotnet publish per-RID → python/vendor/<rid>/
  vendor/                 # per-RID dotnet publish output, picked at wheel-build
```

## Subprocess model

**One host process per Python process; many DocxSession handles inside it.**

- Lazy-spawned by `_Transport.get()` on first `open_session(...)` call.
- `atexit` sends `shutdown` then `wait(timeout=2)` then `terminate()` then `kill()`.
- Stderr drained on a daemon thread into Python's `logging`.
- v1 is sync: a `threading.Lock` serializes one request/response round-trip at a time so two threads sharing a session don't interleave NDJSON lines.
- Request IDs are monotonic per process; the response reader matches replies to callers (so async v0.2 can lift the lock without changing the host).

```python
class _Transport:
    @classmethod
    def get(cls) -> "_Transport":
        if cls._instance is None or cls._instance._proc.poll() is not None:
            cls._instance = cls._spawn()
            atexit.register(cls._instance.shutdown)
        return cls._instance

    def call(self, op: str, args: dict) -> dict:
        with self._lock:
            rid = self._next_id; self._next_id += 1
            self._proc.stdin.write(json.dumps({"id": rid, "op": op, "args": args}).encode() + b"\n")
            self._proc.stdin.flush()
            line = self._proc.stdout.readline()
            if not line:
                raise DocxodusTransportError("host exited unexpectedly")
            reply = json.loads(line)
            if reply["id"] != rid:
                raise DocxodusTransportError(f"id mismatch: expected {rid}, got {reply['id']}")
            if not reply["ok"]:
                raise DocxodusTransportError(reply["error"]["message"])
            return reply["result"]
```

## Lifecycle

```python
with open_session(docx_bytes) as session:        # documented contract
    proj = session.project()
    for placeholder in session.find_placeholders():
        session.replace_match(placeholder.match, "filled value")
    new_bytes = session.save()
# session.close() called automatically; handle returned to pool
```

`__del__` falls back to a best-effort `close_session` for forgotten sessions, never relied on for correctness.

## Type mapping (C# ↔ Python)

Every Python value type is `@dataclass(frozen=True, slots=True)`. Wire keys remain camelCase (same as WASM/TypeScript); decoders map to snake_case fields during deserialization.

| C# | Python | Wire shape |
|---|---|---|
| `Anchor` record | `Anchor(id, kind, scope, unid)` | `{id, kind, scope, unid}` |
| `CharSpan` | `CharSpan(start, length)` | `{start, length}` |
| `Position` enum | `Position(Enum)`: `BEFORE="before"`, `AFTER="after"` | `"before"`/`"after"` |
| `FormatOp` record | `FormatOp(bold=None, italic=None, underline=None, strike=None, code=None, color=None, run_style=None)` | nullable fields omitted when None |
| `EditErrorCode` (22) | `EditErrorCode(Enum)`, values match `EnumToSnake` output | e.g. `"anchor_not_found"` |
| `EditError` | `EditError(code, message, anchor_id=None)` | `{code, message, anchorId?}` |
| `EditResult` | `EditResult(success, created=(), removed=(), modified=(), patch=None, error=None)` | tuples, not lists |
| `MarkdownPatch` | `MarkdownPatch(scope_anchor_id, markdown)` | `{scopeAnchorId, markdown}` |
| `AnchorTarget` | `AnchorTarget(id, kind, scope, unid, part_uri, text_preview)` | flat (see `AppendAnchorTarget`) |
| `AnchorInfo` | `AnchorInfo(id, kind, scope, text_preview)` | `{id, kind, scope, textPreview}` |
| `TextMatch` | `TextMatch(text, enclosing_anchor, span, fragments, context_before, context_after, groups)` | |
| `RunFragment` | `RunFragment(unid, text, span_in_element, formatting)` | |
| `RunFormatting` | `RunFormatting(bold, italic, underline, strike, code, color=None, hyperlink_url=None, run_style=None)` | |
| `CrossBlockMatch` / `BlockSlice` | mirror exactly | |
| `TemplatePlaceholder` / `PlaceholderKind` | `(kind: PlaceholderKind, match: TextMatch, hint=None)` | |
| `PlaceholderKinds` / `ProjectionScopes` ([Flags]) | `IntFlag` | int on the wire |
| `RegexOptions` (.NET subset: `IgnoreCase=1`, `Multiline=2`, `Singleline=16`) | `IntFlag` | int |
| `TrackedChangeMode` | `Enum`: `ACCEPT`, `RENDER_INLINE`, `STRIP_DELETIONS` | snake_case strings |
| `MarkdownProjection` | `MarkdownProjection(markdown, anchor_index: dict[str, AnchorTarget])` | object map keyed by id |
| `DocxSessionSettings` | `DocxSessionSettings(undo_depth=50, validate_raw_ops=False, tracked_changes=TrackedChangeMode.ACCEPT, revision_author=None, persist_anchor_ids=False, smart_quotes=False)` | nested `projection_settings` deferred to v0.2 |
| `DocumentAnnotation` | mirror; `created` stays ISO-8601 `str` in v1 | |
| `AnnotationUpdate` | `AnnotationUpdate(label_id=None, label=None, color=None, author=None, metadata_patch=None)` — all fields optional; `metadata_patch` is `dict[str, str \| None] \| None` | `{labelId?, label?, color?, author?, metadataPatch?}` |

Decoders live in `types.py` — one pair per type, ~5 lines each. The Python wrapper methods are 3-line wrappers: build args dict → `transport.call(op, args)` → decode → return.

The TypeScript types in `npm/src/types.ts` (lines 700-970) already mirror the C# surface — use them as the source of truth for `types.py` shapes.

## Client-side helpers (not wire ops)

A few methods exist as syntactic sugar over already-exposed wire ops; the Python wrapper implements them client-side rather than adding wire ops:

- `session.replace_match(match: TextMatch, replace: str) -> EditResult` →
  `session.replace_text_at_span(match.enclosing_anchor.id, match.span.start, match.span.length, replace)`

  Avoids parsing a full `TextMatch` on the C# side (~80 lines of parser per transport).

- `session.fill_placeholders(picker, options?) -> BulkEditResult` — multi-pass
  template-fill loop mirroring the C# `DocxSession.FillPlaceholders` and the
  TypeScript `session.fillPlaceholders`. Bundles the three foot-guns every
  template-fill agent re-implements: reverse-offset ordering within a paragraph,
  `$`-prefix preservation (`$[___]` → `$0.20` instead of `0.20`), and multi-pass
  iteration for nested AlternativeClause brackets. Runs entirely in Python over
  the existing `find_placeholders` + `replace_match` primitives, matching the TS
  wrapper's no-extra-wire-op design.

## Annotation write surface

Four wire ops added to the Tier E surface (`add_annotation`, `remove_annotation`,
`update_annotation`, `move_annotation`). Python wrapper signatures:

| Method | Signature |
|--------|-----------|
| `add_annotation` | `add_annotation(anchor_id: str, span: CharSpan \| None, annotation: DocumentAnnotation) → EditResult` |
| `remove_annotation` | `remove_annotation(annotation_id: str) → EditResult` |
| `update_annotation` | `update_annotation(annotation_id: str, update: AnnotationUpdate) → EditResult` |
| `move_annotation` | `move_annotation(annotation_id: str, new_anchor_id: str, new_span: CharSpan \| None) → EditResult` |

All four return the standard `EditResult` envelope; `EditResult.annotation_id`
carries the affected id on success. New `EditErrorCode` values:
`DUPLICATE_ANNOTATION_ID`, `ANNOTATION_NOT_FOUND`, `EMPTY_ANNOTATION_SPAN`.
See `docs/architecture/docx_mutation_api.md` § "Tier E: Annotations" for the
full semantics.

## DOCX→HTML conversion

Two public functions expose HTML rendering on the Python surface.

### Module-level function (stateless)

```python
from docx_scalpel import convert_docx_to_html, HtmlOptions

html: str = convert_docx_to_html(docx_bytes, HtmlOptions(page_title="My Doc"))
```

`convert_docx_to_html(data: bytes, options: HtmlOptions | None = None) -> str`
renders the supplied bytes directly in the host process — no persistent session
is created, so it is the right choice when you only need HTML and do not intend
to edit the document. Maps to the `convert_to_html` stdio-host op.

### Session method (renders current edited state)

```python
with open_session(docx_bytes) as session:
    session.replace_text(anchor_id, "updated text")
    html = session.to_html(HtmlOptions(render_tracked_changes=True))
```

`DocxSession.to_html(options: HtmlOptions | None = None) -> str`
renders the session's current, possibly-edited in-memory state. If the session
is in `TrackedChangeMode.RENDER_INLINE`, tracked insertions and deletions appear
as `<ins>`/`<del>` in the HTML output when `render_tracked_changes=True`. Maps
to the `session_to_html` stdio-host op.

### `HtmlOptions` dataclass

`@dataclass(frozen=True, slots=True)` with the following fields and defaults:

| Field | Default | Notes |
|-------|---------|-------|
| `page_title` | `"Document"` | Value of the HTML `<title>` element. |
| `css_class_prefix` | `"docx-"` | Prefix applied to all generated CSS class names. |
| `fabricate_css_classes` | `True` | Emit a `<style>` block of generated classes in the output. |
| `additional_css` | `""` | Extra CSS injected into the `<style>` block verbatim. |
| `comment_render_mode` | `-1` | `-1` = disabled; `0` = endnote-style; `1` = inline; `2` = margin. |
| `comment_css_class_prefix` | `"comment-"` | CSS prefix for comment elements. |
| `pagination_mode` | `0` | `0` = none; `1` = paginated (emulates page breaks). |
| `pagination_scale` | `1.0` | Scale factor applied in paginated mode. |
| `pagination_css_class_prefix` | `"page-"` | CSS prefix for page-break elements. |
| `render_annotations` | `False` | Overlay `ExternalAnnotationProjector` annotation spans. |
| `annotation_label_mode` | `0` | `0` = above; `1` = inline; `2` = tooltip; `3` = none. |
| `annotation_css_class_prefix` | `"annot-"` | CSS prefix for annotation spans. |
| `render_footnotes_and_endnotes` | `False` | Append a footnotes/endnotes section to the HTML. |
| `render_headers_and_footers` | `False` | Include document headers and footers in the output. |
| `render_tracked_changes` | `False` | Render tracked insertions/deletions as `<ins>`/`<del>`. |
| `show_deleted_content` | `True` | When rendering tracked changes, include deleted content. |
| `render_move_operations` | `True` | Distinguish move operations from plain insert/delete. |
| `render_unsupported_content_placeholders` | `False` | Insert placeholder text for unsupported OOXML elements. |
| `document_language` | `None` | BCP-47 language tag (e.g. `"en-US"`). Omitted from the wire when `None`. |

Integer-coded mode fields follow the same conventions as the WASM/npm surface.
`HtmlOptions.to_wire()` always serializes every field except `document_language`,
which is omitted when `None`; the host's `ParseHtmlOptions` then applies the same
per-field defaults for any key that is absent.

### Shared core renderer

Both `convert_docx_to_html` and `DocxSession.to_html` are backed by
`Docxodus.Internal.HtmlConversionOps` — a single `internal` class that
provides:

- `ConvertToHtml(byte[] docxBytes, HtmlConversionOptions options) → string`
- `ConvertToHtml(DocxSession session, HtmlConversionOptions options) → string`
- `ConvertToHtml(int handle, HtmlConversionOptions options) → string`

The WASM `DocumentConverter.ConvertDocxToHtmlComplete` delegates to the same
renderer, ensuring that .NET, WASM/browser, and Python/stdio produce
byte-identical HTML for the same input and options. There is no
Python-specific rendering path.

## Wire-only internals

Decode helpers (`Anchor._from_wire(d)`, `EditResult._from_wire(d)`, etc.) live on
each value type but are deliberately leading-underscore: they are the JSON-to-dataclass
adapter for the stdio NDJSON transport and never need to be called by package
users. Every public `DocxSession` method that returns one of these types decodes
the response itself.

## Distribution

**Per-platform wheels with a bundled self-contained host binary. The user never installs .NET.**

- `dotnet publish tools/python-host/pyhost.csproj -c Release -r <rid> --self-contained -p:PublishSingleFile=true -p:PublishReadyToRun=true -o python/vendor/<rid>` per RID.
- `cibuildwheel` matrix wires the matching `vendor/<rid>/docxodus-pyhost[.exe]` into the wheel's `docxodus/_bin/` directory.
- ~25 MB compressed per RID — unremarkable next to ML wheels.
- The .NET 8 runtime is bundled inside the single-file binary; on first launch it transparently extracts to a temp dir. Pure-pip-installable; zero system deps.

RIDs to ship in v1: `linux-x64`, `linux-arm64`, `osx-x64`, `osx-arm64`, `win-x64`.

`DOCXODUS_HOST` env var overrides the bundled binary path — useful for developers iterating on the host without a full wheel rebuild.

## Testing strategy

Cherry-pick from `Docxodus.Tests/`:

- **`tests/test_smoke.py`** — mirror `Docxodus.Tests/DocxSessionSmokeTest.cs` (`DS999_AgenticWorkflowOnRealDocument`) line-for-line. Exercises every Tier in ~150 lines: Project, ReplaceText, InsertParagraph, SplitParagraph, ApplyFormat, Raw.GetXml/ReplaceXml, Save/reopen round-trip, full Undo chain. This is the v1 acceptance gate.
- **`tests/test_grep.py`** — one Grep + one GrepCrossBlock case; verify `RunFragment.formatting.bold` round-trips correctly through the wire.
- **`tests/test_format.py`** — `apply_format_by_substring` happy path + the offset-trap regression case it exists to prevent.
- **`tests/test_round_trip.py`** — open → mutate → save → reopen → project; markers preserved.
- **`tests/test_errors.py`** — trigger `anchor_not_found`, `malformed_markdown`. Verify Python sees `EditResult(success=False, error=EditError(code=EditErrorCode.ANCHOR_NOT_FOUND, ...))`. **Not** an exception.
- **`tests/test_lifecycle.py`** — open 50 sessions, mutate each, close. Verify host stays alive and `sessions` count returns to 0. Verify `with` block. Verify `atexit` shutdown by running pytest in a subprocess and confirming no orphan host process remains.
- **`tests/test_transport.py`** — mock the host to send malformed replies; assert `DocxodusTransportError`.

Tests import fixtures via `TEST_FILES = Path(__file__).parent.parent.parent / "TestFiles"` so Python tests use the *same* fixture corpus as the C# tests. Cross-language parity check: take the same fixture, run the same edit sequence through the WASM bridge (Playwright) and the Python wrapper, diff the saved bytes — they should be byte-identical. Any divergence is in the host loop or Python decoder, not document semantics.

## Explicitly deferred from v1

- **`async` API** — `asyncio.subprocess` facade is straightforward (one asyncio.Lock + request-id matching already on the wire), but doubles the test surface. v0.2.
- **Multi-process pooling** — one host per Python process is enough for the agentic editing pattern.
- **Native AOT publish** — `PublishReadyToRun` already eliminates JIT warmup; NAOT would require trim-safety auditing of every `PtOpenXmlUtil` reflection path. High risk for marginal payoff.
- **HTML conversion** — landed (`convert_docx_to_html` / `DocxSession.to_html`; see "DOCX→HTML conversion" above). **Comparison / chart extraction** (SkiaSharp-dependent) remain deferred — clean additive in v0.3 (new op names, same wire protocol).
- **Streaming `save`** — even 10 MB documents transit stdio in <100 ms.
- **Schema validation** — JSON Schema per op alongside the host binary so Python can validate before sending. v0.2.
- **Nested `ProjectionSettings`** — bridges don't expose them; defer until both transports gain support.

## Implementation sequencing (follow-up PR)

1. Scaffold `python/` (pyproject.toml, `_transport.py`, `_host_locator.py`, `errors.py`, `enums.py`, `types.py`, `session.py`). Implement 30+ wrapper methods using a `self._call(op, **args)` helper.
2. Write `tests/test_smoke.py` mirroring `DocxSessionSmokeTest.cs`. **Gating acceptance test.**
3. Per-feature tests (`test_grep`, `test_format`, `test_round_trip`, `test_errors`, `test_lifecycle`, `test_transport`).
4. `scripts/build_host.sh` publishing per-RID binaries into `python/vendor/<rid>/`.
5. Wheel packaging via `cibuildwheel` + GitHub Actions matrix.
6. README + docstrings.

## See also

- `docs/architecture/docx_mutation_api.md` — `DocxSession` surface, anchor lifecycle, error catalog, supported markdown subset.
- `docs/architecture/markdown_projection.md` — the projector that defines the anchor space the session operates on.
- `tools/python-host/Program.cs` — the host loop and exit semantics.
- `Docxodus/Internal/DocxSessionOps.cs` — the per-op facade; the canonical list of ops both bridges expose.

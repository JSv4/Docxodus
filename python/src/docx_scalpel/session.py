"""The public ``DocxSession`` class — Python wrapper over ``docxodus-pyhost``.

Every method mirrors a single ``DocxSessionOps`` C# entry point. The session
lives in the host's ``SessionRegistry`` until :meth:`DocxSession.close` is
called (or the host process exits). This is the load-bearing contract:
LLM-agent workflows issue dozens of small edits against one document and
must not pay the parse / Unid annotation / projection cost on every call.

Use the context-manager form whenever possible::

    with open_session(docx_bytes) as session:
        for placeholder in session.find_placeholders():
            session.replace_match(placeholder.match, "filled value")
        new_bytes = session.save()

A best-effort ``__del__`` finalizer closes a forgotten session, but the
context manager is the documented path — finalizers fire late and may
not run at all during interpreter shutdown.
"""

from __future__ import annotations

import base64
from typing import Any, Callable, Iterable, Mapping

from ._transport import call as _call
from .enums import (
    ContextBoundary,
    DiffFormat,
    PlaceholderKinds,
    Position,
    ProjectionDepth,
    ProjectionScopes,
    RegexOptions,
    WhitespaceMode,
)
from .types import (
    AnchorInfo,
    AnchorTarget,
    AnnotationUpdate,
    BlockMetadata,
    BulkEditResult,
    CharSpan,
    CrossBlockMatch,
    DocumentAnnotation,
    DocxSessionSettings,
    EditError,
    EditResult,
    EditSummary,
    FillOptions,
    FindOptions,
    FormatOp,
    HtmlOptions,
    ListMembership,
    MarkdownProjection,
    ReplaceOptions,
    SectionInfo,
    TemplatePlaceholder,
    TextMatch,
)

__all__ = ["DocxSession", "open_session", "ping", "convert_docx_to_html"]


def ping() -> dict[str, Any]:
    """Round-trip the host. Returns ``{pong, version, dotnet, sessions}``."""
    return _call("ping")


def open_session(
    docx_bytes: bytes, settings: DocxSessionSettings | None = None
) -> "DocxSession":
    """Open a new ``DocxSession`` over the given DOCX bytes.

    The bytes are sent base64-encoded over the NDJSON channel. The host parses
    the OOXML, annotates Unids, builds the initial projection, and returns an
    integer handle. The session lives until you call :meth:`DocxSession.close`,
    exit the ``with`` block, or the host exits.
    """
    args: dict[str, Any] = {"docxB64": base64.b64encode(docx_bytes).decode("ascii")}
    if settings is not None:
        args["settings"] = settings.to_wire()
    handle = _call("open_session", args)
    if not isinstance(handle, int):
        raise TypeError(f"open_session: expected int handle, got {type(handle).__name__}: {handle!r}")
    return DocxSession(handle)


def convert_docx_to_html(
    data: bytes, options: HtmlOptions | None = None
) -> str:
    """Convert DOCX bytes to a self-contained HTML string (stateless).

    Mirrors the WASM/npm ``convertDocxToHtml``. Renders the bytes directly in
    the host without creating a persistent session — use
    :meth:`DocxSession.to_html` to render the current state of an already-open
    editing session instead.
    """
    args: dict[str, Any] = {"docxB64": base64.b64encode(data).decode("ascii")}
    if options is not None:
        args["options"] = options.to_wire()
    html = _call("convert_to_html", args)
    if not isinstance(html, str):
        raise TypeError(f"convert_to_html: expected str, got {type(html).__name__}")
    return html


class DocxSession:
    """Handle to one open document session inside ``docxodus-pyhost``.

    Construct via :func:`open_session`; never instantiate directly.
    """

    __slots__ = ("_handle", "_closed")

    def __init__(self, handle: int) -> None:
        self._handle = handle
        self._closed = False

    # -- lifecycle --------------------------------------------------------

    @property
    def handle(self) -> int:
        """Integer handle in the host's ``SessionRegistry``. Stable for the session lifetime."""
        return self._handle

    @property
    def is_closed(self) -> bool:
        return self._closed

    def close(self) -> None:
        """Release the session in the host's ``SessionRegistry``. Idempotent."""
        if self._closed:
            return
        self._closed = True
        try:
            _call("close_session", {"handle": self._handle})
        except Exception:  # noqa: BLE001 — close must never raise out of user code
            pass

    def __enter__(self) -> "DocxSession":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def __del__(self) -> None:
        # Best-effort finalizer for forgotten sessions. Never relied on.
        if not getattr(self, "_closed", True):
            try:
                self.close()
            except Exception:  # noqa: BLE001
                pass

    def __repr__(self) -> str:
        state = "closed" if self._closed else "open"
        return f"DocxSession(handle={self._handle}, {state})"

    # -- core IO ----------------------------------------------------------

    def save(self) -> bytes:
        """Serialize the (mutated) document back to DOCX bytes. Does not close the session."""
        result = self._call("save", {})
        return base64.b64decode(result["docxB64"])

    def to_html(self, options: HtmlOptions | None = None) -> str:
        """Render this session's current (possibly edited) state to HTML."""
        args: dict[str, Any] = {}
        if options is not None:
            args["options"] = options.to_wire()
        html = self._call("session_to_html", args)
        if not isinstance(html, str):
            raise TypeError(f"session_to_html: expected str, got {type(html).__name__}")
        return html

    def undo(self) -> bool:
        """Undo one snapshot. Returns ``True`` if the undo ring had something to pop."""
        return bool(self._call("undo", {}))

    def redo(self) -> bool:
        """Redo one snapshot. Returns ``True`` if the redo ring had something to pop."""
        return bool(self._call("redo", {}))

    # -- projection -------------------------------------------------------

    def project(self) -> MarkdownProjection:
        """Full-document anchor-addressed markdown projection."""
        return MarkdownProjection._from_wire(self._call("project", {}))

    def project_anchor(
        self,
        anchor_id: str,
        depth: ProjectionDepth = ProjectionDepth.SUBTREE_AND_FOLLOWING_SIBLINGS,
    ) -> MarkdownProjection:
        """Scoped re-projection rooted at ``anchor_id``."""
        return MarkdownProjection._from_wire(
            self._call(
                "project_anchor",
                {"anchorId": anchor_id, "depth": int(depth)},
            )
        )

    # -- discovery: grep + find -------------------------------------------

    def grep(
        self,
        pattern: str,
        regex_options: RegexOptions = RegexOptions.NONE,
        scope: ProjectionScopes = ProjectionScopes.BODY,
        context_chars: int = 80,
        whitespace: WhitespaceMode = WhitespaceMode.PRESERVE,
        boundary: ContextBoundary = ContextBoundary.CHAR,
    ) -> tuple[TextMatch, ...]:
        result = self._call(
            "grep",
            {
                "pattern": pattern,
                "regexOptions": int(regex_options),
                "scope": int(scope),
                "contextChars": context_chars,
                "whitespace": int(whitespace),
                "boundary": int(boundary),
            },
        )
        return tuple(TextMatch._from_wire(m) for m in result)

    def grep_cross_block(
        self,
        pattern: str,
        regex_options: RegexOptions = RegexOptions.NONE,
        scope: ProjectionScopes = ProjectionScopes.BODY,
        context_chars: int = 80,
        whitespace: WhitespaceMode = WhitespaceMode.PRESERVE,
        boundary: ContextBoundary = ContextBoundary.CHAR,
    ) -> tuple[CrossBlockMatch, ...]:
        result = self._call(
            "grep_cross_block",
            {
                "pattern": pattern,
                "regexOptions": int(regex_options),
                "scope": int(scope),
                "contextChars": context_chars,
                "whitespace": int(whitespace),
                "boundary": int(boundary),
            },
        )
        return tuple(CrossBlockMatch._from_wire(m) for m in result)

    def find_placeholders(
        self,
        kinds: PlaceholderKinds = PlaceholderKinds.ALL,
        scope: ProjectionScopes = ProjectionScopes.BODY,
        context_chars: int = 80,
        boundary: ContextBoundary = ContextBoundary.CHAR,
    ) -> tuple[TemplatePlaceholder, ...]:
        result = self._call(
            "find_placeholders",
            {
                "kinds": int(kinds),
                "scope": int(scope),
                "contextChars": context_chars,
                "boundary": int(boundary),
            },
        )
        return tuple(TemplatePlaceholder._from_wire(p) for p in result)

    def remaining_placeholders(
        self, kinds: PlaceholderKinds = PlaceholderKinds.ALL
    ) -> tuple[TemplatePlaceholder, ...]:
        result = self._call("remaining_placeholders", {"kinds": int(kinds)})
        return tuple(TemplatePlaceholder._from_wire(p) for p in result)

    def fill_placeholders(
        self,
        picker: Callable[[TemplatePlaceholder], str | None],
        options: FillOptions | None = None,
    ) -> BulkEditResult:
        """Picker-driven template fill — Python mirror of C# ``DocxSession.FillPlaceholders``.

        For every placeholder matching ``options.kinds``, invokes ``picker`` and,
        if the picker returns a non-``None`` string, replaces the placeholder
        (with optional ``$``-prefix preservation per ``options.preserve_dollar_prefix``).
        Iterates until no more placeholders match (or until ``options.max_passes``
        is reached, or a pass makes zero state changes).

        Bundles the three foot-guns every template-fill agent re-implements:

        - **Reverse-offset ordering** across matches within the same paragraph so
          earlier-offset spans stay valid after later edits land.
        - **``$``-prefix preservation** — when a match starts with ``$`` and the
          picker's return value doesn't, the ``$`` is prepended (so ``$[___]`` →
          ``$0.20`` instead of ``0.20``). Disable via ``preserve_dollar_prefix=False``.
        - **Multi-pass convergence** — ``find_placeholders`` returns innermost
          brackets only; stripping one layer can surface a previously-nested outer
          layer. The loop iterates up to ``max_passes`` (default 8) until a pass
          makes no changes.

        The loop runs entirely in Python (no new wire op) — same primitives the
        TypeScript wrapper uses. The picker may be invoked more than once for the
        same logical placeholder when ``options.kinds`` includes
        :attr:`PlaceholderKinds.ALTERNATIVE_CLAUSE` and inner brackets are
        stripped between passes; pickers must therefore be deterministic on
        ``p.match.text`` (return the same result for the same input text).
        Non-deterministic pickers can produce inconsistent fills.

        Returns a :class:`BulkEditResult` with ``filled`` / ``skipped`` / ``passes``
        counts plus ``unfilled`` (placeholders the picker said ``None`` to,
        deduplicated across passes) and ``errors`` (per-replacement failures).
        Raises ``ValueError`` if ``options.max_passes <= 0`` — matches the
        ``ArgumentOutOfRangeException`` the C# API throws.

        See ``docs/architecture/docx_mutation_api.md#fillplaceholders``.
        """
        opts = options or FillOptions()
        if opts.max_passes <= 0:
            raise ValueError("FillOptions.max_passes must be > 0")

        filled = 0
        work_passes = 0
        errors: list[EditError] = []
        unfilled: list[TemplatePlaceholder] = []
        seen_skip_keys: set[tuple[str, int, int]] = set()

        for pass_num in range(1, opts.max_passes + 1):
            placeholders = sorted(
                self.find_placeholders(
                    opts.kinds, opts.scope, opts.context_chars, opts.boundary
                ),
                key=lambda p: (p.match.enclosing_anchor.id, p.match.span.start),
                reverse=True,
            )
            if not placeholders:
                break

            pass_changes = 0
            for p in placeholders:
                pick = picker(p)
                if pick is None:
                    key = (p.match.enclosing_anchor.id, p.match.span.start, p.match.span.length)
                    if key not in seen_skip_keys:
                        seen_skip_keys.add(key)
                        unfilled.append(p)
                    continue

                replacement = pick
                if (
                    opts.preserve_dollar_prefix
                    and p.match.text.startswith("$")
                    and not replacement.startswith("$")
                ):
                    replacement = "$" + replacement

                r = self.replace_match(p.match, replacement)
                if r.success:
                    filled += 1
                    pass_changes += 1
                elif r.error is not None:
                    errors.append(r.error)

            if pass_changes > 0:
                work_passes = pass_num
            if pass_changes == 0:
                break

        # Recompute post-loop so callers can assert `still_present == 0` as the
        # single-call "is the template done?" check (mirrors C# field added in #191).
        still_present = len(
            self.find_placeholders(opts.kinds, opts.scope, opts.context_chars, opts.boundary)
        )

        return BulkEditResult(
            filled=filled,
            skipped=len(unfilled),
            passes=work_passes,
            still_present=still_present,
            unfilled=tuple(unfilled),
            errors=tuple(errors),
        )

    def find_by_text(self, needle: str, options: FindOptions | None = None) -> AnchorTarget | None:
        args: dict[str, Any] = {"needle": needle}
        if options is not None:
            args["options"] = options.to_wire()
        result = self._call("find_by_text", args)
        return AnchorTarget._from_wire(result) if result else None

    def find_all_by_text(
        self, needle: str, options: FindOptions | None = None
    ) -> tuple[AnchorTarget, ...]:
        args: dict[str, Any] = {"needle": needle}
        if options is not None:
            args["options"] = options.to_wire()
        result = self._call("find_all_by_text", args)
        return tuple(AnchorTarget._from_wire(a) for a in result)

    def find_by_regex(
        self,
        pattern: str,
        regex_options: RegexOptions = RegexOptions.NONE,
        options: FindOptions | None = None,
    ) -> tuple[AnchorTarget, ...]:
        args: dict[str, Any] = {"pattern": pattern, "regexOptions": int(regex_options)}
        if options is not None:
            args["options"] = options.to_wire()
        result = self._call("find_by_regex", args)
        return tuple(AnchorTarget._from_wire(a) for a in result)

    def find_by_kind(self, kind: str, scope: str | None = None) -> tuple[AnchorTarget, ...]:
        args: dict[str, Any] = {"kind": kind}
        if scope is not None:
            args["scope"] = scope
        result = self._call("find_by_kind", args)
        return tuple(AnchorTarget._from_wire(a) for a in result)

    def find_by_annotation(self, annotation_id: str) -> tuple[AnchorTarget, ...]:
        result = self._call("find_by_annotation", {"annotationId": annotation_id})
        return tuple(AnchorTarget._from_wire(a) for a in result)

    def find_by_label(self, label_id: str) -> Mapping[str, tuple[AnchorTarget, ...]]:
        result = self._call("find_by_label", {"labelId": label_id})
        return {
            ann_id: tuple(AnchorTarget._from_wire(a) for a in anchors)
            for ann_id, anchors in result.items()
        }

    def find_by_bookmark(self, bookmark_name: str) -> tuple[AnchorTarget, ...]:
        result = self._call("find_by_bookmark", {"bookmarkName": bookmark_name})
        return tuple(AnchorTarget._from_wire(a) for a in result)

    def list_annotations(self) -> tuple[DocumentAnnotation, ...]:
        result = self._call("list_annotations", {})
        return tuple(DocumentAnnotation._from_wire(a) for a in result)

    # -- Tier E: annotations (write surface) -------------------------------

    def add_annotation(
        self,
        anchor_id: str,
        span: CharSpan | None,
        annotation: DocumentAnnotation,
    ) -> EditResult:
        """Annotate a range inside ``anchor_id``.

        When ``span`` is ``None`` the annotation wraps every inline run of
        the block. When ``annotation.id`` is empty, a 16-char hex id is
        auto-generated; check ``EditResult.annotation_id`` for the id used.
        """
        args: dict[str, Any] = {
            "anchorId": anchor_id,
            "annotation": annotation.to_wire(),
        }
        if span is not None:
            args["span"] = {"start": span.start, "length": span.length}
        return EditResult._from_wire(self._call("add_annotation", args))

    def remove_annotation(self, annotation_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call("remove_annotation", {"annotationId": annotation_id})
        )

    def update_annotation(
        self,
        annotation_id: str,
        update: AnnotationUpdate,
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "update_annotation",
                {"annotationId": annotation_id, "update": update.to_wire()},
            )
        )

    def move_annotation(
        self,
        annotation_id: str,
        new_anchor_id: str,
        new_span: CharSpan | None,
    ) -> EditResult:
        args: dict[str, Any] = {
            "annotationId": annotation_id,
            "newAnchorId": new_anchor_id,
        }
        if new_span is not None:
            args["newSpan"] = {"start": new_span.start, "length": new_span.length}
        return EditResult._from_wire(self._call("move_annotation", args))

    # -- discovery: anchor existence + info -------------------------------

    def exists(self, anchor_id: str) -> bool:
        return bool(self._call("exists", {"anchorId": anchor_id}))

    def get_anchor_info(self, anchor_id: str) -> AnchorInfo | None:
        result = self._call("get_anchor_info", {"anchorId": anchor_id})
        return AnchorInfo._from_wire(result) if result else None

    def get_anchor_infos(self, anchor_ids: Iterable[str]) -> dict[str, AnchorInfo | None]:
        result = self._call(
            "get_anchor_infos", {"anchorIds": list(anchor_ids)}
        )
        return {
            aid: AnchorInfo._from_wire(info) if info else None
            for aid, info in result.items()
        }

    def get_block_metadata(self, anchor_id: str) -> BlockMetadata | None:
        """Resolve block-level metadata (style id+name, outline level, list
        membership, formatting probe) for an anchor. Returns None for unknown anchors."""
        result = self._call("get_block_metadata", {"anchorId": anchor_id})
        return BlockMetadata._from_wire(result) if result else None

    def get_block_metadatas(
        self, anchor_ids: Iterable[str]
    ) -> dict[str, BlockMetadata | None]:
        """Bulk variant of :meth:`get_block_metadata`."""
        result = self._call("get_block_metadatas", {"anchorIds": list(anchor_ids)})
        return {
            aid: BlockMetadata._from_wire(meta) if meta else None
            for aid, meta in result.items()
        }

    def get_list_membership(self, anchor_id: str) -> ListMembership | None:
        """Resolve the numbering facts for a list-item paragraph. Returns None
        when the anchor has no w:numPr."""
        result = self._call("get_list_membership", {"anchorId": anchor_id})
        return ListMembership._from_wire(result) if result else None

    def get_section_info(self, anchor_id: str) -> SectionInfo | None:
        """Resolve page-layout info for the w:sectPr that governs an anchor.
        Returns None for anchors outside the body part."""
        result = self._call("get_section_info", {"anchorId": anchor_id})
        return SectionInfo._from_wire(result) if result else None

    # -- discovery: summaries ---------------------------------------------

    def get_edit_summary(self) -> EditSummary:
        return EditSummary._from_wire(self._call("get_edit_summary", {}))

    def get_diff(self, format: DiffFormat = DiffFormat.JSON) -> str:
        return str(self._call("get_diff", {"format": int(format)}))

    # -- Tier A: text mutations -------------------------------------------

    def replace_text(self, anchor_id: str, markdown: str) -> EditResult:
        return EditResult._from_wire(
            self._call("replace_text", {"anchorId": anchor_id, "markdown": markdown})
        )

    def replace_text_at_span(
        self, anchor_id: str, span_start: int, span_length: int, replace: str
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "replace_text_at_span",
                {
                    "anchorId": anchor_id,
                    "spanStart": span_start,
                    "spanLength": span_length,
                    "replace": replace,
                },
            )
        )

    def replace_text_range(
        self,
        anchor_id: str,
        find: str,
        replace: str,
        options: ReplaceOptions | None = None,
    ) -> tuple[EditResult, ...]:
        args: dict[str, Any] = {"anchorId": anchor_id, "find": find, "replace": replace}
        if options is not None:
            args["options"] = options.to_wire()
        result = self._call("replace_text_range", args)
        return tuple(EditResult._from_wire(r) for r in result)

    def replace_inner(
        self,
        match_text: str,
        anchor_id: str,
        span_start: int,
        span_length: int,
        new_inner: str,
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "replace_inner",
                {
                    "matchText": match_text,
                    "anchorId": anchor_id,
                    "spanStart": span_start,
                    "spanLength": span_length,
                    "newInner": new_inner,
                },
            )
        )

    def replace_match(self, match: TextMatch, replace: str) -> EditResult:
        """Client-side sugar over :meth:`replace_text_at_span` keyed by a prior search hit.

        Sends no wire op beyond what ``replace_text_at_span`` already sends —
        the .NET side does not need to re-parse the full ``TextMatch``.
        """
        return self.replace_text_at_span(
            match.enclosing_anchor.id, match.span.start, match.span.length, replace
        )

    def delete_block(self, anchor_id: str) -> EditResult:
        return EditResult._from_wire(self._call("delete_block", {"anchorId": anchor_id}))

    def delete_range(self, from_anchor_id: str, to_anchor_id_exclusive: str) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "delete_range",
                {
                    "fromAnchorId": from_anchor_id,
                    "toAnchorIdExclusive": to_anchor_id_exclusive,
                },
            )
        )

    def delete_section(self, heading_anchor_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call("delete_section", {"headingAnchorId": heading_anchor_id})
        )

    # -- Tier B: structural ------------------------------------------------

    def insert_paragraph(
        self, anchor_id: str, position: Position, markdown: str
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "insert_paragraph",
                {"anchorId": anchor_id, "position": position.value, "markdown": markdown},
            )
        )

    def split_paragraph(self, anchor_id: str, character_offset: int) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "split_paragraph",
                {"anchorId": anchor_id, "characterOffset": character_offset},
            )
        )

    def merge_paragraphs(self, first_anchor_id: str, second_anchor_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "merge_paragraphs",
                {"firstAnchorId": first_anchor_id, "secondAnchorId": second_anchor_id},
            )
        )

    # -- Tier C: formatting -----------------------------------------------

    def apply_format(
        self, anchor_id: str, span: CharSpan | None, op: FormatOp
    ) -> EditResult:
        args: dict[str, Any] = {"anchorId": anchor_id, "op": op.to_wire()}
        if span is not None:
            args["span"] = span.to_wire()
        return EditResult._from_wire(self._call("apply_format", args))

    def apply_format_by_substring(
        self, anchor_id: str, substring: str, op: FormatOp
    ) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "apply_format_by_substring",
                {"anchorId": anchor_id, "substring": substring, "op": op.to_wire()},
            )
        )

    def set_paragraph_style(self, anchor_id: str, style_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "set_paragraph_style",
                {"anchorId": anchor_id, "styleId": style_id},
            )
        )

    def set_list_level(self, anchor_id: str, level_delta: int) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "set_list_level",
                {"anchorId": anchor_id, "levelDelta": level_delta},
            )
        )

    def remove_list_membership(self, anchor_id: str) -> EditResult:
        return EditResult._from_wire(
            self._call("remove_list_membership", {"anchorId": anchor_id})
        )

    # -- Tier D: tables ---------------------------------------------------

    def replace_cell_content(self, cell_anchor_id: str, markdown: str) -> EditResult:
        return EditResult._from_wire(
            self._call(
                "replace_cell_content",
                {"cellAnchorId": cell_anchor_id, "markdown": markdown},
            )
        )

    # -- raw XML escape hatch ---------------------------------------------

    @property
    def raw(self) -> "_RawOps":
        """Sub-proxy exposing the raw-XML escape hatch (``get_xml``/``insert_xml``/``replace_xml``)."""
        return _RawOps(self)

    # -- internals --------------------------------------------------------

    def _call(self, op: str, args: dict[str, Any]) -> Any:
        if self._closed:
            raise ValueError(f"session {self._handle} is closed")
        payload = {"handle": self._handle, **args}
        return _call(op, payload)


class _RawOps:
    """Raw-XML escape hatch bound to a ``DocxSession``."""

    __slots__ = ("_s",)

    def __init__(self, session: DocxSession) -> None:
        self._s = session

    def get_xml(self, anchor_id: str) -> str:
        return str(self._s._call("raw_get_xml", {"anchorId": anchor_id}))

    def insert_xml(self, anchor_id: str, position: Position, xml: str) -> EditResult:
        return EditResult._from_wire(
            self._s._call(
                "raw_insert_xml",
                {"anchorId": anchor_id, "position": position.value, "xml": xml},
            )
        )

    def replace_xml(self, anchor_id: str, xml: str) -> EditResult:
        return EditResult._from_wire(
            self._s._call("raw_replace_xml", {"anchorId": anchor_id, "xml": xml})
        )

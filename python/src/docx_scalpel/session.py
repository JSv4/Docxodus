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
from typing import TYPE_CHECKING, Any, Callable, Iterable, Mapping

if TYPE_CHECKING:
    from types import TracebackType

from ._transport import call as _call
from .enums import (
    ContextBoundary,
    DiffFormat,
    HeaderFooterKind,
    PageNumberField,
    PlaceholderKinds,
    Position,
    ProjectionDepth,
    ProjectionScopes,
    RegexOptions,
    TrackedChangeMode,
    WhitespaceMode,
)
from .types import (
    AnchorInfo,
    AnchorTarget,
    AnnotationUpdate,
    BlockMetadata,
    BulkEditResult,
    CharSpan,
    CommentListEntry,
    CrossBlockMatch,
    DocumentAnnotation,
    DocxDiffConflict,
    DocxDiffConsolidatedRevision,
    DocxDiffConsolidateSettings,
    DocxDiffReviewer,
    DocxDiffRevision,
    DocxDiffSettings,
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
    NumberFormat,
    ParagraphFormatOp,
    ReplaceOptions,
    SectionInfo,
    TemplatePlaceholder,
    TextMatch,
)

__all__ = [
    "DocxSession",
    "open_session",
    "ping",
    "convert_docx_to_html",
    "docx_diff_compare",
    "docx_diff_get_revisions",
    "docx_diff_get_edit_script",
    "docx_diff_accept_revisions",
    "docx_diff_reject_revisions",
    "docx_diff_consolidate",
    "docx_diff_get_conflicts",
    "docx_diff_get_consolidated_revisions",
    "docx_diff_get_consolidated_edit_script",
]


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


# ---------------------------------------------------------------------------
# DocxDiff — IR diff engine (stateless two-document compare)
# ---------------------------------------------------------------------------
#
# These mirror the .NET ``DocxDiff`` static facade and the WASM/npm
# ``docxDiffCompare`` / ``docxDiffGetRevisions`` / ``docxDiffGetEditScript``
# wrappers. ``WmlComparer`` (not exposed in this wrapper yet) remains Docxodus'
# default comparison engine; ``DocxDiff`` is the NEW engine whose differentiators
# are anchor-addressed revisions and the diff-as-data edit script. All three are
# stateless: pass two DOCX byte blobs, get the result — no session.


def _diff_args(left: bytes, right: bytes, settings: DocxDiffSettings | None) -> dict[str, Any]:
    args: dict[str, Any] = {
        "leftB64": base64.b64encode(left).decode("ascii"),
        "rightB64": base64.b64encode(right).decode("ascii"),
    }
    if settings is not None:
        wire = settings.to_wire()
        if wire:
            args["settings"] = wire
    return args


def docx_diff_compare(
    left: bytes, right: bytes, settings: DocxDiffSettings | None = None
) -> bytes:
    """Compare two DOCX blobs; return a redlined DOCX (native tracked-changes markup).

    Mirrors .NET ``DocxDiff.Compare``. The result satisfies the WmlComparer
    contract: accepting its revisions yields ``right``, rejecting them yields
    ``left`` (at the per-block text level).
    """
    result = _call("docx_diff_compare", _diff_args(left, right, settings))
    if not isinstance(result, dict) or "docxB64" not in result:
        raise TypeError(f"docx_diff_compare: expected {{docxB64}}, got {result!r}")
    return base64.b64decode(result["docxB64"])


def docx_diff_get_revisions(
    left: bytes, right: bytes, settings: DocxDiffSettings | None = None
) -> tuple[DocxDiffRevision, ...]:
    """Compare two DOCX blobs; return the anchor-addressed revision list.

    Mirrors .NET ``DocxDiff.GetRevisions`` — the diff-as-revisions view, where
    each revision additionally carries the left/right block anchors it derives
    from.
    """
    result = _call("docx_diff_get_revisions", _diff_args(left, right, settings))
    revisions = result.get("revisions", []) if isinstance(result, dict) else []
    return tuple(DocxDiffRevision._from_wire(r) for r in revisions)


def docx_diff_get_edit_script(
    left: bytes, right: bytes, settings: DocxDiffSettings | None = None
) -> str:
    """Compare two DOCX blobs; return the engine's edit script as a JSON string.

    Mirrors .NET ``DocxDiff.GetEditScriptJson`` — the diff-as-data
    differentiator: the anchor-addressed block-operation list as machine-readable
    JSON, suitable for storage, transport, and review tooling.
    """
    result = _call("docx_diff_get_edit_script", _diff_args(left, right, settings))
    if not isinstance(result, str):
        raise TypeError(
            f"docx_diff_get_edit_script: expected str, got {type(result).__name__}"
        )
    return result


def docx_diff_accept_revisions(redline: bytes) -> bytes:
    """Accept every tracked revision in a redlined DOCX; return the resulting bytes.

    Materializes the "right"/revised side: ``docx_diff_accept_revisions(
    docx_diff_compare(left, right))`` equals ``right`` at the per-block text level.
    Mirrors .NET ``RevisionProcessor.AcceptRevisions`` via ``DocxDiffOps``. The
    byte-in, byte-out counterpart of :func:`docx_diff_compare` — together they let
    a caller verify the round-trip contract of the redline, not just its shape.
    """
    result = _call("docx_diff_accept_revisions", {"docxB64": base64.b64encode(redline).decode("ascii")})
    if not isinstance(result, dict) or "docxB64" not in result:
        raise TypeError(f"docx_diff_accept_revisions: expected {{docxB64}}, got {result!r}")
    return base64.b64decode(result["docxB64"])


def docx_diff_reject_revisions(redline: bytes) -> bytes:
    """Reject every tracked revision in a redlined DOCX; return the resulting bytes.

    Materializes the "left"/original side: ``docx_diff_reject_revisions(
    docx_diff_compare(left, right))`` equals ``left`` at the per-block text level.
    Mirrors .NET ``RevisionProcessor.RejectRevisions`` via ``DocxDiffOps``.
    """
    result = _call("docx_diff_reject_revisions", {"docxB64": base64.b64encode(redline).decode("ascii")})
    if not isinstance(result, dict) or "docxB64" not in result:
        raise TypeError(f"docx_diff_reject_revisions: expected {{docxB64}}, got {result!r}")
    return base64.b64decode(result["docxB64"])


# ---------------------------------------------------------------------------
# DocxDiff consolidate — multi-reviewer composite diff
# ---------------------------------------------------------------------------
#
# These mirror the .NET ``DocxDiff`` consolidate overloads and the WASM/npm
# ``docxDiffConsolidate`` / ``docxDiffGetConflicts`` /
# ``docxDiffGetConsolidatedRevisions`` / ``docxDiffGetConsolidatedEditScript``
# wrappers. All four are stateless: pass a base DOCX byte blob and a sequence
# of reviewer blobs, get the result — no session.


def _consolidate_args(
    base: bytes,
    reviewers: "Iterable[DocxDiffReviewer]",
    settings: "DocxDiffConsolidateSettings | None",
) -> dict[str, Any]:
    args: dict[str, Any] = {
        "baseB64": base64.b64encode(base).decode("ascii"),
        "reviewers": [
            {
                "author": r.author,
                "docB64": base64.b64encode(r.document).decode("ascii"),
            }
            for r in reviewers
        ],
    }
    if settings is not None:
        wire = settings.to_wire()
        if wire:
            args["settings"] = wire
    return args


def docx_diff_consolidate(
    base: bytes,
    reviewers: "Iterable[DocxDiffReviewer]",
    settings: "DocxDiffConsolidateSettings | None" = None,
) -> bytes:
    """Consolidate reviewer DOCX blobs onto a base; return a redlined DOCX.

    Mirrors .NET ``DocxDiff.Consolidate``. Diffs each reviewer's document
    against ``base``, merges the resulting edit scripts, resolves conflicts
    per ``settings.conflict_resolution``, and renders a single tracked-changes
    document. Accepting its revisions yields the consolidated result; rejecting
    them yields ``base``.
    """
    result = _call("docx_diff_consolidate", _consolidate_args(base, reviewers, settings))
    if not isinstance(result, dict) or "docxB64" not in result:
        raise TypeError(
            f"docx_diff_consolidate: expected {{docxB64}}, got {result!r}"
        )
    return base64.b64decode(result["docxB64"])


def docx_diff_get_conflicts(
    base: bytes,
    reviewers: "Iterable[DocxDiffReviewer]",
    settings: "DocxDiffConsolidateSettings | None" = None,
) -> tuple[DocxDiffConflict, ...]:
    """Consolidate reviewer blobs against ``base``; return the conflict list.

    Mirrors .NET ``DocxDiff.GetConflicts`` — the conflicts-as-data view: every
    base span where two or more reviewers made incompatible edits, with each
    reviewer's competing text and the resolution policy that would be applied
    under the current settings.
    """
    result = _call("docx_diff_get_conflicts", _consolidate_args(base, reviewers, settings))
    conflicts = result.get("conflicts", []) if isinstance(result, dict) else []
    return tuple(DocxDiffConflict._from_wire(c) for c in conflicts)


def docx_diff_get_consolidated_revisions(
    base: bytes,
    reviewers: "Iterable[DocxDiffReviewer]",
    settings: "DocxDiffConsolidateSettings | None" = None,
) -> tuple[DocxDiffConsolidatedRevision, ...]:
    """Consolidate reviewer blobs against ``base``; return the merged revision list.

    Mirrors .NET ``DocxDiff.GetConsolidatedRevisions`` — the revisions-as-data
    view of the consolidated edit script, where each revision additionally carries
    the conflict id (if any) it derives from.
    """
    result = _call(
        "docx_diff_get_consolidated_revisions",
        _consolidate_args(base, reviewers, settings),
    )
    revisions = result.get("revisions", []) if isinstance(result, dict) else []
    return tuple(DocxDiffConsolidatedRevision._from_wire(r) for r in revisions)


def docx_diff_get_consolidated_edit_script(
    base: bytes,
    reviewers: "Iterable[DocxDiffReviewer]",
    settings: "DocxDiffConsolidateSettings | None" = None,
) -> str:
    """Consolidate reviewer blobs against ``base``; return the edit script as JSON.

    Mirrors .NET ``DocxDiff.GetConsolidatedEditScriptJson`` — the merged
    anchor-addressed block-operation list as machine-readable JSON, suitable for
    storage, transport, and review tooling.
    """
    result = _call(
        "docx_diff_get_consolidated_edit_script",
        _consolidate_args(base, reviewers, settings),
    )
    if not isinstance(result, str):
        raise TypeError(
            f"docx_diff_get_consolidated_edit_script: expected str, "
            f"got {type(result).__name__}"
        )
    return result


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

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
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

    def save(self, persist_anchor_ids: bool | None = None) -> bytes:
        """Serialize the (mutated) document back to DOCX bytes. Does not close the session.

        ``persist_anchor_ids`` overrides the session's open-time
        :attr:`DocxSessionSettings.persist_anchor_ids` for this save only:
        ``True`` keeps the anchor-id bookkeeping in the bytes so a session opened
        over them later resolves the same anchor ids (an anchor-stable checkpoint,
        at a file-size cost); ``False`` strips it (a clean deliverable from a
        session opened anchor-stable); ``None`` (default) uses the session's setting.
        """
        args: dict[str, Any] = {}
        if persist_anchor_ids is not None:
            args["persistAnchorIds"] = persist_anchor_ids
        result = self._call("save", args)
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

    def set_tracked_changes(self, mode: TrackedChangeMode) -> None:
        """Switch how subsequent mutations are recorded (issue #304).

        Session configuration, not a document mutation: not undoable, and
        already-applied markup is never touched — switching to ``ACCEPT`` does
        not resolve existing revisions, and switching to ``RENDER_INLINE`` does
        not retroactively track prior direct edits.
        """
        self._call("set_tracked_changes", {"mode": mode.value})

    def set_revision_author(self, author: str | None) -> None:
        """Author stamped on subsequent tracked-change markup; ``None`` restores the ``"docxodus"`` default."""
        self._call("set_revision_author", {"author": author})

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

    # -- Headers / footers / page numbers ---------------------------------

    def set_header_text(
        self, anchor_id: str, kind: HeaderFooterKind, markdown: str
    ) -> EditResult:
        """Set the running header story for the section that owns ``anchor_id``.

        ``anchor_id`` is any body block in the target section. Creates the header part +
        ``w:headerReference`` if the story of ``kind`` doesn't exist yet, else replaces its
        content. The created header-paragraph anchors (scope ``hdr{N}``) come back in
        ``EditResult.created`` — pass one to :meth:`insert_page_number_field`.
        """
        return EditResult._from_wire(
            self._call(
                "set_header_text",
                {"anchorId": anchor_id, "kind": kind.value, "markdown": markdown},
            )
        )

    def set_footer_text(
        self, anchor_id: str, kind: HeaderFooterKind, markdown: str
    ) -> EditResult:
        """Set the running footer story for the section that owns ``anchor_id`` — see
        :meth:`set_header_text`. Created footer-paragraph anchors have scope ``ftr{N}``."""
        return EditResult._from_wire(
            self._call(
                "set_footer_text",
                {"anchorId": anchor_id, "kind": kind.value, "markdown": markdown},
            )
        )

    def insert_page_number_field(
        self,
        anchor_id: str,
        field: PageNumberField = PageNumberField.CURRENT_PAGE,
        format: NumberFormat | None = None,
    ) -> EditResult:
        """Append a page-number field to the paragraph ``anchor_id`` (typically a header/footer
        paragraph). ``CURRENT_PAGE`` emits a ``PAGE`` field, ``TOTAL_PAGES`` a ``NUMPAGES`` field.
        Returns the affected paragraph anchor in ``EditResult.modified``.

        ``format`` writes the field's own ``\\*`` general-formatting switch (``PAGE \\* roman`` →
        ``i, ii, iii``). Omitting it — the default — emits a plain field, which is what Word inserts
        and what follows the SECTION's format (:meth:`set_page_numbering`). Prefer the section
        setting for ordinary page numbering: a switch here overrides it for this one field and keeps
        overriding it if the section later changes. ``NumberFormat.BULLET`` is rejected.
        """
        args: dict[str, Any] = {"anchorId": anchor_id, "field": field.value}
        if format is not None:
            args["format"] = format.value
        return EditResult._from_wire(self._call("insert_page_number_field", args))

    def set_page_numbering(
        self,
        anchor_id: str,
        start: int | None = None,
        format: NumberFormat | None = None,
    ) -> EditResult:
        """Set the page-numbering properties (``w:pgNumType``) of the section that owns
        ``anchor_id`` (any body block in that section) — Word's *Format Page Numbers…* dialog.

        ``start`` is the page number the section starts at (e.g. ``1`` to restart at a section
        break); ``format`` is the format its pages use (e.g. ``NumberFormat.LOWER_ROMAN`` for
        ``i, ii, iii`` front matter). Passing ``None`` for either leaves that attribute unchanged,
        so the start can be set without disturbing the format and vice versa. Creates the element,
        and a trailing ``w:sectPr``, if absent.

        Applying values the section already has is a successful no-op that does not consume undo
        history. ``NumberFormat.BULLET`` and a negative ``start`` are rejected with
        ``EditErrorCode.INVALID_PAGE_NUMBERING``.
        """
        op: dict[str, Any] = {}
        if start is not None:
            op["start"] = start
        if format is not None:
            op["format"] = format.value
        return EditResult._from_wire(
            self._call("set_page_numbering", {"anchorId": anchor_id, "op": op})
        )

    def clear_page_numbering(self, anchor_id: str) -> EditResult:
        """Remove the section's page-numbering start/format: it reverts to continuing the previous
        section's numbering in Word's default ``1, 2, 3``. Chapter-numbering attributes
        (``w:chapStyle``/``w:chapSep``) are preserved. A section with nothing to clear is a no-op."""
        return EditResult._from_wire(
            self._call("clear_page_numbering", {"anchorId": anchor_id})
        )

    def ensure_header_footer_visible(
        self, anchor_id: str, kind: HeaderFooterKind
    ) -> EditResult:
        """Make the ``kind`` header/footer stories of the section owning ``anchor_id`` actually
        render: ``FIRST`` sets ``w:titlePg``, ``EVEN`` the document-global ``w:evenAndOddHeaders``;
        ``DEFAULT`` needs no flag and is a successful no-op. Idempotent.

        ``set_header_text``/``set_footer_text`` set these flags while writing content, which covers
        authoring a story from scratch — but not a document that already carries a first/even
        reference with the flag absent (Word leaves exactly that behind when "Different first page"
        is switched off), where the story's content would be present but invisible.
        """
        return EditResult._from_wire(
            self._call(
                "ensure_header_footer_visible",
                {"anchorId": anchor_id, "kind": kind.value},
            )
        )

    # -- Footnotes / endnotes ---------------------------------------------

    def insert_footnote(
        self, anchor_id: str, character_offset: int, markdown: str
    ) -> EditResult:
        """Create a footnote with body ``markdown`` and cite it from the body paragraph
        ``anchor_id``, at ``character_offset`` characters into that paragraph's text (0 = before
        all text, text length = after all of it).

        On a document with no footnotes yet this also creates the footnotes part, Word's two
        reserved separator notes, the ``FootnoteText``/``FootnoteReference`` styles and the
        ``w:footnotePr`` settings declaration; otherwise the existing part is reused. The note id
        is allocated above every id already used in the package, so non-contiguous ids can't
        collide.

        The created note anchors come back in ``EditResult.created`` — the definition (kind
        ``fn``) and its paragraphs (kind ``p``, scope ``fn``) — so the note can immediately be
        edited with :meth:`replace_text` or removed with :meth:`delete_block` (which also drops
        the body-side reference).

        Body paragraphs only: Word does not allow a note reference inside a header/footer story or
        inside another note, so a non-body anchor fails with ``ANCHOR_WRONG_KIND``.
        """
        return EditResult._from_wire(
            self._call(
                "insert_footnote",
                {
                    "anchorId": anchor_id,
                    "characterOffset": character_offset,
                    "markdown": markdown,
                },
            )
        )

    def insert_endnote(
        self, anchor_id: str, character_offset: int, markdown: str
    ) -> EditResult:
        """Create an endnote — see :meth:`insert_footnote`; writes the endnotes part and a
        ``w:endnoteReference``, and the created definition anchor has kind ``en``."""
        return EditResult._from_wire(
            self._call(
                "insert_endnote",
                {
                    "anchorId": anchor_id,
                    "characterOffset": character_offset,
                    "markdown": markdown,
                },
            )
        )

    # -- Comments (issue #300) --------------------------------------------

    def add_comment(
        self,
        anchor_id: str,
        span: CharSpan | None,
        author: str,
        markdown: str,
        initials: str | None = None,
        date: str | None = None,
    ) -> EditResult:
        """Add a **native Word comment** (real ``w:comment`` markup, visible in Word's
        Reviewing pane — not the :meth:`add_annotation` overlay) on the body paragraph
        ``anchor_id``. ``span`` selects the commented character range; ``None`` comments
        the whole block. On a document with no comments yet this also creates the
        comments part and the ``CommentText``/``CommentReference`` styles. ``date``
        (ISO-8601) is written only when provided, keeping output deterministic by
        default.

        The created definition anchor (kind ``cmt``) and its paragraph anchors (kind
        ``p``, scope ``cmt``) come back in ``EditResult.created``, so the comment can
        immediately be edited with :meth:`update_comment` or removed with
        :meth:`remove_comment`.

        Body paragraphs only (Word has no comments-on-comments); a non-body anchor
        fails with ``ANCHOR_WRONG_KIND``, a zero-length span with ``EMPTY_COMMENT_SPAN``.
        """
        args: dict[str, Any] = {
            "anchorId": anchor_id,
            "author": author,
            "markdown": markdown,
        }
        if span is not None:
            args["span"] = {"start": span.start, "length": span.length}
        if initials is not None:
            args["initials"] = initials
        if date is not None:
            args["date"] = date
        return EditResult._from_wire(self._call("add_comment", args))

    def update_comment(self, comment_anchor_id: str, markdown: str) -> EditResult:
        """Replace a comment's body text, addressed by its definition anchor (kind
        ``cmt``); the comment's author/initials/date are preserved, as is the last
        paragraph's ``w14:paraId`` (Word's reply-threading key)."""
        return EditResult._from_wire(
            self._call(
                "update_comment",
                {"anchorId": comment_anchor_id, "markdown": markdown},
            )
        )

    def remove_comment(self, comment_anchor_id: str) -> EditResult:
        """Remove a comment: the definition, its body marker triple everywhere in the
        package, and any ``commentsExtended``/``commentsIds`` threading entries keyed
        by it."""
        return EditResult._from_wire(
            self._call("remove_comment", {"anchorId": comment_anchor_id})
        )

    def list_comments(self) -> tuple[CommentListEntry, ...]:
        """The document's native Word comments in comments-part order."""
        result = self._call("list_comments", {})
        return tuple(CommentListEntry._from_wire(c) for c in result)

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

    def set_paragraph_format(
        self, anchor_id: str, op: ParagraphFormatOp
    ) -> EditResult:
        """Apply paragraph-level formatting to the paragraph at ``anchor_id``: alignment
        (``w:jc``), left-indent delta in twips (``w:ind/@w:left``, clamped at 0), page-break-
        before (``w:pageBreakBefore``), and top/bottom paragraph borders (``w:pBdr``).

        Every field on ``op`` is tri-state: ``None`` leaves that property unchanged. Setting
        ``op.top_border``/``op.bottom_border`` adds or replaces that border edge;
        ``op.clear_borders=True`` removes the whole ``w:pBdr`` before either is applied in the
        same call. An (often empty) paragraph with only a bottom border is what an S-1-style
        horizontal rule is.
        """
        return EditResult._from_wire(
            self._call(
                "set_paragraph_format",
                {"anchorId": anchor_id, "op": op.to_wire()},
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

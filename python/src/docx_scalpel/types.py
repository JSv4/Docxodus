"""Frozen dataclasses mirroring the C# value types in ``Docxodus``.

Wire keys are camelCase (matching the WASM bridge so the JSON shapes are
interchangeable between TypeScript and Python clients). Decoders translate to
snake_case Python fields.

Each type exposes a private ``_from_wire(d)`` classmethod that takes the parsed
JSON dict and returns an instance. The leading underscore marks these as
transport-internal: callers should never need to invoke them — every public
``DocxSession`` method that returns one of these types already decodes for you.
Encoders are simple dict-builders on the encode side (``to_wire()``) and live
where they're used in ``session.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Mapping, Sequence

from .enums import (
    AnchorIdRendering,
    AnchorRenderMode,
    ContextBoundary,
    EditErrorCode,
    EmptyParagraphMode,
    PlaceholderKind,
    PlaceholderKinds,
    ProjectionScopes,
    TableRenderMode,
    TrackedChangeMode,
    WhitespaceMode,
)

__all__ = [
    "Anchor",
    "CharSpan",
    "FormatOp",
    "EditError",
    "EditResult",
    "MarkdownPatch",
    "AnchorTarget",
    "AnchorInfo",
    "BlockMetadata",
    "BulkEditResult",
    "FillOptions",
    "FindOptions",
    "HtmlOptions",
    "ListMembership",
    "NumberFormat",
    "RunFormatting",
    "RunFragment",
    "SectionInfo",
    "TextMatch",
    "BlockSlice",
    "CrossBlockMatch",
    "TemplatePlaceholder",
    "MarkdownProjection",
    "DocxSessionSettings",
    "WmlToMarkdownConverterSettings",
    "DocumentAnnotation",
    "AnnotationUpdate",
    "EditSummary",
    "ReplaceOptions",
]


# ---------------------------------------------------------------------------
# Anchors
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class Anchor:
    """A block-level address into a Docxodus session.

    The ``id`` is the canonical wire form (e.g. ``"p:body:abcd1234..."``) and
    is what every mutation op consumes. ``kind`` / ``scope`` / ``unid`` are the
    decomposed parts.
    """

    id: str
    kind: str
    scope: str
    unid: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "Anchor":
        return cls(id=d["id"], kind=d["kind"], scope=d["scope"], unid=d["unid"])


@dataclass(frozen=True, slots=True)
class AnchorTarget:
    """Search-result anchor with extra metadata (``partUri``, ``textPreview``)."""

    id: str
    kind: str
    scope: str
    unid: str
    part_uri: str
    text_preview: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "AnchorTarget":
        return cls(
            id=d["id"],
            kind=d["kind"],
            scope=d["scope"],
            unid=d["unid"],
            part_uri=d.get("partUri", ""),
            text_preview=d.get("textPreview", ""),
        )


@dataclass(frozen=True, slots=True)
class AnchorInfo:
    """Minimal anchor metadata returned by ``get_anchor_info`` / ``get_anchor_infos``."""

    id: str
    kind: str
    scope: str
    text_preview: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "AnchorInfo":
        return cls(
            id=d["id"],
            kind=d["kind"],
            scope=d["scope"],
            text_preview=d.get("textPreview", ""),
        )


class NumberFormat(str, Enum):
    """Six list formats supported by the list write surface and surfaced
    on ``ListMembership.format``. String-valued so the wire JSON round-trips
    transparently."""

    DECIMAL = "decimal"
    UPPER_LETTER = "upperLetter"
    LOWER_LETTER = "lowerLetter"
    UPPER_ROMAN = "upperRoman"
    LOWER_ROMAN = "lowerRoman"
    BULLET = "bullet"

    @classmethod
    def _from_wire(cls, raw: str) -> "NumberFormat":
        try:
            return cls(raw)
        except ValueError:
            return cls.DECIMAL


@dataclass(frozen=True, slots=True)
class ListMembership:
    """Numbering facts for a list-item paragraph."""

    num_id: int
    abstract_num_id: int
    level: int
    format: NumberFormat
    is_auto_numbered: bool
    from_style: bool
    start_override: int | None = None
    generated_label: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "ListMembership":
        return cls(
            num_id=int(d["numId"]),
            abstract_num_id=int(d["abstractNumId"]),
            level=int(d["level"]),
            format=NumberFormat._from_wire(d["format"]),
            is_auto_numbered=bool(d["isAutoNumbered"]),
            from_style=bool(d["fromStyle"]),
            start_override=int(d["startOverride"]) if "startOverride" in d else None,
            generated_label=d.get("generatedLabel"),
        )


@dataclass(frozen=True, slots=True)
class BlockMetadata:
    """Block-level structural metadata."""

    anchor_id: str
    kind: str
    scope: str
    has_inline_formatting: bool
    style_id: str | None = None
    style_name: str | None = None
    outline_level: int | None = None
    list: ListMembership | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BlockMetadata":
        return cls(
            anchor_id=d["anchorId"],
            kind=d["kind"],
            scope=d["scope"],
            has_inline_formatting=bool(d["hasInlineFormatting"]),
            style_id=d.get("styleId"),
            style_name=d.get("styleName"),
            outline_level=int(d["outlineLevel"]) if "outlineLevel" in d else None,
            list=ListMembership._from_wire(d["list"]) if "list" in d else None,
        )


@dataclass(frozen=True, slots=True)
class SectionInfo:
    """Page-layout snapshot for the w:sectPr that governs an anchor."""

    section_unid: str
    page_width_twips: int
    page_height_twips: int
    landscape: bool
    margin_top_twips: int
    margin_bottom_twips: int
    margin_left_twips: int
    margin_right_twips: int
    columns: int
    header_part_uris: tuple[str, ...]
    footer_part_uris: tuple[str, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "SectionInfo":
        return cls(
            section_unid=d["sectionUnid"],
            page_width_twips=int(d["pageWidthTwips"]),
            page_height_twips=int(d["pageHeightTwips"]),
            landscape=bool(d["landscape"]),
            margin_top_twips=int(d["marginTopTwips"]),
            margin_bottom_twips=int(d["marginBottomTwips"]),
            margin_left_twips=int(d["marginLeftTwips"]),
            margin_right_twips=int(d["marginRightTwips"]),
            columns=int(d["columns"]),
            header_part_uris=tuple(d["headerPartUris"]),
            footer_part_uris=tuple(d["footerPartUris"]),
        )


# ---------------------------------------------------------------------------
# Spans + formatting
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class CharSpan:
    """Half-open character range within an anchor's plain-text projection."""

    start: int
    length: int

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CharSpan":
        return cls(start=int(d["start"]), length=int(d["length"]))

    def to_wire(self) -> dict[str, int]:
        return {"start": self.start, "length": self.length}


@dataclass(frozen=True, slots=True)
class FormatOp:
    """Set of formatting changes to apply.

    Each field is tri-state: ``True`` to turn on, ``False`` to turn off, ``None``
    to leave unchanged. Strings (``color``, ``run_style``) are passed through;
    ``None`` means "don't change", empty string means "clear".
    """

    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    code: bool | None = None
    color: str | None = None
    run_style: str | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.bold is not None: out["bold"] = self.bold
        if self.italic is not None: out["italic"] = self.italic
        if self.underline is not None: out["underline"] = self.underline
        if self.strike is not None: out["strike"] = self.strike
        if self.code is not None: out["code"] = self.code
        if self.color is not None: out["color"] = self.color
        if self.run_style is not None: out["runStyle"] = self.run_style
        return out


@dataclass(frozen=True, slots=True)
class RunFormatting:
    """Resolved run-level formatting for a ``RunFragment``."""

    bold: bool = False
    italic: bool = False
    underline: bool = False
    strike: bool = False
    code: bool = False
    color: str | None = None
    hyperlink_url: str | None = None
    run_style: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RunFormatting":
        return cls(
            bold=bool(d.get("bold", False)),
            italic=bool(d.get("italic", False)),
            underline=bool(d.get("underline", False)),
            strike=bool(d.get("strike", False)),
            code=bool(d.get("code", False)),
            color=d.get("color"),
            hyperlink_url=d.get("hyperlinkUrl"),
            run_style=d.get("runStyle"),
        )


@dataclass(frozen=True, slots=True)
class RunFragment:
    """A single contiguous run inside a ``TextMatch`` slice."""

    unid: str
    text: str
    span_in_element: CharSpan
    formatting: RunFormatting

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RunFragment":
        return cls(
            unid=d["unid"],
            text=d["text"],
            span_in_element=CharSpan._from_wire(d["spanInElement"]),
            formatting=RunFormatting._from_wire(d["formatting"]),
        )


# ---------------------------------------------------------------------------
# Search / Grep results
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TextMatch:
    """A single grep / find-by-text match."""

    text: str
    enclosing_anchor: Anchor
    span: CharSpan
    fragments: tuple[RunFragment, ...] = ()
    context_before: str = ""
    context_after: str = ""
    groups: tuple[str, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TextMatch":
        return cls(
            text=d["text"],
            enclosing_anchor=Anchor._from_wire(d["enclosingAnchor"]),
            span=CharSpan._from_wire(d["span"]),
            fragments=tuple(RunFragment._from_wire(f) for f in d.get("fragments", ())),
            context_before=d.get("contextBefore", ""),
            context_after=d.get("contextAfter", ""),
            groups=tuple(d.get("groups", ())),
        )


@dataclass(frozen=True, slots=True)
class BlockSlice:
    """One block's contribution to a ``CrossBlockMatch``."""

    anchor: Anchor
    span_in_block: CharSpan
    fragments: tuple[RunFragment, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "BlockSlice":
        return cls(
            anchor=Anchor._from_wire(d["anchor"]),
            span_in_block=CharSpan._from_wire(d["spanInBlock"]),
            fragments=tuple(RunFragment._from_wire(f) for f in d.get("fragments", ())),
        )


@dataclass(frozen=True, slots=True)
class CrossBlockMatch:
    """A grep match that spans multiple adjacent blocks."""

    text: str
    enclosing_anchors: tuple[Anchor, ...]
    slices: tuple[BlockSlice, ...]
    context_before: str = ""
    context_after: str = ""
    groups: tuple[str, ...] = ()

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CrossBlockMatch":
        return cls(
            text=d["text"],
            enclosing_anchors=tuple(Anchor._from_wire(a) for a in d.get("enclosingAnchors", ())),
            slices=tuple(BlockSlice._from_wire(s) for s in d.get("slices", ())),
            context_before=d.get("contextBefore", ""),
            context_after=d.get("contextAfter", ""),
            groups=tuple(d.get("groups", ())),
        )


# ---------------------------------------------------------------------------
# Placeholders
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class TemplatePlaceholder:
    """A classified bracketed region from ``find_placeholders``."""

    kind: PlaceholderKind
    match: TextMatch
    alternative_kinds: tuple[PlaceholderKind, ...] = ()
    hint: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "TemplatePlaceholder":
        return cls(
            kind=PlaceholderKind(d["kind"]),
            match=TextMatch._from_wire(d["match"]),
            alternative_kinds=tuple(
                PlaceholderKind(k) for k in d.get("alternativeKinds", ())
            ),
            hint=d.get("hint"),
        )


# ---------------------------------------------------------------------------
# Mutation results
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class EditError:
    """Business-level mutation failure inside a successful ``EditResult`` envelope."""

    code: EditErrorCode
    message: str
    anchor_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditError":
        return cls(
            code=EditErrorCode(d["code"]),
            message=d.get("message", ""),
            anchor_id=d.get("anchorId"),
        )


@dataclass(frozen=True, slots=True)
class MarkdownPatch:
    """A scoped markdown re-projection produced by a successful mutation."""

    scope_anchor_id: str
    markdown: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MarkdownPatch":
        return cls(
            scope_anchor_id=d["scopeAnchorId"],
            markdown=d.get("markdown", ""),
        )


@dataclass(frozen=True, slots=True)
class EditResult:
    """The typed envelope returned by every mutation op.

    ``success=False`` here is a *normal business outcome* (anchor not found,
    malformed markdown, etc.). Transport failures raise ``DocxodusTransportError``
    instead of returning an ``EditResult``.
    """

    success: bool
    created: tuple[Anchor, ...] = ()
    removed: tuple[Anchor, ...] = ()
    modified: tuple[Anchor, ...] = ()
    patch: MarkdownPatch | None = None
    error: EditError | None = None
    annotation_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditResult":
        patch_d = d.get("patch")
        err_d = d.get("error")
        return cls(
            success=bool(d.get("success", False)),
            created=tuple(Anchor._from_wire(a) for a in d.get("created", ())),
            removed=tuple(Anchor._from_wire(a) for a in d.get("removed", ())),
            modified=tuple(Anchor._from_wire(a) for a in d.get("modified", ())),
            patch=MarkdownPatch._from_wire(patch_d) if patch_d else None,
            error=EditError._from_wire(err_d) if err_d else None,
            annotation_id=d.get("annotationId"),
        )


# ---------------------------------------------------------------------------
# Projection
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class MarkdownProjection:
    """The markdown + anchor-index pair returned by ``project`` / ``project_anchor``."""

    markdown: str
    anchor_index: Mapping[str, AnchorTarget]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "MarkdownProjection":
        idx = d.get("anchorIndex", {}) or {}
        # The wire entries don't repeat the id key (it's the dict key); rebuild
        # the AnchorTarget by injecting the id from the surrounding key.
        decoded: dict[str, AnchorTarget] = {}
        for anchor_id, entry in idx.items():
            decoded[anchor_id] = AnchorTarget(
                id=anchor_id,
                kind=entry.get("kind", ""),
                scope=entry.get("scope", ""),
                unid=entry.get("unid", ""),
                part_uri=entry.get("partUri", ""),
                text_preview=entry.get("textPreview", ""),
            )
        return cls(markdown=d.get("markdown", ""), anchor_index=decoded)


# ---------------------------------------------------------------------------
# Annotations
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocumentAnnotation:
    """One stored ``DocumentAnnotation`` returned by ``list_annotations``."""

    id: str
    label_id: str
    label: str
    color: str
    bookmark_name: str
    author: str | None = None
    created: str | None = None  # ISO-8601, kept as a string in v1.
    annotated_text: str | None = None
    metadata: Mapping[str, str] = field(default_factory=dict)

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocumentAnnotation":
        return cls(
            id=d["id"],
            label_id=d.get("labelId", ""),
            label=d.get("label", ""),
            color=d.get("color", ""),
            bookmark_name=d.get("bookmarkName", ""),
            author=d.get("author"),
            created=d.get("created"),
            annotated_text=d.get("annotatedText"),
            metadata=dict(d.get("metadata", {}) or {}),
        )

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {
            "id": self.id,
            "labelId": self.label_id,
            "label": self.label,
            "color": self.color,
            "bookmarkName": self.bookmark_name,
        }
        if self.author is not None:
            wire["author"] = self.author
        if self.created is not None:
            wire["created"] = self.created
        if self.annotated_text is not None:
            wire["annotatedText"] = self.annotated_text
        if self.metadata:
            wire["metadata"] = dict(self.metadata)
        return wire


@dataclass(frozen=True, slots=True)
class AnnotationUpdate:
    """Partial-update payload for :meth:`DocxSession.update_annotation`.

    ``None`` / missing fields leave the existing value unchanged.
    ``metadata_patch`` is a per-key merge: a non-``None`` value sets the
    key, an explicit ``None`` removes it, a missing key leaves it
    unchanged.
    """

    label_id: str | None = None
    label: str | None = None
    color: str | None = None
    author: str | None = None
    metadata_patch: Mapping[str, str | None] | None = None

    def to_wire(self) -> dict[str, Any]:
        wire: dict[str, Any] = {}
        if self.label_id is not None:
            wire["labelId"] = self.label_id
        if self.label is not None:
            wire["label"] = self.label
        if self.color is not None:
            wire["color"] = self.color
        if self.author is not None:
            wire["author"] = self.author
        if self.metadata_patch is not None:
            # Preserve explicit None values — they mean "remove this key".
            wire["metadataPatch"] = dict(self.metadata_patch)
        return wire


# ---------------------------------------------------------------------------
# Edit summary
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class EditSummary:
    """Aggregate snapshot returned by ``get_edit_summary``.

    The shape of the embedded ``remaining_placeholders`` list mirrors
    ``find_placeholders``; ``bare_underscore_runs`` is a list of
    ``{anchor_id, run_unid, text}`` triples.
    """

    total_anchors: int
    remaining_placeholders: tuple[TemplatePlaceholder, ...] = ()
    bare_underscore_runs: tuple[Mapping[str, Any], ...] = ()
    footnote_count: int = 0
    inline_footnote_ref_count: int = 0
    comment_count: int = 0

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "EditSummary":
        return cls(
            total_anchors=int(d.get("totalAnchors", 0)),
            remaining_placeholders=tuple(
                TemplatePlaceholder._from_wire(p) for p in d.get("remainingPlaceholders", ())
            ),
            bare_underscore_runs=tuple(d.get("bareUnderscoreRuns", ())),
            footnote_count=int(d.get("footnoteCount", 0)),
            inline_footnote_ref_count=int(d.get("inlineFootnoteRefCount", 0)),
            comment_count=int(d.get("commentCount", 0)),
        )


# ---------------------------------------------------------------------------
# Option bundles for find / replace
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class FindOptions:
    """Optional filters for ``find_by_text`` / ``find_all_by_text`` / ``find_by_regex``."""

    ignore_case: bool = False
    ignore_whitespace: bool = False
    kind_filter: str | None = None
    scope_filter: ProjectionScopes | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.ignore_case: out["ignoreCase"] = True
        if self.ignore_whitespace: out["ignoreWhitespace"] = True
        if self.kind_filter is not None: out["kindFilter"] = self.kind_filter
        if self.scope_filter is not None: out["scopeFilter"] = int(self.scope_filter)
        return out


@dataclass(frozen=True, slots=True)
class ReplaceOptions:
    """Options for ``replace_text_range``."""

    ignore_case: bool = False
    max_replacements: int | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.ignore_case: out["ignoreCase"] = True
        if self.max_replacements is not None: out["maxReplacements"] = self.max_replacements
        return out


# ---------------------------------------------------------------------------
# DOCX → HTML conversion options (see convert_docx_to_html / DocxSession.to_html)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class HtmlOptions:
    """Options for DOCX→HTML conversion. Mirrors the .NET ``HtmlConversionOptions``.

    Integer-coded modes match the wire contract:
    ``comment_render_mode`` -1=disabled,0=endnote,1=inline,2=margin;
    ``pagination_mode`` 0=none,1=paginated;
    ``annotation_label_mode`` 0=above,1=inline,2=tooltip,3=none.
    """

    page_title: str = "Document"
    css_class_prefix: str = "docx-"
    fabricate_css_classes: bool = True
    additional_css: str = ""
    comment_render_mode: int = -1
    comment_css_class_prefix: str = "comment-"
    pagination_mode: int = 0
    pagination_scale: float = 1.0
    pagination_css_class_prefix: str = "page-"
    render_annotations: bool = False
    annotation_label_mode: int = 0
    annotation_css_class_prefix: str = "annot-"
    render_footnotes_and_endnotes: bool = False
    render_headers_and_footers: bool = False
    render_tracked_changes: bool = False
    show_deleted_content: bool = True
    render_move_operations: bool = True
    render_unsupported_content_placeholders: bool = False
    document_language: str | None = None

    def to_wire(self) -> dict[str, Any]:
        """camelCase keys the host dispatcher's ``ParseHtmlOptions`` reads."""
        wire: dict[str, Any] = {
            "pageTitle": self.page_title,
            "cssClassPrefix": self.css_class_prefix,
            "fabricateCssClasses": self.fabricate_css_classes,
            "additionalCss": self.additional_css,
            "commentRenderMode": self.comment_render_mode,
            "commentCssClassPrefix": self.comment_css_class_prefix,
            "paginationMode": self.pagination_mode,
            "paginationScale": self.pagination_scale,
            "paginationCssClassPrefix": self.pagination_css_class_prefix,
            "renderAnnotations": self.render_annotations,
            "annotationLabelMode": self.annotation_label_mode,
            "annotationCssClassPrefix": self.annotation_css_class_prefix,
            "renderFootnotesAndEndnotes": self.render_footnotes_and_endnotes,
            "renderHeadersAndFooters": self.render_headers_and_footers,
            "renderTrackedChanges": self.render_tracked_changes,
            "showDeletedContent": self.show_deleted_content,
            "renderMoveOperations": self.render_move_operations,
            "renderUnsupportedContentPlaceholders": self.render_unsupported_content_placeholders,
        }
        if self.document_language is not None:
            wire["documentLanguage"] = self.document_language
        return wire


# ---------------------------------------------------------------------------
# Fill placeholders (client-side multi-pass loop; see DocxSession.fill_placeholders)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class FillOptions:
    """Options for :meth:`DocxSession.fill_placeholders`.

    Mirrors the C# ``FillOptions`` record; the multi-pass loop runs client-side
    in Python (no new wire op), so these are consumed locally rather than
    serialized.
    """

    kinds: PlaceholderKinds = PlaceholderKinds.ALL
    scope: ProjectionScopes = ProjectionScopes.BODY
    max_passes: int = 8
    preserve_dollar_prefix: bool = True
    context_chars: int = 80
    boundary: ContextBoundary = ContextBoundary.CHAR


@dataclass(frozen=True, slots=True)
class BulkEditResult:
    """Aggregate result returned by :meth:`DocxSession.fill_placeholders`.

    ``filled`` is the count of picker-returned replacements applied;
    ``skipped`` counts placeholders the picker returned ``None`` for (deduped
    across passes — counted once per placeholder lifetime). ``passes`` is the
    highest iteration pass that actually filled something (``0`` = nothing
    filled, ``1`` = one-shot convergence, higher = multi-pass nested-bracket
    convergence). ``still_present`` is a post-loop ``find_placeholders`` count
    — the trustworthy "is the template done?" check (``skipped > 0 &&
    still_present == 0`` means "picker said no on first sight but later passes
    resolved it"; the canonical case from the NVCA Model COI). Mirrors the
    C# ``BulkEditResult.StillPresent`` added in #191. ``unfilled`` and
    ``errors`` mirror the C# shape.
    """

    filled: int
    skipped: int
    passes: int
    still_present: int
    unfilled: tuple["TemplatePlaceholder", ...] = ()
    errors: tuple["EditError", ...] = ()


# ---------------------------------------------------------------------------
# Session settings (nested projection settings are exposed but not yet
# round-tripped by the bridges per the design doc)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class WmlToMarkdownConverterSettings:
    """Mirror of the C# converter settings, included on ``DocxSessionSettings``."""

    scopes: ProjectionScopes = ProjectionScopes.ALL
    heading_level_offset: int = 0
    anchor_mode: AnchorRenderMode = AnchorRenderMode.BLOCK
    table_mode: TableRenderMode = TableRenderMode.GFM_WITH_OPAQUE_FALLBACK
    table_inline_cell_max: int = 80
    tracked_changes: TrackedChangeMode = TrackedChangeMode.ACCEPT
    resolve_numbering: bool = True
    empty_paragraphs: EmptyParagraphMode = EmptyParagraphMode.ANCHOR_ONLY
    anchor_id_rendering: AnchorIdRendering = AnchorIdRendering.FULL_UNID

    def to_wire(self) -> dict[str, Any]:
        return {
            "scopes": int(self.scopes),
            "headingLevelOffset": self.heading_level_offset,
            "anchorMode": int(self.anchor_mode),
            "tableMode": int(self.table_mode),
            "tableInlineCellMax": self.table_inline_cell_max,
            "trackedChanges": self.tracked_changes.value,
            "resolveNumbering": self.resolve_numbering,
            "emptyParagraphs": int(self.empty_paragraphs),
            "anchorIdRendering": int(self.anchor_id_rendering),
        }


@dataclass(frozen=True, slots=True)
class DocxSessionSettings:
    """Constructor settings passed to ``open_session(bytes, settings=...)``."""

    undo_depth: int = 50
    validate_raw_ops: bool = False
    tracked_changes: TrackedChangeMode = TrackedChangeMode.ACCEPT
    revision_author: str | None = None
    persist_anchor_ids: bool = False
    smart_quotes: bool = False
    capture_initial_projection: bool = True
    projection_settings: WmlToMarkdownConverterSettings | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {
            "undoDepth": self.undo_depth,
            "validateRawOps": self.validate_raw_ops,
            "trackedChanges": self.tracked_changes.value,
            "persistAnchorIds": self.persist_anchor_ids,
            "smartQuotes": self.smart_quotes,
            "captureInitialProjection": self.capture_initial_projection,
        }
        if self.revision_author is not None:
            out["revisionAuthor"] = self.revision_author
        if self.projection_settings is not None:
            out["projectionSettings"] = self.projection_settings.to_wire()
        return out

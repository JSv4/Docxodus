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
    ConflictResolution,
    ContextBoundary,
    DocxDiffFormatComparison,
    DocxDiffRevisionGranularity,
    DocxDiffRevisionType,
    EditErrorCode,
    EmptyParagraphMode,
    HeaderFooterKind,
    LineSpacingRule,
    ParagraphAlignment,
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
    "ParagraphBorderEdge",
    "ParagraphFormatOp",
    "EditError",
    "EditResult",
    "MarkdownPatch",
    "AnchorTarget",
    "AnchorInfo",
    "BlockMetadata",
    "BulkEditResult",
    "FillOptions",
    "FindOptions",
    "HeaderFooterRef",
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
    "DocxDiffSettings",
    "DocxDiffRevision",
    "DocxDiffFormatChange",
    "DocxDiffReviewer",
    "DocxDiffConsolidateSettings",
    "DocxDiffConflictCompetitor",
    "DocxDiffConflict",
    "DocxDiffConsolidatedRevision",
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
class HeaderFooterRef:
    """A ``w:headerReference``/``w:footerReference`` on a section.

    Carries the story kind the reference supplies and the URI of the part holding it,
    so a caller can map a :class:`HeaderFooterKind` to a part — and thence to that
    part's projection anchors, which carry the same ``part_uri`` — instead of guessing
    from part-collection order, which carries no kind information.
    """

    kind: HeaderFooterKind
    part_uri: str
    #: True when this section declares no reference of ``kind`` and the story is inherited from
    #: the nearest preceding section that does (ECMA-376 §17.6.17).
    inherited: bool = False

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "HeaderFooterRef":
        return cls(
            kind=HeaderFooterKind(d["kind"]),
            part_uri=d["partUri"],
            inherited=bool(d.get("inherited", False)),
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
    #: Header references in declaration order, each with its ``w:type``. Describes exactly
    #: the parts :attr:`header_part_uris` lists, plus the kind each supplies.
    header_refs: tuple[HeaderFooterRef, ...] = ()
    #: Footer references in declaration order, each with its ``w:type``.
    footer_refs: tuple[HeaderFooterRef, ...] = ()
    #: The page number this section starts at (``w:pgNumType/@w:start``). ``None`` when the
    #: section continues the previous section's numbering.
    page_number_start: int | None = None
    #: This section's page-number format (``w:pgNumType/@w:fmt``). ``None`` means Word's default
    #: ``1, 2, 3`` — deliberately not reported as ``DECIMAL``, so a caller can tell "inherits"
    #: from "explicitly decimal" and avoid writing an attribute the document never had.
    page_number_format: NumberFormat | None = None

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
            # .get: an older host that predates the refs still decodes.
            header_refs=tuple(HeaderFooterRef._from_wire(r) for r in d.get("headerRefs", ())),
            footer_refs=tuple(HeaderFooterRef._from_wire(r) for r in d.get("footerRefs", ())),
            # Omitted by the host when the attribute is absent — "inherit", not a default value.
            page_number_start=(
                int(d["pageNumberStart"]) if d.get("pageNumberStart") is not None else None
            ),
            page_number_format=(
                NumberFormat._from_wire(d["pageNumberFormat"])
                if d.get("pageNumberFormat") is not None
                else None
            ),
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
class ParagraphBorderEdge:
    """One edge of a paragraph border (a ``w:pBdr`` child — ``w:top``/``w:bottom``) for
    :attr:`ParagraphFormatOp.top_border`/:attr:`ParagraphFormatOp.bottom_border`.

    When an edge is set, fields left ``None`` fall back to the host's defaults (style
    "single", size 6 eighths-of-a-point ≈0.75pt, color "auto", space 1pt).
    """

    style: str | None = None
    size: int | None = None
    color: str | None = None
    space: int | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.style is not None: out["style"] = self.style
        if self.size is not None: out["size"] = self.size
        if self.color is not None: out["color"] = self.color
        if self.space is not None: out["space"] = self.space
        return out


@dataclass(frozen=True, slots=True)
class ParagraphFormatOp:
    """Paragraph-level formatting changes for ``DocxSession.set_paragraph_format``.

    Each field is tri-state: ``None`` leaves that property unchanged. ``clear_borders=True``
    removes the entire ``w:pBdr`` (all paragraph borders) before applying any
    ``top_border``/``bottom_border`` set in this same op — this is how an S-1-style horizontal
    rule paragraph gets (or loses) its bottom border.

    Indent/spacing fields are absolute values in twips (1440 twips = 1 inch; 20 twips = 1pt)
    and must be >= 0. ``first_line_indent`` (``w:ind/@w:firstLine``) and ``hanging_indent``
    (``w:ind/@w:hanging``) share one either/or slot in Word — setting one removes the other,
    and an op carrying both is rejected with ``INVALID_PARAGRAPH_FORMAT``. ``line_spacing``
    (``w:spacing/@w:line``) is measured per ``line_spacing_rule``: 240ths of a line under
    ``AUTO`` (the default when the rule is omitted — 240 = single, 360 = 1.5x, 480 = double),
    twips under ``EXACT``/``AT_LEAST``; a rule without ``line_spacing`` is rejected.
    """

    alignment: ParagraphAlignment | None = None
    indent_delta: int | None = None
    first_line_indent: int | None = None
    hanging_indent: int | None = None
    spacing_before: int | None = None
    spacing_after: int | None = None
    line_spacing: int | None = None
    line_spacing_rule: LineSpacingRule | None = None
    page_break_before: bool | None = None
    top_border: ParagraphBorderEdge | None = None
    bottom_border: ParagraphBorderEdge | None = None
    clear_borders: bool | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.alignment is not None: out["alignment"] = self.alignment.value
        if self.indent_delta is not None: out["indentDelta"] = self.indent_delta
        if self.first_line_indent is not None: out["firstLineIndent"] = self.first_line_indent
        if self.hanging_indent is not None: out["hangingIndent"] = self.hanging_indent
        if self.spacing_before is not None: out["spacingBefore"] = self.spacing_before
        if self.spacing_after is not None: out["spacingAfter"] = self.spacing_after
        if self.line_spacing is not None: out["lineSpacing"] = self.line_spacing
        if self.line_spacing_rule is not None: out["lineSpacingRule"] = self.line_spacing_rule.value
        if self.page_break_before is not None: out["pageBreakBefore"] = self.page_break_before
        if self.top_border is not None: out["topBorder"] = self.top_border.to_wire()
        if self.bottom_border is not None: out["bottomBorder"] = self.bottom_border.to_wire()
        if self.clear_borders is not None: out["clearBorders"] = self.clear_borders
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


@dataclass(frozen=True, slots=True)
class CommentListEntry:
    """One native Word comment, in comments-part order — see ``Session.list_comments``.

    ``anchor_id`` addresses the definition (kind ``cmt``) for ``update_comment``/
    ``remove_comment``; ``date`` is the raw ``w:date`` attribute string (``None`` when
    absent); ``text`` is the flattened body (paragraphs joined by a space, the
    ``w:annotationRef`` mark excluded). ``parent_anchor_id`` and ``resolved`` are
    populated only when the comment has a ``commentsExtended.xml`` entry; ``None``
    distinguishes a legacy/flat comment from an explicitly reopened one. The numeric
    ``w:id`` is deliberately not surfaced — comments are addressed by anchor everywhere.
    """

    anchor_id: str
    author: str
    initials: str | None = None
    date: str | None = None
    text: str = ""
    parent_anchor_id: str | None = None
    resolved: bool | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "CommentListEntry":
        return cls(
            anchor_id=d["anchorId"],
            author=d.get("author", "unknown"),
            initials=d.get("initials"),
            date=d.get("date"),
            text=d.get("text", ""),
            parent_anchor_id=d.get("parentAnchorId"),
            resolved=d.get("resolved"),
        )


@dataclass(frozen=True, slots=True)
class RevisionListEntry:
    """One tracked revision read directly off the live markup — see
    ``Session.list_revisions``.

    ``id`` is stable while the underlying markup exists (derived from the markup's
    own ``w:id`` attributes — resolving OTHER revisions never renames it) and is what
    ``accept_revision``/``reject_revision`` address. ``type`` is ``"insert"``,
    ``"delete"``, ``"move"`` (a linked move pair — both sides resolve together), or
    ``"format"``. ``author``/``date`` are the markup's true ``w:author``/``w:date``
    (``date`` ``None`` when absent). ``text`` is the revision's visible text (deleted
    text for deletions, ``¶`` for a revised paragraph mark, the affected text for
    format changes). ``anchor_id`` is the containing block's anchor when addressable.
    """

    id: str
    type: str
    author: str
    date: str | None = None
    text: str = ""
    anchor_id: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "RevisionListEntry":
        return cls(
            id=d["id"],
            type=d.get("type", ""),
            author=d.get("author", "unknown"),
            date=d.get("date"),
            text=d.get("text", ""),
            anchor_id=d.get("anchorId"),
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
    """Optional filters for ``find_by_text`` / ``find_all_by_text`` / ``find_by_regex``.

    Mirrors the .NET ``FindOptions`` record's two distinct scope controls:

    * ``scopes`` — a :class:`ProjectionScopes` flag set (coarse "which categories
      of part" filter; wire key ``scopes``). Compose with ``|`` to widen, e.g.
      ``ProjectionScopes.HEADERS | ProjectionScopes.FOOTERS``. Defaults to all
      scopes when unset.
    * ``scope_filter`` — a string naming one specific part such as ``"hdr1"``
      (wire key ``scopeFilter``), applied as a finer post-filter on top of
      ``scopes``. Prefer ``scopes`` for whole-category filtering.
    """

    ignore_case: bool = False
    ignore_whitespace: bool = False
    kind_filter: str | None = None
    scopes: ProjectionScopes | None = None
    scope_filter: str | None = None

    def to_wire(self) -> dict[str, Any]:
        out: dict[str, Any] = {}
        if self.ignore_case: out["ignoreCase"] = True
        if self.ignore_whitespace: out["ignoreWhitespace"] = True
        if self.kind_filter is not None: out["kindFilter"] = self.kind_filter
        if self.scopes is not None: out["scopes"] = int(self.scopes)
        if self.scope_filter is not None: out["scopeFilter"] = self.scope_filter
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
# DocxDiff — IR diff engine (stateless two-document compare)
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocxDiffSettings:
    """Settings for the ``docx_diff_*`` module functions. Mirrors the .NET
    ``DocxDiffSettings``; defaults match it field-for-field, so an unset
    ``DocxDiffSettings()`` reproduces the engine's out-of-the-box behavior.

    Every field is sent only when it differs from the default, so the wire
    object stays minimal and the host applies its own defaults for omitted keys.
    """

    author_for_revisions: str = "Open-Xml-PowerTools"
    deterministic: bool = True
    date_time_for_revisions: str | None = None
    #: Accept every tracked revision already present in both inputs before diffing (default False).
    #: ``preserve_input_revisions`` wins when both policies are enabled.
    pre_accept_input_revisions: bool = False
    #: Preserve the inputs' existing tracked revisions Word-style (default False).
    #: This takes precedence over ``pre_accept_input_revisions``.
    preserve_input_revisions: bool = False
    #: Collapse every tracked-revision author in the output to ``author_for_revisions`` (default False).
    #: Matches how author-coloring renderers (LibreOffice) display Word's single-author compare output
    #: when input revisions preserved under their original author would otherwise render in a second color.
    normalize_revision_authors: bool = False
    case_insensitive: bool = False
    culture: str | None = None
    conflate_breaking_and_nonbreaking_spaces: bool = True
    word_separators: str | None = None
    detect_moves: bool = True
    move_similarity_threshold: float = 0.8
    move_minimum_word_count: int = 3
    revision_granularity: DocxDiffRevisionGranularity = DocxDiffRevisionGranularity.FINE
    format_comparison: DocxDiffFormatComparison = DocxDiffFormatComparison.MODELED_ONLY
    #: Compare header/footer stories (default True — Word Compare's own default).
    #: Changed stories get native tracked-changes markup inside their parts; FINE
    #: revisions carry hdr/ftr-scoped anchors; the edit script carries
    #: ``headerFooterOps``. False ignores header/footer scopes entirely.
    compare_headers_footers: bool = True
    #: Track paragraph-and-above property changes (pPr/tcPr/trPr/tblPr/tblGrid/tblPrEx/sectPr) as native
    #: Word markup. Default True; False restores the untracked-right-apply behavior. Consolidate ignores
    #: block-format changes regardless.
    track_block_format_changes: bool = True
    #: Render a run of >=2 adjacent word-matched modified paragraph pairs via a single cross-paragraph
    #: word+pilcrow token-stream diff (the within-run flat-stream shape decoded from Word's compare
    #: output: retained words may cross the pilcrow, paragraph marks are ins/del stream tokens, and the
    #: output paragraph count follows the token-level interleave). Markup (``docx_diff_compare``) only —
    #: the revision list and edit-script JSON are unaffected. Default False.
    cross_paragraph_token_diff: bool = True

    def to_wire(self) -> dict[str, Any]:
        """camelCase keys the host's ``DocxDiffOps.ParseSettings`` reads. Only
        non-default fields are emitted."""
        wire: dict[str, Any] = {}
        if self.author_for_revisions != "Open-Xml-PowerTools":
            wire["authorForRevisions"] = self.author_for_revisions
        if not self.deterministic:
            wire["deterministic"] = False
        if self.date_time_for_revisions is not None:
            wire["dateTimeForRevisions"] = self.date_time_for_revisions
        if self.pre_accept_input_revisions:
            wire["preAcceptInputRevisions"] = True
        if self.preserve_input_revisions:
            wire["preserveInputRevisions"] = True
        if self.normalize_revision_authors:
            wire["normalizeRevisionAuthors"] = True
        if self.case_insensitive:
            wire["caseInsensitive"] = True
        if self.culture is not None:
            wire["culture"] = self.culture
        if not self.conflate_breaking_and_nonbreaking_spaces:
            wire["conflateBreakingAndNonbreakingSpaces"] = False
        if self.word_separators is not None:
            wire["wordSeparators"] = self.word_separators
        if not self.detect_moves:
            wire["detectMoves"] = False
        if self.move_similarity_threshold != 0.8:
            wire["moveSimilarityThreshold"] = self.move_similarity_threshold
        if self.move_minimum_word_count != 3:
            wire["moveMinimumWordCount"] = self.move_minimum_word_count
        if self.revision_granularity != DocxDiffRevisionGranularity.FINE:
            wire["revisionGranularity"] = int(self.revision_granularity)
        if self.format_comparison != DocxDiffFormatComparison.MODELED_ONLY:
            wire["formatComparison"] = int(self.format_comparison)
        if not self.compare_headers_footers:
            wire["compareHeadersFooters"] = False
        if not self.track_block_format_changes:
            wire["trackBlockFormatChanges"] = False
        if not self.cross_paragraph_token_diff:
            wire["crossParagraphTokenDiff"] = False
        return wire


@dataclass(frozen=True, slots=True)
class DocxDiffFormatChange:
    """Details of a ``FORMAT_CHANGED`` revision — the modeled format fields
    before/after plus the names that differ. Mirrors .NET ``DocxDiffFormatChange``.

    ``scope`` names the property container the change describes: ``"run"`` (the default,
    an rPr-grade report) or one of the block-format-change family scopes ``"paragraph"``
    (pPr), ``"tableCell"``/``"tableRow"``/``"table"`` (tcPr/trPr/tblPr+tblGrid), ``"section"``
    (sectPr). Non-run scopes are reported only under Fine revision granularity."""

    old_properties: Mapping[str, str]
    new_properties: Mapping[str, str]
    changed_property_names: Sequence[str]
    scope: str = "run"

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffFormatChange":
        return cls(
            old_properties=dict(d.get("oldProperties") or {}),
            new_properties=dict(d.get("newProperties") or {}),
            changed_property_names=tuple(d.get("changedPropertyNames") or ()),
            scope=str(d.get("scope") or "run"),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffRevision:
    """One consumer revision from ``docx_diff_get_revisions``. Mirrors .NET
    ``DocxDiffRevision`` and carries the IR engine's differentiator: the
    left/right block anchors the revision derives from.

    Anchor presence by ``type`` — each type's PRIMARY anchor is ALWAYS
    present; the opposite anchor MAY also be present for a token-level
    revision. ``INSERTED`` → ``right_anchor`` always (plus ``left_anchor``
    when it is a token-level insert inside a modified block); ``DELETED`` →
    ``left_anchor`` always (plus ``right_anchor`` when token-level);
    ``FORMAT_CHANGED`` → both; ``MOVED`` is EXCLUSIVE: source → ``left_anchor``
    only, destination → ``right_anchor`` only. A token-level revision (an
    insert/delete WITHIN a modified paragraph that exists on both sides)
    carries both enclosing-block anchors; a whole-block insert/delete carries
    only its primary anchor.
    """

    type: DocxDiffRevisionType
    text: str
    author: str
    date: str
    move_group_id: int | None = None
    is_move_source: bool | None = None
    format_change: DocxDiffFormatChange | None = None
    left_anchor: str | None = None
    right_anchor: str | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffRevision":
        fc = d.get("formatChange")
        return cls(
            type=DocxDiffRevisionType._from_wire(d["revisionType"]),
            text=d.get("text", ""),
            author=d.get("author", ""),
            date=d.get("date", ""),
            move_group_id=d.get("moveGroupId"),
            is_move_source=d.get("isMoveSource"),
            format_change=DocxDiffFormatChange._from_wire(fc) if fc else None,
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
        )


# ---------------------------------------------------------------------------
# DocxDiff consolidate — multi-reviewer composite diff
# ---------------------------------------------------------------------------


@dataclass(frozen=True, slots=True)
class DocxDiffReviewer:
    """One reviewer's document for ``docx_diff_consolidate`` and friends.

    ``document`` is the reviewer's DOCX bytes (their redlined or edited version
    of the base); ``author`` is the display name used for their revisions.
    """

    document: bytes
    author: str


@dataclass(frozen=True, slots=True)
class DocxDiffConsolidateSettings:
    """Settings for the ``docx_diff_consolidate_*`` module functions.

    Wraps a :class:`DocxDiffSettings` for the per-reviewer diffs plus
    ``conflict_resolution`` for how to handle competing edits.

    ``to_wire()`` starts from the embedded diff settings wire dict and adds
    ``conflictResolution`` only when it is non-default (i.e. not ``BASE_WINS``),
    matching the sparse-emit convention used by :meth:`DocxDiffSettings.to_wire`.
    """

    diff: DocxDiffSettings = field(default_factory=DocxDiffSettings)
    conflict_resolution: ConflictResolution = ConflictResolution.BASE_WINS

    def to_wire(self) -> dict[str, Any]:
        wire = self.diff.to_wire()
        if self.conflict_resolution != ConflictResolution.BASE_WINS:
            wire["conflictResolution"] = int(self.conflict_resolution)
        return wire


@dataclass(frozen=True, slots=True)
class DocxDiffConflictCompetitor:
    """One reviewer's competing edit within a :class:`DocxDiffConflict`.

    ``author`` is the reviewer's display name; ``result_text`` is the text
    their edit produces at the conflicting span.
    """

    author: str
    result_text: str

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConflictCompetitor":
        return cls(
            author=d.get("author", ""),
            result_text=d.get("resultText", ""),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffConflict:
    """A conflict produced by ``docx_diff_get_conflicts`` — a base span where two
    or more reviewers made incompatible edits.

    ``id`` is a stable integer index for the conflict within this consolidation.
    ``base_anchor`` is the anchor id in the base document. ``token_start`` /
    ``token_end`` delimit the conflicting token range within that block.
    ``policy`` is the :class:`ConflictResolution` that would be applied by
    ``docx_diff_consolidate`` under the current settings. ``competitors`` lists
    each reviewer's competing edit.
    """

    id: int
    base_anchor: str
    token_start: int
    token_end: int
    policy: ConflictResolution
    competitors: tuple[DocxDiffConflictCompetitor, ...]

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConflict":
        return cls(
            id=int(d["id"]),
            base_anchor=d.get("baseAnchor", ""),
            token_start=int(d.get("tokenStart", 0)),
            token_end=int(d.get("tokenEnd", 0)),
            policy=ConflictResolution(int(d.get("policy", 0))),
            competitors=tuple(
                DocxDiffConflictCompetitor._from_wire(c)
                for c in d.get("competitors", ())
            ),
        )


@dataclass(frozen=True, slots=True)
class DocxDiffConsolidatedRevision:
    """One consumer revision from ``docx_diff_get_consolidated_revisions``.

    Mirrors :class:`DocxDiffRevision` with all of its fields plus an optional
    ``conflict_id`` that links this revision to a :class:`DocxDiffConflict`
    when the revision arose from a conflict resolution decision.
    """

    type: DocxDiffRevisionType
    text: str
    author: str
    date: str
    move_group_id: int | None = None
    is_move_source: bool | None = None
    format_change: DocxDiffFormatChange | None = None
    left_anchor: str | None = None
    right_anchor: str | None = None
    conflict_id: int | None = None

    @classmethod
    def _from_wire(cls, d: Mapping[str, Any]) -> "DocxDiffConsolidatedRevision":
        fc = d.get("formatChange")
        return cls(
            type=DocxDiffRevisionType._from_wire(d["revisionType"]),
            text=d.get("text", ""),
            author=d.get("author", ""),
            date=d.get("date", ""),
            move_group_id=d.get("moveGroupId"),
            is_move_source=d.get("isMoveSource"),
            format_change=DocxDiffFormatChange._from_wire(fc) if fc else None,
            left_anchor=d.get("leftAnchor"),
            right_anchor=d.get("rightAnchor"),
            conflict_id=d.get("conflictId"),
        )


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

    #: Maximum undo steps retained (default 20, was 50). Each step is a full document
    #: snapshot, so this is a step count and not a memory bound -- see
    #: ``undo_memory_budget_bytes``.
    undo_depth: int = 20
    #: Approximate ceiling in bytes on memory held by undo/redo snapshots (default 128 MiB).
    #: Oldest history is discarded when exceeded, so on a large document undo may not reach
    #: the full ``undo_depth``; one step is always retained. 0 bounds by depth alone.
    undo_memory_budget_bytes: int = 128 * 1024 * 1024
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
            "undoMemoryBudgetBytes": self.undo_memory_budget_bytes,
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

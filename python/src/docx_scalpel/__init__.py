"""docx-scalpel — anchor-addressed DOCX editing for LLM agents.

A thin client over Docxodus' ``DocxSession``, spoken to via a long-running
``docxodus-pyhost`` subprocess so the session can persist in memory across
many calls.

Quick start::

    from docx_scalpel import open_session

    with open(\"contract.docx\", \"rb\") as f:
        docx_bytes = f.read()

    with open_session(docx_bytes) as session:
        for placeholder in session.find_placeholders():
            session.replace_match(placeholder.match, \"filled value\")
        new_bytes = session.save()

    with open(\"filled.docx\", \"wb\") as f:
        f.write(new_bytes)

The session lives inside a long-running ``docxodus-pyhost`` subprocess and
persists across Python calls until you close it (or exit the ``with`` block).
That persistence is the whole point of this wrapper: LLM agents issue dozens
of small edits to one document, and recreating the session would pay the
parse + Unid annotation + projection cost every time.
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version as _pkg_version

from ._transport import shutdown_host
from .enums import (
    AnchorIdRendering,
    AnchorRenderMode,
    ContextBoundary,
    DiffFormat,
    EditErrorCode,
    EmptyParagraphMode,
    PlaceholderKind,
    PlaceholderKinds,
    Position,
    ProjectionDepth,
    ProjectionScopes,
    RegexOptions,
    TableRenderMode,
    TrackedChangeMode,
    WhitespaceMode,
)
from .errors import DocxodusHostNotFoundError, DocxodusTransportError, DocxScalpelError
from .session import DocxSession, convert_docx_to_html, open_session, ping
from .types import (
    Anchor,
    AnchorInfo,
    AnchorTarget,
    AnnotationUpdate,
    BlockMetadata,
    BlockSlice,
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
    MarkdownPatch,
    MarkdownProjection,
    NumberFormat,
    ReplaceOptions,
    RunFormatting,
    RunFragment,
    SectionInfo,
    TemplatePlaceholder,
    TextMatch,
    WmlToMarkdownConverterSettings,
)

try:
    # Read from the installed wheel's METADATA so __version__ always matches
    # what was actually published, not a constant that drifts from pyproject.
    __version__ = _pkg_version("docx-scalpel")
except PackageNotFoundError:
    # Unpackaged source checkout (e.g. running tests against `git clone` with
    # no editable install yet). Leave a sentinel so callers can detect this.
    __version__ = "0.0.0+source"

__all__ = [
    "__version__",
    # entry points
    "DocxSession",
    "convert_docx_to_html",
    "open_session",
    "ping",
    "shutdown_host",
    # value types
    "Anchor",
    "AnchorInfo",
    "AnchorTarget",
    "BlockMetadata",
    "BlockSlice",
    "BulkEditResult",
    "CharSpan",
    "CrossBlockMatch",
    "AnnotationUpdate",
    "DocumentAnnotation",
    "DocxSessionSettings",
    "EditError",
    "EditResult",
    "EditSummary",
    "FillOptions",
    "FindOptions",
    "FormatOp",
    "HtmlOptions",
    "ListMembership",
    "MarkdownPatch",
    "MarkdownProjection",
    "NumberFormat",
    "ReplaceOptions",
    "RunFormatting",
    "RunFragment",
    "SectionInfo",
    "TemplatePlaceholder",
    "TextMatch",
    "WmlToMarkdownConverterSettings",
    # enums
    "AnchorIdRendering",
    "AnchorRenderMode",
    "ContextBoundary",
    "DiffFormat",
    "EditErrorCode",
    "EmptyParagraphMode",
    "PlaceholderKind",
    "PlaceholderKinds",
    "Position",
    "ProjectionDepth",
    "ProjectionScopes",
    "RegexOptions",
    "TableRenderMode",
    "TrackedChangeMode",
    "WhitespaceMode",
    # errors
    "DocxScalpelError",
    "DocxodusHostNotFoundError",
    "DocxodusTransportError",
]

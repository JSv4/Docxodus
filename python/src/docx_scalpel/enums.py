"""Enums mirroring the C# ``Docxodus`` types.

String-valued enums use the **snake_case wire value** as the enum value, matching
``DocxSessionJson`` output. Flag enums use the **integer bit value** from the C#
``[Flags]`` declaration. The Python wrapper sends these as ints/strings exactly
as the .NET host expects.
"""

from __future__ import annotations

from enum import Enum, IntEnum, IntFlag

__all__ = [
    "Position",
    "EditErrorCode",
    "PlaceholderKind",
    "PlaceholderKinds",
    "ProjectionScopes",
    "ProjectionDepth",
    "ContextBoundary",
    "DiffFormat",
    "WhitespaceMode",
    "TrackedChangeMode",
    "AnchorRenderMode",
    "TableRenderMode",
    "EmptyParagraphMode",
    "AnchorIdRendering",
    "RegexOptions",
]


class Position(str, Enum):
    """Insertion position relative to an anchor."""

    BEFORE = "before"
    AFTER = "after"


class EditErrorCode(str, Enum):
    """All ``EditResult.error.code`` values the .NET surface can emit.

    Values are snake_case strings as serialized by ``DocxSessionJson.EnumToSnake``.
    """

    ANCHOR_NOT_FOUND = "anchor_not_found"
    ANCHOR_WRONG_KIND = "anchor_wrong_kind"
    ANCHORS_NOT_ADJACENT = "anchors_not_adjacent"
    SESSION_DISPOSED = "session_disposed"
    MALFORMED_MARKDOWN = "malformed_markdown"
    UNSUPPORTED_MARKDOWN_SYNTAX = "unsupported_markdown_syntax"
    TABLE_INSERT_NOT_SUPPORTED = "table_insert_not_supported"
    FOOTNOTE_REF_NOT_SUPPORTED = "footnote_ref_not_supported"
    COMMENT_MARKER_NOT_SUPPORTED = "comment_marker_not_supported"
    IMAGE_INSERT_NOT_SUPPORTED = "image_insert_not_supported"
    ANCHOR_TOKEN_IN_PAYLOAD = "anchor_token_in_payload"
    OFFSET_OUT_OF_RANGE = "offset_out_of_range"
    INVALID_POSITION = "invalid_position"
    UNKNOWN_STYLE = "unknown_style"
    INVALID_LIST_LEVEL = "invalid_list_level"
    MALFORMED_XML = "malformed_xml"
    DISALLOWED_NAMESPACE = "disallowed_namespace"
    INCOMPATIBLE_ELEMENT_TYPE = "incompatible_element_type"
    VALIDATION_FAILED = "validation_failed"
    NOTHING_TO_UNDO = "nothing_to_undo"
    NOTHING_TO_REDO = "nothing_to_redo"
    INTERNAL_ERROR = "internal_error"

    @classmethod
    def _missing_(cls, value: object):  # type: ignore[override]
        # Forward-compatibility: a new C# code we don't yet know about
        # decodes to INTERNAL_ERROR rather than raising. The original wire
        # string is still available on EditError.message.
        return cls.INTERNAL_ERROR


class PlaceholderKind(str, Enum):
    """Discriminator for a single ``TemplatePlaceholder``."""

    BLANK_FILL = "blank_fill"
    ALTERNATIVE_CLAUSE = "alternative_clause"
    INSTRUCTION = "instruction"


class PlaceholderKinds(IntFlag):
    """Bitmask filter for ``find_placeholders`` / ``remaining_placeholders``."""

    BLANK_FILL = 1
    ALTERNATIVE_CLAUSE = 2
    INSTRUCTION = 4
    ALL = BLANK_FILL | ALTERNATIVE_CLAUSE | INSTRUCTION


class ProjectionScopes(IntFlag):
    """Which document parts a projection / find operation should include."""

    BODY = 1
    HEADERS = 2
    FOOTERS = 4
    FOOTNOTES = 8
    ENDNOTES = 16
    COMMENTS = 32
    ALL = BODY | HEADERS | FOOTERS | FOOTNOTES | ENDNOTES | COMMENTS


class ProjectionDepth(IntEnum):
    """How much of the document a ``project_anchor`` call returns."""

    SELF_ONLY = 0
    SUBTREE = 1
    SUBTREE_AND_FOLLOWING_SIBLINGS = 2


class ContextBoundary(IntEnum):
    """Where ``contextBefore`` / ``contextAfter`` strings are clipped."""

    CHAR = 0
    BRACKET = 1
    SENTENCE = 2
    COMMA = 3


class DiffFormat(IntEnum):
    """Output format for ``get_diff``. JSON is the only one currently implemented."""

    JSON = 0
    UNIFIED = 1
    SIDE_BY_SIDE = 2


class WhitespaceMode(IntEnum):
    """How ``grep`` / ``grep_cross_block`` handle whitespace before matching."""

    PRESERVE = 0
    NORMALIZE = 1


class DocxDiffRevisionType(str, Enum):
    """Kind of a :class:`DocxDiffRevision` from ``docx_diff_get_revisions``.

    String-valued so the wire JSON (the .NET ``DocxDiffRevisionType`` name)
    round-trips transparently.
    """

    INSERTED = "Inserted"
    DELETED = "Deleted"
    MOVED = "Moved"
    FORMAT_CHANGED = "FormatChanged"

    @classmethod
    def _from_wire(cls, raw: str) -> "DocxDiffRevisionType":
        return cls(raw)


class DocxDiffRevisionGranularity(IntEnum):
    """How ``docx_diff_get_revisions`` projects the edit script to revisions.

    ``FINE`` (the default) is the engine's native one-revision-per-token-span
    grain; ``WML_COMPARER_COMPATIBLE`` coalesces to counts/texts comparable to
    the shipped ``WmlComparer``. Integer-coded to match the .NET enum positions.
    """

    FINE = 0
    WML_COMPARER_COMPATIBLE = 1


class DocxDiffFormatComparison(IntEnum):
    """How ``docx_diff`` compares run formatting.

    ``MODELED_ONLY`` (the default) compares only the modeled rPr fields;
    ``FULL`` includes the unmodeled rPr digest. Integer-coded to match the
    .NET enum positions.
    """

    MODELED_ONLY = 0
    FULL = 1


class TrackedChangeMode(str, Enum):
    """How mutations land in the underlying OOXML."""

    ACCEPT = "accept"
    RENDER_INLINE = "render_inline"
    STRIP_DELETIONS = "strip_deletions"


class AnchorRenderMode(IntEnum):
    BLOCK = 0
    BLOCK_AND_INLINE = 1
    NONE = 2


class TableRenderMode(IntEnum):
    GFM_WITH_OPAQUE_FALLBACK = 0
    ALWAYS_GFM = 1
    ALWAYS_OPAQUE = 2


class EmptyParagraphMode(IntEnum):
    ANCHOR_ONLY = 0
    MARKED_EMPTY = 1
    SUPPRESS = 2


class AnchorIdRendering(IntEnum):
    FULL_UNID = 0
    ABBREVIATED = 1
    SEQUENTIAL = 2


class RegexOptions(IntFlag):
    """Subset of .NET ``System.Text.RegularExpressions.RegexOptions`` we expose.

    Values match the .NET enum exactly so they can be passed through unchanged.
    """

    NONE = 0
    IGNORE_CASE = 1
    MULTILINE = 2
    SINGLELINE = 16

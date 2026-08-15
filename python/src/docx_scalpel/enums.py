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
    "HeaderFooterKind",
    "PageNumberField",
    "ParagraphAlignment",
    "ListFormat",
    "EditErrorCode",
    "MutationBatchMode",
    "MutationPreviewHtmlMode",
    "PlaceholderKind",
    "PlaceholderKinds",
    "ProjectionScopes",
    "HyperlinkKind",
    "ProjectionDepth",
    "ContextBoundary",
    "DiffFormat",
    "WhitespaceMode",
    "DocxDiffRevisionType",
    "DocxDiffRevisionGranularity",
    "DocxDiffFormatComparison",
    "TrackedChangeMode",
    "AnchorRenderMode",
    "TableRenderMode",
    "EmptyParagraphMode",
    "AnchorIdRendering",
    "RegexOptions",
    "ConflictResolution",
    "TableAnchorEntityKind",
    "TableRowHeightRule",
    "TableVerticalMergeRole",
]


class Position(str, Enum):
    """Insertion position relative to an anchor."""

    BEFORE = "before"
    AFTER = "after"


class TableAnchorEntityKind(str, Enum):
    """Structural identity kind in a table-anchor mutation mapping."""

    TABLE = "table"
    ROW = "row"
    COLUMN = "column"
    CELL = "cell"


class TableVerticalMergeRole(str, Enum):
    """A physical cell's role in a Word vertical-merge run."""

    NONE = "none"
    RESTART = "restart"
    CONTINUE = "continue"


class TableRowHeightRule(str, Enum):
    """Interpretation of an explicit table-row height."""

    AUTO = "auto"
    AT_LEAST = "atLeast"
    EXACT = "exact"


class HeaderFooterKind(str, Enum):
    """Which header/footer story ``set_header_text``/``set_footer_text`` targets.

    Maps to the OOXML ``w:type``: ``DEFAULT`` (all pages), ``FIRST`` (first-page-only;
    sets ``w:titlePg``), ``EVEN`` (even pages; sets ``w:evenAndOddHeaders``).
    """

    DEFAULT = "default"
    FIRST = "first"
    EVEN = "even"


class PageNumberField(str, Enum):
    """Which page-number field ``insert_page_number_field`` emits."""

    CURRENT_PAGE = "currentPage"
    TOTAL_PAGES = "totalPages"


class ParagraphAlignment(str, Enum):
    """Paragraph alignment for ``ParagraphFormatOp.alignment`` (maps to ``w:jc``;
    ``JUSTIFY`` writes ``w:val="both"``)."""

    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class LineSpacingRule(str, Enum):
    """How ``ParagraphFormatOp.line_spacing`` is interpreted (``w:spacing/@w:lineRule``).

    Under ``AUTO`` the value is in 240ths of a line (240 = single, 360 = 1.5x,
    480 = double); under ``EXACT``/``AT_LEAST`` it is a height in twips
    (20 twips = 1pt — e.g. 480 = exactly 24pt).
    """

    AUTO = "auto"
    EXACT = "exact"
    AT_LEAST = "atLeast"


class ListFormat(str, Enum):
    """List format for ``apply_list_format`` / ``apply_list_format_range``.

    The plain numbered formats render ``1.`` / ``a.`` / ``i.`` level text; the
    ``*_PARENTHESIS`` variants render ``(1)`` / ``(a)`` / ``(i)`` — same ``w:numFmt``,
    different ``w:lvlText`` (the legal-drafting presets). ``NONE`` strips inline list
    membership.
    """

    NONE = "none"
    BULLET = "bullet"
    DECIMAL = "decimal"
    LOWER_LETTER = "lowerLetter"
    UPPER_LETTER = "upperLetter"
    LOWER_ROMAN = "lowerRoman"
    UPPER_ROMAN = "upperRoman"
    DECIMAL_PARENTHESIS = "decimalParenthesis"
    LOWER_LETTER_PARENTHESIS = "lowerLetterParenthesis"
    UPPER_LETTER_PARENTHESIS = "upperLetterParenthesis"
    LOWER_ROMAN_PARENTHESIS = "lowerRomanParenthesis"
    UPPER_ROMAN_PARENTHESIS = "upperRomanParenthesis"


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
    INVALID_LIST_START_VALUE = "invalid_list_start_value"
    INVALID_PAGE_NUMBERING = "invalid_page_numbering"
    INVALID_PARAGRAPH_FORMAT = "invalid_paragraph_format"
    INVALID_TABLE_STYLING = "invalid_table_styling"
    INVALID_TABLE_MERGE = "invalid_table_merge"
    TABLE_ANCHOR_MIGRATION_REQUIRED = "table_anchor_migration_required"
    MALFORMED_XML = "malformed_xml"
    DISALLOWED_NAMESPACE = "disallowed_namespace"
    INCOMPATIBLE_ELEMENT_TYPE = "incompatible_element_type"
    VALIDATION_FAILED = "validation_failed"
    PRECONDITION_FAILED = "precondition_failed"
    NOTHING_TO_UNDO = "nothing_to_undo"
    NOTHING_TO_REDO = "nothing_to_redo"
    DUPLICATE_ANNOTATION_ID = "duplicate_annotation_id"
    ANNOTATION_NOT_FOUND = "annotation_not_found"
    EMPTY_ANNOTATION_SPAN = "empty_annotation_span"
    EMPTY_COMMENT_SPAN = "empty_comment_span"
    REVISION_NOT_FOUND = "revision_not_found"
    INVALID_BATCH_STEP = "invalid_batch_step"
    HYPERLINK_NOT_FOUND = "hyperlink_not_found"
    BOOKMARK_NOT_FOUND = "bookmark_not_found"
    DUPLICATE_BOOKMARK_NAME = "duplicate_bookmark_name"
    INVALID_BOOKMARK_NAME = "invalid_bookmark_name"
    INVALID_HYPERLINK_TARGET = "invalid_hyperlink_target"
    MISSING_BOOKMARK_TARGET = "missing_bookmark_target"
    BOOKMARK_IN_USE = "bookmark_in_use"
    MANAGED_BOOKMARK = "managed_bookmark"
    EMPTY_HYPERLINK_SPAN = "empty_hyperlink_span"
    UNSUPPORTED_INLINE_BOUNDARY = "unsupported_inline_boundary"
    REVISION_UNSUPPORTED = "revision_unsupported"
    REVISION_MALFORMED = "revision_malformed"
    REVISION_AMBIGUOUS = "revision_ambiguous"
    TRACKED_OPERATION_UNSUPPORTED = "tracked_operation_unsupported"
    UNRESOLVED_STRUCTURAL_REVISION = "unresolved_structural_revision"
    CONTENT_CONTROL_NOT_FOUND = "content_control_not_found"
    CONTENT_CONTROL_MALFORMED = "content_control_malformed"
    CONTENT_CONTROL_UNSUPPORTED = "content_control_unsupported"
    CONTENT_CONTROL_LOCKED = "content_control_locked"
    CONTENT_CONTROL_BOUND = "content_control_bound"
    CONTENT_CONTROL_WRONG_TYPE = "content_control_wrong_type"
    INVALID_CONTENT_CONTROL_VALUE = "invalid_content_control_value"
    CONTENT_CONTROL_PLACEMENT_UNSUPPORTED = "content_control_placement_unsupported"
    CONTENT_CONTROL_NESTED_FILL_UNSUPPORTED = "content_control_nested_fill_unsupported"
    REPEATING_SECTION_CONSTRAINT = "repeating_section_constraint"
    IMAGE_NOT_FOUND = "image_not_found"
    INVALID_IMAGE_DATA = "invalid_image_data"
    UNSUPPORTED_IMAGE_FORMAT = "unsupported_image_format"
    IMAGE_TOO_LARGE = "image_too_large"
    INVALID_IMAGE_DIMENSIONS = "invalid_image_dimensions"
    UNSUPPORTED_IMAGE_MARKUP = "unsupported_image_markup"
    LINKED_IMAGE_READ_ONLY = "linked_image_read_only"
    INVALID_IMAGE_LAYOUT = "invalid_image_layout"
    INTERNAL_ERROR = "internal_error"

    @classmethod
    def _missing_(cls, value: object) -> "EditErrorCode":  # type: ignore[override]
        # Forward-compatibility: a new C# code we don't yet know about
        # decodes to INTERNAL_ERROR rather than raising. The original wire
        # string is still available on EditError.message.
        return cls.INTERNAL_ERROR


class MutationBatchMode(str, Enum):
    """Atomic is the safe default; best-effort explicitly retains partial successes."""

    ATOMIC = "atomic"
    BEST_EFFORT = "best_effort"


class MutationPreviewHtmlMode(str, Enum):
    """Optional HTML a ``preview_batch`` renders from its isolated shadow package.

    ``SCOPED`` requires ``html_anchor_id`` and renders that one block; ``FULL``
    renders the whole predicted document. Both render with tracked changes,
    comments, annotations and notes shown — a preview describes the document the
    batch would produce, not an authoring view of it.
    """

    NONE = "none"
    SCOPED = "scoped"
    FULL = "full"


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


class HyperlinkKind(str, Enum):
    """Native Word hyperlink target representation."""

    EXTERNAL = "external"
    INTERNAL = "internal"


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


class ConflictResolution(IntEnum):
    """How ``docx_diff_consolidate`` resolves competing edits at the same base span.

    Integer-coded to match the .NET ``ConflictResolution`` enum positions.
    """

    BASE_WINS = 0
    FIRST_REVIEWER_WINS = 1
    STACK_ALL = 2


class RegexOptions(IntFlag):
    """Subset of .NET ``System.Text.RegularExpressions.RegexOptions`` we expose.

    Values match the .NET enum exactly so they can be passed through unchanged.
    """

    NONE = 0
    IGNORE_CASE = 1
    MULTILINE = 2
    SINGLELINE = 16

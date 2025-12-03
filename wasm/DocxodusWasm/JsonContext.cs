using System.Text.Json.Serialization;

namespace DocxodusWasm;

/// <summary>
/// JSON serialization context for AOT/trimming-safe serialization.
/// Uses source generators to avoid reflection.
/// </summary>
[JsonSerializable(typeof(ErrorResponse))]
[JsonSerializable(typeof(VersionInfo))]
[JsonSerializable(typeof(RevisionsResponse))]
[JsonSerializable(typeof(RevisionInfo))]
[JsonSerializable(typeof(RevisionInfo[]))]
[JsonSerializable(typeof(FormatChangeInfo))]
[JsonSerializable(typeof(AnnotationInfo))]
[JsonSerializable(typeof(AnnotationInfo[]))]
[JsonSerializable(typeof(AnnotationsResponse))]
[JsonSerializable(typeof(HasAnnotationsResponse))]
[JsonSerializable(typeof(AddAnnotationBase64Response))]
[JsonSerializable(typeof(RemoveAnnotationResponse))]
[JsonSerializable(typeof(AddAnnotationRequest))]
[JsonSerializable(typeof(AddAnnotationResponse))]
[JsonSerializable(typeof(AddAnnotationWithTargetRequest))]
[JsonSerializable(typeof(DocumentStructureResponse))]
[JsonSerializable(typeof(DocumentElementInfo))]
[JsonSerializable(typeof(DocumentElementInfo[]))]
[JsonSerializable(typeof(TableColumnInfoDto))]
[JsonSerializable(typeof(TableColumnInfoDto[]))]
[JsonSerializable(typeof(Dictionary<string, string>))]
[JsonSerializable(typeof(Dictionary<string, DocumentElementInfo>))]
[JsonSerializable(typeof(Dictionary<string, TableColumnInfoDto>))]
[JsonSerializable(typeof(DocumentMetadataResponse))]
[JsonSerializable(typeof(SectionMetadataInfo))]
[JsonSerializable(typeof(SectionMetadataInfo[]))]
// Profiling types
[JsonSerializable(typeof(ProfilingResponse))]
[JsonSerializable(typeof(Dictionary<string, double>))]
// OpenContracts export types
[JsonSerializable(typeof(OpenContractExportResponse))]
[JsonSerializable(typeof(PawlsPageDto))]
[JsonSerializable(typeof(PawlsPageDto[]))]
[JsonSerializable(typeof(PawlsPageBoundaryDto))]
[JsonSerializable(typeof(PawlsTokenDto))]
[JsonSerializable(typeof(PawlsTokenDto[]))]
[JsonSerializable(typeof(OpenContractsAnnotationDto))]
[JsonSerializable(typeof(OpenContractsAnnotationDto[]))]
[JsonSerializable(typeof(OpenContractsRelationshipDto))]
[JsonSerializable(typeof(OpenContractsRelationshipDto[]))]
[JsonSerializable(typeof(TextSpanDto))]
[JsonSerializable(typeof(BoundingBoxDto))]
[JsonSerializable(typeof(TokenIdDto))]
[JsonSerializable(typeof(TokenIdDto[]))]
[JsonSerializable(typeof(OpenContractsSinglePageAnnotationDto))]
[JsonSerializable(typeof(Dictionary<string, OpenContractsSinglePageAnnotationDto>))]
internal partial class DocxodusJsonContext : JsonSerializerContext
{
}

public class ErrorResponse
{
    public string Error { get; set; } = "";
    public string? Type { get; set; }
    public string? StackTrace { get; set; }
}

public class VersionInfo
{
    public string Library { get; set; } = "";
    public string DotnetVersion { get; set; } = "";
    public string Platform { get; set; } = "";
}

public class RevisionsResponse
{
    public RevisionInfo[] Revisions { get; set; } = Array.Empty<RevisionInfo>();
}

public class RevisionInfo
{
    public string Author { get; set; } = "";
    public string Date { get; set; } = "";
    public string RevisionType { get; set; } = "";
    public string Text { get; set; } = "";

    /// <summary>
    /// For Moved revisions, this ID links the source and destination.
    /// Both the "from" and "to" revisions share the same MoveGroupId.
    /// Null for non-move revisions.
    /// </summary>
    public int? MoveGroupId { get; set; }

    /// <summary>
    /// For Moved revisions: true = source (moved FROM here),
    /// false = destination (moved TO here).
    /// Null for non-move revisions.
    /// </summary>
    public bool? IsMoveSource { get; set; }

    /// <summary>
    /// For FormatChanged revisions: details about what formatting changed.
    /// Null for non-format-change revisions.
    /// </summary>
    public FormatChangeInfo? FormatChange { get; set; }
}

/// <summary>
/// Details about formatting changes for FormatChanged revisions.
/// </summary>
public class FormatChangeInfo
{
    /// <summary>
    /// Dictionary of old property names and values.
    /// </summary>
    public Dictionary<string, string>? OldProperties { get; set; }

    /// <summary>
    /// Dictionary of new property names and values.
    /// </summary>
    public Dictionary<string, string>? NewProperties { get; set; }

    /// <summary>
    /// List of property names that changed (e.g., "bold", "italic", "fontSize").
    /// </summary>
    public List<string>? ChangedPropertyNames { get; set; }
}

/// <summary>
/// Information about a document annotation.
/// </summary>
public class AnnotationInfo
{
    /// <summary>
    /// Unique annotation ID.
    /// </summary>
    public string Id { get; set; } = "";

    /// <summary>
    /// Label category/type identifier (e.g., "CLAUSE_TYPE_A", "DATE_REF").
    /// </summary>
    public string LabelId { get; set; } = "";

    /// <summary>
    /// Human-readable label text.
    /// </summary>
    public string Label { get; set; } = "";

    /// <summary>
    /// Highlight color in hex format (e.g., "#FFEB3B").
    /// </summary>
    public string Color { get; set; } = "";

    /// <summary>
    /// Author who created the annotation.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Creation timestamp (ISO 8601).
    /// </summary>
    public string? Created { get; set; }

    /// <summary>
    /// Internal bookmark name.
    /// </summary>
    public string? BookmarkName { get; set; }

    /// <summary>
    /// Start page number (if computed).
    /// </summary>
    public int? StartPage { get; set; }

    /// <summary>
    /// End page number (if computed).
    /// </summary>
    public int? EndPage { get; set; }

    /// <summary>
    /// The annotated text content.
    /// </summary>
    public string? AnnotatedText { get; set; }

    /// <summary>
    /// Custom metadata key-value pairs.
    /// </summary>
    public Dictionary<string, string>? Metadata { get; set; }
}

/// <summary>
/// Response containing all annotations.
/// </summary>
public class AnnotationsResponse
{
    public AnnotationInfo[] Annotations { get; set; } = Array.Empty<AnnotationInfo>();
}

/// <summary>
/// Response for checking if document has annotations.
/// </summary>
public class HasAnnotationsResponse
{
    public bool HasAnnotations { get; set; }
}

/// <summary>
/// Response for AddAnnotation with base64-encoded document bytes.
/// </summary>
public class AddAnnotationBase64Response
{
    public bool Success { get; set; }
    public string DocumentBytes { get; set; } = "";
    public AnnotationInfo? Annotation { get; set; }
}

/// <summary>
/// Response for RemoveAnnotation with base64-encoded document bytes.
/// </summary>
public class RemoveAnnotationResponse
{
    public bool Success { get; set; }
    public string DocumentBytes { get; set; } = "";
}

/// <summary>
/// Request to add an annotation.
/// </summary>
public class AddAnnotationRequest
{
    /// <summary>
    /// Unique annotation ID.
    /// </summary>
    public string Id { get; set; } = "";

    /// <summary>
    /// Label category/type identifier.
    /// </summary>
    public string LabelId { get; set; } = "";

    /// <summary>
    /// Human-readable label text.
    /// </summary>
    public string Label { get; set; } = "";

    /// <summary>
    /// Highlight color in hex format.
    /// </summary>
    public string Color { get; set; } = "#FFEB3B";

    /// <summary>
    /// Author who created the annotation.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Text to search for and annotate.
    /// </summary>
    public string? SearchText { get; set; }

    /// <summary>
    /// Which occurrence to annotate (1-based, default: 1).
    /// </summary>
    public int Occurrence { get; set; } = 1;

    /// <summary>
    /// Start paragraph index (0-based).
    /// </summary>
    public int? StartParagraphIndex { get; set; }

    /// <summary>
    /// End paragraph index (0-based, inclusive).
    /// </summary>
    public int? EndParagraphIndex { get; set; }

    /// <summary>
    /// Custom metadata key-value pairs.
    /// </summary>
    public Dictionary<string, string>? Metadata { get; set; }
}

/// <summary>
/// Response after adding an annotation.
/// </summary>
public class AddAnnotationResponse
{
    /// <summary>
    /// The modified document bytes.
    /// </summary>
    public byte[] DocumentBytes { get; set; } = Array.Empty<byte>();

    /// <summary>
    /// The annotation that was added.
    /// </summary>
    public AnnotationInfo? Annotation { get; set; }
}

/// <summary>
/// Request to add an annotation using flexible targeting.
/// </summary>
public class AddAnnotationWithTargetRequest
{
    /// <summary>
    /// Unique annotation ID.
    /// </summary>
    public string Id { get; set; } = "";

    /// <summary>
    /// Label category/type identifier.
    /// </summary>
    public string LabelId { get; set; } = "";

    /// <summary>
    /// Human-readable label text.
    /// </summary>
    public string Label { get; set; } = "";

    /// <summary>
    /// Highlight color in hex format.
    /// </summary>
    public string Color { get; set; } = "#FFEB3B";

    /// <summary>
    /// Author who created the annotation.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Custom metadata key-value pairs.
    /// </summary>
    public Dictionary<string, string>? Metadata { get; set; }

    // Target specification (one of these should be set):

    /// <summary>
    /// Target by element ID (e.g., "doc/p-0/r-1").
    /// </summary>
    public string? ElementId { get; set; }

    /// <summary>
    /// Element type for index-based targeting: Paragraph, Run, Table, TableRow, TableCell, TableColumn.
    /// </summary>
    public string? ElementType { get; set; }

    /// <summary>
    /// Paragraph index (0-based).
    /// </summary>
    public int? ParagraphIndex { get; set; }

    /// <summary>
    /// Run index within paragraph (0-based).
    /// </summary>
    public int? RunIndex { get; set; }

    /// <summary>
    /// Table index (0-based).
    /// </summary>
    public int? TableIndex { get; set; }

    /// <summary>
    /// Row index within table (0-based).
    /// </summary>
    public int? RowIndex { get; set; }

    /// <summary>
    /// Cell index within row (0-based).
    /// </summary>
    public int? CellIndex { get; set; }

    /// <summary>
    /// Column index for table column targeting (0-based).
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// Text to search for (global or within ElementId).
    /// </summary>
    public string? SearchText { get; set; }

    /// <summary>
    /// Which occurrence of SearchText to target (1-based, default: 1).
    /// </summary>
    public int Occurrence { get; set; } = 1;

    /// <summary>
    /// End paragraph index for range targeting.
    /// </summary>
    public int? RangeEndParagraphIndex { get; set; }
}

/// <summary>
/// Response containing document structure.
/// </summary>
public class DocumentStructureResponse
{
    /// <summary>
    /// Root document element.
    /// </summary>
    public DocumentElementInfo Root { get; set; } = new();

    /// <summary>
    /// All elements indexed by ID for quick lookup.
    /// </summary>
    public Dictionary<string, DocumentElementInfo> ElementsById { get; set; } = new();

    /// <summary>
    /// Table column information indexed by column ID.
    /// </summary>
    public Dictionary<string, TableColumnInfoDto> TableColumns { get; set; } = new();
}

/// <summary>
/// Information about a document element.
/// </summary>
public class DocumentElementInfo
{
    /// <summary>
    /// Unique element ID (path-based, e.g., "doc/tbl-0/tr-1/tc-2").
    /// </summary>
    public string Id { get; set; } = "";

    /// <summary>
    /// Element type: Document, Paragraph, Run, Table, TableRow, TableCell, TableColumn, Hyperlink, Image.
    /// </summary>
    public string Type { get; set; } = "";

    /// <summary>
    /// Preview of text content (first ~100 characters).
    /// </summary>
    public string? TextPreview { get; set; }

    /// <summary>
    /// Position index within parent element.
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// Child elements.
    /// </summary>
    public DocumentElementInfo[] Children { get; set; } = Array.Empty<DocumentElementInfo>();

    /// <summary>
    /// For table rows/cells: the row index.
    /// </summary>
    public int? RowIndex { get; set; }

    /// <summary>
    /// For table cells: the column index.
    /// </summary>
    public int? ColumnIndex { get; set; }

    /// <summary>
    /// For table cells: number of rows this cell spans.
    /// </summary>
    public int? RowSpan { get; set; }

    /// <summary>
    /// For table cells: number of columns this cell spans.
    /// </summary>
    public int? ColumnSpan { get; set; }
}

/// <summary>
/// Information about a table column.
/// </summary>
public class TableColumnInfoDto
{
    /// <summary>
    /// ID of the table this column belongs to.
    /// </summary>
    public string TableId { get; set; } = "";

    /// <summary>
    /// Zero-based column index.
    /// </summary>
    public int ColumnIndex { get; set; }

    /// <summary>
    /// IDs of all cells in this column.
    /// </summary>
    public string[] CellIds { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Total number of rows in this column.
    /// </summary>
    public int RowCount { get; set; }
}

/// <summary>
/// Response containing document metadata for lazy loading pagination.
/// </summary>
public class DocumentMetadataResponse
{
    /// <summary>
    /// Section metadata array.
    /// </summary>
    public SectionMetadataInfo[] Sections { get; set; } = Array.Empty<SectionMetadataInfo>();

    /// <summary>
    /// Total number of paragraphs in the document.
    /// </summary>
    public int TotalParagraphs { get; set; }

    /// <summary>
    /// Total number of tables in the document.
    /// </summary>
    public int TotalTables { get; set; }

    /// <summary>
    /// Whether the document has footnotes.
    /// </summary>
    public bool HasFootnotes { get; set; }

    /// <summary>
    /// Whether the document has endnotes.
    /// </summary>
    public bool HasEndnotes { get; set; }

    /// <summary>
    /// Whether the document has tracked changes.
    /// </summary>
    public bool HasTrackedChanges { get; set; }

    /// <summary>
    /// Whether the document has comments.
    /// </summary>
    public bool HasComments { get; set; }

    /// <summary>
    /// Estimated total page count.
    /// </summary>
    public int EstimatedPageCount { get; set; }
}

/// <summary>
/// Metadata for a single section.
/// </summary>
public class SectionMetadataInfo
{
    /// <summary>
    /// Section index (0-based).
    /// </summary>
    public int SectionIndex { get; set; }

    /// <summary>
    /// Page width in points.
    /// </summary>
    public double PageWidthPt { get; set; }

    /// <summary>
    /// Page height in points.
    /// </summary>
    public double PageHeightPt { get; set; }

    /// <summary>
    /// Top margin in points.
    /// </summary>
    public double MarginTopPt { get; set; }

    /// <summary>
    /// Right margin in points.
    /// </summary>
    public double MarginRightPt { get; set; }

    /// <summary>
    /// Bottom margin in points.
    /// </summary>
    public double MarginBottomPt { get; set; }

    /// <summary>
    /// Left margin in points.
    /// </summary>
    public double MarginLeftPt { get; set; }

    /// <summary>
    /// Content width in points.
    /// </summary>
    public double ContentWidthPt { get; set; }

    /// <summary>
    /// Content height in points.
    /// </summary>
    public double ContentHeightPt { get; set; }

    /// <summary>
    /// Header distance in points.
    /// </summary>
    public double HeaderPt { get; set; }

    /// <summary>
    /// Footer distance in points.
    /// </summary>
    public double FooterPt { get; set; }

    /// <summary>
    /// Number of paragraphs in this section.
    /// </summary>
    public int ParagraphCount { get; set; }

    /// <summary>
    /// Number of tables in this section.
    /// </summary>
    public int TableCount { get; set; }

    /// <summary>
    /// Whether this section has a default header.
    /// </summary>
    public bool HasHeader { get; set; }

    /// <summary>
    /// Whether this section has a default footer.
    /// </summary>
    public bool HasFooter { get; set; }

    /// <summary>
    /// Whether this section has a first page header.
    /// </summary>
    public bool HasFirstPageHeader { get; set; }

    /// <summary>
    /// Whether this section has a first page footer.
    /// </summary>
    public bool HasFirstPageFooter { get; set; }

    /// <summary>
    /// Whether this section has an even page header.
    /// </summary>
    public bool HasEvenPageHeader { get; set; }

    /// <summary>
    /// Whether this section has an even page footer.
    /// </summary>
    public bool HasEvenPageFooter { get; set; }

    /// <summary>
    /// Start paragraph index (0-based, global).
    /// </summary>
    public int StartParagraphIndex { get; set; }

    /// <summary>
    /// End paragraph index (exclusive, global).
    /// </summary>
    public int EndParagraphIndex { get; set; }

    /// <summary>
    /// Start table index (0-based, global).
    /// </summary>
    public int StartTableIndex { get; set; }

    /// <summary>
    /// End table index (exclusive, global).
    /// </summary>
    public int EndTableIndex { get; set; }
}

/// <summary>
/// Response containing profiling timing data for document conversion.
/// </summary>
public class ProfilingResponse
{
    /// <summary>
    /// The generated HTML content.
    /// </summary>
    public string Html { get; set; } = "";

    /// <summary>
    /// Timing breakdown in milliseconds for each operation.
    /// </summary>
    public Dictionary<string, double> Timings { get; set; } = new();

    /// <summary>
    /// Total conversion time in milliseconds.
    /// </summary>
    public double TotalMs { get; set; }

    /// <summary>
    /// Document statistics.
    /// </summary>
    public int ParagraphCount { get; set; }
    public int TableCount { get; set; }
    public int ElementCount { get; set; }
}

#region OpenContracts Export Types

/// <summary>
/// Response containing the OpenContracts export format.
/// </summary>
public class OpenContractExportResponse
{
    /// <summary>
    /// Document title.
    /// </summary>
    public string Title { get; set; } = "";

    /// <summary>
    /// Complete document text content.
    /// </summary>
    public string Content { get; set; } = "";

    /// <summary>
    /// Optional document description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Estimated page count.
    /// </summary>
    public int PageCount { get; set; }

    /// <summary>
    /// PAWLS-format page layout information.
    /// </summary>
    public PawlsPageDto[] PawlsFileContent { get; set; } = Array.Empty<PawlsPageDto>();

    /// <summary>
    /// Document-level labels.
    /// </summary>
    public string[] DocLabels { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Annotations/labeled text spans.
    /// </summary>
    public OpenContractsAnnotationDto[] LabelledText { get; set; } = Array.Empty<OpenContractsAnnotationDto>();

    /// <summary>
    /// Relationships between annotations.
    /// </summary>
    public OpenContractsRelationshipDto[]? Relationships { get; set; }
}

/// <summary>
/// PAWLS page with boundary and token information.
/// </summary>
public class PawlsPageDto
{
    /// <summary>
    /// Page boundary information.
    /// </summary>
    public PawlsPageBoundaryDto Page { get; set; } = new();

    /// <summary>
    /// Tokens on this page.
    /// </summary>
    public PawlsTokenDto[] Tokens { get; set; } = Array.Empty<PawlsTokenDto>();
}

/// <summary>
/// Page boundary information.
/// </summary>
public class PawlsPageBoundaryDto
{
    /// <summary>
    /// Page width in points.
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// Page height in points.
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// Zero-based page index.
    /// </summary>
    public int Index { get; set; }
}

/// <summary>
/// Token with position information.
/// </summary>
public class PawlsTokenDto
{
    /// <summary>
    /// X coordinate (left edge) in points.
    /// </summary>
    public double X { get; set; }

    /// <summary>
    /// Y coordinate (top edge) in points.
    /// </summary>
    public double Y { get; set; }

    /// <summary>
    /// Token width in points.
    /// </summary>
    public double Width { get; set; }

    /// <summary>
    /// Token height in points.
    /// </summary>
    public double Height { get; set; }

    /// <summary>
    /// The text content of this token.
    /// </summary>
    public string Text { get; set; } = "";
}

/// <summary>
/// OpenContracts annotation format.
/// </summary>
public class OpenContractsAnnotationDto
{
    /// <summary>
    /// Unique annotation identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Label/category for this annotation.
    /// </summary>
    public string AnnotationLabel { get; set; } = "";

    /// <summary>
    /// The raw text content of the annotation.
    /// </summary>
    public string RawText { get; set; } = "";

    /// <summary>
    /// Starting page number (0-indexed).
    /// </summary>
    public int Page { get; set; }

    /// <summary>
    /// Position data for the annotation (TextSpanDto or page-indexed dictionary).
    /// </summary>
    public object? AnnotationJson { get; set; }

    /// <summary>
    /// Parent annotation ID for hierarchical annotations.
    /// </summary>
    public string? ParentId { get; set; }

    /// <summary>
    /// Type of annotation (e.g., "text", "structural").
    /// </summary>
    public string? AnnotationType { get; set; }

    /// <summary>
    /// Whether this is a structural element.
    /// </summary>
    public bool Structural { get; set; }
}

/// <summary>
/// Text span with character offsets.
/// </summary>
public class TextSpanDto
{
    /// <summary>
    /// Optional span identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Start character offset (0-indexed).
    /// </summary>
    public int Start { get; set; }

    /// <summary>
    /// End character offset (exclusive).
    /// </summary>
    public int End { get; set; }

    /// <summary>
    /// The text content of this span.
    /// </summary>
    public string Text { get; set; } = "";
}

/// <summary>
/// Per-page annotation position data.
/// </summary>
public class OpenContractsSinglePageAnnotationDto
{
    /// <summary>
    /// Bounding box for the annotation on this page.
    /// </summary>
    public BoundingBoxDto Bounds { get; set; } = new();

    /// <summary>
    /// Token indices that make up this annotation on this page.
    /// </summary>
    public TokenIdDto[] TokensJsons { get; set; } = Array.Empty<TokenIdDto>();

    /// <summary>
    /// Raw text content on this page.
    /// </summary>
    public string RawText { get; set; } = "";
}

/// <summary>
/// Bounding box coordinates.
/// </summary>
public class BoundingBoxDto
{
    /// <summary>
    /// Top edge coordinate.
    /// </summary>
    public double Top { get; set; }

    /// <summary>
    /// Bottom edge coordinate.
    /// </summary>
    public double Bottom { get; set; }

    /// <summary>
    /// Left edge coordinate.
    /// </summary>
    public double Left { get; set; }

    /// <summary>
    /// Right edge coordinate.
    /// </summary>
    public double Right { get; set; }
}

/// <summary>
/// Token identifier.
/// </summary>
public class TokenIdDto
{
    /// <summary>
    /// Zero-based page index.
    /// </summary>
    public int PageIndex { get; set; }

    /// <summary>
    /// Zero-based token index within the page.
    /// </summary>
    public int TokenIndex { get; set; }
}

/// <summary>
/// Relationship between annotations.
/// </summary>
public class OpenContractsRelationshipDto
{
    /// <summary>
    /// Unique relationship identifier.
    /// </summary>
    public string? Id { get; set; }

    /// <summary>
    /// Label describing the relationship type.
    /// </summary>
    public string RelationshipLabel { get; set; } = "";

    /// <summary>
    /// IDs of source annotations.
    /// </summary>
    public string[] SourceAnnotationIds { get; set; } = Array.Empty<string>();

    /// <summary>
    /// IDs of target annotations.
    /// </summary>
    public string[] TargetAnnotationIds { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Whether this is a structural relationship.
    /// </summary>
    public bool Structural { get; set; }
}

#endregion

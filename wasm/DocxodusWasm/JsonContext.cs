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
[JsonSerializable(typeof(AddAnnotationRequest))]
[JsonSerializable(typeof(AddAnnotationResponse))]
[JsonSerializable(typeof(Dictionary<string, string>))]
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

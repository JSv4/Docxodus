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

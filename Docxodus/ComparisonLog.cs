// Copyright (c) John Scrudato. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.Generic;
using System.Linq;

namespace Docxodus;

/// <summary>
/// Severity level for comparison log entries.
/// </summary>
public enum ComparisonLogLevel
{
    /// <summary>Informational message about the comparison process.</summary>
    Info,
    /// <summary>Warning about a potential issue that didn't prevent comparison.</summary>
    Warning,
    /// <summary>Error that may affect comparison results but didn't stop processing.</summary>
    Error
}

/// <summary>
/// A single log entry from the comparison process.
/// </summary>
public class ComparisonLogEntry
{
    /// <summary>Severity level of this entry.</summary>
    public ComparisonLogLevel Level { get; init; }

    /// <summary>
    /// Machine-readable code identifying the type of issue.
    /// Examples: "ORPHANED_FOOTNOTE_REFERENCE", "MISSING_STYLE", "INVALID_NUMBERING"
    /// </summary>
    public string Code { get; init; } = string.Empty;

    /// <summary>Human-readable description of the issue.</summary>
    public string Message { get; init; } = string.Empty;

    /// <summary>Additional context or technical details (optional).</summary>
    public string? Details { get; init; }

    /// <summary>
    /// Location in the document where the issue occurred (optional).
    /// Format: "part/xpath" e.g., "document.xml/w:p[5]/w:r[2]"
    /// </summary>
    public string? Location { get; init; }

    public override string ToString()
    {
        var loc = Location != null ? $" at {Location}" : "";
        var det = Details != null ? $" ({Details})" : "";
        return $"[{Level}] {Code}: {Message}{loc}{det}";
    }
}

/// <summary>
/// Collects warnings and errors during document comparison.
/// Attach to WmlComparerSettings.Log to enable logging.
/// </summary>
/// <remarks>
/// When a ComparisonLog is provided to WmlComparerSettings, the comparison
/// will attempt to continue past recoverable errors (like orphaned footnote
/// references) and log them instead of throwing exceptions.
///
/// Example usage:
/// <code>
/// var log = new ComparisonLog();
/// var settings = new WmlComparerSettings { Log = log };
/// var result = WmlComparer.Compare(doc1, doc2, settings);
///
/// if (log.HasWarnings)
/// {
///     foreach (var warning in log.Warnings)
///         Console.WriteLine(warning);
/// }
/// </code>
/// </remarks>
public class ComparisonLog
{
    private readonly List<ComparisonLogEntry> _entries = new();
    private readonly object _lock = new();

    /// <summary>All log entries collected during the operation.</summary>
    public IReadOnlyList<ComparisonLogEntry> Entries
    {
        get
        {
            lock (_lock)
            {
                return _entries.ToList();
            }
        }
    }

    /// <summary>Number of entries in the log.</summary>
    public int Count
    {
        get
        {
            lock (_lock)
            {
                return _entries.Count;
            }
        }
    }

    /// <summary>Whether any warnings were logged.</summary>
    public bool HasWarnings
    {
        get
        {
            lock (_lock)
            {
                return _entries.Any(e => e.Level == ComparisonLogLevel.Warning);
            }
        }
    }

    /// <summary>Whether any errors were logged.</summary>
    public bool HasErrors
    {
        get
        {
            lock (_lock)
            {
                return _entries.Any(e => e.Level == ComparisonLogLevel.Error);
            }
        }
    }

    /// <summary>Get only info entries.</summary>
    public IEnumerable<ComparisonLogEntry> InfoEntries
    {
        get
        {
            lock (_lock)
            {
                return _entries.Where(e => e.Level == ComparisonLogLevel.Info).ToList();
            }
        }
    }

    /// <summary>Get only warning entries.</summary>
    public IEnumerable<ComparisonLogEntry> Warnings
    {
        get
        {
            lock (_lock)
            {
                return _entries.Where(e => e.Level == ComparisonLogLevel.Warning).ToList();
            }
        }
    }

    /// <summary>Get only error entries.</summary>
    public IEnumerable<ComparisonLogEntry> Errors
    {
        get
        {
            lock (_lock)
            {
                return _entries.Where(e => e.Level == ComparisonLogLevel.Error).ToList();
            }
        }
    }

    /// <summary>
    /// Add an informational log entry.
    /// </summary>
    public void AddInfo(string code, string message, string? details = null, string? location = null)
    {
        Add(ComparisonLogLevel.Info, code, message, details, location);
    }

    /// <summary>
    /// Add a warning log entry.
    /// </summary>
    public void AddWarning(string code, string message, string? details = null, string? location = null)
    {
        Add(ComparisonLogLevel.Warning, code, message, details, location);
    }

    /// <summary>
    /// Add an error log entry.
    /// </summary>
    public void AddError(string code, string message, string? details = null, string? location = null)
    {
        Add(ComparisonLogLevel.Error, code, message, details, location);
    }

    /// <summary>
    /// Add a log entry with the specified level.
    /// </summary>
    public void Add(ComparisonLogLevel level, string code, string message, string? details = null, string? location = null)
    {
        var entry = new ComparisonLogEntry
        {
            Level = level,
            Code = code,
            Message = message,
            Details = details,
            Location = location
        };

        lock (_lock)
        {
            _entries.Add(entry);
        }
    }

    /// <summary>
    /// Clear all log entries.
    /// </summary>
    public void Clear()
    {
        lock (_lock)
        {
            _entries.Clear();
        }
    }
}

/// <summary>
/// Well-known log entry codes used by the comparison engine.
/// </summary>
public static class ComparisonLogCodes
{
    /// <summary>A footnote reference in the document body has no corresponding footnote definition.</summary>
    public const string OrphanedFootnoteReference = "ORPHANED_FOOTNOTE_REFERENCE";

    /// <summary>An endnote reference in the document body has no corresponding endnote definition.</summary>
    public const string OrphanedEndnoteReference = "ORPHANED_ENDNOTE_REFERENCE";

    /// <summary>A style referenced in the document is not defined in styles.xml.</summary>
    public const string MissingStyle = "MISSING_STYLE";

    /// <summary>A numbering definition referenced in the document is missing.</summary>
    public const string MissingNumberingDefinition = "MISSING_NUMBERING_DEFINITION";

    /// <summary>A relationship referenced in the document is missing.</summary>
    public const string MissingRelationship = "MISSING_RELATIONSHIP";

    /// <summary>An image or media file referenced in the document is missing.</summary>
    public const string MissingMedia = "MISSING_MEDIA";

    /// <summary>The document structure contains unexpected or malformed XML.</summary>
    public const string MalformedXml = "MALFORMED_XML";

    /// <summary>A bookmark reference has no corresponding bookmark start/end.</summary>
    public const string OrphanedBookmark = "ORPHANED_BOOKMARK";
}

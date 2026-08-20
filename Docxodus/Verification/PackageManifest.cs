// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text;
using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>
/// An algorithm-labelled digest used by verification artifacts.  The value is lower-case
/// hexadecimal so it is stable across every Docxodus transport.
/// </summary>
public sealed record VerificationDigest
{
    /// <summary>The digest algorithm.  Version 1 manifests use <c>SHA-256</c>.</summary>
    required public string Algorithm { get; init; }

    /// <summary>Lower-case hexadecimal digest bytes.</summary>
    required public string Value { get; init; }
}

/// <summary>Severity of a structured verification finding.</summary>
public enum VerificationFindingSeverity
{
    /// <summary>Informational fact that does not invalidate the package.</summary>
    Info,

    /// <summary>A suspicious condition for which a useful manifest could still be produced.</summary>
    Warning,

    /// <summary>A malformed or unsupported condition that makes the package invalid.</summary>
    Error,
}

/// <summary>
/// Stable location vocabulary shared by package findings and future verification diffs/receipts.
/// Fields that do not apply to a finding are null.
/// </summary>
public sealed record ChangeLocation
{
    /// <summary>Package entry/part URI, including the leading slash.</summary>
    public string? EntryUri { get; init; }

    /// <summary>Relationship owner URI; <c>/</c> means the package itself.</summary>
    public string? OwnerUri { get; init; }

    /// <summary>Relationship ID within <see cref="OwnerUri"/>.</summary>
    public string? RelationshipId { get; init; }

    /// <summary>Raw or resolved target URI implicated by the finding.</summary>
    public string? TargetUri { get; init; }

    /// <summary>Schema/property path for non-entry findings.</summary>
    public string? PropertyPath { get; init; }
}

/// <summary>A machine-readable package validation or safety finding.</summary>
public sealed record VerificationFinding
{
    /// <summary>Stable snake-case code suitable for programmatic matching.</summary>
    required public string Code { get; init; }

    /// <summary>Finding severity.</summary>
    required public VerificationFindingSeverity Severity { get; init; }

    /// <summary>Human-readable diagnostic detail.  Consumers should branch on <see cref="Code"/>.</summary>
    required public string Message { get; init; }

    /// <summary>Optional package location associated with the finding.</summary>
    public ChangeLocation? Location { get; init; }
}

/// <summary>Safety limits applied while inspecting untrusted ZIP/XML input.</summary>
public sealed record PackageManifestOptions
{
    /// <summary>Maximum number of central-directory entries to inspect.</summary>
    public int MaxEntryCount { get; init; } = 10_000;

    /// <summary>
    /// Maximum uncompressed bytes read from one entry. The default preserves the original
    /// manifest behavior by matching the default whole-package budget.
    /// </summary>
    public long MaxEntryUncompressedBytes { get; init; } = 1024L * 1024 * 1024;

    /// <summary>Maximum total declared uncompressed bytes (default 1 GiB).</summary>
    public long MaxTotalUncompressedBytes { get; init; } = 1024L * 1024 * 1024;

    /// <summary>Maximum XML part size parsed for normalization/facts (default 32 MiB).</summary>
    public long MaxXmlPartBytes { get; init; } = 32L * 1024 * 1024;

    /// <summary>Maximum declared uncompressed/compressed ratio accepted for an entry.</summary>
    public double MaxCompressionRatio { get; init; } = 1_000;

    /// <summary>Maximum canonical package URI length.</summary>
    public int MaxUriLength { get; init; } = 2_048;

    internal void Validate()
    {
        if (MaxEntryCount <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxEntryCount));
        if (MaxEntryUncompressedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxEntryUncompressedBytes));
        if (MaxTotalUncompressedBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxTotalUncompressedBytes));
        if (MaxXmlPartBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxXmlPartBytes));
        if (!double.IsFinite(MaxCompressionRatio) || MaxCompressionRatio <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxCompressionRatio));
        if (MaxUriLength <= 0)
            throw new ArgumentOutOfRangeException(nameof(MaxUriLength));
    }
}

/// <summary>One physical ZIP entry.  Duplicate names remain separate occurrences.</summary>
public sealed record PackageManifestEntry
{
    /// <summary>Canonical package URI (leading slash, forward slashes).</summary>
    required public string Uri { get; init; }

    /// <summary>Zero-based occurrence among entries with the same case-insensitive URI.</summary>
    required public int Occurrence { get; init; }

    /// <summary>Resolved OPC MIME type, or null when the package does not declare one.</summary>
    public string? ContentType { get; init; }

    /// <summary><c>override</c>, <c>default</c>, <c>implicit</c>, or <c>unresolved</c>.</summary>
    required public string ContentTypeSource { get; init; }

    /// <summary>
    /// Declared uncompressed byte length. Canonical JSON emits this as a decimal string so ZIP64
    /// values remain lossless in clients whose numeric type cannot represent every <see cref="long"/>.
    /// </summary>
    required public long Size { get; init; }

    /// <summary>Compressed ZIP byte length; also a decimal string in canonical JSON.</summary>
    required public long CompressedSize { get; init; }

    /// <summary>SHA-256 over the exact uncompressed entry bytes; null when reading was unsafe or unavailable.</summary>
    public VerificationDigest? RawBytesDigest { get; init; }

    /// <summary>
    /// SHA-256 over the documented XML normalization; null for non-XML or malformed/limited XML.
    /// </summary>
    public VerificationDigest? NormalizedXmlDigest { get; init; }

    /// <summary>Whether this entry was treated as XML from its MIME type or conventional suffix.</summary>
    required public bool IsXml { get; init; }

    /// <summary>
    /// Whether the ZIP general-purpose flag marks the entry as encrypted, or null when the
    /// central directory could not be parsed authoritatively.
    /// </summary>
    public bool? IsEncrypted { get; init; }
}

/// <summary>One Default or Override declaration from <c>[Content_Types].xml</c>.</summary>
public sealed record PackageContentTypeDeclaration
{
    /// <summary><c>default</c> or <c>override</c>.</summary>
    required public string Kind { get; init; }

    /// <summary>Declared extension spelling for a default, or canonical part URI for an override.</summary>
    required public string Key { get; init; }

    /// <summary>Declared MIME type.</summary>
    required public string ContentType { get; init; }

    /// <summary>Zero-based occurrence of this declaration key in package order.</summary>
    required public int Occurrence { get; init; }
}

/// <summary>One package-level or part-level OPC relationship.</summary>
public sealed record PackageRelationship
{
    /// <summary>Owning part URI; <c>/</c> denotes the package relationship part.</summary>
    required public string OwnerUri { get; init; }

    /// <summary>Relationship identifier, unique within the owner for a valid package.</summary>
    required public string Id { get; init; }

    /// <summary>Relationship type URI.</summary>
    required public string Type { get; init; }

    /// <summary>Target exactly as written in the relationship XML.</summary>
    required public string Target { get; init; }

    /// <summary><c>Internal</c> or <c>External</c>.</summary>
    required public string TargetMode { get; init; }

    /// <summary>Canonical internal target part URI, otherwise null.</summary>
    public string? ResolvedTargetUri { get; init; }

    /// <summary>For internal targets, whether a matching ZIP entry exists; null for external targets.</summary>
    public bool? IsTargetPresent { get; init; }
}

/// <summary>Tracked-revision element counts across all XML parts.</summary>
public sealed record PackageRevisionCounts
{
    public int Insertions { get; init; }
    public int Deletions { get; init; }
    public int MoveFrom { get; init; }
    public int MoveTo { get; init; }
    public int PropertyChanges { get; init; }

    /// <summary>Cell insert/delete/merge revision markers.</summary>
    public int StructuralChanges { get; init; }

    /// <summary>Custom-XML revision range starts, one count per logical range.</summary>
    public int OtherChanges { get; init; }
    public int Total { get; init; }
}

/// <summary>Word comment/threading and Docxodus custom-annotation counts.</summary>
public sealed record PackageAnnotationCounts
{
    public int Comments { get; init; }
    public int CommentReplies { get; init; }
    public int ThreadedCommentMetadata { get; init; }
    public int ResolvedComments { get; init; }
    public int People { get; init; }
    public int DocxodusAnnotations { get; init; }
}

/// <summary>High-signal package and renderer facts derived without opening a mutable SDK package.</summary>
public sealed record PackageManifestFacts
{
    /// <summary>
    /// Exact physical tracked-change carriers found during the bounded manifest XML pass. This is
    /// intentionally internal: schema-v1's public revision counts describe logical families, while
    /// proof admission needs a conservative carrier count before constructing a mutable session.
    /// </summary>
    internal long NativeRevisionCarrierCount { get; init; }

    /// <summary>Whether any physical tracked-change carrier uses the strict Word namespace.</summary>
    internal bool HasStrictRevisionMarkup { get; init; }

    public string? MainDocumentUri { get; init; }
    public bool IsStrictOoxml { get; init; }
    public bool IsMacroEnabled { get; init; }
    public bool HasCoreProperties { get; init; }
    public bool HasExtendedProperties { get; init; }
    public bool HasCustomProperties { get; init; }
    public int SectionCount { get; init; }
    public int ParagraphCount { get; init; }
    public int TableCount { get; init; }
    public int HeaderPartCount { get; init; }
    public int FooterPartCount { get; init; }
    public int FootnoteCount { get; init; }
    public int EndnoteCount { get; init; }
    public int StyleCount { get; init; }
    public int NumberingDefinitionCount { get; init; }
    public int ThemePartCount { get; init; }
    public int MediaPartCount { get; init; }
    public int CustomXmlPartCount { get; init; }
    public int DrawingCount { get; init; }
    public int AltChunkCount { get; init; }
    public int FieldCount { get; init; }
    public PackageRevisionCounts Revisions { get; init; } = new();
    public PackageAnnotationCounts Annotations { get; init; } = new();
}

/// <summary>
/// Versioned, deterministic description of an OOXML package.  Use <see cref="ToJson"/> for the
/// canonical wire representation; its property and collection ordering is part of schema v1.
/// </summary>
public sealed record PackageManifest
{
    /// <summary>Stable schema identifier.</summary>
    public const string SchemaId = "https://docxodus.dev/schemas/verification/package-manifest/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;

    /// <summary>
    /// <c>opc</c>, <c>zip</c>, <c>zip-encrypted</c>, <c>ole-encrypted</c>, <c>ole</c>,
    /// or <c>malformed</c>.
    /// </summary>
    required public string PackageKind { get; init; }

    /// <summary>True only when no error-severity findings were emitted.</summary>
    required public bool IsValid { get; init; }

    /// <summary>SHA-256 of the exact supplied byte array, including ZIP container metadata.</summary>
    required public VerificationDigest RawPackageBytesDigest { get; init; }

    /// <summary>
    /// SHA-256 over URI-ordered entry identities and their exact uncompressed-byte digests. ZIP
    /// timestamps, entry order, and compression do not affect it. Null when any entry could not
    /// safely be read.
    /// </summary>
    public VerificationDigest? OrderedOpcContentDigest { get; init; }

    /// <summary>
    /// SHA-256 over URI/content-type and normalized XML (or raw binary) digests.  XML-only
    /// serialization changes documented by the normalizer do not affect it.
    /// </summary>
    public VerificationDigest? NormalizedSemanticDigest { get; init; }

    public IReadOnlyList<PackageManifestEntry> Entries { get; init; } = Array.Empty<PackageManifestEntry>();
    public IReadOnlyList<PackageContentTypeDeclaration> ContentTypes { get; init; } = Array.Empty<PackageContentTypeDeclaration>();
    public IReadOnlyList<PackageRelationship> Relationships { get; init; } = Array.Empty<PackageRelationship>();
    public PackageManifestFacts Facts { get; init; } = new();
    public IReadOnlyList<VerificationFinding> Findings { get; init; } = Array.Empty<VerificationFinding>();

    /// <summary>Serialize the schema-v1 canonical JSON representation.</summary>
    public string ToJson(bool indented = false) => Encoding.UTF8.GetString(ToJsonBytes(indented));

    /// <summary>Serialize the schema-v1 canonical UTF-8 JSON representation.</summary>
    public byte[] ToJsonBytes(bool indented = false)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = indented }))
        {
            WriteJson(writer);
        }
        return stream.ToArray();
    }

    private void WriteJson(Utf8JsonWriter writer)
    {
        writer.WriteStartObject();
        writer.WriteString("schema", Schema);
        writer.WriteNumber("schemaVersion", SchemaVersion);
        writer.WriteString("packageKind", PackageKind);
        writer.WriteBoolean("isValid", IsValid);
        writer.WritePropertyName("rawPackageBytesDigest");
        WriteDigest(writer, RawPackageBytesDigest);
        writer.WritePropertyName("orderedOpcContentDigest");
        WriteDigest(writer, OrderedOpcContentDigest);
        writer.WritePropertyName("normalizedSemanticDigest");
        WriteDigest(writer, NormalizedSemanticDigest);

        writer.WriteStartArray("entries");
        foreach (var entry in Entries)
        {
            writer.WriteStartObject();
            writer.WriteString("uri", entry.Uri);
            writer.WriteNumber("occurrence", entry.Occurrence);
            WriteNullableString(writer, "contentType", entry.ContentType);
            writer.WriteString("contentTypeSource", entry.ContentTypeSource);
            // ZIP64 sizes can exceed JavaScript's 53-bit safe-integer range.  Decimal strings keep
            // schema-v1 lossless in every JSON client instead of silently rounding large entries.
            writer.WriteString("size", entry.Size.ToString(CultureInfo.InvariantCulture));
            writer.WriteString("compressedSize",
                entry.CompressedSize.ToString(CultureInfo.InvariantCulture));
            writer.WritePropertyName("rawBytesDigest");
            WriteDigest(writer, entry.RawBytesDigest);
            writer.WritePropertyName("normalizedXmlDigest");
            WriteDigest(writer, entry.NormalizedXmlDigest);
            writer.WriteBoolean("isXml", entry.IsXml);
            if (entry.IsEncrypted is { } isEncrypted)
                writer.WriteBoolean("isEncrypted", isEncrypted);
            else
                writer.WriteNull("isEncrypted");
            writer.WriteEndObject();
        }
        writer.WriteEndArray();

        writer.WriteStartArray("contentTypes");
        foreach (var declaration in ContentTypes)
        {
            writer.WriteStartObject();
            writer.WriteString("kind", declaration.Kind);
            writer.WriteString("key", declaration.Key);
            writer.WriteString("contentType", declaration.ContentType);
            writer.WriteNumber("occurrence", declaration.Occurrence);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();

        writer.WriteStartArray("relationships");
        foreach (var relationship in Relationships)
        {
            writer.WriteStartObject();
            writer.WriteString("ownerUri", relationship.OwnerUri);
            writer.WriteString("id", relationship.Id);
            writer.WriteString("type", relationship.Type);
            writer.WriteString("target", relationship.Target);
            writer.WriteString("targetMode", relationship.TargetMode);
            WriteNullableString(writer, "resolvedTargetUri", relationship.ResolvedTargetUri);
            if (relationship.IsTargetPresent is { } targetPresent)
                writer.WriteBoolean("isTargetPresent", targetPresent);
            else
                writer.WriteNull("isTargetPresent");
            writer.WriteEndObject();
        }
        writer.WriteEndArray();

        writer.WritePropertyName("facts");
        WriteFacts(writer, Facts);

        writer.WriteStartArray("findings");
        foreach (var finding in Findings)
        {
            writer.WriteStartObject();
            writer.WriteString("code", finding.Code);
            writer.WriteString("severity", finding.Severity switch
            {
                VerificationFindingSeverity.Info => "info",
                VerificationFindingSeverity.Warning => "warning",
                _ => "error",
            });
            writer.WriteString("message", finding.Message);
            writer.WritePropertyName("location");
            if (finding.Location is null)
            {
                writer.WriteNullValue();
            }
            else
            {
                writer.WriteStartObject();
                WriteNullableString(writer, "entryUri", finding.Location.EntryUri);
                WriteNullableString(writer, "ownerUri", finding.Location.OwnerUri);
                WriteNullableString(writer, "relationshipId", finding.Location.RelationshipId);
                WriteNullableString(writer, "targetUri", finding.Location.TargetUri);
                WriteNullableString(writer, "propertyPath", finding.Location.PropertyPath);
                writer.WriteEndObject();
            }
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
        writer.WriteEndObject();
        writer.Flush();
    }

    private static void WriteFacts(Utf8JsonWriter writer, PackageManifestFacts facts)
    {
        writer.WriteStartObject();
        WriteNullableString(writer, "mainDocumentUri", facts.MainDocumentUri);
        writer.WriteBoolean("isStrictOoxml", facts.IsStrictOoxml);
        writer.WriteBoolean("isMacroEnabled", facts.IsMacroEnabled);
        writer.WriteBoolean("hasCoreProperties", facts.HasCoreProperties);
        writer.WriteBoolean("hasExtendedProperties", facts.HasExtendedProperties);
        writer.WriteBoolean("hasCustomProperties", facts.HasCustomProperties);
        writer.WriteNumber("sectionCount", facts.SectionCount);
        writer.WriteNumber("paragraphCount", facts.ParagraphCount);
        writer.WriteNumber("tableCount", facts.TableCount);
        writer.WriteNumber("headerPartCount", facts.HeaderPartCount);
        writer.WriteNumber("footerPartCount", facts.FooterPartCount);
        writer.WriteNumber("footnoteCount", facts.FootnoteCount);
        writer.WriteNumber("endnoteCount", facts.EndnoteCount);
        writer.WriteNumber("styleCount", facts.StyleCount);
        writer.WriteNumber("numberingDefinitionCount", facts.NumberingDefinitionCount);
        writer.WriteNumber("themePartCount", facts.ThemePartCount);
        writer.WriteNumber("mediaPartCount", facts.MediaPartCount);
        writer.WriteNumber("customXmlPartCount", facts.CustomXmlPartCount);
        writer.WriteNumber("drawingCount", facts.DrawingCount);
        writer.WriteNumber("altChunkCount", facts.AltChunkCount);
        writer.WriteNumber("fieldCount", facts.FieldCount);
        writer.WriteStartObject("revisions");
        writer.WriteNumber("insertions", facts.Revisions.Insertions);
        writer.WriteNumber("deletions", facts.Revisions.Deletions);
        writer.WriteNumber("moveFrom", facts.Revisions.MoveFrom);
        writer.WriteNumber("moveTo", facts.Revisions.MoveTo);
        writer.WriteNumber("propertyChanges", facts.Revisions.PropertyChanges);
        writer.WriteNumber("structuralChanges", facts.Revisions.StructuralChanges);
        writer.WriteNumber("otherChanges", facts.Revisions.OtherChanges);
        writer.WriteNumber("total", facts.Revisions.Total);
        writer.WriteEndObject();
        writer.WriteStartObject("annotations");
        writer.WriteNumber("comments", facts.Annotations.Comments);
        writer.WriteNumber("commentReplies", facts.Annotations.CommentReplies);
        writer.WriteNumber("threadedCommentMetadata", facts.Annotations.ThreadedCommentMetadata);
        writer.WriteNumber("resolvedComments", facts.Annotations.ResolvedComments);
        writer.WriteNumber("people", facts.Annotations.People);
        writer.WriteNumber("docxodusAnnotations", facts.Annotations.DocxodusAnnotations);
        writer.WriteEndObject();
        writer.WriteEndObject();
    }

    private static void WriteDigest(Utf8JsonWriter writer, VerificationDigest? digest)
    {
        if (digest is null)
        {
            writer.WriteNullValue();
            return;
        }
        writer.WriteStartObject();
        writer.WriteString("algorithm", digest.Algorithm);
        writer.WriteString("value", digest.Value);
        writer.WriteEndObject();
    }

    private static void WriteNullableString(Utf8JsonWriter writer, string name, string? value)
    {
        if (value is null)
            writer.WriteNull(name);
        else
            writer.WriteString(name, value);
    }
}

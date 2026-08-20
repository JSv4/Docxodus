// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace Docxodus.Verification;

/// <summary>The operation represented by one semantic document change.</summary>
public enum SemanticChangeOperation
{
    Insert = 0,
    Delete = 1,
    Move = 2,
    Modify = 3,
}

/// <summary>
/// Stable v1 classification of semantic DOCX changes. New families may be appended in future schema
/// versions; existing wire names will not be repurposed.
/// </summary>
public enum SemanticChangeFamily
{
    Text = 0,
    BlockStructure = 1,
    RunFormatting = 2,
    ParagraphFormatting = 3,
    Style = 4,
    Numbering = 5,
    List = 6,
    Table = 7,
    TableRow = 8,
    TableCell = 9,
    TableSpan = 10,
    TableWidth = 11,
    TableStyle = 12,
    Section = 13,
    PageSetup = 14,
    Header = 15,
    Footer = 16,
    Field = 17,
    Footnote = 18,
    Endnote = 19,
    Comment = 20,
    Hyperlink = 21,
    Bookmark = 22,
    ContentControl = 23,
    Image = 24,
    Media = 25,
    Relationship = 26,
    Revision = 27,
    Annotation = 28,
    OpaquePackagePart = 29,
}

/// <summary>The discriminator for a typed <see cref="SemanticValue"/>.</summary>
public enum SemanticValueKind
{
    Absent = 0,
    String = 1,
    Boolean = 2,
    Integer = 3,
    Digest = 4,
    Object = 5,
    Array = 6,
}

/// <summary>A named member in an object-valued semantic value.</summary>
public sealed record SemanticProperty(string Name, SemanticValue Value);

/// <summary>
/// A closed, typed value used by the v1 semantic-change schema. It deliberately avoids arbitrary JSON
/// objects: object members are ordinally sorted and all arrays retain their declared order, making the
/// serialized representation deterministic across runtimes and cultures.
/// </summary>
public sealed class SemanticValue
{
    /// <summary>
    /// Inclusive integer bounds for schema v1. Keeping integer values within ECMAScript's exact
    /// range prevents canonical semantic values from changing when JSON crosses a JavaScript client.
    /// </summary>
    public const long MinSafeInteger = -9_007_199_254_740_991L;
    public const long MaxSafeInteger = 9_007_199_254_740_991L;

    private SemanticValue(
        SemanticValueKind kind,
        string? stringValue = null,
        bool? booleanValue = null,
        long? integerValue = null,
        string? digestAlgorithm = null,
        string? digestProfile = null,
        string? digestValue = null,
        IReadOnlyList<SemanticProperty>? properties = null,
        IReadOnlyList<SemanticValue>? items = null)
    {
        Kind = kind;
        StringValue = stringValue;
        BooleanValue = booleanValue;
        IntegerValue = integerValue;
        DigestAlgorithm = digestAlgorithm;
        DigestProfile = digestProfile;
        DigestValue = digestValue;
        Properties = properties ?? System.Array.Empty<SemanticProperty>();
        Items = items ?? System.Array.Empty<SemanticValue>();
    }

    public SemanticValueKind Kind { get; }
    public string? StringValue { get; }
    public bool? BooleanValue { get; }
    public long? IntegerValue { get; }
    public string? DigestAlgorithm { get; }
    public string? DigestProfile { get; }
    public string? DigestValue { get; }
    public IReadOnlyList<SemanticProperty> Properties { get; }
    public IReadOnlyList<SemanticValue> Items { get; }

    public static SemanticValue Absent { get; } = new(SemanticValueKind.Absent);

    public static SemanticValue String(string? value) =>
        value is null ? Absent : new(SemanticValueKind.String, stringValue: value);

    public static SemanticValue Boolean(bool? value) =>
        value is null ? Absent : new(SemanticValueKind.Boolean, booleanValue: value);

    public static SemanticValue Integer(long? value)
    {
        if (value is null) return Absent;
        if (value is < MinSafeInteger or > MaxSafeInteger)
            throw new ArgumentOutOfRangeException(nameof(value), value,
                $"Semantic integers must be between {MinSafeInteger} and {MaxSafeInteger} inclusive.");
        return new SemanticValue(SemanticValueKind.Integer, integerValue: value);
    }

    /// <summary>
    /// Project an integer that originates in document bytes rather than in modeled IR state.
    /// OOXML attributes such as <c>wp:extent/@cx</c>, <c>w:gridCol/@w</c>, and
    /// <c>w:bookmarkStart/@w:colFirst</c> parse as unbounded <see cref="long"/> values, so a crafted
    /// package can carry one outside the v1 safe range. Such a value is emitted losslessly as an
    /// invariant decimal string instead of throwing: one value's kind degrades, the comparison still
    /// completes, and no two distinct out-of-range values collapse into the same record.
    /// Modeled state that is already <see cref="int"/>-typed calls <see cref="Integer"/> directly so
    /// its range check stays live as an assertion.
    /// </summary>
    internal static SemanticValue IntegerFromDocument(long? value)
    {
        if (value is null) return Absent;
        if (value is < MinSafeInteger or > MaxSafeInteger)
            return String(value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        return new SemanticValue(SemanticValueKind.Integer, integerValue: value);
    }

    /// <summary>
    /// Create a digest value. <paramref name="algorithm"/> names the cryptographic algorithm
    /// (for example <c>SHA-256</c>); <paramref name="profile"/> separately identifies the
    /// domain-specific canonicalization or normalization used before hashing.
    /// </summary>
    public static SemanticValue Digest(string algorithm, string value, string? profile = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(algorithm);
        ArgumentException.ThrowIfNullOrWhiteSpace(value);
        if (profile is not null) ArgumentException.ThrowIfNullOrWhiteSpace(profile);
        return new SemanticValue(SemanticValueKind.Digest,
            digestAlgorithm: algorithm, digestProfile: profile, digestValue: value);
    }

    public static SemanticValue Object(IEnumerable<SemanticProperty> properties)
    {
        ArgumentNullException.ThrowIfNull(properties);
        var ordered = properties
            .OrderBy(property => property.Name, StringComparer.Ordinal)
            .ToArray();
        if (ordered.Any(property => string.IsNullOrWhiteSpace(property.Name)))
            throw new ArgumentException("Semantic object property names must be non-empty.", nameof(properties));
        if (ordered.Select(property => property.Name).Distinct(StringComparer.Ordinal).Count() != ordered.Length)
            throw new ArgumentException("Semantic object property names must be unique.", nameof(properties));
        return new SemanticValue(SemanticValueKind.Object, properties: ordered);
    }

    public static SemanticValue Array(IEnumerable<SemanticValue> items)
    {
        ArgumentNullException.ThrowIfNull(items);
        return new SemanticValue(SemanticValueKind.Array, items: items.ToArray());
    }
}

/// <summary>
/// One stable semantic change. The owning part URI is always present; side-specific anchors and scopes
/// are explicit nullable fields because inserts and deletes naturally have only one side.
/// </summary>
public sealed record SemanticChange
{
    required public string Id { get; init; }
    required public SemanticChangeOperation Operation { get; init; }
    required public SemanticChangeFamily Family { get; init; }
    required public string PartUri { get; init; }
    required public string Path { get; init; }
    public string? LeftAnchor { get; init; }
    public string? RightAnchor { get; init; }
    public string? LeftScope { get; init; }
    public string? RightScope { get; init; }
    public string? MoveId { get; init; }
    required public SemanticValue Before { get; init; }
    required public SemanticValue After { get; init; }
}

/// <summary>
/// Versioned, deterministic semantic comparison of two DOCX packages. This is a durable public schema;
/// it is distinct from the legacy internal edit-script JSON consumed by the redline renderer.
/// </summary>
public sealed class SemanticChangeSet
{
    public const string CurrentSchema = "docxodus.semantic-changes";
    public const int CurrentSchemaVersion = 1;

    public SemanticChangeSet(IReadOnlyList<SemanticChange> changes)
    {
        ArgumentNullException.ThrowIfNull(changes);
        Changes = changes
            .OrderBy(change => change.PartUri, StringComparer.Ordinal)
            .ThenBy(change => change.LeftScope, StringComparer.Ordinal)
            .ThenBy(change => change.RightScope, StringComparer.Ordinal)
            .ThenBy(change => change.LeftAnchor, StringComparer.Ordinal)
            .ThenBy(change => change.RightAnchor, StringComparer.Ordinal)
            .ThenBy(change => (int)change.Family)
            .ThenBy(change => change.Path, StringComparer.Ordinal)
            .ThenBy(change => (int)change.Operation)
            .ThenBy(change => ValueSortKey(change.Before), StringComparer.Ordinal)
            .ThenBy(change => ValueSortKey(change.After), StringComparer.Ordinal)
            .ThenBy(change => change.MoveId, StringComparer.Ordinal)
            .Select((change, index) => change with { Id = $"chg-{index + 1:D6}" })
            .ToArray();
    }

    public string Schema => CurrentSchema;
    public int SchemaVersion => CurrentSchemaVersion;
    public int ChangeCount => Changes.Count;
    public IReadOnlyList<SemanticChange> Changes { get; }

    /// <summary>
    /// The compact canonical JSON form used whenever these schema bytes are hashed or signed.
    /// <see cref="ToJson"/> remains the display-oriented serializer.
    /// </summary>
    public string ToCanonicalJson() => ToJson(indented: false);

    /// <summary>UTF-8 bytes of <see cref="ToCanonicalJson"/>.</summary>
    public byte[] ToCanonicalUtf8Bytes() => Encoding.UTF8.GetBytes(ToCanonicalJson());

    /// <summary>Serialize with a fixed field order and invariant number formatting.</summary>
    public string ToJson(bool indented = true)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer, new JsonWriterOptions { Indented = indented }))
            WriteCanonical(writer);

        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    internal void WriteCanonical(Utf8JsonWriter writer)
    {
        ArgumentNullException.ThrowIfNull(writer);
        writer.WriteStartObject();
        writer.WriteString("schema", Schema);
        writer.WriteNumber("schemaVersion", SchemaVersion);
        writer.WriteNumber("changeCount", ChangeCount);
        writer.WriteStartArray("changes");
        foreach (var change in Changes)
            WriteChange(writer, change);
        writer.WriteEndArray();
        writer.WriteEndObject();
    }

    private static void WriteChange(Utf8JsonWriter writer, SemanticChange change)
    {
        writer.WriteStartObject();
        writer.WriteString("id", change.Id);
        writer.WriteString("operation", OperationName(change.Operation));
        writer.WriteString("family", FamilyName(change.Family));
        writer.WriteString("partUri", change.PartUri);
        writer.WriteString("path", change.Path);
        WriteNullableString(writer, "leftAnchor", change.LeftAnchor);
        WriteNullableString(writer, "rightAnchor", change.RightAnchor);
        WriteNullableString(writer, "leftScope", change.LeftScope);
        WriteNullableString(writer, "rightScope", change.RightScope);
        WriteNullableString(writer, "moveId", change.MoveId);
        writer.WritePropertyName("before");
        WriteValue(writer, change.Before);
        writer.WritePropertyName("after");
        WriteValue(writer, change.After);
        writer.WriteEndObject();
    }

    private static void WriteValue(Utf8JsonWriter writer, SemanticValue value)
    {
        writer.WriteStartObject();
        writer.WriteString("kind", ValueKindName(value.Kind));
        switch (value.Kind)
        {
            case SemanticValueKind.Absent:
                break;
            case SemanticValueKind.String:
                writer.WriteString("value", value.StringValue);
                break;
            case SemanticValueKind.Boolean:
                writer.WriteBoolean("value", value.BooleanValue!.Value);
                break;
            case SemanticValueKind.Integer:
                writer.WriteNumber("value", value.IntegerValue!.Value);
                break;
            case SemanticValueKind.Digest:
                writer.WriteString("algorithm", value.DigestAlgorithm);
                WriteNullableString(writer, "profile", value.DigestProfile);
                writer.WriteString("value", value.DigestValue);
                break;
            case SemanticValueKind.Object:
                writer.WriteStartObject("value");
                foreach (var property in value.Properties)
                {
                    writer.WritePropertyName(property.Name);
                    WriteValue(writer, property.Value);
                }
                writer.WriteEndObject();
                break;
            case SemanticValueKind.Array:
                writer.WriteStartArray("value");
                foreach (var item in value.Items)
                    WriteValue(writer, item);
                writer.WriteEndArray();
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(value), value.Kind, "Unknown semantic value kind.");
        }
        writer.WriteEndObject();
    }

    private static void WriteNullableString(Utf8JsonWriter writer, string name, string? value)
    {
        if (value is null) writer.WriteNull(name);
        else writer.WriteString(name, value);
    }

    private static string ValueSortKey(SemanticValue value)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
            WriteValue(writer, value);
        return Encoding.UTF8.GetString(buffer.ToArray());
    }

    internal static string OperationName(SemanticChangeOperation operation) => operation switch
    {
        SemanticChangeOperation.Insert => "insert",
        SemanticChangeOperation.Delete => "delete",
        SemanticChangeOperation.Move => "move",
        SemanticChangeOperation.Modify => "modify",
        _ => throw new ArgumentOutOfRangeException(nameof(operation), operation, null),
    };

    internal static string FamilyName(SemanticChangeFamily family) => family switch
    {
        SemanticChangeFamily.Text => "text",
        SemanticChangeFamily.BlockStructure => "block_structure",
        SemanticChangeFamily.RunFormatting => "run_formatting",
        SemanticChangeFamily.ParagraphFormatting => "paragraph_formatting",
        SemanticChangeFamily.Style => "style",
        SemanticChangeFamily.Numbering => "numbering",
        SemanticChangeFamily.List => "list",
        SemanticChangeFamily.Table => "table",
        SemanticChangeFamily.TableRow => "table_row",
        SemanticChangeFamily.TableCell => "table_cell",
        SemanticChangeFamily.TableSpan => "table_span",
        SemanticChangeFamily.TableWidth => "table_width",
        SemanticChangeFamily.TableStyle => "table_style",
        SemanticChangeFamily.Section => "section",
        SemanticChangeFamily.PageSetup => "page_setup",
        SemanticChangeFamily.Header => "header",
        SemanticChangeFamily.Footer => "footer",
        SemanticChangeFamily.Field => "field",
        SemanticChangeFamily.Footnote => "footnote",
        SemanticChangeFamily.Endnote => "endnote",
        SemanticChangeFamily.Comment => "comment",
        SemanticChangeFamily.Hyperlink => "hyperlink",
        SemanticChangeFamily.Bookmark => "bookmark",
        SemanticChangeFamily.ContentControl => "content_control",
        SemanticChangeFamily.Image => "image",
        SemanticChangeFamily.Media => "media",
        SemanticChangeFamily.Relationship => "relationship",
        SemanticChangeFamily.Revision => "revision",
        SemanticChangeFamily.Annotation => "annotation",
        SemanticChangeFamily.OpaquePackagePart => "opaque_package_part",
        _ => throw new ArgumentOutOfRangeException(nameof(family), family, null),
    };

    private static string ValueKindName(SemanticValueKind kind) => kind switch
    {
        SemanticValueKind.Absent => "absent",
        SemanticValueKind.String => "string",
        SemanticValueKind.Boolean => "boolean",
        SemanticValueKind.Integer => "integer",
        SemanticValueKind.Digest => "digest",
        SemanticValueKind.Object => "object",
        SemanticValueKind.Array => "array",
        _ => throw new ArgumentOutOfRangeException(nameof(kind), kind, null),
    };
}

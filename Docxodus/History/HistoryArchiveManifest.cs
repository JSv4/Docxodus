// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>One pinned portable document history. Integrity is not authentication or a decision-policy audit.</summary>
public sealed record DocxHistoryArchiveInfo(string DocumentId, HistoryHead Head, int BlobCount, long TotalBlobBytes);

internal sealed record HistoryArchiveManifest(string DocumentId, HistoryHead Head, IReadOnlyList<HistoryBlobReference> Blobs)
{
    internal const string Schema = "https://docxodus.dev/schemas/history/archive/v1";
    internal const string EntryName = "history.json";
    internal DocxHistoryArchiveInfo Info => new(DocumentId, Head, Blobs.Count, Blobs.Sum(b => (long)b.Length));
    internal static string BlobName(HistoryBlobReference blob) => "blobs/" + blob.Digest.Value;

    internal byte[] Encode(DocxHistoryArchiveLimits limits)
    {
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject(); writer.WriteString("schema", Schema); writer.WriteNumber("schemaVersion", 1);
            writer.WriteString("documentId", DocumentId); writer.WriteStartObject("head");
            writer.WriteString("revision", Head.Revision.ToString(CultureInfo.InvariantCulture));
            writer.WritePropertyName("state"); Reference(writer, Head.State); writer.WriteEndObject();
            writer.WriteStartArray("blobs");
            foreach (var blob in Blobs)
            {
                Reference(writer, blob); writer.Flush();
                Budget(buffer.Length <= limits.MaxManifestBytes, "Archive manifest exceeds its byte limit.");
            }
            writer.WriteEndArray(); writer.WriteEndObject();
        }
        Budget(buffer.Length <= limits.MaxManifestBytes, "Archive manifest exceeds its byte limit.");
        return buffer.ToArray();
    }

    internal static HistoryArchiveManifest Decode(byte[] bytes, DocxHistoryArchiveLimits limits)
    {
        Budget(bytes.Length <= limits.MaxManifestBytes, "Archive manifest exceeds its byte limit.");
        try
        {
            using var document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 5 });
            var root = Fields(document.RootElement, "schema", "schemaVersion", "documentId", "head", "blobs");
            if (root["schemaVersion"].GetInt32() != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported history archive version.");
            Require(root["schema"].GetString() == Schema, "Unknown archive schema.");
            var id = root["documentId"].GetString()!; HistoryHeadCodec.Key(id);
            var head = Fields(root["head"], "revision", "state"); var revisionText = head["revision"].GetString();
            Require(long.TryParse(revisionText, NumberStyles.None, CultureInfo.InvariantCulture, out var revision)
                && revision > 0 && revision.ToString(CultureInfo.InvariantCulture) == revisionText, "Invalid archive head revision.");
            var captured = new HistoryHead(revision, Reference(head["state"], limits));
            Budget(root["blobs"].GetArrayLength() <= limits.MaxBlobs, "Too many archive blobs.");
            var blobs = new List<HistoryBlobReference>(); string? previous = null; long total = 0;
            foreach (var entry in root["blobs"].EnumerateArray())
            {
                var blob = Reference(entry, limits);
                Require(previous is null || StringComparer.Ordinal.Compare(previous, blob.Digest.Value) < 0,
                    "Archive blob inventory must be unique and sorted.");
                Budget(blob.Length <= limits.MaxTotalBlobBytes - total, "Archive blob bytes exceed their limit.");
                previous = blob.Digest.Value; total += blob.Length; blobs.Add(blob);
            }
            return new(id, captured, blobs.AsReadOnly());
        }
        catch (Exception error) when (error is JsonException or InvalidOperationException or FormatException or ArgumentException or KeyNotFoundException)
        { throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed history archive manifest."); }
    }

    private static Dictionary<string, JsonElement> Fields(JsonElement element, params string[] names)
    {
        var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var property in element.EnumerateObject())
            Require(names.Contains(property.Name, StringComparer.Ordinal) && result.TryAdd(property.Name, property.Value),
                "Unknown or duplicate archive manifest property.");
        Require(result.Count == names.Length, "Missing archive manifest property.");
        return result;
    }

    private static HistoryBlobReference Reference(JsonElement element, DocxHistoryArchiveLimits limits)
    {
        var fields = Fields(element, "sha256", "length");
        var reference = new HistoryBlobReference(new VerificationDigest
        { Algorithm = "SHA-256", Value = fields["sha256"].GetString()! }, fields["length"].GetInt32());
        HistoryBlobIO.Validate(reference, limits.MaxBlobBytes); return reference;
    }
    private static void Reference(Utf8JsonWriter writer, HistoryBlobReference reference)
    {
        writer.WriteStartObject(); writer.WriteString("sha256", reference.Digest.Value);
        writer.WriteNumber("length", reference.Length); writer.WriteEndObject();
    }
    internal static void Require(bool condition, string message)
    { if (!condition) throw new PackageChangeException(PackageChangeError.InvalidManifest, message); }
    internal static void Budget(bool condition, string message)
    { if (!condition) throw new PackageChangeException(PackageChangeError.ResourceLimit, message); }
}

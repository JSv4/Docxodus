// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>
/// Versioned, bounded immutable history metadata over host-owned blobs. Validates record structure
/// and blob integrity, not referenced-graph existence, cross-record consistency, or authorization.
/// Persist records and their referenced graphs before atomically publishing a head.
/// </summary>
public sealed class HistoryRecordStore
{
    public const int DefaultMaxRecordBytes = 256 * 1024;
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        AllowDuplicateProperties = false,
        RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true,
        MaxDepth = 12,
    };
    private static readonly HistoryRecordsJsonContext JsonContext = new(JsonOptions);
    private readonly IHistoryBlobStore _blobs;
    private readonly int _maxRecordBytes;

    public HistoryRecordStore(IHistoryBlobStore blobs, int maxRecordBytes = DefaultMaxRecordBytes)
    {
        ArgumentNullException.ThrowIfNull(blobs);
        _blobs = blobs;
        _maxRecordBytes = HistoryBlobIO.ValidateLimit(maxRecordBytes);
    }

    public ValueTask<HistoryBlobReference> SaveVersionAsync(DocxVersionRecord record, CancellationToken cancellationToken = default) =>
        SaveAsync(record, "version", cancellationToken);
    public ValueTask<HistoryBlobReference> SaveCommitAsync(PackageHistoryCommitRecord record, CancellationToken cancellationToken = default) =>
        SaveAsync(record, "package-commit", cancellationToken);
    public ValueTask<HistoryBlobReference> SaveStateAsync(DocxHistoryStateRecord record, CancellationToken cancellationToken = default) =>
        SaveAsync(record, "package-state", cancellationToken);
    public ValueTask<DocxVersionRecord> LoadVersionAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default) =>
        LoadAsync<DocxVersionRecord>(reference, "version", cancellationToken);
    public ValueTask<PackageHistoryCommitRecord> LoadCommitAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default) =>
        LoadAsync<PackageHistoryCommitRecord>(reference, "package-commit", cancellationToken);
    public ValueTask<DocxHistoryStateRecord> LoadStateAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default) =>
        LoadAsync<DocxHistoryStateRecord>(reference, "package-state", cancellationToken);

    private async ValueTask<HistoryBlobReference> SaveAsync<T>(T record, string kind, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var normalized = Validate(record);
        var schemaVersion = normalized is DocxHistoryStateRecord { Operation: not null } ? 3
            : normalized is DocxHistoryStateRecord { Requests: not null } ? 2 : 1;
        var bytes = JsonSerializer.SerializeToUtf8Bytes(new HistoryRecordEnvelope<T>
        {
            Record = normalized, Schema = Schema(kind, schemaVersion), SchemaVersion = schemaVersion,
        }, TypeInfo<T>());
        Budget(bytes.Length <= _maxRecordBytes);
        var reference = new HistoryBlobReference(new VerificationDigest
        {
            Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
        }, bytes.Length);
        using var content = new MemoryStream(bytes, writable: false);
        await _blobs.PutAsync(reference, content, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return reference;
    }

    private async ValueTask<T> LoadAsync<T>(HistoryBlobReference reference, string kind, CancellationToken cancellationToken)
    {
        HistoryBlobIO.Validate(reference, _maxRecordBytes);
        cancellationToken.ThrowIfCancellationRequested();
        using var content = await _blobs.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
        if (content is null) throw new PackageChangeException(PackageChangeError.PayloadMissing, "History record is missing.");
        var bytes = await PackageChangeSetCodec.ReadPayloadAsync(reference, content, cancellationToken).ConfigureAwait(false);
        try
        {
            using var header = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 12, AllowDuplicateProperties = false });
            var fields = header.RootElement.EnumerateObject().Select(p => p.Name).ToHashSet(StringComparer.Ordinal);
            Require(fields.SetEquals(["record", "schema", "schemaVersion"]), "Malformed history envelope.");
            var schemaVersion = header.RootElement.GetProperty("schemaVersion").GetInt32();
            if (schemaVersion != 1 && !(schemaVersion is 2 or 3 && kind == "package-state"))
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported history record version.");
            Require(header.RootElement.GetProperty("schema").GetString() == Schema(kind, schemaVersion), "Unexpected history record schema.");
            var envelope = JsonSerializer.Deserialize(bytes, TypeInfo<T>());
            Require(envelope is not null, "History envelope is required.");
            var record = Validate(envelope!.Record);
            if (record is DocxHistoryStateRecord state)
            {
                Require(schemaVersion != 2 || state.Requests is not null, "V2 states require request journals.");
                Require(schemaVersion != 1 || !header.RootElement.GetProperty("record").TryGetProperty("requests", out _),
                    "V1 states cannot declare request journals, including null.");
                Require((state.Operation is not null) == (schemaVersion == 3), "Backend tips require state schema V3.");
                Require(schemaVersion == 3 || (!header.RootElement.GetProperty("record").TryGetProperty("operation", out _)
                    && !header.RootElement.GetProperty("record").TryGetProperty("parentPublication", out _)),
                    "Older states cannot declare V3 publication fields, including null.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            return record;
        }
        catch (Exception error) when (error is JsonException or ArgumentException or InvalidOperationException or FormatException)
        {
            throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed history record.");
        }
    }

    private T Validate<T>(T record)
    {
        ArgumentNullException.ThrowIfNull(record);
        switch (record)
        {
            case DocxVersionRecord version:
                HistoryHeadCodec.Key(version.DocumentId);
                Require(version.Nonce != Guid.Empty && version.Sequence >= 0, "Invalid version identity or sequence.");
                Reference(version.Parent); Reference(version.RestoredFrom); Snapshot(version.Snapshot);
                Require(version.RestoredFrom is null || version.Parent is not null, "A restore requires a parent version.");
                return (T)(object)(version with { Metadata = PrepareMetadata(version.Metadata) });
            case PackageHistoryCommitRecord commit:
                HistoryHeadCodec.Key(commit.DocumentId);
                Require(commit.Sequence > 0 && commit.Epoch >= 0
                    && (commit.Parent is null) == (commit.Sequence == 1), "Invalid commit position.");
                Reference(commit.Parent); RequiredReference(commit.Version);
                Snapshot(commit.Before); Snapshot(commit.After);
                Require(commit.Kind is "import" or "restore", "Unknown package commit kind.");
                Require((commit.Contribution is not null) == (commit.Kind == "import"), "Invalid contribution for commit kind.");
                if (commit.Contribution is not null)
                {
                    HistoryBlobIO.Validate(commit.Contribution, new PackageChangeLimits().MaxManifestBytes);
                    Require(commit.Before.ContentDigest != commit.After.ContentDigest, "An import commit must change package content.");
                }
                if (commit.Kind == "restore") Require(commit.Epoch > 0, "A restore advances the epoch.");
                break;
            case DocxHistoryStateRecord state:
                HistoryHeadCodec.Key(state.DocumentId);
                HistoryRequestJournalStore.ValidateJournal(state.Requests);
                Require(state.Requests is null || state.Requests.DocumentId == state.DocumentId, "Foreign request journal.");
                Require((state.Operation is null) == (state.ParentPublication is null), "Backend tip and publication parent must occur together.");
                Reference(state.Operation);
                if (state.ParentPublication is not null)
                {
                    Require(state.ParentPublication.Revision > 0, "Invalid parent publication revision.");
                    RequiredReference(state.ParentPublication.State);
                }
                Require(state.Sequence >= 0 && state.Epoch >= 0
                    && (state.Commit is null) == (state.Sequence == 0), "Invalid history state position.");
                Reference(state.Commit); RequiredReference(state.Version);
                Snapshot(state.Snapshot); Snapshot(state.InitialSnapshot);
                Require(state.Sequence != 0 || (state.Epoch == 0 && state.Snapshot.ContentDigest == state.InitialSnapshot.ContentDigest),
                    "Initial state must match its initial snapshot.");
                break;
            default: throw new ArgumentException("Unsupported history record type.", nameof(record));
        }
        return record;
    }

    internal static DocxVersionMetadata PrepareMetadata(DocxVersionMetadata metadata)
    {
        ArgumentNullException.ThrowIfNull(metadata);
        Text(metadata.Author, 1024, required: true);
        Text(metadata.Label, 4096); Text(metadata.Message, 16 * 1024);
        ArgumentNullException.ThrowIfNull(metadata.ApplicationMetadata);
        Budget(metadata.ApplicationMetadata.Count <= 64);
        var values = new SortedDictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in metadata.ApplicationMetadata)
        {
            Text(pair.Key, 256, required: true); Text(pair.Value, 4096, required: false);
            Require(pair.Value is not null, "Metadata values cannot be null.");
            values.Add(pair.Key, pair.Value!);
        }
        return metadata with { ApplicationMetadata = new ReadOnlyDictionary<string, string>(values) };
    }

    private void Reference(HistoryBlobReference? reference)
    {
        if (reference is not null) HistoryBlobIO.Validate(reference, _maxRecordBytes);
    }
    private void RequiredReference(HistoryBlobReference reference)
    {
        ArgumentNullException.ThrowIfNull(reference);
        Reference(reference);
    }
    private static void Snapshot(DocxSnapshotReference snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        HistoryBlobIO.Validate(snapshot.Blob, int.MaxValue);
        HistoryBlobIO.Validate(new HistoryBlobReference(snapshot.ContentDigest, 0), int.MaxValue);
    }
    private static void Text(string? value, int maximumLength, bool required = false)
    {
        Require(!required || !string.IsNullOrWhiteSpace(value), "Required metadata text is missing.");
        if (value is null) return;
        Budget(value.Length <= maximumLength);
        _ = new UTF8Encoding(false, true).GetByteCount(value);
    }
    private static string Schema(string kind, int version) => $"https://docxodus.dev/schemas/history/{kind}/v{version}";
    private static JsonTypeInfo<HistoryRecordEnvelope<T>> TypeInfo<T>() =>
        (JsonTypeInfo<HistoryRecordEnvelope<T>>)JsonContext.GetTypeInfo(typeof(HistoryRecordEnvelope<T>))!;
    private static void Require(bool condition, string message)
    {
        if (!condition) throw new PackageChangeException(PackageChangeError.InvalidManifest, message);
    }
    private static void Budget(bool condition)
    {
        if (!condition) throw new PackageChangeException(PackageChangeError.ResourceLimit, "History record exceeds its metadata limit.");
    }
}

internal sealed record HistoryRecordEnvelope<T>
{
    public required T Record { get; init; }
    public required string Schema { get; init; }
    public required int SchemaVersion { get; init; }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, AllowDuplicateProperties = false,
    RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true, MaxDepth = 12)]
[JsonSerializable(typeof(HistoryRecordEnvelope<DocxVersionRecord>))]
[JsonSerializable(typeof(HistoryRecordEnvelope<PackageHistoryCommitRecord>))]
[JsonSerializable(typeof(HistoryRecordEnvelope<DocxHistoryStateRecord>))]
internal partial class HistoryRecordsJsonContext : JsonSerializerContext
{
}

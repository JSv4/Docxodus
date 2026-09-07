// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using Docxodus.History;

namespace Docxodus.Internal;

/// <summary>
/// Shared, reflection-free client boundary for WASM and local agent bindings. Does not own a
/// transport, storage, authorization, or a session. Binary inputs are separate from JSON; binary
/// results use JSON base64. These are package-boundary APIs, not a keystroke recorder.
/// </summary>
public sealed class HistoryClientOps : IDisposable
{
    private readonly DocxVersionHistory _history;
    private readonly DocxHistoryArchive? _archive;
    private bool _disposed;
    public const int MaxArchiveBytes = 64 * 1024 * 1024;
    internal static readonly DocxHistoryArchiveLimits ArchiveLimits = new()
    {
        MaxArchiveBytes = MaxArchiveBytes, MaxBlobBytes = 64 * 1024 * 1024, MaxSnapshotBytes = 64 * 1024 * 1024,
        MaxTotalBlobBytes = 256L * 1024 * 1024, MaxMetadataBytes = 16L * 1024 * 1024,
        MaxManifestBytes = 4 * 1024 * 1024, MaxValidationBytes = 512L * 1024 * 1024,
        MaxExpandedBytes = 2L * 1024 * 1024 * 1024, MaxBlobs = 10_000, MaxEdges = 100_000,
    };
    public DocxHistoryArchiveInfo? ArchiveInfo => _archive?.Info;

    public HistoryClientOps(IHistoryBlobStore blobs, IHistoryHeadStore heads) =>
        _history = new DocxVersionHistory(blobs, heads);

    private HistoryClientOps(DocxHistoryArchive archive) { _archive = archive; _history = archive.Core; }

    /// <summary>Byte bindings have a 64 MiB archive cap; native stream APIs support host-selected larger limits.</summary>
    public static async ValueTask<HistoryClientOps> OpenArchiveAsync(byte[] bytes, CancellationToken cancellationToken = default) =>
        new(await DocxHistoryArchive.OpenAsync(bytes, ArchiveLimits, cancellationToken: cancellationToken).ConfigureAwait(false));

    public void Dispose() { if (_disposed) return; _archive?.Dispose(); _disposed = true; }

    public async Task<string> InvokeAsync(string requestJson, byte[]? docxBytes = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            cancellationToken.ThrowIfCancellationRequested();
            var request = HistoryClientJson.Read<HistoryClientRequest>(requestJson);
            if (request.SchemaVersion != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported history client version.");
            var id = request.DocumentId;
            if (_archive is not null)
            {
                if (id != _archive.DocumentId) throw new DocxHistoryException(DocxHistoryError.ForeignDocument, "Archive belongs to another document.");
                if (request.Operation is "create" or "restore" or "importArchive")
                    return HistoryClientJson.Write(new HistoryClientResult { Success = false, ErrorCode = "ReadOnly", Message = "Archive is read-only; import into host-owned storage to edit." });
            }
            var budget = request.MaxEntriesToScan;
            var result = request.Operation switch
            {
                "read" => new HistoryClientResult { View = await _history.ReadAsync(id, cancellationToken).ConfigureAwait(false) },
                "updates" => new HistoryClientResult { Update = await _history.ReadChangesSinceAsync(id, request.ExpectedHead,
                    budget, cancellationToken).ConfigureAwait(false) },
                "operations" => new HistoryClientResult { OperationUpdate = await _history.ReadOperationsSinceAsync(id, request.ExpectedHead,
                    budget, cancellationToken).ConfigureAwait(false) },
                "getOperation" => new HistoryClientResult { Operation = await _history.GetOperationAsync(id, Required(request.OperationId),
                    cancellationToken).ConfigureAwait(false) },
                "exportOperationProposal" => new HistoryClientResult { Bytes = await _history.ExportOperationProposalAsync(id, Required(request.OperationId),
                    cancellationToken).ConfigureAwait(false) },
                "create" => new HistoryClientResult { View = request.RequestId is null
                    ? await _history.CreateVersionAsync(id, request.ExpectedHead,
                        Required(docxBytes), Required(request.Metadata), cancellationToken).ConfigureAwait(false)
                    : await _history.CreateVersionAsync(id, request.RequestId, request.ExpectedHead,
                        Required(docxBytes), Required(request.Metadata), cancellationToken).ConfigureAwait(false) },
                "list" => new HistoryClientResult { Page = await _history.ListVersionsAsync(id, request.VersionId,
                    request.Limit, cancellationToken).ConfigureAwait(false) },
                "get" => new HistoryClientResult { Version = await _history.GetVersionAsync(id, Required(request.VersionId),
                    cancellationToken).ConfigureAwait(false) },
                "export" => new HistoryClientResult { Bytes = await _history.ExportVersionAsync(id, Required(request.VersionId),
                    cancellationToken).ConfigureAwait(false) },
                "exportDocx" => new HistoryClientResult { Bytes = await _history.Document(id).ExportDocxAsync(request.VersionId,
                    cancellationToken).ConfigureAwait(false) },
                "exportArchive" => await ExportArchiveAsync(id, cancellationToken).ConfigureAwait(false),
                "compare" => await CompareAsync(id, Required(request.BeforeVersionId), Required(request.AfterVersionId), cancellationToken).ConfigureAwait(false),
                "importArchive" => new HistoryClientResult { Import = await _history.ImportScopedArchiveAsync(id.Length == 0 ? null : id,
                    Required(docxBytes), ArchiveLimits, cancellationToken).ConfigureAwait(false) },
                "materialize" => new HistoryClientResult { Bytes = await _history.MaterializeAsync(id, Required(request.Sequence),
                    budget, cancellationToken).ConfigureAwait(false) },
                "replay" => new HistoryClientResult { Bytes = await _history.ReplayAsync(id, Required(request.Sequence),
                    budget, cancellationToken).ConfigureAwait(false) },
                "resolveTime" => new HistoryClientResult { Sequence = await _history.ResolveSequenceAtTimeAsync(id,
                    Required(request.Cutoff), budget, cancellationToken).ConfigureAwait(false) },
                "restore" => new HistoryClientResult { View = request.RequestId is null
                    ? await _history.RestoreVersionAsync(id, Required(request.ExpectedHead),
                        Required(request.VersionId), Required(request.Metadata), cancellationToken).ConfigureAwait(false)
                    : await _history.RestoreVersionAsync(id, request.RequestId, Required(request.ExpectedHead),
                        Required(request.VersionId), Required(request.Metadata), cancellationToken).ConfigureAwait(false) },
                _ => throw new ArgumentException("Unknown history operation."),
            };
            return HistoryClientJson.Write(result);
        }
        catch (Exception error) when (IsClientError(error)) { return Failure(error); }
    }

    private async Task<HistoryClientResult> ExportArchiveAsync(string id, CancellationToken ct)
    {
        using var output = new MemoryStream();
        var info = await _history.ExportHistoryArchiveAsync(id, output, ArchiveLimits, ct).ConfigureAwait(false);
        return new HistoryClientResult { Archive = info, Bytes = output.ToArray() };
    }

    private async Task<HistoryClientResult> CompareAsync(string id, HistoryBlobReference before, HistoryBlobReference after, CancellationToken ct)
    {
        return new HistoryClientResult { Bytes = await _history.Document(id).CompareVersionsToDocxAsync(before, after,
            cancellationToken: ct).ConfigureAwait(false) };
    }

    internal static bool IsClientError(Exception error) => error is DocxHistoryException or PackageChangeException or JsonException
        or ArgumentException or OperationCanceledException or FormatException or ObjectDisposedException;

    internal static string Failure(Exception error)
    {
        var code = error switch
        {
            DocxHistoryException history => history.Code.ToString(),
            PackageChangeException package => package.Code.ToString(),
            OperationCanceledException => "Canceled",
            ObjectDisposedException => "Closed",
            _ => "InvalidRequest",
        };
        return HistoryClientJson.Write(new HistoryClientResult { Success = false, ErrorCode = code, Message = error.Message });
    }

    private static T Required<T>(T? value) where T : class => value ?? throw new ArgumentException("Required history argument is missing.");
    private static T Required<T>(T? value) where T : struct => value ?? throw new ArgumentException("Required history argument is missing.");
}

public sealed record HistoryClientRequest
{
    public required int SchemaVersion { get; init; }
    public required string Operation { get; init; }
    public required string DocumentId { get; init; }
    /// <summary>Optional durable create/restore identity. Null/absent preserves the legacy contract.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? RequestId { get; init; }
    public HistoryHead? ExpectedHead { get; init; }
    public HistoryBlobReference? VersionId { get; init; }
    public HistoryBlobReference? OperationId { get; init; }
    public HistoryBlobReference? BeforeVersionId { get; init; }
    public HistoryBlobReference? AfterVersionId { get; init; }
    public DocxVersionMetadata? Metadata { get; init; }
    public long? Sequence { get; init; }
    public DateTimeOffset? Cutoff { get; init; }
    public int Limit { get; init; } = 25;
    public int MaxEntriesToScan { get; init; } = 10_000;
}

public sealed record HistoryClientResult
{
    public int? Handle { get; init; }
    public DocxHistoryArchiveInfo? Archive { get; init; }
    public DocxHistoryImportResult? Import { get; init; }
    public bool Success { get; init; } = true;
    public string? ErrorCode { get; init; }
    public string? Message { get; init; }
    public DocxHistoryView? View { get; init; }
    public DocxStoredVersion? Version { get; init; }
    public DocxVersionPage? Page { get; init; }
    public long? Sequence { get; init; }
    public byte[]? Bytes { get; init; }
    public DocxHistoryUpdate? Update { get; init; }
    public DocxOperationUpdate? OperationUpdate { get; init; }
    public DocxStoredOperation? Operation { get; init; }
}

/// <summary>Client JSON only: 64-bit positions are decimal strings, preserving precision in JS.</summary>
public static class HistoryClientJson
{
    public const int MaxRequestChars = 512 * 1024;
    private static readonly HistoryClientJsonContext Context = new(new JsonSerializerOptions
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        AllowDuplicateProperties = false, RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true, MaxDepth = 16,
        Converters = { new HistoryInt64JsonConverter() },
    });

    public static string Write<T>(T value) => JsonSerializer.Serialize(value, Info<T>());
    public static T Read<T>(string json)
    {
        ArgumentNullException.ThrowIfNull(json);
        if (json.Length > MaxRequestChars) throw new ArgumentException("History client metadata exceeds its limit.");
        var value = JsonSerializer.Deserialize(json, Info<T>()) ?? throw new JsonException("History metadata is null.");
        if (value is HistoryClientRequest request)
        {
            // Generated init-only construction can supply default(T) for omitted optional
            // properties instead of preserving their C# initializer. Apply wire defaults by
            // presence, never by value, so explicit null/zero remains invalid downstream.
            using var document = JsonDocument.Parse(json);
            var root = document.RootElement;
            var metadata = request.Metadata;
            if (metadata is not null && !root.GetProperty("metadata").TryGetProperty("applicationMetadata", out _))
                metadata = metadata with { ApplicationMetadata = new Dictionary<string, string>() };
            value = (T)(object)(request with
            {
                Metadata = metadata,
                Limit = root.TryGetProperty("limit", out _) ? request.Limit : 25,
                MaxEntriesToScan = root.TryGetProperty("maxEntriesToScan", out _) ? request.MaxEntriesToScan : 10_000,
            });
        }
        return value;
    }
    private static JsonTypeInfo<T> Info<T>() => (JsonTypeInfo<T>)(Context.GetTypeInfo(typeof(T))
        ?? throw new ArgumentException("Unsupported history client JSON type."));
}

internal sealed class HistoryInt64JsonConverter : JsonConverter<long>
{
    public override long Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        if (reader.TokenType == JsonTokenType.String && long.TryParse(reader.GetString(), NumberStyles.None,
            CultureInfo.InvariantCulture, out var value) && value.ToString(CultureInfo.InvariantCulture) == reader.GetString()) return value;
        throw new JsonException("History positions must be canonical nonnegative decimal strings.");
    }
    public override void Write(Utf8JsonWriter writer, long value, JsonSerializerOptions options) =>
        writer.WriteStringValue(value.ToString(CultureInfo.InvariantCulture));
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, AllowDuplicateProperties = false,
    RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true, MaxDepth = 16)]
[JsonSerializable(typeof(HistoryClientRequest))]
[JsonSerializable(typeof(HistoryClientResult))]
[JsonSerializable(typeof(HistoryHead))]
[JsonSerializable(typeof(HistoryBlobReference))]
[JsonSerializable(typeof(HistoryHeadInitializationResult))]
internal partial class HistoryClientJsonContext : JsonSerializerContext { }

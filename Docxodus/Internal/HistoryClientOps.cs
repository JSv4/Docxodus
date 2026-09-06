// Copyright (c) Microsoft. All rights reserved.
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
public sealed class HistoryClientOps
{
    private readonly DocxVersionHistory _history;

    public HistoryClientOps(IHistoryBlobStore blobs, IHistoryHeadStore heads) =>
        _history = new DocxVersionHistory(blobs, heads);

    public async Task<string> InvokeAsync(string requestJson, byte[]? docxBytes = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var request = HistoryClientJson.Read<HistoryClientRequest>(requestJson);
            if (request.SchemaVersion != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported history client version.");
            var id = request.DocumentId;
            var budget = request.MaxEntriesToScan;
            var result = request.Operation switch
            {
                "read" => new HistoryClientResult { View = await _history.ReadAsync(id, cancellationToken).ConfigureAwait(false) },
                "create" => new HistoryClientResult { View = await _history.CreateVersionAsync(id, request.ExpectedHead,
                    Required(docxBytes), Required(request.Metadata), cancellationToken).ConfigureAwait(false) },
                "list" => new HistoryClientResult { Page = await _history.ListVersionsAsync(id, request.VersionId,
                    request.Limit, cancellationToken).ConfigureAwait(false) },
                "get" => new HistoryClientResult { Version = await _history.GetVersionAsync(id, Required(request.VersionId),
                    cancellationToken).ConfigureAwait(false) },
                "export" => new HistoryClientResult { Bytes = await _history.ExportVersionAsync(id, Required(request.VersionId),
                    cancellationToken).ConfigureAwait(false) },
                "materialize" => new HistoryClientResult { Bytes = await _history.MaterializeAsync(id, Required(request.Sequence),
                    budget, cancellationToken).ConfigureAwait(false) },
                "replay" => new HistoryClientResult { Bytes = await _history.ReplayAsync(id, Required(request.Sequence),
                    budget, cancellationToken).ConfigureAwait(false) },
                "resolveTime" => new HistoryClientResult { Sequence = await _history.ResolveSequenceAtTimeAsync(id,
                    Required(request.Cutoff), budget, cancellationToken).ConfigureAwait(false) },
                "restore" => new HistoryClientResult { View = await _history.RestoreVersionAsync(id, Required(request.ExpectedHead),
                    Required(request.VersionId), Required(request.Metadata), cancellationToken).ConfigureAwait(false) },
                _ => throw new ArgumentException("Unknown history operation."),
            };
            return HistoryClientJson.Write(result);
        }
        catch (Exception error) when (error is DocxHistoryException or PackageChangeException or JsonException
            or ArgumentException or OperationCanceledException or FormatException)
        {
            var code = error switch
            {
                DocxHistoryException history => history.Code.ToString(),
                PackageChangeException package => package.Code.ToString(),
                OperationCanceledException => "Canceled",
                _ => "InvalidRequest",
            };
            return HistoryClientJson.Write(new HistoryClientResult { Success = false, ErrorCode = code, Message = error.Message });
        }
    }

    private static T Required<T>(T? value) where T : class => value ?? throw new ArgumentException("Required history argument is missing.");
    private static T Required<T>(T? value) where T : struct => value ?? throw new ArgumentException("Required history argument is missing.");
}

public sealed record HistoryClientRequest
{
    public required int SchemaVersion { get; init; }
    public required string Operation { get; init; }
    public required string DocumentId { get; init; }
    public HistoryHead? ExpectedHead { get; init; }
    public HistoryBlobReference? VersionId { get; init; }
    public DocxVersionMetadata? Metadata { get; init; }
    public long? Sequence { get; init; }
    public DateTimeOffset? Cutoff { get; init; }
    public int Limit { get; init; } = 25;
    public int MaxEntriesToScan { get; init; } = 10_000;
}

public sealed record HistoryClientResult
{
    public bool Success { get; init; } = true;
    public string? ErrorCode { get; init; }
    public string? Message { get; init; }
    public DocxHistoryView? View { get; init; }
    public DocxStoredVersion? Version { get; init; }
    public DocxVersionPage? Page { get; init; }
    public long? Sequence { get; init; }
    public byte[]? Bytes { get; init; }
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
internal partial class HistoryClientJsonContext : JsonSerializerContext { }

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.Json.Serialization.Metadata;
using System.Xml;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Bounded, strict, generated codecs for the backend intent/decision vocabulary.</summary>
internal sealed class DocxOperationStore(IHistoryBlobStore blobs)
{
    internal const int MaxBytes = 512 * 1024;
    private static readonly DocxOperationJsonContext Json = new(new JsonSerializerOptions
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, AllowDuplicateProperties = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true, MaxDepth = 16,
    });

    internal static DocxOperationInput Capture(string documentId, DocxOperationRequest request, byte[]? candidate)
    {
        ArgumentNullException.ThrowIfNull(request);
        return Validate(new DocxOperationInput(documentId, request with
        {
            Metadata = HistoryRecordStore.PrepareMetadata(request.Metadata),
            ReadParts = Array.AsReadOnly((request.ReadParts ?? throw new ArgumentNullException(nameof(request.ReadParts)))
                .OrderBy(x => x, StringComparer.Ordinal).Distinct(StringComparer.Ordinal).ToArray()),
        }, candidate is null ? null : Reference(candidate)));
    }

    internal static byte[] EncodeInput(DocxOperationInput input) => Encode(Validate(input), "operation-input");
    internal ValueTask<HistoryBlobReference> SaveDecisionAsync(DocxOperationRecord record, CancellationToken cancellationToken) =>
        HistoryBlobIO.PutBytesAsync(blobs, Encode(Validate(record), "operation-decision"), cancellationToken);
    internal ValueTask<DocxOperationInput> LoadInputAsync(HistoryBlobReference reference, CancellationToken cancellationToken) =>
        LoadAsync<DocxOperationInput>(reference, "operation-input", cancellationToken);
    internal ValueTask<DocxOperationRecord> LoadDecisionAsync(HistoryBlobReference reference, CancellationToken cancellationToken) =>
        LoadAsync<DocxOperationRecord>(reference, "operation-decision", cancellationToken);

    internal static HistoryBlobReference Reference(byte[] bytes) => new(new VerificationDigest
    { Algorithm = "SHA-256", Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() }, bytes.Length);

    private static byte[] Encode<T>(T value, string kind)
    {
        var result = JsonSerializer.SerializeToUtf8Bytes(new HistoryRecordEnvelope<T>
        { Record = value, SchemaVersion = 1, Schema = Schema(kind) }, TypeInfo<T>());
        Require(result.Length <= MaxBytes, "Operation metadata exceeds its byte limit.");
        return result;
    }

    private async ValueTask<T> LoadAsync<T>(HistoryBlobReference reference, string kind, CancellationToken cancellationToken)
    {
        var bytes = await HistoryBlobIO.ReadBytesAsync(blobs, reference, MaxBytes, cancellationToken).ConfigureAwait(false);
        try
        {
            var envelope = JsonSerializer.Deserialize(bytes, TypeInfo<T>())
                ?? throw new JsonException("Missing operation envelope.");
            if (envelope.SchemaVersion != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported operation codec version.");
            Require(envelope.Schema == Schema(kind), "Unexpected operation schema.");
            var validated = envelope.Record switch
            {
                DocxOperationInput input => (T)(object)Validate(input),
                DocxOperationRecord decision => (T)(object)Validate(decision),
                _ => throw new JsonException("Unexpected operation record."),
            };
            cancellationToken.ThrowIfCancellationRequested();
            return validated;
        }
        catch (Exception error) when (error is JsonException or ArgumentException or XmlException or OverflowException)
        { throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed operation record."); }
    }

    private static DocxOperationInput Validate(DocxOperationInput input)
    {
        HistoryHeadCodec.Key(input.DocumentId);
        var request = input.Request ?? throw new ArgumentNullException(nameof(input.Request));
        HistoryRequestJournalStore.ValidateId(request.RequestId);
        Head(request.Base);
        Require(request.Kind is "text" or "package" or "discard", "Unknown operation kind.");
        Require((request.Kind == "text") == (request.Text is not null), "Text payload disagrees with operation kind.");
        Require((request.Kind == "package") == (input.Candidate is not null), "Candidate bytes disagree with operation kind.");
        Require(request.Kind != "discard" || request.Resolves is not null, "Discard must resolve a published conflict.");
        if (request.Text is not null) Text(request.Text);
        if (input.Candidate is not null) HistoryBlobIO.Validate(input.Candidate, int.MaxValue);
        if (request.Resolves is not null) HistoryBlobIO.Validate(request.Resolves, MaxBytes);
        Require(request.ReadParts is not null && request.ReadParts.Count <= 256, "Too many read dependencies.");
        string? previous = null;
        foreach (var uri in request.ReadParts!)
        {
            Part(uri);
            Require(previous is null || StringComparer.Ordinal.Compare(previous, uri) < 0,
                "Read dependencies must be unique and ordinally sorted.");
            previous = uri;
        }
        return input with { Request = request with
        {
            Metadata = HistoryRecordStore.PrepareMetadata(request.Metadata),
            ReadParts = Array.AsReadOnly(request.ReadParts.ToArray()),
        } };
    }

    private static DocxOperationRecord Validate(DocxOperationRecord record)
    {
        HistoryHeadCodec.Key(record.DocumentId); Head(record.Before);
        Require(record.Revision > 1 && record.Revision - 1 == record.Before.Revision, "Invalid decision revision.");
        HistoryBlobIO.Validate(record.Input, MaxBytes); HistoryBlobIO.Validate(record.Version, int.MaxValue);
        if (record.Parent is not null) HistoryBlobIO.Validate(record.Parent, MaxBytes);
        if (record.ContentCommit is not null) HistoryBlobIO.Validate(record.ContentCommit, int.MaxValue);
        Snapshot(record.ProposedSnapshot); Snapshot(record.AfterSnapshot);
        Require(record.Status is "accepted" or "conflict", "Unknown decision status.");
        Require((record.Status == "conflict") == (record.Conflict is not null), "Conflict reason disagrees with status.");
        Require(record.Conflict is null or "OverlappingText" or "ChangedPart" or "ReadDependencyChanged"
            or "EpochChanged" or "AlreadyResolved" or "UnknownTextChange", "Unknown conflict reason.");
        Require(record.Status != "conflict" || (record.ContentCommit is null && record.AppliedText is null),
            "Conflicting decisions cannot have applied effects.");
        Require(record.AppliedText is null || record.ContentCommit is not null, "Mapped text requires a content commit.");
        if (record.AppliedText is not null) Text(record.AppliedText);
        return record;
    }

    internal static void Part(string uri)
    {
        Require(uri is not null && uri.Length is > 1 and <= 4096 && uri[0] == '/'
            && PackageManifestGenerator.TryCanonicalizeEntryName(uri[1..], out var canonical) && canonical == uri,
            "Expected a canonical absolute OPC part URI.");
    }
    internal static void Text(DocxTextSplice text)
    {
        Part(text.PartUri);
        Require(text.TextNode >= 0 && text.Offset >= 0 && text.DeleteCount >= 0
            && (long)text.Offset + text.DeleteCount <= int.MaxValue && text.Insert is not null
            && text.Insert.Length <= 64 * 1024, "Invalid text splice or payload limit.");
        XmlConvert.VerifyXmlChars(text.Insert!);
    }
    private static void Head(HistoryHead head)
    {
        ArgumentNullException.ThrowIfNull(head); Require(head.Revision > 0, "Invalid base head.");
        HistoryBlobIO.Validate(head.State, int.MaxValue);
    }
    private static void Snapshot(DocxSnapshotReference snapshot)
    {
        ArgumentNullException.ThrowIfNull(snapshot); HistoryBlobIO.Validate(snapshot.Blob, int.MaxValue);
        HistoryBlobIO.Validate(new HistoryBlobReference(snapshot.ContentDigest, 0), 0);
    }
    private static void Require(bool condition, string message)
    { if (!condition) throw new PackageChangeException(PackageChangeError.InvalidManifest, message); }
    private static string Schema(string kind) => $"https://docxodus.dev/schemas/history/{kind}/v1";
    private static JsonTypeInfo<HistoryRecordEnvelope<T>> TypeInfo<T>() =>
        (JsonTypeInfo<HistoryRecordEnvelope<T>>)Json.GetTypeInfo(typeof(HistoryRecordEnvelope<T>))!;
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, AllowDuplicateProperties = false,
    RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true, MaxDepth = 16)]
[JsonSerializable(typeof(HistoryRecordEnvelope<DocxOperationInput>))]
[JsonSerializable(typeof(HistoryRecordEnvelope<DocxOperationRecord>))]
internal partial class DocxOperationJsonContext : JsonSerializerContext { }

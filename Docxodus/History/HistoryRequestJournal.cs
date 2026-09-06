// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>An opaque, document-scoped logical request ID bound to its canonical input digest.</summary>
public sealed record HistoryRequestIdentity(string Id, VerificationDigest Fingerprint);

/// <summary>
/// Atomically publish this value with application state. Current's result is that publication's
/// own head; Index contains earlier receipts. The next publication indexes the now-known head,
/// avoiding a content-hash cycle. Retain the reachable index and original result states.
/// </summary>
public sealed record HistoryRequestJournal(string DocumentId, long Revision,
    HistoryBlobReference? Index, HistoryRequestIdentity? Current);

/// <summary>A durable binding from a logical request to its exact original publication.</summary>
public sealed record HistoryRequestReceipt(string DocumentId, HistoryRequestIdentity Request, HistoryHead Head);

/// <summary>
/// Storage-neutral durable retry lookup. The caller publishes AdvanceAsync's result in the SAME
/// head CAS as its application state. This class writes immutable blobs only; it never publishes
/// independently. No authoritative in-memory cache, transport, clocks, or actor allocator.
/// </summary>
public sealed class HistoryRequestJournalStore
{
    public const int MaxRequestIdChars = 1024;
    public const int MaxNodeBytes = 16 * 1024;
    private static readonly UTF8Encoding Utf8 = new(false, true);
    private static readonly HistoryRequestJsonContext Json = new(new JsonSerializerOptions
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, AllowDuplicateProperties = false,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, RespectNullableAnnotations = true,
        RespectRequiredConstructorParameters = true, MaxDepth = 12,
    });
    private readonly IHistoryBlobStore _blobs;

    public HistoryRequestJournalStore(IHistoryBlobStore blobs) =>
        _blobs = blobs ?? throw new ArgumentNullException(nameof(blobs));

    /// <summary>Resolve a retry before checking its old expected head. ID reuse fails explicitly.</summary>
    public async ValueTask<HistoryHead?> FindAsync(string documentId, HistoryHead? currentHead,
        HistoryRequestJournal? journal, HistoryRequestIdentity request, CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        ValidateIdentity(request); ValidatePublication(documentId, currentHead, journal);
        cancellationToken.ThrowIfCancellationRequested();
        if (journal?.Current?.Id == request.Id)
        {
            Match(journal.Current, request);
            return currentHead;
        }
        var key = Key(request.Id);
        var cursor = journal?.Index;
        HistoryRequestIndexNode? parent = null;
        while (cursor is not null)
        {
            var node = await LoadAsync(documentId, cursor, cancellationToken).ConfigureAwait(false);
            ValidateChild(parent, node, key);
            if (!PrefixMatches(key, node.Key, node.Bit)) return null;
            if (node.Bit == 256)
            {
                Require(node.Receipt!.Request.Id == request.Id, "Request key collision.");
                Match(node.Receipt.Request, request);
                Require(node.Receipt.Head.Revision < currentHead!.Revision, "Indexed receipt does not precede this publication.");
                return node.Receipt.Head;
            }
            parent = node;
            cursor = Bit(key, node.Bit) ? node.One : node.Zero;
        }
        return null;
    }

    /// <summary>
    /// Prepare the next journal. Promote the preceding inline receipt, including when the next
    /// publication is legacy/non-idempotent. Abandoned prepared nodes are safe unreferenced blobs.
    /// </summary>
    public async ValueTask<HistoryRequestJournal?> AdvanceAsync(string documentId, HistoryHead? currentHead,
        HistoryRequestJournal? journal, HistoryRequestIdentity? request, CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        ValidatePublication(documentId, currentHead, journal);
        if (request is not null) { ValidateIdentity(request); ValidateIndexable(documentId, request); }
        cancellationToken.ThrowIfCancellationRequested();
        var index = journal?.Index;
        if (journal?.Current is not null)
            index = await InsertAsync(documentId, index,
                new HistoryRequestReceipt(documentId, journal.Current, currentHead!), cancellationToken).ConfigureAwait(false);
        return index is null && request is null ? null : new HistoryRequestJournal(documentId,
            checked((currentHead?.Revision ?? 0) + 1), index, request);
    }

    internal static void ValidateJournal(HistoryRequestJournal? journal)
    {
        if (journal is null) return;
        HistoryHeadCodec.Key(journal.DocumentId);
        Require(journal.Revision > 0, "Invalid request-journal revision.");
        Require(journal.Index is not null || journal.Current is not null, "Empty request journal.");
        if (journal.Index is not null) HistoryBlobIO.Validate(journal.Index, MaxNodeBytes);
        if (journal.Current is not null) ValidateIdentity(journal.Current);
    }

    internal static void ValidateIdentity(HistoryRequestIdentity request)
    {
        ArgumentNullException.ThrowIfNull(request);
        ValidateId(request.Id);
        HistoryBlobIO.Validate(new HistoryBlobReference(request.Fingerprint, 0), 0);
    }

    internal static void ValidateId(string id)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(id);
        if (id.Length > MaxRequestIdChars) throw new ArgumentException("Request ID exceeds its limit.", nameof(id));
        _ = Utf8.GetByteCount(id);
    }

    internal static void ValidatePublication(string documentId, HistoryHead? head, HistoryRequestJournal? journal)
    {
        ValidateJournal(journal);
        Require(journal is null || (journal.DocumentId == documentId && journal.Revision == head?.Revision),
            "Request journal does not belong to this document publication.");
        if (head is null) return;
        Require(head.Revision > 0, "Invalid request publication revision.");
        HistoryBlobIO.Validate(head.State, int.MaxValue);
    }

    private async ValueTask<HistoryBlobReference> InsertAsync(string documentId, HistoryBlobReference? root,
        HistoryRequestReceipt receipt, CancellationToken cancellationToken)
    {
        var key = Key(receipt.Request.Id);
        var leaf = new HistoryRequestIndexNode
        { DocumentId = documentId, Bit = 256, Key = key, Receipt = receipt, Zero = null, One = null };
        if (root is null) return await SaveAsync(leaf, cancellationToken).ConfigureAwait(false);
        return await InsertAtAsync(root, null).ConfigureAwait(false);

        async ValueTask<HistoryBlobReference> InsertAtAsync(HistoryBlobReference reference, HistoryRequestIndexNode? parent)
        {
            var node = await LoadAsync(documentId, reference, cancellationToken).ConfigureAwait(false);
            ValidateChild(parent, node, key);
            var divergence = FirstDifference(key, node.Key);
            if (divergence < node.Bit)
            {
                var added = await SaveAsync(leaf, cancellationToken).ConfigureAwait(false);
                return await SaveAsync(new HistoryRequestIndexNode
                {
                    DocumentId = documentId, Bit = divergence, Key = Prefix(key, divergence), Receipt = null,
                    Zero = Bit(key, divergence) ? reference : added,
                    One = Bit(key, divergence) ? added : reference,
                }, cancellationToken).ConfigureAwait(false);
            }
            if (node.Bit == 256)
            {
                Require(node.Receipt == receipt, "A request was published more than once or its key collided.");
                return reference;
            }
            var one = Bit(key, node.Bit);
            var child = await InsertAtAsync((one ? node.One : node.Zero)!, node).ConfigureAwait(false);
            return await SaveAsync(one ? node with { One = child } : node with { Zero = child }, cancellationToken).ConfigureAwait(false);
        }
    }

    private async ValueTask<HistoryRequestIndexNode> LoadAsync(string documentId, HistoryBlobReference reference,
        CancellationToken cancellationToken)
    {
        HistoryBlobIO.Validate(reference, MaxNodeBytes);
        cancellationToken.ThrowIfCancellationRequested();
        using var stream = await _blobs.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
        if (stream is null) throw new PackageChangeException(PackageChangeError.PayloadMissing, "Request-index node is missing.");
        var bytes = await PackageChangeSetCodec.ReadPayloadAsync(reference, stream, cancellationToken).ConfigureAwait(false);
        try
        {
            using var header = JsonDocument.Parse(bytes);
            if (header.RootElement.GetProperty("schemaVersion").GetInt32() != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported request-index version.");
            var envelope = JsonSerializer.Deserialize(bytes, Json.HistoryRequestIndexEnvelope)
                ?? throw new JsonException("Missing request-index envelope.");
            Require(envelope.Schema == "https://docxodus.dev/schemas/history/request-index/v1", "Unexpected request-index schema.");
            ValidateNode(envelope.Node);
            Require(envelope.Node.DocumentId == documentId, "Foreign request-index node.");
            return envelope.Node;
        }
        catch (Exception error) when (error is JsonException or ArgumentException or InvalidOperationException or FormatException or KeyNotFoundException)
        { throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed request-index node."); }
    }

    private async ValueTask<HistoryBlobReference> SaveAsync(HistoryRequestIndexNode node, CancellationToken cancellationToken)
    {
        var bytes = Serialize(node);
        if (bytes.Length > MaxNodeBytes) throw new PackageChangeException(PackageChangeError.ResourceLimit, "Request-index node exceeds its limit.");
        return await HistoryBlobIO.PutBytesAsync(_blobs, bytes, cancellationToken).ConfigureAwait(false);
    }

    private static byte[] Serialize(HistoryRequestIndexNode node)
    {
        ValidateNode(node);
        return JsonSerializer.SerializeToUtf8Bytes(new HistoryRequestIndexEnvelope
        { SchemaVersion = 1, Schema = "https://docxodus.dev/schemas/history/request-index/v1", Node = node }, Json.HistoryRequestIndexEnvelope);
    }

    /// <summary>
    /// A receipt published inline must also fit the ONE index node that the next publication
    /// promotes it into. Both IDs are bounded in characters, but JSON escapes each non-ASCII or
    /// reserved character to six bytes, so legal IDs can still exceed a node. Rejecting that here
    /// fails the call supplying the ID instead of the next unrelated publication, which could
    /// otherwise never promote the receipt and would leave the document permanently unpublishable.
    /// The probe uses the largest head a publication can carry, so it never accepts a receipt the
    /// promotion would refuse.
    /// </summary>
    private static void ValidateIndexable(string documentId, HistoryRequestIdentity request)
    {
        var widest = new HistoryHead(long.MaxValue, new HistoryBlobReference(new VerificationDigest
        { Algorithm = "SHA-256", Value = new string('f', 64) }, int.MaxValue));
        var probe = new HistoryRequestIndexNode
        {
            DocumentId = documentId, Bit = 256, Key = Key(request.Id), Zero = null, One = null,
            Receipt = new HistoryRequestReceipt(documentId, request, widest),
        };
        if (Serialize(probe).Length > MaxNodeBytes)
            throw new PackageChangeException(PackageChangeError.ResourceLimit,
                "Document and request ID exceed the request-index node limit.");
    }

    private static void ValidateNode(HistoryRequestIndexNode node)
    {
        ArgumentNullException.ThrowIfNull(node);
        HistoryHeadCodec.Key(node.DocumentId);
        Require(node.Bit is >= 0 and <= 256, "Invalid request-index bit position.");
        Require(node.Key is { Length: 64 } && node.Key.All(c => c is >= '0' and <= '9' or >= 'a' and <= 'f'), "Invalid request-index key.");
        if (node.Bit == 256)
        {
            Require(node.Receipt is not null && node.Zero is null && node.One is null, "Invalid request-index leaf.");
            var receipt = node.Receipt!;
            Require(receipt.DocumentId == node.DocumentId, "Foreign request receipt.");
            ValidateIdentity(receipt.Request);
            ArgumentNullException.ThrowIfNull(receipt.Head);
            Require(receipt.Head.Revision > 0 && node.Key == Key(receipt.Request.Id), "Invalid request receipt identity.");
            HistoryBlobIO.Validate(receipt.Head.State, int.MaxValue);
        }
        else
        {
            Require(node.Receipt is null && node.Zero is not null && node.One is not null && node.Zero != node.One
                && node.Key == Prefix(node.Key, node.Bit), "Invalid request-index branch.");
            HistoryBlobIO.Validate(node.Zero!, MaxNodeBytes); HistoryBlobIO.Validate(node.One!, MaxNodeBytes);
        }
    }

    private static void ValidateChild(HistoryRequestIndexNode? parent, HistoryRequestIndexNode child, string key)
    {
        if (parent is null) return;
        Require(child.Bit > parent.Bit && PrefixMatches(child.Key, parent.Key, parent.Bit)
            && Bit(child.Key, parent.Bit) == Bit(key, parent.Bit), "Request-index path is inconsistent or cyclic.");
    }

    private static void Match(HistoryRequestIdentity stored, HistoryRequestIdentity requested)
    {
        if (stored.Fingerprint != requested.Fingerprint)
            throw new DocxHistoryException(DocxHistoryError.RequestConflict, "Request ID is already bound to different input.");
    }

    private static string Key(string id) => Convert.ToHexString(SHA256.HashData(Utf8.GetBytes(id))).ToLowerInvariant();
    private static bool Bit(string hex, int bit) => (Nibble(hex[bit / 4]) & (8 >> (bit % 4))) != 0;
    private static int Nibble(char c) => c <= '9' ? c - '0' : c - 'a' + 10;
    private static int FirstDifference(string left, string right)
    {
        for (var i = 0; i < 256; i++) if (Bit(left, i) != Bit(right, i)) return i;
        return 256;
    }
    private static bool PrefixMatches(string key, string prefix, int bits) => FirstDifference(key, prefix) >= bits;
    private static string Prefix(string key, int bits)
    {
        var chars = key.ToCharArray();
        for (var i = bits / 4; i < chars.Length; i++)
        {
            var keep = i == bits / 4 ? bits % 4 : 0;
            chars[i] = "0123456789abcdef"[Nibble(chars[i]) & (15 << (4 - keep))];
        }
        return new string(chars);
    }
    private static void Require(bool condition, string message)
    { if (!condition) throw new DocxHistoryException(DocxHistoryError.InvalidHistory, message); }
}

internal sealed record HistoryRequestIndexNode
{
    public required string DocumentId { get; init; }
    public required int Bit { get; init; }
    public required string Key { get; init; }
    public required HistoryRequestReceipt? Receipt { get; init; }
    public required HistoryBlobReference? Zero { get; init; }
    public required HistoryBlobReference? One { get; init; }
}

internal sealed record HistoryRequestIndexEnvelope
{
    public required string Schema { get; init; }
    public required int SchemaVersion { get; init; }
    public required HistoryRequestIndexNode Node { get; init; }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow, AllowDuplicateProperties = false,
    RespectNullableAnnotations = true, RespectRequiredConstructorParameters = true, MaxDepth = 12)]
[JsonSerializable(typeof(HistoryRequestIndexEnvelope))]
internal partial class HistoryRequestJsonContext : JsonSerializerContext { }

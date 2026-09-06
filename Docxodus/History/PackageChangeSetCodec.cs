// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.ObjectModel;
using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>A versioned JSON manifest with payloads stored separately by their SHA-256 identities.</summary>
public static class PackageChangeSetCodec
{
    public const string SchemaId = "https://docxodus.dev/schemas/history/package-changes/v1";

    /// <summary>Persist all payloads before returning a manifest suitable for durable publication.</summary>
    public static async ValueTask<byte[]> SaveAsync(PackageChangeSet changes, IHistoryBlobStore store,
        PackageChangeLimits? limits = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(store);
        cancellationToken.ThrowIfCancellationRequested();
        var manifest = Encode(changes, limits);
        foreach (var digest in changes.PayloadDigests)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var content = changes.OpenPayload(digest);
            await store.PutAsync(new HistoryBlobReference(digest, changes.PayloadLength(digest)), content,
                cancellationToken).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return manifest;
    }

    /// <summary>Encode only metadata; callers must retain every referenced payload separately.</summary>
    public static byte[] Encode(PackageChangeSet changes, PackageChangeLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(changes);
        limits ??= new PackageChangeLimits();
        limits.Validate();
        Budget(changes.Changes.Count <= limits.MaxChanges, "Too many entry changes.");
        Budget(changes.RetainedPayloadBytes <= limits.MaxTotalPayloadBytes, "Payload total exceeds the limit.");
        using var buffer = new BoundedManifestStream(limits.MaxManifestBytes);
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("afterDigest", changes.AfterDigest.Value);
            writer.WriteString("beforeDigest", changes.BeforeDigest.Value);
            writer.WriteStartArray("changes");
            foreach (var change in changes.Changes)
            {
                Budget(change.Uri.Length <= limits.MaxEntryNameLength
                    && (change.BeforeEntryName?.Length ?? 0) <= limits.MaxEntryNameLength
                    && (change.AfterEntryName?.Length ?? 0) <= limits.MaxEntryNameLength, "Entry name exceeds the limit.");
                writer.WriteStartObject();
                writer.WriteString("afterDigest", change.AfterDigest?.Value);
                writer.WriteString("afterEntryName", change.AfterEntryName);
                writer.WriteString("beforeDigest", change.BeforeDigest?.Value);
                writer.WriteString("beforeEntryName", change.BeforeEntryName);
                writer.WriteString("uri", change.Uri);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteStartArray("payloads");
            foreach (var digest in changes.PayloadDigests)
            {
                var length = changes.PayloadLength(digest);
                Budget(length <= limits.MaxPayloadBytes, "Payload exceeds the per-blob limit.");
                writer.WriteStartObject();
                writer.WriteString("digest", digest.Value);
                writer.WriteNumber("length", length);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteString("schema", SchemaId);
            writer.WriteNumber("schemaVersion", 1);
            writer.WriteEndObject();
        }
        return buffer.ToArray();
    }

    /// <summary>
    /// Validate the complete manifest before any storage reads, then check every payload's length
    /// and digest. No change set is returned on corruption, truncation, a missing blob, or cancellation.
    /// </summary>
    public static async ValueTask<PackageChangeSet> LoadAsync(byte[] manifest, IHistoryBlobStore store,
        PackageChangeLimits? limits = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(store);
        cancellationToken.ThrowIfCancellationRequested();
        limits ??= new PackageChangeLimits();
        limits.Validate();
        Budget(manifest.Length <= limits.MaxManifestBytes, "Manifest exceeds the byte limit.");
        var parsed = Parse(manifest, limits);
        var payloads = new Dictionary<VerificationDigest, byte[]>();
        foreach (var reference in parsed.Payloads)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var content = await store.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
            if (content is null)
                throw new PackageChangeException(PackageChangeError.PayloadMissing,
                    $"Missing payload {reference.Digest.Value}.");
            payloads.Add(reference.Digest,
                await ReadPayloadAsync(reference, content, cancellationToken).ConfigureAwait(false));
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new PackageChangeSet(parsed.Before, parsed.After, parsed.Changes,
            new ReadOnlyDictionary<VerificationDigest, byte[]>(payloads));
    }

    internal static async ValueTask<byte[]> ReadPayloadAsync(HistoryBlobReference reference, Stream content,
        CancellationToken cancellationToken)
    {
        var bytes = new byte[reference.Length];
        try
        {
            await content.ReadExactlyAsync(bytes, cancellationToken).ConfigureAwait(false);
            if (await content.ReadAsync(new byte[1], cancellationToken).ConfigureAwait(false) != 0)
                throw new EndOfStreamException();
        }
        catch (EndOfStreamException)
        {
            throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Payload length does not match its reference.");
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() != reference.Digest.Value)
            throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Payload SHA-256 does not match its reference.");
        return bytes;
    }

    private sealed record Manifest(VerificationDigest Before, VerificationDigest After,
        IReadOnlyList<PackageEntryChange> Changes, IReadOnlyList<HistoryBlobReference> Payloads);

    private static Manifest Parse(byte[] bytes, PackageChangeLimits limits)
    {
        try
        {
            using var document = JsonDocument.Parse(bytes, new JsonDocumentOptions { MaxDepth = 8 });
            var root = Object(document.RootElement, "afterDigest", "beforeDigest", "changes", "payloads", "schema", "schemaVersion");
            if (root["schemaVersion"].GetInt32() != 1)
                throw new PackageChangeException(PackageChangeError.UnsupportedVersion, "Unsupported package-change codec version.");
            Require(root["schema"].GetString() == SchemaId, "Unknown manifest schema.");
            var before = Digest(root["beforeDigest"]);
            var after = Digest(root["afterDigest"]);
            var changes = new List<PackageEntryChange>();
            var uris = new HashSet<string>(StringComparer.Ordinal);
            var beforeUris = new HashSet<string>(PackageManifestGenerator.CanonicalPartNameComparer);
            var afterUris = new HashSet<string>(PackageManifestGenerator.CanonicalPartNameComparer);
            var referenced = new HashSet<VerificationDigest>();
            Budget(root["changes"].GetArrayLength() <= limits.MaxChanges, "Too many entry changes.");
            foreach (var value in root["changes"].EnumerateArray())
            {
                var entry = Object(value, "afterDigest", "afterEntryName", "beforeDigest", "beforeEntryName", "uri");
                var uri = entry["uri"].GetString();
                Require(uri is not null, "Entry URI is required.");
                Budget(uri!.Length <= limits.MaxEntryNameLength, "Entry URI exceeds the limit.");
                Require(uris.Add(uri), "Duplicate entry URI.");
                var left = Endpoint(entry["beforeEntryName"], entry["beforeDigest"], uri, limits);
                var right = Endpoint(entry["afterEntryName"], entry["afterDigest"], uri, limits);
                Require(left.Name is null || beforeUris.Add(uri), "Duplicate before entry URI.");
                Require(right.Name is null || afterUris.Add(uri), "Duplicate after entry URI.");
                Require(left.Digest != right.Digest, "An entry change must change content or existence.");
                changes.Add(new PackageEntryChange(uri, left.Name, left.Digest, right.Name, right.Digest));
                if (left.Digest is not null) referenced.Add(left.Digest);
                if (right.Digest is not null) referenced.Add(right.Digest);
            }
            // Empty ZIP directories are retained as effects but excluded from OPC content identity.
            Require(changes.Count != 0 || before == after, "Package identities disagree with the change count.");
            Budget(root["payloads"].GetArrayLength() <= (long)limits.MaxChanges * 2, "Too many payload references.");
            var payloads = new List<HistoryBlobReference>();
            var payloadDigests = new HashSet<VerificationDigest>();
            long total = 0;
            foreach (var value in root["payloads"].EnumerateArray())
            {
                var payload = Object(value, "digest", "length");
                var digest = Digest(payload["digest"]);
                var length = payload["length"].GetInt32();
                Require(length >= 0, "Negative payload length.");
                Budget(length <= limits.MaxPayloadBytes, "Payload exceeds the per-blob limit.");
                total += length;
                Budget(total <= limits.MaxTotalPayloadBytes, "Payload total exceeds the limit.");
                Require(payloadDigests.Add(digest), "Duplicate payload digest.");
                payloads.Add(new HistoryBlobReference(digest, length));
            }
            Require(payloadDigests.SetEquals(referenced), "Payload inventory must exactly cover the entry effects.");
            return new Manifest(before, after,
                Array.AsReadOnly(changes.OrderBy(c => c.Uri, StringComparer.Ordinal).ToArray()),
                Array.AsReadOnly(payloads.OrderBy(p => p.Digest.Value, StringComparer.Ordinal).ToArray()));
        }
        catch (Exception ex) when (ex is JsonException or InvalidOperationException or FormatException)
        {
            throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed package-change manifest.");
        }
    }

    private static (string? Name, VerificationDigest? Digest) Endpoint(JsonElement name, JsonElement digest,
        string uri, PackageChangeLimits limits)
    {
        if (name.ValueKind == JsonValueKind.Null && digest.ValueKind == JsonValueKind.Null) return (null, null);
        var entryName = name.GetString();
        Require(entryName is not null, "An entry digest requires a name.");
        Budget(entryName!.Length <= limits.MaxEntryNameLength, "Entry name exceeds the limit.");
        Require(PackageManifestGenerator.TryCanonicalizeEntryName(entryName, out var canonical) && canonical == uri,
            "Entry name does not map to its declared OPC URI.");
        return (entryName, Digest(digest));
    }

    private static VerificationDigest Digest(JsonElement value)
    {
        var digest = value.GetString();
        Require(digest is { Length: 64 } && digest.All(c => c is >= '0' and <= '9' or >= 'a' and <= 'f'),
            "Expected a lowercase SHA-256 digest.");
        return new VerificationDigest { Algorithm = "SHA-256", Value = digest! };
    }

    private static Dictionary<string, JsonElement> Object(JsonElement value, params string[] properties)
    {
        var result = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        foreach (var property in value.EnumerateObject())
            Require(properties.Contains(property.Name, StringComparer.Ordinal) && result.TryAdd(property.Name, property.Value),
                "Unknown or duplicate JSON property.");
        Require(result.Count == properties.Length, "Missing JSON property.");
        return result;
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new PackageChangeException(PackageChangeError.InvalidManifest, message);
    }

    private static void Budget(bool condition, string message)
    {
        if (!condition) throw new PackageChangeException(PackageChangeError.ResourceLimit, message);
    }

    private sealed class BoundedManifestStream(int maximumBytes) : MemoryStream
    {
        public override void Write(ReadOnlySpan<byte> buffer)
        {
            Budget(Position + buffer.Length <= maximumBytes, "Manifest exceeds the byte limit.");
            base.Write(buffer);
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            Budget(Position + count <= maximumBytes, "Manifest exceeds the byte limit.");
            base.Write(buffer, offset, count);
        }
    }
}

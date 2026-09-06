// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.ObjectModel;
using System.IO.Compression;
using System.Security.Cryptography;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Failure to construct or replay a complete package contribution.</summary>
public enum PackageChangeError
{
    InvalidPackage,
    BaseMismatch,
    PayloadMismatch,
    ResultMismatch,
    InvalidManifest,
    UnsupportedVersion,
    PayloadMissing,
    ResourceLimit,
}

/// <summary>A package contribution failed before publishing any output.</summary>
public sealed class PackageChangeException : Exception
{
    internal PackageChangeException(PackageChangeError code, string message) : base(message) =>
        Code = code;

    internal PackageChangeException(PackageChangeError code, string message, Exception innerException)
        : base(message, innerException) => Code = code;

    public PackageChangeError Code { get; }
}

/// <summary>
/// One entry addition, removal, or replacement. A null digest/name pair denotes absence;
/// an empty entry has a non-null digest. Relationships and content types are ordinary entries.
/// </summary>
public sealed record PackageEntryChange(
    string Uri,
    string? BeforeEntryName,
    VerificationDigest? BeforeDigest,
    string? AfterEntryName,
    VerificationDigest? AfterDigest);

/// <summary>
/// An immutable, reversible contribution between two complete OPC packages. Retains only
/// changed entry payloads, deduplicated by SHA-256, including the before bytes needed to invert.
/// Intended for imported versions and whole-part fallbacks, not the per-keystroke edit path.
/// </summary>
/// <remarks>
/// Replay preserves every uncompressed entry byte, including opaque XML and media. It does
/// not reproduce ZIP timestamps, order, or compression: exact version downloads must retain
/// the original snapshot bytes separately. This primitive does not merge concurrent edits.
/// Replay writes stored ZIP entries to avoid recompression and preserve expansion limits;
/// this temporary package can be larger than a compressed snapshot. Normal save owns compression.
/// </remarks>
public sealed class PackageChangeSet
{
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private readonly IReadOnlyDictionary<VerificationDigest, byte[]> _payloads;

    internal PackageChangeSet(
        VerificationDigest beforeDigest,
        VerificationDigest afterDigest,
        IReadOnlyList<PackageEntryChange> changes,
        IReadOnlyDictionary<VerificationDigest, byte[]> payloads)
    {
        BeforeDigest = beforeDigest;
        AfterDigest = afterDigest;
        Changes = changes;
        _payloads = payloads;
        PayloadDigests = Array.AsReadOnly(payloads.Keys.OrderBy(d => d.Value, StringComparer.Ordinal).ToArray());
        RetainedPayloadBytes = payloads.Values.Sum(bytes => (long)bytes.Length);
    }

    /// <summary>Existing ordered OPC content identity, independent of ZIP serialization.</summary>
    public VerificationDigest BeforeDigest { get; }
    public VerificationDigest AfterDigest { get; }
    public IReadOnlyList<PackageEntryChange> Changes { get; }
    public IReadOnlyList<VerificationDigest> PayloadDigests { get; }
    public long RetainedPayloadBytes { get; }

    /// <summary>A defensive copy of a retained before/after payload.</summary>
    public byte[] GetPayload(VerificationDigest digest) => _payloads[digest].ToArray();

    internal Stream OpenPayload(VerificationDigest digest) => new MemoryStream(_payloads[digest], writable: false);
    internal int PayloadLength(VerificationDigest digest) => _payloads[digest].Length;

    /// <summary>
    /// Compare bounded, valid OPC packages without modifying either input. Only changed
    /// entries are retained; shared, unchanged assets are not copied into the contribution.
    /// </summary>
    public static PackageChangeSet Create(
        byte[] before, byte[] after, PackageManifestOptions? options = null)
    {
        var beforeManifest = Inspect(before, options);
        var afterManifest = Inspect(after, options);
        using var beforeArchive = Open(before);
        using var afterArchive = Open(after);
        var beforeEntries = Index(beforeArchive);
        var afterEntries = Index(afterArchive);
        var beforeIdentities = beforeManifest.Entries.ToDictionary(e => e.Uri, StringComparer.Ordinal);
        var afterIdentities = afterManifest.Entries.ToDictionary(e => e.Uri, StringComparer.Ordinal);
        var changes = new List<PackageEntryChange>();
        var payloads = new Dictionary<VerificationDigest, byte[]>();

        foreach (var uri in beforeIdentities.Keys.Union(afterIdentities.Keys, StringComparer.Ordinal)
                     .OrderBy(uri => uri, StringComparer.Ordinal))
        {
            beforeIdentities.TryGetValue(uri, out var left);
            afterIdentities.TryGetValue(uri, out var right);
            if (left?.RawBytesDigest == right?.RawBytesDigest) continue;

            var leftEntry = left is null ? null : beforeEntries[uri];
            var rightEntry = right is null ? null : afterEntries[uri];
            Retain(leftEntry, left?.RawBytesDigest, payloads);
            Retain(rightEntry, right?.RawBytesDigest, payloads);
            changes.Add(new PackageEntryChange(uri, leftEntry?.FullName, left?.RawBytesDigest,
                rightEntry?.FullName, right?.RawBytesDigest));
        }

        return new PackageChangeSet(beforeManifest.OrderedOpcContentDigest!,
            afterManifest.OrderedOpcContentDigest!, changes.AsReadOnly(),
            new ReadOnlyDictionary<VerificationDigest, byte[]>(payloads));
    }

    /// <summary>Reverse every effect without copying or rereading its immutable payloads.</summary>
    public PackageChangeSet Invert() => new(AfterDigest, BeforeDigest,
        Array.AsReadOnly(Changes.Select(change => new PackageEntryChange(change.Uri,
            change.AfterEntryName, change.AfterDigest, change.BeforeEntryName, change.BeforeDigest)).ToArray()),
        _payloads);

    /// <summary>
    /// Apply atomically to a package with the exact expected OPC content identity. Repacking
    /// the base is allowed; any content difference, including an unrelated part, is rejected.
    /// The input is never edited, and no result is returned until its target digest is verified.
    /// </summary>
    public byte[] Apply(byte[] packageBytes, PackageManifestOptions? options = null)
    {
        var manifest = Inspect(packageBytes, options);
        if (manifest.OrderedOpcContentDigest != BeforeDigest)
            throw new PackageChangeException(PackageChangeError.BaseMismatch,
                "The package content does not match the contribution's expected base.");
        var identities = manifest.Entries.ToDictionary(e => e.Uri, StringComparer.Ordinal);
        foreach (var change in Changes)
        {
            identities.TryGetValue(change.Uri, out var entry);
            if (entry?.RawBytesDigest != change.BeforeDigest)
                throw new PackageChangeException(PackageChangeError.BaseMismatch,
                    $"Entry {change.Uri} does not match the contribution's expected before payload.");
        }
        if (Changes.Count == 0) return packageBytes.ToArray();

        using var source = Open(packageBytes);
        var entries = Index(source);
        var replacements = Changes.ToDictionary(change => change.Uri, StringComparer.Ordinal);
        using var buffer = new MemoryStream();
        using (var output = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var uri in entries.Keys.Union(replacements.Keys, StringComparer.Ordinal)
                         .OrderBy(uri => uri, StringComparer.Ordinal))
            {
                replacements.TryGetValue(uri, out var change);
                if (change is not null && change.AfterDigest is null) continue;
                var name = change?.AfterEntryName ?? entries[uri].FullName;
                // Recompression can turn a valid input into a package that fails the same
                // expansion limits on its next replay. Storage/export owns compression policy.
                var entry = output.CreateEntry(name, CompressionLevel.NoCompression);
                entry.LastWriteTime = ZipEpoch;
                using var destination = entry.Open();
                if (change?.AfterDigest is { } digest)
                    destination.Write(_payloads[digest]);
                else
                {
                    using var original = entries[uri].Open();
                    original.CopyTo(destination);
                }
            }
        }

        var result = buffer.ToArray();
        var resultManifest = Inspect(result, options);
        if (resultManifest.OrderedOpcContentDigest != AfterDigest)
            throw new PackageChangeException(PackageChangeError.ResultMismatch,
                "The reconstructed package does not match the contribution's target content.");
        return result;
    }

    private static PackageManifest Inspect(byte[] bytes, PackageManifestOptions? options)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        var manifest = PackageManifestGenerator.Generate(bytes, options);
        if (!manifest.IsValid || manifest.PackageKind != "opc" || manifest.OrderedOpcContentDigest is null)
            throw new PackageChangeException(PackageChangeError.InvalidPackage,
                "A complete, valid, bounded OPC package is required. " +
                string.Join(", ", manifest.Findings.Where(f => f.Severity == VerificationFindingSeverity.Error)
                    .Select(f => f.Code)));
        return manifest;
    }

    private static ZipArchive Open(byte[] bytes) =>
        new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);

    private static Dictionary<string, ZipArchiveEntry> Index(ZipArchive archive) =>
        archive.Entries.ToDictionary(entry =>
        {
            // Use the inspector's OPC mapping, including percent-encoded Unicode part names.
            if (!PackageManifestGenerator.TryCanonicalizeEntryName(entry.FullName, out var uri))
                throw new PackageChangeException(PackageChangeError.InvalidPackage,
                    "The package contains an invalid entry name.");
            return uri;
        }, StringComparer.Ordinal);

    private static void Retain(ZipArchiveEntry? entry, VerificationDigest? digest,
        IDictionary<VerificationDigest, byte[]> payloads)
    {
        if (entry is null || digest is null || payloads.ContainsKey(digest)) return;
        if (entry.Length > int.MaxValue)
            throw new PackageChangeException(PackageChangeError.InvalidPackage,
                "A changed entry exceeds the supported payload size.");
        var bytes = new byte[(int)entry.Length];
        using var input = entry.Open();
        input.ReadExactly(bytes);
        if (Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() != digest.Value)
            throw new PackageChangeException(PackageChangeError.PayloadMismatch,
                "A changed entry does not match its inspected digest.");
        payloads.Add(digest, bytes);
    }
}

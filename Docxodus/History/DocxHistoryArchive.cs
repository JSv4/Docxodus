// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>
/// A validated, read-only portable history file. Await active calls before Dispose. Stream input
/// must remain unchanged, readable and seekable until disposal; leaveOpen defaults to true.
/// Byte-array input is copied before the first await and owned by the archive.
/// </summary>
public sealed class DocxHistoryArchive : DocxHistoryReader, IDisposable
{
    private readonly HistoryArchiveBlobStore _store;
    internal IReadOnlyList<HistoryBlobReference> Inventory { get; }
    internal IHistoryBlobStore Blobs => _store;
    public DocxHistoryArchiveInfo Info { get; }
    public DocxHistoryView View { get; }

    private DocxHistoryArchive(HistoryArchiveBlobStore store, HistoryArchiveManifest manifest,
        DocxHistoryView view, DocxHistoryArchiveLimits limits, PackageManifestOptions? packageOptions)
        : base(new DocxVersionHistory(store, new PinnedHead(manifest.DocumentId, manifest.Head),
            Math.Min(limits.MaxBlobBytes, limits.MaxSnapshotBytes), limits.MaxRecordBytes, packageOptions), manifest.DocumentId)
    { _store = store; Inventory = manifest.Blobs; Info = manifest.Info; View = view; }

    public static ValueTask<DocxHistoryArchive> OpenAsync(byte[] bytes, DocxHistoryArchiveLimits? limits = null,
        PackageManifestOptions? packageOptions = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(bytes); limits ??= new(); limits.Validate();
        cancellationToken.ThrowIfCancellationRequested();
        HistoryArchiveManifest.Budget(bytes.LongLength <= limits.MaxArchiveBytes, "History archive exceeds its byte limit.");
        return OpenAsync(new MemoryStream(bytes.ToArray(), writable: false), leaveOpen: false, limits, packageOptions, cancellationToken);
    }

    public static async ValueTask<DocxHistoryArchive> OpenAsync(Stream input, bool leaveOpen = true,
        DocxHistoryArchiveLimits? limits = null, PackageManifestOptions? packageOptions = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(input); limits ??= new(); limits.Validate();
        if (!input.CanRead || !input.CanSeek) throw new ArgumentException("Archive input must be readable and seekable.", nameof(input));
        ZipArchive? zip = null; HistoryArchiveBlobStore? store = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var count = await HistoryArchiveZip.PreflightAsync(input, limits, cancellationToken).ConfigureAwait(false);
            zip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
            HistoryArchiveManifest.Require(zip.Entries.Count == count, "ZIP entry count disagrees with preflight.");
            var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
            long total = 0;
            foreach (var entry in zip.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var manifest = entry.FullName == HistoryArchiveManifest.EntryName;
                HistoryArchiveManifest.Require(manifest || (entry.FullName.Length == 70 && entry.FullName.StartsWith("blobs/", StringComparison.Ordinal)
                    && entry.FullName.AsSpan(6).IndexOfAnyExcept("0123456789abcdef") < 0), "Unknown or noncanonical archive entry.");
                HistoryArchiveManifest.Require(entries.TryAdd(entry.FullName, entry), "Duplicate archive ZIP entry.");
                HistoryArchiveManifest.Budget(entry.Length >= 0 && entry.Length <= (manifest ? limits.MaxManifestBytes : limits.MaxBlobBytes),
                    "Archive entry exceeds its byte limit.");
                if (!manifest)
                {
                    HistoryArchiveManifest.Budget(entry.Length <= limits.MaxTotalBlobBytes - total, "Archive entries exceed total blob bytes.");
                    total += entry.Length;
                }
            }
            HistoryArchiveManifest.Require(entries.TryGetValue(HistoryArchiveManifest.EntryName, out var header), "Archive manifest is missing.");
            byte[] metadata;
            using (var content = header!.Open())
            {
                metadata = new byte[(int)header.Length];
                await content.ReadExactlyAsync(metadata, cancellationToken).ConfigureAwait(false);
                HistoryArchiveManifest.Require(await content.ReadAsync(new byte[1], cancellationToken).ConfigureAwait(false) == 0,
                    "Archive manifest length disagrees with ZIP metadata.");
            }
            var parsed = HistoryArchiveManifest.Decode(metadata, limits);
            HistoryArchiveManifest.Require(parsed.Blobs.Count == entries.Count - 1 && parsed.Info.TotalBlobBytes == total,
                "Archive ZIP entries disagree with the manifest inventory.");
            foreach (var blob in parsed.Blobs)
                HistoryArchiveManifest.Require(entries.TryGetValue(HistoryArchiveManifest.BlobName(blob), out var entry)
                    && entry.Length == blob.Length, "Archive inventory blob is missing or has the wrong length.");
            store = new HistoryArchiveBlobStore(zip, input, leaveOpen, entries, limits.MaxBlobBytes);
            var graph = await HistoryArchiveGraph.LoadAsync(parsed.DocumentId, parsed.Head, store, limits,
                packageOptions, cancellationToken).ConfigureAwait(false);
            HistoryArchiveManifest.Require(graph.Inventory.SequenceEqual(parsed.Blobs), "Archive inventory contains unreferenced blobs.");
            cancellationToken.ThrowIfCancellationRequested();
            return new DocxHistoryArchive(store, parsed, graph.View, limits, packageOptions);
        }
        catch (Exception error) when (error is InvalidDataException or EndOfStreamException)
        {
            Cleanup(); throw new PackageChangeException(PackageChangeError.InvalidManifest, "Malformed or truncated history ZIP archive.");
        }
        catch { Cleanup(); throw; }

        void Cleanup()
        {
            if (store is not null) store.Dispose();
            else { zip?.Dispose(); if (!leaveOpen) input.Dispose(); }
        }
    }

    public void Dispose() => _store.Dispose();

    private sealed class PinnedHead(string documentId, HistoryHead head) : IHistoryHeadStore
    {
        public ValueTask<HistoryHead?> ReadAsync(string id, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (id != documentId) throw new DocxHistoryException(DocxHistoryError.ForeignDocument, "Archive belongs to a different document.");
            return ValueTask.FromResult<HistoryHead?>(head);
        }
        public ValueTask<HistoryHead?> TryAdvanceAsync(string id, HistoryHead? expected, HistoryBlobReference state,
            CancellationToken cancellationToken = default) => throw new NotSupportedException("History archives are read-only.");
    }
}

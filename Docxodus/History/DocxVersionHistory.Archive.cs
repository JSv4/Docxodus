// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Stream one captured document's complete referenced history to an empty output. Leaves the
    /// stream open; failure may leave partial bytes, which the host must discard. No source head
    /// is advanced and no shared-store enumeration occurs. Native hosts can use FileStreams.
    /// </summary>
    public async ValueTask<DocxHistoryArchiveInfo> ExportHistoryArchiveAsync(string documentId, Stream output,
        DocxHistoryArchiveLimits? limits = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(output); limits ??= new(); limits.Validate();
        if (!output.CanWrite || (output.CanSeek && (output.Position != 0 || output.Length != 0)))
            throw new ArgumentException("Archive output must be writable and empty.", nameof(output));
        limits = limits with
        {
            MaxSnapshotBytes = Math.Min(limits.MaxSnapshotBytes, _maxSnapshotBytes),
            MaxRecordBytes = Math.Min(limits.MaxRecordBytes, _maxRecordBytes),
        };
        var view = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        var graph = await HistoryArchiveGraph.LoadAsync(documentId, view.Head, _blobs, limits,
            _packageOptions, cancellationToken).ConfigureAwait(false);
        var manifest = new HistoryArchiveManifest(documentId, view.Head, graph.Inventory);
        var metadata = manifest.Encode(limits);
        using var bounded = new HistoryArchiveWriteStream(output, limits.MaxArchiveBytes);
        using (var zip = new ZipArchive(bounded, ZipArchiveMode.Create, leaveOpen: true))
        {
            using (var entry = Entry(zip, HistoryArchiveManifest.EntryName))
                await entry.WriteAsync(metadata, cancellationToken).ConfigureAwait(false);
            foreach (var reference in manifest.Blobs)
            {
                cancellationToken.ThrowIfCancellationRequested();
                using var input = await _blobs.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
                if (input is null) throw new PackageChangeException(PackageChangeError.PayloadMissing, "Export blob disappeared.");
                using var entry = Entry(zip, HistoryArchiveManifest.BlobName(reference));
                await HistoryBlobIO.CopyVerifiedAsync(reference, input, entry, cancellationToken).ConfigureAwait(false);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return manifest.Info;
    }

    private static Stream Entry(ZipArchive zip, string name)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        entry.LastWriteTime = new DateTimeOffset(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
        return entry.Open();
    }
}

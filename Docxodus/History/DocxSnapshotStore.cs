// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Exact immutable DOCX bytes plus the independent ordered OPC content identity.</summary>
public sealed record DocxSnapshotReference(HistoryBlobReference Blob, VerificationDigest ContentDigest);

/// <summary>
/// Capture/export exact WordprocessingML snapshots using host-owned blob storage. No mutable
/// session or serializer touches the supplied bytes. Snapshot inspection is a version/checkpoint
/// boundary operation, not a typing-path operation. Version metadata belongs to the host log.
/// </summary>
public sealed class DocxSnapshotStore
{
    private readonly IHistoryBlobStore _blobs;
    private readonly int _maxSnapshotBytes;
    private readonly PackageManifestOptions? _packageOptions;

    public DocxSnapshotStore(IHistoryBlobStore blobs,
        int maxSnapshotBytes = HistoryBlobIO.DefaultMaxBlobBytes, PackageManifestOptions? packageOptions = null)
    {
        ArgumentNullException.ThrowIfNull(blobs);
        _blobs = blobs;
        _maxSnapshotBytes = HistoryBlobIO.ValidateLimit(maxSnapshotBytes);
        _packageOptions = packageOptions;
    }

    /// <summary>Return a reference only after verified bytes are stored. Failed writes may leave unreferenced blobs.</summary>
    public async ValueTask<DocxSnapshotReference> CaptureAsync(byte[] docxBytes,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(docxBytes);
        cancellationToken.ThrowIfCancellationRequested();
        if (docxBytes.Length > _maxSnapshotBytes)
            throw new PackageChangeException(PackageChangeError.ResourceLimit, "Snapshot exceeds the byte limit.");
        // Isolate the exact captured value before yielding to an asynchronous host adapter.
        var bytes = docxBytes.ToArray();
        var manifest = Inspect(bytes);
        var reference = new HistoryBlobReference(manifest.RawPackageBytesDigest, bytes.Length);
        cancellationToken.ThrowIfCancellationRequested();
        using var content = new MemoryStream(bytes, writable: false);
        await _blobs.PutAsync(reference, content, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return new DocxSnapshotReference(reference, manifest.OrderedOpcContentDigest!);
    }

    /// <summary>Export a fresh owned byte array after verifying length, exact digest, and OPC content identity.</summary>
    public async ValueTask<byte[]> ExportAsync(DocxSnapshotReference snapshot,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        HistoryBlobIO.Validate(snapshot.Blob, _maxSnapshotBytes);
        HistoryBlobIO.Validate(new HistoryBlobReference(snapshot.ContentDigest, 0), _maxSnapshotBytes);
        cancellationToken.ThrowIfCancellationRequested();
        using var content = await _blobs.OpenReadAsync(snapshot.Blob, cancellationToken).ConfigureAwait(false);
        if (content is null)
            throw new PackageChangeException(PackageChangeError.PayloadMissing, "Snapshot bytes are missing.");
        var bytes = await PackageChangeSetCodec.ReadPayloadAsync(snapshot.Blob, content, cancellationToken).ConfigureAwait(false);
        if (Inspect(bytes).OrderedOpcContentDigest != snapshot.ContentDigest)
            throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Snapshot OPC content identity does not match.");
        cancellationToken.ThrowIfCancellationRequested();
        return bytes;
    }

    /// <summary>
    /// Verify both snapshots and return the existing lazy semantic/redline comparison. DocxDiff
    /// owns compatibility policy and limitations, including pre-existing tracked revisions.
    /// </summary>
    public async ValueTask<DocxDiffComparison> CompareAsync(DocxSnapshotReference before,
        DocxSnapshotReference after, DocxDiffSettings? settings = null, CancellationToken cancellationToken = default)
    {
        var left = await ExportAsync(before, cancellationToken).ConfigureAwait(false);
        var right = await ExportAsync(after, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return DocxDiff.CreateComparison(new WmlDocument("before.docx", left), new WmlDocument("after.docx", right), settings);
    }

    internal PackageManifest Inspect(byte[] bytes)
    {
        var manifest = PackageManifestGenerator.Generate(bytes, _packageOptions);
        if (!manifest.IsValid || manifest.PackageKind != "opc" || manifest.Facts.MainDocumentUri is null
            || manifest.OrderedOpcContentDigest is null)
            throw new PackageChangeException(PackageChangeError.InvalidPackage,
                "A complete bounded WordprocessingML package is required for a DOCX snapshot.");
        return manifest;
    }
}

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Restore a verified retained version by appending a new version and reset commit, never
    /// rewriting earlier history. Requires the exact previewed head. The new version's parent
    /// is the current version, RestoredFrom is the selected target, and the epoch advances even
    /// when target content is already current. Does not silently replace any live DocxSession.
    /// Hosts must preserve pre-reset pending work as a recoverable branch when notifying clients.
    /// </summary>
    public ValueTask<DocxHistoryView> RestoreVersionAsync(string documentId, HistoryHead expected,
        HistoryBlobReference targetVersion, DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        RestoreVersionCoreAsync(documentId, expected, targetVersion, metadata, null, cancellationToken);

    /// <summary>Restore with a durable caller-owned retry identity, without repeating the reset.</summary>
    public ValueTask<DocxHistoryView> RestoreVersionAsync(string documentId, string requestId, HistoryHead expected,
        HistoryBlobReference targetVersion, DocxVersionMetadata metadata, CancellationToken cancellationToken = default) =>
        RestoreVersionCoreAsync(documentId, expected, targetVersion, metadata,
            requestId ?? throw new ArgumentNullException(nameof(requestId)), cancellationToken);

    private async ValueTask<DocxHistoryView> RestoreVersionCoreAsync(string documentId, HistoryHead expected,
        HistoryBlobReference targetVersion, DocxVersionMetadata metadata, string? requestId, CancellationToken cancellationToken)
    {
        HistoryHeadCodec.Key(documentId);
        ArgumentNullException.ThrowIfNull(expected);
        ArgumentNullException.ThrowIfNull(targetVersion);
        cancellationToken.ThrowIfCancellationRequested();
        var capturedMetadata = HistoryRecordStore.PrepareMetadata(metadata);
        var request = requestId is null ? null : VersionRequest("restore", documentId, expected,
            capturedMetadata, requestId, null, targetVersion);
        var nonce = Guid.NewGuid();
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false);
        var repeated = await FindRequestAsync(documentId, current, request, cancellationToken).ConfigureAwait(false);
        if (repeated is not null) return repeated;
        if (current?.Head != expected) throw Stale();
        var target = await GetVersionAsync(documentId, targetVersion, cancellationToken).ConfigureAwait(false);
        // Verify exact target bytes and package identity before publishing a reset. Reuse its
        // immutable snapshot reference; restore never reserializes or rewrites the target blob.
        _ = await _snapshots.ExportAsync(target.Record.Snapshot, cancellationToken).ConfigureAwait(false);
        var sequence = checked(current!.State.Sequence + 1);
        var epoch = checked(current.State.Epoch + 1);
        var version = new DocxVersionRecord
        {
            DocumentId = documentId, Metadata = capturedMetadata, Nonce = nonce, Parent = current.State.Version,
            RestoredFrom = targetVersion, Sequence = sequence, Snapshot = target.Record.Snapshot,
        };
        var versionId = await _records.SaveVersionAsync(version, cancellationToken).ConfigureAwait(false);
        var commitId = await _records.SaveCommitAsync(new PackageHistoryCommitRecord
        {
            After = target.Record.Snapshot, Before = current.State.Snapshot, Contribution = null,
            DocumentId = documentId, Epoch = epoch, Kind = "restore", Parent = current.State.Commit,
            Sequence = sequence, Version = versionId,
        }, cancellationToken).ConfigureAwait(false);
        var state = current.State with
        {
            Commit = commitId, Epoch = epoch, Sequence = sequence, Snapshot = target.Record.Snapshot, Version = versionId,
        };
        return await PublishAsync(documentId, expected, state, new DocxStoredVersion(versionId, version),
            cancellationToken, request, current.State.Requests).ConfigureAwait(false);
    }
}

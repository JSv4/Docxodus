// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Runtime.CompilerServices;

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Materialize recorded content at a sequence using the coarse producer's exact snapshot
    /// checkpoints. Every import/restore currently has one. This is not a named-version download:
    /// multiple ZIP serializations can share a content sequence. Use ExportVersionAsync for an
    /// exact named version. The metadata scan is bounded and cancellable; no history is discarded.
    /// </summary>
    public async ValueTask<byte[]> MaterializeAsync(string documentId, long sequence,
        int maxCommitsToScan = 10_000, CancellationToken cancellationToken = default)
    {
        var current = await ReadForReconstructionAsync(documentId, sequence, maxCommitsToScan, cancellationToken).ConfigureAwait(false);
        if (sequence == current.State.Sequence)
            return await _snapshots.ExportAsync(current.State.Snapshot, cancellationToken).ConfigureAwait(false);
        if (sequence == 0)
            return await _snapshots.ExportAsync(current.State.InitialSnapshot, cancellationToken).ConfigureAwait(false);
        await foreach (var commit in ReadBackwardsAsync(documentId, current, maxCommitsToScan, cancellationToken).ConfigureAwait(false))
            if (commit.Sequence == sequence)
                return await _snapshots.ExportAsync(commit.After, cancellationToken).ConfigureAwait(false);
        throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "The requested sequence is unavailable.");
    }

    /// <summary>
    /// Audit retained metadata ancestry and reconstruct a sequence from its initial checkpoint
    /// using recorded package effects and restore snapshots, never high-level commands/clocks/IDs.
    /// This deliberately slower validation path also works without intermediate import snapshots.
    /// Output preserves OPC content, not the original ZIP serialization; it is never published.
    /// </summary>
    public async ValueTask<byte[]> ReplayAsync(string documentId, long sequence,
        int maxCommitsToScan = 10_000, CancellationToken cancellationToken = default)
    {
        var current = await ReadForReconstructionAsync(documentId, sequence, maxCommitsToScan, cancellationToken).ConfigureAwait(false);
        var commits = new List<PackageHistoryCommitRecord>();
        await foreach (var commit in ReadBackwardsAsync(documentId, current, maxCommitsToScan, cancellationToken).ConfigureAwait(false))
            if (commit.Sequence <= sequence) commits.Add(commit);
        commits.Reverse();
        var package = await _snapshots.ExportAsync(current.State.InitialSnapshot, cancellationToken).ConfigureAwait(false);
        var contentDigest = current.State.InitialSnapshot.ContentDigest;
        foreach (var commit in commits)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Consistent(contentDigest == commit.Before.ContentDigest, "Replay base does not match the next commit.");
            if (commit.Kind == "restore")
                package = await _snapshots.ExportAsync(commit.After, cancellationToken).ConfigureAwait(false);
            else
            {
                var limits = new PackageChangeLimits();
                var manifest = await HistoryBlobIO.ReadBytesAsync(_blobs, commit.Contribution!, limits.MaxManifestBytes,
                    cancellationToken).ConfigureAwait(false);
                var changes = await PackageChangeSetCodec.LoadAsync(manifest, _blobs, limits, cancellationToken).ConfigureAwait(false);
                Consistent(changes.BeforeDigest == commit.Before.ContentDigest && changes.AfterDigest == commit.After.ContentDigest,
                    "Recorded effects disagree with the commit endpoints.");
                try { package = changes.Apply(package, _packageOptions); }
                catch (PackageChangeException error) when (error.Code == PackageChangeError.BaseMismatch)
                {
                    // Metadata-only repacks may add/drop empty ZIP directories excluded from
                    // content identity. Align those artifacts to the exact before checkpoint.
                    // Both endpoints remain digest-guarded; a false before payload still fails.
                    var checkpoint = await _snapshots.ExportAsync(commit.Before, cancellationToken).ConfigureAwait(false);
                    package = changes.Apply(checkpoint, _packageOptions);
                }
            }
            contentDigest = commit.After.ContentDigest;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return package;
    }

    private async ValueTask<DocxHistoryView> ReadForReconstructionAsync(string documentId, long sequence,
        int maximum, CancellationToken cancellationToken)
    {
        if (sequence < 0) throw new ArgumentOutOfRangeException(nameof(sequence));
        if (maximum <= 0) throw new ArgumentOutOfRangeException(nameof(maximum));
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        if (sequence > current.State.Sequence) throw new ArgumentOutOfRangeException(nameof(sequence));
        return current;
    }

    private async IAsyncEnumerable<PackageHistoryCommitRecord> ReadBackwardsAsync(string documentId,
        DocxHistoryView current, int maximum, [EnumeratorCancellation] CancellationToken cancellationToken)
    {
        var cursor = current.State.Commit;
        var sequence = current.State.Sequence;
        var epoch = current.State.Epoch;
        var contentDigest = current.State.Snapshot.ContentDigest;
        var scanned = 0;
        while (cursor is not null)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned++ >= maximum)
                throw new DocxHistoryException(DocxHistoryError.TraversalLimit, "History scan budget reached; increase it explicitly to continue.");
            var commit = await _records.LoadCommitAsync(cursor, cancellationToken).ConfigureAwait(false);
            SameDocument(documentId, commit.DocumentId);
            Consistent(commit.Sequence == sequence && commit.Epoch == epoch && commit.After.ContentDigest == contentDigest,
                "Commit sequence, epoch or content chain is discontinuous.");
            var version = await GetVersionAsync(documentId, commit.Version, cancellationToken).ConfigureAwait(false);
            Consistent(version.Record.Sequence == sequence && version.Record.Snapshot == commit.After
                && version.Record.Parent is not null
                && (version.Record.RestoredFrom is not null) == (commit.Kind == "restore"),
                "Commit disagrees with its recorded version.");
            var parentVersion = await GetVersionAsync(documentId, version.Record.Parent!, cancellationToken).ConfigureAwait(false);
            Consistent(parentVersion.Record.Sequence == sequence - 1 && parentVersion.Record.Snapshot == commit.Before,
                "Commit disagrees with its preceding version checkpoint.");
            if (commit.Kind == "restore")
            {
                var target = await GetVersionAsync(documentId, version.Record.RestoredFrom!, cancellationToken).ConfigureAwait(false);
                Consistent(target.Record.Snapshot == commit.After, "Restore target disagrees with its recorded snapshot.");
            }
            yield return commit;
            cursor = commit.Parent;
            sequence--;
            if (commit.Kind == "restore") epoch--;
            contentDigest = commit.Before.ContentDigest;
        }
        Consistent(sequence == 0 && epoch == 0 && contentDigest == current.State.InitialSnapshot.ContentDigest,
            "Commit chain does not reach its initial checkpoint.");
    }
}

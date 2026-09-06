// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

/// <summary>An accepted content commit with its recorded host timestamp and author metadata.</summary>
public sealed record DocxHistoryLogEntry(HistoryBlobReference Id, PackageHistoryCommitRecord Commit, DocxVersionMetadata Metadata);

/// <summary>
/// An immutable captured head and its ascending, contiguous accepted content tail. A first join
/// (After=null) supplies the latest checkpoint with no historical tail. Reset marks first join
/// or an epoch transition; clients must preserve pending work before installing that checkpoint.
/// </summary>
public sealed record DocxHistoryUpdate(HistoryHead? After, DocxHistoryView View,
    IReadOnlyList<DocxHistoryLogEntry> Entries, bool Reset);

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Read host-triggered live updates since an exact previously accepted head. Validates both
    /// version and commit ancestry, rejecting forks/rewinds even when package content matches.
    /// Reads bounded metadata only; hosts retain blobs and provide their own notifications.
    /// Duplicate reads are empty, labels do not fabricate content entries, and timestamps never
    /// replace sequence order. No transport, timer, subscription, or polling loop is installed.
    /// </summary>
    public async ValueTask<DocxHistoryUpdate> ReadChangesSinceAsync(string documentId, HistoryHead? after,
        int maxEntriesToScan = 10_000, CancellationToken cancellationToken = default)
    {
        if (maxEntriesToScan <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesToScan));
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        return await ReadChangesBetweenAsync(documentId, current, after, maxEntriesToScan, cancellationToken).ConfigureAwait(false);
    }

    private async ValueTask<DocxHistoryUpdate> ReadChangesBetweenAsync(string documentId, DocxHistoryView current,
        HistoryHead? after, int maxEntriesToScan, CancellationToken cancellationToken)
    {
        if (after is null) return new DocxHistoryUpdate(null, current, Array.Empty<DocxHistoryLogEntry>(), true);
        if (current.Head == after) return new DocxHistoryUpdate(after, current, Array.Empty<DocxHistoryLogEntry>(), false);
        Consistent(current.Head.Revision > after.Revision, "The supplied history head rewinds or forks the accepted publication.");
        var previous = await ReadHeadViewAsync(documentId, after, cancellationToken).ConfigureAwait(false);
        Consistent(current.State.InitialSnapshot == previous.State.InitialSnapshot
            && current.State.Sequence >= previous.State.Sequence && current.State.Epoch >= previous.State.Epoch,
            "History updates do not extend the accepted checkpoint.");

        var remaining = maxEntriesToScan;
        void Spend()
        {
            if (remaining-- <= 0) throw new DocxHistoryException(DocxHistoryError.TraversalLimit,
                "History update scan budget reached; increase it explicitly to continue.");
        }

        // V3 binds every publication, including a conflict/no-op decision that creates no version.
        // Follow exact parent heads before falling back to V1/V2's one-version-per-publication chain.
        var publication = current;
        while (publication.Head != after && publication.State.ParentPublication is { } parentHead)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Spend();
            Consistent(parentHead.Revision >= after.Revision, "Accepted publication is not an ancestor of the new head.");
            var parent = await ReadHeadViewAsync(documentId, parentHead, cancellationToken).ConfigureAwait(false);
            ValidatePublicationEdge(publication, parent);
            publication = parent;
        }
        Consistent(publication.Head == after || (publication.Head.Revision > after.Revision && previous.State.Operation is null),
            "Accepted publication is not an ancestor of the new head.");

        // Content sequence alone cannot prove ancestry: competing labels and same-content forks
        // also have distinct immutable version chains. Check the accepted legacy version ID.
        var version = publication.Version;
        long versionEdges = 0;
        while (version.Id != previous.Version.Id)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Spend();
            versionEdges++;
            Consistent(version.Record.Parent is not null, "Accepted version is not an ancestor of the new head.");
            var parent = await GetVersionAsync(documentId, version.Record.Parent!, cancellationToken).ConfigureAwait(false);
            Consistent(parent.Record.Sequence >= previous.State.Sequence && parent.Record.Sequence <= version.Record.Sequence
                && version.Record.Sequence - parent.Record.Sequence <= 1,
                "Version ancestry is discontinuous.");
            if (parent.Record.Sequence == version.Record.Sequence)
                Consistent(parent.Record.Snapshot.ContentDigest == version.Record.Snapshot.ContentDigest
                    && version.Record.RestoredFrom is null, "A metadata-only version cannot change content or restore.");
            version = parent;
        }
        Consistent(versionEdges == publication.Head.Revision - after.Revision,
            "Publication revision does not match the accepted version ancestry.");

        var entries = new List<DocxHistoryLogEntry>();
        if (current.State.Sequence == previous.State.Sequence)
        {
            Consistent(current.State.Commit == previous.State.Commit && current.State.Epoch == previous.State.Epoch
                && current.State.Snapshot.ContentDigest == previous.State.Snapshot.ContentDigest,
                "Metadata-only updates disagree with the accepted content tip.");
        }
        else
        {
            Spend(); // Reserve the first commit; the iterator checks its own remaining budget.
            await foreach (var entry in ReadBackwardsAsync(documentId, current, remaining + 1, cancellationToken).ConfigureAwait(false))
            {
                entries.Add(new DocxHistoryLogEntry(entry.Id, entry.Commit, entry.Version.Metadata));
                if (entry.Commit.Sequence == previous.State.Sequence + 1)
                {
                    Consistent(entry.Commit.Parent == previous.State.Commit
                        && entry.Commit.Before.ContentDigest == previous.State.Snapshot.ContentDigest
                        && entry.Commit.Epoch - (entry.Commit.Kind == "restore" ? 1 : 0) == previous.State.Epoch,
                        "Commit tail does not extend the accepted content tip.");
                    break;
                }
            }
            Consistent(entries.Count == current.State.Sequence - previous.State.Sequence,
                "The accepted content tail has a gap.");
            entries.Reverse();
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new DocxHistoryUpdate(after, current, entries.AsReadOnly(), current.State.Epoch != previous.State.Epoch);
    }

    internal static void ValidatePublicationEdge(DocxHistoryView child, DocxHistoryView parent)
    {
        Consistent(child.State.InitialSnapshot == parent.State.InitialSnapshot
            && child.State.Sequence >= parent.State.Sequence && child.State.Sequence - parent.State.Sequence <= 1
            && child.State.Epoch >= parent.State.Epoch && child.State.Epoch - parent.State.Epoch <= 1,
            "Publication ancestry is discontinuous.");
        if (child.Version.Id == parent.Version.Id)
        {
            Consistent(child.State.Snapshot == parent.State.Snapshot && child.State.Sequence == parent.State.Sequence
                && child.State.Commit == parent.State.Commit && child.State.Epoch == parent.State.Epoch
                && child.State.Operation != parent.State.Operation && child.State.Requests?.Current is not null,
                "A version-free publication must record a new identified decision without changing content.");
        }
        else
        {
            Consistent(child.Version.Record.Parent == parent.Version.Id, "Publication skipped its parent version.");
            if (child.State.Sequence == parent.State.Sequence)
                Consistent(child.State.Commit == parent.State.Commit && child.State.Epoch == parent.State.Epoch
                    && child.State.Snapshot.ContentDigest == parent.State.Snapshot.ContentDigest
                    && child.Version.Record.RestoredFrom is null, "A metadata-only publication cannot change content or restore.");
        }
    }
}

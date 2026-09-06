// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

public sealed partial class DocxVersionHistory
{
    /// <summary>
    /// Resolve the greatest committed content sequence whose host-recorded creation time is at
    /// or before the cutoff. Sequence breaks ties and clocks need not be monotonic. Metadata-only
    /// named versions do not invent content history. Sequence zero uses the initial version's
    /// timestamp, not a later label's timestamp. The host owns timestamp authority and provenance.
    /// </summary>
    /// <remarks>
    /// The budget counts traversed commits plus initial-sequence version ancestors; each entry
    /// can require several bounded metadata reads. No snapshot/effect payloads are read. Pass the
    /// result to MaterializeAsync or ReplayAsync for content. The lookup captures one immutable head.
    /// </remarks>
    public async ValueTask<long> ResolveSequenceAtTimeAsync(string documentId, DateTimeOffset cutoff,
        int maxEntriesToScan = 10_000, CancellationToken cancellationToken = default)
    {
        HistoryHeadCodec.Key(documentId);
        cancellationToken.ThrowIfCancellationRequested();
        if (maxEntriesToScan <= 0) throw new ArgumentOutOfRangeException(nameof(maxEntriesToScan));
        var current = await ReadAsync(documentId, cancellationToken).ConfigureAwait(false)
            ?? throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "Document history is unavailable.");
        var initialCursor = current.Version.Id;
        var scanned = 0;
        await foreach (var entry in ReadBackwardsAsync(documentId, current, maxEntriesToScan, cancellationToken).ConfigureAwait(false))
        {
            scanned++;
            if (entry.Version.Metadata.CreatedAt <= cutoff) return entry.Commit.Sequence;
            initialCursor = entry.Version.Parent!;
        }

        // Before the first content commit there may be any number of metadata-only versions.
        // Their timestamps cannot move the beginning of the recorded document history.
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned++ >= maxEntriesToScan)
                throw new DocxHistoryException(DocxHistoryError.TraversalLimit, "History scan budget reached; increase it explicitly to continue.");
            var initial = await GetVersionAsync(documentId, initialCursor, cancellationToken).ConfigureAwait(false);
            Consistent(initial.Record.Sequence == 0 && initial.Record.Snapshot.ContentDigest == current.State.InitialSnapshot.ContentDigest,
                "Initial version ancestry disagrees with its checkpoint.");
            if (initial.Record.Parent is null)
            {
                if (initial.Record.Metadata.CreatedAt <= cutoff) return 0;
                throw new DocxHistoryException(DocxHistoryError.HistoryUnavailable, "No content state was recorded at or before that time.");
            }
            initialCursor = initial.Record.Parent;
        }
    }
}

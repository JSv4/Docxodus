// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

/// <summary>
/// A monotonic publication revision and immutable state-manifest reference. Publication revision
/// is not a document edit/log sequence: named versions can update metadata without a content edit.
/// </summary>
public sealed record HistoryHead(long Revision, HistoryBlobReference State);

/// <summary>
/// Host-owned compare-and-swap publication of one document's state manifest. Persist its complete
/// blob graph first, then advance the head once. A manifest can atomically bind log, checkpoint,
/// version, and deduplication tips without separate mutable indexes. Hosts enforce authorization.
/// </summary>
public interface IHistoryHeadStore
{
    ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default);

    /// <summary>
    /// Publish only if the current head equals expected (null means absent). Return the new head,
    /// or null for a stale expectation. Each success increments Revision even for identical State,
    /// preventing ABA. Publication is the commit point; cancellation after it cannot undo success.
    /// The caller owns validation and durability of the referenced immutable state and its blobs.
    /// </summary>
    ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
        HistoryBlobReference state, CancellationToken cancellationToken = default);
}

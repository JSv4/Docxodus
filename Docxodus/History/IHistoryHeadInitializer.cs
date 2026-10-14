// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>Initialized is true only for this call's insertion; Head is the atomically observed result.</summary>
public sealed record HistoryHeadInitializationResult(bool Initialized, HistoryHead Head);

/// <summary>
/// Optional storage capability for importing exact publication revisions. Must share atomic writer
/// exclusion with TryAdvanceAsync. Never replace an existing head, including an older one.
/// The caller validates and persists all immutable blobs before invoking this commit point.
/// </summary>
public interface IHistoryHeadInitializer : IHistoryHeadStore
{
    /// <summary>
    /// Insert the supplied positive-revision head only when absent; otherwise return the existing
    /// head unchanged. Exact equality means an idempotent retry, not a new publication. Do not
    /// check cancellation after insertion. Storage failures can lose acknowledgements; retry safely.
    /// </summary>
    ValueTask<HistoryHeadInitializationResult> TryInitializeAsync(string documentId, HistoryHead head,
        CancellationToken cancellationToken = default);
}

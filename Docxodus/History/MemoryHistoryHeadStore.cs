// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.Concurrent;

namespace Docxodus.History;

/// <summary>Thread-safe process-local compare-and-swap heads; hosts own aggregate quota/lifetime.</summary>
public sealed class MemoryHistoryHeadStore : IHistoryHeadStore
{
    private readonly ConcurrentDictionary<string, HistoryHead> _heads = new(StringComparer.Ordinal);

    public ValueTask<HistoryHead?> ReadAsync(string documentId, CancellationToken cancellationToken = default)
    {
        var key = HistoryHeadCodec.Key(documentId);
        cancellationToken.ThrowIfCancellationRequested();
        return ValueTask.FromResult(_heads.GetValueOrDefault(key));
    }

    public ValueTask<HistoryHead?> TryAdvanceAsync(string documentId, HistoryHead? expected,
        HistoryBlobReference state, CancellationToken cancellationToken = default)
    {
        var key = HistoryHeadCodec.Key(documentId);
        var next = HistoryHeadCodec.Next(expected, state);
        cancellationToken.ThrowIfCancellationRequested();
        var success = expected is null ? _heads.TryAdd(key, next) : _heads.TryUpdate(key, next, expected);
        return ValueTask.FromResult(success ? next : null);
    }
}

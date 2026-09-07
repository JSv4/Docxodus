// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Collections.Concurrent;
using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Thread-safe, process-local reference store. Hosts own its lifetime and aggregate quota.</summary>
public sealed class MemoryHistoryBlobStore(int maxBlobBytes = HistoryBlobIO.DefaultMaxBlobBytes) : IHistoryBlobStore
{
    private readonly int _maxBlobBytes = HistoryBlobIO.ValidateLimit(maxBlobBytes);
    private readonly ConcurrentDictionary<VerificationDigest, byte[]> _blobs = new();

    public async ValueTask PutAsync(HistoryBlobReference reference, Stream content,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(content);
        HistoryBlobIO.Validate(reference, _maxBlobBytes);
        cancellationToken.ThrowIfCancellationRequested();
        var bytes = await PackageChangeSetCodec.ReadPayloadAsync(reference, content, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        _blobs.TryAdd(reference.Digest, bytes);
    }

    public ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference,
        CancellationToken cancellationToken = default)
    {
        HistoryBlobIO.Validate(reference, _maxBlobBytes);
        cancellationToken.ThrowIfCancellationRequested();
        if (!_blobs.TryGetValue(reference.Digest, out var bytes)) return ValueTask.FromResult<Stream?>(null);
        if (bytes.Length != reference.Length)
            throw new PackageChangeException(PackageChangeError.PayloadMismatch, "Stored blob length does not match its reference.");
        return ValueTask.FromResult<Stream?>(new MemoryStream(bytes, writable: false));
    }
}

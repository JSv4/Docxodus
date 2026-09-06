// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.History;

/// <summary>Content-addressed bytes. Length is the exact uncompressed blob length.</summary>
public sealed record HistoryBlobReference(VerificationDigest Digest, int Length);

/// <summary>
/// Host-owned immutable payload storage. Implementations must verify length and SHA-256 before
/// publishing a blob and make repeated writes of the same content idempotent. A successful put
/// makes complete bytes available to subsequent reads; durability is the host's storage policy.
/// </summary>
public interface IHistoryBlobStore
{
    /// <summary>Does not close or retain the caller-owned content stream after completion.</summary>
    ValueTask PutAsync(HistoryBlobReference reference, Stream content, CancellationToken cancellationToken = default);

    /// <summary>Returns a caller-owned readable stream, or null if the blob is missing.</summary>
    ValueTask<Stream?> OpenReadAsync(HistoryBlobReference reference, CancellationToken cancellationToken = default);
}

/// <summary>Limits checked before loading payloads from a package-change manifest.</summary>
public sealed record PackageChangeLimits
{
    public int MaxManifestBytes { get; init; } = 16 * 1024 * 1024;
    public int MaxChanges { get; init; } = 20_000;
    public int MaxEntryNameLength { get; init; } = 4_096;
    public int MaxPayloadBytes { get; init; } = 256 * 1024 * 1024;
    public long MaxTotalPayloadBytes { get; init; } = 512L * 1024 * 1024;

    internal void Validate()
    {
        if (MaxManifestBytes <= 0 || MaxChanges <= 0 || MaxEntryNameLength <= 0
            || MaxPayloadBytes <= 0 || MaxTotalPayloadBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(PackageChangeLimits), "All limits must be positive.");
    }
}

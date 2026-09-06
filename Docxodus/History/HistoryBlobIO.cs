// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;

namespace Docxodus.History;

internal static class HistoryBlobIO
{
    internal const int DefaultMaxBlobBytes = 256 * 1024 * 1024;

    internal static void Validate(HistoryBlobReference reference, int maximumBytes)
    {
        ArgumentNullException.ThrowIfNull(reference);
        if (reference.Digest is not { Algorithm: "SHA-256", Value.Length: 64 }
            || !reference.Digest.Value.All(c => c is >= '0' and <= '9' or >= 'a' and <= 'f')
            || reference.Length < 0)
            throw new ArgumentException("Expected a lowercase SHA-256 digest and nonnegative length.", nameof(reference));
        if (reference.Length > maximumBytes)
            throw new PackageChangeException(PackageChangeError.ResourceLimit, "Blob exceeds the store's byte limit.");
    }

    internal static int ValidateLimit(int maximumBytes) => maximumBytes > 0 ? maximumBytes
        : throw new ArgumentOutOfRangeException(nameof(maximumBytes));

    // Used for filesystem writes and collision checks without allocating a blob-sized buffer.
    internal static async ValueTask CopyVerifiedAsync(HistoryBlobReference reference, Stream input,
        Stream output, CancellationToken cancellationToken)
    {
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        var buffer = new byte[64 * 1024];
        long length = 0;
        while (true)
        {
            var read = await input.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            length += read;
            if (length > reference.Length) throw Mismatch();
            hash.AppendData(buffer, 0, read);
            await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken).ConfigureAwait(false);
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (length != reference.Length
            || Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant() != reference.Digest.Value)
            throw Mismatch();
    }

    private static PackageChangeException Mismatch() => new(PackageChangeError.PayloadMismatch,
        "Blob content does not match its length and SHA-256 reference.");
}

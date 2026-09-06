// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Docxodus.History;

/// <summary>Explicit archive resource budgets; hosts may choose smaller limits for untrusted uploads.</summary>
public sealed record DocxHistoryArchiveLimits
{
    public long MaxArchiveBytes { get; init; } = 1024L * 1024 * 1024;
    public long MaxTotalBlobBytes { get; init; } = 1024L * 1024 * 1024;
    public long MaxMetadataBytes { get; init; } = 64L * 1024 * 1024;
    /// <summary>Total raw bytes visited during validation, including repeated reads for effect verification.</summary>
    public long MaxValidationBytes { get; init; } = 4L * 1024 * 1024 * 1024;
    /// <summary>Conservative aggregate package-expansion work allowance, reserved before repeated package processing.</summary>
    public long MaxExpandedBytes { get; init; } = 32L * 1024 * 1024 * 1024;
    public int MaxBlobs { get; init; } = 100_000;
    public int MaxEdges { get; init; } = 1_000_000;
    public int MaxManifestBytes { get; init; } = 16 * 1024 * 1024;
    public int MaxBlobBytes { get; init; } = HistoryBlobIO.DefaultMaxBlobBytes;
    public int MaxRecordBytes { get; init; } = HistoryRecordStore.DefaultMaxRecordBytes;

    internal void Validate()
    {
        if (MaxArchiveBytes <= 0 || MaxTotalBlobBytes <= 0 || MaxMetadataBytes <= 0 || MaxValidationBytes <= 0 || MaxExpandedBytes <= 0
            || MaxBlobs <= 0 || MaxEdges <= 0 || MaxManifestBytes <= 0 || MaxBlobBytes <= 0 || MaxRecordBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(DocxHistoryArchiveLimits), "Archive limits must be positive.");
    }
}

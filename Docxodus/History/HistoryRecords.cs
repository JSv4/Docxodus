// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.History;

/// <summary>Host-supplied version metadata; application actor identity is separate from Word revision authors.</summary>
public sealed record DocxVersionMetadata
{
    public IReadOnlyDictionary<string, string> ApplicationMetadata { get; init; } = new Dictionary<string, string>();
    public required string Author { get; init; }
    public required DateTimeOffset CreatedAt { get; init; }
    public string? Label { get; init; }
    public string? Message { get; init; }
}

/// <summary>
/// Immutable named version. Its stored manifest blob reference is its content-addressed version ID.
/// Nonce is assigned once at creation so distinct versions may retain identical snapshots/metadata.
/// </summary>
public sealed record DocxVersionRecord
{
    public required string DocumentId { get; init; }
    public required DocxVersionMetadata Metadata { get; init; }
    public required Guid Nonce { get; init; }
    public required HistoryBlobReference? Parent { get; init; }
    public required HistoryBlobReference? RestoredFrom { get; init; }
    public required long Sequence { get; init; }
    public required DocxSnapshotReference Snapshot { get; init; }
}

/// <summary>
/// Coarse import/restore commit, not a live keystroke effect. Snapshots are exact checkpoint
/// boundaries. An import references the reversible package-change manifest; a restore installs
/// its captured target and advances the epoch. Fine-grained accepted effects use a later codec.
/// </summary>
public sealed record PackageHistoryCommitRecord
{
    public required DocxSnapshotReference After { get; init; }
    public required DocxSnapshotReference Before { get; init; }
    public required HistoryBlobReference? Contribution { get; init; }
    public required string DocumentId { get; init; }
    public required long Epoch { get; init; }
    /// <summary>Exactly "import" or "restore".</summary>
    public required string Kind { get; init; }
    public required HistoryBlobReference? Parent { get; init; }
    public required long Sequence { get; init; }
    public required HistoryBlobReference Version { get; init; }
}

/// <summary>
/// One immutable head manifest for the coarse package-history producer. A single CAS publishes
/// its content, commit and version tips together. Sequence counts content commits; publishing
/// a named version of unchanged content advances only the separate HistoryHead.Revision.
/// </summary>
public sealed record DocxHistoryStateRecord
{
    /// <summary>V2 publication metadata. Absent in legacy V1 states; older records remain readable.</summary>
    [System.Text.Json.Serialization.JsonIgnore(Condition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull)]
    public HistoryRequestJournal? Requests { get; init; }
    public required HistoryBlobReference? Commit { get; init; }
    public required string DocumentId { get; init; }
    public required long Epoch { get; init; }
    public required DocxSnapshotReference InitialSnapshot { get; init; }
    public required long Sequence { get; init; }
    public required DocxSnapshotReference Snapshot { get; init; }
    public required HistoryBlobReference Version { get; init; }
}

#nullable enable

namespace Docxodus.History;

/// <summary>
/// A splice in a w:t element addressed relative to an immutable base publication. TextNode is
/// the zero-based document-order w:t ordinal in PartUri; offsets/counts are UTF-16 boundaries.
/// Known splices preserve that topology. Unmodeled edits of the part form a conflict boundary.
/// </summary>
public sealed record DocxTextSplice(string PartUri, int TextNode, int Offset, int DeleteCount, string Insert);

/// <summary>Host-owned durable intent. Capture/persist once before submitting or retrying.</summary>
public sealed record DocxOperationRequest
{
    public required string RequestId { get; init; }
    public required HistoryHead Base { get; init; }
    /// <summary>Exactly text, package, or discard. Package takes candidate DOCX bytes separately.</summary>
    public required string Kind { get; init; }
    public required DocxVersionMetadata Metadata { get; init; }
    public DocxTextSplice? Text { get; init; }
    /// <summary>Additional whole-part read dependencies, including expected absence at the base.</summary>
    public IReadOnlyList<string> ReadParts { get; init; } = Array.Empty<string>();
    /// <summary>An earlier published conflict explicitly resolved by this NEW request.</summary>
    public HistoryBlobReference? Resolves { get; init; }
}

/// <summary>Canonical immutable input; its manifest digest is the request fingerprint.</summary>
public sealed record DocxOperationInput(string DocumentId, DocxOperationRequest Request, HistoryBlobReference? Candidate);

/// <summary>
/// Durable acceptance/conflict outcome. Revision orders decisions and ordinary publications;
/// ContentCommit is non-null only for actual accepted effects. Conflict preserves ProposedSnapshot
/// without changing the visible document. Parent links only decisions, Before links publications.
/// </summary>
public sealed record DocxOperationRecord
{
    public required string DocumentId { get; init; }
    public required HistoryBlobReference Input { get; init; }
    public required HistoryBlobReference? Parent { get; init; }
    public required long Revision { get; init; }
    public required HistoryHead Before { get; init; }
    public required DocxSnapshotReference ProposedSnapshot { get; init; }
    public required DocxSnapshotReference AfterSnapshot { get; init; }
    public required HistoryBlobReference Version { get; init; }
    public required HistoryBlobReference? ContentCommit { get; init; }
    /// <summary>Exactly accepted or conflict.</summary>
    public required string Status { get; init; }
    public required string? Conflict { get; init; }
    public required DocxTextSplice? AppliedText { get; init; }
}

public sealed record DocxStoredOperation(HistoryBlobReference Id, DocxOperationRecord Record, DocxOperationInput Input);
public sealed record DocxOperationResult(DocxHistoryView View, DocxStoredOperation Operation);
public sealed record DocxOperationUpdate(HistoryHead? After, DocxHistoryView View, IReadOnlyList<DocxStoredOperation> Operations);

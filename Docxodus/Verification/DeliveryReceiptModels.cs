// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using System.Text.Json.Serialization;

namespace Docxodus.Verification;

/// <summary>Controls how much document-derived text a delivery receipt retains.</summary>
public enum DeliveryReceiptPrivacyProfile
{
    /// <summary>Retain identities, counts, and digests, but no text or text summaries.</summary>
    HashOnly,

    /// <summary>Retain identities, counts, digests, and structural summaries without text values.</summary>
    HashAndSummary,

    /// <summary>Retain complete normalized operation and result evidence.</summary>
    FullEvidence,
}

/// <summary>The terminal meaning of a mutation contribution.</summary>
public enum DeliveryTransactionStatus
{
    Committed,
    PartiallyCommitted,
    Failed,
    Prediction,
}

/// <summary>Whether one requested operation produced retained, failed, or rolled-back evidence.</summary>
public enum DeliveryOperationExecutionStatus
{
    NotExecuted,
    Succeeded,
    Failed,
    SucceededRolledBack,
    FailedRolledBack,
}

/// <summary>A reversible transition in delivered-document history.</summary>
public enum DeliveryLineageAction
{
    Undo,
    Redo,
}

/// <summary>Why an observed package change exists.</summary>
public enum DeliveryChangeDisposition
{
    UserRequested,
    Derived,
    Unexpected,
}

/// <summary>Receipt-level package change families derived from two #456 manifests.</summary>
public enum DeliveryPackageChangeKind
{
    PartAdded,
    PartRemoved,
    PartModified,
    RelationshipAdded,
    RelationshipRemoved,
    RelationshipModified,
}

/// <summary>Stable roles for independently verifiable delivery artifacts.</summary>
public enum DeliveryArtifactRole
{
    CleanDocx,
    ReviewDocx,
    Html,
    Pdf,
    PageImage,
    PageMap,
    PackageManifest,
    SemanticDiff,
    ValidationReport,
    ReversibilityProof,
    RenderReport,
    OtherReport,
}

/// <summary>Whether artifact bytes were produced.</summary>
public enum DeliveryArtifactAvailability
{
    Available,
    Unavailable,
}

/// <summary>External evidence consumed by, but not redefined inside, a receipt.</summary>
public enum DeliveryEvidenceKind
{
    SemanticChangeSet,
    ValidationResult,
    RedlineReversibility,
}

/// <summary>The document transition covered by one typed semantic change-set artifact.</summary>
public enum DeliverySemanticComparisonScope
{
    SourceToDelivered,
    Transaction,
}

/// <summary>Mutation-result entities that retain authorship evidence.</summary>
public enum DeliveryAuthoredEntityKind
{
    Revision,
    Comment,
    Annotation,
}

/// <summary>How an entity or anchor changed in one transaction.</summary>
public enum DeliveryObjectChangeKind
{
    Added,
    Removed,
    Modified,
}

/// <summary>
/// Exact identity of one package-backed document version. Digest values use the #456
/// package-manifest contract; the clean delivery identity is independently recomputed from bytes.
/// </summary>
public sealed record DeliveryDocumentIdentity
{
    required public long DocumentVersion { get; init; }
    required public string PackageKind { get; init; }
    required public string PackageManifestSchema { get; init; }
    required public string MainDocumentUri { get; init; }
    required public VerificationDigest RawPackageBytesDigest { get; init; }
    public VerificationDigest? OrderedOpcContentDigest { get; init; }
    public VerificationDigest? NormalizedSemanticDigest { get; init; }

    public static DeliveryDocumentIdentity FromManifest(PackageManifest manifest, long documentVersion)
        => DeliveryPackageManifestAdapter.CreateIdentity(manifest, documentVersion);
}

/// <summary>A privacy-aware text value. Hash and character count are always retained.</summary>
public sealed record DeliveryTextEvidence
{
    required public VerificationDigest Digest { get; init; }
    required public int CharacterCount { get; init; }
    public string? Summary { get; init; }
    public string? Value { get; init; }
}

/// <summary>A stable anchor/object identity reported by an edit result.</summary>
public sealed record DeliveryObjectChange
{
    required public DeliveryObjectChangeKind ChangeKind { get; init; }
    required public string AnchorId { get; init; }
    required public string Kind { get; init; }
    required public string Scope { get; init; }
    required public string Unid { get; init; }
}

/// <summary>One structured EditResult nested beneath a normalized requested operation.</summary>
public sealed record DeliveryOperationResultEvidence
{
    required public VerificationDigest ResultDigest { get; init; }
    required public bool Success { get; init; }
    public string? ErrorCode { get; init; }
    public DeliveryTextEvidence? ErrorMessage { get; init; }
    public IReadOnlyList<DeliveryObjectChange> ObjectChanges { get; init; } =
        Array.Empty<DeliveryObjectChange>();
    public JsonElement? FullResult { get; init; }
}

/// <summary>One normalized request and its result evidence, in execution order.</summary>
public sealed record DeliveryOperationEvidence
{
    required public int Index { get; init; }
    required public string Tool { get; init; }
    required public string Action { get; init; }
    required public VerificationDigest ArgumentsDigest { get; init; }
    public string? ArgumentsSummary { get; init; }
    public JsonElement? Arguments { get; init; }
    required public DeliveryOperationExecutionStatus ExecutionStatus { get; init; }
    required public bool Success { get; init; }
    required public bool RolledBack { get; init; }
    public IReadOnlyList<DeliveryOperationResultEvidence> Results { get; init; } =
        Array.Empty<DeliveryOperationResultEvidence>();
}

/// <summary>A privacy-aware diagnostic attached to an owner revision identity.</summary>
public sealed record DeliveryAuthoredDiagnostic
{
    required public string Code { get; init; }
    required public DeliveryTextEvidence Message { get; init; }
}

/// <summary>Revision, comment, or annotation evidence copied from MutationBatchResult.</summary>
public sealed record DeliveryAuthoredChange
{
    required public DeliveryAuthoredEntityKind EntityKind { get; init; }
    required public DeliveryObjectChangeKind ChangeKind { get; init; }
    required public string EntityId { get; init; }
    required public VerificationDigest SourceDigest { get; init; }
    public string? Author { get; init; }
    public string? Date { get; init; }
    public string? DateUtc { get; init; }
    public string? Type { get; init; }
    public RevisionFamily? Family { get; init; }
    public string? PartUri { get; init; }
    public string? Scope { get; init; }
    public string? AnchorId { get; init; }
    public RevisionResolutionStatus? ResolutionStatus { get; init; }
    public DeliveryAuthoredDiagnostic? Diagnostic { get; init; }
    public IReadOnlyList<string> ConstituentIds { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> ConstituentKeys { get; init; } = Array.Empty<string>();
    public IReadOnlyList<string> AffectedAnchorIds { get; init; } = Array.Empty<string>();
    public DeliveryTextEvidence? Text { get; init; }
    public JsonElement? FullEvidence { get; init; }
}

/// <summary>One committed, failed, partial, or predicted batch contribution.</summary>
public sealed record DeliveryTransactionEntry
{
    required public long Sequence { get; init; }
    required public string EntryId { get; init; }
    public string? TransactionId { get; init; }
    required public string RequestFingerprint { get; init; }
    required public MutationBatchMode Mode { get; init; }
    required public DeliveryTransactionStatus Status { get; init; }
    required public long BaseVersion { get; init; }
    required public long ResultVersion { get; init; }
    required public DeliveryDocumentIdentity BeforeDocument { get; init; }
    required public DeliveryDocumentIdentity AfterDocument { get; init; }
    public VerificationDigest? ReportedPackageContentDigest { get; init; }
    public IReadOnlyList<DeliveryOperationEvidence> Operations { get; init; } =
        Array.Empty<DeliveryOperationEvidence>();
    public IReadOnlyList<DeliveryAuthoredChange> AuthoredChanges { get; init; } =
        Array.Empty<DeliveryAuthoredChange>();
    public IReadOnlyList<DeliveryTextEvidence> Warnings { get; init; } =
        Array.Empty<DeliveryTextEvidence>();
}

/// <summary>An undo or redo transition tied to one existing committed transaction entry.</summary>
public sealed record DeliveryLineageEvent
{
    required public long Sequence { get; init; }
    required public DeliveryLineageAction Action { get; init; }
    required public string AffectedEntryId { get; init; }
    required public DeliveryDocumentIdentity BeforeDocument { get; init; }
    required public DeliveryDocumentIdentity AfterDocument { get; init; }
}

/// <summary>One source-to-delivered part or relationship change.</summary>
public sealed record DeliveryPackageChange
{
    required public string ChangeId { get; init; }
    required public DeliveryPackageChangeKind Kind { get; init; }
    required public ChangeLocation Location { get; init; }
    public DeliveryTextEvidence? Before { get; init; }
    public DeliveryTextEvidence? After { get; init; }
    required public DeliveryChangeDisposition Disposition { get; init; }
    public string? TransactionEntryId { get; init; }
    public int? RequestedOperationIndex { get; init; }
    public string? Derivation { get; init; }
}

/// <summary>One independently hash-addressed output artifact.</summary>
public sealed record DeliveryArtifact
{
    required public string ArtifactId { get; init; }
    required public DeliveryArtifactRole Role { get; init; }
    required public string MediaType { get; init; }
    required public DeliveryArtifactAvailability Availability { get; init; }
    public long? ByteLength { get; init; }
    public VerificationDigest? Digest { get; init; }
    public string? RelativePath { get; init; }
    public string? UnavailableReason { get; init; }
    public long? DocumentVersion { get; init; }
    public VerificationDigest? PackageDigest { get; init; }
    public string? RendererFingerprint { get; init; }
    public VerificationDigest? PageMapDigest { get; init; }
}

/// <summary>A schema-and-digest reference to evidence owned by another verification component.</summary>
public sealed record DeliveryEvidenceReference
{
    required public DeliveryEvidenceKind Kind { get; init; }
    required public string Schema { get; init; }
    required public VerificationDigest Digest { get; init; }
    public string? ArtifactId { get; init; }
    public string? Summary { get; init; }
}

/// <summary>
/// Binding from exact #457 canonical bytes to the document transition they compare. The receipt
/// retains only #457 identity metadata and never flattens or renames its change records.
/// </summary>
public sealed record DeliverySemanticChangeSetBinding
{
    required public DeliverySemanticComparisonScope Scope { get; init; }
    public string? TransactionEntryId { get; init; }
    required public DeliveryDocumentIdentity BeforeDocument { get; init; }
    required public DeliveryDocumentIdentity AfterDocument { get; init; }
    required public string Schema { get; init; }
    required public int SchemaVersion { get; init; }
    required public int ChangeCount { get; init; }
    required public VerificationDigest Digest { get; init; }
    required public string ArtifactId { get; init; }
}

/// <summary>A page-map citation bound to exact package and render evidence.</summary>
public sealed record DeliveryPageCitation
{
    required public string AnchorId { get; init; }
    required public string Scope { get; init; }
    required public long DocumentVersion { get; init; }
    required public VerificationDigest PackageDigest { get; init; }
    required public string RendererFingerprint { get; init; }
    required public VerificationDigest PageMapDigest { get; init; }
    required public string PageMapArtifactId { get; init; }
    required public string RenderArtifactId { get; init; }
    required public VerificationDigest RenderArtifactDigest { get; init; }
    public IReadOnlyList<PageMapPage> Pages { get; init; } = Array.Empty<PageMapPage>();
    public IReadOnlyList<PageMapFragment> Fragments { get; init; } =
        Array.Empty<PageMapFragment>();
}

/// <summary>The canonical, digest-covered body of a delivery receipt.</summary>
public sealed record DeliveryChangeReceiptPayload
{
    public const string SchemaId =
        "https://docxodus.dev/schemas/verification/delivery-change-receipt/v1";
    public const string CanonicalizationId = "docxodus-canonical-json-v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    public string Canonicalization { get; init; } = CanonicalizationId;
    required public DeliveryReceiptPrivacyProfile PrivacyProfile { get; init; }
    required public DeliveryDocumentIdentity SourceDocument { get; init; }
    required public DeliveryDocumentIdentity DeliveredDocument { get; init; }
    public IReadOnlyList<DeliveryTransactionEntry> Transactions { get; init; } =
        Array.Empty<DeliveryTransactionEntry>();
    public IReadOnlyList<DeliveryLineageEvent> Lineage { get; init; } =
        Array.Empty<DeliveryLineageEvent>();
    public IReadOnlyList<DeliveryPackageChange> PackageChanges { get; init; } =
        Array.Empty<DeliveryPackageChange>();
    public bool HasUnexpectedChanges { get; init; }
    public IReadOnlyList<DeliveryEvidenceReference> Evidence { get; init; } =
        Array.Empty<DeliveryEvidenceReference>();
    public IReadOnlyList<DeliverySemanticChangeSetBinding> SemanticChangeSets { get; init; } =
        Array.Empty<DeliverySemanticChangeSetBinding>();
    public IReadOnlyList<DeliveryArtifact> Artifacts { get; init; } =
        Array.Empty<DeliveryArtifact>();
    public IReadOnlyList<DeliveryPageCitation> PageCitations { get; init; } =
        Array.Empty<DeliveryPageCitation>();
    public IReadOnlyList<DeliveryTextEvidence> Warnings { get; init; } =
        Array.Empty<DeliveryTextEvidence>();
}

/// <summary>Hash-addressed delivery receipt envelope.</summary>
public sealed record DeliveryChangeReceipt
{
    required public DeliveryChangeReceiptPayload Payload { get; init; }
    required public VerificationDigest ReceiptDigest { get; init; }

    public byte[] ToJsonBytes(bool indented = false) =>
        DeliveryChangeReceiptSerializer.Serialize(this, indented);

    public byte[] ToJsonBytes(DeliveryReceiptLimits limits, bool indented = false) =>
        DeliveryChangeReceiptSerializer.Serialize(this, limits, indented);

    public string ToJson(bool indented = false) =>
        System.Text.Encoding.UTF8.GetString(ToJsonBytes(indented));

    public string ToJson(DeliveryReceiptLimits limits, bool indented = false) =>
        System.Text.Encoding.UTF8.GetString(ToJsonBytes(limits, indented));
}

/// <summary>Trim/AOT-safe metadata for the durable delivery-receipt wire contract.</summary>
[JsonSourceGenerationOptions(
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    DefaultIgnoreCondition = JsonIgnoreCondition.Never)]
[JsonSerializable(typeof(DeliveryChangeReceiptPayload))]
[JsonSerializable(typeof(DeliveryChangeReceipt))]
[JsonSerializable(typeof(DeliveryTransactionEntry))]
internal partial class DeliveryReceiptJsonContext : JsonSerializerContext
{
}

/// <summary>A stable validation error raised while composing a receipt.</summary>
public sealed class DeliveryReceiptValidationException : ArgumentException
{
    public DeliveryReceiptValidationException(string code, string message)
        : base(message)
    {
        Code = code;
    }

    public string Code { get; }
}

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>How delivery treats revisions that were already present in an input document.</summary>
public enum DeliveryRevisionPolicy
{
    Preserve,
    Accept,
    Reject,
}

/// <summary>Explicit revision policy for pre-existing and newly generated revision sets.</summary>
public sealed record DeliveryBundleRevisionPolicy
{
    required public DeliveryRevisionPolicy PreExistingRevisions { get; init; }
    required public DeliveryRevisionPolicy GeneratedRevisions { get; init; }
}

/// <summary>Whether a requested or implicit artifact is required for this bundle.</summary>
public enum DeliveryArtifactRequiredness
{
    Required,
    Optional,
}

/// <summary>Whether an artifact came from the caller's request or was added by orchestration.</summary>
public enum DeliveryArtifactProvenance
{
    Requested,
    Implicit,
}

/// <summary>Whether artifact bytes were produced.</summary>
public enum DeliveryArtifactAvailability
{
    Available,
    Unavailable,
}

/// <summary>Stable delivery artifact vocabulary. New meanings require new enum members.</summary>
public enum DeliveryArtifactKind
{
    BaselineDocx,
    PolicyBaselineDocx,
    WorkingDocx,
    ReviewDocx,
    FinalDocx,
    StandaloneHtml,
    FinalPdf,
    ReviewPdf,
    PageMap,
    RenderReport,
    BaselinePackageManifest,
    FinalPackageManifest,
    SemanticDelta,
    PackageDelta,
    ValidationReport,
    ReversibilityProof,
    ChangeReceipt,
}

/// <summary>Shared revision/comment presentation vocabulary for rendered artifacts.</summary>
public enum DeliveryReviewProfile
{
    Final,
    Original,
    Markup,
}

/// <summary>Orthogonal presentation of Word review comments in a rendered artifact.</summary>
public enum DeliveryCommentProfile
{
    Hidden,
    Inline,
    Endnotes,
    Margin,
}

/// <summary>Overall result of assembling a delivery bundle.</summary>
public enum DeliveryBundleStatus
{
    Complete,
    Incomplete,
    Failed,
}

/// <summary>Stable meanings for directed relationships between artifacts.</summary>
public enum DeliveryArtifactRelationshipKind
{
    DerivedFrom,
    Describes,
    Validates,
    Proves,
    UsesPageMap,
    RenderedFrom,
    ReceiptFor,
}

/// <summary>One caller-selected artifact and whether its absence is fatal to completion.</summary>
public sealed record DeliveryArtifactRequest
{
    required public string ArtifactId { get; init; }
    required public DeliveryArtifactKind Kind { get; init; }
    required public DeliveryArtifactRequiredness Requiredness { get; init; }
    public DeliveryReviewProfile? ReviewProfile { get; init; }
    public DeliveryCommentProfile? CommentProfile { get; init; }
}

/// <summary>
/// Immutable named package snapshot used to derive bundle identities. Input and returned bytes are
/// copied so later caller mutation cannot change the manifest identity.
/// </summary>
public sealed class DeliveryDocumentSnapshot
{
    private readonly byte[] _bytes;

    public DeliveryDocumentSnapshot(string name, long documentVersion, ReadOnlySpan<byte> bytes)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("A document snapshot name is required.", nameof(name));
        if (documentVersion < 0)
            throw new ArgumentOutOfRangeException(nameof(documentVersion));
        if (bytes.IsEmpty)
            throw new ArgumentException("A document snapshot cannot be empty.", nameof(bytes));

        Name = name;
        DocumentVersion = documentVersion;
        _bytes = bytes.ToArray();
    }

    public string Name { get; }
    public long DocumentVersion { get; }

    /// <summary>A defensive copy of the exact snapshot bytes.</summary>
    public byte[] Bytes => _bytes.ToArray();

    internal byte[] CopyBytes() => _bytes.ToArray();
}

/// <summary>Complete caller intent required before bundle assembly begins.</summary>
public sealed class DeliveryBundleRequest
{
    private readonly DeliveryArtifactRequest[] _artifacts;

    public DeliveryBundleRequest(
        DeliveryDocumentSnapshot baseline,
        DeliveryDocumentSnapshot working,
        DeliveryDocumentSnapshot final,
        DeliveryBundleRevisionPolicy revisionPolicy,
        IEnumerable<DeliveryArtifactRequest> artifacts)
    {
        Baseline = baseline ?? throw new ArgumentNullException(nameof(baseline));
        Working = working ?? throw new ArgumentNullException(nameof(working));
        Final = final ?? throw new ArgumentNullException(nameof(final));
        RevisionPolicy = revisionPolicy ?? throw new ArgumentNullException(nameof(revisionPolicy));
        _artifacts = artifacts?.ToArray() ?? throw new ArgumentNullException(nameof(artifacts));
    }

    public DeliveryDocumentSnapshot Baseline { get; }
    public DeliveryDocumentSnapshot Working { get; }
    public DeliveryDocumentSnapshot Final { get; }
    public DeliveryBundleRevisionPolicy RevisionPolicy { get; }

    /// <summary>A defensive copy of the explicit artifact request list.</summary>
    public IReadOnlyList<DeliveryArtifactRequest> Artifacts => _artifacts.ToArray();

    internal IReadOnlyList<DeliveryArtifactRequest> ArtifactSnapshot => _artifacts;
}

/// <summary>Renderer-owned metadata retained with one output artifact.</summary>
public sealed record DeliveryArtifactRenderMetadataInput
{
    required public DeliveryReviewProfile ReviewProfile { get; init; }
    required public DeliveryCommentProfile CommentProfile { get; init; }
    required public string SourceDocumentName { get; init; }
    required public long SourceDocumentVersion { get; init; }
    required public VerificationDigest SourcePackageDigest { get; init; }
    public string? RendererFingerprint { get; init; }
    public long? PageCount { get; init; }
    public IReadOnlyList<DeliverableRenderDiagnostic> Warnings { get; init; } =
        Array.Empty<DeliverableRenderDiagnostic>();
}

/// <summary>Artifact bytes or an explicit unavailability result supplied to manifest construction.</summary>
public sealed class DeliveryBundleArtifactInput
{
    private readonly byte[]? _bytes;

    private DeliveryBundleArtifactInput(
        string artifactId,
        DeliveryArtifactKind kind,
        string relativePath,
        string mediaType,
        DeliveryArtifactAvailability availability,
        ReadOnlySpan<byte> bytes,
        string? unavailableReason,
        bool isImplicit,
        DeliveryArtifactRequiredness implicitRequiredness,
        DeliveryArtifactRenderMetadataInput? renderMetadata)
    {
        ArtifactId = artifactId;
        Kind = kind;
        RelativePath = relativePath;
        MediaType = mediaType;
        Availability = availability;
        _bytes = availability == DeliveryArtifactAvailability.Available ? bytes.ToArray() : null;
        UnavailableReason = unavailableReason;
        IsImplicit = isImplicit;
        ImplicitRequiredness = implicitRequiredness;
        RenderMetadata = renderMetadata;
    }

    public string ArtifactId { get; }
    public DeliveryArtifactKind Kind { get; }
    public string RelativePath { get; }
    public string MediaType { get; }
    public DeliveryArtifactAvailability Availability { get; }
    public byte[]? Bytes => _bytes?.ToArray();
    public string? UnavailableReason { get; }
    public bool IsImplicit { get; }
    public DeliveryArtifactRequiredness ImplicitRequiredness { get; }
    public DeliveryArtifactRenderMetadataInput? RenderMetadata { get; }

    public static DeliveryBundleArtifactInput Available(
        string artifactId,
        DeliveryArtifactKind kind,
        string relativePath,
        string mediaType,
        ReadOnlySpan<byte> bytes,
        bool isImplicit = false,
        DeliveryArtifactRequiredness implicitRequiredness = DeliveryArtifactRequiredness.Optional,
        DeliveryArtifactRenderMetadataInput? renderMetadata = null) =>
        new(artifactId, kind, relativePath, mediaType, DeliveryArtifactAvailability.Available,
            bytes, null, isImplicit, implicitRequiredness, renderMetadata);

    public static DeliveryBundleArtifactInput Unavailable(
        string artifactId,
        DeliveryArtifactKind kind,
        string relativePath,
        string mediaType,
        string reason,
        bool isImplicit = false,
        DeliveryArtifactRequiredness implicitRequiredness = DeliveryArtifactRequiredness.Optional,
        DeliveryArtifactRenderMetadataInput? renderMetadata = null) =>
        new(artifactId, kind, relativePath, mediaType, DeliveryArtifactAvailability.Unavailable,
            ReadOnlySpan<byte>.Empty, reason, isImplicit, implicitRequiredness, renderMetadata);

    internal byte[]? CopyBytes() => _bytes?.ToArray();
    internal byte[]? OwnedBytes => _bytes;
}

/// <summary>One directed, typed edge between two artifact IDs.</summary>
public sealed record DeliveryArtifactRelationship
{
    required public string RelationshipId { get; init; }
    required public DeliveryArtifactRelationshipKind Kind { get; init; }
    required public string FromArtifactId { get; init; }
    required public string ToArtifactId { get; init; }
}

/// <summary>Resource ceilings applied before expensive hashing or collection traversal.</summary>
public sealed record DeliveryBundleVerificationLimits
{
    public int MaxManifestBytes { get; init; } = 4 * 1024 * 1024;
    public int MaxArtifacts { get; init; } = 1_024;
    public int MaxRelationships { get; init; } = 4_096;
    public int MaxWarningsPerArtifact { get; init; } = 1_024;
    public int MaxStringLength { get; init; } = 4_096;
    public long MaxArtifactBytes { get; init; } = 256L * 1024 * 1024;
    public long MaxTotalArtifactBytes { get; init; } = 512L * 1024 * 1024;

    internal void Validate()
    {
        if (MaxManifestBytes <= 0 || MaxArtifacts <= 0 || MaxRelationships <= 0
            || MaxWarningsPerArtifact <= 0 || MaxStringLength <= 0
            || MaxArtifactBytes <= 0 || MaxTotalArtifactBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(DeliveryBundleVerificationLimits));
    }
}

public enum DeliveryBundleArtifactVerificationStatus
{
    Verified,
    DeclaredUnavailable,
    MissingBytes,
    UnexpectedBytes,
    SizeMismatch,
    DigestMismatch,
    ResourceLimit,
}

/// <summary>Independent verification outcome for one declared artifact.</summary>
public sealed record DeliveryBundleArtifactVerification
{
    required public string ArtifactId { get; init; }
    required public DeliveryBundleArtifactVerificationStatus Status { get; init; }
}

/// <summary>Independent validation result for a manifest and its separately supplied artifact bytes.</summary>
public sealed record DeliveryBundleVerificationResult
{
    required public bool IsValid { get; init; }
    public IReadOnlyList<string> Findings { get; init; } = Array.Empty<string>();
    public IReadOnlyList<DeliveryBundleArtifactVerification> Artifacts { get; init; } =
        Array.Empty<DeliveryBundleArtifactVerification>();
}

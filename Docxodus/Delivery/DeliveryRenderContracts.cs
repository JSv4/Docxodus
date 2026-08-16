// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>
/// Renderer capability declaration. The delivery core never discovers or launches a renderer;
/// adapters supplied by the standalone-export track advertise exactly what they implement.
/// </summary>
public sealed class DeliveryRendererCapabilities
{
    private readonly DeliveryArtifactKind[] _artifactKinds;
    private readonly DeliveryReviewProfile[] _reviewProfiles;
    private readonly DeliveryCommentProfile[] _commentProfiles;

    public DeliveryRendererCapabilities(
        string rendererId,
        IEnumerable<DeliveryArtifactKind> artifactKinds,
        IEnumerable<DeliveryReviewProfile> reviewProfiles,
        IEnumerable<DeliveryCommentProfile> commentProfiles)
    {
        if (string.IsNullOrWhiteSpace(rendererId))
            throw new ArgumentException("A renderer id is required.", nameof(rendererId));
        RendererId = rendererId;
        _artifactKinds = artifactKinds?.Distinct().Order().ToArray()
            ?? throw new ArgumentNullException(nameof(artifactKinds));
        _reviewProfiles = reviewProfiles?.Distinct().Order().ToArray()
            ?? throw new ArgumentNullException(nameof(reviewProfiles));
        _commentProfiles = commentProfiles?.Distinct().Order().ToArray()
            ?? throw new ArgumentNullException(nameof(commentProfiles));
        if (_artifactKinds.Any(kind => !IsRenderKind(kind)))
            throw new ArgumentException("Renderer capabilities contain a non-render artifact kind.",
                nameof(artifactKinds));
        if (_reviewProfiles.Any(profile => !Enum.IsDefined(profile)))
            throw new ArgumentException("Renderer capabilities contain an invalid review profile.",
                nameof(reviewProfiles));
        if (_commentProfiles.Any(profile => !Enum.IsDefined(profile)))
            throw new ArgumentException("Renderer capabilities contain an invalid comment profile.",
                nameof(commentProfiles));
    }

    public string RendererId { get; }
    public IReadOnlyList<DeliveryArtifactKind> ArtifactKinds => _artifactKinds.ToArray();
    public IReadOnlyList<DeliveryReviewProfile> ReviewProfiles => _reviewProfiles.ToArray();
    public IReadOnlyList<DeliveryCommentProfile> CommentProfiles => _commentProfiles.ToArray();

    public bool Supports(
        DeliveryArtifactKind kind,
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile) =>
        _artifactKinds.Contains(kind)
        && _reviewProfiles.Contains(reviewProfile)
        && _commentProfiles.Contains(commentProfile);

    internal static bool IsRenderKind(DeliveryArtifactKind kind) => kind is
        DeliveryArtifactKind.StandaloneHtml or DeliveryArtifactKind.FinalPdf
        or DeliveryArtifactKind.ReviewPdf;
}

/// <summary>One exact-source request passed to a delivery renderer.</summary>
public sealed class DeliveryRenderRequest
{
    private readonly byte[] _sourceBytes;

    public DeliveryRenderRequest(
        string artifactId,
        DeliveryArtifactKind kind,
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile,
        DeliveryDocumentSnapshot sourceDocument)
    {
        if (string.IsNullOrWhiteSpace(artifactId))
            throw new ArgumentException("An artifact id is required.", nameof(artifactId));
        if (!DeliveryRendererCapabilities.IsRenderKind(kind))
            throw new ArgumentOutOfRangeException(nameof(kind));
        ArtifactId = artifactId;
        Kind = kind;
        if (!Enum.IsDefined(reviewProfile))
            throw new ArgumentOutOfRangeException(nameof(reviewProfile));
        if (!Enum.IsDefined(commentProfile))
            throw new ArgumentOutOfRangeException(nameof(commentProfile));
        ReviewProfile = reviewProfile;
        CommentProfile = commentProfile;
        SourceDocumentName = sourceDocument?.Name
            ?? throw new ArgumentNullException(nameof(sourceDocument));
        SourceDocumentVersion = sourceDocument.DocumentVersion;
        _sourceBytes = sourceDocument.CopyBytes();
        SourcePackageDigest = PackageManifestGenerator.Generate(_sourceBytes)
            .RawPackageBytesDigest;
    }

    public string ArtifactId { get; }
    public DeliveryArtifactKind Kind { get; }
    public DeliveryReviewProfile ReviewProfile { get; }
    public DeliveryCommentProfile CommentProfile { get; }
    public string SourceDocumentName { get; }
    public long SourceDocumentVersion { get; }
    public VerificationDigest SourcePackageDigest { get; }
    public byte[] SourceBytes => _sourceBytes.ToArray();

    internal byte[] CopySourceBytes() => _sourceBytes.ToArray();
}

/// <summary>Renderer-produced bytes and the exact metadata needed for delivery verification.</summary>
public sealed class DeliveryRenderResult
{
    private readonly byte[]? _bytes;
    private readonly byte[]? _pageMapBytes;
    private readonly byte[]? _renderReportBytes;
    private readonly DeliverableRenderDiagnostic[] _diagnostics;

    private DeliveryRenderResult(
        DeliveryArtifactAvailability availability,
        ReadOnlySpan<byte> bytes,
        string mediaType,
        string? unavailableReason,
        string? rendererFingerprint,
        long? pageCount,
        ReadOnlySpan<byte> pageMapBytes,
        ReadOnlySpan<byte> renderReportBytes,
        IEnumerable<DeliverableRenderDiagnostic>? diagnostics)
    {
        Availability = availability;
        _bytes = availability == DeliveryArtifactAvailability.Available ? bytes.ToArray() : null;
        MediaType = mediaType;
        UnavailableReason = unavailableReason;
        RendererFingerprint = rendererFingerprint;
        PageCount = pageCount;
        _pageMapBytes = pageMapBytes.IsEmpty ? null : pageMapBytes.ToArray();
        _renderReportBytes = renderReportBytes.IsEmpty ? null : renderReportBytes.ToArray();
        _diagnostics = diagnostics?.ToArray() ?? Array.Empty<DeliverableRenderDiagnostic>();
    }

    public DeliveryArtifactAvailability Availability { get; }
    public byte[]? Bytes => _bytes?.ToArray();
    public string MediaType { get; }
    public string? UnavailableReason { get; }
    public string? RendererFingerprint { get; }
    public long? PageCount { get; }
    public byte[]? PageMapBytes => _pageMapBytes?.ToArray();
    public byte[]? RenderReportBytes => _renderReportBytes?.ToArray();
    public IReadOnlyList<DeliverableRenderDiagnostic> Diagnostics => _diagnostics.ToArray();

    public static DeliveryRenderResult Available(
        ReadOnlySpan<byte> bytes,
        string mediaType,
        string rendererFingerprint,
        long pageCount,
        ReadOnlySpan<byte> pageMapBytes = default,
        ReadOnlySpan<byte> renderReportBytes = default,
        IEnumerable<DeliverableRenderDiagnostic>? diagnostics = null)
    {
        if (bytes.IsEmpty) throw new ArgumentException("Rendered artifact bytes are required.", nameof(bytes));
        if (string.IsNullOrWhiteSpace(mediaType))
            throw new ArgumentException("A rendered artifact media type is required.", nameof(mediaType));
        if (string.IsNullOrWhiteSpace(rendererFingerprint))
            throw new ArgumentException("A renderer fingerprint is required.", nameof(rendererFingerprint));
        if (pageCount <= 0) throw new ArgumentOutOfRangeException(nameof(pageCount));
        return new DeliveryRenderResult(DeliveryArtifactAvailability.Available, bytes, mediaType,
            null, rendererFingerprint, pageCount, pageMapBytes, renderReportBytes, diagnostics);
    }

    public static DeliveryRenderResult Unavailable(string mediaType, string reason)
    {
        if (string.IsNullOrWhiteSpace(mediaType))
            throw new ArgumentException("A rendered artifact media type is required.", nameof(mediaType));
        if (string.IsNullOrWhiteSpace(reason))
            throw new ArgumentException("An unavailability reason is required.", nameof(reason));
        return new DeliveryRenderResult(DeliveryArtifactAvailability.Unavailable,
            ReadOnlySpan<byte>.Empty, mediaType, reason, null, null,
            ReadOnlySpan<byte>.Empty, ReadOnlySpan<byte>.Empty, null);
    }

    internal byte[]? CopyBytes() => _bytes?.ToArray();
    internal byte[]? CopyPageMapBytes() => _pageMapBytes?.ToArray();
    internal byte[]? CopyRenderReportBytes() => _renderReportBytes?.ToArray();
}

/// <summary>Transport-neutral renderer adapter consumed by the delivery orchestrator.</summary>
public interface IDeliveryArtifactRenderer
{
    DeliveryRendererCapabilities Capabilities { get; }

    ValueTask<DeliveryRenderResult> RenderAsync(
        DeliveryRenderRequest request,
        CancellationToken cancellationToken = default);
}

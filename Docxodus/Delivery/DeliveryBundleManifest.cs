// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>Digest identity of one named document snapshot.</summary>
public sealed record DeliveryBundleDocumentIdentity
{
    required public string Name { get; init; }
    required public long DocumentVersion { get; init; }
    required public long ByteLength { get; init; }
    required public VerificationDigest Digest { get; init; }
}

/// <summary>Renderer metadata bound to a particular artifact.</summary>
public sealed record DeliveryArtifactRenderMetadata
{
    required public DeliveryReviewProfile ReviewProfile { get; init; }
    required public DeliveryCommentProfile CommentProfile { get; init; }
    public string? RendererFingerprint { get; init; }
    public long? PageCount { get; init; }
    public IReadOnlyList<string> Warnings { get; init; } = Array.Empty<string>();
}

/// <summary>One hash-addressed artifact entry in a delivery bundle.</summary>
public sealed record DeliveryBundleArtifact
{
    required public string ArtifactId { get; init; }
    required public DeliveryArtifactKind Kind { get; init; }
    required public DeliveryArtifactProvenance Provenance { get; init; }
    required public DeliveryArtifactRequiredness Requiredness { get; init; }
    required public DeliveryArtifactAvailability Availability { get; init; }
    required public string RelativePath { get; init; }
    required public string MediaType { get; init; }
    public long? ByteLength { get; init; }
    public VerificationDigest? Digest { get; init; }
    public string? UnavailableReason { get; init; }
    public DeliveryArtifactRenderMetadata? Render { get; init; }
}

/// <summary>Canonical digest-covered body of a delivery-bundle manifest.</summary>
public sealed record DeliveryBundleManifestPayload
{
    public const string SchemaId =
        "https://docxodus.dev/schemas/delivery/delivery-bundle-manifest/v1";

    public string Schema { get; init; } = SchemaId;
    public int SchemaVersion { get; init; } = 1;
    required public DeliveryBundleStatus Status { get; init; }
    required public DeliveryBundleRevisionPolicy RevisionPolicy { get; init; }
    required public DeliveryBundleDocumentIdentity BaselineDocument { get; init; }
    required public DeliveryBundleDocumentIdentity WorkingDocument { get; init; }
    required public DeliveryBundleDocumentIdentity FinalDocument { get; init; }
    public IReadOnlyList<DeliveryBundleArtifact> Artifacts { get; init; } =
        Array.Empty<DeliveryBundleArtifact>();
    public IReadOnlyList<DeliveryArtifactRelationship> Relationships { get; init; } =
        Array.Empty<DeliveryArtifactRelationship>();
}

/// <summary>Canonical manifest envelope whose digest covers only the payload.</summary>
public sealed record DeliveryBundleManifest
{
    required public DeliveryBundleManifestPayload Payload { get; init; }
    required public VerificationDigest ManifestDigest { get; init; }

    /// <summary>
    /// Materialize one manifest from explicit request intent and produced or unavailable outputs.
    /// Every requested artifact must have exactly one corresponding output record; unrequested
    /// records must opt in as implicit.
    /// </summary>
    public static DeliveryBundleManifest Create(
        DeliveryBundleRequest request,
        IEnumerable<DeliveryBundleArtifactInput> artifactInputs,
        IEnumerable<DeliveryArtifactRelationship>? relationships = null,
        DeliveryBundleStatus? status = null,
        DeliveryBundleVerificationLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(artifactInputs);
        limits ??= new DeliveryBundleVerificationLimits();
        limits.Validate();

        ValidateRevisionPolicy(request.RevisionPolicy);
        var baseline = BuildIdentity(request.Baseline);
        var working = BuildIdentity(request.Working);
        var final = BuildIdentity(request.Final);
        if (new[] { baseline.Name, working.Name, final.Name }
            .Distinct(StringComparer.Ordinal).Count() != 3)
            throw new ArgumentException(
                "Baseline, working, and final document names must be distinct.", nameof(request));

        var requests = request.ArtifactSnapshot.ToArray();
        if (requests.Length > limits.MaxArtifacts)
            throw new ArgumentException("Artifact request count exceeds the configured limit.", nameof(request));
        ValidateRequests(requests, limits);
        var requestsById = requests.ToDictionary(value => value.ArtifactId, StringComparer.Ordinal);

        var inputArray = artifactInputs.ToArray();
        if (inputArray.Length > limits.MaxArtifacts)
            throw new ArgumentException("Artifact count exceeds the configured limit.", nameof(artifactInputs));
        if (inputArray.Any(value => value is null))
            throw new ArgumentException("Artifact inputs cannot contain null entries.", nameof(artifactInputs));
        if (inputArray.GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Any(group => group.Count() != 1))
            throw new ArgumentException("Artifact IDs must be unique.", nameof(artifactInputs));

        var inputIds = inputArray.Select(value => value.ArtifactId).ToHashSet(StringComparer.Ordinal);
        var missing = requests.Where(value => !inputIds.Contains(value.ArtifactId)).ToArray();
        if (missing.Length != 0)
            throw new ArgumentException(
                $"Every requested artifact needs an explicit output record; missing '{missing[0].ArtifactId}'.",
                nameof(artifactInputs));

        var artifacts = new List<DeliveryBundleArtifact>(inputArray.Length);
        var artifactBytes = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var input in inputArray)
        {
            requestsById.TryGetValue(input.ArtifactId, out var requested);
            if (requested is null && !input.IsImplicit)
                throw new ArgumentException(
                    $"Unrequested artifact '{input.ArtifactId}' must be marked implicit.",
                    nameof(artifactInputs));
            if (requested is not null && input.IsImplicit)
                throw new ArgumentException(
                    $"Requested artifact '{input.ArtifactId}' cannot be marked implicit.",
                    nameof(artifactInputs));
            if (requested is not null && requested.Kind != input.Kind)
                throw new ArgumentException(
                    $"Artifact '{input.ArtifactId}' does not match its requested kind.",
                    nameof(artifactInputs));

            var bytes = input.CopyBytes();
            var artifact = new DeliveryBundleArtifact
            {
                ArtifactId = input.ArtifactId,
                Kind = input.Kind,
                Provenance = requested is null
                    ? DeliveryArtifactProvenance.Implicit
                    : DeliveryArtifactProvenance.Requested,
                Requiredness = requested?.Requiredness ?? input.ImplicitRequiredness,
                Availability = input.Availability,
                RelativePath = DeliveryBundlePath.Canonicalize(input.RelativePath),
                MediaType = input.MediaType,
                ByteLength = bytes?.LongLength,
                Digest = bytes is null ? null : DeliveryBundleCanonicalJson.Digest(bytes),
                UnavailableReason = input.UnavailableReason,
                Render = BuildRenderMetadata(input, requested),
            };
            artifacts.Add(artifact);
            if (bytes is not null)
                artifactBytes.Add(artifact.ArtifactId, bytes);
        }
        artifacts.Sort((left, right) => string.CompareOrdinal(left.ArtifactId, right.ArtifactId));

        var relationshipArray = (relationships ?? Array.Empty<DeliveryArtifactRelationship>())
            .ToArray();
        if (relationshipArray.Length > limits.MaxRelationships)
            throw new ArgumentException("Relationship count exceeds the configured limit.", nameof(relationships));
        if (relationshipArray.Any(value => value is null))
            throw new ArgumentException("Relationships cannot contain null entries.", nameof(relationships));
        relationshipArray = relationshipArray
            .OrderBy(value => value.RelationshipId, StringComparer.Ordinal)
            .ToArray();

        var inferredStatus = artifacts.Any(value =>
            value.Requiredness == DeliveryArtifactRequiredness.Required
            && value.Availability == DeliveryArtifactAvailability.Unavailable)
            ? DeliveryBundleStatus.Incomplete
            : DeliveryBundleStatus.Complete;
        var selectedStatus = status ?? inferredStatus;
        if (selectedStatus == DeliveryBundleStatus.Complete
            && inferredStatus != DeliveryBundleStatus.Complete)
            throw new ArgumentException("A complete bundle cannot contain an unavailable required artifact.",
                nameof(status));

        var payload = new DeliveryBundleManifestPayload
        {
            Status = selectedStatus,
            RevisionPolicy = request.RevisionPolicy with { },
            BaselineDocument = baseline,
            WorkingDocument = working,
            FinalDocument = final,
            Artifacts = artifacts.ToArray(),
            Relationships = relationshipArray,
        };
        var manifest = FromPayload(payload);
        var verification = DeliveryBundleVerifier.Verify(manifest, artifactBytes, limits);
        if (!verification.IsValid)
            throw new ArgumentException(
                $"The delivery bundle manifest is invalid: {verification.Findings[0]}",
                nameof(artifactInputs));
        return manifest;
    }

    public byte[] ToJsonBytes(bool indented = false) =>
        JsonSerializer.SerializeToUtf8Bytes(this,
            indented ? DeliveryBundleCanonicalJson.Indented : DeliveryBundleCanonicalJson.Compact);

    public string ToJson(bool indented = false) => Encoding.UTF8.GetString(ToJsonBytes(indented));

    internal static DeliveryBundleManifest FromPayload(DeliveryBundleManifestPayload payload) => new()
    {
        Payload = payload,
        ManifestDigest = DeliveryBundleCanonicalJson.Digest(
            DeliveryBundleCanonicalJson.SerializePayload(payload)),
    };

    private static void ValidateRevisionPolicy(DeliveryBundleRevisionPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(policy);
        if (!Enum.IsDefined(policy.PreExistingRevisions)
            || !Enum.IsDefined(policy.GeneratedRevisions))
            throw new ArgumentOutOfRangeException(nameof(policy));
    }

    private static void ValidateRequests(
        IReadOnlyList<DeliveryArtifactRequest> requests,
        DeliveryBundleVerificationLimits limits)
    {
        foreach (var request in requests)
        {
            if (request is null)
                throw new ArgumentException("Artifact requests cannot contain null entries.");
            DeliveryBundleValidation.RequireString(request.ArtifactId, "artifact request ID", limits);
            if (!Enum.IsDefined(request.Kind) || !Enum.IsDefined(request.Requiredness))
                throw new ArgumentException($"Artifact request '{request.ArtifactId}' has an invalid enum value.");
            if (IsProfiledRenderKind(request.Kind))
            {
                if (request.ReviewProfile is null || !Enum.IsDefined(request.ReviewProfile.Value)
                    || request.CommentProfile is null || !Enum.IsDefined(request.CommentProfile.Value))
                    throw new ArgumentException(
                        $"Render artifact request '{request.ArtifactId}' requires explicit review and comment profiles.");
                if (request.Kind == DeliveryArtifactKind.FinalPdf
                    && request.ReviewProfile != DeliveryReviewProfile.Final)
                    throw new ArgumentException(
                        $"Final PDF request '{request.ArtifactId}' requires the final review profile.");
                if (request.Kind == DeliveryArtifactKind.ReviewPdf
                    && request.ReviewProfile != DeliveryReviewProfile.Markup)
                    throw new ArgumentException(
                        $"Review PDF request '{request.ArtifactId}' requires the markup review profile.");
            }
            else if (request.ReviewProfile is not null || request.CommentProfile is not null)
            {
                throw new ArgumentException(
                    $"Non-render artifact request '{request.ArtifactId}' cannot select render profiles.");
            }
        }
        if (requests.GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Any(group => group.Count() != 1))
            throw new ArgumentException("Artifact request IDs must be unique.");
    }

    private static DeliveryBundleDocumentIdentity BuildIdentity(DeliveryDocumentSnapshot snapshot)
    {
        var bytes = snapshot.CopyBytes();
        return new DeliveryBundleDocumentIdentity
        {
            Name = snapshot.Name,
            DocumentVersion = snapshot.DocumentVersion,
            ByteLength = bytes.LongLength,
            Digest = DeliveryBundleCanonicalJson.Digest(bytes),
        };
    }

    private static DeliveryArtifactRenderMetadata? BuildRenderMetadata(
        DeliveryBundleArtifactInput artifact,
        DeliveryArtifactRequest? request)
    {
        var input = artifact.RenderMetadata;
        if (!IsProfiledRenderKind(artifact.Kind))
        {
            if (input is not null)
                throw new ArgumentException(
                    $"Non-render artifact '{artifact.ArtifactId}' cannot carry render metadata.");
            return null;
        }

        var reviewProfile = request?.ReviewProfile ?? input?.ReviewProfile;
        var commentProfile = request?.CommentProfile ?? input?.CommentProfile;
        if (reviewProfile is null || commentProfile is null)
            throw new ArgumentException(
                $"Render artifact '{artifact.ArtifactId}' requires explicit review and comment profiles.");
        if (input is not null && (input.ReviewProfile != reviewProfile
                                  || input.CommentProfile != commentProfile))
            throw new ArgumentException(
                $"Render metadata for '{artifact.ArtifactId}' does not match the requested profiles.");
        return new DeliveryArtifactRenderMetadata
        {
            ReviewProfile = reviewProfile.Value,
            CommentProfile = commentProfile.Value,
            RendererFingerprint = input?.RendererFingerprint,
            PageCount = input?.PageCount,
            Warnings = (input?.Warnings ?? Array.Empty<string>())
                .OrderBy(value => value, StringComparer.Ordinal)
                .ToArray(),
        };
    }

    internal static bool IsProfiledRenderKind(DeliveryArtifactKind kind) => kind is
        DeliveryArtifactKind.StandaloneHtml
        or DeliveryArtifactKind.FinalPdf
        or DeliveryArtifactKind.ReviewPdf
        or DeliveryArtifactKind.PageMap
        or DeliveryArtifactKind.RenderReport;
}

internal static class DeliveryBundleCanonicalJson
{
    internal static readonly JsonSerializerOptions Compact = Create(indented: false);
    internal static readonly JsonSerializerOptions Indented = Create(indented: true);

    internal static byte[] SerializePayload(DeliveryBundleManifestPayload payload) =>
        JsonSerializer.SerializeToUtf8Bytes(payload, Compact);

    internal static VerificationDigest Digest(ReadOnlySpan<byte> bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    private static JsonSerializerOptions Create(bool indented)
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = false,
            WriteIndented = indented,
            MaxDepth = 64,
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        };
        options.Converters.Add(new JsonStringEnumConverter(
            JsonNamingPolicy.CamelCase, allowIntegerValues: false));
        return options;
    }
}

internal static class DeliveryBundlePath
{
    internal static string Canonicalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || !string.Equals(value, value.Trim(), StringComparison.Ordinal))
            throw new ArgumentException("Artifact relative paths must be non-blank without surrounding whitespace.");
        if (value.Any(char.IsControl))
            throw new ArgumentException("Artifact relative paths cannot contain control characters.");
        if (value.StartsWith('/') || value.StartsWith('\\')
            || (value.Length >= 2 && char.IsAsciiLetter(value[0]) && value[1] == ':'))
            throw new ArgumentException("Artifact paths must be relative.");

        var canonical = value.Replace('\\', '/');
        var segments = canonical.Split('/');
        if (segments.Any(segment => segment.Length == 0 || segment is "." or ".."))
            throw new ArgumentException("Artifact paths cannot contain empty, dot, or parent segments.");
        foreach (var segment in segments)
            ValidateSegment(segment, value);
        return canonical;
    }

    private static void ValidateSegment(string segment, string path)
    {
        if (segment.EndsWith(' ') || segment.EndsWith('.'))
            throw new ArgumentException(
                $"Artifact path has a non-portable trailing character: '{path}'.");
        if (segment.Any(character => character < ' '
                || character is '\0' or '<' or '>' or ':' or '"' or '|' or '?' or '*'))
            throw new ArgumentException(
                $"Artifact path contains a non-portable filename character: '{path}'.");

        var stem = segment.Split('.', 2)[0];
        if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("PRN", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("AUX", StringComparison.OrdinalIgnoreCase)
            || stem.Equals("NUL", StringComparison.OrdinalIgnoreCase)
            || (stem.Length == 4
                && (stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase)
                    || stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase))
                && stem[3] is >= '1' and <= '9'))
            throw new ArgumentException(
                $"Artifact path contains a reserved filename: '{path}'.");
    }
}

internal static class DeliveryBundleValidation
{
    internal static string RequireString(
        string? value,
        string name,
        DeliveryBundleVerificationLimits limits)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Length > limits.MaxStringLength)
            throw new ArgumentException($"{name} must be non-blank and within the configured length limit.");
        return value;
    }
}

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

namespace Docxodus.Delivery;

/// <summary>
/// Immutable in-memory delivery bundle. Artifact bytes are stored separately from the manifest so
/// the manifest can hash every output without introducing a self-referential digest cycle.
/// </summary>
public sealed class DeliveryBundle
{
    public const string ManifestFileName = "bundle-manifest.json";

    private readonly Dictionary<string, byte[]> _artifactBytes;

    private DeliveryBundle(
        DeliveryBundleManifest manifest,
        Dictionary<string, byte[]> artifactBytes,
        DeliveryBundleVerificationResult verification)
    {
        Manifest = manifest;
        _artifactBytes = artifactBytes;
        Verification = verification;
    }

    public DeliveryBundleManifest Manifest { get; }
    public DeliveryBundleVerificationResult Verification { get; }

    /// <summary>Canonical compact manifest bytes used for transport and publication.</summary>
    public byte[] ManifestBytes => Manifest.ToJsonBytes(indented: false);

    /// <summary>Defensive copies keyed by stable artifact ID.</summary>
    public IReadOnlyDictionary<string, byte[]> ArtifactBytes => _artifactBytes
        .ToDictionary(pair => pair.Key, pair => pair.Value.ToArray(), StringComparer.Ordinal);

    public byte[] GetArtifactBytes(string artifactId)
    {
        ArgumentNullException.ThrowIfNull(artifactId);
        return _artifactBytes.TryGetValue(artifactId, out var bytes)
            ? bytes.ToArray()
            : throw new KeyNotFoundException($"Available artifact not found: {artifactId}");
    }

    public static DeliveryBundle Create(
        DeliveryBundleRequest request,
        IEnumerable<DeliveryBundleArtifactInput> artifactInputs,
        IEnumerable<DeliveryArtifactRelationship>? relationships = null,
        DeliveryBundleStatus? status = null,
        DeliveryBundleVerificationLimits? limits = null)
    {
        ArgumentNullException.ThrowIfNull(artifactInputs);
        var materialized = artifactInputs.ToArray();
        var manifest = DeliveryBundleManifest.Create(
            request, materialized, relationships, status, limits);
        var bytes = materialized
            .Where(input => input.Availability == DeliveryArtifactAvailability.Available)
            .ToDictionary(
                input => input.ArtifactId,
                input => input.CopyBytes()
                    ?? throw new InvalidOperationException(
                        $"Available artifact '{input.ArtifactId}' has no bytes."),
                StringComparer.Ordinal);
        var verification = DeliveryBundleVerifier.Verify(manifest, bytes, limits);
        if (!verification.IsValid)
        {
            throw new InvalidDataException(
                $"Delivery bundle verification failed: {verification.Findings[0]}");
        }
        return new DeliveryBundle(manifest, bytes, verification);
    }
}

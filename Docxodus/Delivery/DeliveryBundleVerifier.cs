// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.Delivery;

/// <summary>Portable, independent validation of a bundle manifest and separately supplied bytes.</summary>
public static class DeliveryBundleVerifier
{
    public static DeliveryBundleVerificationResult Verify(
        DeliveryBundleManifest manifest,
        IReadOnlyDictionary<string, byte[]>? artifactBytes = null,
        DeliveryBundleVerificationLimits? limits = null)
    {
        limits ??= new DeliveryBundleVerificationLimits();
        limits.Validate();
        artifactBytes ??= new Dictionary<string, byte[]>();
        var findings = new List<string>();
        var artifactResults = new List<DeliveryBundleArtifactVerification>();

        if (manifest is null)
            return Result(new[] { "manifest_missing" }, artifactResults);
        if (manifest.Payload is null)
            return Result(new[] { "payload_missing" }, artifactResults);

        var payload = manifest.Payload;
        try
        {
            if (manifest.ToJsonBytes().Length > limits.MaxManifestBytes)
                findings.Add("manifest_resource_limit");
        }
        catch (Exception exception) when (exception is JsonException or NotSupportedException
                                             or InvalidOperationException)
        {
            findings.Add("manifest_not_serializable");
        }

        if (!string.Equals(payload.Schema, DeliveryBundleManifestPayload.SchemaId,
                StringComparison.Ordinal) || payload.SchemaVersion != 1)
            findings.Add("unsupported_manifest_schema");
        if (!Enum.IsDefined(payload.Status))
            findings.Add("invalid_bundle_status");
        ValidateRevisionPolicy(payload.RevisionPolicy, findings);
        ValidateDocument(payload.BaselineDocument, "baseline", limits, findings);
        ValidateDocument(payload.WorkingDocument, "working", limits, findings);
        ValidateDocument(payload.FinalDocument, "final", limits, findings);
        if (new[]
            {
                payload.BaselineDocument?.Name,
                payload.WorkingDocument?.Name,
                payload.FinalDocument?.Name,
            }.Where(value => value is not null).Distinct(StringComparer.Ordinal).Count()
            != new[]
            {
                payload.BaselineDocument?.Name,
                payload.WorkingDocument?.Name,
                payload.FinalDocument?.Name,
            }.Count(value => value is not null))
            findings.Add("duplicate_document_name");

        ValidateManifestDigest(manifest.ManifestDigest, payload, findings);

        var artifacts = payload.Artifacts;
        if (artifacts is null)
        {
            findings.Add("artifacts_missing");
            artifacts = Array.Empty<DeliveryBundleArtifact>();
        }
        if (artifacts.Count > limits.MaxArtifacts)
            findings.Add("artifact_count_resource_limit");
        ValidateCanonicalOrder(artifacts.Select(value => value?.ArtifactId),
            "artifact_order_not_canonical", findings);

        var duplicateIds = artifacts.Where(value => value is not null)
            .GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() != 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.Ordinal);
        foreach (var id in duplicateIds.OrderBy(value => value, StringComparer.Ordinal))
            findings.Add($"duplicate_artifact_id:{id}");

        var canonicalPaths = new HashSet<string>(StringComparer.Ordinal);
        long declaredTotal = 0;
        foreach (var artifact in artifacts)
        {
            if (artifact is null)
            {
                findings.Add("null_artifact");
                continue;
            }
            ValidateArtifactMetadata(artifact, limits, canonicalPaths, findings);
            if (artifact.ByteLength is { } length && length >= 0)
            {
                if (length > limits.MaxArtifactBytes)
                    findings.Add($"artifact_resource_limit:{artifact.ArtifactId}");
                if (declaredTotal > limits.MaxTotalArtifactBytes - Math.Min(
                        length, limits.MaxTotalArtifactBytes))
                    declaredTotal = limits.MaxTotalArtifactBytes + 1;
                else
                    declaredTotal += length;
            }
        }
        if (declaredTotal > limits.MaxTotalArtifactBytes)
            findings.Add("total_artifact_resource_limit");

        if (payload.Status == DeliveryBundleStatus.Complete
            && artifacts.Any(value => value is not null
                && value.Requiredness == DeliveryArtifactRequiredness.Required
                && value.Availability != DeliveryArtifactAvailability.Available))
            findings.Add("complete_bundle_missing_required_artifact");

        ValidateRelationships(payload.Relationships, artifacts, limits, findings);
        ValidateRenderSourceBindings(payload, artifacts, payload.Relationships, findings);
        VerifyArtifactBytes(artifacts, artifactBytes, limits, findings, artifactResults);

        return Result(findings, artifactResults);
    }

    /// <summary>Parse bounded JSON, reject duplicate properties, then perform object verification.</summary>
    public static DeliveryBundleVerificationResult VerifyJson(
        ReadOnlySpan<byte> manifestJson,
        IReadOnlyDictionary<string, byte[]>? artifactBytes = null,
        DeliveryBundleVerificationLimits? limits = null)
    {
        limits ??= new DeliveryBundleVerificationLimits();
        limits.Validate();
        if (manifestJson.Length == 0)
            return Result(new[] { "manifest_json_empty" }, Array.Empty<DeliveryBundleArtifactVerification>());
        if (manifestJson.Length > limits.MaxManifestBytes)
            return Result(new[] { "manifest_resource_limit" }, Array.Empty<DeliveryBundleArtifactVerification>());

        try
        {
            using var document = JsonDocument.Parse(manifestJson.ToArray(), new JsonDocumentOptions
            {
                MaxDepth = 64,
                CommentHandling = JsonCommentHandling.Disallow,
                AllowTrailingCommas = false,
            });
            if (HasDuplicateProperties(document.RootElement))
                return Result(new[] { "duplicate_json_property" },
                    Array.Empty<DeliveryBundleArtifactVerification>());
            var manifest = JsonSerializer.Deserialize<DeliveryBundleManifest>(
                manifestJson, DeliveryBundleCanonicalJson.Compact);
            return manifest is null
                ? Result(new[] { "manifest_json_null" }, Array.Empty<DeliveryBundleArtifactVerification>())
                : Verify(manifest, artifactBytes, limits);
        }
        catch (Exception exception) when (exception is JsonException or NotSupportedException
                                             or InvalidOperationException)
        {
            return Result(new[] { "invalid_manifest_json" },
                Array.Empty<DeliveryBundleArtifactVerification>());
        }
    }

    private static void ValidateRevisionPolicy(
        DeliveryBundleRevisionPolicy? policy,
        ICollection<string> findings)
    {
        if (policy is null)
        {
            findings.Add("revision_policy_missing");
            return;
        }
        if (!Enum.IsDefined(policy.PreExistingRevisions))
            findings.Add("invalid_preexisting_revision_policy");
        if (!Enum.IsDefined(policy.GeneratedRevisions))
            findings.Add("invalid_generated_revision_policy");
    }

    private static void ValidateDocument(
        DeliveryBundleDocumentIdentity? document,
        string role,
        DeliveryBundleVerificationLimits limits,
        ICollection<string> findings)
    {
        if (document is null)
        {
            findings.Add($"{role}_document_missing");
            return;
        }
        if (!ValidString(document.Name, limits))
            findings.Add($"invalid_{role}_document_name");
        if (document.DocumentVersion < 0)
            findings.Add($"invalid_{role}_document_version");
        if (document.ByteLength <= 0)
            findings.Add($"invalid_{role}_document_length");
        if (!ValidDigest(document.Digest))
            findings.Add($"invalid_{role}_document_digest");
    }

    private static void ValidateManifestDigest(
        VerificationDigest? digest,
        DeliveryBundleManifestPayload payload,
        ICollection<string> findings)
    {
        if (!ValidDigest(digest))
        {
            findings.Add("invalid_manifest_digest");
            return;
        }
        try
        {
            var canonicalPayload = DeliveryBundleCanonicalJson.SerializePayload(payload);
            if (!DigestMatches(digest!, canonicalPayload))
                findings.Add("manifest_digest_mismatch");
        }
        catch (Exception exception) when (exception is JsonException or NotSupportedException
                                             or InvalidOperationException)
        {
            findings.Add("payload_not_canonicalizable");
        }
    }

    private static void ValidateArtifactMetadata(
        DeliveryBundleArtifact artifact,
        DeliveryBundleVerificationLimits limits,
        ISet<string> paths,
        ICollection<string> findings)
    {
        var id = artifact.ArtifactId ?? "<null>";
        if (!ValidString(artifact.ArtifactId, limits))
            findings.Add($"invalid_artifact_id:{id}");
        if (!Enum.IsDefined(artifact.Kind))
            findings.Add($"invalid_artifact_kind:{id}");
        if (!Enum.IsDefined(artifact.Provenance))
            findings.Add($"invalid_artifact_provenance:{id}");
        if (!Enum.IsDefined(artifact.Requiredness))
            findings.Add($"invalid_artifact_requiredness:{id}");
        if (!Enum.IsDefined(artifact.Availability))
            findings.Add($"invalid_artifact_availability:{id}");
        if (!ValidMediaType(artifact.MediaType, limits))
            findings.Add($"invalid_artifact_media_type:{id}");

        try
        {
            if (!ValidString(artifact.RelativePath, limits))
                throw new ArgumentException("Artifact path is outside the configured string limit.");
            var canonical = DeliveryBundlePath.Canonicalize(artifact.RelativePath);
            if (!string.Equals(canonical, artifact.RelativePath, StringComparison.Ordinal))
                findings.Add($"artifact_path_not_canonical:{id}");
            if (!paths.Add(canonical))
                findings.Add($"duplicate_artifact_path:{canonical}");
            if (paths.Any(existing => !string.Equals(existing, canonical, StringComparison.Ordinal)
                    && string.Equals(existing, canonical, StringComparison.OrdinalIgnoreCase)))
                findings.Add($"case_colliding_artifact_path:{canonical}");
            if (paths.Any(existing => !string.Equals(existing, canonical, StringComparison.Ordinal)
                    && (existing.StartsWith(canonical + '/', StringComparison.OrdinalIgnoreCase)
                        || canonical.StartsWith(existing + '/', StringComparison.OrdinalIgnoreCase))))
                findings.Add($"artifact_path_hierarchy_collision:{canonical}");
        }
        catch (ArgumentException)
        {
            findings.Add($"invalid_artifact_path:{id}");
        }

        if (artifact.Availability == DeliveryArtifactAvailability.Available)
        {
            if (artifact.ByteLength is null || artifact.ByteLength <= 0)
                findings.Add($"available_artifact_length_missing:{id}");
            if (!ValidDigest(artifact.Digest))
                findings.Add($"available_artifact_digest_missing:{id}");
            if (artifact.UnavailableReason is not null)
                findings.Add($"available_artifact_has_reason:{id}");
        }
        else if (artifact.Availability == DeliveryArtifactAvailability.Unavailable)
        {
            if (artifact.ByteLength is not null || artifact.Digest is not null)
                findings.Add($"unavailable_artifact_has_identity:{id}");
            if (!ValidString(artifact.UnavailableReason, limits))
                findings.Add($"unavailable_artifact_reason_missing:{id}");
            if (artifact.Render?.PageCount is not null)
                findings.Add($"unavailable_artifact_has_page_count:{id}");
        }

        ValidateRenderMetadata(artifact, limits, findings);
    }

    private static void ValidateRenderMetadata(
        DeliveryBundleArtifact artifact,
        DeliveryBundleVerificationLimits limits,
        ICollection<string> findings)
    {
        var id = artifact.ArtifactId ?? "<null>";
        bool isProfiledRender = DeliveryBundleManifest.IsProfiledRenderKind(artifact.Kind);
        // A failed render report is available evidence even when failure occurred before renderer
        // identity completed. Layout artifacts still require the complete identity.
        bool needsCompleteRenderIdentity = artifact.Availability == DeliveryArtifactAvailability.Available
            && isProfiledRender && artifact.Kind != DeliveryArtifactKind.RenderReport;
        bool needsPageCount = artifact.Availability == DeliveryArtifactAvailability.Available
            && (artifact.Kind is DeliveryArtifactKind.StandaloneHtml
                or DeliveryArtifactKind.FinalPdf
                or DeliveryArtifactKind.ReviewPdf
                or DeliveryArtifactKind.PageMap);
        if (isProfiledRender && artifact.Render is null)
        {
            findings.Add($"render_metadata_missing:{id}");
            return;
        }
        if (!isProfiledRender && artifact.Render is not null)
        {
            findings.Add($"unexpected_render_metadata:{id}");
            return;
        }
        if (artifact.Render is not { } render) return;

        if (!Enum.IsDefined(render.ReviewProfile))
            findings.Add($"invalid_review_profile:{id}");
        if (!Enum.IsDefined(render.CommentProfile))
            findings.Add($"invalid_comment_profile:{id}");
        if (!ValidString(render.SourceDocumentName, limits))
            findings.Add($"render_source_name_missing:{id}");
        if (render.SourceDocumentVersion < 0)
            findings.Add($"render_source_version_invalid:{id}");
        if (!ValidDigest(render.SourcePackageDigest))
            findings.Add($"render_source_digest_missing:{id}");
        if (needsCompleteRenderIdentity && !ValidString(render.RendererFingerprint, limits))
            findings.Add($"renderer_fingerprint_missing:{id}");
        if (render.RendererFingerprint is not null
            && !ValidString(render.RendererFingerprint, limits))
            findings.Add($"invalid_renderer_fingerprint:{id}");
        if (render.PageCount is <= 0)
            findings.Add($"invalid_render_page_count:{id}");
        if (needsPageCount && render.PageCount is null)
            findings.Add($"render_page_count_missing:{id}");
        if (artifact.Kind == DeliveryArtifactKind.FinalPdf
            && render.ReviewProfile != DeliveryReviewProfile.Final)
            findings.Add($"final_pdf_profile_mismatch:{id}");
        if (artifact.Kind == DeliveryArtifactKind.ReviewPdf
            && render.ReviewProfile != DeliveryReviewProfile.Markup)
            findings.Add($"review_pdf_profile_mismatch:{id}");

        if (render.Warnings is null)
        {
            findings.Add($"render_warnings_missing:{id}");
            return;
        }
        if (render.Warnings.Count > limits.MaxWarningsPerArtifact)
            findings.Add($"render_warning_resource_limit:{id}");
        var orderedWarnings = render.Warnings
            .OrderBy(value => value, DeliveryBundleManifest.RenderDiagnosticComparer.Instance)
            .ToArray();
        if (!render.Warnings.SequenceEqual(orderedWarnings))
            findings.Add($"render_warning_order_not_canonical:{id}");
        if (render.Warnings.Any(value => !ValidRenderWarning(value, limits)))
            findings.Add($"invalid_render_warning:{id}");
        if (render.Warnings.Distinct().Count() != render.Warnings.Count)
            findings.Add($"duplicate_render_warning:{id}");
    }

    private static bool ValidRenderWarning(
        DeliverableRenderDiagnostic? warning,
        DeliveryBundleVerificationLimits limits)
    {
        if (warning is null || !Enum.IsDefined(warning.Kind)
            || !Enum.IsDefined(warning.Severity) || !ValidString(warning.Message, limits))
            return false;
        return OptionalString(warning.Code, limits)
            && OptionalString(warning.Phase, limits)
            && OptionalString(warning.OwningPartUri, limits)
            && OptionalString(warning.AnchorId, limits)
            && OptionalString(warning.Resource, limits)
            && OptionalString(warning.FontName, limits)
            && OptionalString(warning.SubstitutedFontName, limits)
            && OptionalString(warning.Remediation, limits);
    }

    private static bool OptionalString(string? value, DeliveryBundleVerificationLimits limits) =>
        value is null || ValidString(value, limits);

    private static void ValidateRelationships(
        IReadOnlyList<DeliveryArtifactRelationship>? relationships,
        IReadOnlyList<DeliveryBundleArtifact> artifacts,
        DeliveryBundleVerificationLimits limits,
        ICollection<string> findings)
    {
        if (relationships is null)
        {
            findings.Add("relationships_missing");
            return;
        }
        if (relationships.Count > limits.MaxRelationships)
            findings.Add("relationship_count_resource_limit");
        ValidateCanonicalOrder(relationships.Select(value => value?.RelationshipId),
            "relationship_order_not_canonical", findings);

        var ids = artifacts.Where(value => value is not null)
            .Select(value => value.ArtifactId)
            .Where(value => !string.IsNullOrEmpty(value))
            .ToHashSet(StringComparer.Ordinal);
        foreach (var duplicate in relationships.Where(value => value is not null)
                     .GroupBy(value => value.RelationshipId, StringComparer.Ordinal)
                     .Where(group => group.Count() != 1)
                     .Select(group => group.Key)
                     .OrderBy(value => value, StringComparer.Ordinal))
            findings.Add($"duplicate_relationship_id:{duplicate}");

        foreach (var relationship in relationships)
        {
            if (relationship is null)
            {
                findings.Add("null_relationship");
                continue;
            }
            var id = relationship.RelationshipId ?? "<null>";
            if (!ValidString(relationship.RelationshipId, limits))
                findings.Add($"invalid_relationship_id:{id}");
            if (!Enum.IsDefined(relationship.Kind))
                findings.Add($"invalid_relationship_kind:{id}");
            if (!ValidString(relationship.FromArtifactId, limits)
                || !ids.Contains(relationship.FromArtifactId))
                findings.Add($"relationship_source_missing:{id}");
            if (!ValidString(relationship.ToArtifactId, limits)
                || !ids.Contains(relationship.ToArtifactId))
                findings.Add($"relationship_target_missing:{id}");
            if (string.Equals(relationship.FromArtifactId, relationship.ToArtifactId,
                    StringComparison.Ordinal))
                findings.Add($"self_relationship:{id}");
        }
    }

    private static void ValidateRenderSourceBindings(
        DeliveryBundleManifestPayload payload,
        IReadOnlyList<DeliveryBundleArtifact> artifacts,
        IReadOnlyList<DeliveryArtifactRelationship>? relationships,
        ICollection<string> findings)
    {
        if (relationships is null) return;
        var uniqueArtifacts = artifacts.Where(value => value is not null
                && !string.IsNullOrEmpty(value.ArtifactId))
            .GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        foreach (var artifact in uniqueArtifacts.Values.Where(value =>
                     DeliveryBundleManifest.IsProfiledRenderKind(value.Kind)))
        {
            var id = artifact.ArtifactId ?? "<null>";
            var bindings = relationships.Where(relationship => relationship is not null
                    && relationship.Kind == DeliveryArtifactRelationshipKind.RenderedFrom
                    && string.Equals(relationship.FromArtifactId, artifact.ArtifactId,
                        StringComparison.Ordinal))
                .ToArray();
            if (bindings.Length != 1)
            {
                findings.Add(bindings.Length == 0
                    ? $"render_source_relationship_missing:{id}"
                    : $"render_source_relationship_ambiguous:{id}");
                continue;
            }
            if (string.IsNullOrEmpty(bindings[0].ToArtifactId)
                || !uniqueArtifacts.TryGetValue(bindings[0].ToArtifactId, out var source))
                continue;

            var expectedKind = artifact.Render?.ReviewProfile switch
            {
                DeliveryReviewProfile.Final => DeliveryArtifactKind.FinalDocx,
                DeliveryReviewProfile.Original => DeliveryArtifactKind.PolicyBaselineDocx,
                DeliveryReviewProfile.Markup => DeliveryArtifactKind.ReviewDocx,
                _ => (DeliveryArtifactKind?)null,
            };
            if (expectedKind is null || source.Kind != expectedKind)
            {
                findings.Add($"render_source_kind_mismatch:{id}");
                continue;
            }
            if (source.Availability != DeliveryArtifactAvailability.Available
                || !ValidDigest(source.Digest))
            {
                findings.Add($"render_source_artifact_unavailable:{id}");
                continue;
            }
            if (artifact.Render is not { } render || !ValidDigest(render.SourcePackageDigest))
                continue;
            if (!DigestEquals(render.SourcePackageDigest, source.Digest!))
                findings.Add($"render_source_digest_mismatch:{id}");

            var expectedName = render.ReviewProfile switch
            {
                DeliveryReviewProfile.Final => payload.FinalDocument?.Name,
                DeliveryReviewProfile.Original => "policy-baseline",
                DeliveryReviewProfile.Markup => "review",
                _ => null,
            };
            var expectedVersion = render.ReviewProfile switch
            {
                DeliveryReviewProfile.Final => payload.FinalDocument?.DocumentVersion,
                DeliveryReviewProfile.Original => payload.BaselineDocument?.DocumentVersion,
                DeliveryReviewProfile.Markup => payload.FinalDocument?.DocumentVersion,
                _ => null,
            };
            if (!string.Equals(render.SourceDocumentName, expectedName, StringComparison.Ordinal))
                findings.Add($"render_source_name_mismatch:{id}");
            if (render.SourceDocumentVersion != expectedVersion)
                findings.Add($"render_source_version_mismatch:{id}");
        }
    }

    private static void VerifyArtifactBytes(
        IReadOnlyList<DeliveryBundleArtifact> artifacts,
        IReadOnlyDictionary<string, byte[]> artifactBytes,
        DeliveryBundleVerificationLimits limits,
        ICollection<string> findings,
        ICollection<DeliveryBundleArtifactVerification> results)
    {
        var declared = artifacts.Where(value => value is not null
                && !string.IsNullOrEmpty(value.ArtifactId))
            .GroupBy(value => value.ArtifactId, StringComparer.Ordinal)
            .Where(group => group.Count() == 1)
            .ToDictionary(group => group.Key, group => group.Single(), StringComparer.Ordinal);
        long actualTotal = 0;

        foreach (var artifact in declared.Values.OrderBy(value => value.ArtifactId, StringComparer.Ordinal))
        {
            var hasBytes = artifactBytes.TryGetValue(artifact.ArtifactId, out var bytes) && bytes is not null;
            DeliveryBundleArtifactVerificationStatus status;
            if (artifact.Availability == DeliveryArtifactAvailability.Unavailable)
            {
                status = hasBytes
                    ? DeliveryBundleArtifactVerificationStatus.UnexpectedBytes
                    : DeliveryBundleArtifactVerificationStatus.DeclaredUnavailable;
            }
            else if (!hasBytes)
            {
                status = DeliveryBundleArtifactVerificationStatus.MissingBytes;
            }
            else if (bytes!.LongLength > limits.MaxArtifactBytes)
            {
                status = DeliveryBundleArtifactVerificationStatus.ResourceLimit;
            }
            else
            {
                actualTotal = actualTotal > limits.MaxTotalArtifactBytes - bytes.LongLength
                    ? limits.MaxTotalArtifactBytes + 1
                    : actualTotal + bytes.LongLength;
                status = artifact.ByteLength != bytes.LongLength
                    ? DeliveryBundleArtifactVerificationStatus.SizeMismatch
                    : !ValidDigest(artifact.Digest) || !DigestMatches(artifact.Digest!, bytes)
                        ? DeliveryBundleArtifactVerificationStatus.DigestMismatch
                        : DeliveryBundleArtifactVerificationStatus.Verified;
            }

            results.Add(new DeliveryBundleArtifactVerification
            {
                ArtifactId = artifact.ArtifactId,
                Status = status,
            });
            if (status is not (DeliveryBundleArtifactVerificationStatus.Verified
                or DeliveryBundleArtifactVerificationStatus.DeclaredUnavailable))
                findings.Add($"artifact_{StatusCode(status)}:{artifact.ArtifactId}");
        }

        foreach (var supplied in artifactBytes.Keys.Where(key => string.IsNullOrEmpty(key)
                         || !declared.ContainsKey(key))
                     .OrderBy(value => value, StringComparer.Ordinal))
            findings.Add($"undeclared_artifact_bytes:{supplied}");
        if (actualTotal > limits.MaxTotalArtifactBytes)
            findings.Add("total_actual_artifact_resource_limit");
    }

    private static string StatusCode(DeliveryBundleArtifactVerificationStatus status) => status switch
    {
        DeliveryBundleArtifactVerificationStatus.MissingBytes => "bytes_missing",
        DeliveryBundleArtifactVerificationStatus.UnexpectedBytes => "bytes_unexpected",
        DeliveryBundleArtifactVerificationStatus.SizeMismatch => "size_mismatch",
        DeliveryBundleArtifactVerificationStatus.DigestMismatch => "digest_mismatch",
        DeliveryBundleArtifactVerificationStatus.ResourceLimit => "resource_limit",
        _ => "invalid",
    };

    private static bool ValidMediaType(string? value, DeliveryBundleVerificationLimits limits) =>
        ValidString(value, limits)
        && value!.Contains('/', StringComparison.Ordinal)
        && !value.Any(char.IsWhiteSpace)
        && !value.Any(char.IsControl);

    private static bool ValidString(string? value, DeliveryBundleVerificationLimits limits) =>
        !string.IsNullOrWhiteSpace(value)
        && value.Length <= limits.MaxStringLength
        && !value.Any(char.IsControl);

    private static bool ValidDigest(VerificationDigest? digest) =>
        digest is not null
        && string.Equals(digest.Algorithm, "SHA-256", StringComparison.Ordinal)
        && digest.Value is { Length: 64 }
        && digest.Value.All(character => character is >= '0' and <= '9' or >= 'a' and <= 'f');

    private static bool DigestMatches(VerificationDigest digest, ReadOnlySpan<byte> bytes)
    {
        var expected = Convert.FromHexString(digest.Value);
        var actual = SHA256.HashData(bytes);
        return CryptographicOperations.FixedTimeEquals(expected, actual);
    }

    private static bool DigestEquals(VerificationDigest left, VerificationDigest right) =>
        string.Equals(left.Algorithm, right.Algorithm, StringComparison.Ordinal)
        && string.Equals(left.Value, right.Value, StringComparison.Ordinal);

    private static void ValidateCanonicalOrder(
        IEnumerable<string?> values,
        string finding,
        ICollection<string> findings)
    {
        string? previous = null;
        bool first = true;
        foreach (var value in values)
        {
            if (!first && string.CompareOrdinal(previous, value) > 0)
            {
                findings.Add(finding);
                return;
            }
            first = false;
            previous = value;
        }
    }

    private static bool HasDuplicateProperties(JsonElement element)
    {
        switch (element.ValueKind)
        {
            case JsonValueKind.Object:
            {
                var names = new HashSet<string>(StringComparer.Ordinal);
                foreach (var property in element.EnumerateObject())
                {
                    if (!names.Add(property.Name) || HasDuplicateProperties(property.Value))
                        return true;
                }
                break;
            }
            case JsonValueKind.Array:
                foreach (var item in element.EnumerateArray())
                    if (HasDuplicateProperties(item)) return true;
                break;
        }
        return false;
    }

    private static DeliveryBundleVerificationResult Result(
        IEnumerable<string> findings,
        IEnumerable<DeliveryBundleArtifactVerification> artifacts)
    {
        var normalized = findings.Distinct(StringComparer.Ordinal)
            .OrderBy(value => value, StringComparer.Ordinal)
            .ToArray();
        return new DeliveryBundleVerificationResult
        {
            IsValid = normalized.Length == 0,
            Findings = normalized,
            Artifacts = artifacts.OrderBy(value => value.ArtifactId, StringComparer.Ordinal).ToArray(),
        };
    }
}

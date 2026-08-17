// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

public class DeliveryBundleManifestTests
{
    [Fact]
    public void DBM001_Create_IsDeterministicDefensiveAndIndependentlyVerifiable()
    {
        var baselineBytes = new byte[] { 1, 2, 3 };
        var workingBytes = new byte[] { 4, 5, 6 };
        var finalBytes = new byte[] { 7, 8, 9 };
        var baseline = new DeliveryDocumentSnapshot("baseline", 10, baselineBytes);
        var working = new DeliveryDocumentSnapshot("working", 11, workingBytes);
        var final = new DeliveryDocumentSnapshot("final", 12, finalBytes);
        baselineBytes[0] = 99;
        working.Bytes[0] = 99;
        finalBytes[0] = 99;

        var request = Request(baseline, working, final,
            RequestArtifact("review-pdf", DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactRequiredness.Optional,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Margin));
        var finalArtifactBytes = new byte[] { 20, 21, 22, 23 };
        var validationBytes = new byte[] { 30, 31 };
        var outputs = new[]
        {
            DeliveryBundleArtifactInput.Unavailable(
                "review-pdf", DeliveryArtifactKind.ReviewPdf, "review.pdf",
                "application/pdf", "renderer capability is unavailable",
                renderMetadata: new DeliveryArtifactRenderMetadataInput
                {
                    ReviewProfile = DeliveryReviewProfile.Markup,
                    CommentProfile = DeliveryCommentProfile.Margin,
                    SourceDocumentName = "review",
                    SourceDocumentVersion = final.DocumentVersion,
                    SourcePackageDigest = PackageManifestGenerator.Generate(final.Bytes)
                        .RawPackageBytesDigest,
                }),
            DeliveryBundleArtifactInput.Available(
                "review-docx", DeliveryArtifactKind.ReviewDocx, "documents/review.docx",
                DocxMediaType, final.Bytes,
                isImplicit: true,
                implicitRequiredness: DeliveryArtifactRequiredness.Required),
            DeliveryBundleArtifactInput.Available(
                "final-docx", DeliveryArtifactKind.FinalDocx, "documents\\final.docx",
                DocxMediaType, finalArtifactBytes,
                isImplicit: true,
                implicitRequiredness: DeliveryArtifactRequiredness.Required),
            DeliveryBundleArtifactInput.Available(
                "validation", DeliveryArtifactKind.ValidationReport, "reports/validation.json",
                "application/json", validationBytes, isImplicit: true),
        };
        var relationships = new[]
        {
            new DeliveryArtifactRelationship
            {
                RelationshipId = "review-pdf-rendered-from-review",
                Kind = DeliveryArtifactRelationshipKind.RenderedFrom,
                FromArtifactId = "review-pdf",
                ToArtifactId = "review-docx",
            },
            new DeliveryArtifactRelationship
            {
                RelationshipId = "validation-validates-final",
                Kind = DeliveryArtifactRelationshipKind.Validates,
                FromArtifactId = "validation",
                ToArtifactId = "final-docx",
            },
        };

        var first = DeliveryBundleManifest.Create(request, outputs, relationships);
        var second = DeliveryBundleManifest.Create(request, outputs.Reverse(), relationships);

        Assert.Equal(DeliveryBundleStatus.Complete, first.Payload.Status);
        Assert.Equal(DeliveryRevisionPolicy.Preserve,
            first.Payload.RevisionPolicy.PreExistingRevisions);
        Assert.Equal(DeliveryRevisionPolicy.Accept,
            first.Payload.RevisionPolicy.GeneratedRevisions);
        Assert.Equal(10, first.Payload.BaselineDocument.DocumentVersion);
        Assert.Equal(11, first.Payload.WorkingDocument.DocumentVersion);
        Assert.Equal(12, first.Payload.FinalDocument.DocumentVersion);
        Assert.NotEqual(first.Payload.BaselineDocument.Digest,
            first.Payload.WorkingDocument.Digest);
        Assert.NotEqual(first.Payload.WorkingDocument.Digest,
            first.Payload.FinalDocument.Digest);
        Assert.Equal("documents/final.docx",
            first.Payload.Artifacts.Single(value => value.ArtifactId == "final-docx").RelativePath);
        var review = first.Payload.Artifacts.Single(value => value.ArtifactId == "review-pdf");
        Assert.Equal(DeliveryArtifactProvenance.Requested, review.Provenance);
        Assert.Equal(DeliveryArtifactRequiredness.Optional, review.Requiredness);
        Assert.Equal(DeliveryReviewProfile.Markup, review.Render?.ReviewProfile);
        Assert.Equal(DeliveryCommentProfile.Margin, review.Render?.CommentProfile);
        Assert.Equal("review", review.Render?.SourceDocumentName);
        Assert.Equal(final.DocumentVersion, review.Render?.SourceDocumentVersion);
        Assert.Equal(PackageManifestGenerator.Generate(final.Bytes).RawPackageBytesDigest,
            review.Render?.SourcePackageDigest);
        Assert.Equal(first.ManifestDigest, second.ManifestDigest);
        Assert.Equal(first.ToJson(), second.ToJson());

        var bytes = new Dictionary<string, byte[]>
        {
            ["final-docx"] = finalArtifactBytes,
            ["review-docx"] = final.Bytes,
            ["validation"] = validationBytes,
        };
        Assert.True(DeliveryBundleVerifier.Verify(first, bytes).IsValid);
        Assert.True(DeliveryBundleVerifier.VerifyJson(first.ToJsonBytes(indented: true), bytes).IsValid);

        var forgedRender = review with
        {
            Render = review.Render! with
            {
                SourceDocumentVersion = review.Render.SourceDocumentVersion + 1,
            },
        };
        var forged = DeliveryBundleManifest.FromPayload(first.Payload with
        {
            Artifacts = first.Payload.Artifacts.Select(artifact =>
                artifact.ArtifactId == review.ArtifactId ? forgedRender : artifact).ToArray(),
        });
        Assert.Contains("render_source_version_mismatch:review-pdf",
            DeliveryBundleVerifier.Verify(forged, bytes).Findings);
    }

    [Fact]
    public void DBM002_RenderRequestsRequireExplicitOrthogonalProfilesAndMatchingMetadata()
    {
        var snapshots = Snapshots();
        var missingProfile = Request(snapshots.Baseline, snapshots.Working, snapshots.Final,
            RequestArtifact("html", DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required));
        Assert.Throws<ArgumentException>(() => DeliveryBundleManifest.Create(
            missingProfile,
            new[]
            {
                DeliveryBundleArtifactInput.Unavailable(
                    "html", DeliveryArtifactKind.StandaloneHtml, "final.html", "text/html", "not produced"),
            }));

        var nonRenderProfile = Request(snapshots.Baseline, snapshots.Working, snapshots.Final,
            RequestArtifact("final", DeliveryArtifactKind.FinalDocx,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final, DeliveryCommentProfile.Hidden));
        Assert.Throws<ArgumentException>(() => DeliveryBundleManifest.Create(
            nonRenderProfile,
            new[]
            {
                DeliveryBundleArtifactInput.Available(
                    "final", DeliveryArtifactKind.FinalDocx, "final.docx", DocxMediaType,
                    new byte[] { 1 }),
            }));

        var renderRequest = Request(snapshots.Baseline, snapshots.Working, snapshots.Final,
            RequestArtifact("review", DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Inline));
        Assert.Throws<ArgumentException>(() => DeliveryBundleManifest.Create(
            renderRequest,
            new[]
            {
                DeliveryBundleArtifactInput.Available(
                    "review", DeliveryArtifactKind.ReviewPdf, "review.pdf", "application/pdf",
                    MinimalPdf(), renderMetadata: new DeliveryArtifactRenderMetadataInput
                    {
                        ReviewProfile = DeliveryReviewProfile.Final,
                        CommentProfile = DeliveryCommentProfile.Inline,
                        SourceDocumentName = snapshots.Final.Name,
                        SourceDocumentVersion = snapshots.Final.DocumentVersion,
                        SourcePackageDigest = PackageManifestGenerator.Generate(
                            snapshots.Final.Bytes).RawPackageBytesDigest,
                        RendererFingerprint = "renderer-v1",
                        PageCount = 1,
                    }),
            }));
    }

    [Fact]
    public void DBM003_VerifierDetectsByteTamperingMissingBytesAndUnexpectedBytes()
    {
        var finalBytes = new byte[] { 1, 2, 3, 4 };
        var manifest = MinimalManifest(finalBytes, includeUnavailable: true);
        var result = DeliveryBundleVerifier.Verify(manifest, new Dictionary<string, byte[]>
        {
            ["final"] = new byte[] { 1, 2, 3, 5 },
            ["optional-report"] = new byte[] { 9 },
            ["undeclared"] = new byte[] { 8 },
        });

        Assert.False(result.IsValid);
        Assert.Contains("artifact_digest_mismatch:final", result.Findings);
        Assert.Contains("artifact_bytes_unexpected:optional-report", result.Findings);
        Assert.Contains("undeclared_artifact_bytes:undeclared", result.Findings);
    }

    [Fact]
    public void DBM004_VerifierRejectsForgedPayloadShapeEvenWithARecomputedEnvelopeDigest()
    {
        var finalBytes = new byte[] { 1, 2, 3, 4 };
        var manifest = MinimalManifest(finalBytes);
        var artifact = Assert.Single(manifest.Payload.Artifacts);
        var forgedArtifact = artifact with
        {
            RelativePath = "../escape.docx",
            Availability = Docxodus.Delivery.DeliveryArtifactAvailability.Unavailable,
            UnavailableReason = null,
        };
        var forgedPayload = manifest.Payload with
        {
            Status = DeliveryBundleStatus.Complete,
            Artifacts = new[] { forgedArtifact, forgedArtifact },
            Relationships = new[]
            {
                new DeliveryArtifactRelationship
                {
                    RelationshipId = "missing-target",
                    Kind = DeliveryArtifactRelationshipKind.DerivedFrom,
                    FromArtifactId = "final",
                    ToArtifactId = "absent",
                },
            },
        };
        var forged = DeliveryBundleManifest.FromPayload(forgedPayload);

        var result = DeliveryBundleVerifier.Verify(forged,
            new Dictionary<string, byte[]> { ["final"] = finalBytes });

        Assert.False(result.IsValid);
        Assert.Contains("duplicate_artifact_id:final", result.Findings);
        Assert.Contains("invalid_artifact_path:final", result.Findings);
        Assert.Contains("unavailable_artifact_has_identity:final", result.Findings);
        Assert.Contains("unavailable_artifact_reason_missing:final", result.Findings);
        Assert.Contains("complete_bundle_missing_required_artifact", result.Findings);
        Assert.Contains("relationship_target_missing:missing-target", result.Findings);
    }

    [Fact]
    public void DBM005_VerifierChecksEnvelopeDigestAndResourceLimitsBeforeHashing()
    {
        var finalBytes = new byte[] { 1, 2, 3, 4 };
        var manifest = MinimalManifest(finalBytes);
        var tamperedEnvelope = manifest with
        {
            Payload = manifest.Payload with { Status = DeliveryBundleStatus.Failed },
        };
        var digestResult = DeliveryBundleVerifier.Verify(tamperedEnvelope,
            new Dictionary<string, byte[]> { ["final"] = finalBytes });
        Assert.Contains("manifest_digest_mismatch", digestResult.Findings);

        var bounded = DeliveryBundleVerifier.Verify(manifest,
            new Dictionary<string, byte[]> { ["final"] = finalBytes },
            new DeliveryBundleVerificationLimits
            {
                MaxArtifactBytes = 3,
                MaxTotalArtifactBytes = 3,
            });
        Assert.False(bounded.IsValid);
        Assert.Contains("artifact_resource_limit:final", bounded.Findings);
        Assert.Contains(bounded.Artifacts,
            value => value.ArtifactId == "final"
                && value.Status == DeliveryBundleArtifactVerificationStatus.ResourceLimit);
    }

    [Fact]
    public void DBM006_PublishedV1SchemaTracksTheClosedWireVocabulary()
    {
        var schemaPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "../../../../docs/schemas/delivery-bundle-manifest-v1.schema.json"));
        using var schema = JsonDocument.Parse(File.ReadAllBytes(schemaPath));
        var root = schema.RootElement;

        Assert.Equal("https://json-schema.org/draft/2020-12/schema",
            root.GetProperty("$schema").GetString());
        Assert.Equal(DeliveryBundleManifestPayload.SchemaId,
            root.GetProperty("$id").GetString());
        var definitions = root.GetProperty("$defs");
        Assert.Equal(DeliveryBundleManifestPayload.SchemaId,
            definitions.GetProperty("payload").GetProperty("properties")
                .GetProperty("schema").GetProperty("const").GetString());
        Assert.Equal(1, definitions.GetProperty("payload").GetProperty("properties")
            .GetProperty("schemaVersion").GetProperty("const").GetInt32());
        Assert.Equal(CamelNames<DeliveryArtifactKind>(),
            EnumValues(definitions, "artifact", "kind"));
        Assert.Equal(CamelNames<DeliveryArtifactRelationshipKind>(),
            EnumValues(definitions, "relationship", "kind"));
        Assert.Equal(CamelNames<DeliveryRevisionPolicy>(),
            definitions.GetProperty("revisionDisposition").GetProperty("enum")
                .EnumerateArray().Select(value => value.GetString()));
        var renderRequired = definitions.GetProperty("renderMetadata").GetProperty("required")
            .EnumerateArray().Select(value => value.GetString()).ToArray();
        Assert.Contains("sourceDocumentName", renderRequired);
        Assert.Contains("sourceDocumentVersion", renderRequired);
        Assert.Contains("sourcePackageDigest", renderRequired);
        Assert.DoesNotContain("\"maxItems\"", root.GetRawText(), StringComparison.Ordinal);
        Assert.DoesNotContain("\"maxLength\"", root.GetRawText(), StringComparison.Ordinal);
    }

    [Fact]
    public void DBM007_PublishedValidationSchemaTracksAggregateDecisionVocabulary()
    {
        var schemaPath = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "../../../../docs/schemas/delivery-bundle-validation-v1.schema.json"));
        using var schema = JsonDocument.Parse(File.ReadAllBytes(schemaPath));
        var root = schema.RootElement;

        Assert.Equal("https://json-schema.org/draft/2020-12/schema",
            root.GetProperty("$schema").GetString());
        Assert.Equal(DeliveryBundleValidationReport.SchemaId,
            root.GetProperty("$id").GetString());
        Assert.Equal(CamelNames<DeliverableVerificationDecision>(),
            root.GetProperty("properties").GetProperty("decision").GetProperty("enum")
                .EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(CamelNames<DeliveryReviewProfile>(),
            root.GetProperty("$defs").GetProperty("renderCohort")
                .GetProperty("properties").GetProperty("reviewProfile")
                .GetProperty("enum").EnumerateArray().Select(value => value.GetString()));
        Assert.Equal(CamelNames<DeliveryCommentProfile>(),
            root.GetProperty("$defs").GetProperty("renderCohort")
                .GetProperty("properties").GetProperty("commentProfile")
                .GetProperty("enum").EnumerateArray().Select(value => value.GetString()));
        Assert.DoesNotContain("\"maxItems\"", root.GetRawText(), StringComparison.Ordinal);
        Assert.DoesNotContain("\"maxLength\"", root.GetRawText(), StringComparison.Ordinal);
    }

    [Fact]
    public void DBM008_AvailableArtifactsCannotClaimZeroBytes()
    {
        var snapshots = Snapshots();
        var request = Request(snapshots.Baseline, snapshots.Working, snapshots.Final,
            RequestArtifact("empty", DeliveryArtifactKind.ValidationReport,
                DeliveryArtifactRequiredness.Required));

        var exception = Assert.Throws<ArgumentException>(() =>
            DeliveryBundleManifest.Create(request, new[]
            {
                DeliveryBundleArtifactInput.Available(
                    "empty", DeliveryArtifactKind.ValidationReport,
                    "empty.json", "application/json", Array.Empty<byte>()),
            }));

        Assert.Contains("available_artifact_length_missing:empty", exception.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void DBM009_VerifyJsonTreatsNullIdsAndRelationshipEndpointsAsFindings()
    {
        var finalBytes = new byte[] { 1, 2, 3, 4 };
        var manifest = MinimalManifest(finalBytes);
        var artifact = Assert.Single(manifest.Payload.Artifacts);
        var nullArtifactId = DeliveryBundleManifest.FromPayload(manifest.Payload with
        {
            Artifacts = new[] { artifact with { ArtifactId = null! } },
        });

        var artifactResult = DeliveryBundleVerifier.VerifyJson(
            nullArtifactId.ToJsonBytes(),
            new Dictionary<string, byte[]> { ["final"] = finalBytes });

        Assert.False(artifactResult.IsValid);
        Assert.Contains("invalid_artifact_id:<null>", artifactResult.Findings);
        Assert.Contains("undeclared_artifact_bytes:final", artifactResult.Findings);

        var nullRelationshipTarget = DeliveryBundleManifest.FromPayload(manifest.Payload with
        {
            Relationships = new[]
            {
                new DeliveryArtifactRelationship
                {
                    RelationshipId = "null-target",
                    Kind = DeliveryArtifactRelationshipKind.DerivedFrom,
                    FromArtifactId = "final",
                    ToArtifactId = null!,
                },
            },
        });
        var relationshipResult = DeliveryBundleVerifier.VerifyJson(
            nullRelationshipTarget.ToJsonBytes(),
            new Dictionary<string, byte[]> { ["final"] = finalBytes });

        Assert.False(relationshipResult.IsValid);
        Assert.Contains("relationship_target_missing:null-target",
            relationshipResult.Findings);
    }

    [Fact]
    public void DBM010_V1SchemaDoesNotFreezeConfigurableVerifierResourcePolicy()
    {
        var longName = new string('n', 4_097);
        var baseline = new DeliveryDocumentSnapshot(longName + "-baseline", 1, new byte[] { 1 });
        var working = new DeliveryDocumentSnapshot(longName + "-working", 2, new byte[] { 2 });
        var final = new DeliveryDocumentSnapshot(longName + "-final", 3, new byte[] { 3 });
        var request = Request(baseline, working, final,
            RequestArtifact("final", DeliveryArtifactKind.FinalDocx,
                DeliveryArtifactRequiredness.Required));
        var bytes = new byte[] { 4 };
        var limits = new DeliveryBundleVerificationLimits { MaxStringLength = 5_000 };

        var manifest = DeliveryBundleManifest.Create(request, new[]
        {
            DeliveryBundleArtifactInput.Available(
                "final", DeliveryArtifactKind.FinalDocx, "final.docx", DocxMediaType, bytes),
        }, limits: limits);

        Assert.True(DeliveryBundleVerifier.Verify(manifest,
            new Dictionary<string, byte[]> { ["final"] = bytes }, limits).IsValid);
        Assert.False(DeliveryBundleVerifier.Verify(manifest,
            new Dictionary<string, byte[]> { ["final"] = bytes }).IsValid);
    }

    private static DeliveryBundleManifest MinimalManifest(
        byte[] finalBytes,
        bool includeUnavailable = false)
    {
        var snapshots = Snapshots();
        var requests = new List<DeliveryArtifactRequest>
        {
            RequestArtifact("final", DeliveryArtifactKind.FinalDocx,
                DeliveryArtifactRequiredness.Required),
        };
        var outputs = new List<DeliveryBundleArtifactInput>
        {
            DeliveryBundleArtifactInput.Available(
                "final", DeliveryArtifactKind.FinalDocx, "final.docx", DocxMediaType, finalBytes),
        };
        if (includeUnavailable)
        {
            requests.Add(RequestArtifact("optional-report", DeliveryArtifactKind.ValidationReport,
                DeliveryArtifactRequiredness.Optional));
            outputs.Add(DeliveryBundleArtifactInput.Unavailable(
                "optional-report", DeliveryArtifactKind.ValidationReport,
                "validation.json", "application/json", "not requested from validator"));
        }
        return DeliveryBundleManifest.Create(
            Request(snapshots.Baseline, snapshots.Working, snapshots.Final, requests.ToArray()),
            outputs);
    }

    private static DeliveryBundleRequest Request(
        DeliveryDocumentSnapshot baseline,
        DeliveryDocumentSnapshot working,
        DeliveryDocumentSnapshot final,
        params DeliveryArtifactRequest[] requests) => new(
            baseline,
            working,
            final,
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = DeliveryRevisionPolicy.Accept,
            },
            requests);

    private static DeliveryArtifactRequest RequestArtifact(
        string id,
        DeliveryArtifactKind kind,
        DeliveryArtifactRequiredness requiredness,
        DeliveryReviewProfile? reviewProfile = null,
        DeliveryCommentProfile? commentProfile = null) => new()
        {
            ArtifactId = id,
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = reviewProfile,
            CommentProfile = commentProfile,
        };

    private static (DeliveryDocumentSnapshot Baseline, DeliveryDocumentSnapshot Working,
        DeliveryDocumentSnapshot Final) Snapshots() => (
            new DeliveryDocumentSnapshot("baseline", 1, new byte[] { 1 }),
            new DeliveryDocumentSnapshot("working", 2, new byte[] { 2 }),
            new DeliveryDocumentSnapshot("final", 3, new byte[] { 3 }));

    private static byte[] MinimalPdf() => "%PDF-1.7\n%%EOF"u8.ToArray();

    private static IEnumerable<string?> EnumValues(
        JsonElement definitions,
        string definition,
        string property) => definitions.GetProperty(definition).GetProperty("properties")
            .GetProperty(property).GetProperty("enum").EnumerateArray()
            .Select(value => value.GetString());

    private static IEnumerable<string> CamelNames<T>()
        where T : struct, Enum => Enum.GetValues<T>()
            .Select(value => JsonNamingPolicy.CamelCase.ConvertName(value.ToString()));

    private const string DocxMediaType =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
}

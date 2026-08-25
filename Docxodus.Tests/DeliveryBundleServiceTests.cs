// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.Internal;
using Docxodus.Verification;
using Xunit;
using BundleArtifactAvailability = Docxodus.Delivery.DeliveryArtifactAvailability;

namespace Docxodus.Tests;

public sealed class DeliveryBundleServiceTests
{
    [Fact]
    public async Task BuildAsync_ComposesPublishesAndReopensCompleteEvidenceGraph()
    {
        var edit = SingleEdit("Delivery bundle replacement.");
        var request = Request(edit, Artifacts(
            ("baseline-docx", DeliveryArtifactKind.BaselineDocx),
            ("working-docx", DeliveryArtifactKind.WorkingDocx),
            ("final-docx", DeliveryArtifactKind.FinalDocx),
            ("baseline-manifest", DeliveryArtifactKind.BaselinePackageManifest),
            ("final-manifest", DeliveryArtifactKind.FinalPackageManifest),
            ("semantic-source-to-delivered", DeliveryArtifactKind.SemanticDelta),
            ("package-delta", DeliveryArtifactKind.PackageDelta),
            ("validation", DeliveryArtifactKind.ValidationReport),
            ("receipt", DeliveryArtifactKind.ChangeReceipt)),
            ReceiptContext(edit));

        var bundle = await new DeliveryBundleService().BuildAsync(request);

        Assert.True(bundle.Verification.IsValid,
            string.Join(Environment.NewLine, bundle.Verification.Findings));
        Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
        Assert.All(bundle.Manifest.Payload.Artifacts,
            artifact => Assert.Equal(BundleArtifactAvailability.Available,
                artifact.Availability));
        Assert.Contains(bundle.Manifest.Payload.Artifacts, artifact =>
            artifact.Kind == DeliveryArtifactKind.SemanticDelta
            && artifact.Provenance == DeliveryArtifactProvenance.Implicit);
        Assert.Contains(bundle.Manifest.Payload.Relationships, relationship =>
            relationship.Kind == DeliveryArtifactRelationshipKind.ReceiptFor
            && relationship.FromArtifactId == "receipt"
            && relationship.ToArtifactId == "final-docx");

        using var temporary = new TemporaryDirectory();
        var published = Path.Combine(temporary.Path, "bundle");
        DeliveryBundleDirectoryPublisher.Publish(bundle, published);
        AssertPublishedBundleReopens(bundle, published);
        WritePersistentGallery(bundle);

        var returned = bundle.GetArtifactBytes("final-docx");
        var original = (byte[])returned.Clone();
        returned[0] ^= 0xff;
        Assert.Equal(original, bundle.GetArtifactBytes("final-docx"));
        Assert.Equal(edit.BeforeBytes, request.Baseline.Bytes);
        Assert.Equal(edit.AfterBytes, request.Working.Bytes);
    }

    [Fact]
    public async Task BuildAsync_EveryArtifactKind_FormsOneVerifiedPublishedBundle()
    {
        var fixture = TrackedEditAndAcceptance();
        var requests = Enum.GetValues<DeliveryArtifactKind>()
            .Select(kind => kind switch
            {
                DeliveryArtifactKind.StandaloneHtml => RenderArtifact(
                    "standalone-html", kind, DeliveryArtifactRequiredness.Required,
                    DeliveryReviewProfile.Final),
                DeliveryArtifactKind.FinalPdf => RenderArtifact(
                    "final-pdf", kind, DeliveryArtifactRequiredness.Required,
                    DeliveryReviewProfile.Final),
                DeliveryArtifactKind.ReviewPdf => RenderArtifact(
                    "review-pdf", kind, DeliveryArtifactRequiredness.Required,
                    DeliveryReviewProfile.Markup),
                DeliveryArtifactKind.PageMap => RenderArtifact(
                    "page-map", kind, DeliveryArtifactRequiredness.Required,
                    DeliveryReviewProfile.Final),
                DeliveryArtifactKind.RenderReport => RenderArtifact(
                    "render-report", kind, DeliveryArtifactRequiredness.Required,
                    DeliveryReviewProfile.Final),
                _ => new DeliveryArtifactRequest
                {
                    ArtifactId = Kebab(kind),
                    Kind = kind,
                    Requiredness = DeliveryArtifactRequiredness.Required,
                },
            })
            .ToArray();
        var request = new DeliveryBundleBuildRequest(
            new DeliveryDocumentSnapshot(
                "baseline", fixture.FirstResult.BaseVersion, fixture.BaselineBytes),
            new DeliveryDocumentSnapshot(
                "working", fixture.FirstResult.ResultVersion, fixture.WorkingBytes),
            "final",
            fixture.SecondResult.ResultVersion,
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = DeliveryRevisionPolicy.Accept,
            },
            requests,
            new DeliveryReceiptContext(new[]
            {
                new DeliveryReceiptTransactionEvidence(
                    fixture.FirstContribution,
                    new DeliveryDocumentSnapshot(
                        "first-before", fixture.FirstResult.BaseVersion, fixture.BaselineBytes),
                    new DeliveryDocumentSnapshot(
                        "first-after", fixture.FirstResult.ResultVersion, fixture.WorkingBytes)),
                new DeliveryReceiptTransactionEvidence(
                    fixture.SecondContribution,
                    new DeliveryDocumentSnapshot(
                        "second-before", fixture.SecondResult.BaseVersion, fixture.WorkingBytes),
                    new DeliveryDocumentSnapshot(
                        "second-after", fixture.SecondResult.ResultVersion, fixture.FinalBytes)),
            }));

        var bundle = await new DeliveryBundleService(new CapturingRenderer()).BuildAsync(request);

        Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
        Assert.True(bundle.Verification.IsValid,
            string.Join(Environment.NewLine, bundle.Verification.Findings));
        Assert.Equal(
            Enum.GetValues<DeliveryArtifactKind>().Order(),
            bundle.Manifest.Payload.Artifacts.Select(value => value.Kind).Distinct().Order());
        Assert.All(bundle.Manifest.Payload.Artifacts,
            artifact => Assert.Equal(BundleArtifactAvailability.Available,
                artifact.Availability));
        Assert.Contains(bundle.Manifest.Payload.Relationships, relationship =>
            relationship.Kind == DeliveryArtifactRelationshipKind.UsesPageMap);

        using var temporary = new TemporaryDirectory();
        var published = Path.Combine(temporary.Path, "every-artifact-bundle");
        DeliveryBundleDirectoryPublisher.Publish(bundle, published);
        AssertPublishedBundleReopens(bundle, published);
        WritePersistentGallery(bundle, "DOCXODUS_DELIVERY_ALL_ARTIFACT_DIR");
    }

    [Fact]
    public async Task BuildAsync_AbsentRendererReportsOptionalAndFailsRequiredTruthfully()
    {
        var edit = SingleEdit("Renderer availability edit.");
        var optional = Request(edit, new[]
        {
            RenderArtifact(
                "optional-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Optional,
                DeliveryReviewProfile.Final),
        });

        var complete = await new DeliveryBundleService().BuildAsync(optional);

        Assert.Equal(DeliveryBundleStatus.Complete, complete.Manifest.Payload.Status);
        var unavailable = Assert.Single(complete.Manifest.Payload.Artifacts, artifact =>
            artifact.ArtifactId == "optional-html");
        Assert.Equal(BundleArtifactAvailability.Unavailable, unavailable.Availability);
        Assert.Contains("No delivery renderer", unavailable.UnavailableReason,
            StringComparison.Ordinal);
        Assert.Contains(complete.Manifest.Payload.Artifacts, artifact =>
            artifact.Kind == DeliveryArtifactKind.FinalDocx
            && artifact.Provenance == DeliveryArtifactProvenance.Implicit
            && artifact.Availability == BundleArtifactAvailability.Available);

        var required = Request(edit, new[]
        {
            RenderArtifact(
                "required-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final),
        });
        var error = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await new DeliveryBundleService().BuildAsync(required));
        Assert.Equal("required_artifact_unavailable", error.Code);

        var diagnostic = await new DeliveryBundleService().BuildAsync(
            required,
            new DeliveryBundleBuildOptions { ReturnIncompleteBundle = true });
        Assert.Equal(DeliveryBundleStatus.Incomplete, diagnostic.Manifest.Payload.Status);
        Assert.True(diagnostic.Verification.IsValid);
    }

    [Fact]
    public async Task BuildAsync_FailedRendererReportRemainsAvailableEvidence()
    {
        var edit = SingleEdit("Failed render evidence edit.");
        var reportBytes = Encoding.UTF8.GetBytes(
            "{\"schema\":\"https://docxodus.dev/schemas/render/render-report/v1\","
            + "\"schemaVersion\":1,\"status\":\"failed\"}");
        var warning = new DeliverableRenderDiagnostic
        {
            Kind = DeliverableRenderDiagnosticKind.UnsupportedContent,
            Code = "unsupported_revision_story",
            Severity = VerificationFindingSeverity.Error,
            Phase = "package_preflight",
            Message = "A revision story could not be projected.",
            OwningPartUri = "/word/comments.xml",
            Resource = "comment:2",
            Remediation = "Resolve the unsupported revision story.",
        };
        var bundle = await new DeliveryBundleService(
            new FailedEvidenceRenderer(reportBytes, warning)).BuildAsync(
            Request(edit, new[]
            {
                RenderArtifact(
                    "optional-html",
                    DeliveryArtifactKind.StandaloneHtml,
                    DeliveryArtifactRequiredness.Optional,
                    DeliveryReviewProfile.Final),
            }),
            new DeliveryBundleBuildOptions { FailOnDeliverableValidationFailure = false });

        var html = Assert.Single(bundle.Manifest.Payload.Artifacts,
            artifact => artifact.ArtifactId == "optional-html");
        Assert.Equal(BundleArtifactAvailability.Unavailable, html.Availability);
        Assert.Equal("unsupported_revision_story", Assert.Single(html.Render!.Warnings).Code);
        var report = Assert.Single(bundle.Manifest.Payload.Artifacts,
            artifact => artifact.Kind == DeliveryArtifactKind.RenderReport
                && artifact.Render?.ReviewProfile == DeliveryReviewProfile.Final);
        Assert.Equal(BundleArtifactAvailability.Available, report.Availability);
        Assert.Null(report.Render!.RendererFingerprint);
        Assert.Equal(reportBytes, bundle.GetArtifactBytes(report.ArtifactId));
        Assert.True(bundle.Verification.IsValid,
            string.Join(Environment.NewLine, bundle.Verification.Findings));
    }

    [Fact]
    public async Task BuildAsync_RequiredRenderPromotesExplicitOptionalSidecars()
    {
        var edit = SingleEdit("Required review sidecar edit.", tracked: true);
        var request = Request(edit, new[]
        {
            RenderArtifact(
                "required-review-pdf",
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Markup),
            RenderArtifact(
                "optional-review-map",
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactRequiredness.Optional,
                DeliveryReviewProfile.Markup),
            RenderArtifact(
                "optional-review-report",
                DeliveryArtifactKind.RenderReport,
                DeliveryArtifactRequiredness.Optional,
                DeliveryReviewProfile.Markup),
        });
        var service = new DeliveryBundleService(new SidecarFailureRenderer());

        var error = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await service.BuildAsync(request,
                new DeliveryBundleBuildOptions
                {
                    FailOnDeliverableValidationFailure = false,
                }));
        Assert.Equal("required_artifact_unavailable", error.Code);

        var diagnostic = await service.BuildAsync(
            request,
            new DeliveryBundleBuildOptions
            {
                FailOnDeliverableValidationFailure = false,
                ReturnIncompleteBundle = true,
            });
        Assert.Equal(DeliveryBundleStatus.Incomplete, diagnostic.Manifest.Payload.Status);
        foreach (var id in new[] { "optional-review-map", "optional-review-report" })
        {
            var sidecar = Assert.Single(diagnostic.Manifest.Payload.Artifacts,
                artifact => artifact.ArtifactId == id);
            Assert.Equal(DeliveryArtifactRequiredness.Required, sidecar.Requiredness);
            Assert.Equal(DeliveryArtifactProvenance.Requested, sidecar.Provenance);
            Assert.Equal(BundleArtifactAvailability.Unavailable, sidecar.Availability);
        }
    }

    [Fact]
    public async Task BuildAsync_RenderProfilesUseExactPolicySourcesAndRetainMetadata()
    {
        var edit = SingleEdit("Profile-aware render edit.", tracked: true);
        var renderer = new CapturingRenderer();
        var request = Request(edit, new[]
        {
            RenderArtifact(
                "final-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final),
            RenderArtifact(
                "original-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Original),
            RenderArtifact(
                "review-pdf",
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Markup),
        });

        var bundle = await new DeliveryBundleService(renderer).BuildAsync(request);

        Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
        Assert.Equal(9, renderer.Requests.Count);
        Assert.Equal(1, renderer.BatchCount);
        AssertSource("final-html", DeliveryArtifactKind.FinalDocx);
        AssertSource("original-html", DeliveryArtifactKind.PolicyBaselineDocx);
        AssertSource("review-pdf", DeliveryArtifactKind.ReviewDocx);
        foreach (var artifactId in new[] { "final-html", "original-html", "review-pdf" })
        {
            var artifact = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.ArtifactId == artifactId);
            Assert.Equal("test-renderer|engine-1|fonts-1", artifact.Render?.RendererFingerprint);
            Assert.Equal(1, artifact.Render?.PageCount);
            var warning = Assert.Single(artifact.Render?.Warnings
                ?? Array.Empty<DeliverableRenderDiagnostic>());
            Assert.Equal("fixture_warning", warning.Code);
            Assert.Equal("package_preflight", warning.Phase);
            Assert.Equal("fixture diagnostic", warning.Message);
            Assert.Equal("/word/document.xml", warning.OwningPartUri);
            Assert.Equal("comment:7", warning.Resource);
        }
        Assert.Contains(bundle.Manifest.Payload.Artifacts, artifact =>
            artifact.Kind == DeliveryArtifactKind.ReversibilityProof
            && artifact.Provenance == DeliveryArtifactProvenance.Implicit);
        var validationArtifact = Assert.Single(bundle.Manifest.Payload.Artifacts,
            artifact => artifact.Kind == DeliveryArtifactKind.ValidationReport);
        using var validation = JsonDocument.Parse(
            bundle.GetArtifactBytes(validationArtifact.ArtifactId));
        Assert.Equal(DeliveryBundleValidationReport.SchemaId,
            validation.RootElement.GetProperty("schema").GetString());
        var cohorts = validation.RootElement.GetProperty("renderCohorts")
            .EnumerateArray()
            .ToArray();
        Assert.Equal(3, cohorts.Length);
        AssertValidationSource("final", DeliveryArtifactKind.FinalDocx);
        AssertValidationSource("original", DeliveryArtifactKind.PolicyBaselineDocx);
        AssertValidationSource("markup", DeliveryArtifactKind.ReviewDocx);

        void AssertSource(string renderId, DeliveryArtifactKind sourceKind)
        {
            var renderRequest = Assert.Single(renderer.Requests,
                value => value.ArtifactId == renderId);
            var source = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.Kind == sourceKind);
            Assert.Equal(source.Digest, renderRequest.SourcePackageDigest);
            var renderArtifact = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.ArtifactId == renderId);
            Assert.Equal(renderRequest.SourceDocumentName,
                renderArtifact.Render?.SourceDocumentName);
            Assert.Equal(renderRequest.SourceDocumentVersion,
                renderArtifact.Render?.SourceDocumentVersion);
            Assert.Equal(renderRequest.SourcePackageDigest,
                renderArtifact.Render?.SourcePackageDigest);
            Assert.Contains(bundle.Manifest.Payload.Relationships, relationship =>
                relationship.Kind == DeliveryArtifactRelationshipKind.RenderedFrom
                && relationship.FromArtifactId == renderId
                && relationship.ToArtifactId == source.ArtifactId);
        }

        void AssertValidationSource(string reviewProfile, DeliveryArtifactKind sourceKind)
        {
            var cohort = Assert.Single(cohorts, value =>
                value.GetProperty("reviewProfile").GetString() == reviewProfile);
            Assert.Equal("endnotes", cohort.GetProperty("commentProfile").GetString());
            var source = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.Kind == sourceKind);
            var sourceDigest = cohort.GetProperty("sourceDocument")
                .GetProperty("digest")
                .GetProperty("value")
                .GetString();
            Assert.Equal(source.Digest?.Value, sourceDigest);
            var verification = cohort.GetProperty("verification");
            Assert.True(verification.GetProperty("baselineCompared").GetBoolean());
            Assert.Equal(sourceDigest,
                verification.GetProperty("deliverablePackage")
                    .GetProperty("rawPackageBytesDigest")
                    .GetProperty("value")
                    .GetString());
            Assert.Equal(3, verification.GetProperty("companionArtifacts")
                .GetArrayLength());
        }
    }

    [Fact]
    public async Task BuildAsync_MalformedMarkupCohortFailsBundleValidation()
    {
        var edit = SingleEdit("Malformed markup render edit.", tracked: true);
        var request = Request(edit, new[]
        {
            RenderArtifact(
                "review-pdf",
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Markup),
        });

        var error = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await new DeliveryBundleService(new MalformedMarkupRenderer())
                .BuildAsync(request));

        Assert.Equal("deliverable_validation_failed", error.Code);
        Assert.Contains("markup/endnotes", error.Message, StringComparison.Ordinal);
        Assert.Contains("artifact.page_map_malformed", error.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task BuildAsync_GroupsRenderJobsAndInvokesTheBatchSeamExactlyOnce()
    {
        var edit = SingleEdit("Batch grouping edit.");
        var renderer = new CapturingRenderer();
        var request = Request(edit, new[]
        {
            RenderArtifact(
                "final-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final),
            RenderArtifact(
                "final-page-map",
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final),
            RenderArtifact(
                "final-html-margin",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final,
                DeliveryCommentProfile.Margin),
        });

        var bundle = await new DeliveryBundleService(renderer).BuildAsync(request);

        Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
        // Exactly one RenderBatchesAsync call carries every group; the two artifacts sharing
        // source, version, and profiles ride one batch while the differing comment profile
        // splits into its own — with the context this renderer itself described.
        var call = Assert.Single(renderer.Calls);
        Assert.Equal(2, call.Count);
        var shared = Assert.Single(call, batch =>
            batch.Context.CommentProfile == DeliveryCommentProfile.Endnotes);
        var sharedIds = shared.Requests.Select(r => r.ArtifactId).ToArray();
        Assert.Contains("final-html", sharedIds);
        Assert.Contains("final-page-map", sharedIds);
        var split = Assert.Single(call, batch =>
            batch.Context.CommentProfile == DeliveryCommentProfile.Margin);
        Assert.Contains("final-html-margin", split.Requests.Select(r => r.ArtifactId));
        Assert.DoesNotContain("final-html", split.Requests.Select(r => r.ArtifactId));
        foreach (var batch in call)
            Assert.Equal(
                renderer.DescribeBatch(batch.Context.ReviewProfile, batch.Context.CommentProfile),
                batch.Context);
        Assert.Equal(
            call.Select(batch => batch.BatchId).Distinct(StringComparer.Ordinal).Count(),
            call.Count);
    }

    [Fact]
    public async Task BuildAsync_FailsClosedWhenDescribeBatchIsImpureOrForeign()
    {
        var edit = SingleEdit("Impure describe edit.");
        var required = Request(edit, new[]
        {
            RenderArtifact(
                "required-html",
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Final),
        });

        var impure = new ImpureDescribeRenderer();
        var impureError = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await new DeliveryBundleService(impure).BuildAsync(required));
        Assert.Equal("required_artifact_unavailable", impureError.Code);
        Assert.Equal(0, impure.BatchCount);

        var foreign = new ForeignProfileDescribeRenderer();
        var foreignError = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await new DeliveryBundleService(foreign).BuildAsync(required));
        Assert.Equal("required_artifact_unavailable", foreignError.Code);
        Assert.Equal(0, foreign.BatchCount);
    }

    [Fact]
    public async Task BuildAsync_RejectsInvalidPdfProfileBeforeInvokingRenderer()
    {
        var edit = SingleEdit("Invalid profile edit.");
        var renderer = new CapturingRenderer();
        var request = Request(edit, new[]
        {
            RenderArtifact(
                "final-pdf",
                DeliveryArtifactKind.FinalPdf,
                DeliveryArtifactRequiredness.Required,
                DeliveryReviewProfile.Markup),
        });

        await Assert.ThrowsAsync<ArgumentException>(async () =>
            await new DeliveryBundleService(renderer).BuildAsync(request));

        Assert.Empty(renderer.Requests);
    }

    [Fact]
    public async Task BuildAsync_ReceiptUnavailabilityDoesNotLeakPartialTransactionEvidence()
    {
        var untracked = SingleEdit("Missing receipt context edit.");
        var missing = Request(untracked, new[]
        {
            new DeliveryArtifactRequest
            {
                ArtifactId = "optional-receipt",
                Kind = DeliveryArtifactKind.ChangeReceipt,
                Requiredness = DeliveryArtifactRequiredness.Optional,
            },
        });

        var missingBundle = await new DeliveryBundleService().BuildAsync(missing);

        Assert.Equal(DeliveryBundleStatus.Complete, missingBundle.Manifest.Payload.Status);
        Assert.Equal(BundleArtifactAvailability.Unavailable,
            Assert.Single(missingBundle.Manifest.Payload.Artifacts, artifact =>
                artifact.ArtifactId == "optional-receipt").Availability);

        var tracked = SingleEdit("Unreachable receipt endpoint.", tracked: true);
        var rejected = Request(tracked, new[]
        {
            new DeliveryArtifactRequest
            {
                ArtifactId = "required-receipt",
                Kind = DeliveryArtifactKind.ChangeReceipt,
                Requiredness = DeliveryArtifactRequiredness.Required,
            },
        }, ReceiptContext(tracked));
        var error = await Assert.ThrowsAsync<DeliveryBundleException>(async () =>
            await new DeliveryBundleService().BuildAsync(rejected));
        Assert.Equal("required_artifact_unavailable", error.Code);

        var diagnostic = await new DeliveryBundleService().BuildAsync(
            rejected,
            new DeliveryBundleBuildOptions { ReturnIncompleteBundle = true });
        Assert.Equal(DeliveryBundleStatus.Incomplete, diagnostic.Manifest.Payload.Status);
        Assert.Equal(BundleArtifactAvailability.Unavailable,
            Assert.Single(diagnostic.Manifest.Payload.Artifacts, artifact =>
                artifact.ArtifactId == "required-receipt").Availability);
        var semantics = diagnostic.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Kind == DeliveryArtifactKind.SemanticDelta)
            .ToArray();
        Assert.Single(semantics);
        Assert.Equal("semantic-source-to-delivered", semantics[0].ArtifactId);
    }

    private static DeliveryBundleBuildRequest Request(
        EditFixture edit,
        IEnumerable<DeliveryArtifactRequest> artifacts,
        DeliveryReceiptContext? receiptContext = null) => new(
            new DeliveryDocumentSnapshot("baseline", edit.Result.BaseVersion, edit.BeforeBytes),
            new DeliveryDocumentSnapshot("working", edit.Result.ResultVersion, edit.AfterBytes),
            "final",
            edit.Result.ResultVersion,
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = DeliveryRevisionPolicy.Preserve,
                GeneratedRevisions = DeliveryRevisionPolicy.Accept,
            },
            artifacts,
            receiptContext);

    private static DeliveryReceiptContext ReceiptContext(EditFixture edit) => new(
        new[]
            {
                new DeliveryReceiptTransactionEvidence(
                    edit.Contribution,
                    new DeliveryDocumentSnapshot(
                        "transaction-before", edit.Result.BaseVersion, edit.BeforeBytes),
                    new DeliveryDocumentSnapshot(
                        "transaction-after", edit.Result.ResultVersion, edit.AfterBytes)),
            });

    private static DeliveryArtifactRequest[] Artifacts(
        params (string Id, DeliveryArtifactKind Kind)[] artifacts) => artifacts
        .Select(value => new DeliveryArtifactRequest
        {
            ArtifactId = value.Id,
            Kind = value.Kind,
            Requiredness = DeliveryArtifactRequiredness.Required,
        })
        .ToArray();

    private static DeliveryArtifactRequest RenderArtifact(
        string id,
        DeliveryArtifactKind kind,
        DeliveryArtifactRequiredness requiredness,
        DeliveryReviewProfile reviewProfile,
        DeliveryCommentProfile commentProfile = DeliveryCommentProfile.Endnotes) => new()
        {
            ArtifactId = id,
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = reviewProfile,
            CommentProfile = commentProfile,
        };

    private static EditFixture SingleEdit(string replacement, bool tracked = false)
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(source, new DocxSessionSettings
        {
            PersistAnchorIds = true,
            TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
            RevisionAuthor = "Delivery Service Test",
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var anchor = session.Project().AnchorIndex.Keys.First(id =>
            id.StartsWith("p:body:", StringComparison.Ordinal));
        var beforeBytes = session.Save();
        var beforeManifest = PackageManifestGenerator.Generate(beforeBytes);
        var operation = DeliveryNormalizedOperation.Create(
            "docx_edit",
            "replace_text",
            JsonSerializer.Serialize(new { anchorId = anchor, markdown = replacement }));
        var result = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docx_edit",
                "replace_text",
                value => value.ReplaceText(anchor, replacement)),
        });
        Assert.True(result.Success,
            result.Failure is null ? "batch failed" : result.Failure.Error.Message);
        var afterBytes = session.Save();
        var afterManifest = PackageManifestGenerator.Generate(afterBytes);
        var contribution = DeliveryTransactionContribution.FromMutationBatchResult(
            result,
            beforeManifest,
            afterManifest,
            new[] { operation });
        return new EditFixture(beforeBytes, afterBytes, result, contribution);
    }

    private static AcceptedEditFixture TrackedEditAndAcceptance()
    {
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(source, new DocxSessionSettings
        {
            PersistAnchorIds = false,
            TrackedChanges = TrackedChangeMode.RenderInline,
            RevisionAuthor = "Delivery Service Test",
            EmitMarkdownPatch = false,
            CaptureInitialProjection = false,
        });
        var anchor = session.Project().AnchorIndex.Keys.First(id =>
            id.StartsWith("p:body:", StringComparison.Ordinal));
        var baseline = session.Save(persistAnchorIds: false);
        var baselineManifest = PackageManifestGenerator.Generate(baseline);
        var editOperation = DeliveryNormalizedOperation.Create(
            "docx_edit",
            "replace_text",
            JsonSerializer.Serialize(new
            {
                anchorId = anchor,
                markdown = "Every-artifact tracked edit.",
            }));
        var first = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docx_edit",
                "replace_text",
                value => value.ReplaceText(anchor, "Every-artifact tracked edit.")),
        });
        Assert.True(first.Success,
            first.Failure is null ? "tracked batch failed" : first.Failure.Error.Message);
        var working = session.Save(persistAnchorIds: false);
        var workingManifest = PackageManifestGenerator.Generate(working);
        var firstContribution = DeliveryTransactionContribution.FromMutationBatchResult(
            first, baselineManifest, workingManifest, new[] { editOperation });

        var acceptOperation = DeliveryNormalizedOperation.Create(
            "docxodus_track_changes", "accept_all");
        var second = session.ExecuteBatch(new[]
        {
            new MutationBatchStep(
                "docxodus_track_changes",
                "accept_all",
                value => value.AcceptAllRevisions()),
        });
        Assert.True(second.Success,
            second.Failure is null ? "accept batch failed" : second.Failure.Error.Message);
        var final = session.Save(persistAnchorIds: false);
        var finalManifest = PackageManifestGenerator.Generate(final);
        var secondContribution = DeliveryTransactionContribution.FromMutationBatchResult(
            second, workingManifest, finalManifest, new[] { acceptOperation });
        return new AcceptedEditFixture(
            baseline,
            working,
            final,
            first,
            second,
            firstContribution,
            secondContribution);
    }

    private static void AssertPublishedBundleReopens(DeliveryBundle bundle, string directory)
    {
        var manifestBytes = File.ReadAllBytes(
            Path.Combine(directory, DeliveryBundle.ManifestFileName));
        Assert.Equal(bundle.ManifestBytes, manifestBytes);
        var artifactBytes = bundle.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Availability == BundleArtifactAvailability.Available)
            .ToDictionary(
                artifact => artifact.ArtifactId,
                artifact => File.ReadAllBytes(Path.Combine(
                    directory,
                    artifact.RelativePath.Replace('/', Path.DirectorySeparatorChar))),
                StringComparer.Ordinal);
        var verification = DeliveryBundleVerifier.VerifyJson(manifestBytes, artifactBytes);
        Assert.True(verification.IsValid,
            string.Join(Environment.NewLine, verification.Findings));
    }

    private static void WritePersistentGallery(
        DeliveryBundle bundle,
        string environmentVariable = "DOCXODUS_DELIVERY_BUNDLE_ARTIFACT_DIR")
    {
        var root = Environment.GetEnvironmentVariable(environmentVariable);
        if (string.IsNullOrWhiteSpace(root))
            return;

        Directory.CreateDirectory(root);
        var published = Path.Combine(root, "bundle");
        DeliveryBundleDirectoryPublisher.Publish(bundle, published);
        var rows = string.Join(Environment.NewLine,
            bundle.Manifest.Payload.Artifacts.Select(artifact =>
            {
                var availability = artifact.Availability.ToString();
                var link = artifact.Availability == BundleArtifactAvailability.Available
                    ? $"<a href=\"bundle/{Html(artifact.RelativePath)}\">view</a>"
                    : Html(artifact.UnavailableReason ?? "unavailable");
                var source = artifact.Render is null
                    ? string.Empty
                    : $"{Html(artifact.Render.SourceDocumentName)}@"
                      + $"{artifact.Render.SourceDocumentVersion}<br><code>"
                      + $"{Html(artifact.Render.SourcePackageDigest.Value)}</code>";
                return $"<tr><td>{Html(artifact.ArtifactId)}</td><td>{artifact.Kind}</td>"
                    + $"<td>{artifact.Provenance}</td><td>{availability}</td>"
                    + $"<td>{source}</td><td>{link}</td></tr>";
            }));
        var index = "<!doctype html><meta charset=\"utf-8\"><title>Docxodus #465 artifacts</title>"
            + "<style>body{font:15px system-ui;margin:2rem;max-width:1100px}"
            + "table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;"
            + "padding:.45rem;text-align:left}th{background:#eee}code{word-break:break-all}</style>"
            + "<h1>Docxodus #465 delivery-bundle evidence</h1>"
            + $"<p>Status: <strong>{bundle.Manifest.Payload.Status}</strong>; independent verification: "
            + $"<strong>{bundle.Verification.IsValid}</strong>.</p>"
            + (bundle.Manifest.Payload.Artifacts.Any(artifact =>
                    artifact.Render?.RendererFingerprint?.StartsWith(
                        "test-renderer|", StringComparison.Ordinal) == true)
                ? "<p><strong>Rendering note:</strong> HTML, PDF, PageMap, and render-report "
                    + "outputs are deterministic test-adapter fixtures proving orchestration and "
                    + "verification. Production rendering remains gated by epic #434.</p>"
                : string.Empty)
            + $"<p>Manifest digest: <code>{Html(bundle.Manifest.ManifestDigest.Value)}</code>. "
            + "<a href=\"bundle/bundle-manifest.json\">View canonical manifest</a></p>"
            + "<table><thead><tr><th>ID</th><th>Kind</th><th>Provenance</th>"
            + "<th>Availability</th><th>Render source</th>"
            + $"<th>Artifact</th></tr></thead><tbody>{rows}</tbody></table>";
        File.WriteAllText(Path.Combine(root, "index.html"), index, Encoding.UTF8);

        var checksums = bundle.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Availability == BundleArtifactAvailability.Available)
            .Select(artifact => $"{artifact.Digest!.Value}  bundle/{artifact.RelativePath}")
            .Append($"{Sha256(bundle.ManifestBytes)}  bundle/{DeliveryBundle.ManifestFileName}");
        File.WriteAllLines(
            Path.Combine(root, "SHA256SUMS"),
            checksums,
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private static string Html(string value) => HtmlEncoder.Default.Encode(value);

    private static string Sha256(byte[] bytes) => Convert.ToHexString(
        SHA256.HashData(bytes)).ToLowerInvariant();

    private static string Kebab<T>(T value)
        where T : struct, Enum
    {
        var name = value.ToString();
        var builder = new StringBuilder(name.Length + 8);
        for (var index = 0; index < name.Length; index++)
        {
            var character = name[index];
            if (index > 0 && char.IsUpper(character))
                builder.Append('-');
            builder.Append(char.ToLowerInvariant(character));
        }
        return builder.ToString();
    }

    /// <summary>
    /// Shared batch-seam scaffold for the fake renderers: a pure per-pair DescribeBatch whose
    /// digests derive only from the renderer id and the pair, and a RenderBatchesAsync that
    /// counts invocations, records the batches, and flattens each batch to one deterministic
    /// per-request result.
    /// </summary>
    private abstract class BatchRendererTestBase : IDeliveryArtifactRenderer
    {
        private readonly List<DeliveryRenderRequest> _requests = new();
        private readonly List<IReadOnlyList<DeliveryRenderBatch>> _calls = new();

        public abstract DeliveryRendererCapabilities Capabilities { get; }

        public IReadOnlyList<DeliveryRenderRequest> Requests => _requests.ToArray();

        public IReadOnlyList<IReadOnlyList<DeliveryRenderBatch>> Calls => _calls.ToArray();

        public int BatchCount => _calls.Count;

        public virtual DeliveryRenderBatchContext DescribeBatch(
            DeliveryReviewProfile reviewProfile,
            DeliveryCommentProfile commentProfile) => new(
            reviewProfile,
            commentProfile,
            TestDigest($"{Capabilities.RendererId}|layout|{reviewProfile}|{commentProfile}"),
            TestDigest($"{Capabilities.RendererId}|runtime"));

        public ValueTask<IReadOnlyDictionary<string, DeliveryRenderResult>> RenderBatchesAsync(
            IReadOnlyList<DeliveryRenderBatch> batches,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _calls.Add(batches.ToArray());
            var results = new Dictionary<string, DeliveryRenderResult>(StringComparer.Ordinal);
            foreach (var batch in batches)
            {
                foreach (var request in batch.Requests)
                {
                    _requests.Add(request);
                    results.Add(request.ArtifactId, Result(request));
                }
            }
            return ValueTask.FromResult<IReadOnlyDictionary<string, DeliveryRenderResult>>(results);
        }

        protected abstract DeliveryRenderResult Result(DeliveryRenderRequest request);

        protected static VerificationDigest TestDigest(string material) => new()
        {
            Algorithm = "SHA-256",
            Value = Convert.ToHexString(
                SHA256.HashData(Encoding.UTF8.GetBytes(material))).ToLowerInvariant(),
        };
    }

    private sealed class CapturingRenderer : BatchRendererTestBase
    {
        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "test-renderer",
            new[]
            {
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactKind.FinalPdf,
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactKind.RenderReport,
            },
            Enum.GetValues<DeliveryReviewProfile>(),
            Enum.GetValues<DeliveryCommentProfile>());

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request)
        {
            var pageMap = PageMapBytes(request.SourceDocumentVersion);
            var bytes = request.Kind switch
            {
                DeliveryArtifactKind.StandaloneHtml => Encoding.UTF8.GetBytes(
                    $"<!doctype html><html><body><p>{request.ArtifactId}</p></body></html>"),
                DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf =>
                    MinimalPdfBytes(),
                DeliveryArtifactKind.PageMap => pageMap,
                DeliveryArtifactKind.RenderReport => Encoding.UTF8.GetBytes(
                    "{\"schema\":\"docxodus.test.render-report/v1\",\"valid\":true}"),
                _ => throw new ArgumentOutOfRangeException(nameof(request)),
            };
            return DeliveryRenderResult.Available(
                bytes,
                request.Kind switch
                {
                    DeliveryArtifactKind.StandaloneHtml => "text/html",
                    DeliveryArtifactKind.FinalPdf or DeliveryArtifactKind.ReviewPdf =>
                        "application/pdf",
                    _ => "application/json",
                },
                "test-renderer|engine-1|fonts-1",
                pageCount: 1,
                pageMapBytes: pageMap,
                renderReportBytes: Encoding.UTF8.GetBytes("{\"valid\":true}"),
                diagnostics: new[]
                {
                    new DeliverableRenderDiagnostic
                    {
                        Kind = DeliverableRenderDiagnosticKind.Warning,
                        Code = "fixture_warning",
                        Phase = "package_preflight",
                        Message = "fixture diagnostic",
                        OwningPartUri = "/word/document.xml",
                        Resource = "comment:7",
                        Remediation = "Inspect the fixture warning.",
                    },
                });
        }

        private static byte[] PageMapBytes(long documentVersion)
        {
            var map = new PageMap
            {
                Mode = PageMapMode.Paginated,
                Availability = PageMapAvailability.Available,
                DocumentVersion = documentVersion,
                RendererFingerprint = "test-renderer|engine-1|fonts-1",
                Pages = new[]
                {
                    new PageMapPage
                    {
                        PageNumber = 1,
                        PageInSection = 1,
                        Width = 612,
                        Height = 792,
                        SectionIndex = 0,
                        PageName = "page-1",
                    },
                },
                Fragments = new[]
                {
                    new PageMapFragment
                    {
                        FragmentId = "fixture:p1:0",
                        AnchorId = "fixture",
                        FragmentIndex = 0,
                        PageNumber = 1,
                        Geometry = new PageMapRect(72, 72, 200, 20),
                        Story = PageMapStory.Body,
                    },
                },
            };
            return Encoding.UTF8.GetBytes(DocxSessionJson.SerializePageMap(map));
        }

        internal static byte[] MinimalPdfBytes() => Encoding.ASCII.GetBytes(
            "%PDF-1.4\n"
            + "1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n"
            + "2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj\n"
            + "3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >> endobj\n"
            + "xref\n0 4\n0000000000 65535 f \n"
            + "trailer << /Size 4 /Root 1 0 R >>\nstartxref\n0\n%%EOF\n");
    }

    /// <summary>DescribeBatch that returns a different layout digest on every call: the service's
    /// purity probe must fail the render closed without ever invoking the batch seam.</summary>
    private sealed class ImpureDescribeRenderer : BatchRendererTestBase
    {
        private int _describeCalls;

        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "impure-describe-renderer",
            new[] { DeliveryArtifactKind.StandaloneHtml },
            Enum.GetValues<DeliveryReviewProfile>(),
            Enum.GetValues<DeliveryCommentProfile>());

        public override DeliveryRenderBatchContext DescribeBatch(
            DeliveryReviewProfile reviewProfile,
            DeliveryCommentProfile commentProfile) => new(
            reviewProfile,
            commentProfile,
            TestDigest($"impure|{_describeCalls++}"),
            TestDigest("impure|runtime"));

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request) =>
            throw new InvalidOperationException("The batch seam must not be reached.");
    }

    /// <summary>DescribeBatch that answers for a different profile pair than it was asked.</summary>
    private sealed class ForeignProfileDescribeRenderer : BatchRendererTestBase
    {
        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "foreign-profile-renderer",
            new[] { DeliveryArtifactKind.StandaloneHtml },
            Enum.GetValues<DeliveryReviewProfile>(),
            Enum.GetValues<DeliveryCommentProfile>());

        public override DeliveryRenderBatchContext DescribeBatch(
            DeliveryReviewProfile reviewProfile,
            DeliveryCommentProfile commentProfile) => new(
            DeliveryReviewProfile.Markup,
            DeliveryCommentProfile.Hidden,
            TestDigest("foreign|layout"),
            TestDigest("foreign|runtime"));

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request) =>
            throw new InvalidOperationException("The batch seam must not be reached.");
    }

    private sealed class FailedEvidenceRenderer : BatchRendererTestBase
    {
        private readonly byte[] _reportBytes;
        private readonly DeliverableRenderDiagnostic _warning;

        internal FailedEvidenceRenderer(
            byte[] reportBytes,
            DeliverableRenderDiagnostic warning)
        {
            _reportBytes = reportBytes;
            _warning = warning;
        }

        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "failed-evidence-renderer",
            new[]
            {
                DeliveryArtifactKind.StandaloneHtml,
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactKind.RenderReport,
            },
            Enum.GetValues<DeliveryReviewProfile>(),
            Enum.GetValues<DeliveryCommentProfile>());

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request) =>
            request.Kind == DeliveryArtifactKind.RenderReport
                ? DeliveryRenderResult.FailedReport(_reportBytes, diagnostics: new[] { _warning })
                : DeliveryRenderResult.Unavailable(
                    request.Kind == DeliveryArtifactKind.StandaloneHtml
                        ? "text/html"
                        : "application/vnd.docxodus.pagemap+json",
                    "Export host resource_policy_failure at package_preflight.",
                    renderReportBytes: _reportBytes,
                    diagnostics: new[] { _warning });
    }

    private sealed class SidecarFailureRenderer : BatchRendererTestBase
    {
        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "sidecar-failure-renderer",
            new[]
            {
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactKind.RenderReport,
            },
            Enum.GetValues<DeliveryReviewProfile>(),
            Enum.GetValues<DeliveryCommentProfile>());

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request)
        {
            if (request.Kind != DeliveryArtifactKind.ReviewPdf)
                return DeliveryRenderResult.Unavailable(
                    request.Kind == DeliveryArtifactKind.PageMap
                        ? "application/vnd.docxodus.pagemap+json"
                        : "application/vnd.docxodus.render-report+json",
                    "The renderer omitted required cohort evidence.");
            var pageMap = new PageMap
            {
                Mode = PageMapMode.Paginated,
                Availability = PageMapAvailability.Available,
                DocumentVersion = request.SourceDocumentVersion,
                RendererFingerprint = "sidecar-failure-renderer-v1",
                Pages = new[]
                {
                    new PageMapPage
                    {
                        PageNumber = 1,
                        PageInSection = 1,
                        Width = 612,
                        Height = 792,
                        SectionIndex = 0,
                        PageName = "page-1",
                    },
                },
            };
            var pageMapBytes = Encoding.UTF8.GetBytes(DocxSessionJson.SerializePageMap(pageMap));
            return DeliveryRenderResult.Available(
                CapturingRenderer.MinimalPdfBytes(),
                "application/pdf",
                "sidecar-failure-renderer-v1",
                1,
                pageMapBytes,
                Encoding.UTF8.GetBytes("{\"status\":\"complete\"}"));
        }
    }

    private sealed class MalformedMarkupRenderer : BatchRendererTestBase
    {
        private static readonly byte[] MalformedPageMap = Encoding.UTF8.GetBytes("{");

        public override DeliveryRendererCapabilities Capabilities { get; } = new(
            "malformed-markup-renderer",
            new[]
            {
                DeliveryArtifactKind.ReviewPdf,
                DeliveryArtifactKind.PageMap,
                DeliveryArtifactKind.RenderReport,
            },
            new[] { DeliveryReviewProfile.Markup },
            Enum.GetValues<DeliveryCommentProfile>());

        protected override DeliveryRenderResult Result(DeliveryRenderRequest request)
        {
            var bytes = request.Kind switch
            {
                DeliveryArtifactKind.ReviewPdf => CapturingRenderer.MinimalPdfBytes(),
                DeliveryArtifactKind.PageMap => MalformedPageMap,
                DeliveryArtifactKind.RenderReport =>
                    Encoding.UTF8.GetBytes("{\"status\":\"complete\"}"),
                _ => throw new ArgumentOutOfRangeException(nameof(request)),
            };
            return DeliveryRenderResult.Available(
                bytes,
                request.Kind == DeliveryArtifactKind.ReviewPdf
                    ? "application/pdf"
                    : "application/json",
                "malformed-markup-renderer-v1",
                1,
                MalformedPageMap,
                Encoding.UTF8.GetBytes("{\"status\":\"complete\"}"));
        }
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"docxodus-delivery-service-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
    }

    private sealed record EditFixture(
        byte[] BeforeBytes,
        byte[] AfterBytes,
        MutationBatchResult Result,
        DeliveryTransactionContribution Contribution);

    private sealed record AcceptedEditFixture(
        byte[] BaselineBytes,
        byte[] WorkingBytes,
        byte[] FinalBytes,
        MutationBatchResult FirstResult,
        MutationBatchResult SecondResult,
        DeliveryTransactionContribution FirstContribution,
        DeliveryTransactionContribution SecondContribution);
}

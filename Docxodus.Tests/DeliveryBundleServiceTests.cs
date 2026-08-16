// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Docxodus.Delivery;
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
        Assert.Equal(3, renderer.Requests.Count);
        AssertSource("final-html", DeliveryArtifactKind.FinalDocx);
        AssertSource("original-html", DeliveryArtifactKind.PolicyBaselineDocx);
        AssertSource("review-pdf", DeliveryArtifactKind.ReviewDocx);
        foreach (var artifactId in new[] { "final-html", "original-html", "review-pdf" })
        {
            var artifact = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.ArtifactId == artifactId);
            Assert.Equal("test-renderer|engine-1|fonts-1", artifact.Render?.RendererFingerprint);
            Assert.Equal(2, artifact.Render?.PageCount);
            Assert.Equal(new[] { "fixture diagnostic" }, artifact.Render?.Warnings);
        }
        Assert.Contains(bundle.Manifest.Payload.Artifacts, artifact =>
            artifact.Kind == DeliveryArtifactKind.ReversibilityProof
            && artifact.Provenance == DeliveryArtifactProvenance.Implicit);

        void AssertSource(string renderId, DeliveryArtifactKind sourceKind)
        {
            var renderRequest = Assert.Single(renderer.Requests,
                value => value.ArtifactId == renderId);
            var source = Assert.Single(bundle.Manifest.Payload.Artifacts,
                value => value.Kind == sourceKind);
            Assert.Equal(source.Digest, renderRequest.SourcePackageDigest);
        }
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
        DeliveryReviewProfile reviewProfile) => new()
        {
            ArtifactId = id,
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = reviewProfile,
            CommentProfile = DeliveryCommentProfile.Endnotes,
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

    private static void WritePersistentGallery(DeliveryBundle bundle)
    {
        var root = Environment.GetEnvironmentVariable(
            "DOCXODUS_DELIVERY_BUNDLE_ARTIFACT_DIR");
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
                return $"<tr><td>{Html(artifact.ArtifactId)}</td><td>{artifact.Kind}</td>"
                    + $"<td>{artifact.Provenance}</td><td>{availability}</td><td>{link}</td></tr>";
            }));
        var index = "<!doctype html><meta charset=\"utf-8\"><title>Docxodus #465 artifacts</title>"
            + "<style>body{font:15px system-ui;margin:2rem;max-width:1100px}"
            + "table{border-collapse:collapse;width:100%}th,td{border:1px solid #bbb;"
            + "padding:.45rem;text-align:left}th{background:#eee}code{word-break:break-all}</style>"
            + "<h1>Docxodus #465 delivery-bundle evidence</h1>"
            + $"<p>Status: <strong>{bundle.Manifest.Payload.Status}</strong>; independent verification: "
            + $"<strong>{bundle.Verification.IsValid}</strong>.</p>"
            + $"<p>Manifest digest: <code>{Html(bundle.Manifest.ManifestDigest.Value)}</code>. "
            + "<a href=\"bundle/bundle-manifest.json\">View canonical manifest</a></p>"
            + "<table><thead><tr><th>ID</th><th>Kind</th><th>Provenance</th>"
            + $"<th>Availability</th><th>Artifact</th></tr></thead><tbody>{rows}</tbody></table>";
        File.WriteAllText(Path.Combine(root, "index.html"), index, Encoding.UTF8);

        var checksums = bundle.Manifest.Payload.Artifacts
            .Where(artifact => artifact.Availability == BundleArtifactAvailability.Available)
            .Select(artifact => $"{artifact.Digest!.Value}  bundle/{artifact.RelativePath}")
            .Append($"{Sha256(bundle.ManifestBytes)}  bundle/{DeliveryBundle.ManifestFileName}");
        File.WriteAllLines(Path.Combine(root, "SHA256SUMS"), checksums, Encoding.UTF8);
    }

    private static string Html(string value) => HtmlEncoder.Default.Encode(value);

    private static string Sha256(byte[] bytes) => Convert.ToHexString(
        SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed class CapturingRenderer : IDeliveryArtifactRenderer
    {
        private readonly List<DeliveryRenderRequest> _requests = new();

        public DeliveryRendererCapabilities Capabilities { get; } = new(
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

        public IReadOnlyList<DeliveryRenderRequest> Requests => _requests.ToArray();

        public ValueTask<DeliveryRenderResult> RenderAsync(
            DeliveryRenderRequest request,
            CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            _requests.Add(request);
            var bytes = request.Kind == DeliveryArtifactKind.StandaloneHtml
                ? Encoding.UTF8.GetBytes($"<!doctype html><p>{request.ArtifactId}</p>")
                : Encoding.ASCII.GetBytes($"%PDF-1.7\n{request.ArtifactId}\n%%EOF");
            return ValueTask.FromResult(DeliveryRenderResult.Available(
                bytes,
                request.Kind == DeliveryArtifactKind.StandaloneHtml
                    ? "text/html"
                    : "application/pdf",
                "test-renderer|engine-1|fonts-1",
                pageCount: 2,
                pageMapBytes: Encoding.UTF8.GetBytes("{\"pages\":[1,2]}"),
                renderReportBytes: Encoding.UTF8.GetBytes("{\"valid\":true}"),
                diagnostics: new[]
                {
                    new DeliverableRenderDiagnostic
                    {
                        Kind = DeliverableRenderDiagnosticKind.Warning,
                        Message = "fixture diagnostic",
                    },
                }));
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
}

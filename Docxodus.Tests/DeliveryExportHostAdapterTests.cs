// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.Verification;
using Xunit;
using BundleArtifactAvailability = Docxodus.Delivery.DeliveryArtifactAvailability;

namespace Docxodus.Tests;

/// <summary>
/// Unit coverage for the framed-host adapter's wire plan — the seam the real integration test
/// drives end to end. These tests never start Node: they exercise exactly the pre-process
/// contract (batch validation, unsafe-version unavailability, and the single framed request).
/// </summary>
public sealed class DeliveryExportHostAdapterTests
{
    private const string UnrepresentableReason = "document_version_unrepresentable";

    private static DocxodusExportHostRenderer Adapter() => new(new DocxodusExportHostRendererOptions
    {
        // Validation requires absolute existing files; these tests never spawn the process,
        // so any stable readable file stands in for the executables.
        NodeExecutablePath = typeof(DeliveryExportHostAdapterTests).Assembly.Location,
        HostScriptPath = typeof(DocxodusExportHostRenderer).Assembly.Location,
    });

    private static DeliveryRenderRequest Render(
        string artifactId,
        DeliveryArtifactKind kind,
        DeliveryReviewProfile review,
        DeliveryCommentProfile comment,
        byte[] source,
        long version) => new(
        artifactId,
        kind,
        review,
        comment,
        new DeliveryDocumentSnapshot("fixture.docx", version, source));

    [Fact]
    public async Task UnsafeDocumentVersion_IsTypedUnavailabilityBeforeAnyHostFrame()
    {
        var adapter = Adapter();
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var context = adapter.DescribeBatch(
            DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden);
        var batch = new DeliveryRenderBatch("render-0001", context, new[]
        {
            Render("unsafe-html", DeliveryArtifactKind.StandaloneHtml,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden,
                source, long.MaxValue),
            Render("unsafe-report", DeliveryArtifactKind.RenderReport,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden,
                source, long.MaxValue),
        });

        // A version outside JavaScript's safe-integer range must come back as the closed
        // per-artifact reason without the adapter ever building a frame or starting the host —
        // the configured "node executable" here is a .NET assembly, so any spawn would fail
        // with a transport error instead of this typed unavailability.
        var results = await adapter.RenderBatchesAsync(new[] { batch });

        Assert.Equal(2, results.Count);
        foreach (var result in results.Values)
        {
            Assert.Equal(BundleArtifactAvailability.Unavailable, result.Availability);
            Assert.Equal(UnrepresentableReason, result.UnavailableReason);
        }
    }

    [Fact]
    public void BuildHostFramePlan_DeduplicatesSharedSourcesAndSortsArtifactIds()
    {
        var adapter = Adapter();
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var finalContext = adapter.DescribeBatch(
            DeliveryReviewProfile.Final, DeliveryCommentProfile.Endnotes);
        var markupContext = adapter.DescribeBatch(
            DeliveryReviewProfile.Markup, DeliveryCommentProfile.Endnotes);
        var batches = new[]
        {
            new DeliveryRenderBatch("render-0001", finalContext, new[]
            {
                Render("z-final-pdf", DeliveryArtifactKind.FinalPdf,
                    DeliveryReviewProfile.Final, DeliveryCommentProfile.Endnotes, source, 3),
                Render("a-final-html", DeliveryArtifactKind.StandaloneHtml,
                    DeliveryReviewProfile.Final, DeliveryCommentProfile.Endnotes, source, 3),
            }),
            new DeliveryRenderBatch("render-0002", markupContext, new[]
            {
                Render("review-pdf", DeliveryArtifactKind.ReviewPdf,
                    DeliveryReviewProfile.Markup, DeliveryCommentProfile.Endnotes, source, 3),
            }),
        };

        var plan = adapter.BuildHostFramePlan(batches);

        // One shared source crosses the pipe once, and its declared identity is the exact
        // SHA-256 of the frame bytes.
        var frame = Assert.Single(plan.SourceFrames);
        using var control = JsonDocument.Parse(plan.ControlFrame);
        var root = control.RootElement;
        Assert.Equal(1, root.GetProperty("schemaVersion").GetInt32());
        var declaredSource = Assert.Single(root.GetProperty("sources").EnumerateArray());
        Assert.Equal(frame.LongLength, declaredSource.GetProperty("byteLength").GetInt64());
        Assert.Equal(
            Convert.ToHexString(SHA256.HashData(frame)).ToLowerInvariant(),
            declaredSource.GetProperty("sha256").GetString());

        var wireBatches = root.GetProperty("batches").EnumerateArray().ToArray();
        Assert.Equal(2, wireBatches.Length);
        var sourceId = declaredSource.GetProperty("id").GetString();
        foreach (var batch in wireBatches)
            Assert.Equal(sourceId, batch.GetProperty("sourceId").GetString());

        // Artifact request IDs are code-unit sorted, the host's canonical ordering.
        var finalBatch = wireBatches.Single(batch =>
            batch.GetProperty("id").GetString() == "render-0001");
        Assert.Equal(
            new[] { "a-final-html", "z-final-pdf" },
            finalBatch.GetProperty("artifactRequestIds").EnumerateArray()
                .Select(id => id.GetString()).ToArray());
        var finalOptions = finalBatch.GetProperty("options");
        Assert.Equal(new[] { "html", "pdf" },
            finalOptions.GetProperty("outputs").EnumerateArray()
                .Select(output => output.GetString()).ToArray());
        Assert.True(finalOptions.GetProperty("reviewProfileAlreadyApplied").GetBoolean());

        var markupOptions = wireBatches.Single(batch =>
            batch.GetProperty("id").GetString() == "render-0002").GetProperty("options");
        Assert.Equal("markup", markupOptions.GetProperty("reviewProfile").GetString());
        Assert.False(markupOptions.TryGetProperty("reviewProfileAlreadyApplied", out _));
        Assert.Equal(new[] { "pdf" },
            markupOptions.GetProperty("outputs").EnumerateArray()
                .Select(output => output.GetString()).ToArray());
    }

    [Fact]
    public async Task RenderBatchesAsync_RejectsAContextTheAdapterDidNotDescribe()
    {
        var adapter = Adapter();
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var described = adapter.DescribeBatch(
            DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden);
        var tampered = described with
        {
            LayoutOptionsDigest = new VerificationDigest
            {
                Algorithm = "SHA-256",
                Value = new string('0', 64),
            },
        };
        var batch = new DeliveryRenderBatch("render-0001", tampered, new[]
        {
            Render("markup-html", DeliveryArtifactKind.StandaloneHtml,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden, source, 1),
        });

        await Assert.ThrowsAsync<ArgumentException>(async () =>
            await adapter.RenderBatchesAsync(new[] { batch }));
    }

    [Fact]
    public void DescribeBatch_IsPureAndProfileSpecific()
    {
        var adapter = Adapter();
        var first = adapter.DescribeBatch(
            DeliveryReviewProfile.Final, DeliveryCommentProfile.Margin);
        var second = adapter.DescribeBatch(
            DeliveryReviewProfile.Final, DeliveryCommentProfile.Margin);
        Assert.Equal(first, second);

        var other = adapter.DescribeBatch(
            DeliveryReviewProfile.Original, DeliveryCommentProfile.Margin);
        Assert.NotEqual(first.LayoutOptionsDigest, other.LayoutOptionsDigest);
        Assert.Equal(first.RuntimePolicyDigest, other.RuntimePolicyDigest);
    }
}

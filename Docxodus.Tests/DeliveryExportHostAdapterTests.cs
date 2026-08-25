// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Docxodus;
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
    public void ParseResponse_AcceptsTheRealSchemaV2RenderReportShape()
    {
        var adapter = Adapter();
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var context = adapter.DescribeBatch(
            DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden);
        var batch = new DeliveryRenderBatch("render-0001", context, new[]
        {
            Render("review-pdf", DeliveryArtifactKind.ReviewPdf,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden, source, 7),
        });
        var plan = adapter.BuildHostFramePlan(new[] { batch });
        var wire = plan.WireBatches[0];

        const string fingerprint = "a24809a09ef1a55f2053eb5de00a331f556aadac973ce83a3df425ee2fdc82d9";
        var pdfBytes = Encoding.ASCII.GetBytes("%PDF-1.4\n%fixture\n%%EOF\n");
        var pageMapBytes = PortablePageMapBytes(fingerprint, documentVersion: 7);
        var reportBytes = SchemaV2Report(wire, fingerprint, pageMapBytes, pdfBytes);
        var response = HostResponse(wire, fingerprint, pdfBytes, pageMapBytes, reportBytes);

        var results = adapter.ParseResponse(response, plan.WireBatches);

        var result = results["review-pdf"];
        Assert.Equal(BundleArtifactAvailability.Available, result.Availability);
        Assert.Equal(fingerprint, result.RendererFingerprint);
        Assert.Equal(1, result.PageCount);
        Assert.Equal(pdfBytes, result.Bytes);
        Assert.Equal(reportBytes, result.RenderReportBytes);
        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal("font_unavailable", diagnostic.Code);
    }

    [Fact]
    public void ParseResponse_NamesTheExactViolatedBindingWhenRejectingAReport()
    {
        var adapter = Adapter();
        var source = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        var context = adapter.DescribeBatch(
            DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden);
        var batch = new DeliveryRenderBatch("render-0001", context, new[]
        {
            Render("review-pdf", DeliveryArtifactKind.ReviewPdf,
                DeliveryReviewProfile.Markup, DeliveryCommentProfile.Hidden, source, 7),
        });
        var plan = adapter.BuildHostFramePlan(new[] { batch });
        var wire = plan.WireBatches[0];
        const string fingerprint = "a24809a09ef1a55f2053eb5de00a331f556aadac973ce83a3df425ee2fdc82d9";
        var pdfBytes = Encoding.ASCII.GetBytes("%PDF-1.4\n%fixture\n%%EOF\n");
        var pageMapBytes = PortablePageMapBytes(fingerprint, documentVersion: 7);

        // A report regressed to the retired v1 header must be named as a version disagreement —
        // this is the pin that keeps the adapter from silently drifting back to v1 expectations.
        var v1Report = Encoding.UTF8.GetString(
            SchemaV2Report(wire, fingerprint, pageMapBytes, pdfBytes))
            .Replace("render/render-report/v2", "render/render-report/v1")
            .Replace("\"schemaVersion\": 2", "\"schemaVersion\": 1");
        var v1Failure = Assert.Throws<InvalidDataException>(() => adapter.ParseResponse(
            HostResponse(wire, fingerprint, pdfBytes, pageMapBytes,
                Encoding.UTF8.GetBytes(v1Report)),
            plan.WireBatches));
        Assert.Contains("render-report/v1", v1Failure.Message);
        Assert.Contains("requires", v1Failure.Message);

        // A genuine binding violation still fails closed, and the message names the field.
        var reboundReport = Encoding.UTF8.GetString(
            SchemaV2Report(wire, fingerprint, pageMapBytes, pdfBytes))
            .Replace("\"documentVersion\": 7", "\"documentVersion\": 8");
        var bindingFailure = Assert.Throws<InvalidDataException>(() => adapter.ParseResponse(
            HostResponse(wire, fingerprint, pdfBytes, pageMapBytes,
                Encoding.UTF8.GetBytes(reboundReport)),
            plan.WireBatches));
        Assert.Contains("bound to document version 8", bindingFailure.Message);
        Assert.Contains("declared 7", bindingFailure.Message);
    }

    /// <summary>A portable paginated PageMap serialized the way the host frames it.</summary>
    private static byte[] PortablePageMapBytes(string fingerprint, long documentVersion)
    {
        var pageMap = new PageMap
        {
            Mode = PageMapMode.Paginated,
            Availability = PageMapAvailability.Available,
            DocumentVersion = documentVersion,
            RendererFingerprint = fingerprint,
            Pages = new[]
            {
                new PageMapPage
                {
                    PageNumber = 1,
                    PageInSection = 1,
                    Width = 612,
                    Height = 792,
                    SectionIndex = 0,
                    PageName = "docxodus-section-0",
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
        return Encoding.UTF8.GetBytes(Docxodus.Internal.DocxSessionJson.SerializePageMap(pageMap));
    }

    /// <summary>
    /// A schema-v2 render report bound to the given wire batch. The envelope — fonts, font
    /// identity, readiness, resources, normalized options with an explicit
    /// reviewProfileAlreadyApplied=false, policy limits, observed environment — is derived from
    /// the real report the export host emitted in CI run 32860996142
    /// (npm-export/test-artifacts/success/cli-render-report.json), trimmed to one entry per
    /// array; the source, profile, digest, and artifact-ID bindings are computed for the batch
    /// under test.
    /// </summary>
    private static byte[] SchemaV2Report(
        Docxodus.Delivery.DocxodusExportHostRenderer.HostWireBatch wire,
        string fingerprint,
        byte[] pageMapBytes,
        byte[] pdfBytes)
    {
        string Sha(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        var artifactIds = string.Join(",", wire.ArtifactRequestIds.Select(id => $"\"{id}\""));
        var report = $$"""
        {
          "schema": "https://docxodus.dev/schemas/render/render-report/v2",
          "schemaVersion": 2,
          "status": "complete",
          "source": {
            "byteLength": {{wire.SourceByteLength}},
            "documentVersion": {{wire.DocumentVersion}},
            "rawPackageBytesDigest": "{{wire.SourceDigest}}"
          },
          "options": {
            "commentProfile": "hidden",
            "layoutDigest": "df1f75b5acaf148d0ab93e59aac7bbdc6a3278f2e09b657304265d63a361248d",
            "outputs": ["pdf"],
            "policy": {
              "limits": { "compressedDocxBytes": 104857600, "finalPages": 10000 },
              "strictFonts": false,
              "timeoutMs": 120000,
              "unsupportedContent": "warn"
            },
            "reviewProfile": "markup",
            "reviewProfileAlreadyApplied": false,
            "runtimePolicyDigest": "9ae2828608823102feca0e00722e83e1497f0685ee58ef326149c20e1c65efa8",
            "title": ""
          },
          "environment": {
            "fidelityTier": "unbaselined",
            "observed": {
              "architecture": "x64",
              "browserBuild": "143.0.7499.4",
              "browserProduct": "Chromium",
              "operatingSystem": "linux",
              "runtimeKind": "nodeChromium"
            },
            "rendererFingerprint": "{{fingerprint}}",
            "verification": "browserObserved"
          },
          "fontIdentity": {
            "resolutionDigest": "6fcb543dda3df25565daddd0b664bae3e2c2e2004ace1662b07e670719eb2d86",
            "resolverContract": "https://docxodus.dev/contracts/font-resolver/v1",
            "substitutionContractDigest": "2dfdd1841ba5d1f2d6064f798922421564ace117cdb999b2156cbc7891dd4548",
            "substitutionContractVersion": 1
          },
          "fonts": [
            {
              "glyphCoverage": "unverified",
              "requestId": "font-0001",
              "requestedFamilies": ["Calibri"],
              "requestedFamily": "Calibri",
              "requestedFamilyKinds": ["named"],
              "requestedStretch": 100,
              "requestedStyle": "normal",
              "requestedWeight": 400,
              "sampleCodePointCount": 20,
              "sampleDigest": "2ee8ac0dac48ddad7919d9be881540018dbfcbe17125ed46a201d6ae6305b2a7",
              "source": "browser",
              "status": "missing",
              "verified": false
            }
          ],
          "readiness": [
            { "elapsedMs": 642.29, "pending": [], "phase": "browser_launch", "status": "complete" }
          ],
          "resources": [],
          "unsupportedContent": "warn",
          "pages": [
            {
              "height": 792, "pageInSection": 1, "pageName": "docxodus-section-0",
              "pageNumber": 1, "sectionIndex": 0, "width": 612
            }
          ],
          "bindings": {
            "artifactRequestIds": [{{artifactIds}}],
            "pageMapDigest": "{{Sha(pageMapBytes)}}",
            "pdfByteDeterministic": false,
            "pdfDigest": "{{Sha(pdfBytes)}}",
            "volatilePdfMetadata": {
              "creationDate": "2026-08-25T14:43:39.000Z",
              "producer": "Skia/PDF m143"
            }
          },
          "warnings": [
            {
              "code": "font_unavailable",
              "message": "1 font family could not be verified.",
              "phase": "font_loading",
              "remediation": "Supply the family through fontDirectories.",
              "resource": "font:Calibri",
              "severity": "warning"
            }
          ]
        }
        """;
        return Encoding.UTF8.GetBytes(report);
    }

    /// <summary>Assembles the host's framed response: control frame, then artifact frames in
    /// descriptor order (pdf, pageMap, renderReport — the host's per-batch emission order).</summary>
    private static byte[] HostResponse(
        Docxodus.Delivery.DocxodusExportHostRenderer.HostWireBatch wire,
        string fingerprint,
        byte[] pdfBytes,
        byte[] pageMapBytes,
        byte[] reportBytes)
    {
        string Sha(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
        object Descriptor(string id, string kind, string mediaType, byte[] bytes) => new
        {
            id,
            batchId = wire.Batch.BatchId,
            kind,
            mediaType,
            byteLength = bytes.Length,
            sha256 = Sha(bytes),
        };
        var control = JsonSerializer.SerializeToUtf8Bytes(new
        {
            schemaVersion = 1,
            batches = new[]
            {
                new
                {
                    id = wire.Batch.BatchId,
                    sourceId = wire.SourceId,
                    pageCount = 1,
                    rendererFingerprint = fingerprint,
                    artifacts = new Dictionary<string, string>
                    {
                        ["pdf"] = "b0-pdf",
                        ["pageMap"] = "b0-pageMap",
                        ["renderReport"] = "b0-renderReport",
                    },
                },
            },
            artifacts = new[]
            {
                Descriptor("b0-pdf", "pdf", "application/pdf", pdfBytes),
                Descriptor("b0-pageMap", "pageMap", "application/json; charset=utf-8", pageMapBytes),
                Descriptor("b0-renderReport", "renderReport", "application/json; charset=utf-8", reportBytes),
            },
        });
        static byte[] Frame(byte[] payload)
        {
            var frame = new byte[payload.Length + 4];
            frame[0] = (byte)(payload.Length >> 24);
            frame[1] = (byte)(payload.Length >> 16);
            frame[2] = (byte)(payload.Length >> 8);
            frame[3] = (byte)payload.Length;
            payload.CopyTo(frame, 4);
            return frame;
        }
        return Frame(control)
            .Concat(Frame(pdfBytes))
            .Concat(Frame(pageMapBytes))
            .Concat(Frame(reportBytes))
            .ToArray();
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

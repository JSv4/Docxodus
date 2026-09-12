// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
using Docxodus.McpServer;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #747: the full deliverable-verification request — options, expected semantic and
/// package changes, companion artifacts — crosses every transport as one wire shape and yields
/// the same canonical report the typed .NET call produces. One fixture drives the equivalence:
/// a baseline, a deliverable with one paragraph changed, no approved package changes (so the
/// change is unexpected under <c>failOnUnexpectedChanges</c>), and a PDF companion whose
/// source-package digest is the baseline's (so it is stale).
/// </summary>
[Collection("MCP session registry isolation")]
public sealed class DeliverableVerificationRequestJsonTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"verify-request-{Guid.NewGuid():N}");
    private readonly byte[] _baseline;
    private readonly byte[] _deliverable;
    private readonly byte[] _companion = System.Text.Encoding.ASCII.GetBytes("%PDF-1.4\n%stale\n");
    private readonly string _expectedSemantic;

    public DeliverableVerificationRequestJsonTests()
    {
        Directory.CreateDirectory(_root);
        _baseline = DocxSessionTests.BuildDS001_SimpleTwoParagraphs();
        using var session = new DocxSession(_baseline);
        var first = session.Project().AnchorIndex.Values
            .First(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "p").Anchor.Id;
        Assert.True(session.ReplaceText(first, "Changed for delivery.").Success);
        _deliverable = session.Save(persistAnchorIds: false);
        // Expectations are sourced from the same byte-level comparison the verifier runs; a
        // live session's opening-to-current listing carries session-minted anchors instead.
        _expectedSemantic = SemanticDiff.Compare(_baseline, _deliverable).ToCanonicalJson();
    }

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private static VerificationDigest Sha256(byte[] bytes) => new()
    {
        Algorithm = "SHA-256",
        Value = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
    };

    /// <summary>The wire request; <paramref name="companionBytesB64"/> null omits the bytes field.</summary>
    private string WireRequest(string? companionBytesB64) =>
        JsonSerializer.Serialize(new
        {
            options = new { failOnUnexpectedChanges = true, mode = "standard" },
            expectedPackageChanges = Array.Empty<object>(),
            companionArtifacts = new[]
            {
                new
                {
                    artifactId = "pdf-1",
                    role = "pdf",
                    mediaType = "application/pdf",
                    bytesB64 = companionBytesB64,
                    pageCount = 1,
                    rendererFingerprint = "renderer/1.0",
                    sourcePackageDigest = new { algorithm = "SHA-256", value = Sha256(_baseline).Value },
                    renderDiagnostics = new[]
                    {
                        new { kind = "missingFont", message = "Aptos substituted", severity = "warning" },
                    },
                },
            },
        });

    private string TypedReport() => DeliverableVerifier.VerifyDeliverable(
        new DeliverableVerificationRequest
        {
            BaselineBytes = _baseline,
            DeliverableBytes = _deliverable,
            ExpectedPackageChanges = Array.Empty<DeliverablePackageChangeExpectation>(),
            CompanionArtifacts = new[]
            {
                new DeliverableCompanionArtifactInput
                {
                    ArtifactId = "pdf-1",
                    Role = DeliverableArtifactRole.Pdf,
                    MediaType = "application/pdf",
                    Availability = DeliverableArtifactAvailability.Available,
                    Bytes = _companion,
                    PageCount = 1,
                    RendererFingerprint = "renderer/1.0",
                    SourcePackageDigest = Sha256(_baseline),
                    RenderDiagnostics = new[]
                    {
                        new DeliverableRenderDiagnostic
                        {
                            Kind = DeliverableRenderDiagnosticKind.MissingFont,
                            Message = "Aptos substituted",
                        },
                    },
                },
            },
        },
        new DeliverableVerificationOptions { FailOnUnexpectedChanges = true }).ToCanonicalJson();

    [Fact]
    public void DVR747a_WireRequest_ProducesTheTypedCallsCanonicalReport_ThroughFacadeHostAndMcp()
    {
        var typed = TypedReport();
        var facade = VerificationOps.VerifyDeliverable(
            _deliverable, _baseline, WireRequest(Convert.ToBase64String(_companion)));
        Assert.Equal(typed, facade);

        using var report = JsonDocument.Parse(typed);
        var codes = report.RootElement.GetProperty("findings").EnumerateArray()
            .Select(f => f.GetProperty("code").GetString()).ToList();
        Assert.Contains("delta.package_change_unexpected", codes);
        Assert.Contains("artifact.source_digest_mismatch", codes);
        Assert.Equal("failed", report.RootElement.GetProperty("decision").GetString());
        Assert.Single(report.RootElement.GetProperty("companionArtifacts").EnumerateArray());

        // The stdio host takes the same object beside its base64 package fields.
        var hostArgs = JsonSerializer.SerializeToElement(new
        {
            docxB64 = Convert.ToBase64String(_deliverable),
            baselineB64 = Convert.ToBase64String(_baseline),
            request = JsonDocument.Parse(WireRequest(Convert.ToBase64String(_companion))).RootElement,
        });
        Assert.Equal(typed, HostVerify(hostArgs));

        // MCP verifies the live session's clean save against its opening bytes; the same
        // request, with the companion named by a path inside the document store.
        var path = Path.Combine(_root, "deliverable.docx");
        File.WriteAllBytes(path, _baseline);
        File.WriteAllBytes(Path.Combine(_root, "evidence.pdf"), _companion);
        var store = new SessionStore(new LocalFileDocumentStore(_root));
        try
        {
            var sessionId = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_open",
                JsonSerializer.SerializeToElement(new { path }))).RootElement
                .GetProperty("sessionId").GetString()!;
            var markdown = JsonDocument.Parse(Dispatcher.Call(store, "docxodus_get_content",
                JsonSerializer.SerializeToElement(new { sessionId, format = "markdown" }))).RootElement;
            var anchor = markdown.GetProperty("anchorIndex").EnumerateObject().First().Name;
            Dispatcher.Call(store, "docxodus_edit", JsonSerializer.SerializeToElement(new
            {
                sessionId, action = "replace_text", anchorId = anchor, markdown = "Changed for delivery.",
            }));
            var mcpRequest = JsonDocument.Parse(WireRequest(null)).RootElement;
            var verification = JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, JsonElement>>(
                mcpRequest.GetRawText())!;
            var companion = JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, JsonElement>>(
                verification["companionArtifacts"][0].GetRawText())!;
            companion.Remove("bytesB64");
            companion["path"] = JsonSerializer.SerializeToElement("evidence.pdf");
            verification["companionArtifacts"] = JsonSerializer.SerializeToElement(new[] { companion });
            var mcp = Dispatcher.Call(store, "docxodus_get_content", JsonSerializer.SerializeToElement(new
            {
                sessionId, format = "verification", verification,
            }));

            // The session's package bytes differ from the fixture's saved bytes only by identity,
            // so compare what the request controls: the finding codes and the decision.
            using var mcpReport = JsonDocument.Parse(mcp);
            var mcpCodes = mcpReport.RootElement.GetProperty("findings").EnumerateArray()
                .Select(f => f.GetProperty("code").GetString()).ToList();
            Assert.Contains("delta.package_change_unexpected", mcpCodes);
            Assert.Contains("artifact.source_digest_mismatch", mcpCodes);
            Assert.Equal("failed", mcpReport.RootElement.GetProperty("decision").GetString());
            Assert.Equal("pdf-1", mcpReport.RootElement.GetProperty("companionArtifacts")[0]
                .GetProperty("artifactId").GetString());
        }
        finally
        {
            store.CloseAll();
        }
    }

    [Fact]
    public void DVR747b_EmptyRequest_IsTheDefaultPolicy()
    {
        Assert.Equal(
            VerificationOps.VerifyDeliverable(_baseline, _deliverable),
            VerificationOps.VerifyDeliverable(_deliverable, _baseline, null));
        Assert.Equal(
            VerificationOps.VerifyDeliverable(_deliverable),
            VerificationOps.VerifyDeliverable(_deliverable, null, "{}"));
    }

    [Fact]
    public void DVR747c_DefaultsComeFromTheModel_AndUnknownPropertiesAreRejected()
    {
        var parsed = DeliverableVerificationRequestJson.Parse("{\"options\":{\"maxFindings\":7}}");
        var defaults = new DeliverableVerificationOptions();
        Assert.Equal(7, parsed.Options.MaxFindings);
        Assert.Equal(defaults.MaxCompanionArtifacts, parsed.Options.MaxCompanionArtifacts);
        Assert.Equal(defaults.PackageManifestOptions.MaxEntryCount, parsed.Options.PackageManifestOptions.MaxEntryCount);
        Assert.Equal(defaults.RequireNoPlaceholders, parsed.Options.RequireNoPlaceholders);

        var unknown = Assert.Throws<ArgumentException>(() =>
            DeliverableVerificationRequestJson.Parse("{\"options\":{\"failOnUnexpectedChange\":true}}"));
        Assert.Contains("failOnUnexpectedChange", unknown.Message);
        Assert.Contains("failOnUnexpectedChanges", unknown.Message);
        Assert.Throws<ArgumentException>(() =>
            DeliverableVerificationRequestJson.Parse("{\"options\":{\"mode\":\"lenient\"}}"));
    }

    [Fact]
    public void DVR747d_LimitsApplyBeforeDecoding()
    {
        // Not base64 at all: a parser that decoded first would raise a format error instead.
        var oversized = new string('!', 4096);
        var tooLarge = Assert.Throws<ArgumentException>(() => DeliverableVerificationRequestJson.Parse(
            "{\"options\":{\"maxCompanionArtifactBytes\":1024},\"companionArtifacts\":["
            + "{\"artifactId\":\"a\",\"role\":\"pdf\",\"mediaType\":\"application/pdf\",\"bytesB64\":\"" + oversized + "\"}]}"));
        Assert.Contains("before decoding", tooLarge.Message);

        var tooMany = Assert.Throws<ArgumentException>(() => DeliverableVerificationRequestJson.Parse(
            "{\"options\":{\"maxCompanionArtifacts\":1},\"companionArtifacts\":["
            + "{\"artifactId\":\"a\",\"role\":\"pdf\",\"mediaType\":\"application/pdf\",\"bytesB64\":\"!!!!\"},"
            + "{\"artifactId\":\"b\",\"role\":\"pdf\",\"mediaType\":\"application/pdf\",\"bytesB64\":\"!!!!\"}]}"));
        Assert.Contains("maxCompanionArtifacts", tooMany.Message);

        var tooManyDiagnostics = Assert.Throws<ArgumentException>(() => DeliverableVerificationRequestJson.Parse(
            "{\"options\":{\"maxRenderDiagnostics\":1},\"companionArtifacts\":["
            + "{\"artifactId\":\"a\",\"role\":\"pdf\",\"mediaType\":\"application/pdf\",\"renderDiagnostics\":["
            + "{\"kind\":\"warning\",\"message\":\"x\"},{\"kind\":\"warning\",\"message\":\"y\"}]}]}"));
        Assert.Contains("maxRenderDiagnostics", tooManyDiagnostics.Message);

        var tooManyExpected = Assert.Throws<ArgumentException>(() => DeliverableVerificationRequestJson.Parse(
            "{\"options\":{\"maxExpectedChanges\":1},\"expectedPackageChanges\":["
            + "{\"kind\":\"entryModified\",\"location\":{\"entryUri\":\"/word/document.xml\"}},"
            + "{\"kind\":\"entryModified\",\"location\":{\"entryUri\":\"/word/styles.xml\"}}]}"));
        Assert.Contains("maxExpectedChanges", tooManyExpected.Message);
    }

    [Fact]
    public void DVR747e_ExpectedSemanticChanges_RoundTripFromTheSessionsOwnOutput()
    {
        // Re-serialized by a client (indented): parsed structurally, not byte-exact.
        var pretty = JsonSerializer.Serialize(JsonDocument.Parse(_expectedSemantic).RootElement,
            new JsonSerializerOptions { WriteIndented = true });
        var request = "{\"options\":{\"failOnUnexpectedChanges\":true},\"expectedSemanticChanges\":" + pretty + "}";

        using var stateless = JsonDocument.Parse(VerificationOps.VerifyDeliverable(_deliverable, _baseline, request));
        AssertSemanticDeltaApproved(stateless.RootElement);

        // The session path: the same edit on a live session, verified against its opening bytes.
        var handle = DocxSessionOps.OpenSession(_baseline, null);
        try
        {
            using var projection = JsonDocument.Parse(DocxSessionOps.Project(handle));
            var anchor = projection.RootElement.GetProperty("anchorIndex").EnumerateObject()
                .First(e => e.Value.GetProperty("scope").GetString() == "body"
                    && e.Value.GetProperty("kind").GetString() == "p").Name;
            DocxSessionOps.ReplaceText(handle, anchor, "Changed for delivery.");
            using var live = JsonDocument.Parse(DocxSessionOps.VerifyDeliverable(handle, request));
            AssertSemanticDeltaApproved(live.RootElement);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    private static void AssertSemanticDeltaApproved(JsonElement report)
    {
        var deltaFindings = report.GetProperty("findings").EnumerateArray()
            .Where(f => f.GetProperty("code").GetString()!.StartsWith("delta.semantic", StringComparison.Ordinal))
            .Select(f => f.GetProperty("code").GetString() + ": " + f.GetProperty("message").GetString())
            .ToList();
        Assert.True(deltaFindings.Count == 0, string.Join("\n", deltaFindings));
        Assert.True(report.GetProperty("semanticDelta").GetProperty("changeCount").GetInt32() > 0);
    }

    private static string HostVerify(JsonElement args) =>
        Docxodus.PyHost.Dispatcher.Dispatch("verify_deliverable", args);
}

// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Internal;
using Docxodus.McpServer;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Transport-seam coverage for the redline reversibility proof. The proof engine itself is
/// covered by <see cref="RedlineReversibilityProofTests"/>; these tests assert only that the
/// shared facade and the MCP action hand callers the same canonical schema-v1 document the
/// in-process verifier produces, with the inputs they actually named.
/// </summary>
public sealed class RedlineReversibilityTransportTests
{
    private const string Schema =
        "https://docxodus.dev/schemas/verification/redline-reversibility-proof/v1";

    [Fact]
    public void RT001_FacadeReturnsCanonicalDeterministicProofBoundToItsThreeInputs()
    {
        var (baseline, intendedFinal, redline) = Triple();

        var json = VerificationOps.ProveRedlineReversibility(baseline, intendedFinal, redline);

        // Determinism is the property a receipt digest depends on.
        Assert.Equal(json, VerificationOps.ProveRedlineReversibility(
            baseline, intendedFinal, redline));

        // The facade must not reshape the proof: it is byte-identical to the engine's canonical
        // form, so a receipt that hashes one hashes the other.
        Assert.Equal(
            RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline)
                .Proof.ToCanonicalJson(),
            json);

        using var proof = JsonDocument.Parse(json);
        Assert.Equal(Schema, proof.RootElement.GetProperty("schema").GetString());
        Assert.Equal(1, proof.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(Digest(baseline), RawDigest(proof.RootElement, "baselinePackage"));
        Assert.Equal(Digest(intendedFinal), RawDigest(proof.RootElement, "intendedFinalPackage"));
        Assert.Equal(Digest(redline), RawDigest(proof.RootElement, "redlinePackage"));
        Assert.NotEqual(JsonValueKind.Null,
            proof.RootElement.GetProperty("acceptToFinal").ValueKind);
        Assert.NotEqual(JsonValueKind.Null,
            proof.RootElement.GetProperty("rejectToBaseline").ValueKind);
        Assert.NotEmpty(proof.RootElement.GetProperty("revisionClassifications").EnumerateArray());
    }

    [Fact]
    public void RT002_LoweredProofLimitConstrainsTheProofRatherThanRejectingItAfterwards()
    {
        var (baseline, intendedFinal, redline) = Triple();

        var json = VerificationOps.ProveRedlineReversibility(
            baseline,
            intendedFinal,
            redline,
            new RedlineReversibilityProofOptions { MaxRevisionElements = 1 });

        using var proof = JsonDocument.Parse(json);
        Assert.False(proof.RootElement.GetProperty("success").GetBoolean());
        // Fail-closed: the paths are not attempted at all rather than reported on partial work.
        Assert.Equal(JsonValueKind.Null, proof.RootElement.GetProperty("acceptToFinal").ValueKind);
        Assert.Equal(JsonValueKind.Null,
            proof.RootElement.GetProperty("rejectToBaseline").ValueKind);
        Assert.Contains(
            proof.RootElement.GetProperty("findings").EnumerateArray(),
            finding => finding.GetProperty("severity").GetString() == "error");
    }

    [Fact]
    public void RT003_McpProvesTheSessionCheckpointAgainstTwoInScopeDocuments()
    {
        var (baseline, intendedFinal, redline) = Triple();
        var root = Path.Combine(Path.GetTempPath(), $"reversibility-transport-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            var baselinePath = Path.Combine(root, "baseline.docx");
            var finalPath = Path.Combine(root, "final.docx");
            var redlinePath = Path.Combine(root, "redline.docx");
            File.WriteAllBytes(baselinePath, baseline);
            File.WriteAllBytes(finalPath, intendedFinal);
            File.WriteAllBytes(redlinePath, redline);

            using var opened = JsonDocument.Parse(Dispatcher.Call(
                store, "docxodus_open", Json($$"""{"path":{{Quote(redlinePath)}}}""")));
            var sessionId = Quote(opened.RootElement.GetProperty("sessionId").GetString()!);

            var proofJson = Dispatcher.Call(store, "docxodus_track_changes", Json($$"""
                {"sessionId":{{sessionId}},"action":"prove_reversibility",
                 "baselinePath":{{Quote(baselinePath)}},
                 "intendedFinalPath":{{Quote(finalPath)}}}
                """));

            using var proof = JsonDocument.Parse(proofJson);
            Assert.Equal(Schema, proof.RootElement.GetProperty("schema").GetString());
            Assert.Equal(Digest(baseline), RawDigest(proof.RootElement, "baselinePackage"));
            Assert.Equal(Digest(intendedFinal), RawDigest(proof.RootElement, "intendedFinalPackage"));

            // The package under proof is the session's clean-save checkpoint — the bytes an agent
            // would ship — so it is a real digest, and proving twice yields the same document.
            Assert.Matches("^[0-9a-f]{64}$", RawDigest(proof.RootElement, "redlinePackage")!);
            Assert.Equal(proofJson, Dispatcher.Call(store, "docxodus_track_changes", Json($$"""
                {"sessionId":{{sessionId}},"action":"prove_reversibility",
                 "baselinePath":{{Quote(baselinePath)}},
                 "intendedFinalPath":{{Quote(finalPath)}}}
                """)));

            // A location outside the store's scope is refused by the store's containment check,
            // not by a failed read: the file exists and holds a valid package, so only the
            // resolver can be the thing that rejects it.
            var outsidePath = Path.Combine(Path.GetTempPath(),
                $"reversibility-outside-{Guid.NewGuid():N}.docx");
            File.WriteAllBytes(outsidePath, baseline);
            try
            {
                var escape = Assert.Throws<McpToolException>(() => Dispatcher.Call(
                    store, "docxodus_track_changes", Json($$"""
                        {"sessionId":{{sessionId}},"action":"prove_reversibility",
                         "baselinePath":{{Quote(outsidePath)}},
                         "intendedFinalPath":{{Quote(finalPath)}}}
                        """)));
                Assert.Contains("outside this server's document scope", escape.Message);
            }
            finally
            {
                File.Delete(outsidePath);
            }
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void RT004_ProveReversibilityIsAdvertisedAndIsNotBatchable()
    {
        var trackChanges = Assert.Single(
            ToolCatalog.Tools.Where(tool => tool.Name == "docxodus_track_changes"));
        using var schema = JsonDocument.Parse(trackChanges.InputSchemaJson);
        var properties = schema.RootElement.GetProperty("properties");
        Assert.Contains("prove_reversibility", properties
            .GetProperty("action").GetProperty("enum").EnumerateArray()
            .Select(value => value.GetString()));
        Assert.True(properties.TryGetProperty("baselinePath", out _));
        Assert.True(properties.TryGetProperty("intendedFinalPath", out _));

        // Read-only evidence must stay out of the mutation batch: docxodus_mutations drives
        // RunTrackChangesAction, which never sees this action.
        var mutations = Assert.Single(
            ToolCatalog.Tools.Where(tool => tool.Name == "docxodus_mutations"));
        using var batchSchema = JsonDocument.Parse(mutations.InputSchemaJson);
        Assert.DoesNotContain("prove_reversibility", batchSchema.RootElement.ToString());
    }

    [Fact]
    public void RT005_BatchStepRefusalIsBehavioralNotJustSchemaText()
    {
        var root = Path.Combine(Path.GetTempPath(), $"reversibility-batch-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            var path = Path.Combine(root, "document.docx");
            File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
            using var opened = JsonDocument.Parse(Dispatcher.Call(
                store, "docxodus_open", Json($$"""{"path":{{Quote(path)}}}""")));
            var sessionId = Quote(opened.RootElement.GetProperty("sessionId").GetString()!);

            // The whitelist in ValidateMutationBatchAction, not just the tool schema, must
            // refuse the action — nothing mutated, the batch reports the step as invalid.
            var response = Dispatcher.Call(store, "docxodus_mutations", Json($$"""
                {"sessionId":{{sessionId}},"steps":[
                    {"tool":"docxodus_track_changes","args":{"action":"prove_reversibility"} }]}
                """));
            using var batch = JsonDocument.Parse(response);
            Assert.False(batch.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains(
                "unsupported or read-only batch action: docxodus_track_changes/prove_reversibility",
                response);
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }

    private static (byte[] Baseline, byte[] IntendedFinal, byte[] Redline) Triple()
    {
        var baseline = Document("The original clause.");
        var intendedFinal = Document("The revised clause.");
        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal),
            new DocxDiffSettings { AuthorForRevisions = "Comparison Engine" }).DocumentByteArray;
        return (baseline, intendedFinal, redline);
    }

    private static byte[] Document(string text)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            document.Save();
        }
        return stream.ToArray();
    }

    private static string Digest(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string? RawDigest(JsonElement root, string package) => root
        .GetProperty(package)
        .GetProperty("rawPackageBytesDigest")
        .GetProperty("value")
        .GetString();

    private static string Quote(string value) => JsonSerializer.Serialize(value);

    private static JsonElement Json(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

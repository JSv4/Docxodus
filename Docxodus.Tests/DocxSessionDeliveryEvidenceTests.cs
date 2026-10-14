// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.Delivery;
using Docxodus.Internal;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #748: a session opened with <c>CaptureDeliveryEvidence</c> records every version step
/// as it executes, and a delivery mints a change receipt from that history — or says exactly why
/// it cannot.
/// </summary>
[Collection("MCP session registry isolation")]
public class DocxSessionDeliveryEvidenceTests
{
    [Fact]
    public void DE748a_CapturedHistory_MintsAVerifiableReceipt()
    {
        using var session = Open();
        var anchors = BodyParagraphs(session);

        var atomic = session.ExecuteBatch(new[]
        {
            Step("docx_edit", "replace_text", new { anchorId = anchors[0], markdown = "First." },
                s => s.ReplaceText(anchors[0], "First.")),
            Step("docx_edit", "replace_text", new { anchorId = anchors[1], markdown = "Second." },
                s => s.ReplaceText(anchors[1], "Second.")),
        });
        Assert.True(atomic.Success, Describe(atomic));

        // A typed direct call: exact packages, unknown request.
        Assert.True(session.InsertParagraph(anchors[1], Position.After, "Direct paragraph.").Success);
        _ = session.Save();

        var bestEffort = session.ExecuteBatch(new[]
        {
            Step("docx_edit", "replace_text", new { anchorId = anchors[0], markdown = "Third." },
                s => s.ReplaceText(anchors[0], "Third.")),
            Step("docx_edit", "replace_text", new { anchorId = anchors[1], markdown = "Fourth." },
                s => s.ReplaceText(anchors[1], "Fourth.")),
        }, MutationBatchMode.BestEffort);
        Assert.True(bestEffort.Success, Describe(bestEffort));

        var failed = session.ExecuteBatch(new[]
        {
            Step("docx_edit", "replace_text", new { anchorId = "p:body:missing", markdown = "Never." },
                s => s.ReplaceText("p:body:missing", "Never.")),
        });
        Assert.False(failed.Success);

        Assert.True(session.Undo());
        Assert.True(session.Redo());

        var status = session.GetDeliveryEvidenceStatus();
        Assert.True(status.Enabled);
        Assert.Null(status.UnavailableReason);
        Assert.Equal(5, status.TransactionCount);
        Assert.Equal(2, status.LineageEventCount);
        Assert.Equal(1, status.UnlabeledTransactionCount);
        Assert.Equal(session.Version, status.CurrentVersion);

        var bundle = session.BuildDeliveryReceipt();
        Assert.Equal(DeliveryBundleStatus.Complete, bundle.Manifest.Payload.Status);
        Assert.True(bundle.Verification.IsValid, string.Join("; ", bundle.Verification.Findings));
        var receiptBytes = bundle.GetArtifactBytes("change-receipt");
        var verification = DeliveryChangeReceiptVerifier.VerifyJson(receiptBytes, ReceiptArtifacts(bundle));
        Assert.True(verification.IsValid, string.Join("; ", verification.Findings));

        using var receipt = JsonDocument.Parse(receiptBytes);
        var payload = receipt.RootElement.GetProperty("payload");
        var transactions = payload.GetProperty("transactions").EnumerateArray().ToArray();
        Assert.Equal(5, transactions.Length);
        Assert.Equal(
            new[] { "committed", "committed", "committed", "committed", "failed" },
            transactions.Select(t => t.GetProperty("status").GetString()));
        Assert.Equal(
            new[] { "atomic", "atomic", "bestEffort", "bestEffort", "atomic" },
            transactions.Select(t => t.GetProperty("mode").GetString()));
        Assert.Equal(2, transactions[0].GetProperty("operations").GetArrayLength());
        var unlabeled = transactions[1].GetProperty("operations")[0];
        Assert.Equal("docx_session", unlabeled.GetProperty("tool").GetString());
        Assert.Equal("unlabeled_mutation", unlabeled.GetProperty("action").GetString());
        Assert.Equal(
            new[] { "undo", "redo" },
            payload.GetProperty("lineage").EnumerateArray().Select(l => l.GetProperty("action").GetString()));
        Assert.Equal(0, payload.GetProperty("sourceDocument").GetProperty("documentVersion").GetInt64());
        Assert.Equal(session.Version, payload.GetProperty("deliveredDocument").GetProperty("documentVersion").GetInt64());
        Assert.Contains(bundle.Manifest.Payload.Artifacts,
            artifact => artifact.Kind == DeliveryArtifactKind.SemanticDelta
                && artifact.ArtifactId.StartsWith("semantic-transaction-", StringComparison.Ordinal));

        // The delivered document is the recorded current state, opened as a document it carries
        // exactly the committed edits.
        using var delivered = new DocxSession(bundle.GetArtifactBytes("final-docx"));
        var deliveredMarkdown = delivered.Project().Markdown;
        Assert.Contains("Third.", deliveredMarkdown, StringComparison.Ordinal);
        Assert.Contains("Fourth.", deliveredMarkdown, StringComparison.Ordinal);
        Assert.Contains("Direct paragraph.", deliveredMarkdown, StringComparison.Ordinal);
        Assert.DoesNotContain("Never.", deliveredMarkdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DE748b_CleanCheckpoint_HasTheCleanSavePayloadsAndDoesNotTouchTheSession()
    {
        using var session = Open(capture: false);
        var anchors = BodyParagraphs(session);
        Assert.True(session.ReplaceText(anchors[0], "Edited before the checkpoint.").Success);
        var version = session.Version;
        var hash = session.GetPackageContentHash();

        var checkpoint = session.SerializeCleanCheckpoint();

        // Part payloads are the clean save's (projector bookkeeping stripped); only the
        // package-level relationship and content-type files may be re-serialized by the clone.
        Assert.Equal(PartPayloads(session.Save(persistAnchorIds: false)), PartPayloads(checkpoint));
        Assert.Equal(version, session.Version);
        Assert.Equal(hash, session.GetPackageContentHash());
        Assert.Contains("Edited before the checkpoint.", session.Project().Markdown, StringComparison.Ordinal);
    }

    private static Dictionary<string, string> PartPayloads(byte[] package)
    {
        using var archive = new System.IO.Compression.ZipArchive(new System.IO.MemoryStream(package));
        return archive.Entries
            .Where(entry => entry.FullName.StartsWith("word/", StringComparison.Ordinal)
                && !entry.FullName.Contains("/_rels/", StringComparison.Ordinal))
            .ToDictionary(entry => entry.FullName, entry =>
            {
                using var stream = new System.IO.MemoryStream();
                entry.Open().CopyTo(stream);
                return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(stream.ToArray()));
            }, StringComparer.Ordinal);
    }

    [Fact]
    public void DE748c_TransactionalRetry_IsOneEntryUnderItsIdentity()
    {
        var handle = DocxSessionOps.OpenSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = true, CaptureDeliveryEvidence = true });
        try
        {
            var anchor = FirstAnchor(handle);
            var request = JsonSerializer.SerializeToElement(new
            {
                handle,
                mode = "atomic",
                steps = new[] { new { operation = "insert_paragraph", args = new { anchorId = anchor, markdown = "Once." } } },
            });
            IEnumerable<MutationBatchStep> Steps() => new[]
            {
                DocxSessionOps.SerializedBatchStep("docx_scalpel", "insert_paragraph",
                    () => DocxSessionOps.InsertParagraph(handle, anchor, Position.After, "Once."),
                    argumentsJson: JsonSerializer.Serialize(new { anchorId = anchor, markdown = "Once." })),
            };
            var first = DocxSessionOps.ExecuteBatchTransactional(handle, "tx-1", request, MutationBatchMode.Atomic, Steps);
            var retry = DocxSessionOps.ExecuteBatchTransactional(handle, "tx-1", request, MutationBatchMode.Atomic, Steps);
            Assert.Equal(first, retry);

            using var status = JsonDocument.Parse(DocxSessionOps.GetDeliveryEvidenceStatus(handle));
            Assert.Equal(1, status.RootElement.GetProperty("transactionCount").GetInt32());
            Assert.True(status.RootElement.GetProperty("unavailableReason").ValueKind == JsonValueKind.Null);

            using var bundle = JsonDocument.Parse(DocxSessionOps.BuildDeliveryReceipt(handle, null));
            Assert.Equal("complete", bundle.RootElement.GetProperty("status").GetString());
            Assert.True(bundle.RootElement.GetProperty("verified").GetBoolean());
            var receiptArtifact = bundle.RootElement.GetProperty("artifacts").EnumerateArray()
                .Single(artifact => artifact.GetProperty("artifactId").GetString() == "change-receipt");
            Assert.Equal("available", receiptArtifact.GetProperty("availability").GetString());
            using var receipt = JsonDocument.Parse(Convert.FromBase64String(receiptArtifact.GetProperty("bytes").GetString()!));
            var entry = Assert.Single(receipt.RootElement.GetProperty("payload").GetProperty("transactions").EnumerateArray());
            Assert.Equal("tx-1", entry.GetProperty("transactionId").GetString());
            Assert.StartsWith("sha256:", entry.GetProperty("requestFingerprint").GetString(), StringComparison.Ordinal);
            var operation = Assert.Single(entry.GetProperty("operations").EnumerateArray());
            Assert.Equal("insert_paragraph", operation.GetProperty("action").GetString());
            Assert.True(bundle.RootElement.GetProperty("evidence").GetProperty("enabled").GetBoolean());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DE748d_RetentionBounds_MakeTheReceiptExplicitlyUnavailable()
    {
        using var session = Open();
        session.DeliveryEvidence!.SetLimits(maxStates: 2, byteBudget: long.MaxValue);
        var anchors = BodyParagraphs(session);
        Assert.True(session.ReplaceText(anchors[0], "One.").Success);
        Assert.True(session.ReplaceText(anchors[0], "Two.").Success);

        var status = session.GetDeliveryEvidenceStatus();
        Assert.True(status.Enabled);
        Assert.Contains("retention exceeded", status.UnavailableReason, StringComparison.Ordinal);
        Assert.Equal(0, status.TransactionCount);

        var bundle = session.BuildDeliveryReceipt();
        Assert.Equal(DeliveryBundleStatus.Incomplete, bundle.Manifest.Payload.Status);
        var receipt = Assert.Single(bundle.Manifest.Payload.Artifacts, a => a.Kind == DeliveryArtifactKind.ChangeReceipt);
        Assert.Equal(Delivery.DeliveryArtifactAvailability.Unavailable, receipt.Availability);
        Assert.NotNull(receipt.UnavailableReason);
        // The delivered document itself is still exact.
        using var delivered = new DocxSession(bundle.GetArtifactBytes("final-docx"));
        Assert.Contains("Two.", delivered.Project().Markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void DE748e_NotEnabled_IsExplicitOnEverySurface()
    {
        using var session = Open(capture: false);
        var status = session.GetDeliveryEvidenceStatus();
        Assert.False(status.Enabled);
        Assert.Equal(DocxSession.NotCapturingDeliveryEvidence, status.UnavailableReason);

        var export = session.ExportDeliveryEvidence();
        Assert.Null(export.ReceiptContext);
        Assert.Equal(DocxSession.NotCapturingDeliveryEvidence, export.UnavailableReason);

        var bundle = session.BuildDeliveryReceipt();
        Assert.Equal(DeliveryBundleStatus.Incomplete, bundle.Manifest.Payload.Status);
        Assert.Equal(Delivery.DeliveryArtifactAvailability.Unavailable,
            Assert.Single(bundle.Manifest.Payload.Artifacts, a => a.Kind == DeliveryArtifactKind.ChangeReceipt).Availability);

        Assert.Throws<ArgumentException>(() => new DocxSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { CaptureDeliveryEvidence = true, CaptureInitialProjection = false }));
    }

    [Fact]
    public void DE748f_TamperedArtifacts_FailVerification()
    {
        using var session = Open();
        var anchors = BodyParagraphs(session);
        Assert.True(session.ExecuteBatch(new[]
        {
            Step("docx_edit", "replace_text", new { anchorId = anchors[0], markdown = "Delivered." },
                s => s.ReplaceText(anchors[0], "Delivered.")),
        }).Success);
        var bundle = session.BuildDeliveryReceipt();
        var receiptBytes = bundle.GetArtifactBytes("change-receipt");
        var artifacts = ReceiptArtifacts(bundle);
        Assert.True(DeliveryChangeReceiptVerifier.VerifyJson(receiptBytes, artifacts).IsValid);

        var tampered = new Dictionary<string, byte[]>(artifacts, StringComparer.Ordinal);
        var docx = tampered["final-docx"].ToArray();
        docx[^1] ^= 0x5A;
        tampered["final-docx"] = docx;
        var altered = DeliveryChangeReceiptVerifier.VerifyJson(receiptBytes, tampered);
        Assert.False(altered.IsValid);
        Assert.Contains(altered.Artifacts, artifact => artifact.ArtifactId == "final-docx"
            && artifact.Status != DeliveryArtifactVerificationStatus.Verified);

        var edited = System.Text.Encoding.UTF8.GetString(receiptBytes).Replace("\"committed\"", "\"failed\"", StringComparison.Ordinal);
        Assert.False(DeliveryChangeReceiptVerifier.VerifyJson(
            System.Text.Encoding.UTF8.GetBytes(edited), artifacts).IsValid);
    }

    [Fact]
    public void DE748g_BestEffortIdentity_RidesOnTheFirstStepEntry()
    {
        var handle = DocxSessionOps.OpenSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = true, CaptureDeliveryEvidence = true });
        try
        {
            var anchor = FirstAnchor(handle);
            var request = JsonSerializer.SerializeToElement(new { handle, mode = "best_effort", steps = new[] { "a", "b" } });
            _ = DocxSessionOps.ExecuteBatchTransactional(handle, "tx-best", request, MutationBatchMode.BestEffort, () => new[]
            {
                DocxSessionOps.SerializedBatchStep("docx_scalpel", "insert_paragraph",
                    () => DocxSessionOps.InsertParagraph(handle, anchor, Position.After, "A.")),
                DocxSessionOps.SerializedBatchStep("docx_scalpel", "replace_text",
                    () => DocxSessionOps.ReplaceText(handle, "p:body:missing", "B.")),
            });

            using var bundle = JsonDocument.Parse(DocxSessionOps.BuildDeliveryReceipt(handle, "{\"privacyProfile\":\"full_evidence\"}"));
            var receiptArtifact = bundle.RootElement.GetProperty("artifacts").EnumerateArray()
                .Single(artifact => artifact.GetProperty("artifactId").GetString() == "change-receipt");
            Assert.Equal("available", receiptArtifact.GetProperty("availability").GetString());
            using var receipt = JsonDocument.Parse(Convert.FromBase64String(receiptArtifact.GetProperty("bytes").GetString()!));
            var payload = receipt.RootElement.GetProperty("payload");
            Assert.Equal("fullEvidence", payload.GetProperty("privacyProfile").GetString());
            var entries = payload.GetProperty("transactions").EnumerateArray().ToArray();
            Assert.Equal(2, entries.Length);
            Assert.Equal("tx-best", entries[0].GetProperty("transactionId").GetString());
            Assert.Equal("committed", entries[0].GetProperty("status").GetString());
            Assert.Equal(JsonValueKind.Null, entries[1].GetProperty("transactionId").ValueKind);
            Assert.Equal("failed", entries[1].GetProperty("status").GetString());
            Assert.Contains(payload.GetProperty("warnings").EnumerateArray(),
                warning => (warning.GetProperty("value").GetString() ?? string.Empty).Contains("per-step", StringComparison.Ordinal));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DE748h_CommittedPreview_IsALabeledTransaction()
    {
        using var session = Open();
        var anchors = BodyParagraphs(session);
        var preview = session.PreviewBatch(new[]
        {
            new MutationBatchStep("docx_edit", "replace_text", s => s.ReplaceText(anchors[0], "Previewed.")),
        }, options: new MutationBatchPreviewOptions { Retain = true });
        Assert.True(session.CommitPreview(preview.Retention!.PreviewId).Success);

        var status = session.GetDeliveryEvidenceStatus();
        Assert.Null(status.UnavailableReason);
        Assert.Equal(1, status.TransactionCount);
        Assert.Equal(0, status.UnlabeledTransactionCount);
        var export = session.ExportDeliveryEvidence();
        var entry = Assert.Single(export.ReceiptContext!.Transactions);
        var operation = Assert.Single(entry.Contribution.Operations);
        Assert.Equal("docx_session", operation.Tool);
        Assert.Equal("commit_preview", operation.Action);
    }

    private static MutationBatchStep Step(string tool, string action, object args, Func<DocxSession, EditResult> mutation) =>
        new(tool, action, mutation) { ArgumentsJson = JsonSerializer.Serialize(args) };

    private static Dictionary<string, byte[]> ReceiptArtifacts(DeliveryBundle bundle) =>
        bundle.Manifest.Payload.Artifacts
            .Where(a => a.Availability == Delivery.DeliveryArtifactAvailability.Available && a.Kind != DeliveryArtifactKind.ChangeReceipt)
            .ToDictionary(a => a.ArtifactId, a => bundle.GetArtifactBytes(a.ArtifactId), StringComparer.Ordinal);

    private static string FirstAnchor(int handle)
    {
        using var projection = JsonDocument.Parse(DocxSessionOps.Project(handle));
        return projection.RootElement.GetProperty("anchorIndex").EnumerateObject()
            .First(property => property.Name.StartsWith("p:body:", StringComparison.Ordinal)).Name;
    }

    private static string Describe(MutationBatchResult result) =>
        result.Failure is null
            ? "no failure envelope"
            : $"{result.Failure.Index}:{result.Failure.Action}:{result.Failure.Error.Code}:{result.Failure.Error.Message}";

    private static DocxSession Open(bool capture = true) =>
        new(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { PersistAnchorIds = true, CaptureDeliveryEvidence = capture });

    private static string[] BodyParagraphs(DocxSession session) =>
        session.Project().AnchorIndex.Keys
            .Where(id => id.StartsWith("p:body:", StringComparison.Ordinal))
            .ToArray();
}

// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #748 over the MCP server: a session opened with <c>captureDeliveryEvidence</c> records
/// its direct tool calls and batches, and <c>docxodus_deliver</c> mints an available change
/// receipt that <c>docxodus_verify_receipt</c> verifies.
/// </summary>
[Collection("MCP session registry isolation")]
public sealed class McpDeliveryEvidenceTests : IDisposable
{
    private readonly string _root;
    private readonly string _path;
    private readonly SessionStore _store;

    public McpDeliveryEvidenceTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-delivery-evidence-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _path = Path.Combine(_root, "document.docx");
        File.WriteAllBytes(_path, DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void MCP748a_DirectCallsAndBatches_AreDescribedTransactionsInAVerifiedReceipt()
    {
        var sessionId = OpenSession(capture: true);
        var anchor = FirstAnchor(sessionId);

        // A direct tool call, described by the dispatcher.
        var direct = J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "replace_text",
            anchorId = anchor,
            markdown = "Directly edited.",
        }))));
        Assert.True(direct.GetProperty("success").GetBoolean());

        // A transactional batch, described by its steps and its journal identity.
        var batchArgs = JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = "tx-deliver",
            steps = new[]
            {
                new
                {
                    tool = "docxodus_create",
                    args = new { action = "insert_paragraph", anchorId = anchor, position = "after", markdown = "Batched." },
                },
            },
        });
        var first = Dispatcher.Call(_store, "docxodus_mutations", J(batchArgs));
        Assert.Equal(first, Dispatcher.Call(_store, "docxodus_mutations", J(batchArgs)));

        // A failed direct call: recorded as a failure that changed nothing.
        var failed = J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "replace_text",
            anchorId = "p:body:missing",
            markdown = "Never.",
        }))));
        Assert.False(failed.GetProperty("success").GetBoolean());

        // Undo/redo become lineage; a save is not a transaction.
        Assert.True(J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new { sessionId, action = "undo" }))))
            .GetProperty("success").GetBoolean());
        Assert.True(J(Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new { sessionId, action = "redo" }))))
            .GetProperty("success").GetBoolean());
        _ = Dispatcher.Call(_store, "docxodus_save", J(JsonSerializer.Serialize(new { sessionId, path = "saved.docx" })));

        var version = DocxSessionOps.GetVersion(_store.Get(sessionId).Handle);
        var bundle = J(Dispatcher.Call(_store, "docxodus_deliver", J(DeliverArgs(sessionId, finalVersion: version))));
        var evidence = bundle.GetProperty("evidence");
        Assert.True(evidence.GetProperty("unavailableReason").ValueKind == JsonValueKind.Null,
            evidence.GetProperty("unavailableReason").ToString());
        Assert.Equal("complete", bundle.GetProperty("status").GetString());
        Assert.True(bundle.GetProperty("verified").GetBoolean());
        Assert.True(evidence.GetProperty("enabled").GetBoolean());
        Assert.Equal(3, evidence.GetProperty("transactionCount").GetInt32());
        Assert.Equal(2, evidence.GetProperty("lineageEventCount").GetInt32());
        Assert.Equal(0, evidence.GetProperty("unlabeledTransactionCount").GetInt32());

        var artifacts = bundle.GetProperty("artifacts").EnumerateArray().ToArray();
        var receiptArtifact = artifacts.Single(a => a.GetProperty("artifactId").GetString() == "receipt");
        Assert.Equal("available", receiptArtifact.GetProperty("availability").GetString());
        var receiptBytes = Convert.FromBase64String(receiptArtifact.GetProperty("bytes").GetString()!);

        using var receipt = JsonDocument.Parse(receiptBytes);
        var payload = receipt.RootElement.GetProperty("payload");
        var transactions = payload.GetProperty("transactions").EnumerateArray().ToArray();
        Assert.Equal(3, transactions.Length);
        Assert.Equal("docxodus_edit", transactions[0].GetProperty("operations")[0].GetProperty("tool").GetString());
        Assert.Equal("replace_text", transactions[0].GetProperty("operations")[0].GetProperty("action").GetString());
        Assert.Equal("tx-deliver", transactions[1].GetProperty("transactionId").GetString());
        Assert.Equal("failed", transactions[2].GetProperty("status").GetString());
        Assert.Equal(2, payload.GetProperty("lineage").GetArrayLength());

        // The portable verifier accepts the receipt against the bundle's own artifact bytes.
        var receiptPath = Path.Combine(_root, "receipt.json");
        File.WriteAllBytes(receiptPath, receiptBytes);
        var artifactPaths = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var artifact in artifacts)
        {
            var id = artifact.GetProperty("artifactId").GetString()!;
            if (id == "receipt" || artifact.GetProperty("availability").GetString() != "available") continue;
            var file = Path.Combine(_root, id + ".bin");
            File.WriteAllBytes(file, Convert.FromBase64String(artifact.GetProperty("bytes").GetString()!));
            artifactPaths[id] = file;
        }
        var verification = J(Dispatcher.Call(_store, "docxodus_verify_receipt", J(JsonSerializer.Serialize(new
        {
            receiptPath,
            artifactPaths,
        }))));
        Assert.True(verification.GetProperty("isValid").GetBoolean(),
            string.Join("; ", verification.GetProperty("findings").EnumerateArray().Select(f => f.GetString())));
    }

    [Fact]
    public void MCP748b_WithoutCapture_TheReceiptIsExplicitlyUnavailable()
    {
        var sessionId = OpenSession(capture: false);
        var anchor = FirstAnchor(sessionId);
        _ = Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "replace_text",
            anchorId = anchor,
            markdown = "Unrecorded.",
        })));

        var bundle = J(Dispatcher.Call(_store, "docxodus_deliver", J(DeliverArgs(sessionId, returnIncomplete: true))));
        Assert.Equal("incomplete", bundle.GetProperty("status").GetString());
        var evidence = bundle.GetProperty("evidence");
        Assert.False(evidence.GetProperty("enabled").GetBoolean());
        Assert.Contains("DeliveryEvidence", evidence.GetProperty("unavailableReason").GetString(), StringComparison.Ordinal);
        var receiptArtifact = bundle.GetProperty("artifacts").EnumerateArray()
            .Single(a => a.GetProperty("artifactId").GetString() == "receipt");
        Assert.Equal("unavailable", receiptArtifact.GetProperty("availability").GetString());
        Assert.False(receiptArtifact.TryGetProperty("bytes", out _));
    }

    [Fact]
    public void MCP748c_ABaselineThatIsNotTheOpenedPackage_CannotBeAttested()
    {
        var other = Path.Combine(_root, "other.docx");
        File.WriteAllBytes(other, DocxSession.CreateBlankDocxBytes());
        var sessionId = OpenSession(capture: true);
        var anchor = FirstAnchor(sessionId);
        _ = Dispatcher.Call(_store, "docxodus_edit", J(JsonSerializer.Serialize(new
        {
            sessionId,
            action = "replace_text",
            anchorId = anchor,
            markdown = "Edited.",
        })));

        var bundle = J(Dispatcher.Call(_store, "docxodus_deliver", J(DeliverArgs(sessionId, baseline: "other.docx", returnIncomplete: true))));
        Assert.Equal("incomplete", bundle.GetProperty("status").GetString());
        Assert.Contains("not the package this session opened",
            bundle.GetProperty("evidence").GetProperty("unavailableReason").GetString(), StringComparison.Ordinal);
    }

    private static string DeliverArgs(
        string sessionId, string baseline = "document.docx", bool returnIncomplete = true, long finalVersion = 1) =>
        JsonSerializer.Serialize(new
        {
            sessionId,
            baselinePath = baseline,
            baselineDocumentVersion = 0,
            finalDocumentName = "final",
            finalDocumentVersion = finalVersion,
            revisionPolicy = new { preExistingRevisions = "preserve", generatedRevisions = "preserve" },
            returnIncompleteBundle = returnIncomplete,
            failOnDeliverableValidationFailure = false,
            changeReceipt = new { privacyProfile = "hashAndSummary" },
            artifacts = new[]
            {
                new { artifactId = "final", kind = "finalDocx", requiredness = "required" },
                new { artifactId = "semantic", kind = "semanticDelta", requiredness = "required" },
                new { artifactId = "receipt", kind = "changeReceipt", requiredness = "required" },
            },
        });

    private string OpenSession(bool capture)
    {
        var opened = J(Dispatcher.Call(_store, "docxodus_open",
            J(JsonSerializer.Serialize(new { path = _path, captureDeliveryEvidence = capture }))));
        return opened.GetProperty("sessionId").GetString()!;
    }

    private string FirstAnchor(string sessionId) =>
        J(Dispatcher.Call(_store, "docxodus_get_content",
            J(JsonSerializer.Serialize(new { sessionId, format = "markdown" }))))
        .GetProperty("anchorIndex").EnumerateObject()
        .First(property => property.Name.StartsWith("p:body:", StringComparison.Ordinal)).Name;

    private static JsonElement J(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

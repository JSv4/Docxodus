// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Security.Cryptography;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Transport-seam coverage for the default deliverable verification operation.</summary>
public sealed class DeliverableVerificationTransportTests
{
    private const string Schema =
        "https://docxodus.dev/schemas/verification/deliverable-verification/v1";

    [Fact]
    public void VT001_InternalStatelessAndSessionOpsReturnCanonicalByteBoundReports()
    {
        var bytes = DocxSessionOps.CreateBlankDocx();
        var expectedDigest = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

        var statelessJson = VerificationOps.VerifyDeliverable(bytes);
        Assert.Equal(statelessJson, VerificationOps.VerifyDeliverable(bytes));
        using (var stateless = JsonDocument.Parse(statelessJson))
        {
            Assert.Equal(Schema, stateless.RootElement.GetProperty("schema").GetString());
            Assert.Equal("standard", stateless.RootElement.GetProperty("mode").GetString());
            Assert.Matches("^[a-z]", stateless.RootElement.GetProperty("decision").GetString()!);
            Assert.False(stateless.RootElement.GetProperty("baselineCompared").GetBoolean());
            Assert.Equal(JsonValueKind.Null,
                stateless.RootElement.GetProperty("baselinePackage").ValueKind);
            Assert.Equal(expectedDigest, stateless.RootElement
                .GetProperty("deliverablePackage")
                .GetProperty("rawPackageBytesDigest")
                .GetProperty("value")
                .GetString());
        }

        var comparedJson = VerificationOps.VerifyDeliverable(bytes, bytes);
        Assert.Equal(comparedJson, VerificationOps.VerifyDeliverable(bytes, bytes));
        using (var compared = JsonDocument.Parse(comparedJson))
        {
            Assert.True(compared.RootElement.GetProperty("baselineCompared").GetBoolean());
            Assert.Equal(expectedDigest, compared.RootElement
                .GetProperty("baselinePackage")
                .GetProperty("rawPackageBytesDigest")
                .GetProperty("value")
                .GetString());
            Assert.Equal(expectedDigest, compared.RootElement
                .GetProperty("deliverablePackage")
                .GetProperty("rawPackageBytesDigest")
                .GetProperty("value")
                .GetString());
        }

        var handle = DocxSessionOps.OpenSession(bytes, settings: null);
        try
        {
            var checkpointDigest = Convert.ToHexString(
                SHA256.HashData(DocxSessionOps.Save(handle))).ToLowerInvariant();
            var version = DocxSessionOps.GetVersion(handle);
            using var session = JsonDocument.Parse(DocxSessionOps.VerifyDeliverable(handle));
            Assert.Equal(Schema, session.RootElement.GetProperty("schema").GetString());
            Assert.True(session.RootElement.GetProperty("baselineCompared").GetBoolean());
            Assert.Equal(expectedDigest, session.RootElement
                .GetProperty("baselinePackage")
                .GetProperty("rawPackageBytesDigest")
                .GetProperty("value")
                .GetString());
            Assert.Equal(checkpointDigest, session.RootElement
                .GetProperty("deliverablePackage")
                .GetProperty("rawPackageBytesDigest")
                .GetProperty("value")
                .GetString());
            Assert.Equal(version, DocxSessionOps.GetVersion(handle));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void VT002_McpVerificationFormatIsAdvertisedDocumentWideAndTyped()
    {
        var root = Path.Combine(Path.GetTempPath(), $"verification-transport-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        var store = new SessionStore(new LocalFileDocumentStore(root));
        try
        {
            var path = Path.Combine(root, "document.docx");
            File.WriteAllBytes(path, DocxSessionOps.CreateBlankDocx());
            using var opened = JsonDocument.Parse(Dispatcher.Call(
                store,
                "docxodus_open",
                Json($$"""{"path":{{JsonSerializer.Serialize(path)}}}""")));
            var sessionId = opened.RootElement.GetProperty("sessionId").GetString()!;
            var sessionArg = JsonSerializer.Serialize(sessionId);

            using var report = JsonDocument.Parse(Dispatcher.Call(
                store,
                "docxodus_get_content",
                Json($$"""{"sessionId":{{sessionArg}},"format":"verification"}""")));
            Assert.Equal(Schema, report.RootElement.GetProperty("schema").GetString());
            Assert.Equal(1, report.RootElement.GetProperty("schemaVersion").GetInt32());
            Assert.True(report.RootElement.GetProperty("baselineCompared").GetBoolean());

            foreach (var anchorValue in new[] { "null", "false", "42", "{}", "[]", "\"p:body:any\"" })
            {
                Assert.Throws<McpToolException>(() => Dispatcher.Call(
                    store,
                    "docxodus_get_content",
                    Json($$"""{"sessionId":{{sessionArg}},"format":"verification","anchorId":{{anchorValue}}}""")));
            }

            var getContent = Assert.Single(
                ToolCatalog.Tools.Where(tool => tool.Name == "docxodus_get_content"));
            using var schema = JsonDocument.Parse(getContent.InputSchemaJson);
            Assert.Contains("verification", schema.RootElement
                .GetProperty("properties")
                .GetProperty("format")
                .GetProperty("enum")
                .EnumerateArray()
                .Select(value => value.GetString()));
        }
        finally
        {
            store.CloseAll();
            Directory.Delete(root, recursive: true);
        }
    }

    private static JsonElement Json(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

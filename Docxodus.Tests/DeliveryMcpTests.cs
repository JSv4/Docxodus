// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

public sealed class DeliveryMcpTests : IDisposable
{
    private readonly string _root;
    private readonly string _baselinePath;
    private readonly SessionStore _store;

    public DeliveryMcpTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"docxodus-delivery-mcp-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _baselinePath = Path.Combine(_root, "baseline.docx");
        File.WriteAllBytes(_baselinePath, DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    [Fact]
    public void Deliver_ReturnsBytesFromSharedServiceThatVerifyIndependently()
    {
        var sessionId = Open();
        var anchorId = FirstBodyAnchor(sessionId);
        var edit = Parse(Dispatcher.Call(_store, "docxodus_edit", Parse(
            $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"replace_text","anchorId":{{JsonSerializer.Serialize(anchorId)}},"markdown":"MCP delivery edit."}""")));
        Assert.True(edit.GetProperty("success").GetBoolean());

        var response = Parse(Dispatcher.Call(_store, "docxodus_deliver", Parse(
            $$"""
            {
              "sessionId": {{JsonSerializer.Serialize(sessionId)}},
              "baselinePath": "baseline.docx",
              "baselineDocumentVersion": 0,
              "finalDocumentName": "final",
              "finalDocumentVersion": 1,
              "revisionPolicy": {
                "preExistingRevisions": "preserve",
                "generatedRevisions": "accept"
              },
              "artifacts": [
                { "artifactId": "final", "kind": "finalDocx", "requiredness": "required" },
                { "artifactId": "semantic", "kind": "semanticDelta", "requiredness": "required" },
                { "artifactId": "package", "kind": "packageDelta", "requiredness": "required" },
                { "artifactId": "validation", "kind": "validationReport", "requiredness": "required" },
                { "artifactId": "html", "kind": "standaloneHtml", "requiredness": "optional", "reviewProfile": "final", "commentProfile": "endnotes" }
              ]
            }
            """)));

        Assert.Equal("complete", response.GetProperty("status").GetString());
        Assert.True(response.GetProperty("verified").GetBoolean());
        Assert.True(response.GetProperty("manifestVerified").GetBoolean());
        Assert.Equal("passed", response.GetProperty("deliverableDecision").GetString());
        var artifacts = response.GetProperty("artifacts").EnumerateArray().ToArray();
        var html = Assert.Single(artifacts, value =>
            value.GetProperty("artifactId").GetString() == "html");
        Assert.Equal("unavailable", html.GetProperty("availability").GetString());
        Assert.Contains("renderer", html.GetProperty("unavailableReason").GetString(),
            StringComparison.OrdinalIgnoreCase);

        var manifestBytes = response.GetProperty("manifestBytes").GetBytesFromBase64();
        var available = artifacts
            .Where(value => value.GetProperty("availability").GetString() == "available")
            .ToDictionary(
                value => value.GetProperty("artifactId").GetString()!,
                value => value.GetProperty("bytes").GetBytesFromBase64(),
                StringComparer.Ordinal);
        var verification = DeliveryBundleVerifier.VerifyJson(manifestBytes, available);
        Assert.True(verification.IsValid,
            string.Join(Environment.NewLine, verification.Findings));
        Assert.True(available["final"].Length > 0);
    }

    [Fact]
    public void Deliver_IsAdvertisedWithExplicitArtifactAndRevisionPolicySchema()
    {
        var tools = Parse(Program.BuildToolsListResult()).GetProperty("tools")
            .EnumerateArray().ToArray();

        var delivery = Assert.Single(tools, tool =>
            tool.GetProperty("name").GetString() == "docxodus_deliver");
        var schema = delivery.GetProperty("inputSchema");
        Assert.Contains("artifacts", schema.GetProperty("required")
            .EnumerateArray().Select(value => value.GetString()));
        Assert.Contains("revisionPolicy", schema.GetProperty("required")
            .EnumerateArray().Select(value => value.GetString()));
        var artifactSchema = schema.GetProperty("properties").GetProperty("artifacts");
        Assert.Equal(DeliveryArtifactRequestRules.MaximumArtifactCount,
            artifactSchema.GetProperty("maxItems").GetInt32());
        Assert.Equal(3, artifactSchema.GetProperty("items").GetProperty("allOf")
            .GetArrayLength());

        var sessionId = Open();
        var error = Assert.Throws<McpToolException>(() => Dispatcher.Call(
            _store,
            "docxodus_deliver",
            Parse($$"""
            {
              "sessionId": {{JsonSerializer.Serialize(sessionId)}},
              "baselinePath": "baseline.docx",
              "baselineDocumentVersion": 0,
              "finalDocumentName": "final",
              "finalDocumentVersion": 0,
              "revisionPolicy": {
                "preExistingRevisions": "preserve",
                "generatedRevisions": "accept"
              },
              "artifacts": []
            }
            """)));
        Assert.Contains("at least one artifact", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Deliver_RejectsInvalidProfilesAndUnknownPropertiesBeforeStoreIo()
    {
        var documents = new CountingDocumentStore();
        using var store = new SessionStoreScope(documents);
        var session = store.Store.Open(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            location: null,
            new DocxSessionSettings());
        var invalid = Parse($$"""
            {
              "sessionId": {{JsonSerializer.Serialize(session.Id)}},
              "baselinePath": "baseline.docx",
              "baselineDocumentVersion": 0,
              "finalDocumentName": "final",
              "finalDocumentVersion": 1,
              "revisionPolicy": {
                "preExistingRevisions": "preserve",
                "generatedRevisions": "accept"
              },
              "artifacts": [
                { "artifactId": "pdf", "kind": "finalPdf", "requiredness": "required", "reviewProfile": "original", "commentProfile": "hidden" }
              ],
              "typo": true
            }
            """);

        var exception = Assert.Throws<McpToolException>(() =>
            DeliveryTool.Execute(store.Store, session, invalid));

        Assert.Contains("unknown docxodus_deliver argument", exception.Message,
            StringComparison.Ordinal);
        Assert.Equal(0, documents.ResolveCount);
        Assert.Equal(0, documents.ReadCount);
    }

    [Fact]
    public void LocalStore_BoundedReadRejectsSparseOversizeBeforeAllocation()
    {
        var path = Path.Combine(_root, "oversize.docx");
        using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write))
            stream.SetLength(DeliveryArtifactRequestRules.MaximumInputPackageBytes + 1);
        var documents = new LocalFileDocumentStore(_root);

        var exception = Assert.Throws<McpToolException>(() => documents.Read(
            documents.Resolve("oversize.docx"),
            DeliveryArtifactRequestRules.MaximumInputPackageBytes));

        Assert.Contains("read limit", exception.Message, StringComparison.Ordinal);
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root))
            Directory.Delete(_root, recursive: true);
    }

    private string Open()
    {
        var result = Parse(Dispatcher.Call(_store, "docxodus_open", Parse(
            """{"path":"baseline.docx"}""")));
        return result.GetProperty("sessionId").GetString()!;
    }

    private string FirstBodyAnchor(string sessionId)
    {
        var result = Parse(Dispatcher.Call(_store, "docxodus_get_content", Parse(
            $$"""{"sessionId":{{JsonSerializer.Serialize(sessionId)}},"format":"blocks"}""")));
        return result.GetProperty("blocks").EnumerateObject().First().Name;
    }

    private static JsonElement Parse(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }

    private sealed class CountingDocumentStore : IDocumentStore
    {
        public string Kind => "counting";
        public string RootDescription => "counting";
        internal int ResolveCount { get; private set; }
        internal int ReadCount { get; private set; }

        public string Resolve(string location)
        {
            ResolveCount++;
            return location;
        }

        public byte[] Read(string resolvedLocation)
        {
            ReadCount++;
            throw new InvalidOperationException("read should not be reached");
        }

        public void Write(string resolvedLocation, byte[] bytes) =>
            throw new InvalidOperationException("write should not be reached");
    }

    private sealed class SessionStoreScope : IDisposable
    {
        internal SessionStoreScope(IDocumentStore documents) => Store = new SessionStore(documents);
        internal SessionStore Store { get; }
        public void Dispose() => Store.CloseAll();
    }
}

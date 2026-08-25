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
}

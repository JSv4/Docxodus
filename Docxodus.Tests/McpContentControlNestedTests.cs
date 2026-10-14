// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>Issue #763 over the MCP server: nested policies, child fills and the per-operation matrix.</summary>
[Collection("MCP session registry isolation")]
public sealed class McpContentControlNestedTests : IDisposable
{
    private readonly string _root;
    private readonly string _path;
    private readonly SessionStore _store;

    public McpContentControlNestedTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-cc-nested-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _path = Path.Combine(_root, "controls.docx");
        File.WriteAllBytes(_path, DocxSessionContentControlTests.BuildFixture());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void MCP763a_NestedPolicyAndChildFills_RideTheToolArguments()
    {
        var sessionId = OpenSession(tracked: false);
        var outer = Control(sessionId, "100");
        var inner = Control(sessionId, "101");
        var outerAnchor = outer.GetProperty("anchorId").GetString()!;
        var innerAnchor = inner.GetProperty("anchorId").GetString()!;
        Assert.Equal(new[] { innerAnchor },
            outer.GetProperty("nestedControlAnchorIds").EnumerateArray().Select(a => a.GetString()));
        var operations = outer.GetProperty("operations").EnumerateArray().ToArray();
        Assert.Contains(operations, op => op.GetProperty("operation").GetString() == "fill_rich_text"
            && op.GetProperty("nestedControls").GetString() == "preserve" && op.GetProperty("canMutate").GetBoolean());
        Assert.Contains(operations, op => op.GetProperty("operation").GetString() == "fill_text"
            && op.GetProperty("nestedControls").GetString() == "refuse" && !op.GetProperty("canMutate").GetBoolean());

        var refused = Call("docxodus_content_controls", new { sessionId, action = "fill_text", anchorId = outerAnchor, text = "x" });
        Assert.Equal("content_control_nested_fill_unsupported", refused.GetProperty("error").GetProperty("code").GetString());

        var preserved = Call("docxodus_content_controls", new
        {
            sessionId,
            action = "fill_text",
            anchorId = outerAnchor,
            text = "Outer via MCP",
            nestedControls = "preserve",
            childFills = new System.Collections.Generic.Dictionary<string, string> { [innerAnchor] = "Inner via MCP" },
        });
        Assert.True(preserved.GetProperty("success").GetBoolean(), preserved.ToString());
        Assert.Equal("Inner via MCP", Control(sessionId, "101").GetProperty("text").GetString());
        Assert.Contains("Outer via MCP", Control(sessionId, "100").GetProperty("text").GetString(), StringComparison.Ordinal);

        // A direct call answers an unknown policy as a typed edit error, like every other bad option.
        var unknownPolicy = Call("docxodus_content_controls", new
        {
            sessionId,
            action = "fill_text",
            anchorId = outerAnchor,
            text = "x",
            nestedControls = "sometimes",
        });
        Assert.Equal("invalid_content_control_value", unknownPolicy.GetProperty("error").GetProperty("code").GetString());
        // A batch step validates the same option before the batch starts.
        var batch = Call("docxodus_mutations", new
        {
            sessionId,
            steps = new[]
            {
                new { tool = "docxodus_content_controls", args = new { action = "fill_text", anchorId = outerAnchor, text = "x", nestedControls = "sometimes" } },
            },
        });
        Assert.False(batch.GetProperty("success").GetBoolean());
        Assert.Equal("invalid_batch_step", batch.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        Assert.Contains("nestedControls", batch.GetProperty("failure").GetProperty("error").GetProperty("message").GetString(), StringComparison.Ordinal);

        var replaced = Call("docxodus_content_controls", new
        {
            sessionId,
            action = "fill_text",
            anchorId = outerAnchor,
            text = "flat",
            nestedControls = "replace",
        });
        Assert.True(replaced.GetProperty("success").GetBoolean(), replaced.ToString());
        Assert.Equal(innerAnchor, Assert.Single(replaced.GetProperty("removed").EnumerateArray()).GetProperty("id").GetString());
    }

    [Fact]
    public void MCP763b_TrackedMode_FillsRecordRevisionsAndStateChangesExplainThemselves()
    {
        var sessionId = OpenSession(tracked: true);
        var inner = Control(sessionId, "101");
        Assert.True(inner.GetProperty("canMutate").GetBoolean());
        var checkbox = Control(sessionId, "102");
        Assert.False(checkbox.GetProperty("canMutate").GetBoolean());
        Assert.Contains("w14:checked", checkbox.GetProperty("unsupportedReason").GetString(), StringComparison.Ordinal);

        var filled = Call("docxodus_content_controls", new
        {
            sessionId,
            action = "fill_text",
            anchorId = inner.GetProperty("anchorId").GetString(),
            text = "tracked via MCP",
        });
        Assert.True(filled.GetProperty("success").GetBoolean(), filled.ToString());
        var revisions = Call("docxodus_track_changes", new { sessionId, action = "list" })
            .GetProperty("revisions").EnumerateArray().Select(r => r.GetProperty("type").GetString()).ToArray();
        Assert.Contains("insert", revisions);
        Assert.Contains("delete", revisions);
        var refused = Call("docxodus_content_controls", new
        {
            sessionId,
            action = "set_checked",
            anchorId = checkbox.GetProperty("anchorId").GetString(),
            @checked = true,
        });
        Assert.Equal("tracked_operation_unsupported", refused.GetProperty("error").GetProperty("code").GetString());
    }

    private JsonElement Call(string tool, object args) =>
        J(Dispatcher.Call(_store, tool, J(JsonSerializer.Serialize(args))));

    private JsonElement Control(string sessionId, string nativeId) =>
        Call("docxodus_content_controls", new { sessionId, action = "list" })
            .GetProperty("contentControls").EnumerateArray()
            .Single(control => control.TryGetProperty("nativeId", out var id) && id.GetString() == nativeId);

    private string OpenSession(bool tracked)
    {
        var opened = J(Dispatcher.Call(_store, "docxodus_open", J(JsonSerializer.Serialize(new
        {
            path = _path,
            trackedChanges = tracked ? "render_inline" : "accept",
        }))));
        return opened.GetProperty("sessionId").GetString()!;
    }

    private static JsonElement J(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

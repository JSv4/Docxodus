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

/// <summary>Issue #760 over the MCP server: retain a preview, then commit it exactly as previewed.</summary>
[Collection("MCP session registry isolation")]
public sealed class McpPreviewCommitTests : IDisposable
{
    private readonly string _root;
    private readonly string _path;
    private readonly SessionStore _store;

    public McpPreviewCommitTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"mcp-preview-commit-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _path = Path.Combine(_root, "document.docx");
        File.WriteAllBytes(_path, DocxSession.CreateBlankDocxBytes());
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    [Fact]
    public void MCP760a_RetainedPreview_CommitsExactlyAsPreviewedAndRefusesOnceStale()
    {
        var sessionId = OpenSession();
        var anchor = FirstAnchor(sessionId);

        var preview = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "preview",
            retainPreview = true,
            steps = new[] { InsertStep(anchor, "Committed exactly as previewed.") },
        }))));
        Assert.True(preview.GetProperty("success").GetBoolean());
        Assert.True(preview.GetProperty("preview").GetBoolean());
        var retention = preview.GetProperty("retention");
        var previewId = retention.GetProperty("previewId").GetString()!;
        Assert.Equal(0, retention.GetProperty("baseVersion").GetInt64());
        var previewedHash = preview.GetProperty("packageHash").GetString()!;
        var createdId = preview.GetProperty("steps")[0].GetProperty("results")[0]
            .GetProperty("created")[0].GetProperty("id").GetString()!;
        Assert.DoesNotContain("Committed exactly as previewed.", GetMarkdown(sessionId), StringComparison.Ordinal);

        var commit = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            commitPreviewId = previewId,
        }))));
        Assert.True(commit.GetProperty("success").GetBoolean());
        Assert.False(commit.GetProperty("preview").GetBoolean());
        Assert.Equal(previewedHash, commit.GetProperty("packageHash").GetString());
        Assert.Equal(1, commit.GetProperty("resultVersion").GetInt64());
        Assert.Equal(previewId, commit.GetProperty("retention").GetProperty("previewId").GetString());
        Assert.Equal(createdId, commit.GetProperty("steps")[0].GetProperty("results")[0]
            .GetProperty("created")[0].GetProperty("id").GetString());
        Assert.DoesNotContain(commit.GetProperty("warnings").EnumerateArray(),
            warning => warning.GetString()!.Contains("may be generated", StringComparison.Ordinal));
        var markdown = GetMarkdown(sessionId);
        Assert.Contains("Committed exactly as previewed.", markdown, StringComparison.Ordinal);
        Assert.Contains(createdId, markdown, StringComparison.Ordinal);

        // A second preview, then an unrelated live edit: the commit is refused without editing.
        var second = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "atomic",
            preview = true,
            retainPreview = true,
            steps = new[] { InsertStep(anchor, "Never lands.") },
        }))));
        var secondId = second.GetProperty("retention").GetProperty("previewId").GetString()!;
        _ = Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            steps = new[] { InsertStep(anchor, "Intervening.") },
        })));
        var stale = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            commitPreviewId = secondId,
        }))));
        Assert.False(stale.GetProperty("success").GetBoolean());
        Assert.Equal("preview_stale",
            stale.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());
        Assert.DoesNotContain("Never lands.", GetMarkdown(sessionId), StringComparison.Ordinal);
    }

    [Fact]
    public void MCP760b_CommitUnderATransactionId_ReplaysAndArgumentsAreValidated()
    {
        var sessionId = OpenSession();
        var anchor = FirstAnchor(sessionId);
        var preview = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            mode = "preview",
            retainPreview = true,
            steps = new[] { InsertStep(anchor, "Once, even when retried.") },
        }))));
        var previewId = preview.GetProperty("retention").GetProperty("previewId").GetString()!;

        var first = Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = "tx-commit",
            commitPreviewId = previewId,
        })));
        var retry = Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            transactionId = "tx-commit",
            commitPreviewId = previewId,
        })));
        Assert.Equal(first, retry);
        var committed = J(first);
        Assert.True(committed.GetProperty("success").GetBoolean());
        Assert.Equal("tx-commit", committed.GetProperty("transaction").GetProperty("transactionId").GetString());
        Assert.Equal(1, GetMarkdown(sessionId).Split("Once, even when retried.").Length - 1);

        // Consumed: without a transaction the id is gone.
        var gone = J(Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
        {
            sessionId,
            commitPreviewId = previewId,
        }))));
        Assert.Equal("preview_not_found",
            gone.GetProperty("failure").GetProperty("error").GetProperty("code").GetString());

        // Argument shape: retention is a preview-only option, and a commit takes no steps.
        Assert.Contains("retainPreview", Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
            {
                sessionId,
                retainPreview = true,
                steps = new[] { InsertStep(anchor, "x") },
            })))).Message, StringComparison.Ordinal);
        Assert.Contains("steps", Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
            {
                sessionId,
                commitPreviewId = "pv-x",
                steps = new[] { InsertStep(anchor, "x") },
            })))).Message, StringComparison.Ordinal);
        Assert.Contains("preview", Assert.Throws<McpToolException>(() =>
            Dispatcher.Call(_store, "docxodus_mutations", J(JsonSerializer.Serialize(new
            {
                sessionId,
                mode = "preview",
                commitPreviewId = "pv-x",
            })))).Message, StringComparison.Ordinal);
    }

    private static object InsertStep(string anchor, string markdown) => new
    {
        tool = "docxodus_create",
        args = new { action = "insert_paragraph", anchorId = anchor, position = "after", markdown },
    };

    private string OpenSession()
    {
        var opened = J(Dispatcher.Call(_store, "docxodus_open", J(JsonSerializer.Serialize(new { path = _path }))));
        return opened.GetProperty("sessionId").GetString()!;
    }

    private string FirstAnchor(string sessionId) =>
        J(Dispatcher.Call(_store, "docxodus_get_content",
            J(JsonSerializer.Serialize(new { sessionId, format = "markdown" }))))
        .GetProperty("anchorIndex").EnumerateObject().First().Name;

    private string GetMarkdown(string sessionId) =>
        J(Dispatcher.Call(_store, "docxodus_get_content",
            J(JsonSerializer.Serialize(new { sessionId, format = "markdown" }))))
        .GetProperty("markdown").GetString()!;

    private static JsonElement J(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }
}

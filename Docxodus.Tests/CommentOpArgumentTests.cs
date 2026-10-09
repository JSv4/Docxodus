// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The native-comment ops take one name per argument and one default per op on every transport
/// (issue #1014). The comment is <c>commentAnchorId</c>; the stdio host still accepts its old
/// spellings (<c>parentAnchorId</c> for a reply, <c>anchorId</c> for update, resolve and remove) as
/// deprecated aliases. <see cref="DocxSessionOps"/> owns the defaults: an omitted <c>markdown</c> on
/// add or reply is an empty comment, and an omitted <c>resolved</c> resolves.
/// </summary>
public sealed class CommentOpArgumentTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"comment-args-{Guid.NewGuid():N}");
    private readonly SessionStore _store;

    public CommentOpArgumentTests()
    {
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        Directory.Delete(_root, recursive: true);
    }

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    private static void Seed(int handle, out string paragraph, out string comment)
    {
        var session = SessionRegistry.Get(handle);
        paragraph = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        paragraph = session.ReplaceText(paragraph, "Hello comment world").Modified.Select(a => a.Id).FirstOrDefault() ?? paragraph;
        comment = session.AddComment(paragraph, null, "Seed", "Seed comment").Created.First(a => a.Kind == "cmt").Id;
    }

    /// <summary>Run <paramref name="call"/> against a fresh stdio-host session; return its result and comments.</summary>
    private static (JsonElement Result, CommentListEntry[] Comments) Stdio(
        string op, Func<int, string, string, Dictionary<string, object>> args)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        try
        {
            Seed(handle, out var paragraph, out var comment);
            var a = args(handle, paragraph, comment);
            a["handle"] = handle;
            var result = J(JsonDocument.Parse(Docxodus.PyHost.Dispatcher.Dispatch(op, J(a))).RootElement);
            return (result, SessionRegistry.Get(handle).ListComments().ToArray());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    private (JsonElement Result, CommentListEntry[] Comments) Mcp(
        string action, Func<string, string, Dictionary<string, object>> args)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(JsonDocument.Parse(Dispatcher.Call(_store, "docxodus_open", J(new { path }))).RootElement)
            .GetProperty("sessionId").GetString()!;
        var handle = _store.Get(sessionId).Handle;
        Seed(handle, out var paragraph, out var comment);
        var a = args(paragraph, comment);
        a["sessionId"] = sessionId;
        a["action"] = action;
        var result = J(JsonDocument.Parse(Dispatcher.Call(_store, "docxodus_comment", J(a))).RootElement);
        return (result, SessionRegistry.Get(handle).ListComments().ToArray());
    }

    private static void AssertSucceeded(JsonElement result) =>
        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());

    [Theory]
    [InlineData("update_comment")]
    [InlineData("set_comment_resolved")]
    [InlineData("remove_comment")]
    [InlineData("add_comment_reply")]
    public void StdioHost_TakesTheCommentAsCommentAnchorId(string op)
    {
        var (result, _) = Stdio(op, (_, _, comment) => new()
        {
            ["commentAnchorId"] = comment, ["markdown"] = "Text", ["resolved"] = true, ["author"] = "Reviewer",
        });
        AssertSucceeded(result);
    }

    [Theory]
    [InlineData("update_comment", "anchorId")]
    [InlineData("set_comment_resolved", "anchorId")]
    [InlineData("remove_comment", "anchorId")]
    [InlineData("add_comment_reply", "parentAnchorId")]
    public void StdioHost_StillAcceptsTheDeprecatedSpelling(string op, string alias)
    {
        var (result, _) = Stdio(op, (_, _, comment) => new()
        {
            [alias] = comment, ["markdown"] = "Text", ["resolved"] = true, ["author"] = "Reviewer",
        });
        AssertSucceeded(result);
    }

    [Fact]
    public void StdioHost_RefusesTheTwoSpellingsDisagreeing() =>
        Assert.Throws<ArgumentException>(() => Stdio("remove_comment", (_, paragraph, comment) => new()
        {
            ["commentAnchorId"] = comment, ["anchorId"] = paragraph,
        }));

    [Fact]
    public void ResolveWithoutResolved_ResolvesOnBothTransports()
    {
        var (stdio, stdioComments) = Stdio("set_comment_resolved", (_, _, comment) => new() { ["commentAnchorId"] = comment });
        var (mcp, mcpComments) = Mcp("resolve", (_, comment) => new() { ["commentAnchorId"] = comment });

        AssertSucceeded(stdio);
        AssertSucceeded(mcp);
        Assert.True(Assert.Single(stdioComments).Resolved);
        Assert.True(Assert.Single(mcpComments).Resolved);
    }

    [Fact]
    public void AddAndReplyWithoutMarkdown_AddAnEmptyCommentOnBothTransports()
    {
        var (stdioAdd, stdioAdded) = Stdio("add_comment", (_, paragraph, _) => new() { ["anchorId"] = paragraph, ["author"] = "Reviewer" });
        var (mcpAdd, mcpAdded) = Mcp("add", (paragraph, _) => new() { ["anchorId"] = paragraph, ["author"] = "Reviewer" });
        var (stdioReply, stdioReplied) = Stdio("add_comment_reply", (_, _, comment) => new() { ["commentAnchorId"] = comment, ["author"] = "Reviewer" });
        var (mcpReply, mcpReplied) = Mcp("reply", (_, comment) => new() { ["commentAnchorId"] = comment, ["author"] = "Reviewer" });

        foreach (var result in new[] { stdioAdd, mcpAdd, stdioReply, mcpReply })
            AssertSucceeded(result);
        foreach (var comments in new[] { stdioAdded, mcpAdded, stdioReplied, mcpReplied })
            Assert.Equal("", comments.Single(c => c.Author == "Reviewer").Text);
    }
}

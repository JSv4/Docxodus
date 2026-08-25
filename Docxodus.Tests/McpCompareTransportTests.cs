// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Transport-seam coverage for the sessionless <c>docxodus_compare</c> tool. The comparison
/// engines are covered by their own suites; these tests assert only the seam: store-scoped path
/// resolution on every input and the output, honest revision attribution in the summary and the
/// written redline, the two-form contract, and that the tool cannot be smuggled into a mutation
/// batch.
/// </summary>
public sealed class McpCompareTransportTests
{
    [Fact]
    public void CT001_TwoWayCompareWritesAnAttributedRedlineAndSummarizesIt()
    {
        using var workspace = new StoreWorkspace();
        workspace.Write("baseline.docx", Document("The original clause."));
        workspace.Write("revised.docx", Document("The revised clause."));

        var response = workspace.Call("docxodus_compare", $$"""
            {"baselinePath":"baseline.docx","revisedPath":"revised.docx",
             "author":"Reviewer A","outputPath":"redline.docx"}
            """);

        using var summary = JsonDocument.Parse(response);
        var total = summary.RootElement.GetProperty("revisions").GetProperty("total").GetInt32();
        Assert.True(total > 0, response);
        Assert.Equal(
            total,
            summary.RootElement.GetProperty("revisions").GetProperty("byAuthor")
                .GetProperty("Reviewer A").GetInt32());

        // The written file, opened like any other document, carries the same live markup the
        // summary claimed — the summary cannot drift from the artifact.
        var sessionId = workspace.OpenSession("redline.docx");
        using var listed = JsonDocument.Parse(workspace.Call("docxodus_track_changes", $$"""
            {"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"list"}
            """));
        var revisions = listed.RootElement.GetProperty("revisions").EnumerateArray().ToList();
        Assert.Equal(total, revisions.Count);
        Assert.All(revisions, revision =>
            Assert.Equal("Reviewer A", revision.GetProperty("author").GetString()));
    }

    [Fact]
    public void CT002_ConsolidateAttributesEachReviewersChangesToTheirAuthor()
    {
        using var workspace = new StoreWorkspace();
        workspace.Write("base.docx", Document("Alpha stays. Beta stays."));
        workspace.Write("a.docx", Document("Alpha revised. Beta stays."));
        workspace.Write("b.docx", Document("Alpha stays. Beta revised."));

        var response = workspace.Call("docxodus_compare", """
            {"baselinePath":"base.docx","revisedPaths":["a.docx","b.docx"],
             "authors":["Reviewer A","Reviewer B"],"outputPath":"consolidated.docx"}
            """);

        using var summary = JsonDocument.Parse(response);
        var byAuthor = summary.RootElement.GetProperty("revisions").GetProperty("byAuthor");
        Assert.True(byAuthor.GetProperty("Reviewer A").GetInt32() > 0, response);
        Assert.True(byAuthor.GetProperty("Reviewer B").GetInt32() > 0, response);

        var sessionId = workspace.OpenSession("consolidated.docx");
        using var listed = JsonDocument.Parse(workspace.Call("docxodus_track_changes", $$"""
            {"sessionId":{{JsonSerializer.Serialize(sessionId)}},"action":"list"}
            """));
        var authors = listed.RootElement.GetProperty("revisions").EnumerateArray()
            .Select(revision => revision.GetProperty("author").GetString())
            .ToHashSet();
        Assert.Contains("Reviewer A", authors);
        Assert.Contains("Reviewer B", authors);
    }

    [Fact]
    public void CT003_EveryPathIsResolvedInsideTheDocumentScope()
    {
        using var workspace = new StoreWorkspace();
        workspace.Write("baseline.docx", Document("Text."));
        workspace.Write("revised.docx", Document("Changed text."));
        var outside = Path.Combine(
            Path.GetTempPath(), $"compare-outside-{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(outside, Document("Text."));
        try
        {
            foreach (var request in new[]
                     {
                         $$"""
                           {"baselinePath":{{JsonSerializer.Serialize(outside)}},
                            "revisedPath":"revised.docx","outputPath":"out.docx"}
                           """,
                         $$"""
                           {"baselinePath":"baseline.docx",
                            "revisedPath":{{JsonSerializer.Serialize(outside)}},"outputPath":"out.docx"}
                           """,
                         $$"""
                           {"baselinePath":"baseline.docx","revisedPath":"revised.docx",
                            "outputPath":{{JsonSerializer.Serialize(outside)}}}
                           """,
                     })
            {
                var escape = Assert.Throws<McpToolException>(
                    () => workspace.Call("docxodus_compare", request));
                Assert.Contains("outside this server's document scope", escape.Message);
            }
        }
        finally
        {
            File.Delete(outside);
        }
    }

    [Fact]
    public void CT004_TheTwoFormsAreExclusiveAndAuthorsMustMatch()
    {
        using var workspace = new StoreWorkspace();
        workspace.Write("baseline.docx", Document("Text."));
        workspace.Write("revised.docx", Document("Changed text."));

        var both = Assert.Throws<McpToolException>(() => workspace.Call("docxodus_compare", """
            {"baselinePath":"baseline.docx","revisedPath":"revised.docx",
             "revisedPaths":["revised.docx","revised.docx"],"outputPath":"out.docx"}
            """));
        Assert.Contains("exactly one of", both.Message);

        var neither = Assert.Throws<McpToolException>(() => workspace.Call("docxodus_compare", """
            {"baselinePath":"baseline.docx","outputPath":"out.docx"}
            """));
        Assert.Contains("exactly one of", neither.Message);

        var mismatched = Assert.Throws<McpToolException>(() => workspace.Call("docxodus_compare", """
            {"baselinePath":"baseline.docx","revisedPaths":["revised.docx","revised.docx"],
             "authors":["Only One"],"outputPath":"out.docx"}
            """));
        Assert.Contains("authors must match revisedPaths", mismatched.Message);
    }

    [Fact]
    public void CT005_CompareIsNotAcceptedAsAMutationBatchStep()
    {
        using var workspace = new StoreWorkspace();
        workspace.Write("document.docx", Document("Text."));
        var sessionId = workspace.OpenSession("document.docx");

        // Behavioral, not schema text: the batch validator's own allowlist must refuse the tool —
        // nothing executed, the batch reports the step as invalid.
        var response = workspace.Call("docxodus_mutations", $$$"""
            {"sessionId":{{{JsonSerializer.Serialize(sessionId)}}},"steps":[
                {"tool":"docxodus_compare","args":{"action":"compare",
                 "baselinePath":"document.docx","revisedPath":"document.docx",
                 "outputPath":"out.docx"}}]}
            """);
        using var batch = JsonDocument.Parse(response);
        Assert.False(batch.RootElement.GetProperty("success").GetBoolean());
        Assert.Contains("unsupported or read-only batch action: docxodus_compare", response);
    }

    /// <summary>A disposable store rooted in a private directory, mirroring the MCP server.</summary>
    private sealed class StoreWorkspace : IDisposable
    {
        private readonly string _root =
            Path.Combine(Path.GetTempPath(), $"docxodus-compare-{Guid.NewGuid():N}");

        private readonly SessionStore _store;

        public StoreWorkspace()
        {
            Directory.CreateDirectory(_root);
            _store = new SessionStore(new LocalFileDocumentStore(_root));
        }

        public void Write(string name, byte[] bytes) =>
            File.WriteAllBytes(Path.Combine(_root, name), bytes);

        public string Call(string tool, string argsJson)
        {
            using var document = JsonDocument.Parse(argsJson);
            return Dispatcher.Call(_store, tool, document.RootElement.Clone());
        }

        public string OpenSession(string name)
        {
            using var opened = JsonDocument.Parse(Call(
                "docxodus_open",
                $$"""{"path":{{JsonSerializer.Serialize(name)}}}"""));
            return opened.RootElement.GetProperty("sessionId").GetString()!;
        }

        public void Dispose()
        {
            _store.CloseAll();
            Directory.Delete(_root, recursive: true);
        }
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
}

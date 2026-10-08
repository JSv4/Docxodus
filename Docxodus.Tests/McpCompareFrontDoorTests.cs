// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using Docxodus.McpServer;
using Xunit;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;

namespace Docxodus.Tests;

/// <summary>
/// The MCP <c>docxodus_compare</c> tool is a front door: it compares the accepted view of each input, as
/// <see cref="DocxCompare.Compare"/> and Word's Compare do (issue #1006). It used to run the raw engine
/// defaults, so an input that already carried tracked changes produced a different redline from the agent
/// server than from the CLI or the browser.
/// </summary>
public sealed class McpCompareFrontDoorTests
{
    // Both versions carry another reviewer's pending changes. The body ones ("BodyPrior") differ between the
    // versions; the header and footnote ones ("ScopePrior") are identical, so those parts are carried over
    // unchanged. The raw engine diffs the accepted body but copies carried-over parts verbatim, so
    // "ScopePrior"'s markup leaks into its redline; the front door accepts both inputs first.
    private static WmlDocument Baseline() => MultiScopeRevisionDoc(
        "Body alpha", "BodyPrior", "ins-a", "del-a", "ScopePrior", "hdr-prior", "fn-prior");

    private static WmlDocument Revised(string body = "Body gamma") => MultiScopeRevisionDoc(
        body, "BodyPrior", "ins-g", "del-g", "ScopePrior", "hdr-prior", "fn-prior");

    [Fact]
    public void TwoWayCompare_OfRevisionBearingInputs_MatchesTheDocxCompareFrontDoor()
    {
        using var workspace = new Workspace();
        workspace.Write("baseline.docx", Baseline().DocumentByteArray);
        workspace.Write("revised.docx", Revised().DocumentByteArray);
        var frontDoor = DocxCompare.Compare(Baseline(), Revised(), new DocxDiffSettings { AuthorForRevisions = "Agent" });

        workspace.Call("docxodus_compare",
            """{"baselinePath":"baseline.docx","revisedPath":"revised.docx","author":"Agent","outputPath":"mcp.docx"}""");

        var mcp = workspace.Read("mcp.docx");
        Assert.Equal(new[] { "Agent" }, RevisionAuthorsAllScopes(frontDoor).OrderBy(a => a));
        Assert.Equal(RevisionAuthorsAllScopes(frontDoor).OrderBy(a => a), RevisionAuthorsAllScopes(mcp).OrderBy(a => a));
        Assert.Equal(AllScopesText(frontDoor), AllScopesText(mcp));
    }

    [Fact]
    public void FanOutCompare_OfRevisionBearingInputs_MatchesTheDocxCompareFrontDoor()
    {
        using var workspace = new Workspace();
        workspace.Write("baseline.docx", Baseline().DocumentByteArray);
        workspace.Write("gamma.docx", Revised().DocumentByteArray);
        workspace.Write("delta.docx", Revised("Body delta").DocumentByteArray);

        workspace.Call("docxodus_compare",
            """{"baselinePath":"baseline.docx","mode":"fan_out","revisedPaths":["gamma.docx","delta.docx"],"outputPaths":["mcp-gamma.docx","mcp-delta.docx"]}""");

        foreach (var (revised, output) in new[] { (Revised(), "mcp-gamma.docx"), (Revised("Body delta"), "mcp-delta.docx") })
        {
            var frontDoor = DocxCompare.Compare(Baseline(), revised);
            var mcp = workspace.Read(output);
            Assert.DoesNotContain("ScopePrior", RevisionAuthorsAllScopes(mcp));
            Assert.Equal(RevisionAuthorsAllScopes(frontDoor).OrderBy(a => a), RevisionAuthorsAllScopes(mcp).OrderBy(a => a));
            Assert.Equal(AllScopesText(frontDoor), AllScopesText(mcp));
        }
    }

    /// <summary>Identical packages take the front door's exact no-op: the stored redline is the input
    /// itself, byte for byte, as <see cref="DocxCompare.Compare"/> returns it.</summary>
    [Fact]
    public void TwoWayCompare_OfIdenticalPackages_WritesTheFrontDoorsExactClone()
    {
        using var workspace = new Workspace();
        var bytes = Baseline().DocumentByteArray;
        workspace.Write("baseline.docx", bytes);
        workspace.Write("same.docx", bytes);

        var response = workspace.Call("docxodus_compare",
            """{"baselinePath":"baseline.docx","revisedPath":"same.docx","outputPath":"mcp.docx"}""");

        Assert.Equal(bytes, workspace.Read("mcp.docx").DocumentByteArray);
        using var summary = JsonDocument.Parse(response);
        Assert.Equal(0, summary.RootElement.GetProperty("revisions").GetProperty("total").GetInt32());
    }

    /// <summary>A disposable MCP document store in a private directory.</summary>
    private sealed class Workspace : IDisposable
    {
        private readonly string _root = Path.Combine(Path.GetTempPath(), $"docxodus-front-door-{Guid.NewGuid():N}");
        private readonly SessionStore _store;

        public Workspace()
        {
            Directory.CreateDirectory(_root);
            _store = new SessionStore(new LocalFileDocumentStore(_root));
        }

        public void Write(string name, byte[] bytes) => File.WriteAllBytes(Path.Combine(_root, name), bytes);

        public string Call(string tool, string argsJson)
        {
            using var document = JsonDocument.Parse(argsJson);
            return Dispatcher.Call(_store, tool, document.RootElement.Clone());
        }

        public WmlDocument Read(string name) => new(name, File.ReadAllBytes(Path.Combine(_root, name)));

        public void Dispose()
        {
            _store.CloseAll();
            Directory.Delete(_root, recursive: true);
        }
    }
}

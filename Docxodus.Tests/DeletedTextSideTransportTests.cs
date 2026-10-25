using System.Text.Json;
using System.Xml.Linq;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The optional <c>deletedTextSide</c> of the span ops (issue #1290) reaches the session through the
/// stdio host and the MCP server: each test calls the transport for real on <c>The tenant shall
/// [deleted: promptly ]repaint</c> and checks where the insertion landed.
/// </summary>
public sealed class DeletedTextSideTransportTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"deleted-side-{Guid.NewGuid():N}");
    private readonly SessionStore _store;

    public DeletedTextSideTransportTests()
    {
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        Directory.Delete(_root, recursive: true);
    }

    private static byte[] Fixture() => DocxSessionSpanRevisionTests.Build(DocxSessionSpanRevisionTests.TenantPromptlyRepaint());

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(value as string ?? JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    /// <summary>Whether the paragraph's inserted "Y" follows its deletion.</summary>
    private static bool InsertionFollowsDeletion(DocxSession session)
    {
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
        var paragraph = XElement.Parse(session.Raw.GetXml(anchor));
        var deletion = paragraph.Elements().ToList().FindIndex(e => e.Name == W.del);
        var insertion = paragraph.Elements().ToList().FindIndex(e => e.Name == W.ins);
        Assert.True(insertion >= 0 && deletion >= 0, paragraph.ToString());
        return insertion > deletion;
    }

    [Theory]
    [InlineData("replace_text_at_span", null, false)]
    [InlineData("replace_text_at_span", "after", true)]
    [InlineData("replace_text_at_span_with_format", "after", true)]
    [InlineData("replace_text_at_span_with_format", "before", false)]
    public void Stdio_SpanInsertion_GoesOnTheRequestedSideOfTheDeletion(string op, string? side, bool after)
    {
        var handle = DocxSessionOps.OpenSession(Fixture(), new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
            var args = new Dictionary<string, object?>
            {
                ["handle"] = handle, ["anchorId"] = anchor, ["spanStart"] = 17, ["spanLength"] = 0, ["replace"] = "Y",
            };
            if (side is not null) args["deletedTextSide"] = side;
            var result = J(Docxodus.PyHost.Dispatcher.Dispatch(op, J(args)));
            Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
            Assert.Equal(after, InsertionFollowsDeletion(session));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("after", true)]
    public void Mcp_SpanInsertion_GoesOnTheRequestedSideOfTheDeletion(string? side, bool after)
    {
        var path = Path.Combine(_root, "tenant.docx");
        File.WriteAllBytes(path, Fixture());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path, trackedChanges = "render_inline" })))
            .GetProperty("sessionId").GetString()!;
        var session = SessionRegistry.Get(_store.Get(sessionId).Handle);
        var anchor = session.Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
        var args = new Dictionary<string, object?>
        {
            ["sessionId"] = sessionId, ["action"] = "replace_text_at_span_with_format",
            ["anchorId"] = anchor, ["spanStart"] = 17, ["spanLength"] = 0, ["replace"] = "Y",
        };
        if (side is not null) args["deletedTextSide"] = side;

        var result = J(Dispatcher.Call(_store, "docxodus_edit", J(args)));

        Assert.True(result.GetProperty("success").GetBoolean(), result.GetRawText());
        Assert.Equal(after, InsertionFollowsDeletion(session));
    }

    [Fact]
    public void Stdio_AnUnknownSide_IsRefused()
    {
        var handle = DocxSessionOps.OpenSession(Fixture(), null);
        try
        {
            var anchor = SessionRegistry.Get(handle).Project().AnchorIndex.Values.Single(a => a.Anchor.Kind == "p").Anchor.Id;
            Assert.ThrowsAny<Exception>(() => Docxodus.PyHost.Dispatcher.Dispatch("replace_text_at_span", J(new
            {
                handle, anchorId = anchor, spanStart = 17, spanLength = 0, replace = "Y", deletedTextSide = "sideways",
            })));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }
}

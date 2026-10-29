using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

public sealed class DocxSessionAnchorLocationTransportTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "anchor-locations-" + Guid.NewGuid().ToString("N"));
    private readonly SessionStore _store;
    private readonly int _handle;
    private readonly string _sessionId;

    public DocxSessionAnchorLocationTransportTests()
    {
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
        var bytes = DocxSession.CreateBlankDocxBytes();
        _handle = DocxSessionOps.OpenSession(bytes, new DocxSessionSettings { CaptureInitialProjection = false });
        var path = Path.Combine(_root, "source.docx");
        File.WriteAllBytes(path, bytes);
        _sessionId = Parse(Dispatcher.Call(_store, "docxodus_open", Json(new { path, captureInitialProjection = false })))
            .GetProperty("sessionId").GetString()!;
    }

    public void Dispose()
    {
        DocxSessionOps.CloseSession(_handle);
        _store.CloseAll();
        Directory.Delete(_root, true);
    }

    private static JsonElement Json(object value) => Parse(JsonSerializer.Serialize(value));
    private static JsonElement Parse(string json)
    {
        using var document = JsonDocument.Parse(json);
        return document.RootElement.Clone();
    }

    private JsonElement Stdio(Dictionary<string, object?> args)
    {
        args["handle"] = _handle;
        return Parse(Docxodus.PyHost.Dispatcher.Dispatch("get_anchor_locations", Json(args)));
    }

    private JsonElement Mcp(Dictionary<string, object?> args)
    {
        args["sessionId"] = _sessionId;
        args["format"] = "anchor_locations";
        return Parse(Dispatcher.Call(_store, "docxodus_get_content", Json(args)));
    }

    [Fact]
    public void Facade_stdio_and_MCP_return_the_same_read_only_location_envelope()
    {
        var expected = Parse(DocxSessionOps.GetAnchorLocations(_handle));
        Assert.Equal(expected.GetRawText(), Stdio(new()).GetRawText());
        Assert.Equal(expected.GetRawText(), Mcp(new()).GetRawText());
        Assert.Equal(0, expected.GetProperty("version").GetInt64());
        Assert.Contains(expected.GetProperty("locations").EnumerateObject(), property =>
            property.Value.GetProperty("partUri").GetString() == "/word/document.xml"
            && property.Value.GetProperty("elementPath").GetString() == "/*[1]/*[1]/*[1]");
        Assert.False(SessionRegistry.Get(_handle).Undo());
        Assert.Equal(0, SessionRegistry.Get(_store.Get(_sessionId).Handle).Version);
    }

    [Fact]
    public void Null_all_empty_none_and_requested_unknown_keys_are_preserved_on_both_transports()
    {
        var all = Stdio(new()).GetProperty("locations");
        var key = all.EnumerateObject().First().Name;
        foreach (var call in new Func<Dictionary<string, object?>, JsonElement>[] { Stdio, Mcp })
        {
            Assert.Equal(all.GetRawText(), call(new() { ["anchorIds"] = null }).GetProperty("locations").GetRawText());
            Assert.Empty(call(new() { ["anchorIds"] = Array.Empty<string>() }).GetProperty("locations").EnumerateObject());
            const string unknown = "not \" an anchor";
            var subset = call(new() { ["anchorIds"] = new[] { key, unknown, key } }).GetProperty("locations");
            Assert.Equal(new[] { key, unknown }, subset.EnumerateObject().Select(p => p.Name));
            Assert.Equal(JsonValueKind.Null, subset.GetProperty(unknown).ValueKind);
            Assert.Equal(all.GetProperty(key).GetRawText(), subset.GetProperty(key).GetRawText());
        }
    }

    [Theory]
    [InlineData("42")]
    [InlineData("\"p:body:unknown\"")]
    [InlineData("[null]")]
    [InlineData("[true]")]
    [InlineData("[\"p:body:unknown\",42]")]
    public void Invalid_anchor_arrays_are_refused_consistently(string value)
    {
        var invalid = Parse(value);
        foreach (var call in new Func<Dictionary<string, object?>, JsonElement>[] { Stdio, Mcp })
        {
            var error = Assert.ThrowsAny<Exception>(() => call(new() { ["anchorIds"] = invalid }));
            Assert.Contains("anchorIds", error.Message);
        }
        Assert.Equal(0, SessionRegistry.Get(_handle).Version);
    }

    [Fact]
    public void MCP_refuses_a_singular_anchor_argument_instead_of_silently_reading_every_anchor()
    {
        var error = Assert.Throws<McpToolException>(() => Mcp(new() { ["anchorId"] = "p:body:unknown" }));
        Assert.Contains("anchorIds", error.Message);
    }
}

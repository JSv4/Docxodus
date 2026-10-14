// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Reflection;
using System.Text.Json;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Some session ops reach MCP through a selector on a read tool (<c>docxodus_get_content</c>'s
/// <c>format</c>, say) rather than through an <c>action</c> (issue #1026). An op description
/// records such a route beside its <c>absent</c> reason as
/// <c>"route": { "tool": …, "selector": { "&lt;property&gt;": "&lt;value&gt;" } }</c>. These tests
/// check every recorded route the way the drift tests check an action: the tool's schema
/// advertises the selector value, and calling the tool with it returns what the facade returns.
/// Only one-to-one routes are recorded: the selector takes the op's arguments unchanged and
/// returns the facade's result unwrapped.
/// </summary>
public sealed class SessionOpRouteTests : IDisposable
{
    private const string RepoRoot = "../../../..";

    private readonly string _root = Path.Combine(Path.GetTempPath(), $"op-routes-{Guid.NewGuid():N}");
    private readonly SessionStore _store;

    public SessionOpRouteTests()
    {
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        Directory.Delete(_root, recursive: true);
    }

    private sealed record Route(string Family, string Op, string Facade, int ArgCount, string? Absent,
        string Tool, IReadOnlyDictionary<string, string> Selector);

    private static IReadOnlyList<Route> Routes()
    {
        var routes = new List<Route>();
        foreach (var file in Directory.GetFiles(Path.Combine(RepoRoot, "tools", "op-descriptions"), "*.json"))
        {
            using var doc = JsonDocument.Parse(File.ReadAllText(file));
            if (!doc.RootElement.TryGetProperty("ops", out var ops)) continue;
            var family = doc.RootElement.GetProperty("family").GetString()!;
            foreach (var op in ops.EnumerateArray())
            {
                var mcp = op.GetProperty("transports").GetProperty("mcp");
                if (!mcp.TryGetProperty("route", out var route)) continue;
                routes.Add(new Route(
                    family,
                    op.GetProperty("op").GetString()!,
                    op.GetProperty("facade").GetString()!,
                    op.GetProperty("args").GetArrayLength(),
                    mcp.TryGetProperty("absent", out var absent) ? absent.GetString() : null,
                    route.GetProperty("tool").GetString()!,
                    route.GetProperty("selector").EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!)));
            }
        }
        return routes;
    }

    public static IEnumerable<object[]> RoutedOps => Routes().Select(r => new object[] { r.Family, r.Op });

    private static Route Routed(string family, string op) => Routes().Single(r => r.Family == family && r.Op == op);

    private static JsonElement J(object value)
    {
        using var doc = JsonDocument.Parse(JsonSerializer.Serialize(value));
        return doc.RootElement.Clone();
    }

    private static JsonElement J(string json)
    {
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.Clone();
    }

    [Theory]
    [InlineData("queries", "getVersion", "format", "version")]
    [InlineData("diff", "getSemanticChanges", "format", "semantic_changes")]
    public void TheOneToOneSelectorRoutesAreRecorded(string family, string op, string selector, string value)
    {
        var route = Routed(family, op);
        Assert.Equal("docxodus_get_content", route.Tool);
        Assert.Equal(value, Assert.Single(route.Selector, kv => kv.Key == selector).Value);
    }

    [Theory]
    [MemberData(nameof(RoutedOps))]
    public void ARoutedOpIsStillRecordedAbsentAsAnAction(string family, string op)
    {
        var route = Routed(family, op);
        Assert.False(string.IsNullOrWhiteSpace(route.Absent), $"{family}/{op} records a route but no absent reason");
        Assert.NotEmpty(route.Selector);
    }

    [Theory]
    [MemberData(nameof(RoutedOps))]
    public void TheMcpSchemaAdvertisesTheSelectorValue(string family, string op)
    {
        var route = Routed(family, op);
        var tool = Assert.Single(ToolCatalog.Tools, t => t.Name == route.Tool);
        using var schema = JsonDocument.Parse(tool.InputSchemaJson);
        var properties = schema.RootElement.GetProperty("properties");
        foreach (var (property, value) in route.Selector)
        {
            Assert.True(properties.TryGetProperty(property, out var p), $"{route.Tool} has no {property} property");
            Assert.Contains(value, p.GetProperty("enum").EnumerateArray().Select(v => v.GetString()));
        }
    }

    [Theory]
    [MemberData(nameof(RoutedOps))]
    public void CallingTheRouteReturnsTheFacadesResult(string family, string op)
    {
        var route = Routed(family, op);
        Assert.True(route.ArgCount == 0, $"{family}/{op}: only argument-free routes are checked by call so far");

        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path }))).GetProperty("sessionId").GetString()!;
        var handle = _store.Get(sessionId).Handle;
        // One real edit, so the version is past its opening value and the semantic change set is not empty.
        var session = SessionRegistry.Get(handle);
        var paragraph = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        Assert.True(session.ReplaceText(paragraph, "Routed through a selector").Success);

        var args = route.Selector.ToDictionary(kv => kv.Key, kv => (object)kv.Value);
        args["sessionId"] = sessionId;
        var viaMcp = J(Dispatcher.Call(_store, route.Tool, J(args)));

        var facade = typeof(DocxSessionOps).GetMethod(route.Facade, BindingFlags.Public | BindingFlags.Static, new[] { typeof(int) })!;
        var direct = J((string)facade.Invoke(null, new object[] { handle })!);

        Assert.True(JsonElement.DeepEquals(direct, viaMcp), $"{route.Tool} {string.Join(",", route.Selector)} returned {viaMcp}, the facade {direct}");
    }
}

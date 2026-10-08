using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using Docxodus.Internal;
using Docxodus.McpServer;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Checks every transport against the checked-in op description of a session op family
/// (<c>tools/op-descriptions/</c>, issue #982): the facade method, the MCP tool schema, the stdio
/// host's and the MCP server's argument names and defaults (by calling them), the WASM bridge
/// export, the npm wrapper and the Python client. A transport that renames, drops or re-defaults
/// an argument without the description changing fails here. Test IDs use the OPD prefix.
/// </summary>
public class SessionOpDescriptionDriftTests : IDisposable
{
    private const string RepoRoot = "../../../..";

    private readonly string _root;
    private readonly SessionStore _store;

    public SessionOpDescriptionDriftTests()
    {
        _root = Path.Combine(Path.GetTempPath(), $"op-descriptions-{Guid.NewGuid():N}");
        Directory.CreateDirectory(_root);
        _store = new SessionStore(new LocalFileDocumentStore(_root));
    }

    public void Dispose()
    {
        _store.CloseAll();
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    // ─── The description ────────────────────────────────────────────────

    private sealed record Arg(string Name, string Type, bool Required);

    private sealed record Op(
        string Name,
        string Facade,
        IReadOnlyList<Arg> Args,
        string Wasm,
        string Npm,
        string Python,
        string StdioOp,
        IReadOnlyDictionary<string, string> StdioArgNames,
        string McpTool,
        string McpAction,
        IReadOnlyDictionary<string, string> McpArgNames,
        IReadOnlyDictionary<string, JsonElement> McpDefaults)
    {
        public string StdioName(string arg) => StdioArgNames.TryGetValue(arg, out var n) ? n : arg;

        public string McpName(string arg) => McpArgNames.TryGetValue(arg, out var n) ? n : arg;
    }

    private static IReadOnlyList<Op> Family(string family)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(Path.Combine(RepoRoot, "tools", "op-descriptions", $"{family}.json")));
        static IReadOnlyDictionary<string, string> Names(JsonElement t) =>
            t.TryGetProperty("argNames", out var n)
                ? n.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!)
                : new Dictionary<string, string>();
        return doc.RootElement.GetProperty("ops").EnumerateArray().Select(op =>
        {
            var t = op.GetProperty("transports");
            var mcp = t.GetProperty("mcp");
            return new Op(
                op.GetProperty("op").GetString()!,
                op.GetProperty("facade").GetString()!,
                op.GetProperty("args").EnumerateArray().Select(a => new Arg(
                    a.GetProperty("name").GetString()!,
                    a.GetProperty("type").GetString()!,
                    a.GetProperty("required").GetBoolean())).ToList(),
                t.GetProperty("wasm").GetProperty("method").GetString()!,
                t.GetProperty("npm").GetProperty("method").GetString()!,
                t.GetProperty("python").GetProperty("method").GetString()!,
                t.GetProperty("stdio").GetProperty("op").GetString()!,
                Names(t.GetProperty("stdio")),
                mcp.GetProperty("tool").GetString()!,
                mcp.GetProperty("action").GetString()!,
                Names(mcp),
                mcp.TryGetProperty("defaults", out var d)
                    ? d.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.Clone())
                    : new Dictionary<string, JsonElement>());
        }).ToList();
    }

    private static IReadOnlyList<Op> Comments => Family("comments");

    public static IEnumerable<object[]> CommentOps => Family("comments").Select(o => new object[] { o.Name });

    private static Op CommentOp(string name) => Comments.Single(o => o.Name == name);

    // ─── Static surfaces ────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(CommentOps))]
    public void OPD001_TheFacadeOwnsTheOp(string name)
    {
        var op = CommentOp(name);
        var method = typeof(DocxSessionOps).GetMethod(op.Facade, BindingFlags.Public | BindingFlags.Static);
        Assert.True(method != null, $"DocxSessionOps.{op.Facade} is missing");
        var parameters = method!.GetParameters().Select(p => p.Name).ToList();
        Assert.Equal("handle", parameters[0]);
    }

    [Theory]
    [MemberData(nameof(CommentOps))]
    public void OPD002_TheMcpSchemaAdvertisesTheActionAndEveryArgument(string name)
    {
        var op = CommentOp(name);
        var tool = Assert.Single(ToolCatalog.Tools, t => t.Name == op.McpTool);
        using var schema = JsonDocument.Parse(tool.InputSchemaJson);
        var properties = schema.RootElement.GetProperty("properties");
        var actions = properties.GetProperty("action").GetProperty("enum").EnumerateArray().Select(a => a.GetString());
        Assert.Contains(op.McpAction, actions);
        foreach (var arg in op.Args)
            Assert.True(properties.TryGetProperty(op.McpName(arg.Name), out _),
                $"{op.McpTool} does not advertise {op.McpName(arg.Name)} for {op.Name}");
    }

    [Theory]
    [MemberData(nameof(CommentOps))]
    public void OPD003_TheWasmBridgeExportsTheOp(string name)
    {
        var op = CommentOp(name);
        var bridge = File.ReadAllText(Path.Combine(RepoRoot, "wasm", "DocxodusWasm", "DocxSessionBridge.cs"));
        Assert.Matches(new Regex($@"public static string {op.Wasm}\(\s*int h\b"), bridge);
    }

    [Theory]
    [MemberData(nameof(CommentOps))]
    public void OPD004_TheNpmWrapperCallsTheBridgeExport(string name)
    {
        var op = CommentOp(name);
        var body = MethodBody(File.ReadAllText(Path.Combine(RepoRoot, "npm", "src", "session.ts")),
            new Regex($@"^  {op.Npm}\(", RegexOptions.Multiline), new Regex(@"^  \}", RegexOptions.Multiline));
        Assert.True(body != null, $"npm DocxSession.{op.Npm} is missing");
        Assert.Contains($"this.wasm.{op.Wasm}(", body);
    }

    [Theory]
    [MemberData(nameof(CommentOps))]
    public void OPD005_ThePythonClientSendsTheStdioOpWithItsArgumentNames(string name)
    {
        var op = CommentOp(name);
        var body = MethodBody(File.ReadAllText(Path.Combine(RepoRoot, "python", "src", "docx_scalpel", "session.py")),
            new Regex($@"^    def {op.Python}\(", RegexOptions.Multiline), new Regex(@"^    (def |@)", RegexOptions.Multiline));
        Assert.True(body != null, $"Python DocxSession.{op.Python} is missing");
        Assert.Contains($"\"{op.StdioOp}\"", body);
        foreach (var arg in op.Args.Where(a => a.Required))
            Assert.Contains($"\"{op.StdioName(arg.Name)}\"", body);
    }

    /// <summary>The text from the line <paramref name="start"/> matches up to the next line <paramref name="next"/> matches.</summary>
    private static string? MethodBody(string source, Regex start, Regex next)
    {
        var m = start.Match(source);
        if (!m.Success) return null;
        var end = next.Match(source, m.Index + m.Length);
        return source[m.Index..(end.Success ? end.Index : source.Length)];
    }

    // ─── The dispatchers, by calling them ───────────────────────────────

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

    /// <summary>The canonical argument values for each comment op against the fixture's paragraph and comment.</summary>
    private static Dictionary<string, object> Values(Op op, string paragraph, string comment) =>
        op.Args.ToDictionary(a => a.Name, a => (object)(a.Name switch
        {
            "anchorId" => paragraph,
            "commentAnchorId" => comment,
            "span" => new { start = 0, length = 5 },
            "author" => "Reviewer",
            "initials" => "RV",
            "date" => "2026-01-01T00:00:00Z",
            "markdown" => "A remark",
            "resolved" => true,
            _ => throw new InvalidOperationException($"no fixture value for {a.Name}"),
        }));

    /// <summary>Give the session's first paragraph text and one comment on it; return both anchors.</summary>
    private static void Seed(int handle, out string paragraph, out string comment)
    {
        var session = SessionRegistry.Get(handle);
        paragraph = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;
        paragraph = session.ReplaceText(paragraph, "Hello comment world").Modified.Select(a => a.Id).FirstOrDefault() ?? paragraph;
        var added = session.AddComment(paragraph, null, "Seed", "Seed comment");
        Assert.True(added.Success, added.Error?.Message);
        comment = added.Created.First(a => a.Kind == "cmt").Id;
    }

    private static int OpenFixture(out string paragraph, out string comment)
    {
        var handle = DocxSessionOps.OpenSession(DocxSession.CreateBlankDocxBytes(), null);
        Seed(handle, out paragraph, out comment);
        return handle;
    }

    private string OpenMcpFixture(out string paragraph, out string comment)
    {
        var path = Path.Combine(_root, $"{Guid.NewGuid():N}.docx");
        File.WriteAllBytes(path, DocxSession.CreateBlankDocxBytes());
        var sessionId = J(Dispatcher.Call(_store, "docxodus_open", J(new { path }))).GetProperty("sessionId").GetString()!;
        Seed(_store.Get(sessionId).Handle, out paragraph, out comment);
        return sessionId;
    }

    private static bool Succeeded(Func<string> call)
    {
        try
        {
            var result = J(call());
            return result.ValueKind != JsonValueKind.Object
                || !result.TryGetProperty("success", out var success)
                || success.GetBoolean();
        }
        catch (Exception e) when (e is FormatException or McpToolException or KeyNotFoundException or InvalidOperationException or ArgumentException)
        {
            return false;
        }
    }

    // Every op but the revision-targeted add, which needs a tracked revision in the fixture.
    private static readonly string[] Exercised = { "addComment", "addCommentReply", "updateComment", "setCommentResolved", "listComments", "removeComment" };

    [Fact]
    public void OPD010_TheStdioHostAcceptsTheDescribedArgumentNames()
    {
        foreach (var op in Comments.Where(o => Exercised.Contains(o.Name)))
        {
            var handle = OpenFixture(out var paragraph, out var comment);
            try
            {
                var args = Values(op, paragraph, comment).ToDictionary(kv => op.StdioName(kv.Key), kv => kv.Value);
                args["handle"] = handle;
                Assert.True(Succeeded(() => Docxodus.PyHost.Dispatcher.Dispatch(op.StdioOp, J(args))),
                    $"stdio {op.StdioOp} refused the described arguments");
            }
            finally
            {
                DocxSessionOps.CloseSession(handle);
            }
        }
    }

    [Fact]
    public void OPD011_TheStdioHostRequiresWhatTheDescriptionRequires()
    {
        foreach (var op in Comments.Where(o => Exercised.Contains(o.Name)))
        {
            foreach (var required in op.Args.Where(a => a.Required))
            {
                var handle = OpenFixture(out var paragraph, out var comment);
                try
                {
                    var args = Values(op, paragraph, comment)
                        .Where(kv => kv.Key != required.Name)
                        .ToDictionary(kv => op.StdioName(kv.Key), kv => kv.Value);
                    args["handle"] = handle;
                    Assert.False(Succeeded(() => Docxodus.PyHost.Dispatcher.Dispatch(op.StdioOp, J(args))),
                        $"stdio {op.StdioOp} accepted a call without {op.StdioName(required.Name)}, which the description requires");
                }
                finally
                {
                    DocxSessionOps.CloseSession(handle);
                }
            }
        }
    }

    [Fact]
    public void OPD012_TheMcpServerAcceptsTheDescribedArgumentNames()
    {
        foreach (var op in Comments.Where(o => Exercised.Contains(o.Name)))
        {
            var sessionId = OpenMcpFixture(out var paragraph, out var comment);
            var args = Values(op, paragraph, comment).ToDictionary(kv => op.McpName(kv.Key), kv => kv.Value);
            args["sessionId"] = sessionId;
            args["action"] = op.McpAction;
            Assert.True(Succeeded(() => Dispatcher.Call(_store, op.McpTool, J(args))),
                $"MCP {op.McpTool}/{op.McpAction} refused the described arguments");
        }
    }

    [Fact]
    public void OPD013_TheMcpServerRequiresWhatTheDescriptionRequires_UnlessItRecordsADefault()
    {
        foreach (var op in Comments.Where(o => Exercised.Contains(o.Name)))
        {
            foreach (var required in op.Args.Where(a => a.Required))
            {
                var sessionId = OpenMcpFixture(out var paragraph, out var comment);
                var args = Values(op, paragraph, comment)
                    .Where(kv => kv.Key != required.Name)
                    .ToDictionary(kv => op.McpName(kv.Key), kv => kv.Value);
                args["sessionId"] = sessionId;
                args["action"] = op.McpAction;
                var defaulted = op.McpDefaults.ContainsKey(required.Name);
                Assert.True(defaulted == Succeeded(() => Dispatcher.Call(_store, op.McpTool, J(args))),
                    defaulted
                        ? $"MCP {op.McpAction} no longer defaults {required.Name}, as the description records"
                        : $"MCP {op.McpAction} accepted a call without {required.Name}, which the description requires");
            }
        }
    }
}

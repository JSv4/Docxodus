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

    /// <summary>One transport's exposure of an op: its name for the op (null when the description
    /// records the transport as absent, with <see cref="Absent"/> saying why), and its recorded
    /// argument spellings and defaults.</summary>
    /// <param name="Flattens">MCP only: an object argument the tool takes as separate top-level
    /// properties, as argument → (field → MCP property).</param>
    /// <param name="Aliases">MCP only: further tool/action pairs that call the same op.</param>
    /// <param name="ArgsFrom">Python only: <c>file:[Class.]function</c> that builds the wire
    /// arguments for the wrapper, when the wrapper does not spell them itself.</param>
    private sealed record Transport(
        string? Name,
        string? Action,
        string? Absent,
        IReadOnlyDictionary<string, string> ArgNames,
        IReadOnlyDictionary<string, JsonElement> Defaults,
        IReadOnlyList<string> Requires,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, string>> Flattens,
        IReadOnlyList<(string Tool, string Action)> Aliases,
        string? ArgsFrom)
    {
        public bool Exposed => Absent is null;

        public string ArgName(string arg) => ArgNames.TryGetValue(arg, out var n) ? n : arg;
    }

    private sealed record Op(
        string Family,
        string Name,
        string Facade,
        IReadOnlyList<Arg> Args,
        Transport Wasm,
        Transport Npm,
        Transport Python,
        Transport Stdio,
        Transport Mcp)
    {
        public string StdioOp => Stdio.Name!;
        public string McpTool => Mcp.Name!;
        public string McpAction => Mcp.Action!;
        public IReadOnlyDictionary<string, JsonElement> McpDefaults => Mcp.Defaults;

        public string StdioName(string arg) => Stdio.ArgName(arg);

        public string McpName(string arg) => Mcp.ArgName(arg);

        public IEnumerable<(string Name, Transport Transport)> Transports =>
            new[] { ("wasm", Wasm), ("npm", Npm), ("python", Python), ("stdio", Stdio), ("mcp", Mcp) };
    }

    private static readonly string DescriptionDirectory = Path.Combine(RepoRoot, "tools", "op-descriptions");

    /// <summary>The file listing the facade methods that are deliberately not described, each with a reason.</summary>
    private const string NotDescribedFile = "not-described.json";

    private static IEnumerable<string> FamilyFiles =>
        Directory.GetFiles(DescriptionDirectory, "*.json")
            .Where(f => Path.GetFileName(f) != NotDescribedFile)
            .OrderBy(f => f, StringComparer.Ordinal);

    private static IReadOnlyList<Op> Family(string family) => Load(Path.Combine(DescriptionDirectory, $"{family}.json"));

    private static IReadOnlyList<Op> Load(string file)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(file));
        static IReadOnlyDictionary<string, string> Names(JsonElement t) =>
            t.TryGetProperty("argNames", out var n)
                ? n.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!)
                : new Dictionary<string, string>();
        static Transport Read(JsonElement transports, string key, string nameKey)
        {
            var t = transports.GetProperty(key);
            return new Transport(
                t.TryGetProperty(nameKey, out var n) ? n.GetString() : null,
                t.TryGetProperty("action", out var a) ? a.GetString() : null,
                t.TryGetProperty("absent", out var absent) ? absent.GetString() : null,
                Names(t),
                t.TryGetProperty("defaults", out var d)
                    ? d.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.Clone())
                    : new Dictionary<string, JsonElement>(),
                t.TryGetProperty("requires", out var r)
                    ? r.EnumerateArray().Select(v => v.GetString()!).ToList()
                    : new List<string>(),
                t.TryGetProperty("flattens", out var f)
                    ? f.EnumerateObject().ToDictionary(p => p.Name,
                        p => (IReadOnlyDictionary<string, string>)p.Value.EnumerateObject()
                            .ToDictionary(q => q.Name, q => q.Value.GetString()!))
                    : new Dictionary<string, IReadOnlyDictionary<string, string>>(),
                t.TryGetProperty("aliases", out var al)
                    ? al.EnumerateArray().Select(v => (v.GetProperty("tool").GetString()!, v.GetProperty("action").GetString()!)).ToList()
                    : new List<(string, string)>(),
                t.TryGetProperty("argsFrom", out var from) ? from.GetString() : null);
        }
        var family = doc.RootElement.GetProperty("family").GetString()!;
        return doc.RootElement.GetProperty("ops").EnumerateArray().Select(op =>
        {
            var t = op.GetProperty("transports");
            return new Op(
                family,
                op.GetProperty("op").GetString()!,
                op.GetProperty("facade").GetString()!,
                op.GetProperty("args").EnumerateArray().Select(a => new Arg(
                    a.GetProperty("name").GetString()!,
                    a.GetProperty("type").GetString()!,
                    a.GetProperty("required").GetBoolean())).ToList(),
                Read(t, "wasm", "method"),
                Read(t, "npm", "method"),
                Read(t, "python", "method"),
                Read(t, "stdio", "op"),
                Read(t, "mcp", "tool"));
        }).ToList();
    }

    private static IReadOnlyList<Op> AllDescribed => FamilyFiles.SelectMany(Load).ToList();

    private static IReadOnlyList<Op> Comments => Family("comments");

    /// <summary>Every described op, as (family, op) — the static checks run over all of them.</summary>
    public static IEnumerable<object[]> DescribedOps => AllDescribed.Select(o => new object[] { o.Family, o.Name });

    private static Op Described(string family, string name) => Family(family).Single(o => o.Name == name);

    // ─── The description is complete and well formed ────────────────────

    /// <summary>
    /// Every public facade method is described, or listed in <c>not-described.json</c> with a reason
    /// (lifecycle, transactions, previews, delivery evidence: the composite and stateful ops the design
    /// keeps hand-written). A new op that is neither fails here, so it cannot bypass the description.
    /// </summary>
    [Fact]
    public void OPD000_EveryFacadeMethodIsDescribedOrDeliberatelyNot()
    {
        var facade = typeof(DocxSessionOps).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(m => m.Name).ToHashSet(StringComparer.Ordinal);
        var described = AllDescribed.Select(o => o.Facade).ToHashSet(StringComparer.Ordinal);
        using var excludedDoc = JsonDocument.Parse(File.ReadAllText(Path.Combine(DescriptionDirectory, NotDescribedFile)));
        var excluded = excludedDoc.RootElement.GetProperty("methods").EnumerateObject()
            .ToDictionary(p => p.Name, p => p.Value.GetString());

        Assert.All(excluded, e => Assert.False(string.IsNullOrWhiteSpace(e.Value), $"{e.Key} is excluded without a reason"));
        Assert.Empty(facade.Except(described).Except(excluded.Keys).OrderBy(n => n));
        Assert.Empty(described.Intersect(excluded.Keys).OrderBy(n => n));
        Assert.Empty(described.Concat(excluded.Keys).Except(facade).OrderBy(n => n));
    }

    [Fact]
    public void OPD006_EveryDescriptionIsWellFormed()
    {
        var ops = AllDescribed;
        Assert.Empty(ops.GroupBy(o => (o.Family, o.Name)).Where(g => g.Count() > 1).Select(g => g.Key));
        foreach (var op in ops)
        {
            var names = op.Args.Select(a => a.Name).ToList();
            Assert.True(names.Count == names.Distinct().Count(), $"{op.Name} names an argument twice");
            foreach (var (transport, t) in op.Transports)
            {
                var where = $"{op.Family}/{op.Name} {transport}";
                Assert.True(t.Exposed ? t.Name is not null : !string.IsNullOrWhiteSpace(t.Absent),
                    $"{where} is neither named nor recorded absent with a reason");
                Assert.All(t.ArgNames.Keys.Concat(t.Defaults.Keys).Concat(t.Requires).Concat(t.Flattens.Keys),
                    arg => Assert.True(names.Contains(arg), $"{where} records a divergence for unknown argument {arg}"));
                Assert.All(t.Requires, arg => Assert.False(op.Args.Single(a => a.Name == arg).Required,
                    $"{where} records requiring {arg}, which the description already requires"));
            }
            Assert.True(!op.Mcp.Exposed || op.Mcp.Action is not null, $"{op.Name}: an MCP exposure needs an action");
        }
    }

    // ─── Static surfaces ────────────────────────────────────────────────

    [Theory]
    [MemberData(nameof(DescribedOps))]
    public void OPD001_TheFacadeOwnsTheOp(string family, string name)
    {
        var op = Described(family, name);
        var method = typeof(DocxSessionOps).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .FirstOrDefault(m => m.Name == op.Facade);
        Assert.True(method != null, $"DocxSessionOps.{op.Facade} is missing");
        var parameters = method!.GetParameters().Select(p => p.Name).ToList();
        Assert.Equal("handle", parameters[0]);
    }

    [Theory]
    [MemberData(nameof(DescribedOps))]
    public void OPD002_TheMcpSchemaAdvertisesTheActionAndEveryArgument(string family, string name)
    {
        var op = Described(family, name);
        if (!op.Mcp.Exposed) return;
        foreach (var (toolName, action) in op.Mcp.Aliases.Prepend((op.McpTool, op.McpAction)))
        {
            var tool = Assert.Single(ToolCatalog.Tools, t => t.Name == toolName);
            using var schema = JsonDocument.Parse(tool.InputSchemaJson);
            var properties = schema.RootElement.GetProperty("properties");
            var actions = properties.GetProperty("action").GetProperty("enum").EnumerateArray().Select(a => a.GetString());
            Assert.Contains(action, actions);
            foreach (var arg in op.Args)
            {
                var advertised = op.Mcp.Flattens.TryGetValue(arg.Name, out var fields)
                    ? fields.Values
                    : new[] { op.McpName(arg.Name) };
                foreach (var property in advertised)
                    Assert.True(properties.TryGetProperty(property, out _),
                        $"{toolName} does not advertise {property} for {op.Name}");
            }
        }
    }

    [Theory]
    [MemberData(nameof(DescribedOps))]
    public void OPD003_TheWasmBridgeExportsTheOp(string family, string name)
    {
        var op = Described(family, name);
        if (!op.Wasm.Exposed) return;
        var bridge = File.ReadAllText(Path.Combine(RepoRoot, "wasm", "DocxodusWasm", "DocxSessionBridge.cs"));
        Assert.Matches(new Regex($@"public static \w+ {op.Wasm.Name}\(\s*int (h|handle)\b"), bridge);
    }

    [Theory]
    [MemberData(nameof(DescribedOps))]
    public void OPD004_TheNpmWrapperCallsTheBridgeExport(string family, string name)
    {
        var op = Described(family, name);
        if (!op.Npm.Exposed) return;
        var body = MethodBody(File.ReadAllText(Path.Combine(RepoRoot, "npm", "src", "session.ts")),
            new Regex($@"^  (async )?{op.Npm.Name}\(", RegexOptions.Multiline), new Regex(@"^  \}", RegexOptions.Multiline));
        Assert.True(body != null, $"npm DocxSession.{op.Npm.Name} is missing");
        if (op.Wasm.Exposed)
            Assert.Contains($"this.wasm.{op.Wasm.Name}(", body);
    }

    [Theory]
    [MemberData(nameof(DescribedOps))]
    public void OPD005_ThePythonClientSendsTheStdioOpWithItsArgumentNames(string family, string name)
    {
        var op = Described(family, name);
        if (!op.Python.Exposed) return;
        var body = MethodBody(File.ReadAllText(Path.Combine(RepoRoot, "python", "src", "docx_scalpel", "session.py")),
            new Regex($@"^    def {op.Python.Name}\(", RegexOptions.Multiline), new Regex(@"^    (def |@)", RegexOptions.Multiline));
        Assert.True(body != null, $"Python DocxSession.{op.Python.Name} is missing");
        Assert.True(op.Stdio.Exposed, $"Python {op.Python.Name} is described, but the stdio op it sends is recorded absent");
        Assert.Contains($"\"{op.StdioOp}\"", body);
        var argumentSource = body!;
        if (op.Python.ArgsFrom is { } from)
        {
            var helper = PythonFunctionBody(from);
            Assert.True(helper != null, $"Python {op.Python.Name} builds its arguments in {from}, which is missing");
            Assert.Contains(from.Split(':')[1].Split('.')[^1], body);
            argumentSource += helper;
        }
        foreach (var arg in op.Args.Where(a => a.Required))
            Assert.Contains($"\"{op.StdioName(arg.Name)}\"", argumentSource);
    }

    /// <summary>The body of <c>file:[Class.]function</c> in the Python package.</summary>
    private static string? PythonFunctionBody(string reference)
    {
        var (file, qualified) = (reference.Split(':')[0], reference.Split(':')[1]);
        var source = File.ReadAllText(Path.Combine(RepoRoot, "python", "src", "docx_scalpel", file));
        var parts = qualified.Split('.');
        var from = 0;
        if (parts.Length == 2)
        {
            var cls = new Regex($@"^class {parts[0]}\b", RegexOptions.Multiline).Match(source);
            if (!cls.Success) return null;
            from = cls.Index;
        }
        var def = new Regex($@"^(\s*)def {parts[^1]}\(", RegexOptions.Multiline).Match(source, from);
        if (!def.Success) return null;
        var indent = def.Groups[1].Value;
        var end = new Regex($@"^{indent}(def |@|class )|^\S", RegexOptions.Multiline).Match(source, def.Index + def.Length);
        return source[def.Index..(end.Success ? end.Index : source.Length)];
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

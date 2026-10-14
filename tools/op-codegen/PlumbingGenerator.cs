// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace Docxodus.OpCodegen;

/// <summary>One generated piece of source: a whole file, or a marker-delimited region inside a
/// hand-written file when <see cref="Region"/> is set.</summary>
/// <param name="Path">Repository-relative path, with forward slashes.</param>
/// <param name="Content">The generated text, LF line endings.</param>
/// <param name="Region">The region name, for output spliced between
/// <c>// BEGIN GENERATED &lt;region&gt;</c> and <c>// END GENERATED &lt;region&gt;</c> lines.</param>
public sealed record GeneratedOutput(string Path, string Content, string? Region);

/// <summary>
/// Generates the per-transport session plumbing from the op descriptions (issue #1027): for every
/// family marked <c>"generate": true</c>, the stdio host's argument parsing, the MCP server's
/// argument parsing and batch-step validation, the MCP schema properties and the WASM
/// <c>[JSExport]</c> shims. The output is checked in; <c>GeneratedSessionPlumbingTests</c> fails
/// when it is stale. See <c>docs/architecture/session_op_descriptions.md</c>.
/// </summary>
public static class PlumbingGenerator
{
    /// <summary>Where the descriptions live, relative to the repository root.</summary>
    public const string DescriptionDirectory = "tools/op-descriptions";

    /// <summary>The hand-written WASM bridge the shims are spliced into. A separate file is not an
    /// option: the description drift test reads the bridge exports from this file.</summary>
    public const string WasmBridgePath = "wasm/DocxodusWasm/DocxSessionBridge.cs";

    private const string NotDescribedFile = "not-described.json";

    private static readonly JsonSerializerOptions SchemaJson = new() { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };

    /// <summary>The descriptions of every family, in file order.</summary>
    public static IReadOnlyList<(string File, FamilyDescription Family)> LoadFamilies(string repoRoot) =>
        Directory.GetFiles(Path.Combine(repoRoot, DescriptionDirectory), "*.json")
            .Where(f => Path.GetFileName(f) != NotDescribedFile)
            .OrderBy(f => f, StringComparer.Ordinal)
            .Select(f => (Path.GetFileName(f), FamilyDescription.Parse(File.ReadAllText(f))))
            .ToList();

    /// <summary>The generated output for every family marked for generation.</summary>
    public static IReadOnlyList<GeneratedOutput> Generate(string repoRoot)
    {
        var families = LoadFamilies(repoRoot).Where(f => f.Family.Generate).ToList();

        // Each family's generated dispatch owns its stdio ops and MCP tools outright: it throws on
        // a name it does not know, and the schema constants are named by tool.
        foreach (var (what, key) in new (string, Func<OpDescription, string>)[] { ("stdio op", o => o.StdioOp), ("MCP tool", o => o.McpTool) })
        {
            var shared = families.SelectMany(f => f.Family.Ops.Select(key).Distinct().Select(k => (Key: k, f.Family.Family)))
                .GroupBy(x => x.Key).FirstOrDefault(g => g.Count() > 1);
            if (shared is not null)
                throw new CodegenException(
                    $"the {what} {shared.Key} spans the generated families {string.Join(" and ", shared.Select(x => x.Family))}, which the generator does not support yet");
        }

        return families.SelectMany(f => GenerateFamily(f.Family, f.File)).ToList();
    }

    /// <summary>The generated output for one family.</summary>
    /// <param name="sourceFile">The description's file name, quoted in the generated headers.</param>
    public static IReadOnlyList<GeneratedOutput> GenerateFamily(FamilyDescription family, string sourceFile)
    {
        var source = $"{DescriptionDirectory}/{sourceFile}";
        var name = Pascal(family.Family);
        foreach (var op in family.Ops)
        {
            foreach (var arg in op.Args)
            {
                try
                {
                    _ = arg.Kind;
                }
                catch (CodegenException e)
                {
                    throw new CodegenException($"{family.Family}/{op.Name}: {e.Message}");
                }

                if (arg.Name is "args" or "handle" or "h" or "action" or "tool")
                    throw new CodegenException($"{family.Family}/{op.Name}: argument {arg.Name} collides with a generated parameter name");
            }
        }

        return new[]
        {
            new GeneratedOutput($"tools/python-host/Generated/{name}Ops.cs", StdioEmitter.Emit(family, source), null),
            new GeneratedOutput($"tools/mcp-server/Generated/{name}Ops.cs", McpEmitter.Emit(family, source), null),
            new GeneratedOutput(WasmBridgePath, WasmEmitter.Emit(family, source), family.Family),
        };
    }

    /// <summary>The full text each output's file should have, given the files as they are now
    /// (<paramref name="read"/> returns null for a missing file). Several outputs may target one
    /// file; regions are spliced in order. Results are LF without a byte-order mark.</summary>
    public static IReadOnlyDictionary<string, string> Render(IEnumerable<GeneratedOutput> outputs, Func<string, string?> read)
    {
        var files = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var output in outputs)
        {
            if (output.Region is null)
            {
                files[output.Path] = output.Content;
                continue;
            }

            var current = files.TryGetValue(output.Path, out var spliced) ? spliced : Normalize(read(output.Path)
                ?? throw new CodegenException($"{output.Path} is missing; it must hold the GENERATED {output.Region} region"));
            files[output.Path] = Splice(current, output);
        }

        return files;
    }

    /// <summary>LF line endings, no byte-order mark: the form generated text is compared in.</summary>
    public static string Normalize(string text) => text.TrimStart('﻿').Replace("\r\n", "\n");

    /// <summary>The marker line that opens a generated region.</summary>
    public static string BeginMarker(string region) => $"// BEGIN GENERATED {region}";

    /// <summary>The marker line that closes a generated region.</summary>
    public static string EndMarker(string region) => $"// END GENERATED {region}";

    private static string Splice(string text, GeneratedOutput output)
    {
        var lines = text.Split('\n').ToList();
        var begin = lines.FindIndex(l => l.Trim() == BeginMarker(output.Region!));
        var end = lines.FindIndex(l => l.Trim() == EndMarker(output.Region!));
        if (begin < 0 || end < begin)
            throw new CodegenException(
                $"{output.Path} has no \"{BeginMarker(output.Region!)}\" ... \"{EndMarker(output.Region!)}\" region to generate into");
        var generated = (output.Content.EndsWith('\n') ? output.Content[..^1] : output.Content).Split('\n');
        lines.RemoveRange(begin + 1, end - begin - 1);
        lines.InsertRange(begin + 1, generated);
        return string.Join('\n', lines);
    }

    // ─── Shared helpers for the emitters ────────────────────────────────

    /// <summary><c>comments</c> → <c>Comments</c>; <c>headers-footers</c> → <c>HeadersFooters</c>;
    /// <c>add_comment</c> → <c>AddComment</c>.</summary>
    internal static string Pascal(string name) =>
        string.Concat(name.Split('-', '_').Where(p => p.Length > 0).Select(p => char.ToUpperInvariant(p[0]) + p[1..]));

    /// <summary>The MCP tool's name without its <c>docxodus_</c> prefix: <c>docxodus_comment</c> → <c>comment</c>.</summary>
    internal static string ToolNoun(string tool) => tool.StartsWith("docxodus_", StringComparison.Ordinal) ? tool["docxodus_".Length..] : tool;

    internal static string Header(string source) =>
        $"// Generated by tools/op-codegen from {source}.\n"
        + "// Do not edit by hand: change the description and run\n"
        + "// `dotnet run --project tools/op-codegen` (docs/architecture/session_op_descriptions.md).\n";

    /// <summary>A C# string literal.</summary>
    internal static string Quote(string s) => "\"" + s.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    /// <summary>A JSON string, as the MCP schema writes it.</summary>
    internal static string JsonQuote(string s) => JsonSerializer.Serialize(s, SchemaJson);

    /// <summary>"a or b", "a, b or c".</summary>
    internal static string OrList(IReadOnlyList<string> items) =>
        items.Count == 1 ? items[0] : string.Join(", ", items.Take(items.Count - 1)) + " or " + items[^1];

    /// <summary>A C# local for an argument, escaped when it is a keyword.</summary>
    internal static string Local(string name) => CSharpKeywords.Contains(name) ? "@" + name : name;

    private static readonly HashSet<string> CSharpKeywords = new(StringComparer.Ordinal)
    {
        "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char", "checked", "class", "const",
        "continue", "decimal", "default", "delegate", "do", "double", "else", "enum", "event", "explicit", "extern",
        "false", "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit", "in", "int", "interface",
        "internal", "is", "lock", "long", "namespace", "new", "null", "object", "operator", "out", "override",
        "params", "private", "protected", "public", "readonly", "ref", "return", "sbyte", "sealed", "short",
        "sizeof", "stackalloc", "static", "string", "struct", "switch", "this", "throw", "true", "try", "typeof",
        "uint", "ulong", "unchecked", "unsafe", "ushort", "using", "virtual", "void", "volatile", "while",
    };

    /// <summary>The ops in <paramref name="ops"/> grouped by transport name, in first-appearance order.</summary>
    internal static IReadOnlyList<SharedName> GroupBy(IEnumerable<OpDescription> ops, Func<OpDescription, string> key)
    {
        var groups = new List<SharedName>();
        foreach (var op in ops)
        {
            var existing = groups.FindIndex(g => g.Key == key(op));
            if (existing < 0) groups.Add(new SharedName(key(op), new List<OpDescription> { op }));
            else ((List<OpDescription>)groups[existing].Ops).Add(op);
        }

        return groups;
    }

    /// <summary>A facade call, one argument per line.</summary>
    internal static string Call(OpDescription op, string handle, IEnumerable<string> values, string indent)
    {
        var all = new[] { handle }.Concat(values).ToList();
        if (all.Count == 1) return $"DocxSessionOps.{op.Facade}({handle})";
        var sb = new StringBuilder($"DocxSessionOps.{op.Facade}(");
        for (var i = 0; i < all.Count; i++)
        {
            sb.Append('\n').Append(indent).Append(all[i].Replace("\n", "\n" + indent));
            sb.Append(i < all.Count - 1 ? "," : ")");
        }

        return sb.ToString();
    }
}

/// <summary>
/// Ops that one transport reaches under one name, such as <c>addComment</c> and
/// <c>addCommentToRevision</c>, which are both the stdio <c>add_comment</c> op and the MCP
/// <c>docxodus_comment add</c> action. The caller picks one by naming exactly one op's
/// <see cref="Selector"/>: the one required argument no other op in the group has.
/// </summary>
internal sealed record SharedName(string Key, IReadOnlyList<OpDescription> Ops)
{
    public bool IsShared => Ops.Count > 1;

    /// <summary>The arguments every op in the group takes, in the first op's order.</summary>
    public IReadOnlyList<OpArg> Common => Ops[0].Args.Where(a => Ops.All(o => o.Args.Any(b => b.Name == a.Name))).ToList();

    /// <summary>The arguments of <paramref name="op"/> that some other op in the group lacks.</summary>
    public IReadOnlyList<OpArg> Exclusive(OpDescription op) =>
        op.Args.Where(a => Ops.Any(o => o.Args.All(b => b.Name != a.Name))).ToList();

    /// <summary>The required argument that picks <paramref name="op"/>.</summary>
    public OpArg Selector(OpDescription op)
    {
        var required = Exclusive(op).Where(a => a.Required).ToList();
        if (required.Count != 1 || required[0].Kind != ArgKind.String)
            throw new CodegenException(
                $"{Key}: {op.Name} shares the name with {string.Join(", ", Ops.Where(o => o != op).Select(o => o.Name))}, so it needs exactly one required string argument the others lack to be picked by");
        if (op.StdioAliases.ContainsKey(required[0].Name))
            throw new CodegenException($"{Key}: {op.Name} is picked by {required[0].Name}, which cannot also have a deprecated alias");
        return required[0];
    }

    /// <summary>The optional arguments only <paramref name="op"/> takes.</summary>
    public IReadOnlyList<OpArg> Extras(OpDescription op) => Exclusive(op).Where(a => !a.Required).ToList();

    /// <summary>Optional arguments of other ops that <paramref name="op"/> refuses when it is picked.</summary>
    public IReadOnlyList<string> Forbidden(OpDescription op) =>
        Ops.Where(o => o != op).SelectMany(Extras).Select(a => a.Name)
            .Where(n => op.Args.All(a => a.Name != n)).Distinct().ToList();

    /// <summary>Every op sharing the name must agree on the arguments they share.</summary>
    public void CheckConsistent()
    {
        foreach (var arg in Ops.SelectMany(o => o.Args).GroupBy(a => a.Name))
        {
            if (arg.Select(a => (a.Type, a.Required, a.HasDefault)).Distinct().Count() > 1)
                throw new CodegenException($"{Key}: the ops sharing this name describe {arg.Key} differently");
        }

        foreach (var arg in Common)
        {
            if (Ops.Select(o => o.StdioAliases.TryGetValue(arg.Name, out var alias) ? alias : null).Distinct().Count() > 1)
                throw new CodegenException($"{Key}: the ops sharing this name give {arg.Name} different deprecated aliases");
        }

        foreach (var op in Ops) _ = Selector(op);
    }

    /// <summary>The "exactly one target" refusal: <c>anchorId (with optional span) or revisionId</c>.</summary>
    public string TargetList(bool withExtras) => PlumbingGenerator.OrList(Ops.Select(o =>
        Selector(o).Name + (withExtras && Extras(o).Count > 0 ? $" (with optional {string.Join(", ", Extras(o).Select(a => a.Name))})" : "")).ToList());

    /// <summary>The expression that is true when the caller names a number of targets other than one.</summary>
    public string NotExactlyOne() =>
        "(" + string.Join(" + ", Ops.Select(o => $"({PlumbingGenerator.Local(Selector(o).Name)} is null ? 0 : 1)")) + ") != 1";

    public static string Present(string name) =>
        $"args.ValueKind == JsonValueKind.Object && args.TryGetProperty({PlumbingGenerator.Quote(name)}, out _)";
}

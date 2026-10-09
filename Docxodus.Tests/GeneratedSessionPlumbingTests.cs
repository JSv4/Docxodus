// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Docxodus.Internal;
using Docxodus.McpServer;
using Docxodus.OpCodegen;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The session plumbing generated from the op descriptions (issue #1027, stage 3 of #982):
/// <c>tools/op-codegen</c> emits the stdio host's and the MCP server's argument parsing, the MCP
/// schema properties and the WASM shims for every family marked <c>"generate": true</c>. These
/// tests fail when the checked-in output is stale, and pin the assumptions the generator makes.
/// Test IDs use the OPG prefix.
/// </summary>
public class GeneratedSessionPlumbingTests
{
    private const string RepoRoot = "../../../..";

    private const string Regenerate = "run `dotnet run --project tools/op-codegen` and commit the result";

    private static string? Read(string path)
    {
        var full = Path.Combine(RepoRoot, path);
        return File.Exists(full) ? File.ReadAllText(full) : null;
    }

    private static IEnumerable<(string File, FamilyDescription Family)> GeneratedFamilies =>
        PlumbingGenerator.LoadFamilies(RepoRoot).Where(f => f.Family.Generate);

    public static IEnumerable<object[]> GeneratedOps =>
        GeneratedFamilies.SelectMany(f => f.Family.Ops.Select(o => new object[] { f.Family.Family, o.Name }));

    /// <summary>The text each generated file would have if the generator ran on the checked-in descriptions.</summary>
    private static IReadOnlyDictionary<string, string> Expected(IEnumerable<GeneratedOutput> outputs) =>
        PlumbingGenerator.Render(outputs, Read);

    [Fact]
    public void OPG001_TheCheckedInPlumbingIsWhatTheDescriptionsGenerate()
    {
        var expected = Expected(PlumbingGenerator.Generate(RepoRoot));
        Assert.NotEmpty(expected);
        foreach (var (path, text) in expected)
        {
            var current = Read(path);
            Assert.True(current is not null, $"{path} is missing; {Regenerate}");
            Assert.True(PlumbingGenerator.Normalize(current!) == text, $"{path} is stale; {Regenerate}");
        }
    }

    [Fact]
    public void OPG002_TheCommentFamilyIsGeneratedForStdioMcpAndWasm()
    {
        var comments = Assert.Single(GeneratedFamilies, f => f.Family.Family == "comments");
        var paths = PlumbingGenerator.GenerateFamily(comments.Family, comments.File).Select(o => o.Path).ToList();
        Assert.Equal(
            new[] { "tools/python-host/Generated/CommentsOps.cs", "tools/mcp-server/Generated/CommentsOps.cs", PlumbingGenerator.WasmBridgePath },
            paths);
    }

    /// <summary>
    /// The staleness check sees a change to the description, not just to the generated files: a
    /// reworded MCP property, a renamed stdio alias or an argument made required each change the output.
    /// </summary>
    [Theory]
    [InlineData("add/reply: comment author (required).", "add/reply: who wrote the comment (required).")]
    [InlineData("\"commentAnchorId\": \"parentAnchorId\"", "\"commentAnchorId\": \"parentId\"")]
    [InlineData("{ \"name\": \"resolved\", \"type\": \"boolean\", \"required\": false, \"default\": true }", "{ \"name\": \"resolved\", \"type\": \"boolean\", \"required\": true }")]
    public void OPG003_ChangingTheDescriptionMakesTheCheckedInPlumbingStale(string find, string replace)
    {
        var json = PlumbingGenerator.Normalize(File.ReadAllText(Path.Combine(RepoRoot, PlumbingGenerator.DescriptionDirectory, "comments.json")));
        Assert.Contains(find, json);
        var mutated = FamilyDescription.Parse(json.Replace(find, replace));

        var expected = Expected(PlumbingGenerator.GenerateFamily(mutated, "comments.json"));
        Assert.Contains(expected, kv => PlumbingGenerator.Normalize(Read(kv.Key)!) != kv.Value);
    }

    /// <summary>
    /// The generated code passes the description's arguments to the facade by position, so the
    /// description must list them in the facade's parameter order, with matching types. A swap of
    /// two strings would still compile; this catches it by kind and count at least, and by name
    /// where the facade spells the parameter the same way.
    /// </summary>
    [Theory]
    [MemberData(nameof(GeneratedOps))]
    public void OPG004_TheDescriptionListsTheArgumentsInFacadeParameterOrder(string family, string name)
    {
        var op = GeneratedFamilies.Single(f => f.Family.Family == family).Family.Ops.Single(o => o.Name == name);
        static List<ParameterInfo> Described(MethodInfo m) => m.GetParameters().Skip(1).Where(p => p.Name != "preconditions").ToList();
        var method = Assert.Single(
            typeof(DocxSessionOps).GetMethods(BindingFlags.Public | BindingFlags.Static),
            m => m.Name == op.Facade && Described(m).Count == op.Args.Count);
        var parameters = Described(method);
        for (var i = 0; i < parameters.Count; i++)
        {
            var (arg, parameter) = (op.Args[i], parameters[i]);
            var expected = arg.Kind switch
            {
                ArgKind.String => new[] { typeof(string) },
                ArgKind.Boolean => new[] { typeof(bool), typeof(bool?) },
                _ => new[] { typeof(CharSpan?) },
            };
            Assert.True(expected.Contains(parameter.ParameterType),
                $"{op.Name}: argument {i} ({arg.Name}, {arg.Type}) meets facade parameter {parameter.Name} of type {parameter.ParameterType}");
            var others = op.Args.Where((_, j) => j != i).Select(a => a.Name);
            Assert.False(others.Contains(parameter.Name, StringComparer.Ordinal),
                $"{op.Name}: facade parameter {parameter.Name} is at position {i}, but the description puts {parameter.Name} elsewhere");
        }
    }

    /// <summary>A tool's schema merges generated and hand-written properties; a name spelled by both
    /// would be a duplicate JSON key, which JSON parsers resolve silently.</summary>
    [Fact]
    public void OPG005_NoMcpToolSchemaAdvertisesAPropertyTwice()
    {
        foreach (var tool in ToolCatalog.Tools)
        {
            var reader = new Utf8JsonReader(Encoding.UTF8.GetBytes(tool.InputSchemaJson));
            var names = new Stack<HashSet<string>>();
            while (reader.Read())
            {
                switch (reader.TokenType)
                {
                    case JsonTokenType.StartObject:
                        names.Push(new HashSet<string>(StringComparer.Ordinal));
                        break;
                    case JsonTokenType.EndObject:
                        names.Pop();
                        break;
                    case JsonTokenType.PropertyName:
                        var property = reader.GetString()!;
                        Assert.True(names.Peek().Add(property), $"{tool.Name} names {property} twice in one schema object");
                        break;
                }
            }
        }
    }

    [Fact]
    public void OPG006_TheGeneratorRefusesAFamilyThatStillRecordsADivergence()
    {
        var json = CommentsWith(root =>
            root["ops"]![0]!["transports"]!["mcp"]!["argNames"] = new JsonObject { ["anchorId"] = "paragraphId" });
        var e = Assert.Throws<CodegenException>(() => FamilyDescription.Parse(json));
        Assert.Contains("argNames", e.Message);
    }

    [Fact]
    public void OPG007_TheGeneratorRefusesAnMcpPropertyWithoutProse()
    {
        var json = CommentsWith(root => ((JsonObject)root["mcp"]!["docxodus_comment"]!["properties"]!).Remove("resolved"));
        var e = Assert.Throws<CodegenException>(() => PlumbingGenerator.GenerateFamily(FamilyDescription.Parse(json), "comments.json"));
        Assert.Contains("resolved", e.Message);
    }

    [Fact]
    public void OPG008_TheGeneratorRefusesAnArgumentTypeItCannotRead()
    {
        var json = CommentsWith(root => root["ops"]![0]!["args"]![0]!["type"] = "listFormat");
        var e = Assert.Throws<CodegenException>(() => PlumbingGenerator.GenerateFamily(FamilyDescription.Parse(json), "comments.json"));
        Assert.Contains("listFormat", e.Message);
    }

    /// <summary>
    /// <c>ToolCatalog</c> splices the generated schema fragments into each tool's hand-written
    /// schema. A tool that stops splicing them, or splices them out of order, fails here even though
    /// the generated file itself is current.
    /// </summary>
    [Fact]
    public void OPG009_EachToolSchemaCarriesItsGeneratedActionsAndProperties()
    {
        foreach (var (_, family) in GeneratedFamilies)
        {
            foreach (var tool in family.Ops.Select(o => o.McpTool).Distinct())
            {
                var definition = Assert.Single(ToolCatalog.Tools, t => t.Name == tool);
                using var schema = JsonDocument.Parse(definition.InputSchemaJson);
                var properties = schema.RootElement.GetProperty("properties");
                var actions = properties.GetProperty("action").GetProperty("enum").EnumerateArray().Select(a => a.GetString()!).ToList();
                var generated = family.Ops.Where(o => o.McpTool == tool).Select(o => o.McpAction).Distinct().ToList();
                var start = actions.IndexOf(generated[0]);
                Assert.True(start >= 0 && actions.Skip(start).Take(generated.Count).SequenceEqual(generated),
                    $"{tool}'s action enum does not carry the generated actions {string.Join(", ", generated)} in order");

                var names = properties.EnumerateObject().Select(p => p.Name).ToList();
                var described = family.McpProperties[tool];
                var at = names.IndexOf(described[0].Name);
                Assert.True(at >= 0 && names.Skip(at).Take(described.Count).SequenceEqual(described.Select(p => p.Name)),
                    $"{tool}'s schema does not carry the generated properties in order");
                foreach (var property in described)
                    Assert.Equal(property.Description, properties.GetProperty(property.Name).GetProperty("description").GetString());
            }
        }
    }

    /// <summary>Generated code lives only where the generator puts it: no generated file outlives
    /// its family, and each generated WASM shim is defined once, inside its family's region.</summary>
    [Fact]
    public void OPG010_NoGeneratedCodeLivesOutsideTheGeneratorsOutputs()
    {
        var outputs = PlumbingGenerator.Generate(RepoRoot);
        var paths = outputs.Select(o => o.Path).ToHashSet(StringComparer.Ordinal);
        foreach (var directory in new[] { "tools/python-host/Generated", "tools/mcp-server/Generated" })
        {
            var full = Path.Combine(RepoRoot, directory);
            if (!Directory.Exists(full)) continue;
            foreach (var file in Directory.GetFiles(full, "*", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(RepoRoot, file).Replace('\\', '/');
                Assert.True(paths.Contains(relative), $"{relative} is not generated by any family; delete it or mark its family \"generate\"");
            }
        }

        var bridge = PlumbingGenerator.Normalize(Read(PlumbingGenerator.WasmBridgePath)!);
        foreach (var (_, family) in GeneratedFamilies)
        {
            var begin = bridge.IndexOf(PlumbingGenerator.BeginMarker(family.Family), StringComparison.Ordinal);
            var end = bridge.IndexOf(PlumbingGenerator.EndMarker(family.Family), StringComparison.Ordinal);
            Assert.True(begin >= 0 && end > begin, $"the bridge has no generated region for {family.Family}");
            foreach (var op in family.Ops)
            {
                var exports = System.Text.RegularExpressions.Regex.Matches(bridge, $@"public static \w+ {op.WasmMethod}\(");
                var export = Assert.Single(exports);
                Assert.True(export.Index > begin && export.Index < end, $"{op.WasmMethod} is defined outside the generated {family.Family} region");
            }
        }
    }

    private static string CommentsWith(Action<JsonNode> mutate)
    {
        var root = JsonNode.Parse(File.ReadAllText(Path.Combine(RepoRoot, PlumbingGenerator.DescriptionDirectory, "comments.json")))!;
        mutate(root);
        return root.ToJsonString();
    }
}

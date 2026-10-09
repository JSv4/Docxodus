// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;

namespace Docxodus.OpCodegen;

/// <summary>A description the generator cannot turn into code, with the reason.</summary>
public sealed class CodegenException(string message) : Exception(message);

/// <summary>How an argument travels: the generator knows how every transport reads and checks each kind.</summary>
public enum ArgKind
{
    /// <summary>A JSON string: the description types <c>anchor</c>, <c>string</c> and <c>markdown</c>.</summary>
    String,

    /// <summary>A JSON boolean: the description type <c>boolean</c>.</summary>
    Boolean,

    /// <summary>A <c>{ start, length }</c> object, a <c>CharSpan?</c> on the facade: the description type <c>charSpan</c>.</summary>
    CharSpan,
}

/// <summary>One argument of a described op, in facade parameter order.</summary>
/// <param name="Name">The canonical wire name.</param>
/// <param name="Type">The description's type name.</param>
/// <param name="Required">Whether every transport requires it.</param>
/// <param name="HasDefault">Whether the facade applies a default when it is absent.</param>
public sealed record OpArg(string Name, string Type, bool Required, bool HasDefault)
{
    /// <summary>The kind the description's <see cref="Type"/> maps to.</summary>
    public ArgKind Kind => Type switch
    {
        "anchor" or "string" or "markdown" => ArgKind.String,
        "boolean" => ArgKind.Boolean,
        "charSpan" => ArgKind.CharSpan,
        _ => throw new CodegenException($"argument {Name} has type {Type}, which the generator does not support yet"),
    };
}

/// <summary>One described op, as the generator reads it.</summary>
/// <param name="StdioAliases">Deprecated spellings the stdio host still accepts, as argument → old name.</param>
public sealed record OpDescription(
    string Name,
    string Facade,
    IReadOnlyList<OpArg> Args,
    string WasmMethod,
    string? WasmDoc,
    string StdioOp,
    IReadOnlyDictionary<string, string> StdioAliases,
    string McpTool,
    string McpAction);

/// <summary>One MCP schema property and its prose, in the order the tool advertises it.</summary>
public sealed record McpProperty(string Name, string Description);

/// <summary>A family description (<c>tools/op-descriptions/&lt;family&gt;.json</c>).</summary>
/// <param name="McpProperties">For each MCP tool, the properties the family's actions advertise, in order.</param>
public sealed record FamilyDescription(
    string Family,
    bool Generate,
    IReadOnlyList<OpDescription> Ops,
    IReadOnlyDictionary<string, IReadOnlyList<McpProperty>> McpProperties)
{
    /// <summary>Parse a family description. A family marked <c>generate</c> must be complete: every
    /// transport exposed and no recorded divergence, because a generator cannot emit two spellings
    /// of one argument.</summary>
    public static FamilyDescription Parse(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var family = root.GetProperty("family").GetString()!;
        var generate = root.TryGetProperty("generate", out var g) && g.GetBoolean();

        var mcp = new Dictionary<string, IReadOnlyList<McpProperty>>(StringComparer.Ordinal);
        if (root.TryGetProperty("mcp", out var tools))
        {
            foreach (var tool in tools.EnumerateObject())
            {
                mcp[tool.Name] = tool.Value.GetProperty("properties").EnumerateObject()
                    .Select(p => new McpProperty(
                        p.Name,
                        p.Value.TryGetProperty("description", out var d) ? d.GetString()! : throw new CodegenException(
                            $"{family}: mcp.{tool.Name}.properties.{p.Name} has no description")))
                    .ToList();
            }
        }

        var ops = new List<OpDescription>();
        foreach (var op in root.GetProperty("ops").EnumerateArray())
        {
            var name = op.GetProperty("op").GetString()!;
            var args = op.GetProperty("args").EnumerateArray().Select(a => new OpArg(
                a.GetProperty("name").GetString()!,
                a.GetProperty("type").GetString()!,
                a.GetProperty("required").GetBoolean(),
                a.TryGetProperty("default", out _))).ToList();
            var transports = op.GetProperty("transports");
            if (generate)
            {
                foreach (var transport in transports.EnumerateObject())
                {
                    foreach (var divergence in new[] { "absent", "argNames", "defaults", "requires", "flattens", "aliases" })
                    {
                        if (transport.Value.TryGetProperty(divergence, out _))
                            throw new CodegenException(
                                $"{family}/{name}: the {transport.Name} transport records \"{divergence}\"; resolve the divergence before generating the family");
                    }
                }
            }

            static string? Text(JsonElement e, string key) => e.TryGetProperty(key, out var v) ? v.GetString() : null;
            var wasm = transports.GetProperty("wasm");
            var stdio = transports.GetProperty("stdio");
            var mcpTransport = transports.GetProperty("mcp");
            var aliases = stdio.TryGetProperty("deprecatedAliases", out var al)
                ? al.EnumerateObject().ToDictionary(p => p.Name, p => p.Value.GetString()!, StringComparer.Ordinal)
                : new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var aliased in aliases.Keys)
            {
                if (args.All(a => a.Name != aliased))
                    throw new CodegenException($"{family}/{name}: stdio deprecatedAliases names unknown argument {aliased}");
            }

            ops.Add(new OpDescription(
                name,
                op.GetProperty("facade").GetString()!,
                args,
                Text(wasm, "method") ?? "",
                Text(wasm, "doc"),
                Text(stdio, "op") ?? "",
                aliases,
                Text(mcpTransport, "tool") ?? "",
                Text(mcpTransport, "action") ?? ""));
        }

        return new FamilyDescription(family, generate, ops, mcp);
    }
}

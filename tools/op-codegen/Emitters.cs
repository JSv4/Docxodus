// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using static Docxodus.OpCodegen.PlumbingGenerator;

namespace Docxodus.OpCodegen;

/// <summary>
/// Emits the body of a shared-name dispatch (see <see cref="SharedName"/>) for a transport that
/// parses JSON arguments: read every selector, refuse anything but exactly one, refuse another
/// op's optional arguments, then call the picked op's facade method.
/// </summary>
internal static class SharedNameParse
{
    public static void Emit(
        StringBuilder sb,
        SharedName group,
        string signature,
        string label,
        string exception,
        string handle,
        Func<OpArg, OpDescription, string> read)
    {
        group.CheckConsistent();
        sb.Append($"    /// <summary><c>{label}</c>: {string.Join(" or ", group.Ops.Select(o => o.Name))}, picked by which of {OrList(group.Ops.Select(o => group.Selector(o).Name).ToList())} the caller names.</summary>\n");
        sb.Append($"    {signature}\n    {{\n");
        foreach (var op in group.Ops)
        {
            var selector = group.Selector(op).Name;
            sb.Append($"        var {Local(selector)} = OptStr(args, {Quote(selector)});\n");
        }

        var refusals = new List<string> { group.NotExactlyOne() };
        foreach (var op in group.Ops)
        {
            foreach (var forbidden in group.Forbidden(op))
                refusals.Add($"({Local(group.Selector(op).Name)} is not null && {SharedName.Present(forbidden)})");
        }

        sb.Append("        if (").Append(string.Join("\n            || ", refusals)).Append(")\n");
        sb.Append($"            throw new {exception}({Quote($"{label} requires exactly one target: {group.TargetList(withExtras: true)}")});\n");
        for (var i = 0; i < group.Ops.Count; i++)
        {
            var op = group.Ops[i];
            var selector = group.Selector(op).Name;
            var last = i == group.Ops.Count - 1;
            var values = op.Args.Select(a => a.Name == selector ? Local(selector) + (last ? "!" : "") : read(a, op));
            if (last)
            {
                sb.Append("        return ").Append(Call(op, handle, values, "            ")).Append(";\n");
            }
            else
            {
                sb.Append($"        if ({Local(selector)} is not null)\n");
                sb.Append("            return ").Append(Call(op, handle, values, "                ")).Append(";\n");
            }
        }

        sb.Append("    }\n");
    }
}

/// <summary>The stdio host's argument parsing (<c>tools/python-host/Generated/</c>).</summary>
internal static class StdioEmitter
{
    public static string Emit(FamilyDescription family, string source)
    {
        var name = Pascal(family.Family);
        var groups = GroupBy(family.Ops, o => o.StdioOp);
        var sb = new StringBuilder(Header(source));
        sb.Append("\nusing System;\nusing System.Text.Json;\nusing Docxodus.Internal;\n\nnamespace Docxodus.PyHost;\n\n");
        sb.Append($"/// <summary>The stdio host's {family.Family} ops: argument parsing into <see cref=\"DocxSessionOps\"/> calls.</summary>\n");
        sb.Append("internal static partial class Dispatcher\n{\n");
        sb.Append($"    /// <summary>Whether <paramref name=\"op\"/> is one of the generated {family.Family} ops.</summary>\n");
        sb.Append($"    private static bool IsGenerated{name}Op(string op) => op\n");
        for (var i = 0; i < groups.Count; i++)
            sb.Append(i == 0 ? "        is " : "        or ").Append(Quote(groups[i].Key)).Append(i == groups.Count - 1 ? ";\n" : "\n");
        sb.Append($"\n    /// <summary>Parse <paramref name=\"args\"/> for a generated {family.Family} op and call the facade.</summary>\n");
        sb.Append($"    private static string DispatchGenerated{name}Op(string op, JsonElement args) => op switch\n    {{\n");
        foreach (var group in groups)
        {
            sb.Append($"        {Quote(group.Key)} => ");
            if (group.IsShared)
            {
                sb.Append($"DispatchGenerated{Pascal(group.Key)}(args),\n");
                continue;
            }

            var op = group.Ops[0];
            sb.Append(Call(op, "Handle(args)", op.Args.Select(a => Read(a, op)), "            ")).Append(",\n");
        }

        sb.Append($"        _ => throw new InvalidOperationException($\"{{op}} is not a generated {family.Family} op\"),\n    }};\n");
        foreach (var group in groups.Where(g => g.IsShared))
        {
            sb.Append('\n');
            SharedNameParse.Emit(sb, group, $"private static string DispatchGenerated{Pascal(group.Key)}(JsonElement args)",
                group.Key, "FormatException", "Handle(args)", Read);
        }

        sb.Append("}\n");
        return sb.ToString();
    }

    private static string Read(OpArg arg, OpDescription op)
    {
        var n = Quote(arg.Name);
        var alias = op.StdioAliases.TryGetValue(arg.Name, out var a) ? Quote(a) : null;
        return (arg.Kind, arg.Required, alias) switch
        {
            (ArgKind.String, true, null) => $"Str(args, {n})",
            (ArgKind.String, true, _) => $"DocxSessionJson.AliasedString(args, {n}, {alias})\n    ?? throw new FormatException({Quote($"args missing string \"{arg.Name}\"")})",
            (ArgKind.String, false, null) => $"OptStr(args, {n})",
            (ArgKind.String, false, _) => $"DocxSessionJson.AliasedString(args, {n}, {alias})",
            (ArgKind.Boolean, true, null) => $"OptBool(args, {n})\n    ?? throw new FormatException({Quote($"args missing boolean \"{arg.Name}\"")})",
            (ArgKind.Boolean, true, _) => $"DocxSessionJson.AliasedBool(args, {n}, {alias})\n    ?? throw new FormatException({Quote($"args missing boolean \"{arg.Name}\"")})",
            (ArgKind.Boolean, false, null) => $"OptBool(args, {n})",
            (ArgKind.Boolean, false, _) => $"DocxSessionJson.AliasedBool(args, {n}, {alias})",
            (ArgKind.CharSpan, false, null) => $"ParseOptionalSpan(args, {n})",
            _ => throw new CodegenException($"{op.Name}: the stdio host cannot read a {(arg.Required ? "required" : "optional")} {arg.Type} argument{(alias is null ? "" : " with a deprecated alias")} ({arg.Name}) yet"),
        };
    }
}

/// <summary>The MCP server's argument parsing, batch-step validation and schema properties
/// (<c>tools/mcp-server/Generated/</c>). The grouped-intent tool shells stay hand-written and
/// call these.</summary>
internal static class McpEmitter
{
    public static string Emit(FamilyDescription family, string source)
    {
        // Deprecated aliases are a stdio-host concession; the MCP server takes canonical names only.
        var name = Pascal(family.Family);
        var groups = GroupBy(family.Ops, o => o.McpTool + "\n" + o.McpAction);
        var sb = new StringBuilder(Header(source));
        sb.Append("\nusing System.Text.Json;\nusing Docxodus.Internal;\n\nnamespace Docxodus.McpServer;\n\n");
        sb.Append($"/// <summary>The MCP server's {family.Family} actions: argument parsing into <see cref=\"DocxSessionOps\"/> calls and batch-step validation.</summary>\n");
        sb.Append("internal static partial class Dispatcher\n{\n");

        sb.Append($"    /// <summary>Parse <paramref name=\"args\"/> for a generated {family.Family} action and call the facade.</summary>\n");
        sb.Append($"    private static string RunGenerated{name}Action(int handle, string tool, string action, JsonElement args) => (tool, action) switch\n    {{\n");
        foreach (var group in groups)
        {
            var op = group.Ops[0];
            sb.Append($"        ({Quote(op.McpTool)}, {Quote(op.McpAction)}) => ");
            if (group.IsShared)
            {
                sb.Append($"RunGenerated{SharedMethod(op)}(handle, args),\n");
                continue;
            }

            sb.Append(Call(op, "handle", op.Args.Select(a => Read(a, op)), "            ")).Append(",\n");
        }

        sb.Append("        _ => throw new McpToolException($\"unknown {tool} action: {action}\"),\n    };\n\n");

        sb.Append($"    /// <summary>Check a batch step's arguments for a generated {family.Family} action without running it.</summary>\n");
        sb.Append($"    private static void ValidateGenerated{name}Arguments(string tool, string action, JsonElement args)\n    {{\n");
        sb.Append("        switch ((tool, action))\n        {\n");
        foreach (var group in groups)
        {
            var op = group.Ops[0];
            sb.Append($"            case ({Quote(op.McpTool)}, {Quote(op.McpAction)}):\n");
            if (group.IsShared)
                sb.Append($"                ValidateGenerated{SharedMethod(op)}(args);\n");
            else
                foreach (var line in Validation(op.Args)) sb.Append("                ").Append(line).Append('\n');
            sb.Append("                break;\n");
        }

        sb.Append("        }\n    }\n");

        foreach (var group in groups.Where(g => g.IsShared))
        {
            var op = group.Ops[0];
            var label = $"{op.McpTool} {op.McpAction}";
            sb.Append('\n');
            SharedNameParse.Emit(sb, group, $"private static string RunGenerated{SharedMethod(op)}(int handle, JsonElement args)",
                label, "McpToolException", "handle", Read);
            sb.Append('\n');
            EmitSharedValidation(sb, group, op);
        }

        sb.Append("}\n");
        EmitSchema(sb, family);
        return sb.ToString();
    }

    private static string SharedMethod(OpDescription op) => Pascal(ToolNoun(op.McpTool)) + Pascal(op.McpAction);

    private static string Read(OpArg arg, OpDescription op) => (arg.Kind, arg.Required) switch
    {
        (ArgKind.String, true) => $"Str(args, {Quote(arg.Name)})",
        (ArgKind.String, false) => $"OptStr(args, {Quote(arg.Name)})",
        (ArgKind.Boolean, true) => $"RequiredBool(args, {Quote(arg.Name)})",
        (ArgKind.Boolean, false) => $"OptBool(args, {Quote(arg.Name)})",
        (ArgKind.CharSpan, false) => $"ParseSpan(args, {Quote(arg.Name)})",
        _ => throw new CodegenException($"{op.Name}: the MCP server cannot read a required {arg.Type} argument ({arg.Name}) yet"),
    };

    /// <summary>The statements that check <paramref name="args"/> the way the facade call will read them.</summary>
    private static IEnumerable<string> Validation(IEnumerable<OpArg> args)
    {
        var list = args.ToList();
        var strings = list.Where(a => a.Required && a.Kind == ArgKind.String).Select(a => Quote(a.Name)).ToList();
        if (strings.Count > 0) yield return $"RequireStrings(args, {string.Join(", ", strings)});";
        foreach (var arg in list.Where(a => a.Required && a.Kind == ArgKind.Boolean))
            yield return $"_ = RequiredBool(args, {Quote(arg.Name)});";
        foreach (var arg in list.Where(a => !a.Required))
        {
            yield return arg.Kind switch
            {
                ArgKind.String => $"ValidateOptionalString(args, {Quote(arg.Name)});",
                ArgKind.Boolean => $"ValidateOptionalBool(args, {Quote(arg.Name)});",
                _ => $"ValidateOptionalSpan(args, {Quote(arg.Name)});",
            };
        }
    }

    /// <summary>Batch-step validation for a shared action: exactly one selector, then the shared
    /// arguments, then no other op's optional arguments, then this op's own optional arguments.</summary>
    private static void EmitSharedValidation(StringBuilder sb, SharedName group, OpDescription first)
    {
        var label = $"{first.McpTool} {first.McpAction}";
        sb.Append($"    /// <summary>Check a <c>{label}</c> batch step's arguments without running it.</summary>\n");
        sb.Append($"    private static void ValidateGenerated{SharedMethod(first)}(JsonElement args)\n    {{\n");
        foreach (var op in group.Ops)
        {
            var selector = group.Selector(op).Name;
            sb.Append($"        var {Local(selector)} = OptionalStringValue(args, {Quote(selector)});\n");
        }

        sb.Append($"        if ({group.NotExactlyOne()})\n");
        sb.Append($"            throw new McpToolException({Quote($"{label} requires exactly one target: {group.TargetList(withExtras: false)}")});\n");
        foreach (var line in Validation(group.Common)) sb.Append("        ").Append(line).Append('\n');
        foreach (var op in group.Ops)
        {
            var selector = group.Selector(op).Name;
            foreach (var forbidden in group.Forbidden(op))
            {
                sb.Append($"        if ({Local(selector)} is not null && {SharedName.Present(forbidden)})\n");
                sb.Append($"            throw new McpToolException({Quote($"{selector} {ToolNoun(first.McpTool)} targets cannot include {forbidden}")});\n");
            }
        }

        foreach (var op in group.Ops)
        {
            foreach (var line in Validation(group.Extras(op))) sb.Append("        ").Append(line).Append('\n');
        }

        sb.Append("    }\n");
    }

    /// <summary>For each MCP tool the family uses, its generated action enum members and schema
    /// properties, for <c>ToolCatalog</c> to splice into the tool's hand-written input schema.</summary>
    private static void EmitSchema(StringBuilder sb, FamilyDescription family)
    {
        sb.Append($"\n/// <summary>The MCP schema fragments generated from the {family.Family} description.</summary>\n");
        sb.Append("internal static partial class GeneratedMcpSchema\n{\n");
        var tools = family.Ops.Select(o => o.McpTool).Distinct().ToList();
        for (var t = 0; t < tools.Count; t++)
        {
            var tool = tools[t];
            var ops = family.Ops.Where(o => o.McpTool == tool).ToList();
            var args = ops.SelectMany(o => o.Args).GroupBy(a => a.Name).ToDictionary(g => g.Key, g => g.ToList());
            foreach (var (argName, uses) in args)
            {
                if (uses.Select(a => a.Kind).Distinct().Count() > 1)
                    throw new CodegenException($"{family.Family}: {tool} advertises {argName} once, but its actions describe it with different types");
            }

            if (!family.McpProperties.TryGetValue(tool, out var properties))
                throw new CodegenException($"{family.Family}: mcp.{tool}.properties is missing; it carries the schema prose for {tool}'s arguments");
            var missing = args.Keys.Where(a => properties.All(p => p.Name != a)).ToList();
            if (missing.Count > 0)
                throw new CodegenException($"{family.Family}: mcp.{tool}.properties has no description for {string.Join(", ", missing)}");
            var unknown = properties.Where(p => !args.ContainsKey(p.Name)).Select(p => p.Name).ToList();
            if (unknown.Count > 0)
                throw new CodegenException($"{family.Family}: mcp.{tool}.properties describes {string.Join(", ", unknown)}, which no {family.Family} action takes");

            var constant = Pascal(tool);
            var actions = ops.Select(o => o.McpAction).Distinct().Select(a => Quote(a));
            sb.Append($"    /// <summary>The generated <c>{tool}</c> actions, as members of the schema's <c>action</c> enum.</summary>\n");
            sb.Append($"    public const string {constant}Actions = \"\"\"\n        {string.Join(", ", actions)}\n        \"\"\";\n\n");
            sb.Append($"    /// <summary>The generated <c>{tool}</c> argument properties, comma-separated, without a trailing comma.</summary>\n");
            sb.Append($"    public const string {constant}Properties = \"\"\"\n");
            for (var i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                if (property.Description.Contains("\"\"\"", StringComparison.Ordinal))
                    throw new CodegenException($"{family.Family}: the description of {property.Name} cannot contain three double quotes in a row");
                var type = args[property.Name][0].Kind switch
                {
                    ArgKind.String => "\"type\": \"string\"",
                    ArgKind.Boolean => "\"type\": \"boolean\"",
                    _ => "\"type\": \"object\", \"properties\": { \"start\": { \"type\": \"integer\" }, \"length\": { \"type\": \"integer\" } }",
                };

                // Lines after the first carry the four-space indent the property list has in
                // ToolCatalog, so the spliced schema reads as if it were written there.
                sb.Append(i == 0 ? "        " : "            ");
                sb.Append($"{JsonQuote(property.Name)}: {{ {type}, \"description\": {JsonQuote(property.Description)} }}");
                sb.Append(i < properties.Count - 1 ? ",\n" : "\n");
            }

            sb.Append("        \"\"\";\n");
            if (t < tools.Count - 1) sb.Append('\n');
        }

        sb.Append("}\n");
    }
}

/// <summary>The WASM <c>[JSExport]</c> shims, spliced into the bridge's generated region.</summary>
internal static class WasmEmitter
{
    public static string Emit(FamilyDescription family, string source)
    {
        var sb = new StringBuilder();
        foreach (var line in Header(source).TrimEnd('\n').Split('\n')) sb.Append("    ").Append(line).Append('\n');
        foreach (var op in family.Ops)
        {
            var shim = op.Args.Select(a => Shim(a, op)).ToList();
            sb.Append('\n');
            if (op.WasmDoc is { } doc)
            {
                sb.Append("    /// <summary>\n");
                foreach (var line in Wrap(doc, 100 - "    /// ".Length)) sb.Append("    /// ").Append(line).Append('\n');
                sb.Append("    /// </summary>\n");
            }

            sb.Append("    [JSExport]\n");
            var parameters = new[] { "int h" }.Concat(shim.Select(s => s.Parameter)).ToList();
            var signature = $"    public static string {op.WasmMethod}({string.Join(", ", parameters)}) =>";
            if (signature.Length <= 100)
                sb.Append(signature).Append('\n');
            else
                sb.Append($"    public static string {op.WasmMethod}(\n        ").Append(string.Join(",\n        ", parameters)).Append(") =>\n");
            sb.Append("        ").Append(Call(op, "h", shim.Select(s => s.Value), "            ")).Append(";\n");
        }

        return sb.Append('\n').ToString();
    }

    /// <summary>The shim's parameter and the value it passes to the facade. JavaScript passes every
    /// argument positionally, so an optional string with no facade default arrives as <c>""</c>
    /// when absent and becomes null here; one with a default passes through for the facade to apply.</summary>
    private static (string Parameter, string Value) Shim(OpArg arg, OpDescription op)
    {
        var n = Local(arg.Name);
        return (arg.Kind, arg.Required, arg.HasDefault) switch
        {
            (ArgKind.String, true, _) or (ArgKind.String, false, true) => ($"string {n}", n),
            (ArgKind.String, false, false) => ($"string {n}", $"string.IsNullOrEmpty({n}) ? null : {n}"),
            (ArgKind.Boolean, true, _) or (ArgKind.Boolean, false, true) => ($"bool {n}", n),
            (ArgKind.CharSpan, false, _) => ($"string {arg.Name}Json", $"ParseSpan({arg.Name}Json)"),
            _ => throw new CodegenException($"{op.Name}: the WASM bridge cannot pass a {(arg.Required ? "required" : "optional, undefaulted")} {arg.Type} argument ({arg.Name}) yet"),
        };
    }

    private static IEnumerable<string> Wrap(string text, int width)
    {
        var line = new StringBuilder();
        foreach (var word in text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            if (line.Length > 0 && line.Length + 1 + word.Length > width)
            {
                yield return line.ToString();
                line.Clear();
            }

            if (line.Length > 0) line.Append(' ');
            line.Append(word);
        }

        if (line.Length > 0) yield return line.ToString();
    }
}

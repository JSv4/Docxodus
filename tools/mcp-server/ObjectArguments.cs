// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text.Json;
using System.Text.Json.Nodes;

namespace Docxodus.McpServer;

/// <summary>
/// One deprecated top-level MCP property that stands for a field of an op's object argument
/// (issue #1025).
/// </summary>
/// <param name="Field">The field inside the object argument, as every other transport names it.</param>
/// <param name="Property">The flat top-level MCP property that used to carry it.</param>
/// <param name="Kind">The JSON kind the flat property must have (<c>string</c>, <c>boolean</c>,
/// <c>integer</c>, <c>array</c>, <c>object</c>), or null to pass any value through unchecked.</param>
/// <param name="Negated">The flat boolean is the negation of the field
/// (<c>caseSensitive</c> = !<c>options.ignoreCase</c>).</param>
/// <param name="McpDefault">The field's value, as JSON text, when the caller sends neither it nor
/// its alias, where MCP's default differs from the facade's; null to leave the facade's.</param>
internal sealed record FlatAlias(
    string Field, string Property, string? Kind = null, bool Negated = false, string? McpDefault = null);

/// <summary>An op's object argument (<c>options</c>, <c>rule</c>, <c>spec</c>) and the deprecated
/// flat properties that alias its fields.</summary>
internal sealed record ObjectArgument(string Name, IReadOnlyList<FlatAlias> Aliases);

/// <summary>
/// The object arguments the MCP tools take for the ops whose facade, stdio host, WASM and npm
/// surfaces take one options object (issue #1025). MCP takes the same nested object, parsed by the
/// same <c>DocxSessionJson</c> parser, so every field the facade supports is reachable. The flat
/// top-level properties MCP used to spread the object across are kept as deprecated aliases; this
/// table is the one place they are listed, and <see cref="Read"/> is the one place they are merged.
/// No new flat alias should ever be added: a new option is a new field of the object.
/// </summary>
internal static class ObjectArguments
{
    private static readonly ObjectArgument ContentControlOptions = new("options", new FlatAlias[]
    {
        new("bindingPolicy", "bindingPolicy", "string"),
        new("nestedControls", "nestedControls", "string"),
        new("childFills", "childFills", "object"),
    });

    private static readonly ObjectArgument TableInsertOptions = new("options", new FlatAlias[]
    {
        new("borderless", "borderless", "boolean"),
        new("cellAlignment", "cellAlignment", "string"),
        new("cellContents", "cellContents", "array"),
        new("columnWidths", "columnWidths", "array"),
    });

    /// <summary>The reference-table parsers read their fields leniently, as they did when MCP
    /// handed them the whole argument object, so these aliases carry no kind.</summary>
    private static readonly ObjectArgument TableOfContentsOptions = new("options", new FlatAlias[]
    {
        new("levels", "levels"),
        new("hyperlinks", "hyperlinks"),
        new("title", "title"),
        new("rightTabPos", "rightTabPos"),
        new("hideTabAndPageNumbersInWeb", "hideTabAndPageNumbersInWeb"),
        new("useOutlineLevels", "useOutlineLevels"),
    });

    /// <summary>The (tool, action) pairs that take an object argument, with its flat aliases.</summary>
    internal static readonly IReadOnlyDictionary<(string Tool, string Action), ObjectArgument> ByAction =
        new Dictionary<(string, string), ObjectArgument>
        {
            [("docxodus_edit", "replace_text_range")] = new("options", new FlatAlias[]
            {
                // MCP has always matched case-insensitively unless told otherwise (caseSensitive
                // defaulted to false); the facade's ignoreCase defaults to false. Kept.
                new("ignoreCase", "caseSensitive", "boolean", Negated: true, McpDefault: "true"),
            }),
            [("docxodus_create", "insert_table")] = TableInsertOptions,
            [("docxodus_table", "insert")] = TableInsertOptions,
            [("docxodus_create", "insert_horizontal_rule")] = new("rule", new FlatAlias[]
            {
                new("style", "ruleStyle", "string"),
            }),
            [("docxodus_create", "insert_table_of_contents")] = TableOfContentsOptions,
            [("docxodus_create", "insert_table_of_figures")] = new("options", new FlatAlias[]
            {
                new("captionLabel", "captionLabel"),
                new("hyperlinks", "hyperlinks"),
                new("rightTabPos", "rightTabPos"),
            }),
            [("docxodus_create", "insert_table_of_authorities")] = new("options", new FlatAlias[]
            {
                new("category", "category"),
                new("hyperlinks", "hyperlinks"),
                new("entryPageSeparator", "entryPageSeparator"),
                new("rightTabPos", "rightTabPos"),
            }),
            [("docxodus_links", "insert_cross_reference")] = new("options", new FlatAlias[]
            {
                new("referenceNumber", "referenceNumber", "boolean"),
                new("hyperlink", "hyperlink", "boolean"),
                new("includePosition", "includePosition", "boolean"),
            }),
            [("docxodus_table", "set_borders")] = new("spec", new FlatAlias[]
            {
                new("scope", "borderScope", "string"),
                new("style", "borderStyle", "string"),
                new("size", "borderSize", "integer"),
                new("color", "borderColor", "string"),
            }),
            [("docxodus_content_controls", "fill_text")] = ContentControlOptions,
            [("docxodus_content_controls", "fill_rich_text")] = ContentControlOptions,
            [("docxodus_content_controls", "fill_picture")] = ContentControlOptions,
            [("docxodus_content_controls", "select_item")] = ContentControlOptions,
            [("docxodus_content_controls", "set_checked")] = ContentControlOptions,
            [("docxodus_content_controls", "set_date")] = ContentControlOptions,
            [("docxodus_content_controls", "add_repeating_item")] = ContentControlOptions,
        };

    /// <summary>
    /// The object argument of <paramref name="tool"/>/<paramref name="action"/>: the nested object
    /// (absent or null reads as empty), with each deprecated flat alias the caller sent folded in.
    /// A flat alias and its nested field that both appear must agree, as with
    /// <c>DocxSessionJson.AliasedArgument</c>; otherwise the call is refused rather than one of the
    /// two silently winning. Field values are not otherwise checked here: the shared parser the
    /// facade applies decides what a field means, exactly as on every other transport.
    /// </summary>
    internal static JsonElement Read(JsonElement args, string tool, string action)
    {
        var argument = ByAction[(tool, action)];
        var merged = new JsonObject();
        if (args.ValueKind == JsonValueKind.Object
            && args.TryGetProperty(argument.Name, out var nested)
            && nested.ValueKind != JsonValueKind.Null)
        {
            if (nested.ValueKind != JsonValueKind.Object)
                throw new McpToolException($"argument \"{argument.Name}\" must be an object");
            merged = JsonNode.Parse(nested.GetRawText())!.AsObject();
        }

        foreach (var alias in argument.Aliases)
        {
            if (args.ValueKind != JsonValueKind.Object || !args.TryGetProperty(alias.Property, out var flat))
            {
                if (alias.McpDefault is not null && !merged.ContainsKey(alias.Field))
                    merged[alias.Field] = JsonNode.Parse(alias.McpDefault);
                continue;
            }
            CheckKind(alias, flat);
            var value = alias.Negated ? JsonValue.Create(!flat.GetBoolean()) : JsonNode.Parse(flat.GetRawText());
            if (merged.TryGetPropertyValue(alias.Field, out var existing))
            {
                if (!JsonNode.DeepEquals(existing, value))
                    throw new McpToolException(
                        $"\"{argument.Name}.{alias.Field}\" and its deprecated alias \"{alias.Property}\" "
                        + $"disagree; pass only \"{argument.Name}.{alias.Field}\"");
                continue;
            }
            merged[alias.Field] = value;
        }
        return ToElement(merged);
    }

    /// <summary>The raw JSON text of <see cref="Read"/>, for facade methods that take the object as text.</summary>
    internal static string ReadJson(JsonElement args, string tool, string action) =>
        Read(args, tool, action).GetRawText();

    private static void CheckKind(FlatAlias alias, JsonElement flat)
    {
        var ok = alias.Kind switch
        {
            null => true,
            "string" => flat.ValueKind == JsonValueKind.String,
            "boolean" => flat.ValueKind is JsonValueKind.True or JsonValueKind.False,
            "integer" => flat.ValueKind == JsonValueKind.Number && flat.TryGetInt32(out _),
            "array" => flat.ValueKind == JsonValueKind.Array,
            "object" => flat.ValueKind == JsonValueKind.Object,
            _ => throw new InvalidOperationException($"unknown alias kind {alias.Kind}"),
        };
        if (!ok)
            throw new McpToolException($"argument \"{alias.Property}\" must be a{(alias.Kind is "integer" or "array" or "object" ? "n" : "")} {alias.Kind}");
    }

    private static JsonElement ToElement(JsonObject value)
    {
        using var document = JsonDocument.Parse(value.ToJsonString());
        return document.RootElement.Clone();
    }
}

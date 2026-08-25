// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Json;
using System.Text.RegularExpressions;

namespace Docxodus.Tests.Eval;

/// <summary>
/// A deliberately small JSON-Schema interpreter covering exactly the constructs
/// <c>eval/scenario.schema.json</c> uses: <c>$ref</c> into <c>$defs</c>, <c>type</c>,
/// <c>enum</c>, <c>required</c>, <c>properties</c>, <c>additionalProperties</c> (false or a
/// schema), <c>items</c>, <c>minItems</c>, <c>minLength</c>, <c>minProperties</c>,
/// <c>minimum</c>, and <c>pattern</c>. The suite has no JSON-Schema package dependency, and a
/// generic validator is not the point — the point is that a scenario the schema rejects must
/// also be rejected where the tests run, so the two contracts cannot drift silently.
/// </summary>
internal static class EvalSchemaValidator
{
    /// <summary>Validate one instance against the schema, returning human-readable errors.</summary>
    public static IReadOnlyList<string> Validate(JsonElement instance, JsonElement schema)
    {
        var errors = new List<string>();
        Check(instance, schema, schema, "$", errors);
        return errors;
    }

    private static void Check(
        JsonElement instance, JsonElement schema, JsonElement root, string path, List<string> errors)
    {
        if (schema.TryGetProperty("$ref", out var reference))
        {
            Check(instance, ResolveRef(root, reference.GetString()!), root, path, errors);
            return;
        }

        if (schema.TryGetProperty("type", out var type)
            && !TypeMatches(instance, type.GetString()!))
        {
            errors.Add($"{path}: expected {type.GetString()}, found {instance.ValueKind}");
            return;
        }

        if (schema.TryGetProperty("enum", out var allowed)
            && !allowed.EnumerateArray().Any(candidate => JsonEquals(candidate, instance)))
            errors.Add($"{path}: value {instance.GetRawText()} is not one of {allowed.GetRawText()}");

        if (instance.ValueKind == JsonValueKind.String)
        {
            var value = instance.GetString()!;
            if (schema.TryGetProperty("minLength", out var minLength)
                && value.Length < minLength.GetInt32())
                errors.Add($"{path}: string shorter than minLength {minLength.GetInt32()}");
            if (schema.TryGetProperty("pattern", out var pattern)
                && !Regex.IsMatch(value, pattern.GetString()!))
                errors.Add($"{path}: \"{value}\" does not match pattern {pattern.GetString()}");
        }

        if (instance.ValueKind == JsonValueKind.Number
            && schema.TryGetProperty("minimum", out var minimum)
            && instance.GetDouble() < minimum.GetDouble())
            errors.Add($"{path}: {instance.GetRawText()} is below minimum {minimum.GetRawText()}");

        if (instance.ValueKind == JsonValueKind.Array)
        {
            var count = instance.GetArrayLength();
            if (schema.TryGetProperty("minItems", out var minItems)
                && count < minItems.GetInt32())
                errors.Add($"{path}: {count} item(s), minItems is {minItems.GetInt32()}");
            if (schema.TryGetProperty("items", out var items))
            {
                var index = 0;
                foreach (var item in instance.EnumerateArray())
                    Check(item, items, root, $"{path}[{index++}]", errors);
            }
        }

        if (instance.ValueKind == JsonValueKind.Object)
            CheckObject(instance, schema, root, path, errors);
    }

    private static void CheckObject(
        JsonElement instance, JsonElement schema, JsonElement root, string path, List<string> errors)
    {
        if (schema.TryGetProperty("minProperties", out var minProperties)
            && instance.EnumerateObject().Count() < minProperties.GetInt32())
            errors.Add($"{path}: fewer than minProperties {minProperties.GetInt32()} member(s)");

        if (schema.TryGetProperty("required", out var required))
        {
            foreach (var name in required.EnumerateArray())
            {
                if (!instance.TryGetProperty(name.GetString()!, out _))
                    errors.Add($"{path}: missing required property '{name.GetString()}'");
            }
        }

        var hasProperties = schema.TryGetProperty("properties", out var properties);
        var hasAdditional = schema.TryGetProperty("additionalProperties", out var additional);
        foreach (var member in instance.EnumerateObject())
        {
            if (hasProperties && properties.TryGetProperty(member.Name, out var memberSchema))
            {
                Check(member.Value, memberSchema, root, $"{path}.{member.Name}", errors);
            }
            else if (hasAdditional)
            {
                if (additional.ValueKind == JsonValueKind.False)
                    errors.Add($"{path}: property '{member.Name}' is not allowed");
                else
                    Check(member.Value, additional, root, $"{path}.{member.Name}", errors);
            }
        }
    }

    private static JsonElement ResolveRef(JsonElement root, string reference)
    {
        if (!reference.StartsWith("#/", StringComparison.Ordinal))
            throw new InvalidOperationException($"unsupported $ref '{reference}'");
        var current = root;
        foreach (var segment in reference[2..].Split('/'))
            current = current.GetProperty(segment);
        return current;
    }

    private static bool TypeMatches(JsonElement instance, string type) => type switch
    {
        "object" => instance.ValueKind == JsonValueKind.Object,
        "array" => instance.ValueKind == JsonValueKind.Array,
        "string" => instance.ValueKind == JsonValueKind.String,
        "integer" => instance.ValueKind == JsonValueKind.Number
            && instance.TryGetInt64(out _),
        "number" => instance.ValueKind == JsonValueKind.Number,
        "boolean" => instance.ValueKind is JsonValueKind.True or JsonValueKind.False,
        _ => throw new InvalidOperationException($"unsupported schema type '{type}'"),
    };

    private static bool JsonEquals(JsonElement left, JsonElement right) =>
        left.ValueKind == right.ValueKind && left.GetRawText() == right.GetRawText();
}

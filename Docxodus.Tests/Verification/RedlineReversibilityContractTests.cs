// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus.Internal;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests.Verification;

/// <summary>
/// Pins the emitted redline-reversibility proof JSON to the published v1 schema
/// (<c>docs/schemas/redline-reversibility-proof-v1.schema.json</c>), the way
/// <see cref="SemanticChangeSetContractTests"/> pins the semantic-changes wire shape. The
/// validator below implements exactly the draft 2020-12 subset the checked-in schema uses and
/// fails closed on any keyword it does not know, so evolving the schema past that subset breaks
/// this test loudly instead of silently validating nothing.
/// </summary>
public class RedlineReversibilityContractTests
{
    private static readonly string SchemaPath = Path.GetFullPath(Path.Combine(
        AppContext.BaseDirectory,
        "../../../../docs/schemas/redline-reversibility-proof-v1.schema.json"));

    [Fact]
    public void Emitted_proof_json_conforms_to_the_published_v1_schema()
    {
        using var schema = JsonDocument.Parse(File.ReadAllBytes(SchemaPath));
        Assert.Equal(RedlineReversibilityProof.SchemaId,
            schema.RootElement.GetProperty("$id").GetString());
        Assert.Equal(RedlineReversibilityProof.SchemaId,
            schema.RootElement.GetProperty("properties")
                .GetProperty("schema").GetProperty("const").GetString());

        // A completed two-path proof: classifications, both paths, divergences and findings
        // are all populated, so every referenced $defs shape is exercised.
        var baseline = Document("The original clause.");
        var intendedFinal = Document("The revised clause.");
        var redline = EngineRedline(baseline, intendedFinal);
        var completed = RedlineReversibilityVerifier.Prove(baseline, intendedFinal, redline).Proof;
        Assert.NotEmpty(completed.RevisionClassifications);
        Assert.NotNull(completed.AcceptToFinal);
        using (var instance = JsonDocument.Parse(completed.ToCanonicalUtf8Bytes()))
            Assert.Empty(ValidationErrors(schema.RootElement, instance.RootElement));

        // A fail-closed proof (malformed redline package): both paths are null, exercising
        // the schema's nullablePath branches.
        var failed = RedlineReversibilityVerifier.Prove(
            baseline, baseline, new byte[] { 1, 2, 3 }).Proof;
        Assert.Null(failed.AcceptToFinal);
        using (var instance = JsonDocument.Parse(failed.ToCanonicalUtf8Bytes()))
            Assert.Empty(ValidationErrors(schema.RootElement, instance.RootElement));

        // The shared facade every transport routes through emits the same conformant shape.
        var wireJson = VerificationOps.ProveRedlineReversibility(baseline, intendedFinal, redline);
        using (var instance = JsonDocument.Parse(wireJson))
            Assert.Empty(ValidationErrors(schema.RootElement, instance.RootElement));
    }

    // Negative control: the validator must actually reject nonconforming documents, or the
    // conformance assertion above proves nothing. One mutation per constraint kind the
    // schema relies on.
    [Fact]
    public void Schema_validation_rejects_each_constraint_violation_kind()
    {
        using var schema = JsonDocument.Parse(File.ReadAllBytes(SchemaPath));
        var baseline = Document("The original clause.");
        var intendedFinal = Document("The revised clause.");
        var redline = EngineRedline(baseline, intendedFinal);
        var canonicalJson = RedlineReversibilityVerifier
            .Prove(baseline, intendedFinal, redline).Proof.ToCanonicalJson();

        // The unmutated document validates; every mutation below must not.
        using (var wellFormed = JsonDocument.Parse(canonicalJson))
            Assert.Empty(ValidationErrors(schema.RootElement, wellFormed.RootElement));

        AssertRejected(schema, canonicalJson, // required
            node => node.AsObject().Remove("success"));
        AssertRejected(schema, canonicalJson, // additionalProperties: false
            node => node.AsObject()["unexpected"] = true);
        AssertRejected(schema, canonicalJson, // const
            node => node.AsObject()["schemaVersion"] = 2);
        AssertRejected(schema, canonicalJson, // type
            node => node.AsObject()["success"] = "yes");
        AssertRejected(schema, canonicalJson, // pattern
            node => node["baselinePackage"]!["rawPackageBytesDigest"]!["value"] = "not-a-digest");
        AssertRejected(schema, canonicalJson, // enum
            node => node["revisionClassifications"]![0]!["disposition"] = "bogus");
        AssertRejected(schema, canonicalJson, // oneOf (neither object-shaped nor null)
            node => node.AsObject()["acceptToFinal"] = 42);
    }

    private static void AssertRejected(
        JsonDocument schema, string canonicalJson, Action<JsonNode> mutate)
    {
        var node = JsonNode.Parse(canonicalJson)!;
        mutate(node);
        using var mutated = JsonDocument.Parse(node.ToJsonString());
        Assert.NotEmpty(ValidationErrors(schema.RootElement, mutated.RootElement));
    }

    // ------------------------------------------------------------------
    // Minimal JSON-schema validator: exactly the draft 2020-12 subset the checked-in proof
    // schema uses, failing closed on anything else.
    // ------------------------------------------------------------------

    private static readonly string[] KnownKeywords =
    {
        "$schema", "$id", "title", "description", "$defs",
        "type", "required", "properties", "additionalProperties",
        "enum", "const", "oneOf", "items", "$ref", "pattern", "minLength", "minimum",
    };

    private static IReadOnlyList<string> ValidationErrors(
        JsonElement schemaRoot, JsonElement instance)
    {
        var errors = new List<string>();
        Validate(schemaRoot, instance, schemaRoot, "$", errors);
        return errors;
    }

    private static void Validate(
        JsonElement schema, JsonElement instance, JsonElement root, string path,
        List<string> errors)
    {
        foreach (var keyword in schema.EnumerateObject())
            Assert.Contains(keyword.Name, KnownKeywords);

        if (schema.TryGetProperty("$ref", out var reference))
        {
            Validate(ResolveRef(root, reference.GetString()!), instance, root, path, errors);
            return;
        }

        if (schema.TryGetProperty("oneOf", out var oneOf))
        {
            var matches = oneOf.EnumerateArray().Count(branch =>
            {
                var branchErrors = new List<string>();
                Validate(branch, instance, root, path, branchErrors);
                return branchErrors.Count == 0;
            });
            if (matches != 1)
                errors.Add($"{path}: matched {matches} oneOf branches, expected exactly 1");
            return; // The schema never combines oneOf with sibling constraints.
        }

        if (schema.TryGetProperty("const", out var constant)
            && !JsonElement.DeepEquals(constant, instance))
            errors.Add($"{path}: does not equal const {constant.GetRawText()}");

        if (schema.TryGetProperty("enum", out var enumeration)
            && !enumeration.EnumerateArray().Any(option => JsonElement.DeepEquals(option, instance)))
            errors.Add($"{path}: {instance.GetRawText()} is not in the enum");

        if (schema.TryGetProperty("type", out var type) && !TypeMatches(type, instance))
            errors.Add($"{path}: {instance.ValueKind} does not satisfy type {type.GetRawText()}");

        if (instance.ValueKind == JsonValueKind.String)
        {
            var value = instance.GetString()!;
            if (schema.TryGetProperty("pattern", out var pattern)
                && !Regex.IsMatch(value, pattern.GetString()!))
                errors.Add($"{path}: '{value}' does not match pattern {pattern.GetString()}");
            if (schema.TryGetProperty("minLength", out var minLength)
                && value.Length < minLength.GetInt32())
                errors.Add($"{path}: shorter than minLength {minLength.GetInt32()}");
        }

        if (instance.ValueKind == JsonValueKind.Number
            && schema.TryGetProperty("minimum", out var minimum)
            && instance.GetDouble() < minimum.GetDouble())
            errors.Add($"{path}: below minimum {minimum.GetRawText()}");

        if (instance.ValueKind == JsonValueKind.Object)
        {
            var hasProperties = schema.TryGetProperty("properties", out var properties);
            if (schema.TryGetProperty("required", out var required))
            {
                foreach (var name in required.EnumerateArray())
                {
                    if (!instance.TryGetProperty(name.GetString()!, out _))
                        errors.Add($"{path}: missing required property '{name.GetString()}'");
                }
            }

            foreach (var property in instance.EnumerateObject())
            {
                if (hasProperties
                    && properties.TryGetProperty(property.Name, out var propertySchema))
                {
                    Validate(propertySchema, property.Value, root,
                        $"{path}.{property.Name}", errors);
                }
                else if (schema.TryGetProperty("additionalProperties", out var additional)
                    && additional.ValueKind == JsonValueKind.False)
                {
                    errors.Add($"{path}: unexpected property '{property.Name}'");
                }
            }
        }

        if (instance.ValueKind == JsonValueKind.Array
            && schema.TryGetProperty("items", out var items))
        {
            var index = 0;
            foreach (var element in instance.EnumerateArray())
                Validate(items, element, root, $"{path}[{index++}]", errors);
        }
    }

    private static JsonElement ResolveRef(JsonElement root, string reference)
    {
        Assert.StartsWith("#/$defs/", reference);
        return root.GetProperty("$defs").GetProperty(reference["#/$defs/".Length..]);
    }

    private static bool TypeMatches(JsonElement type, JsonElement instance) =>
        type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Any(entry => TypeNameMatches(entry.GetString()!, instance))
            : TypeNameMatches(type.GetString()!, instance);

    private static bool TypeNameMatches(string name, JsonElement instance) => name switch
    {
        "object" => instance.ValueKind == JsonValueKind.Object,
        "array" => instance.ValueKind == JsonValueKind.Array,
        "string" => instance.ValueKind == JsonValueKind.String,
        "boolean" => instance.ValueKind is JsonValueKind.True or JsonValueKind.False,
        "null" => instance.ValueKind == JsonValueKind.Null,
        "integer" => instance.ValueKind == JsonValueKind.Number && instance.TryGetInt64(out _),
        "number" => instance.ValueKind == JsonValueKind.Number,
        _ => throw new InvalidOperationException($"Unsupported schema type '{name}'."),
    };

    // ------------------------------------------------------------------
    // Fixtures.
    // ------------------------------------------------------------------

    private static byte[] EngineRedline(byte[] baseline, byte[] intendedFinal) =>
        DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("final.docx", intendedFinal),
            new DocxDiffSettings { AuthorForRevisions = "Comparison Engine" }).DocumentByteArray;

    private static byte[] Document(string text)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(new Run(new Text(text)))));
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            document.Save();
        }

        return stream.ToArray();
    }
}

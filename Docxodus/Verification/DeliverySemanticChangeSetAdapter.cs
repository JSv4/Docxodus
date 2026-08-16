// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System.Text.Encodings.Web;
using System.Text.Json;

namespace Docxodus.Verification;

internal sealed record DeliverySemanticChangeSetProjection(
    string Schema,
    int SchemaVersion,
    int ChangeCount,
    byte[] CanonicalBytes,
    VerificationDigest Digest);

/// <summary>
/// The sole #457 integration seam. It reconstructs the complete closed typed contract and accepts
/// bytes only when they are byte-for-byte identical to <see cref="SemanticChangeSet.ToCanonicalUtf8Bytes"/>.
/// </summary>
internal static class DeliverySemanticChangeSetAdapter
{
    public static DeliverySemanticChangeSetProjection Project(
        SemanticChangeSet changeSet,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(changeSet);
        ArgumentNullException.ThrowIfNull(limits);
        ValidateTyped(changeSet, limits);
        var bytes = SerializeExactBounded(changeSet, limits);
        var inspected = InspectExact(bytes, limits);
        if (!string.Equals(inspected.Schema, changeSet.Schema, StringComparison.Ordinal)
            || inspected.SchemaVersion != changeSet.SchemaVersion
            || inspected.ChangeCount != changeSet.ChangeCount)
        {
            throw Invalid("Semantic change-set canonical bytes are inconsistent.");
        }
        return inspected;
    }

    public static DeliverySemanticChangeSetProjection InspectExact(
        ReadOnlySpan<byte> bytes,
        DeliveryReceiptLimits limits)
    {
        ArgumentNullException.ThrowIfNull(limits);
        DeliveryReceiptResourceBudget.Bytes(
            bytes.Length,
            limits.MaxSemanticEvidenceBytes,
            "semantic_resource_limit",
            "SemanticChangeSet artifact");

        // Strict UTF-8/JSON/depth/duplicate-property gate. The generic canonical bytes are not
        // used as #457's digest or representation because #457 owns a different fixed field order.
        _ = DeliveryReceiptCanonicalJson.CanonicalizeBounded(
            bytes,
            limits,
            limits.MaxSemanticEvidenceBytes,
            "semantic_resource_limit");
        using var document = JsonDocument.Parse(bytes.ToArray(), new JsonDocumentOptions
        {
            AllowTrailingCommas = false,
            CommentHandling = JsonCommentHandling.Disallow,
            MaxDepth = limits.MaxJsonDepth,
        });
        var budget = new DeliveryReceiptResourceBudget(
            limits, limits.MaxSemanticEvidenceBytes, "semantic_resource_limit");
        budget.AddSerializedBytes(64, "semantic root");
        var root = document.RootElement;
        ExactObject(root, budget, "schema", "schemaVersion", "changeCount", "changes");
        var schema = RequiredString(root, "schema", budget);
        if (!string.Equals(schema, SemanticChangeSet.CurrentSchema, StringComparison.Ordinal)
            || !root.GetProperty("schemaVersion").TryGetInt32(out var schemaVersion)
            || schemaVersion != SemanticChangeSet.CurrentSchemaVersion
            || !root.GetProperty("changeCount").TryGetInt32(out var declaredCount)
            || declaredCount < 0)
        {
            throw Invalid("Semantic change-set root identity is invalid.");
        }

        var changesElement = root.GetProperty("changes");
        if (changesElement.ValueKind != JsonValueKind.Array)
            throw Invalid("Semantic change-set changes must be an array.");
        var changes = new List<SemanticChange>();
        foreach (var change in changesElement.EnumerateArray())
        {
            budget.AddItems(1, "semantic changes");
            changes.Add(ParseChange(change, budget));
        }
        if (changes.Count != declaredCount)
            throw Invalid("Semantic change-set count does not match its changes array.");
        var typed = new SemanticChangeSet(changes);
        var canonicalBytes = SerializeExactBounded(typed, limits);
        if (!bytes.SequenceEqual(canonicalBytes))
        {
            throw Invalid(
                "Semantic artifact bytes are not the exact canonical #457 representation.");
        }
        return new DeliverySemanticChangeSetProjection(
            typed.Schema,
            typed.SchemaVersion,
            typed.ChangeCount,
            canonicalBytes,
            DeliveryReceiptCanonicalJson.Digest(canonicalBytes));
    }

    private static SemanticChange ParseChange(
        JsonElement element,
        DeliveryReceiptResourceBudget budget)
    {
        ExactObject(
            element,
            budget,
            "id",
            "operation",
            "family",
            "partUri",
            "path",
            "leftAnchor",
            "rightAnchor",
            "leftScope",
            "rightScope",
            "moveId",
            "before",
            "after");
        budget.AddSerializedBytes(192, "semantic change");
        var id = RequiredString(element, "id", budget);
        if (!ValidChangeId(id))
            throw Invalid("Semantic change id is invalid.");
        var operationName = RequiredString(element, "operation", budget);
        var familyName = RequiredString(element, "family", budget);
        var partUri = RequiredString(element, "partUri", budget);
        if (!partUri.StartsWith("/", StringComparison.Ordinal))
            throw Invalid("Semantic change partUri must be package-absolute.");
        return new SemanticChange
        {
            Id = id,
            Operation = ParseOperation(operationName),
            Family = ParseFamily(familyName),
            PartUri = partUri,
            Path = RequiredString(element, "path", budget, allowEmpty: true),
            LeftAnchor = NullableString(element, "leftAnchor", budget),
            RightAnchor = NullableString(element, "rightAnchor", budget),
            LeftScope = NullableString(element, "leftScope", budget),
            RightScope = NullableString(element, "rightScope", budget),
            MoveId = NullableString(element, "moveId", budget),
            Before = ParseValue(element.GetProperty("before"), budget, depth: 4),
            After = ParseValue(element.GetProperty("after"), budget, depth: 4),
        };
    }

    private static SemanticValue ParseValue(
        JsonElement element,
        DeliveryReceiptResourceBudget budget,
        int depth)
    {
        budget.Depth(depth, "semantic value");
        budget.AddSerializedBytes(24, "semantic value");
        if (element.ValueKind != JsonValueKind.Object)
            throw Invalid("Semantic values must be non-null objects.");
        var kindName = RequiredString(element, "kind", budget);
        switch (kindName)
        {
            case "absent":
                ExactObject(element, budget, "kind");
                return SemanticValue.Absent;
            case "string":
                ExactObject(element, budget, "kind", "value");
                return SemanticValue.String(RequiredString(
                    element, "value", budget, allowEmpty: true));
            case "boolean":
                ExactObject(element, budget, "kind", "value");
                if (element.GetProperty("value").ValueKind is not
                    (JsonValueKind.True or JsonValueKind.False))
                {
                    throw Invalid("Semantic Boolean value is invalid.");
                }
                return SemanticValue.Boolean(element.GetProperty("value").GetBoolean());
            case "integer":
                ExactObject(element, budget, "kind", "value");
                if (!element.GetProperty("value").TryGetInt64(out var integer))
                    throw Invalid("Semantic integer value is outside Int64.");
                return SemanticValue.Integer(integer);
            case "digest":
                ExactObject(element, budget, "kind", "algorithm", "profile", "value");
                var algorithm = RequiredString(element, "algorithm", budget);
                var digestValue = RequiredString(element, "value", budget);
                var profile = NullableString(element, "profile", budget);
                if (profile is not null && string.IsNullOrWhiteSpace(profile))
                    throw Invalid("Semantic digest profile cannot be blank.");
                return SemanticValue.Digest(algorithm, digestValue, profile);
            case "object":
                ExactObject(element, budget, "kind", "value");
                var objectElement = element.GetProperty("value");
                if (objectElement.ValueKind != JsonValueKind.Object)
                    throw Invalid("Semantic object value must be an object.");
                var properties = new List<SemanticProperty>();
                foreach (var property in objectElement.EnumerateObject())
                {
                    budget.AddItems(1, "semantic object properties");
                    budget.String(property.Name, "semantic property name");
                    if (string.IsNullOrWhiteSpace(property.Name))
                        throw Invalid("Semantic object property names must be non-blank.");
                    properties.Add(new SemanticProperty(
                        property.Name, ParseValue(property.Value, budget, depth + 2)));
                }
                return SemanticValue.Object(properties);
            case "array":
                ExactObject(element, budget, "kind", "value");
                var arrayElement = element.GetProperty("value");
                if (arrayElement.ValueKind != JsonValueKind.Array)
                    throw Invalid("Semantic array value must be an array.");
                var items = new List<SemanticValue>();
                foreach (var item in arrayElement.EnumerateArray())
                {
                    budget.AddItems(1, "semantic array items");
                    items.Add(ParseValue(item, budget, depth + 2));
                }
                return SemanticValue.Array(items);
            default:
                throw Invalid($"Unknown semantic value kind '{kindName}'.");
        }
    }

    private static void ValidateTyped(
        SemanticChangeSet changeSet,
        DeliveryReceiptLimits limits)
    {
        if (!string.Equals(changeSet.Schema, SemanticChangeSet.CurrentSchema,
                StringComparison.Ordinal)
            || changeSet.SchemaVersion != SemanticChangeSet.CurrentSchemaVersion
            || changeSet.ChangeCount != changeSet.Changes.Count)
        {
            throw Invalid("Semantic change-set identity is inconsistent.");
        }
        var budget = new DeliveryReceiptResourceBudget(
            limits, limits.MaxSemanticEvidenceBytes, "semantic_resource_limit");
        budget.AddSerializedBytes(64, "semantic root");
        budget.AddItems(changeSet.Changes.Count, "semantic changes");
        foreach (var change in changeSet.Changes)
        {
            if (change is null || change.Before is null || change.After is null
                || !Enum.IsDefined(change.Operation) || !Enum.IsDefined(change.Family)
                || string.IsNullOrWhiteSpace(change.PartUri)
                || !change.PartUri.StartsWith("/", StringComparison.Ordinal)
                || change.Path is null)
            {
                throw Invalid("Semantic change contains invalid typed fields.");
            }
            budget.String(change.Id, "semantic change id");
            budget.String(change.PartUri, "semantic part URI");
            budget.String(change.Path, "semantic path");
            budget.String(change.LeftAnchor, "semantic anchor");
            budget.String(change.RightAnchor, "semantic anchor");
            budget.String(change.LeftScope, "semantic scope");
            budget.String(change.RightScope, "semantic scope");
            budget.String(change.MoveId, "semantic move id");
            budget.AddSerializedBytes(192, "semantic change");
            ValidateTypedValue(change.Before, budget, depth: 4);
            ValidateTypedValue(change.After, budget, depth: 4);
        }
    }

    private static void ValidateTypedValue(
        SemanticValue value,
        DeliveryReceiptResourceBudget budget,
        int depth)
    {
        ArgumentNullException.ThrowIfNull(value);
        budget.Depth(depth, "semantic value");
        budget.AddSerializedBytes(24, "semantic value");
        switch (value.Kind)
        {
            case SemanticValueKind.Absent:
                break;
            case SemanticValueKind.String when value.StringValue is not null:
                budget.String(value.StringValue, "semantic string value");
                break;
            case SemanticValueKind.Boolean when value.BooleanValue is not null:
            case SemanticValueKind.Integer when value.IntegerValue is not null:
                break;
            case SemanticValueKind.Digest
                when !string.IsNullOrWhiteSpace(value.DigestAlgorithm)
                    && !string.IsNullOrWhiteSpace(value.DigestValue)
                    && (value.DigestProfile is null
                        || !string.IsNullOrWhiteSpace(value.DigestProfile)):
                budget.String(value.DigestAlgorithm, "semantic digest algorithm");
                budget.String(value.DigestProfile, "semantic digest profile");
                budget.String(value.DigestValue, "semantic digest value");
                break;
            case SemanticValueKind.Object:
                budget.AddItems(value.Properties.Count, "semantic object properties");
                string? priorName = null;
                foreach (var property in value.Properties)
                {
                    if (property is null || string.IsNullOrWhiteSpace(property.Name)
                        || property.Value is null
                        || (priorName is not null && string.CompareOrdinal(
                            priorName, property.Name) >= 0))
                    {
                        throw Invalid("Semantic object properties are invalid or noncanonical.");
                    }
                    budget.String(property.Name, "semantic property name");
                    ValidateTypedValue(property.Value, budget, depth + 2);
                    priorName = property.Name;
                }
                break;
            case SemanticValueKind.Array:
                budget.AddItems(value.Items.Count, "semantic array items");
                foreach (var item in value.Items)
                {
                    if (item is null)
                        throw Invalid("Semantic arrays cannot contain null.");
                    ValidateTypedValue(item, budget, depth + 2);
                }
                break;
            default:
                throw Invalid("Semantic value fields do not match their kind.");
        }
    }

    private static void ExactObject(
        JsonElement element,
        DeliveryReceiptResourceBudget budget,
        params string[] names)
    {
        if (element.ValueKind != JsonValueKind.Object)
            throw Invalid("Expected a semantic JSON object.");
        int count = 0;
        foreach (var property in element.EnumerateObject())
        {
            budget.AddItems(1, "semantic object fields");
            count++;
            if (count > names.Length
                || !names.Contains(property.Name, StringComparer.Ordinal))
            {
                throw Invalid("Semantic JSON object fields are missing or unknown.");
            }
        }
        if (count != names.Length)
            throw Invalid("Semantic JSON object fields are missing or unknown.");
    }

    private static string RequiredString(
        JsonElement owner,
        string name,
        DeliveryReceiptResourceBudget budget,
        bool allowEmpty = false)
    {
        if (!owner.TryGetProperty(name, out var element)
            || element.ValueKind != JsonValueKind.String)
        {
            throw Invalid($"Semantic field '{name}' must be a string.");
        }
        var value = element.GetString()!;
        budget.String(value, $"semantic field '{name}'");
        if (!allowEmpty && string.IsNullOrWhiteSpace(value))
            throw Invalid($"Semantic field '{name}' cannot be blank.");
        return value;
    }

    private static string? NullableString(
        JsonElement owner,
        string name,
        DeliveryReceiptResourceBudget budget)
    {
        var element = owner.GetProperty(name);
        if (element.ValueKind == JsonValueKind.Null)
            return null;
        if (element.ValueKind != JsonValueKind.String)
            throw Invalid($"Semantic field '{name}' must be a string or null.");
        var value = element.GetString()!;
        budget.String(value, $"semantic field '{name}'");
        return value;
    }

    private static SemanticChangeOperation ParseOperation(string value)
    {
        foreach (var candidate in Enum.GetValues<SemanticChangeOperation>())
        {
            if (string.Equals(
                    SemanticChangeSet.OperationName(candidate), value, StringComparison.Ordinal))
            {
                return candidate;
            }
        }
        throw Invalid($"Unknown semantic operation '{value}'.");
    }

    private static SemanticChangeFamily ParseFamily(string value)
    {
        foreach (var candidate in Enum.GetValues<SemanticChangeFamily>())
        {
            if (string.Equals(
                    SemanticChangeSet.FamilyName(candidate), value, StringComparison.Ordinal))
            {
                return candidate;
            }
        }
        throw Invalid($"Unknown semantic family '{value}'.");
    }

    private static bool ValidChangeId(string value) =>
        value.StartsWith("chg-", StringComparison.Ordinal)
        && value.Length >= 10
        && value.AsSpan(4).IndexOfAnyExceptInRange('0', '9') < 0;

    private static byte[] SerializeExactBounded(
        SemanticChangeSet changeSet,
        DeliveryReceiptLimits limits)
    {
        using var stream = new DeliveryReceiptBoundedMemoryStream(
            limits.MaxSemanticEvidenceBytes,
            "semantic_resource_limit",
            "SemanticChangeSet");
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.Default,
            Indented = false,
            MaxDepth = limits.MaxJsonDepth,
        }))
        {
            changeSet.WriteCanonical(writer);
        }
        return stream.ToArray();
    }

    private static DeliveryReceiptValidationException Invalid(string message) =>
        new("invalid_semantic_change_set", message);
}

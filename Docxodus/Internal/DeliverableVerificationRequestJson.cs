// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using Docxodus.Verification;

namespace Docxodus.Internal;

/// <summary>
/// The one wire form of the full deliverable-verification request (issue #747). Package bytes
/// travel natively on each transport; this object carries what the typed .NET request adds —
/// policy and inspection-limit options, the approved semantic and package deltas, and companion
/// renderer artifacts with their diagnostics. Every transport (WASM/npm, the stdio host and
/// docx-scalpel, MCP) parses the identical shape here and hands the result to the same
/// <see cref="DeliverableVerifier"/>, so equivalent requests produce one canonical report.
/// Defaults are read from <see cref="DeliverableVerificationOptions"/> itself, never restated.
/// Every count and byte limit is enforced on the encoded input before anything is decoded,
/// copied, or inspected, and unknown properties are rejected so a mistyped policy field cannot
/// silently fall back to the default.
/// </summary>
internal static class DeliverableVerificationRequestJson
{
    internal sealed record Parsed(
        DeliverableVerificationOptions Options,
        SemanticChangeSet? ExpectedSemanticChanges,
        IReadOnlyList<DeliverablePackageChangeExpectation> ExpectedPackageChanges,
        IReadOnlyList<DeliverableCompanionArtifactInput> CompanionArtifacts);

    private static readonly JsonDocumentOptions DocumentOptions = new()
    {
        AllowTrailingCommas = false,
        CommentHandling = JsonCommentHandling.Disallow,
        MaxDepth = 64,
    };

    public static Parsed Parse(string? requestJson)
    {
        if (string.IsNullOrWhiteSpace(requestJson))
        {
            return new Parsed(
                new DeliverableVerificationOptions(),
                null,
                Array.Empty<DeliverablePackageChangeExpectation>(),
                Array.Empty<DeliverableCompanionArtifactInput>());
        }

        using var document = JsonDocument.Parse(requestJson, DocumentOptions);
        var root = document.RootElement;
        Require(root.ValueKind == JsonValueKind.Object, "verification request must be a JSON object");
        RejectUnknown(root, "$", "options", "expectedSemanticChanges", "expectedPackageChanges",
            "companionArtifacts");

        var options = ParseOptions(Optional(root, "options"));
        options.Validate();

        // Expected-change budgets are checked against declared counts before any item is parsed.
        int expectedBudget = options.MaxExpectedChanges;
        SemanticChangeSet? semantic = null;
        if (Optional(root, "expectedSemanticChanges") is { } semanticElement)
        {
            Require(semanticElement.ValueKind == JsonValueKind.Object,
                "expectedSemanticChanges must be the canonical semantic-changes object");
            if (semanticElement.TryGetProperty("changeCount", out var declared)
                && declared.TryGetInt32(out var declaredCount))
            {
                Require(declaredCount <= expectedBudget,
                    $"expectedSemanticChanges.changeCount ({declaredCount}) exceeds maxExpectedChanges ({options.MaxExpectedChanges})");
            }
            try
            {
                semantic = DeliverySemanticChangeSetAdapter.Parse(
                    semanticElement,
                    new DeliveryReceiptLimits { MaxCollectionItems = Math.Max(1, expectedBudget) });
            }
            catch (DeliveryReceiptValidationException ex)
            {
                throw new ArgumentException("expectedSemanticChanges: " + ex.Message, ex);
            }
            expectedBudget -= semantic.ChangeCount;
        }

        var packageChanges = ParsePackageExpectations(
            Optional(root, "expectedPackageChanges"), expectedBudget, options);
        var companions = ParseCompanions(Optional(root, "companionArtifacts"), options);
        return new Parsed(options, semantic, packageChanges, companions);
    }

    // ─── Options ──────────────────────────────────────────────────────────

    private static DeliverableVerificationOptions ParseOptions(JsonElement? element)
    {
        var defaults = new DeliverableVerificationOptions();
        if (element is not { } options) return defaults;
        Require(options.ValueKind == JsonValueKind.Object, "options must be an object");
        RejectUnknown(options, "options",
            "mode", "openXmlVersion", "failOnUnexpectedChanges", "requireNoPlaceholders",
            "detectBracketedAlternativeClauses", "editorialMarkers", "placeholderTokens",
            "maxPackageBytes", "maxFindings", "maxDetectorNodes", "maxDetectorRelationships",
            "maxDetectorTextCharacters", "maxDetectorRegexMatches", "maxDetectorSteps",
            "maxCompanionArtifactBytes", "maxTotalCompanionArtifactBytes",
            "maxCompanionArtifacts", "maxRenderDiagnostics", "maxExpectedChanges",
            "maxReportedDeltaChanges", "packageManifest");

        var manifestDefaults = defaults.PackageManifestOptions;
        var manifest = manifestDefaults;
        if (Optional(options, "packageManifest") is { } manifestElement)
        {
            Require(manifestElement.ValueKind == JsonValueKind.Object,
                "options.packageManifest must be an object");
            RejectUnknown(manifestElement, "options.packageManifest",
                "maxEntryCount", "maxEntryUncompressedBytes", "maxTotalUncompressedBytes",
                "maxXmlPartBytes", "maxCompressionRatio", "maxUriLength");
            manifest = manifestDefaults with
            {
                MaxEntryCount = Int32(manifestElement, "maxEntryCount", manifestDefaults.MaxEntryCount),
                MaxEntryUncompressedBytes = Int64(manifestElement, "maxEntryUncompressedBytes",
                    manifestDefaults.MaxEntryUncompressedBytes),
                MaxTotalUncompressedBytes = Int64(manifestElement, "maxTotalUncompressedBytes",
                    manifestDefaults.MaxTotalUncompressedBytes),
                MaxXmlPartBytes = Int64(manifestElement, "maxXmlPartBytes", manifestDefaults.MaxXmlPartBytes),
                MaxCompressionRatio = Double(manifestElement, "maxCompressionRatio",
                    manifestDefaults.MaxCompressionRatio),
                MaxUriLength = Int32(manifestElement, "maxUriLength", manifestDefaults.MaxUriLength),
            };
        }

        return defaults with
        {
            Mode = Enum<DeliverableVerificationMode>(options, "mode", defaults.Mode),
            OpenXmlVersion = Enum<FileFormatVersions>(options, "openXmlVersion", defaults.OpenXmlVersion),
            FailOnUnexpectedChanges = Bool(options, "failOnUnexpectedChanges", defaults.FailOnUnexpectedChanges),
            RequireNoPlaceholders = Bool(options, "requireNoPlaceholders", defaults.RequireNoPlaceholders),
            DetectBracketedAlternativeClauses = Bool(options, "detectBracketedAlternativeClauses",
                defaults.DetectBracketedAlternativeClauses),
            EditorialMarkers = Strings(options, "editorialMarkers", defaults.EditorialMarkers),
            PlaceholderTokens = Strings(options, "placeholderTokens", defaults.PlaceholderTokens),
            MaxPackageBytes = Int64(options, "maxPackageBytes", defaults.MaxPackageBytes),
            MaxFindings = Int32(options, "maxFindings", defaults.MaxFindings),
            MaxDetectorNodes = Int32(options, "maxDetectorNodes", defaults.MaxDetectorNodes),
            MaxDetectorRelationships = Int32(options, "maxDetectorRelationships", defaults.MaxDetectorRelationships),
            MaxDetectorTextCharacters = Int64(options, "maxDetectorTextCharacters", defaults.MaxDetectorTextCharacters),
            MaxDetectorRegexMatches = Int32(options, "maxDetectorRegexMatches", defaults.MaxDetectorRegexMatches),
            MaxDetectorSteps = Int64(options, "maxDetectorSteps", defaults.MaxDetectorSteps),
            MaxCompanionArtifactBytes = Int64(options, "maxCompanionArtifactBytes", defaults.MaxCompanionArtifactBytes),
            MaxTotalCompanionArtifactBytes = Int64(options, "maxTotalCompanionArtifactBytes",
                defaults.MaxTotalCompanionArtifactBytes),
            MaxCompanionArtifacts = Int32(options, "maxCompanionArtifacts", defaults.MaxCompanionArtifacts),
            MaxRenderDiagnostics = Int32(options, "maxRenderDiagnostics", defaults.MaxRenderDiagnostics),
            MaxExpectedChanges = Int32(options, "maxExpectedChanges", defaults.MaxExpectedChanges),
            MaxReportedDeltaChanges = Int32(options, "maxReportedDeltaChanges", defaults.MaxReportedDeltaChanges),
            PackageManifestOptions = manifest,
        };
    }

    // ─── Expected package changes ─────────────────────────────────────────

    private static IReadOnlyList<DeliverablePackageChangeExpectation> ParsePackageExpectations(
        JsonElement? element, int budget, DeliverableVerificationOptions options)
    {
        if (element is not { } array) return Array.Empty<DeliverablePackageChangeExpectation>();
        Require(array.ValueKind == JsonValueKind.Array, "expectedPackageChanges must be an array");
        var count = array.GetArrayLength();
        Require(count <= budget,
            $"expectedPackageChanges has {count} entries; with the expected semantic changes that exceeds maxExpectedChanges ({options.MaxExpectedChanges})");
        var result = new List<DeliverablePackageChangeExpectation>(count);
        int index = 0;
        foreach (var item in array.EnumerateArray())
        {
            var path = $"expectedPackageChanges[{index++}]";
            Require(item.ValueKind == JsonValueKind.Object, path + " must be an object");
            RejectUnknown(item, path, "kind", "location", "beforeDigest", "afterDigest",
                "beforeValue", "afterValue");
            result.Add(new DeliverablePackageChangeExpectation
            {
                Kind = RequiredEnum<DeliverablePackageChangeKind>(item, "kind", path),
                Location = ParseLocation(Required(item, "location", path), path + ".location"),
                BeforeDigest = ParseDigest(Optional(item, "beforeDigest"), path + ".beforeDigest"),
                AfterDigest = ParseDigest(Optional(item, "afterDigest"), path + ".afterDigest"),
                BeforeValue = String(item, "beforeValue", path),
                AfterValue = String(item, "afterValue", path),
            });
        }
        return result;
    }

    private static ChangeLocation ParseLocation(JsonElement element, string path)
    {
        Require(element.ValueKind == JsonValueKind.Object, path + " must be an object");
        RejectUnknown(element, path, "entryUri", "ownerUri", "relationshipId", "targetUri", "propertyPath");
        return new ChangeLocation
        {
            EntryUri = String(element, "entryUri", path),
            OwnerUri = String(element, "ownerUri", path),
            RelationshipId = String(element, "relationshipId", path),
            TargetUri = String(element, "targetUri", path),
            PropertyPath = String(element, "propertyPath", path),
        };
    }

    private static VerificationDigest? ParseDigest(JsonElement? element, string path)
    {
        if (element is not { } digest) return null;
        Require(digest.ValueKind == JsonValueKind.Object, path + " must be an object");
        RejectUnknown(digest, path, "algorithm", "value");
        return new VerificationDigest
        {
            Algorithm = String(digest, "algorithm", path) ?? throw Missing(path + ".algorithm"),
            Value = String(digest, "value", path) ?? throw Missing(path + ".value"),
        };
    }

    // ─── Companion artifacts ──────────────────────────────────────────────

    private static IReadOnlyList<DeliverableCompanionArtifactInput> ParseCompanions(
        JsonElement? element, DeliverableVerificationOptions options)
    {
        if (element is not { } array) return Array.Empty<DeliverableCompanionArtifactInput>();
        Require(array.ValueKind == JsonValueKind.Array, "companionArtifacts must be an array");
        var count = array.GetArrayLength();
        Require(count <= options.MaxCompanionArtifacts,
            $"companionArtifacts has {count} entries, more than maxCompanionArtifacts ({options.MaxCompanionArtifacts})");

        var result = new List<DeliverableCompanionArtifactInput>(count);
        long totalBytes = 0;
        int totalDiagnostics = 0;
        int index = 0;
        foreach (var item in array.EnumerateArray())
        {
            var path = $"companionArtifacts[{index++}]";
            Require(item.ValueKind == JsonValueKind.Object, path + " must be an object");
            RejectUnknown(item, path, "artifactId", "role", "mediaType", "availability", "bytesB64",
                "unavailableReason", "pageCount", "rendererFingerprint", "sourcePackageDigest",
                "pageMapDigest", "renderDiagnostics");

            // Byte limits are enforced on the encoded length, before decoding allocates anything.
            byte[]? bytes = null;
            if (Optional(item, "bytesB64") is { } encoded)
            {
                Require(encoded.ValueKind == JsonValueKind.String, path + ".bytesB64 must be a base64 string");
                var encodedLength = encoded.GetString()!.Length;
                var decodedUpperBound = (long)encodedLength / 4 * 3 + 2;
                Require(decodedUpperBound <= options.MaxCompanionArtifactBytes,
                    $"{path}.bytesB64 would decode to more than maxCompanionArtifactBytes ({options.MaxCompanionArtifactBytes}); refused before decoding");
                totalBytes += decodedUpperBound;
                Require(totalBytes <= options.MaxTotalCompanionArtifactBytes,
                    $"companionArtifacts exceed maxTotalCompanionArtifactBytes ({options.MaxTotalCompanionArtifactBytes}) at {path}; refused before decoding");
                try
                {
                    bytes = Convert.FromBase64String(encoded.GetString()!);
                }
                catch (FormatException ex)
                {
                    throw new ArgumentException(path + ".bytesB64 is not valid base64", ex);
                }
            }

            var diagnostics = Array.Empty<DeliverableRenderDiagnostic>();
            if (Optional(item, "renderDiagnostics") is { } diagnosticsElement)
            {
                Require(diagnosticsElement.ValueKind == JsonValueKind.Array,
                    path + ".renderDiagnostics must be an array");
                totalDiagnostics += diagnosticsElement.GetArrayLength();
                Require(totalDiagnostics <= options.MaxRenderDiagnostics,
                    $"renderDiagnostics across companionArtifacts exceed maxRenderDiagnostics ({options.MaxRenderDiagnostics}) at {path}");
                diagnostics = diagnosticsElement.EnumerateArray()
                    .Select((diagnostic, i) => ParseDiagnostic(diagnostic, $"{path}.renderDiagnostics[{i}]"))
                    .ToArray();
            }

            result.Add(new DeliverableCompanionArtifactInput
            {
                ArtifactId = String(item, "artifactId", path) ?? throw Missing(path + ".artifactId"),
                Role = RequiredEnum<DeliverableArtifactRole>(item, "role", path),
                MediaType = String(item, "mediaType", path) ?? throw Missing(path + ".mediaType"),
                Availability = Enum(item, "availability", DeliverableArtifactAvailability.Available),
                Bytes = bytes,
                UnavailableReason = String(item, "unavailableReason", path),
                PageCount = Optional(item, "pageCount") is { } pageCount
                    ? pageCount.ValueKind == JsonValueKind.Number && pageCount.TryGetInt64(out var pages)
                        ? pages
                        : throw new ArgumentException(path + ".pageCount must be an integer")
                    : null,
                RendererFingerprint = String(item, "rendererFingerprint", path),
                SourcePackageDigest = ParseDigest(Optional(item, "sourcePackageDigest"), path + ".sourcePackageDigest"),
                PageMapDigest = ParseDigest(Optional(item, "pageMapDigest"), path + ".pageMapDigest"),
                RenderDiagnostics = diagnostics,
            });
        }
        return result;
    }

    private static DeliverableRenderDiagnostic ParseDiagnostic(JsonElement element, string path)
    {
        Require(element.ValueKind == JsonValueKind.Object, path + " must be an object");
        RejectUnknown(element, path, "kind", "message", "severity", "code", "phase");
        return new DeliverableRenderDiagnostic
        {
            Kind = RequiredEnum<DeliverableRenderDiagnosticKind>(element, "kind", path),
            Message = String(element, "message", path) ?? throw Missing(path + ".message"),
            Severity = Enum(element, "severity", VerificationFindingSeverity.Warning),
            Code = String(element, "code", path),
            Phase = String(element, "phase", path),
        };
    }

    // ─── Primitives ───────────────────────────────────────────────────────

    private static JsonElement? Optional(JsonElement element, string name) =>
        element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null
            ? value
            : null;

    private static JsonElement Required(JsonElement element, string name, string path) =>
        Optional(element, name) ?? throw Missing(path + "." + name);

    private static string? String(JsonElement element, string name, string path)
    {
        if (Optional(element, name) is not { } value) return null;
        Require(value.ValueKind == JsonValueKind.String, $"{path}.{name} must be a string");
        return value.GetString();
    }

    private static bool Bool(JsonElement element, string name, bool fallback)
    {
        if (Optional(element, name) is not { } value) return fallback;
        Require(value.ValueKind is JsonValueKind.True or JsonValueKind.False, $"options.{name} must be a boolean");
        return value.GetBoolean();
    }

    private static int Int32(JsonElement element, string name, int fallback)
    {
        if (Optional(element, name) is not { } value) return fallback;
        Require(value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var parsed),
            $"options.{name} must be a 32-bit integer");
        return value.GetInt32();
    }

    private static double Double(JsonElement element, string name, double fallback)
    {
        if (Optional(element, name) is not { } value) return fallback;
        Require(value.ValueKind == JsonValueKind.Number, $"options.{name} must be a number");
        return value.GetDouble();
    }

    private static long Int64(JsonElement element, string name, long fallback)
    {
        if (Optional(element, name) is not { } value) return fallback;
        Require(value.ValueKind == JsonValueKind.Number && value.TryGetInt64(out _),
            $"options.{name} must be an integer");
        return value.GetInt64();
    }

    private static IReadOnlyList<string> Strings(JsonElement element, string name, IReadOnlyList<string> fallback)
    {
        if (Optional(element, name) is not { } value) return fallback;
        Require(value.ValueKind == JsonValueKind.Array, $"options.{name} must be an array of strings");
        return value.EnumerateArray().Select(item =>
        {
            Require(item.ValueKind == JsonValueKind.String, $"options.{name} must contain only strings");
            return item.GetString()!;
        }).ToArray();
    }

    private static T Enum<T>(JsonElement element, string name, T fallback)
        where T : struct, System.Enum =>
        Optional(element, name) is { } value ? ParseEnum<T>(value, name) : fallback;

    private static T RequiredEnum<T>(JsonElement element, string name, string path)
        where T : struct, System.Enum =>
        ParseEnum<T>(Required(element, name, path), path + "." + name);

    private static T ParseEnum<T>(JsonElement value, string path)
        where T : struct, System.Enum
    {
        Require(value.ValueKind == JsonValueKind.String, path + " must be a string");
        // Wire vocabulary is camelCase; the model's names are PascalCase.
        if (System.Enum.TryParse<T>(value.GetString(), ignoreCase: true, out var parsed)
            && System.Enum.IsDefined(parsed))
            return parsed;
        throw new ArgumentException(
            $"{path} must be one of: " + string.Join(", ",
                System.Enum.GetNames<T>().Select(n => char.ToLowerInvariant(n[0]) + n[1..])));
    }

    private static void RejectUnknown(JsonElement element, string path, params string[] known)
    {
        foreach (var property in element.EnumerateObject())
        {
            if (Array.IndexOf(known, property.Name) < 0)
                throw new ArgumentException(
                    $"{path} has an unknown property '{property.Name}'; known: {string.Join(", ", known)}");
        }
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new ArgumentException(message);
    }

    private static ArgumentException Missing(string path) =>
        new(path + " is required");
}

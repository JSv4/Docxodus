// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using Docxodus.Verification;

namespace Docxodus.Internal;

/// <summary>
/// Single-owner wire facade for the portable delivery-receipt verify surface (issue #520).
/// Every transport — WASM bridge, npm/TypeScript, stdio python-host, MCP server — routes
/// through this string-in/string-out entry point, so the wire contract lives in exactly one
/// place: artifacts travel as a JSON object of <c>{"artifactId": "&lt;base64&gt;"}</c>, enums
/// serialize snake_case, and the result mirrors <see cref="DeliveryReceiptVerificationResult"/>.
/// Receipt BUILDING deliberately stays on the typed .NET surface
/// (<see cref="DeliveryChangeReceiptBuilder"/>) and the delivery-bundle operation that drives
/// it — the receipt JSON itself is portable, so remote consumers verify; they do not compose.
/// </summary>
public static class DeliveryOps
{
    /// <summary>
    /// Parse and verify a portable JSON delivery change receipt against optionally supplied
    /// artifact bytes. Never throws for malformed wire input: a bad artifacts object or an
    /// unparsable receipt returns a structured invalid verdict whose findings carry the reason.
    /// </summary>
    /// <param name="receiptJson">The receipt envelope (<c>{"payload":…, "receiptDigest":…}</c>).</param>
    /// <param name="artifactsBase64Json">JSON object mapping artifact id to base64 content;
    /// null or empty verifies the receipt envelope alone (recorded artifacts report missing).</param>
    public static string VerifyChangeReceiptJson(string receiptJson, string? artifactsBase64Json)
    {
        ArgumentNullException.ThrowIfNull(receiptJson);
        Dictionary<string, byte[]> artifacts;
        try
        {
            artifacts = ParseArtifacts(artifactsBase64Json);
        }
        catch (Exception ex) when (ex is JsonException or FormatException or ArgumentException)
        {
            return Serialize(new DeliveryReceiptVerificationResult
            {
                IsValid = false,
                ReceiptDigestValid = false,
                ContractValid = false,
                CitationBindingsValid = false,
                Findings = new[] { $"malformed_artifacts:{ex.GetType().Name}" },
            });
        }

        return Serialize(DeliveryChangeReceiptVerifier.VerifyJson(receiptJson, artifacts));
    }

    /// <summary>Cap on the bytes a transport returns inline for one bundle (manifest plus artifacts).</summary>
    public const long DefaultMaxReturnedBytes = 64L * 1024 * 1024;

    /// <summary>
    /// The delivery-bundle wire shape every transport publishes: bundle status, manifest (as an
    /// object and as its canonical bytes), and each artifact with base64 bytes or the reason it
    /// is unavailable. <paramref name="evidence"/>, when given, is the session's recorder status
    /// so a caller reading an unavailable change receipt sees why in the same response.
    /// </summary>
    public static string SerializeBundle(
        Delivery.DeliveryBundle bundle,
        long maxReturnedBytes = DefaultMaxReturnedBytes,
        Delivery.DeliveryEvidenceStatus? evidence = null)
    {
        ArgumentNullException.ThrowIfNull(bundle);
        var manifestBytes = bundle.ManifestBytes;
        var returnedBytes = manifestBytes.LongLength;
        foreach (var artifact in bundle.Manifest.Payload.Artifacts)
        {
            if (artifact.ByteLength is { } length)
            {
                if (returnedBytes > maxReturnedBytes - Math.Min(length, maxReturnedBytes))
                    throw new InvalidOperationException(
                        $"delivery bundle exceeds the {maxReturnedBytes}-byte return limit; use the CLI or programmatic API");
                returnedBytes += length;
            }
        }
        var buffer = new System.Buffers.ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("status", CamelCase(bundle.Manifest.Payload.Status));
            writer.WriteBoolean("verified", bundle.Verification.IsValid);
            writer.WriteBoolean("manifestVerified", bundle.Verification.IsValid);
            writer.WritePropertyName("manifest");
            using (var manifest = JsonDocument.Parse(manifestBytes))
                manifest.RootElement.WriteTo(writer);
            writer.WriteBase64String("manifestBytes", manifestBytes);
            writer.WriteStartArray("artifacts");
            foreach (var artifact in bundle.Manifest.Payload.Artifacts)
            {
                writer.WriteStartObject();
                writer.WriteString("artifactId", artifact.ArtifactId);
                writer.WriteString("kind", CamelCase(artifact.Kind));
                writer.WriteString("requiredness", CamelCase(artifact.Requiredness));
                writer.WriteString("availability", CamelCase(artifact.Availability));
                writer.WriteString("relativePath", artifact.RelativePath);
                writer.WriteString("mediaType", artifact.MediaType);
                if (artifact.Availability == Delivery.DeliveryArtifactAvailability.Available)
                    writer.WriteBase64String("bytes", bundle.GetArtifactBytes(artifact.ArtifactId));
                else
                    writer.WriteString("unavailableReason", artifact.UnavailableReason);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            if (evidence is not null)
            {
                writer.WritePropertyName("evidence");
                WriteEvidenceStatus(writer, evidence);
            }
            writer.WriteEndObject();
        }
        return Encoding.UTF8.GetString(buffer.WrittenSpan);
    }

    /// <summary>The recorder status wire shape shared by every transport.</summary>
    public static string SerializeEvidenceStatus(Delivery.DeliveryEvidenceStatus status)
    {
        ArgumentNullException.ThrowIfNull(status);
        var buffer = new System.Buffers.ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
            WriteEvidenceStatus(writer, status);
        return Encoding.UTF8.GetString(buffer.WrittenSpan);
    }

    private static void WriteEvidenceStatus(Utf8JsonWriter writer, Delivery.DeliveryEvidenceStatus status)
    {
        writer.WriteStartObject();
        writer.WriteBoolean("enabled", status.Enabled);
        writer.WriteNumber("transactionCount", status.TransactionCount);
        writer.WriteNumber("lineageEventCount", status.LineageEventCount);
        writer.WriteNumber("unlabeledTransactionCount", status.UnlabeledTransactionCount);
        writer.WriteNumber("retainedStateCount", status.RetainedStateCount);
        writer.WriteNumber("retainedBytes", status.RetainedBytes);
        writer.WriteNumber("sourceVersion", status.SourceVersion);
        writer.WriteNumber("currentVersion", status.CurrentVersion);
        if (status.UnavailableReason is null) writer.WriteNull("unavailableReason");
        else writer.WriteString("unavailableReason", status.UnavailableReason);
        writer.WriteEndObject();
    }

    /// <summary><c>{"privacyProfile","failOnUnexpectedChanges"}</c>; null or empty means the defaults.</summary>
    public static Delivery.DeliveryReceiptBuildOptions ParseReceiptBuildOptions(string? optionsJson)
    {
        var options = new Delivery.DeliveryReceiptBuildOptions();
        if (string.IsNullOrWhiteSpace(optionsJson)) return options;
        using var document = JsonDocument.Parse(optionsJson);
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Null) return options;
        if (root.ValueKind != JsonValueKind.Object)
            throw new ArgumentException("delivery receipt options must be a JSON object");
        if (root.TryGetProperty("privacyProfile", out var profile) && profile.ValueKind != JsonValueKind.Null)
        {
            if (profile.ValueKind != JsonValueKind.String)
                throw new ArgumentException("privacyProfile must be a string");
            options = options with { PrivacyProfile = ParsePrivacyProfile(profile.GetString()!) };
        }
        if (root.TryGetProperty("failOnUnexpectedChanges", out var fail) && fail.ValueKind != JsonValueKind.Null)
        {
            if (fail.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
                throw new ArgumentException("failOnUnexpectedChanges must be a boolean");
            options = options with { FailOnUnexpectedChanges = fail.GetBoolean() };
        }
        return options;
    }

    /// <summary>Accepts the receipt's camelCase spelling and snake_case: <c>hashOnly</c>, <c>hash_and_summary</c>, ….</summary>
    public static DeliveryReceiptPrivacyProfile ParsePrivacyProfile(string value)
    {
        var compact = value.Replace("_", string.Empty, StringComparison.Ordinal)
            .Replace("-", string.Empty, StringComparison.Ordinal);
        foreach (var candidate in Enum.GetValues<DeliveryReceiptPrivacyProfile>())
        {
            if (string.Equals(candidate.ToString(), compact, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }
        throw new ArgumentException($"unknown privacy profile: {value}");
    }

    /// <summary><c>[{"tool","action","args"?}]</c> as a transport describes the batch it is about to run.</summary>
    public static IReadOnlyList<(string Tool, string Action, string? ArgumentsJson)> ParseEvidenceOperations(string operationsJson)
    {
        ArgumentNullException.ThrowIfNull(operationsJson);
        using var document = JsonDocument.Parse(operationsJson);
        if (document.RootElement.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("evidence operations must be a JSON array");
        var operations = new List<(string, string, string?)>();
        foreach (var element in document.RootElement.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object
                || !element.TryGetProperty("tool", out var tool) || tool.ValueKind != JsonValueKind.String
                || !element.TryGetProperty("action", out var action) || action.ValueKind != JsonValueKind.String)
                throw new ArgumentException("each evidence operation needs string \"tool\" and \"action\"");
            string? args = element.TryGetProperty("args", out var argsElement)
                && argsElement.ValueKind == JsonValueKind.Object
                ? argsElement.GetRawText()
                : null;
            operations.Add((tool.GetString()!, action.GetString()!, args));
        }
        return operations;
    }

    /// <summary><c>{"transactionId","requestFingerprint"}</c> or null/empty for none.</summary>
    public static DeliveryTransactionIdentity? ParseTransactionIdentity(string? identityJson)
    {
        if (string.IsNullOrWhiteSpace(identityJson)) return null;
        using var document = JsonDocument.Parse(identityJson);
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Null) return null;
        if (root.ValueKind != JsonValueKind.Object
            || !root.TryGetProperty("transactionId", out var id) || id.ValueKind != JsonValueKind.String
            || !root.TryGetProperty("requestFingerprint", out var fingerprint) || fingerprint.ValueKind != JsonValueKind.String)
            throw new ArgumentException("a transaction identity needs string \"transactionId\" and \"requestFingerprint\"");
        return new DeliveryTransactionIdentity
        {
            TransactionId = id.GetString()!,
            RequestFingerprint = fingerprint.GetString()!,
        };
    }

    /// <summary>The <c>steps</c> array of a transport-composed batch result.</summary>
    public static IReadOnlyList<MutationBatchStepResult> ParseBatchSteps(string stepsJson)
    {
        ArgumentNullException.ThrowIfNull(stepsJson);
        using var document = JsonDocument.Parse(stepsJson);
        if (document.RootElement.ValueKind != JsonValueKind.Array)
            throw new ArgumentException("batch steps must be a JSON array");
        var steps = new List<MutationBatchStepResult>();
        foreach (var element in document.RootElement.EnumerateArray())
        {
            if (element.ValueKind != JsonValueKind.Object
                || !element.TryGetProperty("index", out var index) || !index.TryGetInt32(out var indexValue)
                || !element.TryGetProperty("tool", out var tool) || tool.ValueKind != JsonValueKind.String
                || !element.TryGetProperty("action", out var action) || action.ValueKind != JsonValueKind.String
                || !element.TryGetProperty("results", out var results) || results.ValueKind != JsonValueKind.Array)
                throw new ArgumentException("each batch step needs index, tool, action and a results array");
            var rolledBack = element.TryGetProperty("rolledBack", out var rb) && rb.ValueKind == JsonValueKind.True;
            steps.Add(new MutationBatchStepResult(
                indexValue,
                tool.GetString()!,
                action.GetString()!,
                DocxSessionJson.DeserializeEditResults(results.GetRawText()),
                rolledBack));
        }
        return steps;
    }

    private static string CamelCase<T>(T value)
        where T : struct, Enum =>
        JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    private static Dictionary<string, byte[]> ParseArtifacts(string? artifactsBase64Json)
    {
        var artifacts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(artifactsBase64Json)) return artifacts;
        using var document = JsonDocument.Parse(artifactsBase64Json);
        if (document.RootElement.ValueKind != JsonValueKind.Object)
            throw new JsonException("artifacts must be a JSON object of {artifactId: base64}");
        foreach (var property in document.RootElement.EnumerateObject())
        {
            if (property.Value.ValueKind != JsonValueKind.String)
                throw new JsonException($"artifact '{property.Name}' content must be a base64 string");
            artifacts[property.Name] = Convert.FromBase64String(property.Value.GetString()!);
        }
        return artifacts;
    }

    private static string Serialize(DeliveryReceiptVerificationResult result)
    {
        var sb = new StringBuilder(256);
        sb.Append("{\"isValid\":").Append(result.IsValid ? "true" : "false")
          .Append(",\"receiptDigestValid\":").Append(result.ReceiptDigestValid ? "true" : "false")
          .Append(",\"contractValid\":").Append(result.ContractValid ? "true" : "false")
          .Append(",\"citationBindingsValid\":").Append(result.CitationBindingsValid ? "true" : "false")
          .Append(",\"artifacts\":[");
        for (int i = 0; i < result.Artifacts.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var artifact = result.Artifacts[i];
            sb.Append("{\"artifactId\":").Append(DocxSessionJson.JsonString(artifact.ArtifactId))
              .Append(",\"status\":\"").Append(DocxSessionJson.EnumToSnake(artifact.Status)).Append('"');
            if (artifact.ExpectedLength is { } expectedLength)
                sb.Append(",\"expectedLength\":").Append(expectedLength);
            if (artifact.ActualLength is { } actualLength)
                sb.Append(",\"actualLength\":").Append(actualLength);
            AppendDigest(sb, "expectedDigest", artifact.ExpectedDigest);
            AppendDigest(sb, "actualDigest", artifact.ActualDigest);
            sb.Append('}');
        }

        sb.Append("],\"findings\":[");
        for (int i = 0; i < result.Findings.Count; i++)
        {
            if (i > 0) sb.Append(',');
            sb.Append(DocxSessionJson.JsonString(result.Findings[i]));
        }

        sb.Append("]}");
        return sb.ToString();
    }

    private static void AppendDigest(StringBuilder sb, string name, VerificationDigest? digest)
    {
        if (digest is null) return;
        sb.Append(",\"").Append(name).Append("\":{\"algorithm\":")
          .Append(DocxSessionJson.JsonString(digest.Algorithm))
          .Append(",\"value\":").Append(DocxSessionJson.JsonString(digest.Value))
          .Append('}');
    }
}

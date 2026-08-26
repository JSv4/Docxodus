// Copyright (c) Microsoft. All rights reserved.
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

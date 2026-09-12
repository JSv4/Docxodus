#nullable enable

using System.Buffers;
using System.Text.Json;
using Docxodus.Delivery;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>
/// Agent-server adapter for the shared delivery service. The adapter owns only MCP parsing and
/// byte serialization; artifact planning, policy, validation, and verification stay in core.
/// </summary>
internal static class DeliveryTool
{
    internal const long MaxReturnedBytes = 64L * 1024 * 1024;

    internal static string Execute(SessionStore store, DocSession session, JsonElement args)
    {
        if (args.ValueKind != JsonValueKind.Object)
            throw new McpToolException("docxodus_deliver arguments must be an object");

        var baselineLocation = store.Documents.Resolve(String(args, "baselinePath"));
        var baselineBytes = store.Documents.Read(baselineLocation);
        var baselineVersion = NonNegativeLong(args, "baselineDocumentVersion");
        var finalVersion = NonNegativeLong(args, "finalDocumentVersion");
        var finalName = String(args, "finalDocumentName");
        var policy = Object(args, "revisionPolicy");
        var artifactArray = Array(args, "artifacts");
        if (artifactArray.GetArrayLength() == 0)
            throw new McpToolException("docxodus_deliver requires at least one artifact");
        var artifacts = artifactArray.EnumerateArray()
            .Select(ParseArtifact)
            .ToArray();

        // A requested change receipt is minted from the session's host-captured evidence
        // (issue #748). The working document is then the recorder's own current state, so the
        // delivered bytes are the last recorded after-state by construction; without a receipt
        // request the clean save is used as before.
        var evidence = artifacts.Any(artifact => artifact.Kind == DeliveryArtifactKind.ChangeReceipt)
            ? ExportEvidence(session, args, baselineBytes)
            : null;
        var workingBytes = evidence?.Working.Bytes
            ?? DocxSessionOps.Save(session.Handle, persistAnchorIds: false);

        var request = new DeliveryBundleBuildRequest(
            new DeliveryDocumentSnapshot(
                "baseline:" + Path.GetFileName(baselineLocation),
                baselineVersion,
                baselineBytes),
            new DeliveryDocumentSnapshot(
                "working:" + session.Id,
                evidence?.Working.DocumentVersion ?? DocxSessionOps.GetVersion(session.Handle),
                workingBytes),
            finalName,
            finalVersion,
            new DeliveryBundleRevisionPolicy
            {
                PreExistingRevisions = RevisionPolicy(
                    String(policy, "preExistingRevisions"), "preExistingRevisions"),
                GeneratedRevisions = RevisionPolicy(
                    String(policy, "generatedRevisions"), "generatedRevisions"),
            },
            artifacts,
            evidence?.ReceiptContext);
        var options = new DeliveryBundleBuildOptions
        {
            ReturnIncompleteBundle = OptionalBoolean(args, "returnIncompleteBundle", false),
            FailOnDeliverableValidationFailure = OptionalBoolean(
                args, "failOnDeliverableValidationFailure", true),
        };

        try
        {
            var rendererOptions = DocxodusExportHostRendererOptions.FromEnvironment();
            var renderer = rendererOptions is null
                ? null
                : new DocxodusExportHostRenderer(rendererOptions);
            var bundle = new DeliveryBundleService(renderer)
                .BuildAsync(request, options)
                .AsTask()
                .GetAwaiter()
                .GetResult();
            return DeliveryOps.SerializeBundle(bundle, MaxReturnedBytes, evidence?.Status);
        }
        catch (DeliveryBundleException ex)
        {
            throw new McpToolException($"{ex.Code}: {ex.Message}");
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException
                                   or IOException or UnauthorizedAccessException)
        {
            throw new McpToolException($"delivery_configuration_failed: {ex.Message}");
        }
    }

    /// <summary>
    /// The session's captured evidence for this delivery. The receipt's source document is the
    /// package the session opened, so a baseline that is a different package cannot be attested
    /// by this history; the receipt is then unavailable with that reason rather than rejected
    /// deep inside lineage validation.
    /// </summary>
    private static DeliveryEvidenceExport ExportEvidence(DocSession session, JsonElement args, byte[] baselineBytes)
    {
        var options = args.TryGetProperty("changeReceipt", out var receipt) && receipt.ValueKind == JsonValueKind.Object
            ? DeliveryOps.ParseReceiptBuildOptions(receipt.GetRawText())
            : new DeliveryReceiptBuildOptions();
        var export = DocxSessionOps.ExportDeliveryEvidence(session.Handle, options);
        if (export.ReceiptContext is null) return export;
        string? mismatch = null;
        if (!baselineBytes.AsSpan().SequenceEqual(export.Source.Bytes))
            mismatch = "the delivery baseline is not the package this session opened, so the captured history cannot attest it";
        else if (NonNegativeLong(args, "baselineDocumentVersion") != export.Source.DocumentVersion)
            mismatch = $"baselineDocumentVersion must be {export.Source.DocumentVersion}, the version the captured history starts at";
        else if (NonNegativeLong(args, "finalDocumentVersion") != export.Working.DocumentVersion)
            mismatch = $"finalDocumentVersion must be {export.Working.DocumentVersion}, the session version the captured history ends at";
        if (mismatch is null) return export;
        return new DeliveryEvidenceExport(
            null,
            export.Source,
            export.Working,
            export.Status with { UnavailableReason = mismatch });
    }

    private static DeliveryArtifactRequest ParseArtifact(JsonElement value)
    {
        if (value.ValueKind != JsonValueKind.Object)
            throw new McpToolException("each delivery artifact must be an object");
        var kind = EnumValue<DeliveryArtifactKind>(String(value, "kind"), "artifact kind");
        var requiredness = EnumValue<DeliveryArtifactRequiredness>(
            String(value, "requiredness"), "artifact requiredness");
        var review = OptionalString(value, "reviewProfile") is { } reviewName
            ? EnumValue<DeliveryReviewProfile>(reviewName, "review profile")
            : (DeliveryReviewProfile?)null;
        var comments = OptionalString(value, "commentProfile") is { } commentName
            ? EnumValue<DeliveryCommentProfile>(commentName, "comment profile")
            : (DeliveryCommentProfile?)null;
        return new DeliveryArtifactRequest
        {
            ArtifactId = String(value, "artifactId"),
            Kind = kind,
            Requiredness = requiredness,
            ReviewProfile = review,
            CommentProfile = comments,
        };
    }

    private static string Name<T>(T value)
        where T : struct, Enum =>
        JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    private static T EnumValue<T>(string value, string name)
        where T : struct, Enum
    {
        var compact = value.Replace("-", string.Empty, StringComparison.Ordinal)
            .Replace("_", string.Empty, StringComparison.Ordinal);
        foreach (var candidate in Enum.GetValues<T>())
        {
            var candidateName = candidate.ToString();
            if (string.Equals(candidateName, compact, StringComparison.OrdinalIgnoreCase))
                return candidate;
        }
        throw new McpToolException($"unknown {name}: {value}");
    }

    private static DeliveryRevisionPolicy RevisionPolicy(string value, string name) =>
        EnumValue<DeliveryRevisionPolicy>(value, name);

    private static JsonElement Object(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Object)
            throw new McpToolException($"missing required object argument \"{name}\"");
        return value;
    }

    private static JsonElement Array(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.Array)
            throw new McpToolException($"missing required array argument \"{name}\"");
        return value;
    }

    private static string String(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || value.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(value.GetString()))
            throw new McpToolException($"missing required string argument \"{name}\"");
        return value.GetString()!;
    }

    private static string? OptionalString(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value))
            return null;
        if (value.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(value.GetString()))
            throw new McpToolException($"optional argument \"{name}\" must be a non-blank string");
        return value.GetString();
    }

    private static long NonNegativeLong(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var value) || !value.TryGetInt64(out var number)
            || number < 0)
            throw new McpToolException($"argument \"{name}\" must be a non-negative integer");
        return number;
    }

    private static bool OptionalBoolean(JsonElement args, string name, bool defaultValue)
    {
        if (!args.TryGetProperty(name, out var value))
            return defaultValue;
        if (value.ValueKind is not (JsonValueKind.True or JsonValueKind.False))
            throw new McpToolException($"argument \"{name}\" must be a boolean");
        return value.GetBoolean();
    }
}

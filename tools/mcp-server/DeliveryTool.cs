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
        var workingBytes = DocxSessionOps.Save(session.Handle, persistAnchorIds: false);
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

        var request = new DeliveryBundleBuildRequest(
            new DeliveryDocumentSnapshot(
                "baseline:" + Path.GetFileName(baselineLocation),
                baselineVersion,
                baselineBytes),
            new DeliveryDocumentSnapshot(
                "working:" + session.Id,
                DocxSessionOps.GetVersion(session.Handle),
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
            artifacts);
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
            return Serialize(bundle);
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

    private static string Serialize(DeliveryBundle bundle)
    {
        var returnedBytes = bundle.ManifestBytes.LongLength;
        foreach (var artifact in bundle.Manifest.Payload.Artifacts)
        {
            if (artifact.ByteLength is { } length)
            {
                if (returnedBytes > MaxReturnedBytes - Math.Min(length, MaxReturnedBytes))
                    throw new McpToolException(
                        $"delivery bundle exceeds the {MaxReturnedBytes}-byte MCP return limit; use the CLI or programmatic API");
                returnedBytes += length;
            }
        }
        var buffer = new ArrayBufferWriter<byte>();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteString("status", Name(bundle.Manifest.Payload.Status));
            writer.WriteBoolean("verified", bundle.Verification.IsValid);
            writer.WriteBoolean("manifestVerified", bundle.Verification.IsValid);
            writer.WritePropertyName("manifest");
            using (var manifest = JsonDocument.Parse(bundle.ManifestBytes))
                manifest.RootElement.WriteTo(writer);
            writer.WriteBase64String("manifestBytes", bundle.ManifestBytes);
            writer.WriteStartArray("artifacts");
            foreach (var artifact in bundle.Manifest.Payload.Artifacts)
            {
                writer.WriteStartObject();
                writer.WriteString("artifactId", artifact.ArtifactId);
                writer.WriteString("kind", Name(artifact.Kind));
                writer.WriteString("requiredness", Name(artifact.Requiredness));
                writer.WriteString("availability", Name(artifact.Availability));
                writer.WriteString("relativePath", artifact.RelativePath);
                writer.WriteString("mediaType", artifact.MediaType);
                if (artifact.Availability == DeliveryArtifactAvailability.Available)
                    writer.WriteBase64String("bytes", bundle.GetArtifactBytes(artifact.ArtifactId));
                else
                    writer.WriteString("unavailableReason", artifact.UnavailableReason);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            writer.WriteEndObject();
        }
        return System.Text.Encoding.UTF8.GetString(buffer.WrittenSpan);
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

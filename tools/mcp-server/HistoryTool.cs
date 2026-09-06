#nullable enable

using System.Text;
using System.Text.Json;
using Docxodus.History;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>Session-capability-scoped history over host-configured storage, using existing dispatch.</summary>
internal static class HistoryTool
{
    public const string RootVariable = "DOCXODUS_HISTORY_ROOT";

    public static HistoryClientOps? Configure(string? root)
    {
        if (string.IsNullOrWhiteSpace(root)) return null;
        if (!Path.IsPathFullyQualified(root)) throw new McpToolException($"{RootVariable} must be an absolute host-owned path.");
        try
        {
            return new HistoryClientOps(new FileHistoryBlobStore(Path.Combine(root, "blobs")),
                new FileHistoryHeadStore(Path.Combine(root, "heads")));
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException)
        { throw new McpToolException($"Cannot initialize {RootVariable}: {error.Message}"); }
    }

    public static string Execute(SessionStore store, DocSession session, JsonElement args)
    {
        var history = store.History ?? throw new McpToolException($"History is disabled; the host must configure {RootVariable}.");
        var documentId = session.Location ?? throw new McpToolException("History requires a session opened from a scoped document location.");
        if (args.GetRawText().Length > HistoryClientJson.MaxRequestChars)
            throw new McpToolException("History request metadata exceeds its limit.");
        var names = new HashSet<string>(StringComparer.Ordinal);
        foreach (var property in args.EnumerateObject())
            if (!names.Add(property.Name)) throw new McpToolException("Duplicate history argument: " + property.Name);
        if (!args.TryGetProperty("action", out var actionValue) || actionValue.ValueKind != JsonValueKind.String)
            throw new McpToolException("History action is required and must be a string.");
        var action = actionValue.GetString()!;
        using var buffer = new MemoryStream();
        using (var writer = new Utf8JsonWriter(buffer))
        {
            writer.WriteStartObject();
            writer.WriteNumber("schemaVersion", 1);
            writer.WriteString("documentId", documentId);
            writer.WriteString("operation", action == "render" ? "materialize" : action);
            foreach (var property in args.EnumerateObject())
            {
                if (property.Name is "sessionId" or "action") continue;
                if (property.Name is "schemaVersion" or "documentId" or "operation")
                    throw new McpToolException("History identity and schema are assigned by the session capability, not caller arguments.");
                property.WriteTo(writer);
            }
            writer.WriteEndObject();
        }
        var request = HistoryClientJson.Read<HistoryClientRequest>(Encoding.UTF8.GetString(buffer.ToArray()));
        if (action == "render")
        {
            if ((request.Sequence is null) == (request.Cutoff is null))
                throw new McpToolException("History render requires exactly one sequence or cutoff.");
            if (request.Cutoff is not null)
            {
                var resolvedJson = history.InvokeAsync(HistoryClientJson.Write(request with { Operation = "resolveTime" })).GetAwaiter().GetResult();
                var resolved = HistoryClientJson.Read<HistoryClientResult>(resolvedJson);
                if (!resolved.Success) return resolvedJson;
                request = request with { Sequence = resolved.Sequence };
            }
            var resultJson = history.InvokeAsync(HistoryClientJson.Write(request)).GetAwaiter().GetResult();
            using var result = JsonDocument.Parse(resultJson);
            if (!result.RootElement.GetProperty("success").GetBoolean()) return resultJson;
            var bytes = result.RootElement.GetProperty("bytes").GetBytesFromBase64();
            var html = HtmlConversionOps.ConvertToHtml(bytes, new HtmlConversionOptions
            { RenderFootnotesAndEndnotes = true, RenderHeadersAndFooters = true, RenderTrackedChanges = true });
            return "{\"success\":true,\"sequence\":" + JsonRpcIo.JsonString(request.Sequence!.Value.ToString(System.Globalization.CultureInfo.InvariantCulture))
                + ",\"html\":" + JsonRpcIo.JsonString(html) + "}";
        }
        // The existing session gate covers capture + publication. No session/source-file mutation
        // occurs, including restore: it appends history only and leaves open local work untouched.
        var captured = action == "create" ? DocxSessionOps.Save(session.Handle, persistAnchorIds: false) : null;
        return history.InvokeAsync(HistoryClientJson.Write(request), captured).GetAwaiter().GetResult();
    }
}

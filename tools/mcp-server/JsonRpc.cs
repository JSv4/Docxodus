#nullable enable

using System.Text.Json;

namespace Docxodus.McpServer;

/// <summary>
/// A parsed JSON-RPC 2.0 request/notification line. <see cref="Id"/> is null for
/// notifications (no response expected) — <c>notifications/initialized</c> is the only
/// one this server receives.
/// </summary>
internal readonly struct JsonRpcRequest
{
    public JsonElement? Id { get; init; }
    required public string Method { get; init; }
    public JsonElement Params { get; init; }
}

/// <summary>JSON-RPC error codes this server raises. Business-level tool failures are never
/// reported this way — see <see cref="McpToolException"/> — only transport/protocol problems.</summary>
internal static class JsonRpcErrorCodes
{
    public const int ParseError = -32700;
    public const int InvalidRequest = -32600;
    public const int MethodNotFound = -32601;
    public const int InvalidParams = -32602;
    public const int InternalError = -32603;
}

/// <summary>
/// Reads/writes newline-delimited JSON-RPC 2.0 messages on stdio — the transport the Model
/// Context Protocol's stdio binding specifies (one complete, embedded-newline-free JSON value
/// per line, both directions). Mirrors the framing style of <c>tools/python-host/Program.cs</c>'s
/// NDJSON host, just against the MCP wire shape instead of the bespoke <c>{id,op,args}</c> one.
/// </summary>
internal static class JsonRpcIo
{
    public static JsonRpcRequest? ParseRequest(string line)
    {
        using var doc = JsonDocument.Parse(line);
        var root = doc.RootElement.Clone();
        if (!root.TryGetProperty("method", out var methodEl) || methodEl.ValueKind != JsonValueKind.String)
            throw new JsonException("request missing string \"method\"");

        JsonElement? id = root.TryGetProperty("id", out var idEl) ? idEl : null;
        var paramsEl = root.TryGetProperty("params", out var p) ? p : default;
        return new JsonRpcRequest { Id = id, Method = methodEl.GetString()!, Params = paramsEl };
    }

    public static void WriteResult(System.IO.TextWriter w, JsonElement id, string resultJson)
    {
        w.Write("{\"jsonrpc\":\"2.0\",\"id\":");
        w.Write(id.GetRawText());
        w.Write(",\"result\":");
        w.Write(resultJson);
        w.Write("}\n");
    }

    public static void WriteError(System.IO.TextWriter w, JsonElement? id, int code, string message)
    {
        w.Write("{\"jsonrpc\":\"2.0\",\"id\":");
        w.Write(id?.GetRawText() ?? "null");
        w.Write(",\"error\":{\"code\":");
        w.Write(code);
        w.Write(",\"message\":");
        w.Write(JsonString(message));
        w.Write("}}\n");
    }

    public static string JsonString(string s) => JsonSerializer.Serialize(s);
}

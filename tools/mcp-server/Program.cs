#nullable enable

using System;
using System.IO;
using System.Text;
using System.Text.Json;
using Docxodus.Internal;

namespace Docxodus.McpServer;

/// <summary>
/// Stdio Model Context Protocol server. Reads one JSON-RPC 2.0 message per line from stdin,
/// dispatches <c>tools/call</c> into <see cref="Dispatcher"/> (which in turn calls the shared
/// <see cref="DocxSessionOps"/>/<see cref="Docxodus.Internal.DocxDiffOps"/> facades), and writes
/// one JSON-RPC message per line to stdout. Diagnostic output goes to stderr only — stdout is
/// reserved for the protocol, exactly like <c>tools/python-host/Program.cs</c>'s NDJSON host.
///
/// Business-level tool failures (bad arguments, unknown session, an underlying Docxodus
/// <see cref="EditResult"/> with <c>success:false</c>) are reported as a normal <c>tools/call</c>
/// result with <c>isError: true</c>, never as a JSON-RPC error — that's reserved for
/// transport-level problems (malformed JSON, unknown method).
/// </summary>
internal static class Program
{
    private const string ProtocolVersion = "2024-11-05";

    public static int Main(string[] args)
    {
        SessionStore store;
        IDocumentStore documents;
        try
        {
            documents = DocumentStores.FromEnvironment();
            store = new SessionStore(documents);
        }
        catch (McpToolException ex)
        {
            // Misconfigured storage is fatal at startup rather than surfaced per tool call: a
            // server that answered tools/list and then failed every open would be far harder for
            // an operator to diagnose than one that refuses to start with the reason.
            Console.Error.WriteLine($"[mcp:fatal] storage configuration: {ex.Message}");
            return 1;
        }

        using var stdin = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false))
        {
            NewLine = "\n",
            AutoFlush = false,
        };

        Console.Error.WriteLine(
            $"docxodus-mcp ready (storage: {documents.Kind}, root: {documents.RootDescription})");

        string? line;
        while ((line = stdin.ReadLine()) is not null)
        {
            if (line.Length == 0) continue;

            JsonRpcRequest request;
            try
            {
                var parsed = JsonRpcIo.ParseRequest(line);
                if (parsed is null) continue;
                request = parsed.Value;
            }
            catch (JsonException ex)
            {
                JsonRpcIo.WriteError(stdout, null, JsonRpcErrorCodes.ParseError, ex.Message);
                stdout.Flush();
                continue;
            }

            // Notifications carry no "id" and expect no response.
            bool isNotification = request.Id is null;

            try
            {
                var resultJson = HandleMethod(store, request, out var shouldExit);
                if (!isNotification)
                    JsonRpcIo.WriteResult(stdout, request.Id!.Value, resultJson);
                stdout.Flush();
                if (shouldExit) break;
            }
            catch (MethodNotFoundException ex)
            {
                if (!isNotification)
                    JsonRpcIo.WriteError(stdout, request.Id, JsonRpcErrorCodes.MethodNotFound, ex.Message);
                stdout.Flush();
            }
            catch (InvalidParamsException ex)
            {
                if (!isNotification)
                    JsonRpcIo.WriteError(stdout, request.Id, JsonRpcErrorCodes.InvalidParams, ex.Message);
                stdout.Flush();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[mcp:error] {request.Method}: {ex}");
                if (!isNotification)
                    JsonRpcIo.WriteError(stdout, request.Id, JsonRpcErrorCodes.InternalError, ex.Message);
                stdout.Flush();
            }
        }

        store.CloseAll();
        return 0;
    }

    private static string HandleMethod(SessionStore store, JsonRpcRequest request, out bool shouldExit)
    {
        shouldExit = false;
        switch (request.Method)
        {
            case "initialize":
                return "{\"protocolVersion\":\"" + ProtocolVersion
                    + "\",\"capabilities\":{\"tools\":{}},\"serverInfo\":{\"name\":\"docxodus-mcp\",\"version\":\"1.0.0\"}}";

            case "notifications/initialized":
                return "null"; // notification; response is discarded by the caller anyway

            case "ping":
                return "{}";

            case "tools/list":
                return BuildToolsListResult();

            case "tools/call":
                return HandleToolsCall(store, request.Params);

            case "shutdown":
                shouldExit = true;
                return "null";

            default:
                throw new MethodNotFoundException($"method not found: {request.Method}");
        }
    }

    internal static string BuildToolsListResult()
    {
        var sb = new StringBuilder(4096);
        sb.Append("{\"tools\":[");
        for (int i = 0; i < ToolCatalog.Tools.Count; i++)
        {
            if (i > 0) sb.Append(',');
            var t = ToolCatalog.Tools[i];
            sb.Append("{\"name\":").Append(JsonRpcIo.JsonString(t.Name))
              .Append(",\"description\":").Append(JsonRpcIo.JsonString(t.Description))
              .Append(",\"inputSchema\":").Append(JsonSerializer.Serialize(JsonDocument.Parse(t.InputSchemaJson).RootElement))
              .Append('}');
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string HandleToolsCall(SessionStore store, JsonElement callParams)
    {
        if (callParams.ValueKind != JsonValueKind.Object
            || !callParams.TryGetProperty("name", out var nameEl) || nameEl.ValueKind != JsonValueKind.String)
            throw new InvalidParamsException("tools/call params missing string \"name\"");
        var toolName = nameEl.GetString()!;
        var arguments = callParams.TryGetProperty("arguments", out var a) ? a : default;

        try
        {
            var resultJson = Dispatcher.Call(store, toolName, arguments);
            var isError = LooksLikeFailure(resultJson);
            return $"{{\"content\":[{{\"type\":\"text\",\"text\":{JsonRpcIo.JsonString(resultJson)}}}],\"isError\":{(isError ? "true" : "false")}}}";
        }
        catch (Exception ex)
        {
            // Any dispatch failure (bad arguments, unknown session, unknown tool/action, an
            // internal exception) is a TOOL error, not a JSON-RPC protocol error — see the
            // class doc. The agent sees the message in the result content and can react to it.
            return $"{{\"content\":[{{\"type\":\"text\",\"text\":{JsonRpcIo.JsonString(ex.Message)}}}],\"isError\":true}}";
        }
    }

    /// <summary>Heuristic: a top-level <c>{"success":false,...}</c> — the shape every
    /// Docxodus <c>EditResult</c> serializes to — is surfaced as an MCP tool error so
    /// the calling agent notices without having to parse the text content first.</summary>
    private static bool LooksLikeFailure(string resultJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(resultJson);
            return doc.RootElement.ValueKind == JsonValueKind.Object
                && doc.RootElement.TryGetProperty("success", out var s)
                && s.ValueKind == JsonValueKind.False;
        }
        catch (JsonException)
        {
            return false;
        }
    }
}

internal sealed class MethodNotFoundException : Exception
{
    public MethodNotFoundException(string message) : base(message) { }
}

internal sealed class InvalidParamsException : Exception
{
    public InvalidParamsException(string message) : base(message) { }
}

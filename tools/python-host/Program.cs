#nullable enable

using System;
using System.IO;
using System.Text;
using System.Text.Json;
using Docxodus;
using Docxodus.Internal;

namespace Docxodus.PyHost;

/// <summary>
/// Stdio NDJSON host for the python-docxodus wrapper. Reads one JSON request
/// per line from stdin, dispatches via <see cref="Dispatcher"/> into the shared
/// <see cref="DocxSessionOps"/> facade, and writes one JSON response per line
/// to stdout. Diagnostic output goes to stderr only — stdout is reserved for
/// the protocol.
///
/// Sessions live across requests, keyed by the integer handle returned from
/// <c>open_session</c>. The host exits cleanly on <c>shutdown</c> or when
/// stdin closes (the Python parent process exits).
///
/// Request:  {"id": &lt;int&gt;, "op": &lt;string&gt;, "args": &lt;object&gt;}
/// Success:  {"id": &lt;int&gt;, "ok": true,  "result": &lt;any&gt;}
/// Failure:  {"id": &lt;int&gt;, "ok": false, "error": {"code": &lt;str&gt;, "message": &lt;str&gt;, "trace"?: &lt;str&gt;}}
///
/// "Failure" here means transport-level failure (host crash, malformed
/// request, unknown op). Per-op business outcomes — anchor-not-found,
/// malformed-markdown, etc. — are returned as <c>EditResult</c> with
/// <c>success: false</c> inside a normal <c>ok: true</c> envelope.
/// </summary>
internal static class Program
{
    public static int Main(string[] args)
    {
        using var stdin = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8);
        using var stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false))
        {
            NewLine = "\n",
            AutoFlush = false,
        };

        Console.Error.WriteLine("docxodus-pyhost ready");

        string? line;
        while ((line = stdin.ReadLine()) is not null)
        {
            if (line.Length == 0) continue;

            long requestId = -1;
            string? op = null;
            try
            {
                using var doc = JsonDocument.Parse(line);
                var root = doc.RootElement;
                if (!root.TryGetProperty("id", out var idEl) || idEl.ValueKind != JsonValueKind.Number)
                    throw new FormatException("request missing numeric \"id\"");
                requestId = idEl.GetInt64();

                if (!root.TryGetProperty("op", out var opEl) || opEl.ValueKind != JsonValueKind.String)
                    throw new FormatException("request missing string \"op\"");
                op = opEl.GetString()!;

                var argsEl = root.TryGetProperty("args", out var a) ? a : default;

                if (op == "shutdown")
                {
                    WriteOk(stdout, requestId, "null");
                    stdout.Flush();
                    SessionRegistry.DisposeAll();
                    return 0;
                }

                var resultJson = Dispatcher.Dispatch(op, argsEl);
                WriteOk(stdout, requestId, resultJson);
            }
            catch (UnknownOpException ex)
            {
                WriteError(stdout, requestId, "unknown_op", ex.Message, null);
            }
            catch (JsonException ex)
            {
                WriteError(stdout, requestId, "malformed_request", ex.Message, null);
            }
            catch (FormatException ex)
            {
                WriteError(stdout, requestId, "malformed_request", ex.Message, null);
            }
            catch (ArgumentException ex)
            {
                // Includes SessionRegistry.Get's "unknown session handle" — distinct from
                // EditError(SessionDisposed), which is returned via the OK envelope by ops
                // that survive a closed-handle check.
                WriteError(stdout, requestId, "invalid_argument", ex.Message, null);
            }
            catch (Exception ex)
            {
                WriteError(stdout, requestId, "internal_error", ex.Message, ex.StackTrace);
            }

            stdout.Flush();
        }

        SessionRegistry.DisposeAll();
        return 0;
    }

    private static void WriteOk(StreamWriter w, long id, string resultJson)
    {
        w.Write("{\"id\":");
        w.Write(id);
        w.Write(",\"ok\":true,\"result\":");
        w.Write(resultJson);
        w.Write("}\n");
    }

    private static void WriteError(StreamWriter w, long id, string code, string message, string? trace)
    {
        w.Write("{\"id\":");
        w.Write(id);
        w.Write(",\"ok\":false,\"error\":{\"code\":");
        w.Write(DocxSessionJson.JsonString(code));
        w.Write(",\"message\":");
        w.Write(DocxSessionJson.JsonString(message));
        if (trace is not null)
        {
            w.Write(",\"trace\":");
            w.Write(DocxSessionJson.JsonString(trace));
        }
        w.Write("}}\n");
    }
}

internal sealed class UnknownOpException : Exception
{
    public UnknownOpException(string op) : base($"unknown op: {op}") { }
}

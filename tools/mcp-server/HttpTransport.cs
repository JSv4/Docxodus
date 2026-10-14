#nullable enable

using System;
using System.IO;
using System.Net;
using System.Text;

namespace Docxodus.McpServer;

/// <summary>
/// Minimal MCP streamable-HTTP transport (<c>--http PORT</c>): each POST carries one JSON-RPC
/// message and gets a single <c>application/json</c> response — the non-streaming shape the
/// streamable-HTTP binding allows a server to choose. No SSE, no session-id handshake, no TLS:
/// this exists so the stdio server can be put behind a tunnel (e.g. <c>ngrok http PORT</c>) and
/// pointed at from a ChatGPT Apps / remote-MCP developer setup, which cannot spawn a local
/// process. Requests are processed one at a time: <see cref="Run"/> handles each request inline
/// on the serial accept loop, and the lock below pins that even if handling ever moves off it.
/// <see cref="SessionStore"/> does now serialize per session rather than requiring a
/// single-threaded caller, but nothing in the shipped server exercises that: this transport is
/// serial and stdio is single-threaded, so the per-session dispatch gate is forward-looking.
/// Making concurrent requests real means dispatching handling off the accept loop AND dropping
/// this lock — a deliberate behaviour change, not a comment fix.
/// </summary>
internal static class HttpTransport
{
    public static int Run(SessionStore store, IDocumentStore documents, int port)
    {
        using var listener = new HttpListener();
        listener.Prefixes.Add($"http://localhost:{port}/");
        listener.Start();
        Console.Error.WriteLine(
            $"docxodus-mcp listening on http://localhost:{port}/ (streamable HTTP, storage: {documents.Kind}, root: {documents.RootDescription})");

        var gate = new object();
        while (true)
        {
            HttpListenerContext context;
            try { context = listener.GetContext(); }
            catch (HttpListenerException) { break; }
            catch (ObjectDisposedException) { break; }

            try
            {
                Handle(store, gate, context);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[mcp:http] {ex.Message}");
                try { context.Response.Abort(); } catch { /* already gone */ }
            }
        }

        store.CloseAll();
        return 0;
    }

    private static void Handle(SessionStore store, object gate, HttpListenerContext context)
    {
        var response = context.Response;
        // Permissive CORS: browser-embedded MCP clients preflight; a stdio-equivalent local dev
        // server has nothing origin-specific to protect beyond what its document scope enforces.
        response.Headers["Access-Control-Allow-Origin"] = "*";
        response.Headers["Access-Control-Allow-Methods"] = "POST, GET, DELETE, OPTIONS";
        response.Headers["Access-Control-Allow-Headers"] = "content-type, mcp-session-id, mcp-protocol-version, authorization";

        switch (context.Request.HttpMethod)
        {
            case "OPTIONS":
                response.StatusCode = 204;
                response.Close();
                return;

            case "DELETE": // session termination — nothing held per-connection here
                response.StatusCode = 200;
                response.Close();
                return;

            case "POST":
                break;

            default: // no SSE stream to offer on GET
                response.StatusCode = 405;
                response.Headers["Allow"] = "POST, DELETE, OPTIONS";
                response.Close();
                return;
        }

        string body;
        using (var reader = new StreamReader(context.Request.InputStream, Encoding.UTF8))
            body = reader.ReadToEnd();

        string? reply;
        lock (gate)
            reply = Program.ProcessMessage(store, body, out _); // shutdown is a stdio concept; ignored here

        if (reply is null)
        {
            response.StatusCode = 202; // notification accepted, nothing to say
            response.Close();
            return;
        }

        var bytes = Encoding.UTF8.GetBytes(reply);
        response.StatusCode = 200;
        response.ContentType = "application/json";
        response.ContentLength64 = bytes.Length;
        response.OutputStream.Write(bytes, 0, bytes.Length);
        response.Close();
    }
}

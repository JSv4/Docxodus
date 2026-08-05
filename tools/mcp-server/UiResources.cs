#nullable enable

using System.Text;
using System.Text.Json;

namespace Docxodus.McpServer;

/// <summary>
/// MCP Apps (extension <c>io.modelcontextprotocol/ui</c>, spec 2026-01-26) support: the embedded
/// <c>ui://</c> viewer template, the <c>resources/list</c>/<c>resources/read</c> handlers that
/// serve it, the <c>_meta</c> stamped onto preview-capable tools in <c>tools/list</c>, and the
/// tool-result wrapping that routes rendered HTML to the widget without flooding the model's
/// context. ChatGPT's Apps SDK reads the same <c>_meta.ui.resourceUri</c> field (the
/// <c>openai/*</c> keys are compatibility aliases), so one template serves both hosts; the
/// viewer itself detects <c>window.openai</c> vs the MCP Apps postMessage bridge at runtime.
/// See docs/architecture/docx_agent_server.md ("Inline preview / MCP Apps").
/// </summary>
internal static class UiResources
{
    public const string ViewerUri = "ui://docxodus/viewer.html";
    public const string ViewerMimeType = "text/html;profile=mcp-app";
    public const string ExtensionId = "io.modelcontextprotocol/ui";

    /// <summary>Result-<c>_meta</c> key carrying the rendered document HTML. Hosts deliver the
    /// full tool result to the widget (MCP Apps: <c>ui/notifications/tool-result</c>; ChatGPT:
    /// <c>window.openai.toolResponseMetadata</c>) but only <c>content</c>/<c>structuredContent</c>
    /// reach the model — so a multi-hundred-KB render costs the conversation nothing.</summary>
    public const string HtmlMetaKey = "docxodus/html";

    /// <summary>Extra <c>_meta</c> for a tool descriptor in <c>tools/list</c>, or null for tools
    /// with no UI. <c>ui.resourceUri</c> is the MCP Apps standard key; the <c>openai/*</c> keys
    /// are ChatGPT's documented aliases (<c>widgetAccessible</c> lets the widget itself call
    /// <c>docxodus_preview</c> to refresh after edits).</summary>
    public static string? ToolMetaJson(string toolName) => toolName switch
    {
        "docxodus_open" or "docxodus_preview" =>
            "{\"ui\":{\"resourceUri\":\"" + ViewerUri + "\",\"visibility\":[\"model\",\"app\"]}"
            + ",\"openai/outputTemplate\":\"" + ViewerUri + "\""
            + ",\"openai/widgetAccessible\":true"
            + ",\"openai/toolInvocation/invoking\":\"Rendering document…\""
            + ",\"openai/toolInvocation/invoked\":\"Rendered document preview\"}",
        _ => null,
    };

    /// <summary>The <c>_meta.ui</c> block advertised on the viewer resource itself. No external
    /// domains are declared: the template is fully self-contained, so it renders under the spec's
    /// default CSP (<c>script-src 'self' 'unsafe-inline'</c>, no network) in every host.</summary>
    private const string ResourceUiMetaJson =
        "{\"ui\":{\"prefersBorder\":true,\"csp\":{\"connectDomains\":[],\"resourceDomains\":[]}}}";

    public static string BuildResourcesListResult() =>
        "{\"resources\":[{\"uri\":\"" + ViewerUri + "\""
        + ",\"name\":\"Docxodus document preview\""
        + ",\"description\":\"Inline HTML viewer for a Docxodus editing session. Rendered by the host next to docxodus_open/docxodus_preview results.\""
        + ",\"mimeType\":\"" + ViewerMimeType + "\""
        + ",\"_meta\":" + ResourceUiMetaJson + "}]}";

    public static string BuildResourcesReadResult(JsonElement readParams)
    {
        if (readParams.ValueKind != JsonValueKind.Object
            || !readParams.TryGetProperty("uri", out var uriEl) || uriEl.ValueKind != JsonValueKind.String)
            throw new InvalidParamsException("resources/read params missing string \"uri\"");
        var uri = uriEl.GetString()!;
        if (uri != ViewerUri)
            throw new InvalidParamsException($"unknown resource: {uri}");

        return "{\"contents\":[{\"uri\":\"" + ViewerUri + "\""
            + ",\"mimeType\":\"" + ViewerMimeType + "\""
            + ",\"text\":" + JsonRpcIo.JsonString(ViewerHtml)
            + ",\"_meta\":" + ResourceUiMetaJson + "}]}";
    }

    /// <summary>
    /// Wrap a dispatcher result into the MCP <c>tools/call</c> result envelope. The default shape
    /// (everything in one text content block) is unchanged; the two widget-bearing tools get the
    /// richer shape: <c>docxodus_open</c> additionally exposes <c>structuredContent</c> so its
    /// widget instance learns the session id, and <c>docxodus_preview</c> moves the rendered HTML
    /// out of the model-visible content into result <c>_meta</c>.
    /// </summary>
    public static string WrapToolResult(string toolName, string resultJson, bool isError)
    {
        if (!isError)
        {
            if (toolName == "docxodus_preview")
            {
                using var doc = JsonDocument.Parse(resultJson);
                var root = doc.RootElement;
                var html = root.GetProperty("html").GetString() ?? string.Empty;
                var sessionId = root.GetProperty("sessionId").GetString();
                string? anchorId = root.TryGetProperty("anchorId", out var a) && a.ValueKind == JsonValueKind.String
                    ? a.GetString() : null;

                var structured = new StringBuilder("{\"sessionId\":")
                    .Append(JsonRpcIo.JsonString(sessionId!));
                if (anchorId is not null)
                    structured.Append(",\"anchorId\":").Append(JsonRpcIo.JsonString(anchorId));
                structured.Append(",\"htmlLength\":").Append(html.Length).Append('}');

                // content text = the structuredContent summary, NOT the HTML: the model needs to
                // know the render happened (and how big it was), never the markup itself.
                return "{\"content\":[{\"type\":\"text\",\"text\":" + JsonRpcIo.JsonString(structured.ToString()) + "}]"
                    + ",\"structuredContent\":" + structured
                    + ",\"_meta\":{\"" + HtmlMetaKey + "\":" + JsonRpcIo.JsonString(html) + "}"
                    + ",\"isError\":false}";
            }

            if (toolName == "docxodus_open")
            {
                // The open result is already a small JSON object ({sessionId, path}); mirror it as
                // structuredContent so the widget (which only sees structured data, not content
                // text) can pick up the session id and fetch its first render.
                return "{\"content\":[{\"type\":\"text\",\"text\":" + JsonRpcIo.JsonString(resultJson) + "}]"
                    + ",\"structuredContent\":" + resultJson
                    + ",\"isError\":false}";
            }
        }

        return "{\"content\":[{\"type\":\"text\",\"text\":" + JsonRpcIo.JsonString(resultJson) + "}]"
            + ",\"isError\":" + (isError ? "true" : "false") + "}";
    }

    /// <summary>
    /// The viewer template. Fully self-contained (no external fetches, so it passes the default
    /// MCP Apps CSP) and dual-host: with <c>window.openai</c> present it reads
    /// <c>toolOutput</c>/<c>toolResponseMetadata</c> and refreshes via <c>callTool</c>; otherwise
    /// it speaks MCP Apps JSON-RPC over <c>postMessage</c> (<c>ui/initialize</c> handshake,
    /// <c>ui/notifications/tool-*</c> notifications, <c>tools/call</c> for refresh).
    /// </summary>
    public const string ViewerHtml = """"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Docxodus document preview</title>
<style>
  html, body { margin: 0; padding: 0; background: #fff; color: #1a1a1a;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif; }
  #dxo-bar { display: flex; align-items: center; gap: 8px; padding: 6px 10px;
    border-bottom: 1px solid #e3e3e3; background: #fafafa; position: sticky; top: 0;
    font-size: 12px; }
  #dxo-bar strong { font-size: 12px; font-weight: 600; }
  #dxo-status { color: #6b6b6b; flex: 1; overflow: hidden; text-overflow: ellipsis;
    white-space: nowrap; }
  #dxo-refresh { font-size: 12px; padding: 2px 10px; border: 1px solid #c9c9c9;
    border-radius: 4px; background: #fff; cursor: pointer; }
  #dxo-refresh:hover { background: #f0f0f0; }
  #dxo-content { padding: 16px 20px; overflow: auto; }
</style>
</head>
<body>
<div id="dxo-bar">
  <strong>Docxodus</strong>
  <span id="dxo-status">loading&hellip;</span>
  <button id="dxo-refresh" type="button">Refresh</button>
</div>
<div id="dxo-content"></div>
<script>
(function () {
  "use strict";
  var state = { sessionId: null, anchorId: null, rendered: false };
  var statusEl, contentEl;

  function setStatus(text) { if (statusEl) statusEl.textContent = text; }

  function renderDocumentHtml(fullHtml) {
    var doc = new DOMParser().parseFromString(fullHtml, "text/html");
    var stale = document.querySelectorAll("style[data-docxodus]");
    for (var i = 0; i < stale.length; i++) stale[i].parentNode.removeChild(stale[i]);
    var styles = doc.querySelectorAll("style");
    for (var j = 0; j < styles.length; j++) {
      var copy = document.createElement("style");
      copy.setAttribute("data-docxodus", "");
      copy.textContent = styles[j].textContent;
      document.head.appendChild(copy);
    }
    contentEl.innerHTML = "";
    var nodes = doc.body ? doc.body.childNodes : [];
    for (var k = 0; k < nodes.length; k++)
      contentEl.appendChild(document.importNode(nodes[k], true));
    state.rendered = true;
    setStatus(state.sessionId ? "session " + state.sessionId
      + (state.anchorId ? " · " + state.anchorId : "") : "rendered");
  }

  // Accept a tools/call result object from any host path; returns true if HTML was rendered.
  function acceptToolResult(result) {
    if (!result || typeof result !== "object") return false;
    var sc = result.structuredContent || {};
    if (typeof sc.sessionId === "string") state.sessionId = sc.sessionId;
    if (typeof sc.anchorId === "string") state.anchorId = sc.anchorId;
    var meta = result._meta || {};
    var html = meta["docxodus/html"];
    if (typeof html === "string" && html.length > 0) { renderDocumentHtml(html); return true; }
    return false;
  }

  // ---- MCP Apps postMessage JSON-RPC bridge ------------------------------
  var pending = {}, nextRpcId = 1;
  function rpcRequest(method, params) {
    return new Promise(function (resolve, reject) {
      var id = nextRpcId++;
      pending[id] = { resolve: resolve, reject: reject };
      window.parent.postMessage({ jsonrpc: "2.0", id: id, method: method, params: params }, "*");
    });
  }

  window.addEventListener("message", function (ev) {
    var msg = ev.data;
    if (!msg || msg.jsonrpc !== "2.0") return;
    if (msg.id !== undefined && msg.method === undefined) {
      var waiter = pending[msg.id];
      if (!waiter) return;
      delete pending[msg.id];
      if (msg.error) waiter.reject(new Error(msg.error.message || "host error"));
      else waiter.resolve(msg.result);
      return;
    }
    if (!msg.method) return;
    if (msg.id !== undefined) {
      // Host request (ping, ui/resource-teardown): acknowledge; nothing to tear down here.
      window.parent.postMessage({ jsonrpc: "2.0", id: msg.id, result: {} }, "*");
      return;
    }
    var params = msg.params || {};
    if (msg.method === "ui/notifications/tool-result") {
      var result = params.result !== undefined ? params.result : params;
      if (!acceptToolResult(result) && state.sessionId && !state.rendered) fetchPreview();
    } else if (msg.method === "ui/notifications/tool-input") {
      var args = (params.arguments !== undefined ? params.arguments : params) || {};
      if (typeof args.sessionId === "string") state.sessionId = args.sessionId;
      if (typeof args.anchorId === "string") state.anchorId = args.anchorId;
    }
  });

  // ---- Preview fetch (both hosts) ----------------------------------------
  function callPreview() {
    var args = { sessionId: state.sessionId };
    if (state.anchorId) args.anchorId = state.anchorId;
    if (window.openai && typeof window.openai.callTool === "function")
      return Promise.resolve(window.openai.callTool("docxodus_preview", args));
    return rpcRequest("tools/call", { name: "docxodus_preview", arguments: args });
  }

  function fetchPreview() {
    if (!state.sessionId) { setStatus("no session id available"); return; }
    setStatus("rendering…");
    callPreview().then(function (result) {
      if (!acceptToolResult(result)) setStatus("preview returned no HTML");
    }, function (err) {
      setStatus("error: " + (err && err.message ? err.message : String(err)));
    });
  }

  // ---- Boot ---------------------------------------------------------------
  window.addEventListener("DOMContentLoaded", function () {
    statusEl = document.getElementById("dxo-status");
    contentEl = document.getElementById("dxo-content");
    document.getElementById("dxo-refresh").addEventListener("click", fetchPreview);

    if (window.openai) {
      // ChatGPT Apps SDK host: tool data is exposed as globals, updated via an event.
      var readGlobals = function () {
        var ok = acceptToolResult({
          structuredContent: window.openai.toolOutput || {},
          _meta: window.openai.toolResponseMetadata || {}
        });
        if (!ok && state.sessionId && !state.rendered) fetchPreview();
      };
      window.addEventListener("openai:set_globals", readGlobals);
      readGlobals();
      if (!state.rendered && !state.sessionId) setStatus("waiting for document…");
      return;
    }

    // MCP Apps host: handshake, then the host pushes tool-input/tool-result notifications.
    rpcRequest("ui/initialize", { protocolVersion: "2026-01-26", capabilities: {} })
      .then(function () { if (!state.rendered) setStatus("waiting for document…"); },
            function () { /* tolerate hosts that skip the handshake */ });
    setTimeout(function () {
      if (!state.rendered && state.sessionId) fetchPreview();
    }, 1500);
  });
})();
</script>
</body>
</html>
"""";
}

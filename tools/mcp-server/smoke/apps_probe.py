#!/usr/bin/env python3
"""Smoke-test the MCP Apps (inline preview) surface of docxodus-mcp.

Exercises, over stdio: initialize (extension capability), resources/list,
resources/read (the ui:// viewer template), tools/list (_meta.ui on the
widget-bearing tools), docxodus_open (structuredContent), docxodus_preview
(HTML in result _meta, none in the model-visible content), and a single-block
preview. Then repeats initialize + preview over the --http transport.

Usage:
    python3 apps_probe.py --server bin/Debug/net10.0/docxodus-mcp.dll --docx ../../TestFiles/HC001-5DayTourPlanTemplate.docx
"""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
from pathlib import Path
from typing import Any

VIEWER_URI = "ui://docxodus/viewer.html"
VIEWER_MIME = "text/html;profile=mcp-app"
EXTENSION_ID = "io.modelcontextprotocol/ui"
HTML_META_KEY = "docxodus/html"

_checks = 0


def check(condition: bool, label: str) -> None:
    global _checks
    _checks += 1
    if condition:
        print(f"  ok  {label}")
    else:
        print(f"FAIL  {label}", file=sys.stderr)
        raise SystemExit(1)


class StdioClient:
    def __init__(self, argv: list[str], root: Path) -> None:
        self.process = subprocess.Popen(
            argv,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.DEVNULL,
            text=True,
            env={**os.environ, "DOCXODUS_STORAGE_ROOT": str(root)},
        )
        self.next_id = 1

    def request(self, method: str, params: Any = None) -> dict[str, Any]:
        request_id = self.next_id
        self.next_id += 1
        message: dict[str, Any] = {"jsonrpc": "2.0", "id": request_id, "method": method}
        if params is not None:
            message["params"] = params
        assert self.process.stdin and self.process.stdout
        self.process.stdin.write(json.dumps(message, separators=(",", ":")) + "\n")
        self.process.stdin.flush()
        for line in self.process.stdout:
            if not line.strip():
                continue
            reply = json.loads(line)
            if reply.get("id") == request_id:
                if "error" in reply:
                    raise RuntimeError(f"{method}: {reply['error']}")
                return reply["result"]
        raise RuntimeError(f"server closed before response to {method}")

    def close(self) -> None:
        try:
            self.request("shutdown")
        except Exception:
            pass
        self.process.wait(timeout=10)


def tool_result_text(result: dict[str, Any]) -> Any:
    text = result["content"][0]["text"]
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return text


def run_stdio(server_argv: list[str], root: Path, docx_name: str) -> None:
    print("stdio transport:")
    client = StdioClient(server_argv, root)
    try:
        init = client.request("initialize", {
            "protocolVersion": "2026-01-26",
            "capabilities": {},
            "clientInfo": {"name": "apps_probe", "version": "0"},
        })
        check(init["protocolVersion"] == "2026-01-26", "initialize echoes requested protocolVersion")
        check(EXTENSION_ID in init["capabilities"].get("extensions", {}),
              "initialize declares the MCP Apps extension capability")
        check(VIEWER_MIME in init["capabilities"]["extensions"][EXTENSION_ID]["mimeTypes"],
              "extension capability lists the mcp-app mimeType")
        check("resources" in init["capabilities"], "initialize declares resources capability")

        resources = client.request("resources/list")["resources"]
        viewer = next((r for r in resources if r["uri"] == VIEWER_URI), None)
        check(viewer is not None, "resources/list contains the viewer template")
        check(viewer["mimeType"] == VIEWER_MIME, "viewer resource has the mcp-app mimeType")

        contents = client.request("resources/read", {"uri": VIEWER_URI})["contents"][0]
        check(contents["mimeType"] == VIEWER_MIME, "resources/read returns the mcp-app mimeType")
        check("<!DOCTYPE html>" in contents["text"] and "docxodus_preview" in contents["text"],
              "viewer template is self-contained HTML that calls docxodus_preview")
        check("ui" in contents.get("_meta", {}), "viewer resource carries _meta.ui")

        tools = {t["name"]: t for t in client.request("tools/list")["tools"]}
        check("docxodus_preview" in tools, "tools/list advertises docxodus_preview")
        for name in ("docxodus_open", "docxodus_preview"):
            meta = tools[name].get("_meta", {})
            check(meta.get("ui", {}).get("resourceUri") == VIEWER_URI,
                  f"{name} carries _meta.ui.resourceUri")
            check(meta.get("openai/outputTemplate") == VIEWER_URI,
                  f"{name} carries the openai/outputTemplate alias")
        check("_meta" not in tools["docxodus_save"], "non-widget tools carry no _meta")

        opened = client.request("tools/call", {
            "name": "docxodus_open", "arguments": {"path": docx_name}})
        session_id = opened["structuredContent"]["sessionId"]
        check(bool(session_id), "docxodus_open exposes sessionId via structuredContent")

        preview = client.request("tools/call", {
            "name": "docxodus_preview", "arguments": {"sessionId": session_id}})
        html = preview["_meta"][HTML_META_KEY]
        check(html.lstrip().lower().startswith(("<html", "<!doctype")),
              "docxodus_preview puts the rendered document in result _meta")
        check(preview["structuredContent"]["htmlLength"] == len(html),
              "structuredContent.htmlLength matches the rendered HTML")
        check("<html" not in preview["content"][0]["text"],
              "model-visible content contains no markup")

        anchors = tool_result_text(client.request("tools/call", {
            "name": "docxodus_search",
            "arguments": {"sessionId": session_id, "mode": "kind", "query": "p", "maxResults": 1}}))
        anchor_id = anchors["matches"][0]["id"] if anchors.get("matches") else None
        if anchor_id:
            block = client.request("tools/call", {
                "name": "docxodus_preview",
                "arguments": {"sessionId": session_id, "anchorId": anchor_id}})
            check(len(block["_meta"][HTML_META_KEY]) > 0, "single-block preview renders HTML")
            check(block["structuredContent"]["anchorId"] == anchor_id,
                  "single-block preview echoes the anchor id")
        else:
            print("  ..  no paragraph anchor found; skipping single-block preview")

        client.request("tools/call", {"name": "docxodus_close", "arguments": {"sessionId": session_id}})
    finally:
        client.close()


def run_http(server_argv: list[str], root: Path, docx_name: str, port: int) -> None:
    print("streamable-HTTP transport:")
    process = subprocess.Popen(
        server_argv + ["--http", str(port)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        env={**os.environ, "DOCXODUS_STORAGE_ROOT": str(root)},
    )
    try:
        url = f"http://localhost:{port}/mcp"
        next_id = [1]

        def post(method: str, params: Any = None) -> dict[str, Any]:
            message: dict[str, Any] = {"jsonrpc": "2.0", "id": next_id[0], "method": method}
            next_id[0] += 1
            if params is not None:
                message["params"] = params
            body = json.dumps(message).encode()
            request = urllib.request.Request(
                url, data=body, headers={"Content-Type": "application/json"})
            for attempt in range(20):
                try:
                    with urllib.request.urlopen(request, timeout=10) as reply:
                        return json.loads(reply.read())["result"]
                except (urllib.error.URLError, ConnectionError):
                    if attempt == 19:
                        raise
                    time.sleep(0.5)
            raise RuntimeError("unreachable")

        init = post("initialize", {"protocolVersion": "2026-01-26", "capabilities": {},
                                   "clientInfo": {"name": "apps_probe", "version": "0"}})
        check(EXTENSION_ID in init["capabilities"].get("extensions", {}),
              "HTTP initialize declares the MCP Apps extension")

        opened = post("tools/call", {"name": "docxodus_open", "arguments": {"path": docx_name}})
        session_id = opened["structuredContent"]["sessionId"]
        preview = post("tools/call", {"name": "docxodus_preview",
                                      "arguments": {"sessionId": session_id}})
        check(len(preview["_meta"][HTML_META_KEY]) > 0, "HTTP preview renders HTML in _meta")
        post("tools/call", {"name": "docxodus_close", "arguments": {"sessionId": session_id}})
    finally:
        process.terminate()
        process.wait(timeout=10)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--server", required=True,
                        help="Path to docxodus-mcp.dll (run with 'dotnet') or the apphost binary")
    parser.add_argument("--docx", required=True, help="A .docx to open and render")
    parser.add_argument("--dotnet", default="dotnet", help="dotnet executable for a .dll server")
    parser.add_argument("--port", type=int, default=8931, help="Port for the --http leg")
    args = parser.parse_args()

    server = Path(args.server).resolve()
    argv = [args.dotnet, str(server)] if server.suffix == ".dll" else [str(server)]

    with tempfile.TemporaryDirectory(prefix="docxodus-apps-probe-") as scratch:
        root = Path(scratch)
        docx = Path(args.docx).resolve()
        shutil.copy(docx, root / docx.name)
        run_stdio(argv, root, docx.name)
        run_http(argv, root, docx.name, args.port)

    print(f"all {_checks} checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

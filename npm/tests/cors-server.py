#!/usr/bin/env python3
"""Static file server with CORS headers — emulates a CDN origin for tests.

The cdn-embed spec loads dist/embed.bundle.js (and, transitively, the WASM
assets in dist/wasm/) from THIS server while the page itself is served from the
main test server, so every fetch is genuinely cross-origin — the same shape as
loading the published package from jsDelivr/unpkg. Mirrors the CDN's contract:
`Access-Control-Allow-Origin: *` and a correct `application/wasm` MIME type.

Usage: cors-server.py PORT DIRECTORY
"""
import http.server
import functools
import sys


class CorsHandler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".wasm": "application/wasm",
        ".js": "text/javascript",
        ".mjs": "text/javascript",
    }

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        super().end_headers()

    def log_message(self, *args):
        pass  # keep Playwright's webServer output quiet


if __name__ == "__main__":
    port = int(sys.argv[1])
    directory = sys.argv[2]
    handler = functools.partial(CorsHandler, directory=directory)
    http.server.ThreadingHTTPServer(("127.0.0.1", port), handler).serve_forever()

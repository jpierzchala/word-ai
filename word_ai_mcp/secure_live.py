"""Isolated live-only runtime. No imports of the legacy file tools/session store.

HTTPS is exclusively an Office add-in API. MCP is newline JSON-RPC over a
container-local Unix socket, reached via `docker exec -i ... stdio`.
Document content and capabilities never persist; restarting revokes everything.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import os
import re
import secrets
import socket
import socketserver
import ssl
import sys
import threading
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

MAX_BODY = 4 * 1024 * 1024  # bounded 2 MiB image + Base64/JSON overhead
SOCKET = "/run/word-ai/mcp.sock"
TERMINAL = {"succeeded", "failed", "cancelled", "expired", "unknown"}
if __package__:
    from .document_state import State, TOOLS
else:
    from document_state import State, TOOLS


def rpc(state, req):
    rid = req.get("id")
    method = req.get("method")
    if method == "notifications/initialized" or (rid is None and str(method).startswith("notifications/")):
        return None
    try:
        if method == "initialize":
            result = {"protocolVersion": req.get("params", {}).get("protocolVersion", "2025-11-25"),
                      "serverInfo": {"name": "word-ai-secure-live", "version": "2.0.0"}, "capabilities": {"tools": {}}}
        elif method == "ping":
            result = {}
        elif method in {"resources/list", "resources/templates/list", "prompts/list"}:
            result = {"resources" if method == "resources/list" else "resourceTemplates" if method == "resources/templates/list" else "prompts": []}
        elif method == "tools/list":
            result = {"tools": [{"name": name, "description": desc,
                                  "inputSchema": {"type": "object", "properties": {k: v if isinstance(v, dict) else {"type": v} for k, v in props.items()},
                                                  "required": list(props), "additionalProperties": False},
                                  "annotations": {"readOnlyHint": name not in {"word_session_connect", "word_session_apply", "word_session_cancel"},
                                                  "destructiveHint": name == "word_session_apply", "openWorldHint": False}}
                                 for name, (desc, props) in TOOLS.items()]}
        elif method == "tools/call":
            params = req.get("params", {})
            try:
                value = state.call(params.get("name"), params.get("arguments", {}))
                result = {"content": [{"type": "text", "text": json.dumps(value, ensure_ascii=False)}], "isError": False}
            except (ValueError, PermissionError) as exc:
                result = {"content": [{"type": "text", "text": str(exc)}], "isError": True}
        else:
            raise ValueError("Unsupported method")
        return {"jsonrpc": "2.0", "id": rid, "result": result}
    except (ValueError, TypeError, KeyError):
        return {"jsonrpc": "2.0", "id": rid, "error": {"code": -32600, "message": "Invalid request"}}


def handler(state, static, port=3100):
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_):
            pass  # Never log credentials, request bodies or document URLs.

        def setup(self):
            super().setup()
            self.connection.settimeout(10)

        def send(self, status, value, kind="application/json"):
            data = json.dumps(value).encode() if kind == "application/json" else value
            self.send_response(status)
            self.send_header("Content-Type", kind)
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Referrer-Policy", "no-referrer")
            self.send_header("Content-Security-Policy", "default-src 'none'; script-src 'self' https://appsforoffice.microsoft.com; style-src 'self'; connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'none'")
            self.end_headers()
            self.wfile.write(data)

        def valid_origin(self):
            return (self.headers.get_all("Host") == [f"localhost:{port}"] and
                    self.headers.get("Origin") in (None, f"https://localhost:{port}"))

        def do_OPTIONS(self):
            self.send(403, {"error": "Cross-origin requests disabled"})

        def do_GET(self):
            if not self.valid_origin():
                self.send(403, {"error": "Host or Origin rejected"})
                return
            # Office appends host metadata to SourceLocation. Route only the
            # exact path; do not interpret, persist or log its query values.
            path = self.path.partition("?")[0]
            files = {"/": ("taskpane.html", "text/html; charset=utf-8"),
                     "/taskpane.html": ("taskpane.html", "text/html; charset=utf-8"),
                     "/taskpane.js": ("taskpane.js", "text/javascript; charset=utf-8"),
                     "/safety.js": ("safety.js", "text/javascript; charset=utf-8"),
                     "/document-edit.js": ("document-edit.js", "text/javascript; charset=utf-8"),
                     "/live-ops.js": ("live-ops.js", "text/javascript; charset=utf-8"),
                     "/section-move.js": ("section-move.js", "text/javascript; charset=utf-8"),
                     "/taskpane.css": ("taskpane.css", "text/css; charset=utf-8"),
                     "/icon-32.png": ("icon-32.png", "image/png"),
                     "/icon-64.png": ("icon-64.png", "image/png")}
            if path == "/health":
                self.send(200, {"ok": True, "profile": "secure-live"})
            elif path in files:
                name, kind = files[path]
                self.send(200, (static / name).read_bytes(), kind)
            else:
                self.send(404, {"error": "Not found"})

        def do_POST(self):
            if not self.valid_origin():
                self.send(403, {"error": "Host or Origin rejected"})
                return
            routes = {f"/office/{x}": x for x in ("pair", "pair-status", "pair-cancel", "poll", "begin", "result", "disconnect")}
            if self.path not in routes:
                self.send(404, {"error": "Not found"})
                return
            auth = self.headers.get("Authorization", "")
            if self.path == "/office/pair" and self.headers.get("Origin") != f"https://localhost:{port}":
                self.send(403, {"error": "Pairing must originate in the local Word panel"})
                return
            if self.path != "/office/pair" and (not auth.startswith("Bearer ") or len(auth) < 30):
                self.send(401, {"error": "Authentication required"})
                return
            try:
                if self.headers.get("Transfer-Encoding") or len(self.headers.get_all("Content-Length", [])) != 1:
                    raise ValueError("Invalid framing")
                length = int(self.headers.get("Content-Length", "0"))
                if not 0 < length <= MAX_BODY or self.headers.get_content_type() != "application/json":
                    raise ValueError("Invalid body size or content type")
                body = json.loads(self.rfile.read(length))
                if not isinstance(body, dict):
                    raise ValueError("Expected JSON object")
                value = state.office(routes[self.path], body, auth[7:])
                self.send(200, value)
            except PermissionError:
                self.send(401, {"error": "Unauthorized"})
            except (ValueError, TypeError, KeyError):
                self.send(400, {"error": "Request rejected: invalid, stale or mismatched state"})
    return Handler


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=["serve", "stdio"])
    args = parser.parse_args()
    if args.mode == "stdio":
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as client:
            client.connect(SOCKET)
            stream = client.makefile("rwb")
            for line in sys.stdin:
                stream.write(line.encode())
                stream.flush()
                response = stream.readline(MAX_BODY * 2)
                if not response:
                    break
                if response.strip() != b"null":
                    sys.stdout.write(response.decode())
                    sys.stdout.flush()
        return
    os.umask(0o077)
    state = State()

    class RpcHandler(socketserver.StreamRequestHandler):
        def handle(self):
            while line := self.rfile.readline(MAX_BODY + 1):
                if len(line) > MAX_BODY:
                    return
                try:
                    request = json.loads(line)
                    response = rpc(state, request) if isinstance(request, dict) else None
                except (ValueError, TypeError):
                    response = {"jsonrpc": "2.0", "id": None, "error": {"code": -32700, "message": "Invalid JSON"}}
                self.wfile.write(json.dumps(response).encode() + b"\n")
                self.wfile.flush()

    Path(SOCKET).parent.mkdir(parents=True, exist_ok=True)
    Path(SOCKET).unlink(missing_ok=True)
    local = socketserver.ThreadingUnixStreamServer(SOCKET, RpcHandler)
    local.daemon_threads = True
    threading.Thread(target=local.serve_forever, daemon=True).start()
    def cleanup():
        while True:
            time.sleep(5)
            with state.lock:
                state.clean()
    threading.Thread(target=cleanup, daemon=True).start()
    context = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
    context.minimum_version = ssl.TLSVersion.TLSv1_2
    context.load_cert_chain("/secrets/localhost.crt", "/secrets/localhost.key")
    class HttpsServer(ThreadingHTTPServer):
        def get_request(self):
            connection, address = super().get_request()
            # Handshake in a request thread with the handler's timeout, never
            # in the accept loop (one silent connection must not block others).
            return context.wrap_socket(connection, server_side=True, do_handshake_on_connect=False), address
    server = HttpsServer(("0.0.0.0", 3100), handler(state, Path("/app/static")))
    server.serve_forever()


if __name__ == "__main__":
    main()

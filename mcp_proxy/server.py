"""
Catalyst retailer pipeline — MCP front door.

Claude's connector check won't follow the 302 that every Apps Script web app
answers with, so this sits in front: it speaks Streamable HTTP MCP on /mcp and
forwards each JSON-RPC message to the Apps Script deployment, which holds all
the actual logic (tools, sheet access, Salesforce rules).

Config, both from the environment:
  APPS_SCRIPT_URL   the /exec URL of the Retailers sync deployment  (required)
  MCP_BEARER        optional shared secret; when set, requests must carry
                    "Authorization: Bearer <secret>"

Stdlib only, so the image stays a plain python:slim with no install step.
"""

import json
import os
import urllib.error
import urllib.request
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

APPS_SCRIPT_URL = os.environ.get("APPS_SCRIPT_URL", "").strip()
MCP_BEARER = os.environ.get("MCP_BEARER", "").strip()
PORT = int(os.environ.get("PORT", "8080"))
UPSTREAM_TIMEOUT = 120


def call_upstream(payload):
    """POST one JSON-RPC message to Apps Script. urllib follows its redirect."""
    req = urllib.request.Request(
        APPS_SCRIPT_URL,
        method="POST",
        data=json.dumps(payload).encode(),
        headers={"Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=UPSTREAM_TIMEOUT) as r:
        body = r.read().decode("utf-8").strip()
    return json.loads(body) if body else None


class Handler(BaseHTTPRequestHandler):
    server_version = "catalyst-mcp/1.0"

    # ── plumbing ──────────────────────────────────────────────────────────
    def _send(self, code, body=None, ctype="application/json"):
        raw = b"" if body is None else json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers",
                         "Content-Type, Authorization, Mcp-Session-Id, MCP-Protocol-Version")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.end_headers()
        if raw:
            self.wfile.write(raw)

    def _rpc_error(self, rid, code, message, http=200):
        self._send(http, {"jsonrpc": "2.0", "id": rid, "error": {"code": code, "message": message}})

    def _authorized(self):
        if not MCP_BEARER:
            return True
        return self.headers.get("Authorization", "") == f"Bearer {MCP_BEARER}"

    def log_message(self, fmt, *args):   # one tidy line per request
        print(f"{self.address_string()} {fmt % args}", flush=True)

    # ── routes ────────────────────────────────────────────────────────────
    def do_OPTIONS(self):
        self._send(204, None)

    def do_GET(self):
        path = self.path.split("?")[0].rstrip("/") or "/"
        if path in ("/", "/health"):
            self._send(200, {"ok": True, "service": "catalyst-retailers-mcp",
                             "upstream_configured": bool(APPS_SCRIPT_URL),
                             "auth_required": bool(MCP_BEARER)})
            return
        if path == "/mcp":
            # No server-initiated stream; the spec allows refusing the SSE channel.
            self._send(405, {"error": "this server answers MCP on POST /mcp only"})
            return
        self._send(404, {"error": "not found"})

    def do_POST(self):
        path = self.path.split("?")[0].rstrip("/") or "/"
        if path != "/mcp":
            self._send(404, {"error": "not found"})
            return
        if not self._authorized():
            self._send(401, {"error": "unauthorized"})
            return
        if not APPS_SCRIPT_URL:
            self._rpc_error(None, -32603, "APPS_SCRIPT_URL is not set on the proxy")
            return

        try:
            length = int(self.headers.get("Content-Length") or 0)
            payload = json.loads(self.rfile.read(length) or b"{}")
        except (ValueError, json.JSONDecodeError):
            self._rpc_error(None, -32700, "parse error", http=400)
            return

        # A notification has no id and expects 202 with no body.
        is_notification = isinstance(payload, dict) and "id" not in payload
        rid = payload.get("id") if isinstance(payload, dict) else None

        try:
            result = call_upstream(payload)
        except urllib.error.HTTPError as e:
            self._rpc_error(rid, -32603, f"upstream HTTP {e.code}")
            return
        except Exception as e:                      # timeout, DNS, bad JSON back
            self._rpc_error(rid, -32603, f"upstream call failed: {e}")
            return

        if is_notification or result is None:
            self._send(202, None)
            return
        self._send(200, result)


if __name__ == "__main__":
    if not APPS_SCRIPT_URL:
        print("[warn] APPS_SCRIPT_URL is empty — /mcp will return errors", flush=True)
    print(f"[start] listening on :{PORT} -> {APPS_SCRIPT_URL[:60]}", flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()

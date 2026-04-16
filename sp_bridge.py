#!/usr/bin/env python3
"""
SharePoint Bridge Server — local HTTP proxy for uploading files to SharePoint.

Runs on port 8081. The browser app POSTs a file here, and this server
uploads it to SharePoint via Microsoft Graph API using Azure CLI auth
(no admin consent required).

Usage:
    python sp_bridge.py

Requires:
    pip install msal requests
"""

import json, os, sys, re
from http.server import HTTPServer, BaseHTTPRequestHandler
import msal, requests

PORT = 8081

# Azure CLI public client — pre-consented, no admin needed
CLIENT_ID = "04b07795-8ddb-461a-bbee-02f9e1bf7b46"
AUTHORITY = "https://login.microsoftonline.com/organizations"
SCOPES    = ["https://graph.microsoft.com/.default"]
GRAPH     = "https://graph.microsoft.com/v1.0"

TOKEN_CACHE_FILE = os.path.join(os.path.dirname(__file__), "_token_cache.bin")


# ── Auth ─────────────────────────────────────────────────────────────────────

def _build_app():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_FILE):
        cache.deserialize(open(TOKEN_CACHE_FILE, "r").read())
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    return app, cache


def _save_cache(cache):
    if cache.has_state_changed:
        with open(TOKEN_CACHE_FILE, "w") as f:
            f.write(cache.serialize())


def get_token():
    app, cache = _build_app()
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]

    print("[bridge] Opening browser for sign-in ...")
    result = app.acquire_token_interactive(SCOPES)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description", str(result)))
    _save_cache(cache)
    return result["access_token"]


# ── Graph helpers ────────────────────────────────────────────────────────────

def resolve_drive(token, sp_host, site_path, lib_name):
    """Resolve drive-id for a SharePoint doc library."""
    url = f"{GRAPH}/sites/{sp_host}:{site_path}"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    site_id = r.json()["id"]

    url = f"{GRAPH}/sites/{site_id}/drives"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    for d in r.json()["value"]:
        if d["name"] == lib_name:
            return d["id"]

    names = [d["name"] for d in r.json()["value"]]
    raise ValueError(f"Drive '{lib_name}' not found. Available: {names}")


def upload_file(token, drive_id, filename, data):
    """Upload file bytes to SharePoint via Graph API."""
    url = f"{GRAPH}/drives/{drive_id}/root:/{filename}:/content"
    r = requests.put(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/octet-stream"},
        data=data,
    )
    r.raise_for_status()
    item = r.json()
    return item.get("webUrl", item.get("id", "ok"))


def parse_sharepoint_url(url):
    """
    Parse a SharePoint URL like:
      https://tenant.sharepoint.com/sites/SiteName/DocLibName
    Returns (host, site_path, lib_name)
    """
    m = re.match(
        r'https?://([^/]+\.sharepoint\.com)(/sites/[^/]+)/([^/?#]+)',
        url.strip().rstrip('/')
    )
    if not m:
        raise ValueError(
            f"Invalid SharePoint URL: {url}\n"
            f"Expected format: https://<tenant>.sharepoint.com/sites/<SiteName>/<DocLibName>"
        )
    return m.group(1), m.group(2), m.group(3)


# ── HTTP Handler ─────────────────────────────────────────────────────────────

class BridgeHandler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def do_POST(self):
        if self.path != "/upload":
            self._error(404, "Not found")
            return

        try:
            sp_url   = self.headers.get("X-SP-Url", "")
            filename = self.headers.get("X-Filename", "upload.xlsx")
            length   = int(self.headers.get("Content-Length", 0))

            if not sp_url:
                self._error(400, "Missing X-SP-Url header")
                return
            if length == 0:
                self._error(400, "Empty body")
                return

            sp_host, site_path, lib_name = parse_sharepoint_url(sp_url)
            body = self.rfile.read(length)

            print(f"[bridge] Uploading '{filename}' to {sp_host}{site_path}/{lib_name} ...")
            token = get_token()
            drive_id = resolve_drive(token, sp_host, site_path, lib_name)
            web_url = upload_file(token, drive_id, filename, body)

            result = {"ok": True, "webUrl": web_url, "filename": filename}
            self._json_response(200, result)
            print(f"[bridge] Uploaded: {web_url}")

        except Exception as e:
            print(f"[bridge] Error: {e}")
            self._error(500, str(e))

    def do_GET(self):
        if self.path == "/health":
            self._json_response(200, {"ok": True, "service": "sp-bridge"})
        else:
            self._error(404, "Not found")

    def _cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS, GET")
        self.send_header("Access-Control-Allow-Headers",
                         "Content-Type, X-SP-Url, X-Filename")

    def _json_response(self, code, obj):
        body = json.dumps(obj).encode()
        self.send_response(code)
        self._cors_headers()
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _error(self, code, msg):
        self._json_response(code, {"ok": False, "error": msg})

    def log_message(self, fmt, *args):
        pass  # suppress default logging


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    # Pre-authenticate so the user doesn't wait on first upload
    print("[bridge] Authenticating ...")
    try:
        get_token()
        print("[bridge] Authenticated")
    except Exception as e:
        print(f"[bridge] Auth failed: {e}")
        sys.exit(1)

    server = HTTPServer(("127.0.0.1", PORT), BridgeHandler)
    print(f"[bridge] Listening on http://127.0.0.1:{PORT}")
    print(f"[bridge] POST /upload  — upload file to SharePoint")
    print(f"[bridge] GET  /health  — health check")
    print(f"[bridge] Press Ctrl+C to stop")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[bridge] Stopped")
        server.server_close()


if __name__ == "__main__":
    main()

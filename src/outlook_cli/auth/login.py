"""Capture an MSAL.js session via a Bookmarklet.

Flow (matches what the user does in their normal browser, no automation):

1. CLI prints instructions to create a bookmarklet and start the server.
2. CLI spins up a temporary HTTP server on localhost.
3. User signs into Outlook in their own browser.
4. User clicks the bookmarklet.
5. The bookmarklet generates a temporary localhost HTML page via data URI or window.open,
   bypassing Outlook's strict Content Security Policy (connect-src).
6. That localhost page POSTs the localStorage dump to the CLI's localhost server.
7. CLI parses the MSAL.js state and persists credentials to disk.
"""

from __future__ import annotations

import base64
import http.server
import json
import sys
import threading
from dataclasses import dataclass, field
from datetime import UTC, datetime
from typing import Any

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.errors import UserError

OUTLOOK_LOGIN_URL = "https://outlook.cloud.microsoft/mail/"

# Outlook's strict CSP blocks connect-src (fetch) and form-action (forms),
# and TrustedTypes blocks document.write.
# We bypass all of this by encoding the payload into the URL hash (fragment)
# and opening a new tab to localhost. The localhost page then reads its own hash
# and POSTs it to the CLI.
#
# Only MSAL credential/account entries are sent. A web client can keep megabytes
# of app state in localStorage; shipping all of it overflows the browser's URL
# length limit and the fragment arrives truncated (JSONDecodeError mid-string).
# The filter matches exactly what ``parse_msal_localstorage`` consumes —
# RefreshToken, IdToken, and account entries — and skips AccessToken blobs.
# Falls back to the full dump if nothing matches, in case MSAL's schema shifts.
_BOOKMARKLET_TEMPLATE = (
    "javascript:(function(){"
    "var o={},i,k,v;"
    "for(i=0;i<localStorage.length;i++){"
    "k=localStorage.key(i);v=localStorage.getItem(k);"
    "if(v&&(v.indexOf('\"RefreshToken\"')>-1||v.indexOf('\"IdToken\"')>-1||"
    "(v.indexOf('\"homeAccountId\"')>-1&&v.indexOf('\"username\"')>-1)))o[k]=v;"
    "}"
    "if(!Object.keys(o).length)o=Object.fromEntries(Object.entries(localStorage));"
    "window.open('http://127.0.0.1:__PORT__/auth#'+"
    "encodeURIComponent(JSON.stringify(o)),'_blank');"
    "})()"
)

_AUTH_PAGE_HTML = """<html>
<head><title>Outlook CLI: Authenticating...</title></head>
<body style="font-family: sans-serif; text-align: center; padding-top: 50px;">
<h1 id="status">Outlook CLI: Processing...</h1>
<p>Do not close this tab until you see a success message.</p>
<script>
var s = document.getElementById("status");
var payload = "";
try {
  payload = decodeURIComponent(window.location.hash.substring(1));
} catch (e) {
  s.innerText = "Error: Failed to decode payload.";
}
if (payload) {
  fetch('/submit', { method: 'POST', body: payload }).then(function(r) {
    if (r.ok) {
      s.innerText = "Success! You can close this tab.";
      s.style.color = "green";
    } else {
      s.innerText = "Error: CLI rejected the payload.";
      s.style.color = "red";
    }
  }).catch(function() {
    s.innerText = "Error: Failed to send to CLI.";
    s.style.color = "red";
  });
}
</script>
</body>
</html>
"""


@dataclass
class ParsedSession:
    refresh_token: str
    client_id: str
    tenant_id: str
    home_account_id: str
    username: str
    id_token_claims: dict[str, Any] = field(default_factory=dict)
    is_foci: bool = False


def _decode_jwt_payload(jwt: str) -> dict[str, Any]:
    parts = jwt.split(".")
    if len(parts) < 2:
        return {}
    payload_b64 = parts[1]
    padded = payload_b64 + "=" * (-len(payload_b64) % 4)
    try:
        decoded: dict[str, Any] = json.loads(base64.urlsafe_b64decode(padded).decode("utf-8"))
        return decoded
    except (ValueError, UnicodeDecodeError):
        return {}


def _parsed_entries(storage: dict[str, str]) -> list[dict[str, Any]]:
    """Return all JSON-object values from a localStorage dump."""
    entries: list[dict[str, Any]] = []
    for value in storage.values():
        try:
            candidate = json.loads(value)
        except (json.JSONDecodeError, TypeError):
            continue
        if isinstance(candidate, dict):
            entries.append(candidate)
    return entries


def _last_updated(entry: dict[str, Any]) -> int:
    try:
        return int(entry.get("lastUpdatedAt", 0))
    except (TypeError, ValueError):
        return 0


def parse_msal_localstorage(storage: dict[str, str]) -> ParsedSession:
    """Extract session credentials from a dump of MSAL.js localStorage."""
    entries = _parsed_entries(storage)

    refresh_candidates = [
        e for e in entries if e.get("credentialType") == "RefreshToken" and e.get("secret")
    ]
    if not refresh_candidates:
        raise UserError(
            "No MSAL refresh token found in the captured data. "
            "Make sure you signed in at https://outlook.cloud.microsoft/ and ran "
            "the bookmarklet on that page."
        )
    # Prefer the most recently issued refresh token (newer MSAL schema version).
    refresh_entry = max(refresh_candidates, key=_last_updated)

    refresh_token = refresh_entry["secret"]
    client_id = refresh_entry["clientId"]
    home_account_id = refresh_entry["homeAccountId"]
    tenant_id = home_account_id.split(".", 1)[1] if "." in home_account_id else ""

    username = ""
    for entry in entries:
        if (
            entry.get("authorityType")
            and entry.get("homeAccountId") == home_account_id
            and entry.get("username")
        ):
            username = entry["username"]
            break

    id_claims: dict[str, Any] = {}
    id_candidates = [
        e
        for e in entries
        if e.get("credentialType") == "IdToken"
        and e.get("clientId") == client_id
        and e.get("homeAccountId") == home_account_id
    ]
    if id_candidates:
        idtok = max(id_candidates, key=_last_updated)
        id_claims = _decode_jwt_payload(idtok.get("secret", ""))

    return ParsedSession(
        refresh_token=refresh_token,
        client_id=client_id,
        tenant_id=tenant_id,
        home_account_id=home_account_id,
        username=username,
        id_token_claims=id_claims,
        is_foci=str(refresh_entry.get("familyId", "")) == "1",
    )


def _make_credentials(session: ParsedSession) -> Credentials:
    return Credentials(
        version=1,
        acquired_at=datetime.now(UTC).isoformat().replace("+00:00", "Z"),
        tenant_id=session.tenant_id,
        client_id=session.client_id,
        home_account_id=session.home_account_id,
        username=session.username,
        refresh_token=session.refresh_token,
        id_token_claims=session.id_token_claims,
    )


class _AuthServer(http.server.HTTPServer):
    """HTTPServer with typed slots for the captured payload + error."""

    captured_storage: dict[str, Any] | None = None
    capture_error: str | None = None


class BookmarkletAuthHandler(http.server.BaseHTTPRequestHandler):
    server: _AuthServer

    def do_GET(self) -> None:
        if self.path.startswith("/auth"):
            self.send_response(200)
            self.send_header("Content-Type", "text/html")
            self.end_headers()
            self.wfile.write(_AUTH_PAGE_HTML.encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self) -> None:
        if self.path == "/submit":
            content_length = int(self.headers.get("Content-Length", 0))
            post_data = self.rfile.read(content_length)

            try:
                self.server.captured_storage = json.loads(post_data.decode("utf-8"))
            except (json.JSONDecodeError, UnicodeDecodeError) as e:
                # Byte count makes a truncated fragment obvious in the error.
                self.server.capture_error = f"{e} (received {len(post_data)} bytes)"

            self.send_response(200)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()

            # Shut down the server on a separate thread to unblock serve_forever
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format: str, *args: Any) -> None:
        pass


def capture_session_via_bookmarklet(
    *,
    port: int = 49152,
) -> ParsedSession:
    """Print instructions and start a temporary server to catch the bookmarklet POST."""
    server: _AuthServer | None = None
    for p in range(port, port + 100):
        try:
            server = _AuthServer(("127.0.0.1", p), BookmarkletAuthHandler)
            port = p
            break
        except OSError:
            continue

    if not server:
        raise UserError("Could not bind to a local port for the bookmarklet server.")

    bookmarklet = _BOOKMARKLET_TEMPLATE.replace("__PORT__", str(port))

    bar = "=" * 64
    print("", file=sys.stderr)
    print(bar, file=sys.stderr)
    print(" Sign in to Microsoft 365", file=sys.stderr)
    print(bar, file=sys.stderr)
    print("", file=sys.stderr)
    print(
        "  1. Create a new bookmark in your browser (drag it to your bookmarks bar).",
        file=sys.stderr,
    )
    print("     Name it: Outlook CLI Login", file=sys.stderr)
    print(f"     URL: {bookmarklet}", file=sys.stderr)
    print("", file=sys.stderr)
    print("  2. Open this URL in your browser:", file=sys.stderr)
    print(f"       {OUTLOOK_LOGIN_URL}", file=sys.stderr)
    print("", file=sys.stderr)
    print(
        "  3. Sign in normally. Once your inbox loads, click the 'Outlook CLI Login'",
        file=sys.stderr,
    )
    print("     bookmark you just made.", file=sys.stderr)
    print(bar, file=sys.stderr)
    print("\nWaiting for authentication (press Ctrl+C to cancel)...", file=sys.stderr)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nLogin cancelled.", file=sys.stderr)
        sys.exit(1)
    finally:
        server.server_close()

    if server.capture_error:
        raise UserError(f"Failed to parse data from bookmarklet: {server.capture_error}")

    storage = server.captured_storage
    if not storage:
        raise UserError("No data received from bookmarklet.")

    if not isinstance(storage, dict):
        raise UserError("Received JSON is not an object.")

    return parse_msal_localstorage(storage)


def save_session(session: ParsedSession) -> Credentials:
    """Persist a ParsedSession to ~/.config/outlook-cli/credentials.json."""
    creds = _make_credentials(session)
    save(creds)
    return creds

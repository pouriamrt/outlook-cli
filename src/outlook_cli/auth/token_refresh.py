"""Mint short-lived access tokens from the stored refresh token. Cache by scope."""

from __future__ import annotations

import json
import logging
import os
import time
from pathlib import Path
from typing import Any

import httpx
from filelock import FileLock

from outlook_cli.auth.token_store import load, rotate_refresh_token
from outlook_cli.config import access_tokens_path, cache_home, ensure_dirs, http_verify
from outlook_cli.errors import SessionExpired

logger = logging.getLogger(__name__)

GRAPH_SCOPE = "https://outlook.office.com/.default"
_SKEW_SECONDS = 60


def _default_scope() -> str:
    """Resolve the scope to request, honouring an env override.

    The One Outlook web client only has consent for the Outlook REST audience,
    not Microsoft Graph, so users on those tenants must export
    ``OUTLOOK_CLI_API_SCOPE=https://outlook.office.com/.default``.
    """
    return os.environ.get("OUTLOOK_CLI_API_SCOPE", GRAPH_SCOPE)


def _cache_lock() -> Path:
    return cache_home() / "access_tokens.lock"


def _read_cache() -> dict[str, dict[str, Any]]:
    p = access_tokens_path()
    if not p.exists():
        return {}
    try:
        data: dict[str, dict[str, Any]] = json.loads(p.read_text(encoding="utf-8"))
        return data
    except (json.JSONDecodeError, OSError):
        return {}


def _write_cache(cache: dict[str, dict[str, Any]]) -> None:
    ensure_dirs()
    p = access_tokens_path()
    tmp = p.with_suffix(".json.tmp")
    payload = json.dumps(cache, indent=2)
    if os.name == "nt":
        tmp.write_text(payload, encoding="utf-8")
    else:
        tmp.unlink(missing_ok=True)
        fd = os.open(str(tmp), os.O_CREAT | os.O_WRONLY | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(payload)
    os.replace(tmp, p)


def _cached_token(scope: str) -> str | None:
    with FileLock(str(_cache_lock())):
        cache = _read_cache()
    entry = cache.get(scope)
    if entry is None:
        return None
    if entry["expires_at"] - _SKEW_SECONDS <= time.time():
        return None
    token: str = entry["access_token"]
    return token


def _store_token(scope: str, access_token: str, expires_in: int) -> None:
    with FileLock(str(_cache_lock())):
        cache = _read_cache()
        cache[scope] = {
            "access_token": access_token,
            "expires_at": int(time.time()) + int(expires_in),
        }
        _write_cache(cache)


def _redeem(scope: str) -> str:
    creds = load()
    url = f"https://login.microsoftonline.com/{creds.tenant_id}/oauth2/v2.0/token"
    body = {
        "client_id": creds.client_id,
        "grant_type": "refresh_token",
        "refresh_token": creds.refresh_token,
        "scope": scope,
    }
    # SPA-issued refresh tokens (AADSTS9002327) require an Origin header
    # matching a registered SPA redirect URI. outlook.cloud.microsoft is the
    # page MSAL.js ran on when the token was minted.
    headers = {"Origin": "https://outlook.cloud.microsoft"}
    with httpx.Client(timeout=30.0, verify=http_verify()) as client:
        r = client.post(url, data=body, headers=headers)
    if r.status_code == 400:
        try:
            err_payload = r.json()
        except ValueError:
            err_payload = {}
        err = err_payload.get("error", "")
        if err in ("invalid_grant", "interaction_required"):
            raise SessionExpired("Session expired. Run 'outlook login' to re-authenticate.")
        # Surface the full AAD error so the user knows what to fix (e.g.
        # invalid_scope, unauthorized_client, consent_required).
        desc = err_payload.get("error_description", r.text)
        raise SessionExpired(
            f"Token endpoint rejected refresh (HTTP 400, {err or 'unknown'}):\n{desc}"
        )
    r.raise_for_status()
    payload = r.json()
    access_token: str = payload["access_token"]
    expires_in: int = payload.get("expires_in", 3599)
    new_rt: str | None = payload.get("refresh_token")
    if new_rt and new_rt != creds.refresh_token:
        logger.debug("Refresh token rotated.")
        rotate_refresh_token(new_rt)
    _store_token(scope, access_token, expires_in)
    return access_token


def get_token(scope: str | None = None) -> str:
    """Return a valid access token for the given scope, minting if needed."""
    if scope is None:
        scope = _default_scope()
    cached = _cached_token(scope)
    if cached:
        return cached
    return _redeem(scope)

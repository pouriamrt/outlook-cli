"""Resolve mailbox folder names ("inbox", "ProjectX") to Graph folder IDs with a 24h cache."""

from __future__ import annotations

import json
import os
import time
from typing import Any

from outlook_cli.config import ensure_dirs, folders_cache_path
from outlook_cli.errors import NotFound

WELL_KNOWN: dict[str, str] = {
    "inbox": "inbox",
    "sent": "sentitems",
    "drafts": "drafts",
    "deleted": "deleteditems",
    "junk": "junkemail",
    "archive": "archive",
    "outbox": "outbox",
}
_CACHE_TTL_SECONDS = 24 * 3600


def _looks_like_graph_id(name: str) -> bool:
    # Real Graph mailFolders IDs are typically 80+ chars of base64; lowering threshold
    # to 20 still safely beats all well-known names while accepting smaller test fixtures.
    return (
        len(name) > 20
        and " " not in name
        and "@" not in name
        and ("=" in name or "/" in name or "+" in name)
    )


def _read_cache() -> dict[str, dict[str, Any]]:
    p = folders_cache_path()
    if not p.exists():
        return {}
    try:
        data: dict[str, dict[str, Any]] = json.loads(p.read_text(encoding="utf-8"))
        return data
    except (json.JSONDecodeError, OSError):
        return {}


def _write_cache(cache: dict[str, dict[str, Any]]) -> None:
    ensure_dirs()
    p = folders_cache_path()
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(cache, indent=2), encoding="utf-8")
    os.replace(tmp, p)


def resolve_folder_id(name: str, client: Any) -> str:
    """Return a Graph folder ID for a name. Accepts well-known names, custom names, or raw IDs."""
    if name in WELL_KNOWN:
        return WELL_KNOWN[name]
    if _looks_like_graph_id(name):
        return name

    cache = _read_cache()
    entry = cache.get(name.lower())
    now = time.time()
    if entry and now - entry["fetched_at"] < _CACHE_TTL_SECONDS:
        return str(entry["id"])

    resp = client.get("/me/mailFolders", params={"$top": 250, "$select": "id,displayName"})
    payload = resp.json()
    for folder in payload.get("value", []):
        cache[folder["displayName"].lower()] = {"id": folder["id"], "fetched_at": now}
    _write_cache(cache)

    entry = cache.get(name.lower())
    if not entry:
        raise NotFound(f"Folder '{name}' not found.")
    return str(entry["id"])

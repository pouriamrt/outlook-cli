"""Per-command-family caches: index → Graph ID."""

from __future__ import annotations

import json
import os
from collections.abc import Callable
from pathlib import Path
from typing import Any

from outlook_cli.config import cal_index_path, ensure_dirs, mail_index_path
from outlook_cli.errors import NotFound

_PATHS: dict[str, Callable[[], Path]] = {
    "mail": mail_index_path,
    "cal": cal_index_path,
}


def _path(family: str) -> Path:
    return _PATHS[family]()


def store(family: str, items: list[dict[str, Any]]) -> None:
    """Write [{ "index": int, "id": str }, ...] to the cache for the family."""
    ensure_dirs()
    payload = {
        "family": family,
        "items": [{"index": it["index"], "id": it["id"]} for it in items],
    }
    p = _path(family)
    tmp = p.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    os.replace(tmp, p)


def resolve(family: str, ref: int | str) -> str:
    """Given an int (short index) or string (Graph ID), return the Graph ID."""
    if isinstance(ref, str):
        if len(ref) > 20 and " " not in ref and ("=" in ref or "/" in ref or "+" in ref):
            return ref
        try:
            ref = int(ref)
        except ValueError as exc:
            raise NotFound(f"Cannot interpret '{ref}' as index or ID.") from exc

    p = _path(family)
    if not p.exists():
        raise NotFound(f"No prior listing. Run 'outlook {family} list' first.")
    payload = json.loads(p.read_text(encoding="utf-8"))
    for item in payload.get("items", []):
        if item["index"] == ref:
            return str(item["id"])
    raise NotFound(f"No item with index {ref} in last {family} listing.")

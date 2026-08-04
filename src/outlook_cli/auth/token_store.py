"""Persist and rotate Microsoft refresh-token credentials."""

import json
import os
from pathlib import Path
from typing import Any

from filelock import FileLock
from pydantic import BaseModel, ConfigDict, Field

from outlook_cli.config import config_home, credentials_path, ensure_dirs
from outlook_cli.errors import SessionExpired


class Credentials(BaseModel):
    model_config = ConfigDict(frozen=True)
    version: int = 1
    acquired_at: str
    tenant_id: str
    client_id: str
    home_account_id: str
    username: str
    refresh_token: str
    id_token_claims: dict[str, Any] = Field(default_factory=dict)


def _lock_path() -> Path:
    return config_home() / "credentials.lock"


def _atomic_write_json(target: Path, payload: str) -> None:
    """Atomically write JSON payload to target with mode 0600 on POSIX from creation."""
    tmp = target.with_suffix(".json.tmp")
    if os.name == "nt":
        tmp.write_text(payload, encoding="utf-8")
    else:
        # On POSIX, create the tmp file with restrictive permissions from the start
        # to avoid a brief world-readable window between open and chmod.
        tmp.unlink(missing_ok=True)
        fd = os.open(str(tmp), os.O_CREAT | os.O_WRONLY | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(payload)
    os.replace(tmp, target)


def save(creds: Credentials) -> None:
    """Atomically persist Credentials to disk with mode 0600 on POSIX."""
    ensure_dirs()
    target = credentials_path()
    payload = creds.model_dump_json(indent=2)
    with FileLock(str(_lock_path())):
        _atomic_write_json(target, payload)


def load() -> Credentials:
    """Read Credentials from disk. Raises SessionExpired if missing or malformed."""
    target = credentials_path()
    if not target.exists():
        raise SessionExpired("Not logged in. Run 'outlook login'.")
    with FileLock(str(_lock_path())):
        try:
            data = json.loads(target.read_text(encoding="utf-8"))
            return Credentials.model_validate(data)
        except (json.JSONDecodeError, ValueError) as exc:
            raise SessionExpired(
                f"Credentials file is corrupt ({exc}). Run 'outlook login'."
            ) from exc


def rotate_refresh_token(new_rt: str) -> None:
    """Replace just the refresh_token field, leaving other fields untouched.

    Holds the credentials lock across the entire read-modify-write to prevent
    concurrent CLI processes from interleaving updates.

    Note: Credentials is frozen, so we use model_copy(update=...) rather than
    direct attribute assignment.
    """
    target = credentials_path()
    with FileLock(str(_lock_path())):
        # Read current credentials inside the lock
        if not target.exists():
            raise SessionExpired("Not logged in. Run 'outlook login'.")
        try:
            data = json.loads(target.read_text(encoding="utf-8"))
            creds = Credentials.model_validate(data)
        except (json.JSONDecodeError, ValueError) as exc:
            raise SessionExpired(
                f"Credentials file is corrupt ({exc}). Run 'outlook login'."
            ) from exc
        # Produce updated copy (model_copy handles frozen models correctly)
        new_creds = creds.model_copy(update={"refresh_token": new_rt})
        payload = new_creds.model_dump_json(indent=2)
        _atomic_write_json(target, payload)


def delete() -> None:
    """Remove credentials and the lock file. Idempotent."""
    with FileLock(str(_lock_path())):
        credentials_path().unlink(missing_ok=True)
